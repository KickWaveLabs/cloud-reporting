#!/usr/bin/env python3
"""Pay-as-you-go list prices for the meters that differ between AKS and Container Apps.

Source: Azure Retail Prices API (https://prices.azure.com/api/retail/prices), public and anonymous:
no sign-in and no call to the bank's tenant. Prices are list prices in USD, without contract
discounts; the workbook compares them with the effective rate measured in Cost Management.

Meters fetched per region:
  Container Apps: every Consumption meter (vCPU and memory active/idle per second, requests,
                  environment management, private endpoint, dedicated plan).
  AKS: cluster tier (Standard uptime SLA, Long Term Support).
  Node VMs: Linux pay-as-you-go and 1/3-year reservations for each VM size in the inventory
            node pools, plus --vm-sizes.
  OS disks: Premium SSD and Standard SSD tiers in --disk-tiers.
  Network: public IP (Standard, regional) and Standard Load Balancer (global meters).
  Defender for Containers: per vCore-hour (global meter).

Outputs (in --output-dir):
  prices.csv          one row per meter with a `key` the workbook looks up, region and the query used
  prices_summary.json
"""

import argparse
import json
import os
import re
import time

import requests

import aa_common as aa

azure = aa.azure
ENDPOINT = "https://prices.azure.com/api/retail/prices"
API = "2023-01-01-preview"

# (key, filter template, row predicate). {region} is replaced per region; global meters omit it.
STATIC = [
    ("aca", "serviceName eq 'Azure Container Apps' and armRegionName eq '{region}' and priceType eq 'Consumption'", None),
    ("aks_tier", "serviceName eq 'Azure Kubernetes Service' and armRegionName eq '{region}' "
                 "and productName eq 'Azure Kubernetes Service'", None),
    ("public_ip", "serviceName eq 'Virtual Network' and armRegionName eq '{region}' and productName eq 'IP Addresses' "
                  "and priceType eq 'Consumption'", lambda item: item["skuName"] == "Standard"),
    ("load_balancer", "serviceName eq 'Load Balancer' and armRegionName eq 'Global' and priceType eq 'Consumption'",
     lambda item: item["skuName"] == "Standard"),
    ("defender_containers", "serviceName eq 'Microsoft Defender for Cloud' and contains(meterName, 'Containers Node')",
     None),
]

ACA_KEYS = {
    "Standard vCPU Active Usage": "aca_vcpu_active_s",
    "Standard vCPU Idle Usage": "aca_vcpu_idle_s",
    "Standard Memory Active Usage": "aca_mem_active_gib_s",
    "Standard Memory Idle Usage": "aca_mem_idle_gib_s",
    "Standard Requests": "aca_requests_million",
    "Environment Management Hour": "aca_env_management_hour",
    "Environment Private Endpoint": "aca_env_private_endpoint_hour",
    "Dedicated Plan Management": "aca_dedicated_management_hour",
    "Dedicated vCPU Usage": "aca_dedicated_vcpu_hour",
    "Dedicated Memory Usage": "aca_dedicated_mem_gib_hour",
}


def canonical_vm_size(size):
    """ARM accepts any case, the Retail Prices API does not: Standard_F4S_v2 -> Standard_F4s_v2."""
    return re.sub(r"^(Standard_[A-Z]+\d+)([A-Za-z]*)", lambda match: match.group(1) + match.group(2).lower(), size)


def fetch(filter_text, session):
    items, url, params = [], ENDPOINT, {"api-version": API, "$filter": filter_text}
    while url:
        for attempt in range(1, 6):
            response = session.get(url, params=params, timeout=60)
            if response.status_code == 429 or response.status_code >= 500:
                time.sleep(5 * attempt)
                continue
            break
        response.raise_for_status()
        body = response.json()
        items.extend(body.get("Items", []))
        url, params = body.get("NextPageLink"), None
        time.sleep(0.3)
    return items


def key_for(group, item):
    meter = item["meterName"]
    if group == "aca":
        return ACA_KEYS.get(meter, "aca_other")
    if group == "aks_tier":
        return {"Standard Uptime SLA": "aks_standard_tier_hour",
                "Standard Long Term Support": "aks_lts_hour"}.get(meter, "aks_other")
    if group == "public_ip":
        return "public_ip_standard_hour" if meter == "Standard IPv4 Static Public IP" else "public_ip_other"
    if group == "load_balancer":
        return {"Standard Included LB Rules and Outbound Rules": "lb_standard_rules_hour",
                "Standard Overage LB Rules and Outbound Rules": "lb_standard_overage_rule_hour",
                "Standard Data Processed": "lb_standard_data_gb"}.get(meter, "lb_other")
    if group == "defender_containers":
        return "defender_containers_vcore_hour" if item["retailPrice"] else "defender_other"
    if group == "vm":
        term = item.get("reservationTerm") or ""
        suffix = {"1 Year": "_1y", "3 Years": "_3y"}.get(term, "_payg")
        return f"vm_{item['armSkuName']}{suffix}"
    if group == "disk":
        return f"disk_{item['skuName'].split()[0]}_month"
    return group


def row_for(group, item, region, filter_text):
    price = item["retailPrice"]
    unit = item["unitOfMeasure"]
    # Reservations are quoted for the whole term; the workbook needs USD per hour like pay-as-you-go.
    hourly = ""
    if item.get("type") == "Reservation":
        hours = {"1 Year": 8760, "3 Years": 26280}.get(item.get("reservationTerm"), 0)
        hourly = round(price / hours, 6) if hours else ""
    return {
        "key": key_for(group, item), "region": region, "service": item["serviceName"], "product": item["productName"],
        "sku": item["skuName"], "meter": item["meterName"], "unit": unit, "retail_price": price,
        "hourly_equivalent": hourly if hourly != "" else (price if unit in ("1 Hour", "1/Hour") else ""),
        "price_type": item.get("type", ""), "reservation_term": item.get("reservationTerm") or "",
        "tier_minimum_units": item.get("tierMinimumUnits", 0), "effective_start": item.get("effectiveStartDate", "")[:10],
        "meter_id": item.get("meterId", ""), "arm_sku": item.get("armSkuName", ""), "query": filter_text,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output-dir", default=aa.results_path("aa-prices"))
    parser.add_argument("--inventory", default=aa.results_path("aa-inventory"),
                        help="Folder with aks_node_pools.csv; its VM sizes are priced too")
    parser.add_argument("--regions", default="eastus2,centralus", help="Comma-separated ARM region names")
    parser.add_argument("--vm-sizes", default="Standard_D8s_v5,Standard_D8ds_v5,Standard_D8as_v5,Standard_D8s_v6",
                        help="VM sizes priced even if not in the inventory")
    parser.add_argument("--disk-tiers", default="P6,P10,P15,E6,E10", help="Managed disk tiers for OS disks")
    parser.add_argument("--delimiter", default=",")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)

    sizes = {size.strip() for size in args.vm_sizes.split(",") if size.strip()}
    sizes |= {canonical_vm_size(row["vm_size"])
              for row in aa.read_csv(os.path.join(args.inventory, "aks_node_pools.csv")) if row.get("vm_size")}
    regions = [region.strip() for region in args.regions.split(",") if region.strip()]
    session = requests.Session()
    rows, queries, seen_global = [], [], set()
    for region in regions:
        jobs = [(group, template.format(region=region), predicate) for group, template, predicate in STATIC]
        for size in sorted(sizes):
            jobs.append(("vm", f"serviceName eq 'Virtual Machines' and armRegionName eq '{region}' "
                               f"and armSkuName eq '{size}'",
                         lambda item: (item["productName"].startswith("Virtual Machines")
                                       and not item["productName"].endswith("Windows")
                                       and "Spot" not in item["skuName"] and "Low Priority" not in item["skuName"]
                                       and item.get("type") in ("Consumption", "Reservation"))))
        for tier in args.disk_tiers.split(","):
            product = "Premium SSD Managed Disks" if tier.startswith("P") else "Standard SSD Managed Disks"
            jobs.append(("disk", f"serviceName eq 'Storage' and armRegionName eq '{region}' and productName eq '{product}' "
                                 f"and skuName eq '{tier.strip()} LRS'",
                         lambda item: item["meterName"].endswith("Disk")))
        for group, filter_text, predicate in jobs:
            global_meter = "armRegionName" not in filter_text or "'Global'" in filter_text
            if global_meter and filter_text in seen_global:
                continue
            seen_global.add(filter_text)
            items = [item for item in fetch(filter_text, session) if predicate is None or predicate(item)]
            queries.append({"group": group, "region": "Global" if global_meter else region, "filter": filter_text,
                            "rows": len(items)})
            azure.LOG.info("%-20s %-10s %3d meters", group, region, len(items))
            for item in items:
                # Global meters are listed once per region so the workbook can look them up by region.
                for target in (regions if global_meter else [region]):
                    rows.append(row_for(group, item, target, filter_text))

    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("prices.csv", rows)
    missing = [query for query in queries if not query["rows"]]
    output.json("prices_summary.json", {
        "generated_at": azure.iso(azure.utc_now()), "source": ENDPOINT, "api_version": API, "currency": "USD",
        "regions": regions, "vm_sizes": sorted(sizes), "queries": queries, "queries_without_rows": missing,
        "note": "List prices (pay-as-you-go and reservations). Contract discounts are not included.",
    })
    print(json.dumps({"rows": len(rows), "queries_without_rows": len(missing)}))


if __name__ == "__main__":
    main()
