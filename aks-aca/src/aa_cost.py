#!/usr/bin/env python3
"""Actual monthly cost of each AKS cluster, Container Apps environment and App Service plan, by resource and meter.

Uses the Cost Management query API at subscription scope, only on subscriptions that have AKS, Container
Apps or App Service. One query per subscription and family (AKS/ACA, App Service), cached on disk so a
rerun resumes after throttling. App Service and Functions cover PROD only by default
(--app-service-environments), since the deck only shows the PROD estate.

Default cost type is AmortizedCost: node VMs covered by reservations or savings plans show 0 in
ActualCost (the purchase is billed apart), so only amortized cost gives the effective rate per VM hour.

A cluster's cost is the cluster resource itself (tier, Defender, cost analysis add-on) plus every
resource in its node resource group (VMs, OS disks, load balancer, public IPs). A Container Apps
environment's cost is its apps, jobs and environment plus its infrastructure resource group. An App
Service plan's cost is the plan plus its sites and slots; owner_type is the plan's runtime from the
inventory (APP_SERVICE, or FUNCTIONS for Functions-only plans). A site without a plan in the inventory
counts under its own type.

Outputs (in --output-dir):
  cost_by_resource.csv  month x owner (cluster, environment or plan) x resource x meter: cost and quantity
  cost_errors.csv
  cost_summary.json     period, subscriptions queried and every request body
"""

import json
import os
import time

import aa_common as aa

azure = aa.azure

WEB_TYPES = ("microsoft.web/serverfarms", "microsoft.web/sites", "microsoft.web/sites/slots")


def last_month():
    start, end = aa.last_full_month()
    return start.isoformat(), end.isoformat()


def resource_type(resource_id):
    """microsoft.compute/virtualmachinescalesets from an ARM id (last provider segment, child types joined)."""
    parts = resource_id.strip("/").split("/")
    if "providers" not in parts:
        return ""
    index = len(parts) - 1 - parts[::-1].index("providers")
    tail = parts[index + 1:]
    return "/".join([tail[0]] + tail[1::2]) if tail else ""


def billing_month(record):
    """Month (YYYY-MM) of a Monthly row: BillingMonth is a date string, UsageDate an int like 20261001."""
    value = record.get("BillingMonth") or record.get("UsageDate") or ""
    text = str(value)
    return f"{text[:4]}-{text[4:6]}" if text.isdigit() else text[:7]


def owners(clusters, environments, apps):
    """Map resource groups and resource ids to the cluster or environment that owns their cost."""
    by_group, by_resource = {}, {}
    for row in clusters:
        owner = {"owner_type": "AKS", "owner": row["name"], "cod_app": row.get("cod_app", ""),
                 "environment": row["environment"],
                 "location": row["location"], "owner_id": row["id"].lower()}
        by_resource[row["id"].lower()] = owner
        if row.get("node_resource_group"):
            by_group[(row["subscription_id"].lower(), row["node_resource_group"].lower())] = owner
    env_owner = {}
    for row in environments:
        owner = {"owner_type": "ACA", "owner": row["name"], "cod_app": row.get("cod_app", ""),
                 "environment": row["environment"],
                 "location": row["location"], "owner_id": row["id"].lower()}
        env_owner[row["name"].lower()] = owner
        by_resource[row["id"].lower()] = owner
        if row.get("infrastructure_resource_group"):
            by_group[(row["subscription_id"].lower(), row["infrastructure_resource_group"].lower())] = owner
    for row in apps:
        owner = env_owner.get(str(row.get("managed_environment", "")).lower())
        if owner:
            by_resource[row["id"].lower()] = owner
    return by_group, by_resource


def web_owners(plans, sites):
    """Map App Service plans, sites and slots to the plan that owns their cost."""
    by_resource = {}
    for row in plans:
        by_resource[row["id"].lower()] = {"owner_type": row.get("runtime") or "APP_SERVICE", "owner": row["name"],
                                          "cod_app": row.get("cod_app", ""), "environment": row["environment"],
                                          "location": row["location"], "owner_id": row["id"].lower()}
    for row in sites:
        owner = by_resource.get(str(row.get("plan_id", "")).lower()) or {
            "owner_type": "FUNCTIONS" if row.get("type") == "FUNCTION" else "APP_SERVICE", "owner": row["name"],
            "cod_app": row.get("cod_app", ""), "environment": row["environment"], "location": row["location"],
            "owner_id": row["id"].lower()}
        by_resource[row["id"].lower()] = owner
    return by_resource


def site_of(resource_id):
    """Site id of a slot id (.../sites/<name>/slots/<slot> -> .../sites/<name>); other ids unchanged."""
    marker = resource_id.find("/slots/")
    return resource_id[:marker] if "/providers/microsoft.web/sites/" in resource_id and marker > 0 else resource_id


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-cost"))
    aa.add_common_arguments(parser)
    parser.set_defaults(environments="ALL")
    parser.add_argument("--from-date", default=None, help="Start date (YYYY-MM-DD). Default: first day of last month")
    parser.add_argument("--to-date", default=None, help="End date (YYYY-MM-DD). Default: last day of last month")
    parser.add_argument("--pause", type=float, default=2.0, help="Seconds each worker waits between queries")
    parser.add_argument("--cost-workers", type=int, default=2, help="Concurrent Cost Management queries")
    parser.add_argument("--cost-type", default="AmortizedCost", choices=["AmortizedCost", "ActualCost"],
                        help="AmortizedCost spreads reservations and savings plans over the resources that use them")
    parser.add_argument("--app-service-environments", default="PROD",
                        help="Environments whose App Service and Functions plans are costed ('ALL', or 'NONE' to skip)")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)
    start, end = last_month()
    start, end = args.from_date or start, args.to_date or end
    environments = aa.wanted_environments(args)
    web_value = args.app_service_environments.strip().upper()
    web_environments = None if web_value == "ALL" else (
        set() if web_value == "NONE" else {item.strip() for item in web_value.split(",") if item.strip()})

    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    aca_envs = aa.read_csv(os.path.join(args.inventory, "aca_environments.csv"))
    apps = aa.read_csv(os.path.join(args.inventory, "aca_apps.csv"))
    plans = aa.read_csv(os.path.join(args.inventory, "app_service_plans.csv"))
    sites = aa.read_csv(os.path.join(args.inventory, "app_services.csv"))
    if not clusters and not aca_envs:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv or aca_environments.csv in " + args.inventory)
    if web_value != "NONE" and (not plans or "runtime" not in plans[0] or (sites and "plan_id" not in sites[0])):
        raise SystemExit("App Service costs need a newer inventory: run aa_inventory.py again, "
                         "or pass --app-service-environments NONE to skip them.")
    if environments is not None:
        clusters = [row for row in clusters if row["environment"] in environments]
        aca_envs = [row for row in aca_envs if row["environment"] in environments]
    if web_environments is not None:
        plans = [row for row in plans if row["environment"] in web_environments]
        sites = [row for row in sites if row["environment"] in web_environments]
    by_group, by_resource = owners(clusters, aca_envs, apps)
    web_by_resource = web_owners(plans, sites)
    families = {
        "containers": {"subscriptions": sorted({row["subscription_id"] for row in clusters + aca_envs}),
                       "types": [aa.AKS_TYPE] + list(aa.ACA_TYPES), "cache": f"{args.cost_type}_{start}_{end}"},
        "app_service": {"subscriptions": sorted({row["subscription_id"] for row in plans + sites}),
                        "types": list(WEB_TYPES), "cache": f"web_{args.cost_type}_{start}_{end}"},
    }

    ctx = azure.AzureContext(args)
    period = {"type": args.cost_type, "timeframe": "Custom", "timePeriod": {"from": start, "to": end}}

    def fetch(family, sub_id):
        cache_dir = os.path.join(args.output_dir, "_cache", families[family]["cache"])
        os.makedirs(cache_dir, exist_ok=True)
        cache_file = os.path.join(cache_dir, f"{sub_id}.json")
        if os.path.exists(cache_file):
            with open(cache_file, encoding="utf-8") as handle:
                return json.load(handle)
        filters = [{"dimensions": {"name": "ResourceType", "operator": "In", "values": families[family]["types"]}}]
        if family == "containers":
            groups = sorted({group for (sub, group) in by_group if sub == sub_id.lower()})
            if groups:
                filters.append({"dimensions": {"name": "ResourceGroupName", "operator": "In", "values": groups}})
        body = dict(period, dataset={
            "granularity": "Monthly",
            "aggregation": {"cost": {"name": "CostUSD", "function": "Sum"},
                            "quantity": {"name": "UsageQuantity", "function": "Sum"}},
            "grouping": [{"type": "Dimension", "name": "ResourceId"},
                         {"type": "Dimension", "name": "MeterCategory"},
                         {"type": "Dimension", "name": "Meter"}],
            "filter": {"or": filters} if len(filters) > 1 else filters[0],
        })
        result = ctx.try_arm(f"cost-by-resource-{family}", sub_id, azure.cost_query, ctx, sub_id, body)
        if result is None:
            return None
        columns, rows = result
        payload = [dict(zip(columns, row)) for row in rows]
        with open(cache_file, "w", encoding="utf-8") as handle:
            json.dump(payload, handle)
        time.sleep(args.pause)
        return payload

    jobs = [(family, sub_id) for family in families for sub_id in families[family]["subscriptions"]]
    results = ctx.parallel(lambda job: fetch(*job), jobs, label="cost", workers=args.cost_workers)
    rows, unmatched = [], {"containers": 0, "app_service": 0}
    for (family, sub_id), records in zip(jobs, results):
        for record in records or []:
            resource_id = str(record.get("ResourceId", "")).lower()
            parts = azure.parse_resource_id(resource_id)
            if family == "containers":
                owner = by_resource.get(resource_id) or by_group.get((sub_id.lower(), parts["resource_group"]))
            else:
                # Plans and sites of environments out of scope share the subscription: they stay unmatched.
                owner = web_by_resource.get(resource_id) or web_by_resource.get(site_of(resource_id))
            if not owner:
                unmatched[family] += 1
                continue
            month = billing_month(record)
            rows.append(dict(owner, month=month, subscription_id=sub_id, resource_group=parts.get("resource_group", ""),
                             resource_type=resource_type(resource_id), resource_name=parts.get("name", ""),
                             meter_category=record.get("MeterCategory", ""), meter=record.get("Meter", ""),
                             cost_usd=round(record.get("CostUSD", 0) or 0, 4),
                             usage_quantity=round(record.get("UsageQuantity", 0) or 0, 4)))

    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("cost_by_resource.csv", rows)
    output.csv("cost_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    output.json("cost_summary.json", azure.run_metadata(ctx, {
        "period": {"from": start, "to": end},
        "cost_type": args.cost_type,
        "subscriptions_queried": len(families["containers"]["subscriptions"]),
        "app_service_subscriptions_queried": len(families["app_service"]["subscriptions"]),
        "app_service_environments": sorted(web_environments) if web_environments is not None else "ALL",
        "owners": {"AKS": len(clusters), "ACA_environments": len(aca_envs),
                   "APP_SERVICE_plans": sum((row.get("runtime") or "APP_SERVICE") == "APP_SERVICE" for row in plans),
                   "FUNCTIONS_plans": sum(row.get("runtime") == "FUNCTIONS" for row in plans)},
        "rows_without_owner": unmatched["containers"],
        "app_service_rows_out_of_scope": unmatched["app_service"],
    }))


if __name__ == "__main__":
    main()
