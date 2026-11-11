#!/usr/bin/env python3
"""Actual monthly cost of each AKS cluster and Container Apps environment, by resource and meter.

Uses the Cost Management query API at subscription scope, only on subscriptions that have AKS or
Container Apps. One query per subscription, cached on disk so a rerun resumes after throttling.

Default cost type is AmortizedCost: node VMs covered by reservations or savings plans show 0 in
ActualCost (the purchase is billed apart), so only amortized cost gives the effective rate per VM hour.

A cluster's cost is the cluster resource itself (tier, Defender, cost analysis add-on) plus every
resource in its node resource group (VMs, OS disks, load balancer, public IPs). A Container Apps
environment's cost is its apps, jobs and environment plus its infrastructure resource group.

Outputs (in --output-dir):
  cost_by_resource.csv  month x owner (cluster or environment) x resource x meter: cost and quantity
  cost_errors.csv
  cost_summary.json     period, subscriptions queried and every request body
"""

import json
import os
import sys
import time
from datetime import date, timedelta

import aa_common as aa

bu = aa.bu
sys.path.insert(0, os.path.join(os.path.dirname(aa.TOPIC_DIR), "backup", "src"))
from bu_backup_cost import cost_query  # noqa: E402  (same Cost Management client as backup)


def last_three_months():
    first = date.today().replace(day=1)
    end = first - timedelta(days=1)
    start = (first - timedelta(days=85)).replace(day=1)
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
        owner = {"owner_type": "AKS", "owner": row["name"], "environment": row["environment"],
                 "location": row["location"], "owner_id": row["id"].lower()}
        by_resource[row["id"].lower()] = owner
        if row.get("node_resource_group"):
            by_group[(row["subscription_id"].lower(), row["node_resource_group"].lower())] = owner
    env_owner = {}
    for row in environments:
        owner = {"owner_type": "ACA", "owner": row["name"], "environment": row["environment"],
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


def main():
    parser = bu.build_parser(__doc__.splitlines()[0], aa.results_path("aa-cost"))
    aa.add_common_arguments(parser)
    parser.set_defaults(environments="ALL")
    parser.add_argument("--from-date", default=None, help="Start date (YYYY-MM-DD). Default: first day, three months ago")
    parser.add_argument("--to-date", default=None, help="End date (YYYY-MM-DD). Default: last day of last month")
    parser.add_argument("--pause", type=float, default=2.0, help="Seconds each worker waits between queries")
    parser.add_argument("--cost-workers", type=int, default=2, help="Concurrent Cost Management queries")
    parser.add_argument("--cost-type", default="AmortizedCost", choices=["AmortizedCost", "ActualCost"],
                        help="AmortizedCost spreads reservations and savings plans over the resources that use them")
    args = parser.parse_args()
    bu.setup_logging(args.verbose)
    start, end = last_three_months()
    start, end = args.from_date or start, args.to_date or end
    environments = aa.wanted_environments(args)

    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    aca_envs = aa.read_csv(os.path.join(args.inventory, "aca_environments.csv"))
    apps = aa.read_csv(os.path.join(args.inventory, "aca_apps.csv"))
    if not clusters and not aca_envs:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv or aca_environments.csv in " + args.inventory)
    if environments is not None:
        clusters = [row for row in clusters if row["environment"] in environments]
        aca_envs = [row for row in aca_envs if row["environment"] in environments]
    by_group, by_resource = owners(clusters, aca_envs, apps)
    subscriptions = sorted({row["subscription_id"] for row in clusters + aca_envs})

    ctx = bu.AzureContext(args)
    period = {"type": args.cost_type, "timeframe": "Custom", "timePeriod": {"from": start, "to": end}}
    cache_dir = os.path.join(args.output_dir, "_cache", f"{args.cost_type}_{start}_{end}")
    os.makedirs(cache_dir, exist_ok=True)

    def fetch(sub_id):
        cache_file = os.path.join(cache_dir, f"{sub_id}.json")
        if os.path.exists(cache_file):
            with open(cache_file, encoding="utf-8") as handle:
                return json.load(handle)
        groups = sorted({group for (sub, group) in by_group if sub == sub_id.lower()})
        filters = [{"dimensions": {"name": "ResourceType", "operator": "In",
                                   "values": [aa.AKS_TYPE] + list(aa.ACA_TYPES)}}]
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
        result = ctx.try_arm("cost-by-resource", sub_id, cost_query, ctx, sub_id, body)
        if result is None:
            return None
        columns, rows = result
        payload = [dict(zip(columns, row)) for row in rows]
        with open(cache_file, "w", encoding="utf-8") as handle:
            json.dump(payload, handle)
        time.sleep(args.pause)
        return payload

    results = ctx.parallel(fetch, subscriptions, label="cost", workers=args.cost_workers)
    rows, unmatched = [], 0
    for sub_id, records in zip(subscriptions, results):
        for record in records or []:
            resource_id = str(record.get("ResourceId", "")).lower()
            parts = bu.parse_resource_id(resource_id)
            owner = by_resource.get(resource_id) or by_group.get((sub_id.lower(), parts["resource_group"]))
            if not owner:
                unmatched += 1
                continue
            month = billing_month(record)
            rows.append(dict(owner, month=month, subscription_id=sub_id, resource_group=parts.get("resource_group", ""),
                             resource_type=resource_type(resource_id), resource_name=parts.get("name", ""),
                             meter_category=record.get("MeterCategory", ""), meter=record.get("Meter", ""),
                             cost_usd=round(record.get("CostUSD", 0) or 0, 4),
                             usage_quantity=round(record.get("UsageQuantity", 0) or 0, 4)))

    output = bu.Output(args.output_dir, args.delimiter)
    output.csv("cost_by_resource.csv", rows)
    output.csv("cost_errors.csv", ctx.errors.rows, bu.ErrorLog.FIELDS)
    output.json("cost_summary.json", bu.run_metadata(ctx, {
        "period": {"from": start, "to": end},
        "cost_type": args.cost_type,
        "subscriptions_queried": len(subscriptions),
        "owners": {"AKS": len(clusters), "ACA_environments": len(aca_envs)},
        "rows_without_owner": unmatched,
    }))


if __name__ == "__main__":
    main()
