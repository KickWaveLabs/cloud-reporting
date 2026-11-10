#!/usr/bin/env python3
"""Actual monthly cost per subscription by resource type and meter, plus backup vault cost per vault.

Uses the Cost Management query API at subscription scope (no management group access needed).
The API throttles, so subscriptions are queried sequentially with retries on 429.

Outputs:
  cost_by_meter.csv            subscription x resource type x meter category x meter (all spend)
  cost_by_vault.csv            Recovery Services / Backup vault cost by vault and meter
  cost_by_storage_account.csv  storage account cost by meter, for the subscriptions with the most storage spend
  cost_errors.csv
  cost_summary.json            run metadata and evidence (every request body and call pattern)

The query API returns the same ActualCost data shown in the portal under Cost Management > Cost analysis
(actual cost, with the contract discounts applied).
"""

import json
import os
import time
from datetime import date, timedelta

from collections import Counter

import requests

import bu_common as common

API = "2023-11-01"
VAULT_TYPES = ["microsoft.recoveryservices/vaults", "microsoft.dataprotection/backupvaults"]


def previous_month():
    first = date.today().replace(day=1)
    last = first - timedelta(days=1)
    return last.replace(day=1).isoformat(), last.isoformat()


def cost_query(ctx, subscription_id, body):
    """POST a Cost Management query and follow nextLink. Returns (columns, rows)."""
    url = f"{common.ARM_ENDPOINT}/subscriptions/{subscription_id}/providers/Microsoft.CostManagement/query"
    columns, rows = [], []
    params = {"api-version": API}
    grouping = "+".join(group["name"] for group in body["dataset"].get("grouping", []))
    ctx.evidence.add_body(f"Cost Management query by {grouping}", url, API, body)
    while url:
        for attempt in range(1, 8):
            headers = {"Authorization": f"Bearer {ctx._bearer()}", "Content-Type": "application/json"}
            try:
                response = ctx._session().post(url, params=params, json=body, headers=headers, timeout=180)
            except requests.RequestException as error:
                if attempt == 7:
                    raise common.ArmError("", "ConnectionError", str(error), url) from error
                time.sleep(10 * attempt)
                continue
            ctx.evidence.add_call("POST", url, API, response.status_code)
            if response.status_code == 429 or response.status_code >= 500:
                retry = next((value for key, value in response.headers.items()
                              if key.lower().endswith("retry-after") and value.isdigit()), None)
                time.sleep(int(retry) if retry else 10 * attempt)
                continue
            break
        if response.status_code >= 400:
            error = response.json().get("error", {}) if response.content else {}
            raise common.ArmError(response.status_code, error.get("code", ""), error.get("message", ""), url)
        properties = response.json().get("properties", {})
        columns = [column["name"] for column in properties.get("columns", [])]
        rows.extend(properties.get("rows", []))
        url = properties.get("nextLink")
        params = None
    return columns, rows


def main():
    parser = common.build_parser(__doc__.splitlines()[0], common.results_path("backup-cost"))
    parser.add_argument("--from-date", default=None, help="Start date (YYYY-MM-DD). Default: first day of last month")
    parser.add_argument("--to-date", default=None, help="End date (YYYY-MM-DD). Default: last day of last month")
    parser.add_argument("--pause", type=float, default=1.0, help="Seconds each worker waits between queries")
    parser.add_argument("--cost-workers", type=int, default=4, help="Concurrent Cost Management queries")
    parser.add_argument("--storage-top", type=int, default=10,
                        help="Subscriptions (by storage spend) to break down per storage account. 0 disables it")
    args = parser.parse_args()
    common.setup_logging(args.verbose)
    start, end = previous_month()
    start, end = args.from_date or start, args.to_date or end

    ctx = common.AzureContext(args)
    ctx.load_subscriptions()
    vault_rows = ctx.query(
        "resources | where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')"
        " | distinct subscriptionId",
        partition=False, label="vault_subscriptions")
    vault_subscriptions = {common.normalize(row["subscriptionId"]) for row in vault_rows}

    period = {"type": "ActualCost", "timeframe": "Custom", "timePeriod": {"from": start, "to": end}}
    cache_dir = os.path.join(args.output_dir, "_cache", f"{start}_{end}")
    os.makedirs(cache_dir, exist_ok=True)

    def fetch(subscription):
        """Query one subscription; results are cached on disk so a rerun resumes."""
        sub_id = subscription["id"]
        cache_file = os.path.join(cache_dir, f"{sub_id}.json")
        if os.path.exists(cache_file):
            with open(cache_file, encoding="utf-8") as handle:
                return json.load(handle)
        meters, vaults = [], []
        body = dict(period, dataset={
            "granularity": "None",
            "aggregation": {"cost": {"name": "CostUSD", "function": "Sum"},
                            "quantity": {"name": "UsageQuantity", "function": "Sum"}},
            "grouping": [{"type": "Dimension", "name": "ResourceType"},
                         {"type": "Dimension", "name": "MeterCategory"},
                         {"type": "Dimension", "name": "Meter"}],
        })
        result = ctx.try_arm("cost-by-meter", sub_id, cost_query, ctx, sub_id, body)
        if result is None:
            return None
        columns, rows = result
        for row in rows:
            record = dict(zip(columns, row))
            meters.append({
                "subscription": subscription["name"], "subscription_id": sub_id,
                "resource_type": record.get("ResourceType", ""), "meter_category": record.get("MeterCategory", ""),
                "meter": record.get("Meter", ""), "cost_usd": round(record.get("CostUSD", 0) or 0, 4),
                "usage_quantity": round(record.get("UsageQuantity", 0) or 0, 4),
            })
        if common.normalize(sub_id) in vault_subscriptions:
            time.sleep(args.pause)
            body = dict(period, dataset={
                "granularity": "None",
                "aggregation": {"cost": {"name": "CostUSD", "function": "Sum"},
                                "quantity": {"name": "UsageQuantity", "function": "Sum"}},
                "grouping": [{"type": "Dimension", "name": "ResourceId"}, {"type": "Dimension", "name": "Meter"}],
                "filter": {"dimensions": {"name": "ResourceType", "operator": "In", "values": VAULT_TYPES}},
            })
            result = ctx.try_arm("cost-by-vault", sub_id, cost_query, ctx, sub_id, body)
            if result is None:
                return None
            columns, rows = result
            for row in rows:
                record = dict(zip(columns, row))
                vaults.append({
                    "subscription": subscription["name"], "subscription_id": sub_id,
                    "vault_id": record.get("ResourceId", ""), "meter": record.get("Meter", ""),
                    "cost_usd": round(record.get("CostUSD", 0) or 0, 4),
                    "usage_quantity": round(record.get("UsageQuantity", 0) or 0, 4),
                })
        payload = {"meters": meters, "vaults": vaults}
        with open(cache_file, "w", encoding="utf-8") as handle:
            json.dump(payload, handle)
        time.sleep(args.pause)
        return payload

    subscriptions = sorted(ctx.subscriptions.values(), key=lambda value: value["name"].lower())
    results = ctx.parallel(fetch, subscriptions, label="cost", workers=args.cost_workers)
    meter_rows = [row for result in results if result for row in result["meters"]]
    vault_cost_rows = [row for result in results if result for row in result["vaults"]]

    storage_spend = Counter()
    for row in meter_rows:
        if row["resource_type"].lower() == "microsoft.storage/storageaccounts":
            storage_spend[row["subscription_id"]] += row["cost_usd"]
    account_rows = []
    for sub_id, _ in storage_spend.most_common(args.storage_top):
        body = dict(period, dataset={
            "granularity": "None",
            "aggregation": {"cost": {"name": "CostUSD", "function": "Sum"},
                            "quantity": {"name": "UsageQuantity", "function": "Sum"}},
            "grouping": [{"type": "Dimension", "name": "ResourceId"}, {"type": "Dimension", "name": "Meter"}],
            "filter": {"dimensions": {"name": "ResourceType", "operator": "In", "values": ["microsoft.storage/storageaccounts"]}},
        })
        result = ctx.try_arm("cost-by-storage-account", sub_id, cost_query, ctx, sub_id, body)
        if result:
            columns, rows = result
            for row in rows:
                record = dict(zip(columns, row))
                account_rows.append({
                    "subscription": ctx.subscription_name(sub_id), "subscription_id": sub_id,
                    "resource_id": record.get("ResourceId", ""), "meter": record.get("Meter", ""),
                    "cost_usd": round(record.get("CostUSD", 0) or 0, 4),
                    "usage_quantity": round(record.get("UsageQuantity", 0) or 0, 4),
                })
        time.sleep(args.pause)

    output = common.Output(args.output_dir, args.delimiter)
    output.csv("cost_by_meter.csv", meter_rows)
    output.csv("cost_by_vault.csv", vault_cost_rows)
    output.csv("cost_by_storage_account.csv", account_rows)
    output.csv("cost_errors.csv", ctx.errors.rows, common.ErrorLog.FIELDS)
    output.json("cost_summary.json", common.run_metadata(ctx, {"period": {"from": start, "to": end}}))


if __name__ == "__main__":
    main()
