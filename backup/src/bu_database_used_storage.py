#!/usr/bin/env python3
"""Used and allocated storage of each Azure SQL database, from Azure Monitor metrics.

database_inventory.csv only has max_size_gb (the size limit). This script reads that inventory and adds
the space actually used, which is the protected size a backup vendor would quote against.

Outputs:
  database_used_storage.csv       one row per Azure SQL database: used, allocated and max size (GB)
  database_used_storage_errors.csv
  database_used_storage_summary.json  totals and evidence (calls executed)

Dependencies:
  pip install azure-identity azure-mgmt-resourcegraph requests
"""

import csv
import os
from datetime import timedelta

import bu_common as common

METRICS_API = "2023-10-01"
METRICS = ("storage", "allocated_data_storage")
GB = 1024 ** 3


def main():
    parser = common.build_parser(__doc__.splitlines()[0], common.results_path("backup-database"))
    parser.add_argument("--database-inventory",
                        default=common.results_path("backup-database", "database_inventory.csv"),
                        help="database_inventory.csv written by bu_database_inventory.py")
    parser.add_argument("--ltr-only", action="store_true", help="Only databases with long-term retention")
    args = parser.parse_args()
    common.setup_logging(args.verbose)
    ctx = common.AzureContext(args)
    ctx.load_subscriptions()

    with open(args.database_inventory, encoding="utf-8-sig", newline="") as handle:
        databases = [row for row in csv.DictReader(handle)
                     if row["service_type"] == "AZURE_SQL_DB"
                     and common.normalize(row["subscription_id"]) in ctx.subscriptions
                     and (not args.ltr_only or row["ltr_configured"] == "YES")]
    common.LOG.info("Azure SQL databases in scope: %d", len(databases))

    end = common.utc_now()
    timespan = f"{common.iso(end - timedelta(hours=6))}/{common.iso(end)}"

    def fetch(row):
        body = ctx.try_arm("sql_used_storage", row["id"], ctx.arm_get,
                           f"{row['id']}/providers/Microsoft.Insights/metrics", METRICS_API,
                           {"metricnames": ",".join(METRICS), "aggregation": "Maximum",
                            "interval": "PT1H", "timespan": timespan}) or {}
        latest = {}
        for metric in body.get("value", []):
            points = [point.get("maximum") for series in metric.get("timeseries", [])
                      for point in series.get("data", []) if point.get("maximum") is not None]
            if points:
                latest[metric["name"]["value"]] = points[-1]
        return latest

    rows = []
    for row, latest in zip(databases, ctx.parallel(fetch, databases, label="sql_used_storage")):
        latest = latest or {}
        rows.append({
            "subscription": row["subscription"], "resource_group": row["resource_group"],
            "server_name": row["server_name"], "database_name": row["database_name"],
            "environment": row["environment"], "sku_tier": row["sku_tier"],
            "ltr_configured": row["ltr_configured"],
            "used_gb": round(latest["storage"] / GB, 2) if "storage" in latest else "",
            "allocated_gb": (round(latest["allocated_data_storage"] / GB, 2)
                             if "allocated_data_storage" in latest else ""),
            "max_size_gb": row["max_size_gb"], "id": row["id"],
        })

    def total_tb(selected, key):
        return round(sum(float(row[key]) for row in selected if row[key] not in ("", None)) / 1024, 2)

    ltr = [row for row in rows if row["ltr_configured"] == "YES"]
    totals = {
        "databases": len(rows),
        "with_used_metric": sum(1 for row in rows if row["used_gb"] != ""),
        "used_tb": total_tb(rows, "used_gb"),
        "max_size_tb": total_tb(rows, "max_size_gb"),
        "ltr_databases": len(ltr),
        "ltr_used_tb": total_tb(ltr, "used_gb"),
        "ltr_max_size_tb": total_tb(ltr, "max_size_gb"),
        "metric_window": timespan,
    }

    os.makedirs(args.output_dir, exist_ok=True)
    output = common.Output(args.output_dir, args.delimiter)
    output.csv("database_used_storage.csv", rows)
    output.csv("database_used_storage_errors.csv", ctx.errors.rows, common.ErrorLog.FIELDS)
    output.json("database_used_storage_summary.json", common.run_metadata(ctx, {"totals": totals}))
    common.LOG.info("Used: %s TB of %s TB max size (LTR: %s TB of %s TB)", totals["used_tb"],
                    totals["max_size_tb"], totals["ltr_used_tb"], totals["ltr_max_size_tb"])


if __name__ == "__main__":
    main()
