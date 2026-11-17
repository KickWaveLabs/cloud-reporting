#!/usr/bin/env python3
"""Microservices running in each AKS cluster on sample days, from Container Insights, kept as a growing history.

A sample is the 1st and the 15th of each month, at 10:00-11:00 Lima time (15:00-16:00 UTC): a microservice
counts when at least one of its pods was seen in that hour. A microservice is a ReplicaSet workload outside
platform namespaces and not a known tool (aa_common.is_microservice), the same rule as the rest of aks-aca.

The history lives in --output-dir (results/aa-history by default) and is never cleared: each run adds the
sample days that are not there yet (or every day with --full) and keeps the rest. Container Insights keeps a
limited window, so run this at least once a month and the history keeps growing after the workspace has
dropped the old data. Days older than the workspace retention come back empty and are not written.

Run aa_inventory.py first. Clusters are read from the inventory of today, so a cluster deleted before today
has no new samples, but its rows from earlier runs stay.

Outputs (in --output-dir):
  microservice_history.csv   sample date x cluster: cluster_id, cod_app, microservices, namespaces
  history_summary.json       sample days queried, clusters, and microservices per day (PROD, all categories)
"""

import os
import time
from collections import defaultdict
from datetime import date, datetime, timedelta

import aa_common as aa
import aa_container_insights as ci

azure = aa.azure

HISTORY_FILE = "microservice_history.csv"
SAMPLE_DAYS = (1, 15)
SAMPLE_HOUR_UTC = 15  # 10:00 Lima

SAMPLE_QUERY = """
KubePodInventory
| where {windows}
| where ControllerKind == 'ReplicaSet' and isnotempty(ControllerName)
| extend Workload = {workload}
| summarize by Day = format_datetime(startofday(TimeGenerated), 'yyyy-MM-dd'), Namespace, Workload
"""


def sample_days(start, end):
    """The 1st and the 15th of every month between start and end, inclusive."""
    days, current = [], start.replace(day=1)
    while current <= end:
        days.extend(current.replace(day=day) for day in SAMPLE_DAYS if start <= current.replace(day=day) <= end)
        current = (current + timedelta(days=32)).replace(day=1)
    return days


def windows(days):
    return " or ".join(f"TimeGenerated between (datetime({day.isoformat()}T{SAMPLE_HOUR_UTC:02d}:00:00Z) .. "
                       f"datetime({day.isoformat()}T{SAMPLE_HOUR_UTC + 1:02d}:00:00Z))" for day in days)


def cluster_samples(ctx, client, cluster, days):
    """Microservices and namespaces per sample day of one cluster; None when the query fails."""
    if not days:
        return []
    kql = SAMPLE_QUERY.format(windows=windows(days), workload=ci.WORKLOAD_NAME)
    timespan = f"{days[0].isoformat()}T00:00:00Z/{(days[-1] + timedelta(days=1)).isoformat()}T00:00:00Z"
    rows = ctx.try_arm("history", cluster["name"], client.query, cluster["id"], kql, timespan)
    if rows is None:
        return None
    workloads, namespaces = defaultdict(set), defaultdict(set)
    for row in rows:
        if aa.is_microservice(row["Namespace"], row["Workload"]):
            workloads[row["Day"]].add((row["Namespace"], row["Workload"]))
            namespaces[row["Day"]].add(row["Namespace"])
    return [{"date": day, "cluster": cluster["name"], "cluster_id": cluster["id"].lower(),
             "cod_app": cluster.get("cod_app", ""), "environment": cluster["environment"],
             "microservices": len(items), "namespaces": len(namespaces[day])}
            for day, items in sorted(workloads.items())]


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-history"))
    aa.add_common_arguments(parser)
    parser.add_argument("--start", default=None,
                        help="First sample day (YYYY-MM-DD). Default: one year back; older days than the workspace "
                             "keeps come back empty")
    parser.add_argument("--end", default=None, help="Last sample day (YYYY-MM-DD). Default: yesterday (Lima)")
    parser.add_argument("--clusters", default="", help="Comma-separated cluster names. Default: clusters in --environments")
    parser.add_argument("--full", action="store_true", help="Query every sample day again, not only the missing ones")
    parser.add_argument("--pause", type=float, default=2.0, help="Seconds between Log Analytics queries")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)

    today = datetime.now(aa.LIMA).date()
    end = date.fromisoformat(args.end) if args.end else today - timedelta(days=1)
    start = date.fromisoformat(args.start) if args.start else end.replace(day=1) - timedelta(days=365)
    days = sample_days(start, end)
    environments = aa.wanted_environments(args)
    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    if not clusters:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv in " + args.inventory)
    wanted = {name.strip().lower() for name in args.clusters.split(",") if name.strip()}
    clusters = [row for row in clusters if row.get("container_insights", "").lower() == "true"
                and ((wanted and row["name"].lower() in wanted)
                     or (not wanted and (environments is None or row["environment"] in environments)))]

    history = aa.read_csv(os.path.join(args.output_dir, HISTORY_FILE))
    done = defaultdict(set)
    for row in history:
        done[row["cluster_id"]].add(row["date"])
    ctx = azure.AzureContext(args)
    client = ci.LogsClient(ctx, args.pause)
    new_rows, started = [], time.monotonic()
    for number, cluster in enumerate(clusters, 1):
        # A day already in the history is skipped unless --full; the last sample day is always refreshed.
        pending = [day for day in days if args.full or day.isoformat() not in done[cluster["id"].lower()]
                   or day == days[-1]]
        azure.LOG.info("History %d/%d %s: %d days (%.0f min)", number, len(clusters), cluster["name"], len(pending),
                       (time.monotonic() - started) / 60)
        new_rows.extend(cluster_samples(ctx, client, cluster, pending) or [])

    replaced = {(row["cluster_id"], row["date"]) for row in new_rows}
    merged = [row for row in history if (row["cluster_id"], row["date"]) not in replaced] + new_rows
    merged.sort(key=lambda row: (row["date"], row["cluster"]))
    ctx.evidence.queries.append({"source": "Log Analytics (Container Insights)", "label": "microservice-history",
                                 "query": " ".join(SAMPLE_QUERY.replace("{workload}", ci.WORKLOAD_NAME).split()),
                                 "rows": len(new_rows), "subscriptions": "", "seconds": "",
                                 "timestamp": azure.iso(azure.utc_now())})
    per_day = defaultdict(int)
    for row in merged:
        if row["environment"] == "PROD":
            per_day[row["date"]] += int(row["microservices"])
    output = azure.Output(args.output_dir, args.delimiter)
    output.csv(HISTORY_FILE, merged)
    output.csv("history_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    output.json("history_summary.json", azure.run_metadata(ctx, {
        "sample_days": [day.isoformat() for day in days], "clusters_queried": len(clusters),
        "new_rows": len(new_rows), "prod_microservices_per_day_all_categories": dict(sorted(per_day.items())),
    }))


if __name__ == "__main__":
    main()
