#!/usr/bin/env python3
"""Capacity and latency signals: unschedulable pods, OOMKilled containers and APIM backend latency.

Three read-only checks that test the usage figures from the other side: low CPU use is only waste if the
clusters are not short of capacity (unschedulable pods) and the containers are not short of memory
(OOMKilled). Backend latency per API tells how much a Container Apps cold start would add.

  AKS platform metric   cluster_autoscaler_unschedulable_pods_count (Average, hourly), per cluster.
                        Only emitted when the cluster autoscaler is on; other clusters return no points.
  Container Insights    KubePodInventory rows whose ContainerLastStatus is OOMKilled. One OOM kill = one
                        (pod, container, restart count) with that last status, so a status repeated in
                        every inventory snapshot counts once. A kill from before the window counts in its
                        first month. Only counts leave the workspace: no logs, labels or messages.
  APIM platform metrics BackendDuration and Duration (milliseconds, Average and Maximum) per API and month,
                        one data point per month (interval FULL), so the average is per request.

Run aa_inventory.py first. The window matches aa_metrics.py: the last 92 days up to yesterday (Lima).

Outputs (in --output-dir):
  aks_autoscaler_monthly.csv  cluster x month: hours with data, hours with unschedulable pods (all, office),
                              highest hourly average of unschedulable pods
  aks_oomkilled_monthly.csv   cluster x namespace x workload x month: OOM kills and pods affected
  apim_latency_monthly.csv    APIM x API x month: backend and gateway duration, average and maximum (ms)
  signals_errors.csv
  signals_summary.json
"""

import json
import os
import time

import aa_common as aa
import aa_container_insights as ci

azure = aa.azure

AUTOSCALER_METRIC = "cluster_autoscaler_unschedulable_pods_count"
LATENCY_METRICS = ["BackendDuration", "Duration"]

OOM_QUERY = """
KubePodInventory
| where TimeGenerated between (datetime({start}) .. datetime({end}))
| where ContainerLastStatus has 'OOMKilled' and isnotempty(ControllerName)
| extend Workload = {workload}
| summarize FirstSeen = min(TimeGenerated) by PodUid, ContainerName, ContainerRestartCount, Namespace, Workload,
    ControllerKind
| extend Month = format_datetime(datetime_utc_to_local(FirstSeen, 'America/Lima'), 'yyyy-MM')
| summarize OomKills = count(), Pods = dcount(PodUid) by Month, Namespace, Workload, ControllerKind
"""


def autoscaler_rows(ctx, cluster, start, end, chunk_days):
    hours = {}
    for begin, finish in aa.chunks(start, end, chunk_days):
        values = ctx.try_arm("aks-autoscaler", cluster["name"], aa.get_metrics, ctx, cluster["id"],
                             [AUTOSCALER_METRIC], begin, finish, "PT1H", "Average")
        for metric in values or []:
            for series in metric.get("timeseries", []):
                hours.update(aa.series_points(series, "average"))
    months = {}
    for moment, value in hours.items():
        entry = months.setdefault(aa.lima_month(moment), {"hours": 0, "unschedulable": 0, "office": 0, "max": 0.0})
        entry["hours"] += 1
        if value > 0:
            entry["unschedulable"] += 1
            entry["office"] += aa.is_office(moment)
            entry["max"] = max(entry["max"], value)
    return [{"cluster": cluster["name"], "environment": cluster["environment"], "month": month,
             "hours_with_data": entry["hours"], "hours_with_unschedulable": entry["unschedulable"],
             "office_hours_with_unschedulable": entry["office"], "unschedulable_pods_max": round(entry["max"], 2)}
            for month, entry in sorted(months.items())]


def oom_rows(ctx, client, cluster, start, end):
    kql = OOM_QUERY.format(start=aa.iso_z(start), end=aa.iso_z(end), workload=ci.WORKLOAD_NAME)
    rows = ctx.try_arm("ci-oomkilled", cluster["name"], client.query, cluster["id"], kql,
                       f"{aa.iso_z(start)}/{aa.iso_z(end)}")
    if rows is None:
        return None
    return [{"cluster": cluster["name"], "environment": cluster["environment"], "month": row["Month"],
             "namespace": row["Namespace"], "workload": row["Workload"], "kind": row["ControllerKind"],
             "system_namespace": "YES" if row["Namespace"] in aa.SYSTEM_NAMESPACES
                                 or str(row["Namespace"]).startswith("aks-istio") else "NO",
             "oom_kills": row["OomKills"], "pods": row["Pods"]} for row in rows]


def latency_rows(ctx, service, start, end, top):
    out = {}
    for month, begin, finish in ci.month_windows(start, end):
        values = ctx.try_arm("apim-latency", service["name"], aa.get_metrics, ctx, service["id"], LATENCY_METRICS,
                             begin, finish, "FULL", "Average,Maximum", "ApiId eq '*'", top=top)
        for metric in values or []:
            prefix = "backend_duration" if metric["name"]["value"] == "BackendDuration" else "duration"
            for series in metric.get("timeseries", []):
                api_id = aa.series_dimensions(series).get("apiid", "")
                row = out.setdefault((month, api_id), {"apim": service["name"], "api_id": api_id, "month": month})
                for aggregation, suffix in (("average", "avg_ms"), ("maximum", "max_ms")):
                    points = [value for _, value in aa.series_points(series, aggregation)]
                    if points:
                        row[f"{prefix}_{suffix}"] = round(points[0], 1)
    # An API without requests in the month comes back as 0 ms: leave it out.
    return [row for _, row in sorted(out.items()) if row.get("duration_max_ms")]


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-signals"))
    aa.add_common_arguments(parser)
    parser.add_argument("--start", default=None, help="First day (YYYY-MM-DD, Lima). Default: 92 days before --end")
    parser.add_argument("--end", default=None, help="Last day (YYYY-MM-DD, Lima). Default: yesterday")
    parser.add_argument("--clusters", default="", help="Comma-separated cluster names. Default: clusters in --environments")
    parser.add_argument("--apim-names", default="", help="Comma-separated APIM name fragments. Default: APIM in --environments")
    parser.add_argument("--skip-autoscaler", action="store_true")
    parser.add_argument("--skip-oom", action="store_true")
    parser.add_argument("--skip-apim", action="store_true")
    parser.add_argument("--chunk-days", type=int, default=31, help="Days per autoscaler metrics call (hourly points)")
    parser.add_argument("--top", type=int, default=1000, help="Max APIs per APIM metrics call")
    parser.add_argument("--pause", type=float, default=2.0, help="Seconds between Log Analytics queries")
    parser.set_defaults(arm_rps=2.0, workers=2)
    args = parser.parse_args()
    azure.setup_logging(args.verbose)
    start, end = aa.metrics_window(args.start, args.end)
    environments = aa.wanted_environments(args)

    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    services = aa.read_csv(os.path.join(args.inventory, "apim_services.csv"))
    if not clusters and not services:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv or apim_services.csv in " + args.inventory)
    wanted = {name.strip().lower() for name in args.clusters.split(",") if name.strip()}
    clusters = [row for row in clusters
                if (wanted and row["name"].lower() in wanted)
                or (not wanted and (environments is None or row["environment"] in environments))]
    names = [name.strip().lower() for name in args.apim_names.split(",") if name.strip()]
    services = [row for row in services
                if (names and any(name in row["name"].lower() for name in names))
                or (not names and (environments is None or row["environment"] in environments))]

    ctx = azure.AzureContext(args)
    azure.LOG.info("Window %s to %s (UTC); %d clusters, %d APIM", aa.iso_z(start), aa.iso_z(end), len(clusters), len(services))
    output = azure.Output(args.output_dir, args.delimiter)

    autoscaler = []
    if not args.skip_autoscaler:
        for result in ctx.parallel(lambda cluster: autoscaler_rows(ctx, cluster, start, end, args.chunk_days),
                                   clusters, label="aks-autoscaler"):
            autoscaler.extend(result or [])
        output.csv("aks_autoscaler_monthly.csv", autoscaler)

    ooms, oom_clusters = [], 0
    if not args.skip_oom:
        client = ci.LogsClient(ctx, args.pause)
        with_ci = [row for row in clusters if row.get("container_insights", "").lower() == "true"]
        for row in clusters:
            if row not in with_ci:
                ctx.errors.add("ci-oomkilled", row["name"], "inventory", code="NotEnabled",
                               message="Container Insights add-on is not enabled on this cluster")
        started = time.monotonic()
        for number, cluster in enumerate(with_ci, 1):  # sequential: Log Analytics allows few concurrent queries
            azure.LOG.info("OOMKilled %d/%d %s (%.0f min)", number, len(with_ci), cluster["name"],
                           (time.monotonic() - started) / 60)
            rows = oom_rows(ctx, client, cluster, start, end)
            if rows is not None:
                oom_clusters += 1
                ooms.extend(rows)
        ctx.evidence.queries.append({"source": "Log Analytics (Container Insights)", "label": "oomkilled",
                                     "query": " ".join(OOM_QUERY.replace("{workload}", ci.WORKLOAD_NAME).split()),
                                     "rows": len(ooms), "subscriptions": "", "seconds": "",
                                     "timestamp": azure.iso(azure.utc_now())})
        output.csv("aks_oomkilled_monthly.csv", ooms)

    latency = []
    if not args.skip_apim:
        for service in services:
            latency.extend(latency_rows(ctx, service, start, end, args.top))
        output.csv("apim_latency_monthly.csv", latency)

    output.csv("signals_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    summary = {"autoscaler_rows": len(autoscaler), "oom_rows": len(ooms), "oom_clusters_queried": oom_clusters,
               "latency_rows": len(latency), "errors": len(ctx.errors.rows)}
    output.json("signals_summary.json", azure.run_metadata(ctx, {
        "window_utc": {"from": aa.iso_z(start), "to": aa.iso_z(end)},
        "office_hours": "Monday-Friday 08:00-20:00 America/Lima (UTC-5), holidays not excluded",
        "clusters": [row["name"] for row in clusters], "apim": [row["name"] for row in services], **summary}))
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
