#!/usr/bin/env python3
"""Capacity, memory and latency signals: unschedulable pods, OOM kills, stack and memory per container, APIM latency.

Three read-only checks that test the usage figures from the other side: low CPU use is only waste if the
clusters are not short of capacity (unschedulable pods) and the containers are not short of memory
(OOMKilled). Backend latency per API tells how much a Container Apps cold start would add.

  AKS platform metric   cluster_autoscaler_unschedulable_pods_count (Average, hourly), per cluster.
                        Only emitted when the cluster autoscaler is on; other clusters return no points.
  Container Insights    KubePodInventory rows whose ContainerLastStatus is OOMKilled. One OOM kill = one
                        (pod, container, restart count) with that last status, so a status repeated in
                        every inventory snapshot counts once. A kill from before the window counts in its
                        first month. Only counts leave the workspace: no logs, labels or messages.
  Container Insights    Stack per workload and memory per replica split into app and Istio sidecar, office hours
                        of the last --stack-days days (memory per container is stable, so a week is enough). The
                        stack is classified inside the workspace from ContainerInventory (command and image of
                        every container of the pod, including the init container that injects a Java agent);
                        only the label leaves it, never the command, the image or the environment variables.
  APIM platform metrics BackendDuration and Duration (milliseconds, Average and Maximum) per API and month,
                        one data point per month (interval FULL), so the average is per request.

Run aa_inventory.py first. The window matches aa_metrics.py: the last full month by default (Lima).

Outputs (in --output-dir):
  aks_autoscaler_monthly.csv  cluster x month: hours with data, hours with unschedulable pods (all, office),
                              highest hourly average of unschedulable pods
  aks_oomkilled_monthly.csv   cluster x namespace x workload x month: OOM kills and pods affected
  aks_stack_memory.csv        cluster x namespace x workload: stack, replicas, app and sidecar memory per replica
  apim_latency_monthly.csv    APIM x API x month: backend and gateway duration, average and maximum (ms)
  signals_errors.csv
  signals_summary.json
"""

import json
import os
import time
from datetime import timedelta

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


STACK_QUERY = """
let t0 = datetime({start}); let t1 = datetime({end});
let clues = ContainerInventory
| where TimeGenerated between (t0 .. t1)
| summarize arg_max(TimeGenerated, Command, Image) by ContainerID
| extend c = tolower(Command), i = tolower(Image)
| project ContainerID,
    Quarkus = toint(c has 'quarkus' or i has 'quarkus'),
    Spring = toint(c has 'springframework' or c has 'jarlauncher' or i has 'spring'),
    Java = toint(c has 'java' or c has 'javaagent' or i has 'java' or i has 'openjdk' or i has 'temurin' or i has 'jre'),
    Dotnet = toint(c has 'dotnet' or i has 'dotnet' or i has 'aspnet'),
    Node = toint(c has 'node' or i has 'node'),
    Python = toint(c has 'python' or c has 'gunicorn' or c has 'uvicorn' or i has 'python');
KubePodInventory
| where TimeGenerated between (t0 .. t1) and isnotempty(ControllerName) and isnotempty(ContainerID)
| extend Workload = {workload}
| distinct ContainerID, Namespace, Workload, ControllerKind
| join kind=inner clues on ContainerID
| summarize Quarkus = max(Quarkus), Spring = max(Spring), Java = max(Java), Dotnet = max(Dotnet), Node = max(Node),
    Python = max(Python) by Namespace, Workload, ControllerKind
| extend Stack = case(Quarkus == 1, 'Quarkus', Spring == 1, 'Spring Boot', Java == 1, 'Java', Dotnet == 1, '.NET',
    Node == 1, 'Node.js', Python == 1, 'Python', 'Unidentified')
| project Namespace, Workload, ControllerKind, Stack
"""

MEMORY_QUERY = """
let t0 = datetime({start}); let t1 = datetime({end});
let pods = KubePodInventory
| where TimeGenerated between (t0 .. t1) and isnotempty(ControllerName)
| extend Workload = {workload}
| distinct PodUid, Namespace, Workload, ControllerKind;
Perf
| where TimeGenerated between (t0 .. t1)
| where ObjectName == 'K8SContainer' and CounterName == 'memoryWorkingSetBytes'
| extend Local = datetime_utc_to_local(TimeGenerated, 'America/Lima')
| where dayofweek(Local) between (1d .. 5d) and hourofday(Local) between (8 .. 19)
| extend PodUid = extract(@'/([^/]+)/[^/]+$', 1, InstanceName), Container = extract(@'/([^/]+)$', 1, InstanceName)
| summarize Value = avg(CounterValue) by Hour = bin(TimeGenerated, 1h), PodUid,
    Role = iff(Container == 'istio-proxy', 'sidecar', 'app')
| join kind=inner pods on PodUid
| summarize Total = sum(Value), Pods = dcount(PodUid) by Hour, Namespace, Workload, ControllerKind, Role
| summarize PerReplica = avg(Total / Pods), Replicas = avg(Pods) by Namespace, Workload, ControllerKind, Role
"""


def stack_rows(ctx, client, cluster, start, end):
    """Stack and memory per replica (app and sidecar) per workload; None when a query fails."""
    timespan = f"{aa.iso_z(start)}/{aa.iso_z(end)}"
    arguments = {"start": aa.iso_z(start), "end": aa.iso_z(end), "workload": ci.WORKLOAD_NAME}
    stacks = ctx.try_arm("ci-stack", cluster["name"], client.query, cluster["id"], STACK_QUERY.format(**arguments), timespan)
    memory = ctx.try_arm("ci-memory", cluster["name"], client.query, cluster["id"], MEMORY_QUERY.format(**arguments),
                         timespan)
    if stacks is None or memory is None:
        return None
    by_workload = {}
    for row in memory:
        key = (row["Namespace"], row["Workload"], row["ControllerKind"])
        entry = by_workload.setdefault(key, {"replicas_avg": 0.0, "app_mem_gib_per_replica": "",
                                             "sidecar_mem_gib_per_replica": ""})
        entry[f"{row['Role']}_mem_gib_per_replica"] = round(row["PerReplica"] / 1024 ** 3, 3)
        if row["Role"] == "app":
            entry["replicas_avg"] = round(row["Replicas"], 2)
    stack_of = {(row["Namespace"], row["Workload"], row["ControllerKind"]): row["Stack"] for row in stacks}
    return [{"cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "namespace": namespace,
             "workload": workload, "kind": kind,
             "system_namespace": "YES" if namespace in aa.SYSTEM_NAMESPACES or str(namespace).startswith("aks-istio") else "NO",
             "stack": stack_of.get((namespace, workload, kind), "Unidentified"), **figures}
            for (namespace, workload, kind), figures in sorted(by_workload.items())]


def autoscaler_rows(ctx, cluster, start, end, chunk_days):
    azure.LOG.info("Unschedulable pods %s", cluster["name"])
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
    return [{"cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "month": month,
             "hours_with_data": entry["hours"], "hours_with_unschedulable": entry["unschedulable"],
             "office_hours_with_unschedulable": entry["office"], "unschedulable_pods_max": round(entry["max"], 2)}
            for month, entry in sorted(months.items())]


def oom_rows(ctx, client, cluster, start, end):
    kql = OOM_QUERY.format(start=aa.iso_z(start), end=aa.iso_z(end), workload=ci.WORKLOAD_NAME)
    rows = ctx.try_arm("ci-oomkilled", cluster["name"], client.query, cluster["id"], kql,
                       f"{aa.iso_z(start)}/{aa.iso_z(end)}")
    if rows is None:
        return None
    return [{"cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "month": row["Month"],
             "namespace": row["Namespace"], "workload": row["Workload"], "kind": row["ControllerKind"],
             "system_namespace": "YES" if row["Namespace"] in aa.SYSTEM_NAMESPACES
                                 or str(row["Namespace"]).startswith("aks-istio") else "NO",
             "oom_kills": row["OomKills"], "pods": row["Pods"]} for row in rows]


def latency_rows(ctx, service, start, end, top):
    out = {}
    for month, begin, finish in ci.month_windows(start, end):
        azure.LOG.info("APIM latency %s %s", service["name"], month)
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
    parser.add_argument("--start", default=None, help="First day (YYYY-MM-DD, Lima). Default: first day of last month")
    parser.add_argument("--end", default=None, help="Last day (YYYY-MM-DD, Lima). Default: last day of last month")
    parser.add_argument("--clusters", default="", help="Comma-separated cluster names. Default: clusters in --environments")
    parser.add_argument("--apim-names", default="", help="Comma-separated APIM name fragments. Default: APIM in --environments")
    parser.add_argument("--skip-autoscaler", action="store_true")
    parser.add_argument("--skip-oom", action="store_true")
    parser.add_argument("--skip-apim", action="store_true")
    parser.add_argument("--skip-stack", action="store_true")
    parser.add_argument("--stack-days", type=int, default=7, help="Days before --end for the stack and memory split")
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
    cluster_codes = aa.cluster_cod_app(aa.cod_app_by_cluster(clusters))
    api_codes = aa.cod_app_by_api(aa.read_csv(os.path.join(args.inventory, "apim_apis.csv")))
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
        output.csv("aks_autoscaler_monthly.csv", aa.with_cod_app(autoscaler, "cluster", cluster_codes))

    client = ci.LogsClient(ctx, args.pause)
    with_ci = [row for row in clusters if row.get("container_insights", "").lower() == "true"]
    if not (args.skip_oom and args.skip_stack):
        for row in clusters:
            if row not in with_ci:
                ctx.errors.add("container-insights", row["name"], "inventory", code="NotEnabled",
                               message="Container Insights add-on is not enabled on this cluster")
    ooms, oom_clusters = [], 0
    if not args.skip_oom:
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
        output.csv("aks_oomkilled_monthly.csv", aa.with_cod_app(ooms, "cluster", cluster_codes))

    stack, stack_clusters = [], 0
    if not args.skip_stack:
        stack_start = max(start, end - timedelta(days=args.stack_days))
        started = time.monotonic()
        for number, cluster in enumerate(with_ci, 1):
            azure.LOG.info("Stack and memory %d/%d %s (%.0f min)", number, len(with_ci), cluster["name"],
                           (time.monotonic() - started) / 60)
            rows = stack_rows(ctx, client, cluster, stack_start, end)
            if rows is not None:
                stack_clusters += 1
                stack.extend(rows)
        for label, query in (("stack", STACK_QUERY), ("memory-by-container", MEMORY_QUERY)):
            ctx.evidence.queries.append({"source": "Log Analytics (Container Insights)", "label": label,
                                         "query": " ".join(query.replace("{workload}", ci.WORKLOAD_NAME).split()),
                                         "rows": len(stack), "subscriptions": "", "seconds": "",
                                         "timestamp": azure.iso(azure.utc_now())})
        output.csv("aks_stack_memory.csv", aa.with_cod_app(stack, "cluster", cluster_codes))

    latency = []
    if not args.skip_apim:
        for service in services:
            latency.extend(latency_rows(ctx, service, start, end, args.top))
        output.csv("apim_latency_monthly.csv", aa.with_cod_app(
            latency, "api_id", lambda row: api_codes.get((row["apim"], row["api_id"]), "")))

    output.csv("signals_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    summary = {"autoscaler_rows": len(autoscaler), "oom_rows": len(ooms), "oom_clusters_queried": oom_clusters,
               "stack_rows": len(stack), "stack_clusters_queried": stack_clusters, "latency_rows": len(latency),
               "errors": len(ctx.errors.rows)}
    output.json("signals_summary.json", azure.run_metadata(ctx, {
        "window_utc": {"from": aa.iso_z(start), "to": aa.iso_z(end)},
        "office_hours": "Monday-Friday 08:00-20:00 America/Lima (UTC-5), holidays not excluded",
        "clusters": [row["name"] for row in clusters], "apim": [row["name"] for row in services], **summary}))
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
