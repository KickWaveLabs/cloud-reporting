#!/usr/bin/env python3
"""Workloads, replicas and per-replica CPU/memory per AKS cluster, from Container Insights.

Reads the Container Insights tables of each cluster through a resource-scoped Log Analytics query
(GET, read-only). Only inventory and performance tables are used: KubePodInventory, KubeServices and
Perf (K8SContainer counters). Container logs, pod labels and environment variables are never read.
Every query aggregates in Kusto, so only monthly figures leave the workspace.

Run aa_inventory.py first: it says which clusters have Container Insights (container_insights = True).
Results depend on the workspace retention; months older than that come back empty and are reported.

Outputs (in --output-dir):
  aks_workloads_monthly.csv  cluster x namespace x workload x month x office/off-hours: replicas (avg,
                             max) and per-counter usage and requests (total and per replica)
  aks_services.csv           cluster x namespace x service x type
  ci_coverage.csv            cluster x month: rows returned, so empty months are visible
  ci_errors.csv
  ci_summary.json
"""

import os
import time
from datetime import timedelta

import requests

import aa_common as aa

azure = aa.azure

LOGS_ENDPOINT = "https://api.loganalytics.io/v1"
LOGS_SCOPE = "https://api.loganalytics.io/.default"

# Workload name from the pod's controller: a ReplicaSet or CronJob Job loses its generated suffix. Shared with
# aa_signals.py; passed as a format() argument, so its braces stay single.
WORKLOAD_NAME = """case(
    ControllerKind == 'ReplicaSet', replace_regex(ControllerName, @'-[a-z0-9]{6,10}$', ''),
    ControllerKind == 'Job', replace_regex(ControllerName, @'-[0-9]{8,}$', ''),
    ControllerKind == 'EphemeralRunner', replace_regex(ControllerName, @'-[a-z0-9]{5}-runner-[a-z0-9]{5}$', ''),
    ControllerName)"""

WORKLOADS_QUERY = """
let t0 = datetime({start}); let t1 = datetime({end});
let pods = KubePodInventory
| where TimeGenerated between (t0 .. t1) and isnotempty(ControllerName)
| extend Workload = {workload}
| distinct PodUid, Namespace, Workload, ControllerKind;
Perf
| where TimeGenerated between (t0 .. t1)
| where ObjectName == 'K8SContainer'
| where CounterName in ('cpuUsageNanoCores', 'memoryWorkingSetBytes', 'cpuRequestNanoCores', 'memoryRequestBytes')
| extend PodUid = extract(@'/([^/]+)/[^/]+$', 1, InstanceName)
| summarize Value = avg(CounterValue) by Hour = bin(TimeGenerated, 1h), PodUid, InstanceName, CounterName
| summarize Value = sum(Value) by Hour, PodUid, CounterName
| join kind=inner pods on PodUid
| summarize Replicas = dcount(PodUid), Total = sum(Value) by Hour, Namespace, Workload, ControllerKind, CounterName
| extend Local = datetime_utc_to_local(Hour, 'America/Lima')
| extend Month = format_datetime(Local, 'yyyy-MM'),
    Office = dayofweek(Local) between (1d .. 5d) and hourofday(Local) between (8 .. 19)
| summarize Hours = count(), AvgReplicas = avg(Replicas), MaxReplicas = max(Replicas),
    AvgTotal = avg(Total), P95Total = percentile(Total, 95), MaxTotal = max(Total), AvgPerReplica = avg(Total / Replicas)
    by Month, Office, Namespace, Workload, ControllerKind, CounterName
"""

SERVICES_QUERY = """
KubeServices
| where TimeGenerated between (datetime({start}) .. datetime({end}))
| summarize LastSeen = max(TimeGenerated) by Namespace, ServiceName, ServiceType
"""

# Counters in Container Insights units -> (output name, divisor to vCPU or GiB)
COUNTERS = {
    "cpuUsageNanoCores": ("cpu_used", 1e9),
    "cpuRequestNanoCores": ("cpu_request", 1e9),
    "memoryWorkingSetBytes": ("mem_used_gib", 1024 ** 3),
    "memoryRequestBytes": ("mem_request_gib", 1024 ** 3),
}


class LogsClient:
    """Resource-scoped Log Analytics queries with GET, retries on 429/5xx and evidence of each call."""

    def __init__(self, ctx, pause):
        self.ctx = ctx
        self.pause = pause
        self.session = requests.Session()
        self._token = None

    def _bearer(self):
        if not self._token or self._token.expires_on - 300 < time.time():
            self._token = self.ctx.credential.get_token(LOGS_SCOPE)
        return self._token.token

    def query(self, resource_id, kql, timespan):
        url = f"{LOGS_ENDPOINT}{resource_id}/query"
        params = {"query": " ".join(kql.split()), "timespan": timespan}
        for attempt in range(1, 7):
            time.sleep(self.pause)
            response = self.session.get(url, params=params, timeout=300,
                                        headers={"Authorization": f"Bearer {self._bearer()}", "Prefer": "wait=300"})
            if response.status_code in (429, 500, 502, 503, 504) and attempt < 6:
                retry = response.headers.get("Retry-After")
                time.sleep(int(retry) if retry and retry.isdigit() else 10 * attempt)
                continue
            self.ctx.evidence.add_call("GET", url, "logs-v1", response.status_code)
            if response.status_code >= 400:
                error = response.json().get("error", {}) if response.content else {}
                raise azure.ArmError(response.status_code, error.get("code", ""), error.get("message", response.reason), url)
            tables = response.json().get("tables", [])
            if not tables:
                return []
            columns = [column["name"] for column in tables[0]["columns"]]
            return [dict(zip(columns, row)) for row in tables[0]["rows"]]
        raise azure.ArmError("", "RetryExhausted", "Retries exhausted", url)


def month_windows(start, end):
    """Calendar months (Lima) inside [start, end), as (label, start, end) in UTC."""
    current = start
    while current < end:
        local = current.astimezone(aa.LIMA)
        following_local = (local.replace(day=1) + timedelta(days=32)).replace(day=1, hour=0, minute=0, second=0)
        following = min(end, following_local.astimezone(start.tzinfo))
        yield local.strftime("%Y-%m"), current, following
        current = following


def workload_rows(cluster, rows):
    merged = {}
    for row in rows:
        key = (row["Month"], row["Office"], row["Namespace"], row["Workload"], row["ControllerKind"])
        target = merged.setdefault(key, {
            "cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "month": row["Month"],
            "period": "OFFICE" if row["Office"] else "OFFHOURS", "namespace": row["Namespace"],
            "workload": row["Workload"], "kind": row["ControllerKind"],
            "system_namespace": "YES" if row["Namespace"] in aa.SYSTEM_NAMESPACES
                                or str(row["Namespace"]).startswith("aks-istio") else "NO",
            "hours": row["Hours"], "replicas_avg": round(row["AvgReplicas"], 2), "replicas_max": row["MaxReplicas"],
        })
        name, divisor = COUNTERS[row["CounterName"]]
        target[f"{name}_avg"] = round(row["AvgTotal"] / divisor, 4)
        target[f"{name}_p95"] = round(row["P95Total"] / divisor, 4)
        target[f"{name}_max"] = round(row["MaxTotal"] / divisor, 4)
        target[f"{name}_per_replica"] = round(row["AvgPerReplica"] / divisor, 4)
    return list(merged.values())


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-container-insights"))
    aa.add_common_arguments(parser)
    parser.add_argument("--start", default=None, help="First day (YYYY-MM-DD, Lima). Default: first day of last month")
    parser.add_argument("--end", default=None, help="Last day (YYYY-MM-DD, Lima). Default: last day of last month")
    parser.add_argument("--clusters", default="", help="Comma-separated cluster names. Default: clusters in --environments")
    parser.add_argument("--pause", type=float, default=2.0, help="Seconds between Log Analytics queries")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)
    start, end = aa.metrics_window(args.start, args.end)
    environments = aa.wanted_environments(args)

    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    if not clusters:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv in " + args.inventory)
    cluster_codes = aa.cluster_cod_app(aa.cod_app_by_cluster(clusters))
    wanted = {name.strip().lower() for name in args.clusters.split(",") if name.strip()}
    clusters = [row for row in clusters
                if (wanted and row["name"].lower() in wanted)
                or (not wanted and (environments is None or row["environment"] in environments))]
    without = [row["name"] for row in clusters if row.get("container_insights", "").lower() != "true"]
    clusters = [row for row in clusters if row.get("container_insights", "").lower() == "true"]

    ctx = azure.AzureContext(args)
    for name in without:
        ctx.errors.add("container-insights", name, "inventory", code="NotEnabled",
                       message="Container Insights add-on is not enabled on this cluster")
    client = LogsClient(ctx, args.pause)
    workloads, services, coverage = [], [], []
    started = time.monotonic()
    for number, cluster in enumerate(clusters, 1):  # sequential: Log Analytics allows few concurrent queries per user
        azure.LOG.info("Container Insights %d/%d %s (%.0f min)", number, len(clusters), cluster["name"],
                       (time.monotonic() - started) / 60)
        for month, begin, finish in month_windows(start, end):
            timespan = f"{aa.iso_z(begin)}/{aa.iso_z(finish)}"
            kql = WORKLOADS_QUERY.format(start=aa.iso_z(begin), end=aa.iso_z(finish), workload=WORKLOAD_NAME)
            rows = ctx.try_arm("ci-workloads", cluster["name"], client.query, cluster["id"], kql, timespan)
            workloads.extend(workload_rows(cluster, rows or []))
            coverage.append({"cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "month": month, "rows": len(rows or []),
                             "status": "ERROR" if rows is None else ("EMPTY" if not rows else "OK")})
        kql = SERVICES_QUERY.format(start=aa.iso_z(start), end=aa.iso_z(end))
        rows = ctx.try_arm("ci-services", cluster["name"], client.query, cluster["id"], kql,
                           f"{aa.iso_z(start)}/{aa.iso_z(end)}")
        for row in rows or []:
            services.append({"cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"],
                             "namespace": row["Namespace"], "service": row["ServiceName"],
                             "service_type": row["ServiceType"], "last_seen": row["LastSeen"]})

    ctx.evidence.queries.append({"source": "Log Analytics (Container Insights)", "label": "workloads",
                                 "query": " ".join(WORKLOADS_QUERY.replace("{workload}", WORKLOAD_NAME).split()), "rows": len(workloads),
                                 "subscriptions": "", "seconds": "", "timestamp": azure.iso(azure.utc_now())})
    ctx.evidence.queries.append({"source": "Log Analytics (Container Insights)", "label": "services",
                                 "query": " ".join(SERVICES_QUERY.split()), "rows": len(services),
                                 "subscriptions": "", "seconds": "", "timestamp": azure.iso(azure.utc_now())})
    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("aks_workloads_monthly.csv", aa.with_cod_app(workloads, "cluster", cluster_codes))
    output.csv("aks_services.csv", aa.with_cod_app(services, "cluster", cluster_codes))
    output.csv("ci_coverage.csv", aa.with_cod_app(coverage, "cluster", cluster_codes))
    output.csv("ci_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    output.json("ci_summary.json", azure.run_metadata(ctx, {
        "window_utc": {"from": aa.iso_z(start), "to": aa.iso_z(end)},
        "clusters_queried": [row["name"] for row in clusters],
        "clusters_without_container_insights": without,
        "coverage": azure.count_by(coverage, "status"),
    }))


if __name__ == "__main__":
    main()
