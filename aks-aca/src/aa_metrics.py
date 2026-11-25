#!/usr/bin/env python3
"""Monthly AKS node usage and APIM requests per API, from Azure Monitor platform metrics.

Platform metrics are kept for 93 days. The default window is the last full calendar month (Lima), the
same period as aa_cost.py; --start and --end pick another one inside the retention. Office hours are
Monday to Friday, 8:00 to 20:00 Lima time (UTC-5); holidays count as working days.

AKS (per cluster and node pool, hourly points):
  node_cpu_usage_millicores, node_memory_working_set_bytes (Average and Maximum), split by node.
  kube_node_status_allocatable_cpu_cores and _memory_bytes (Average).
APIM (per API, 5-minute points):
  Requests (Total), split by ApiId. The 5-minute buckets give the off-hours wake-ups that Container
  Apps would bill: each run of buckets with traffic, plus the scale-to-zero cooldown.
  By default the APIM of the layers in aa_common.TRAFFIC_LAYERS (BS; UX, Open Banking, Yape and Core only stand in
  for applications without BS traffic), in --apim-environments (PROD; DESA requests are test traffic, for a one-off
  check); --apim-layers picks other layers and --apim-names picks services by name.
  Every APIM row carries apim_layer.

Outputs (in --output-dir):
  aks_usage_monthly.csv   cluster x pool x month: nodes, allocatable, used (office avg, off-hours avg,
                          p95, max) for CPU and memory
  apim_api_monthly.csv    APIM x API x month: requests (office / off-hours), peak 5 minutes and hour,
                          share of office buckets with traffic, off-hours wake-ups and active buckets
  apim_hourly_profile.csv APIM x API x weekday/weekend x Lima hour: average requests per hour
  entry_points_monthly.csv the same figures for entry points outside APIM: the TotalRequests of every Application
                          Gateway of an AKS application (AGIC sends to the pods, past APIM and the load balancer), and
                          the IncomingMessages of the Event Hubs / Service Bus namespaces of AKS applications with no
                          AKS-backed PROD API (event-driven: their load is messages); apim_layer says which
  entry_points_hourly_profile.csv the same profile for those entry points
  apim_api_hourly.csv     APIM x API x Lima hour with traffic: requests in the hour and its busiest 5-minute bucket
  entry_points_hourly.csv the same hourly series for the entry points
  metrics_errors.csv
  metrics_summary.json
"""

import json
import os
from collections import defaultdict
from datetime import timedelta

import aa_common as aa

azure = aa.azure

NODE_METRICS = ["node_cpu_usage_millicores", "node_memory_working_set_bytes"]
ALLOCATABLE_METRICS = ["kube_node_status_allocatable_cpu_cores", "kube_node_status_allocatable_memory_bytes"]
GIB = 1024 ** 3
BUCKET_SECONDS = 300


def percentile(values, share):
    if not values:
        return ""
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, int(round(share * (len(ordered) - 1)))))
    return ordered[index]


def mean(values):
    return sum(values) / len(values) if values else ""


# ---------------------------------------------------------------------------
# AKS
# ---------------------------------------------------------------------------


def cluster_usage(ctx, cluster, pools, start, end, args):
    """Hourly sums per pool, then monthly figures. Pool comes from the nodepool dimension."""
    # Per (pool, hour): CPU millicores and memory bytes summed over the pool's nodes, and the nodes seen.
    hourly = defaultdict(lambda: {"cpu_avg": 0.0, "cpu_max": 0.0, "mem_avg": 0.0, "mem_max": 0.0, "nodes": set()})
    allocatable = defaultdict(dict)
    for begin, finish in aa.chunks(start, end, args.chunk_days):
        pieces = ctx.try_arm("aks-node-metrics", cluster["name"], aa.get_metrics_split, ctx, cluster["id"],
                             NODE_METRICS, begin, finish, "PT1H", "Average,Maximum", "node eq '*' and nodepool eq '*'",
                             top=1000)
        for metric in (metric for values in pieces or [] for metric in values):
            name = metric["name"]["value"]
            kind = "cpu" if "cpu" in name else "mem"
            for series in metric.get("timeseries", []):
                dims = aa.series_dimensions(series)
                pool = dims.get("nodepool", "")
                for aggregation, suffix in (("average", "avg"), ("maximum", "max")):
                    for moment, value in aa.series_points(series, aggregation):
                        bucket = hourly[(pool, moment)]
                        bucket[f"{kind}_{suffix}"] += value
                        bucket["nodes"].add(dims.get("node", ""))
        values = ctx.try_arm("aks-allocatable", cluster["name"], aa.get_metrics, ctx, cluster["id"],
                             ALLOCATABLE_METRICS, begin, finish, "PT1H", "Average")
        for metric in values or []:
            kind = "cpu" if "cpu" in metric["name"]["value"] else "mem"
            for series in metric.get("timeseries", []):
                for moment, value in aa.series_points(series, "average"):
                    allocatable[moment][kind] = value

    pool_modes = {pool["pool"]: pool["mode"] for pool in pools}
    groups = defaultdict(list)
    for (pool, moment), bucket in hourly.items():
        groups[(pool, aa.lima_month(moment))].append((moment, bucket))
    rows = []
    for (pool, month), points in sorted(groups.items()):
        office = [bucket for moment, bucket in points if aa.is_office(moment)]
        off = [bucket for moment, bucket in points if not aa.is_office(moment)]
        cpu_all = [bucket["cpu_avg"] / 1000 for _, bucket in points]
        mem_all = [bucket["mem_avg"] / GIB for _, bucket in points]
        rows.append({
            "cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "month": month, "pool": pool,
            "mode": pool_modes.get(pool, ""), "hours": len(points),
            "nodes_avg": round(mean([len(bucket["nodes"]) for _, bucket in points]), 2),
            "nodes_max": max(len(bucket["nodes"]) for _, bucket in points),
            "cpu_used_office_avg": round(mean([b["cpu_avg"] / 1000 for b in office]), 3) if office else "",
            "cpu_used_offhours_avg": round(mean([b["cpu_avg"] / 1000 for b in off]), 3) if off else "",
            "cpu_used_min": round(min(cpu_all), 3),
            "cpu_used_p95": round(percentile(cpu_all, 0.95), 3),
            "cpu_used_max": round(max(bucket["cpu_max"] / 1000 for _, bucket in points), 3),
            "mem_used_office_avg_gib": round(mean([b["mem_avg"] / GIB for b in office]), 2) if office else "",
            "mem_used_offhours_avg_gib": round(mean([b["mem_avg"] / GIB for b in off]), 2) if off else "",
            "mem_used_min_gib": round(min(mem_all), 2),
            "mem_used_p95_gib": round(percentile(mem_all, 0.95), 2),
            "mem_used_max_gib": round(max(bucket["mem_max"] / GIB for _, bucket in points), 2),
        })
    # Cluster-wide allocatable, one row per month with pool = "*".
    by_month = defaultdict(list)
    for moment, values in allocatable.items():
        by_month[aa.lima_month(moment)].append(values)
    for month, values in sorted(by_month.items()):
        rows.append({
            "cluster": cluster["name"], "cluster_id": cluster["id"].lower(), "environment": cluster["environment"], "month": month, "pool": "*",
            "mode": "ALLOCATABLE", "hours": len(values),
            "cpu_allocatable_avg": round(mean([v["cpu"] for v in values if "cpu" in v] or [0]), 2),
            "mem_allocatable_avg_gib": round(mean([v["mem"] / GIB for v in values if "mem" in v] or [0]), 2),
        })
    return rows


# ---------------------------------------------------------------------------
# APIM
# ---------------------------------------------------------------------------


class ApiState:
    """Running monthly figures for one API, fed with 5-minute buckets in time order."""

    def __init__(self, cooldown):
        self.cooldown = cooldown
        self.months = defaultdict(lambda: defaultdict(float))
        self.hours = defaultdict(float)          # (month, local hour start) -> requests
        self.hour_peaks = defaultdict(float)     # (month, local hour start) -> busiest 5-minute bucket in that hour
        self.profile = defaultdict(float)        # (day type, hour) -> requests
        self.in_run = False
        self.last = None

    def add(self, moment, requests):
        month = aa.lima_month(moment)
        figures = self.months[month]
        office = aa.is_office(moment)
        local = moment.astimezone(aa.LIMA)
        if self.last is not None and moment - self.last > timedelta(seconds=BUCKET_SECONDS):
            self.in_run = False  # a gap with no points means no traffic
        self.last = moment
        figures["requests"] += requests
        figures["office_requests" if office else "offhours_requests"] += requests
        figures["office_buckets" if office else "offhours_buckets"] += 1
        figures["peak_5min"] = max(figures["peak_5min"], requests)
        hour_start = local.replace(minute=0, second=0, microsecond=0)
        self.hours[(month, hour_start)] += requests
        self.hour_peaks[(month, hour_start)] = max(self.hour_peaks[(month, hour_start)], requests)
        self.profile[("weekday" if local.weekday() < 5 else "weekend", local.hour)] += requests
        active = requests > 0
        if office:
            figures["office_active_buckets"] += 1 if active else 0
            self.in_run = False  # the minimum replica is already running
            return
        if active:
            figures["offhours_active_buckets"] += 1
            if not self.in_run:
                figures["offhours_wakeups"] += 1
            self.in_run = True
        else:
            self.in_run = False

    def rows(self, apim, api_id, calendar):
        for month, figures in sorted(self.months.items()):
            peak_hour = max((value for (m, _), value in self.hours.items() if m == month), default=0)
            # Buckets without traffic may come back with no point, so the denominator is the calendar.
            office_buckets = calendar.get((month, True), 0)
            yield {
                "apim": apim, "api_id": api_id, "month": month,
                "requests": int(figures["requests"]),
                "office_requests": int(figures["office_requests"]),
                "offhours_requests": int(figures["offhours_requests"]),
                "peak_5min_requests": int(figures["peak_5min"]),
                "peak_hour_requests": int(peak_hour),
                "office_hours_in_window": round(office_buckets * BUCKET_SECONDS / 3600, 1),
                "buckets_with_points": int(figures["office_buckets"] + figures["offhours_buckets"]),
                "office_active_share": round(figures["office_active_buckets"] / office_buckets, 4) if office_buckets else "",
                "offhours_wakeups": int(figures["offhours_wakeups"]),
                "offhours_active_buckets": int(figures["offhours_active_buckets"]),
                # Container Apps estimate: active while there is traffic, then idle for the cooldown.
                "offhours_active_seconds": int(figures["offhours_active_buckets"] * BUCKET_SECONDS),
                "offhours_idle_seconds": int(figures["offhours_wakeups"] * self.cooldown),
            }

    def hourly_rows(self, apim, api_id):
        """One row per hour with traffic: requests in the hour and its busiest 5-minute bucket (Lima time). The
        snapshot adds them up per cluster and application to draw min / average / peak per hour of the day."""
        for (month, hour_start), value in sorted(self.hours.items()):
            if value > 0:
                yield {"apim": apim, "api_id": api_id, "month": month,
                       "hour_start": hour_start.strftime("%Y-%m-%dT%H:00"), "requests": int(value),
                       "peak_5min_requests": int(self.hour_peaks[(month, hour_start)])}

    def profile_rows(self, apim, api_id, day_counts):
        for (day_type, hour), value in sorted(self.profile.items()):
            days = day_counts.get(day_type) or 1
            yield {"apim": apim, "api_id": api_id, "day_type": day_type, "hour": hour,
                   "avg_requests_per_hour": round(value / days, 1)}


def apim_requests(ctx, service, start, end, args):
    states = defaultdict(lambda: ApiState(args.cooldown))
    for begin, finish in aa.chunks(start, end, args.apim_chunk_days):
        pieces = ctx.try_arm("apim-requests", service["name"], aa.get_metrics_split, ctx, service["id"], ["Requests"],
                             begin, finish, "PT5M", "Total", "ApiId eq '*'", top=args.top)
        for values in pieces or []:
            series_list = (values or [{}])[0].get("timeseries", [])
            if len(series_list) >= args.top:
                ctx.errors.add("apim-requests", service["name"], "metrics", code="TopReached",
                               message=f"{len(series_list)} APIs returned; raise --top to get every API")
            for series in series_list:
                api_id = aa.series_dimensions(series).get("apiid", "")
                for moment, value in aa.series_points(series, "total"):
                    states[api_id].add(moment, value)
    return states


ENTRY_METRICS = {"APP_GATEWAY": "TotalRequests", "EVENT_HUB": "IncomingMessages", "SERVICE_BUS": "IncomingMessages"}


def entry_requests(ctx, resource, start, end, args):
    """Requests of an Application Gateway, or messages of an Event Hubs / Service Bus namespace, in 5-minute buckets."""
    state = ApiState(args.cooldown)
    for begin, finish in aa.chunks(start, end, args.apim_chunk_days):
        pieces = ctx.try_arm("entry-points", resource["name"], aa.get_metrics_split, ctx, resource["id"],
                             [ENTRY_METRICS[resource["kind"]]], begin, finish, "PT5M", "Total")
        for values in pieces or []:
            for series in (values or [{}])[0].get("timeseries", []):
                for moment, value in aa.series_points(series, "total"):
                    state.add(moment, value)
    return state


def entry_points(inventory, clusters, apis, environments):
    """Entry points outside APIM of the AKS applications in --environments: every Application Gateway, and the
    Event Hubs / Service Bus namespaces of applications without any AKS-backed PROD API (event-driven)."""
    with_api = {row.get("backend_cod_app") for row in apis
                if row.get("apim_environment", "PROD") == "PROD"
                and (row.get("backend_platform") or aa.backend_platform(row.get("backend_host"))) == "AKS"}
    aks_apps = {row["cod_app"] for row in clusters if row.get("cod_app") and row["cod_app"] not in aa.INFRASTRUCTURE_APPS
                and (environments is None or row["environment"] in environments)}

    def wanted(row):
        return (environments is None or row["environment"] in environments) and row.get("cod_app") in aks_apps

    gateways = [dict(row, kind="APP_GATEWAY") for row in aa.read_csv(os.path.join(inventory, "app_gateways.csv"))
                if wanted(row)]
    queues = [row for row in aa.read_csv(os.path.join(inventory, "messaging_namespaces.csv"))
              if wanted(row) and row["cod_app"] not in with_api]
    return gateways + queues


def bucket_calendar(start, end):
    """Number of 5-minute buckets per (Lima month, is office) in [start, end)."""
    counts = defaultdict(int)
    moment = start
    while moment < end:
        counts[(aa.lima_month(moment), aa.is_office(moment))] += 1
        moment += timedelta(seconds=BUCKET_SECONDS)
    return counts


def day_counts(start, end):
    counts = defaultdict(int)
    day = start.astimezone(aa.LIMA).date()
    last = (end - timedelta(seconds=1)).astimezone(aa.LIMA).date()
    while day <= last:
        counts["weekday" if day.weekday() < 5 else "weekend"] += 1
        day += timedelta(days=1)
    return counts


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-metrics"))
    aa.add_common_arguments(parser)
    parser.add_argument("--start", default=None, help="First day (YYYY-MM-DD, Lima). Default: first day of last month")
    parser.add_argument("--end", default=None, help="Last day (YYYY-MM-DD, Lima). Default: last day of last month")
    parser.add_argument("--clusters", default="", help="Comma-separated cluster names. Default: every cluster in --environments")
    parser.add_argument("--apim-names", default="",
                        help="Comma-separated APIM name fragments; overrides --apim-layers")
    parser.add_argument("--apim-environments", default="PROD",
                        help="Environments of the APIM services measured. DESA requests are test traffic: a one-off "
                             "check of activity, never added to PROD. Default: %(default)s")
    parser.add_argument("--apim-layers", default=",".join(aa.TRAFFIC_LAYERS),
                        help="Comma-separated APIM layers to measure (BS, UX, CORE, OPEN_BANKING, YAPE). Default: %(default)s")
    parser.add_argument("--skip-aks", action="store_true")
    parser.add_argument("--skip-apim", action="store_true")
    parser.add_argument("--skip-entry-points", action="store_true",
                        help="Do not measure Application Gateways nor the Event Hubs / Service Bus namespaces of "
                             "applications with no API on APIM")
    parser.add_argument("--chunk-days", type=int, default=31, help="Days per AKS metrics call (hourly points)")
    parser.add_argument("--apim-chunk-days", type=int, default=2, help="Days per APIM metrics call (5-minute points)")
    parser.add_argument("--top", type=int, default=2000, help="Max APIs per APIM metrics call")
    parser.add_argument("--cooldown", type=int, default=300,
                        help="Seconds a Container App replica stays up after the last request (scale rule cooldown)")
    aa.add_cache_argument(parser)
    parser.set_defaults(arm_rps=2.0, workers=2)
    args = parser.parse_args()
    azure.setup_logging(args.verbose)
    start, end = aa.metrics_window(args.start, args.end)
    environments = aa.wanted_environments(args)

    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    pools = aa.read_csv(os.path.join(args.inventory, "aks_node_pools.csv"))
    services = aa.read_csv(os.path.join(args.inventory, "apim_services.csv"))
    cluster_codes = aa.cluster_cod_app(aa.cod_app_by_cluster(clusters))
    api_inventory = aa.read_csv(os.path.join(args.inventory, "apim_apis.csv"))
    api_codes = aa.cod_app_by_api(api_inventory)
    entries = entry_points(args.inventory, clusters, api_inventory, environments)
    if not clusters and not services:
        raise SystemExit("Run aa_inventory.py first: no aks_clusters.csv or apim_services.csv in " + args.inventory)
    wanted_clusters = {name.strip().lower() for name in args.clusters.split(",") if name.strip()}
    clusters = [row for row in clusters
                if (wanted_clusters and row["name"].lower() in wanted_clusters)
                or (not wanted_clusters and (environments is None or row["environment"] in environments))]
    names = [name.strip().lower() for name in args.apim_names.split(",") if name.strip()]
    layers = {layer.strip().upper() for layer in args.apim_layers.split(",") if layer.strip()}
    apim_environments = {item.strip().upper() for item in args.apim_environments.split(",") if item.strip()} or None
    layer_of = {row["name"]: row.get("apim_layer") or aa.apim_layer(row["name"], row.get("custom_domains")) for row in services}
    services = [row for row in services
                if (names and any(name in row["name"].lower() for name in names))
                or (not names and (apim_environments is None or row["environment"] in apim_environments)
                    and layer_of[row["name"]] in layers)]

    ctx = azure.AzureContext(args)
    ctx.query_cache = aa.query_cache(args)  # every metrics GET of a finished window (aa.get_metrics)
    azure.LOG.info("Window %s to %s (UTC); %d clusters, %d APIM", aa.iso_z(start), aa.iso_z(end), len(clusters), len(services))
    output = azure.Output(args.output_dir, args.delimiter)

    usage_rows = []
    if not args.skip_aks:
        pools_by_cluster = defaultdict(list)
        for pool in pools:
            pools_by_cluster[pool["cluster_id"]].append(pool)
        for result in ctx.parallel(lambda cluster: cluster_usage(ctx, cluster, pools_by_cluster[cluster["id"].lower()],
                                                                 start, end, args), clusters, label="aks-metrics"):
            usage_rows.extend(result or [])
        output.csv("aks_usage_monthly.csv", aa.with_cod_app(usage_rows, "cluster", cluster_codes))

    api_rows, profile_rows, hourly_rows = [], [], []
    if not args.skip_apim:
        days = day_counts(start, end)
        calendar = bucket_calendar(start, end)
        for service in services:  # one APIM at a time, chunks in time order
            for api_id, state in sorted(apim_requests(ctx, service, start, end, args).items()):
                api_rows.extend(state.rows(service["name"], api_id, calendar))
                profile_rows.extend(state.profile_rows(service["name"], api_id, days))
                hourly_rows.extend(state.hourly_rows(service["name"], api_id))
        api_code = lambda row: api_codes.get((row["apim"], row["api_id"]), "")
        layer = lambda row: layer_of.get(row["apim"], "")
        output.csv("apim_api_monthly.csv", aa.with_cod_app(aa.insert_column(api_rows, "apim", "apim_layer", layer),
                                                           "api_id", api_code))
        output.csv("apim_hourly_profile.csv", aa.with_cod_app(aa.insert_column(profile_rows, "apim", "apim_layer", layer),
                                                              "api_id", api_code))
        output.csv("apim_api_hourly.csv", aa.insert_column(hourly_rows, "apim", "apim_layer", layer))

    entry_rows, entry_profile, entry_hourly = [], [], []
    if not args.skip_entry_points and entries:
        days = day_counts(start, end)
        calendar = bucket_calendar(start, end)
        for result, entry in zip(ctx.parallel(lambda item: entry_requests(ctx, item, start, end, args), entries,
                                              label="entry-points"), entries):
            if result is None:
                continue
            metric = ENTRY_METRICS[entry["kind"]]
            entry_rows.extend(dict(row, apim_layer=entry["kind"], cod_app=entry["cod_app"])
                              for row in result.rows(entry["name"], metric, calendar))
            entry_profile.extend(dict(row, apim_layer=entry["kind"], cod_app=entry["cod_app"])
                                 for row in result.profile_rows(entry["name"], metric, days))
            entry_hourly.extend(dict(row, apim_layer=entry["kind"], cod_app=entry["cod_app"])
                                for row in result.hourly_rows(entry["name"], metric))
        output.csv("entry_points_monthly.csv", entry_rows)
        output.csv("entry_points_hourly_profile.csv", entry_profile)
        output.csv("entry_points_hourly.csv", entry_hourly)

    output.csv("metrics_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    output.json("metrics_summary.json", azure.run_metadata(ctx, {
        "window_utc": {"from": aa.iso_z(start), "to": aa.iso_z(end)},
        "cache": ctx.query_cache.summary(),
        "office_hours": "Monday-Friday 08:00-20:00 America/Lima (UTC-5), holidays not excluded",
        "cooldown_seconds": args.cooldown,
        "clusters": [row["name"] for row in clusters],
        "apim": [row["name"] for row in services],
        "apim_layers": {row["name"]: layer_of[row["name"]] for row in services},
        "aks_rows": len(usage_rows), "apim_api_rows": len(api_rows),
        "entry_points": [f"{row['kind']}:{row['name']}" for row in entries] if not args.skip_entry_points else [],
    }))
    print(json.dumps({"aks_rows": len(usage_rows), "apim_api_rows": len(api_rows), "errors": len(ctx.errors.rows)}))


if __name__ == "__main__":
    main()
