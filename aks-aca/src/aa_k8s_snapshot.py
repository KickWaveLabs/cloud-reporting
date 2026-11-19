#!/usr/bin/env python3
"""HPA and deployment configuration of one day, joined with the replica and usage history of Container Insights.

Reads two exports from the platform team, kept outside git in --inputs (.local/aks-aca-inputs by default):
  reportHpa*.csv|xlsx         one row per HPA: min, max, current replicas, last scale time, replica trend
  reportDeployment*.csv|xlsx  one row per deployment: CPU and memory requests and limits
and, optionally, cluster_overrides.csv in the same folder (cluster_id or cluster, category, reason) to correct
the category of a cluster by hand.

The exports are a snapshot of the export day (configuration, no history). Behaviour comes from
aa_container_insights.py (replicas and usage per workload and month) and node size from aa_metrics.py, so run
them first. No Azure call is made here. Clusters are joined by resource id; a name is only used when a
Container Insights or metrics row has no cluster_id and the name is unique in the inventory.

Only the columns listed in HPA_COLUMNS and DEPLOYMENT_COLUMNS are read; the rest of each export (subscription,
application, owner and provisioning columns) is never copied to the results. `bia` is the business impact
level of the application (MUY ALTA to BAJA, kept as the bank writes it); a cluster takes the highest level
among its microservices.

A microservice is a deployment outside platform namespaces and not a known tool (aa_common.is_microservice).
A domain is a namespace with microservices. The layer is the first word of the deployment name.

Cluster category (only APPLICATION clusters enter the analysis):
  INFRASTRUCTURE   its app code is in aa_common.INFRASTRUCTURE_APPS, or it only runs tools
  APPLICATION      at least one microservice in the export or in Container Insights, or only the application's
                   own gateway (aa_common.APP_GATEWAY_PREFIXES)
  NO_APPLICATIONS  in the export or with Container Insights, but without a single microservice
  NO_DATA          neither in the export nor with Container Insights data

A cluster is a technical fit for Container Apps when it has 1 to 15 microservices, every one fits a Container
Apps replica by its request (4 vCPU / 8 GiB) and Container Insights shows no StatefulSet among its microservices.

Replicas needed by CPU = office CPU used by all replicas / (request per replica x 70%, a typical HPA target).
A minimum above it is more than the load asks for; above max(2, needed) it is more even keeping two replicas
for availability.

HPA behaviour uses the office-hours average of replicas in the period, not the hourly maximum: Container
Insights counts the pods seen in each hour, so every rolling deployment briefly adds the new pods to the old
ones and the hourly maximum overstates scaling. At the minimum = average <= 110% of minReplicas; near the
maximum = average >= 90% of maxReplicas. ci_replicas_max_office (hourly maximum in office hours) is shown next
to hpa_max with that caveat.

Time without scaling: days from lastScaleTime to the export day. With min = max the date is when the HPA was
set, not a scaling event, so no_scale_cause separates FIXED_MIN_MAX from MIN_ABOVE_NEED (the minimum exceeds
what the CPU needs) and LOAD_BELOW_TARGET (the load never reaches the HPA target).

Traffic per application: each APIM API maps to an app code through its backend (backend_cod_app in apim_apis.csv,
see aa_inventory.py); a <service>.<namespace>.svc.cluster.local backend (self-hosted gateway) takes the code of the
PROD cluster that runs that namespace in Container Insights. Only APIs whose backend runs on AKS count.
One APIM layer per application (traffic_layer): the first of BS, UX, Open Banking, Yape and Core with traffic in the
month; layers are never added, since a call goes on from one gateway to the next. traffic_status says why an
application has no traffic: ZERO_IN_MONTH (APIs on a measured APIM, no requests), LAYER_NOT_MEASURED (APIs only on
layers the metrics did not cover) or NO_APIM_API; traffic_note adds what the data shows (new cluster, only event
consumers, own ingress load balancer, no microservice running).
peak_tps_upper_bound adds each API's 5-minute peak, so it is an upper bound;
busiest_hour_tps adds the APIs hour by hour (average weekday or weekend profile of the month) and is what the
100 TPS rule uses. busiest_hour says which hour it is (day type and Lima hour); busiest_office_hour_tps and
busiest_offhours_hour_tps split it between the ADT_019 office window (weekdays 8:00-19:59) and the rest.
offhours_active_share is the share of off-hours 5-minute intervals with traffic (busiest API of the app); above
50% the app cannot scale to zero, so it goes to AKS (ADT_019: never warm 24/7). An hour has traffic when it carries at least 10% of the busiest hour.

Units as exported: CPU in millicores; memory in MB (10^6 bytes, so 2Gi shows as 2147.483648).

Outputs (in --output-dir):
  hpa_snapshot.csv         one row per HPA with its Container Insights history and flags
  deployment_snapshot.csv  one row per deployment with requests vs measured usage and flags
  cluster_snapshot.csv     one row per cluster (inventory and export): category, microservices, namespaces, HPAs
  capability_domains.csv   per capability keyword (notification, token, login...): PROD namespaces and
                           deployments of APPLICATION clusters whose name carries it
  application_snapshot.csv one row per PROD application (app code of APPLICATION clusters): microservices (and the
                           projection with DESA: PROD plus the DESA microservices not in PROD yet, by name without
                           version; it feeds the more-than-15 rules), cost of the last month, APIM traffic (layer, TPS, hours with traffic), HPAs
                           that do not scale, and the runtime the ADT_019 tree gives it (AKS, ACA, ACA_WITH_REVIEW)
                           with the reason
  snapshot_summary.json    input files, snapshot date, categories, rows without a match and counts per flag (PROD)
and, in --history-dir (results/aa-history, never cleared), export_history.csv: one row per export day with the
PROD clusters by category and the PROD microservices of APPLICATION clusters.
"""

import argparse
import csv
import glob
import io
import math
import os
import re
import statistics
from calendar import monthrange
from collections import Counter, defaultdict
from datetime import date, datetime

import aa_common as aa

azure = aa.azure

DEFAULT_INPUTS = os.path.join(os.path.dirname(aa.TOPIC_DIR), ".local", "aks-aca-inputs")
OVERRIDES_FILE = "cluster_overrides.csv"
EXPORT_HISTORY_FILE = "export_history.csv"
MB_TO_GIB = 1e6 / 1024 ** 3
ACA_MAX_VCPU, ACA_MAX_GIB = 4, 8
OVER_REQUEST_FACTOR = 3
HIGH_LIMIT_FACTOR = 10
NODE_SHARE_TOO_BIG = 0.5

# Output name -> column in the export (matched without case).
HPA_COLUMNS = {"hpa": "HPAs", "deployment": "Deployment", "namespace": "Namespace", "last_scale_time": "lastScaleTime",
               "replicas_now": "Val", "trend_max_export": "maxHPA", "trend": "ValTrend", "hpa_min": "currentMin",
               "hpa_max": "currentMax", "cluster": "resourceName", "cluster_id": "resourceId",
               "environment_export": "environment"}
DEPLOYMENT_COLUMNS = {"namespace": "Namespace", "service": "ServiceName", "deployment": "Deployment",
                      "cpu_request_m": "CpuRequest", "cpu_limit_m": "CpuLimit", "mem_request_mb": "MemoryRequest",
                      "mem_limit_mb": "MemoryLimit", "cluster": "resourceName", "cluster_id": "resourceId",
                      "environment_export": "environment", "bia": "bia"}
ENVIRONMENTS = {"prod": "PROD", "cert": "CERT", "desa": "DESA", "dev": "DESA"}
BIA_ORDER = ("MUY ALTA", "ALTA", "MEDIA", "BAJA", "MUY BAJA")
NO_BIA = "SIN BIA"
ADT_MAX_MICROSERVICES = 15
LAYERS = ("business", "channel", "app", "async", "support", "core")
APP_KINDS = ("ReplicaSet", "StatefulSet")
# Capability -> words in the deployment name. Words with a hyphen match as text, single words as a whole word.
CAPABILITIES = {
    "notification": ("notification", "notifications"),
    "operation-number": ("operation-number",),
    "authorization": ("authorization", "authorize", "authoriza"),
    "one-time-password": ("one-time-password", "otp"),
    "token": ("token", "tokens"),
    "session": ("session", "sessions"),
    "login": ("login",),
}
# Days without scaling -> bucket, checked in order.
SCALE_BUCKETS = ((30, "UNDER_1_MONTH"), (90, "1_TO_3_MONTHS"), (365, "3_TO_12_MONTHS"))
OVER_1_YEAR = "OVER_1_YEAR"
# ADT_019 thresholds for AKS, and the share of the busiest hour above which an hour counts as "with traffic".
ADT_MAX_TPS, ADT_MAX_ACTIVE_HOURS = 100, 10
# Share of off-hours 5-minute intervals with traffic above which an application cannot scale to zero: on
# Container Apps it would stay warm around the clock, which ADT_019 rules out (never warm 24/7).
OFFHOURS_ACTIVE_MAX = 0.5
ACTIVE_HOUR_SHARE = 0.1
CRITICAL_BIA = ("MUY ALTA", "ALTA")


def normalize_bia(value):
    text = " ".join(str(value or "").upper().split())
    return text if text and text != "-" else NO_BIA


def most_critical(levels):
    """Highest BIA among the given levels; unknown labels rank below the known ones, SIN BIA last."""
    known = [level for level in levels if level != NO_BIA]
    if not known:
        return NO_BIA
    return min(known, key=lambda level: BIA_ORDER.index(level) if level in BIA_ORDER else len(BIA_ORDER))


def layer_of(deployment):
    first = str(deployment or "").lower().split("-")[0]
    return first if first in LAYERS else "other"


def capabilities_of(deployment):
    name = str(deployment or "").lower()
    words = set(name.split("-"))
    return [capability for capability, keys in CAPABILITIES.items()
            if any((key in name) if "-" in key else (key in words) for key in keys)]


AT_MIN_FACTOR, NEAR_MAX_FACTOR = 1.1, 0.9
# Replicas the load needs: CPU used in office hours over the request at a typical HPA target utilization.
CPU_TARGET = 0.7
HA_REPLICAS = 2


def behaviour_from_average(average, low, high):
    """Same codes as behaviour(), from the office-hours average of replicas."""
    if None in (low, high):
        return "NO_DATA"
    if low >= high:
        return "FIXED"
    if average in (None, ""):
        return "NO_DATA"
    if average <= low * AT_MIN_FACTOR:
        return "NEVER_SCALES"
    return "AT_MAX" if average >= high * NEAR_MAX_FACTOR else "SCALES"


def replicas_needed(ci):
    """Replicas the measured office CPU needs at CPU_TARGET of the request, or None without data."""
    used, request, replicas = (ci.get(key) for key in ("ci_cpu_used_per_replica", "ci_cpu_request_per_replica",
                                                       "ci_replicas_avg_office"))
    if not all(isinstance(value, (int, float)) for value in (used, request, replicas)) or request <= 0:
        return None
    return max(1, math.ceil(used * replicas / (request * CPU_TARGET) - 1e-9))


def behaviour(peak, low, high):
    """FIXED (min = max), NEVER_SCALES (peak <= min), AT_MAX (peak >= max), SCALES, or NO_DATA."""
    if None in (low, high):
        return "NO_DATA"
    if low >= high:
        return "FIXED"
    if peak in (None, ""):
        return "NO_DATA"
    if peak <= low:
        return "NEVER_SCALES"
    return "AT_MAX" if peak >= high else "SCALES"


def scale_bucket(days):
    if days in (None, ""):
        return "NO_DATA"
    return next((label for limit, label in SCALE_BUCKETS if days < limit), OVER_1_YEAR)


def no_scale_cause(fixed, behaviour_ci, min_above_need):
    if fixed:
        return "FIXED_MIN_MAX"
    if behaviour_ci == "NO_DATA":
        return "NO_DATA"
    if behaviour_ci in ("SCALES", "AT_MAX"):
        return "SCALES"
    return "MIN_ABOVE_NEED" if min_above_need else "LOAD_BELOW_TARGET"


def yes(value):
    return "YES" if value else "NO"


def to_float(value):
    if value in (None, ""):
        return None
    try:
        number = float(str(value).strip())
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def parse_trend(value):
    text = str(value or "").strip().strip("[]")
    return [number for number in (to_float(part) for part in text.split(",")) if number is not None]


def find_input(folder, prefix):
    """Newest reportX*.csv or reportX*.xlsx in the folder."""
    matches = [path for pattern in (f"{prefix}*.csv", f"{prefix}*.xlsx") for path in glob.glob(os.path.join(folder, pattern))]
    if not matches:
        raise SystemExit(f"No {prefix}*.csv or {prefix}*.xlsx in {folder}. Copy the export there or pass --inputs.")
    return max(matches, key=os.path.getmtime)


def read_rows(path):
    if path.lower().endswith(".xlsx"):
        from openpyxl import load_workbook
        sheet = load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = sheet.iter_rows(values_only=True)
        header = [str(cell or "").strip() for cell in next(rows)]
        return [dict(zip(header, ["" if cell is None else cell for cell in row])) for row in rows if any(row)]
    with open(path, "rb") as handle:
        raw = handle.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:  # exports saved by Excel on Windows use cp1252
        text = raw.decode("cp1252")
    delimiter = max(";,\t", key=text.split("\n")[0].count)
    return list(csv.DictReader(io.StringIO(text, newline=""), delimiter=delimiter))


def select(rows, columns, path, ids_by_name):
    """Keep only the wanted columns, renamed; stop with the list of missing ones. Adds the cluster's resource id."""
    present = {key.strip().lower(): key for key in (rows[0] if rows else {})}
    missing = [source for source in columns.values() if source.lower() not in present]
    if missing:
        raise SystemExit(f"{os.path.basename(path)} lacks the columns {', '.join(missing)}. "
                         f"Columns found: {', '.join(present.values())}")
    known = {cluster_id for ids in ids_by_name.values() for cluster_id in ids}
    out = []
    for row in rows:
        record = {name: row[present[source.lower()]] for name, source in columns.items()}
        own = "/" + str(record.get("cluster_id") or "").strip().strip("/").lower()
        if not record.get("cluster"):
            record["cluster"] = own.split("/")[-1]
        # The export's id is used when it matches the inventory; otherwise the name decides.
        record["cluster_id"] = own if own in known else aa.cluster_id_of({"cluster": record["cluster"]}, ids_by_name)
        out.append(record)
    return out


def snapshot_date(path, override):
    if override:
        return date.fromisoformat(override)
    return datetime.fromtimestamp(os.path.getmtime(path)).date()


def cluster_key(row):
    """Join key of a cluster: its resource id, else its name (clusters outside the inventory)."""
    return row["cluster_id"] or "name:" + str(row["cluster"]).lower()


def ci_history(rows, ids_by_name):
    """Container Insights rows grouped by (cluster key, namespace, workload): replicas and per-replica usage."""
    history = defaultdict(lambda: {"months": set(), "office_hours": 0.0, "offhours_hours": 0.0, "replicas_max": 0.0,
                                   "replicas_max_office": 0.0, "sums": defaultdict(float)})
    weighted = ("replicas_avg", "cpu_used_per_replica", "cpu_request_per_replica", "mem_used_gib_per_replica",
                "mem_request_gib_per_replica")
    for row in rows:
        key = (cluster_key(dict(row, cluster_id=aa.cluster_id_of(row, ids_by_name))), row["namespace"], row["workload"])
        entry = history[key]
        hours = to_float(row.get("hours")) or 0
        entry["months"].add(row["month"])
        replicas_max = to_float(row.get("replicas_max")) or 0
        entry["replicas_max"] = max(entry["replicas_max"], replicas_max)
        period = "office" if row["period"] == "OFFICE" else "offhours"
        if period == "office":
            entry["replicas_max_office"] = max(entry["replicas_max_office"], replicas_max)
        entry[f"{period}_hours"] += hours
        for field in weighted:
            value = to_float(row.get(field))
            if value is not None:
                entry["sums"][f"{period}_{field}"] += value * hours
    result = {}
    for key, entry in history.items():
        def average(period, field):
            hours = entry[f"{period}_hours"]
            return round(entry["sums"][f"{period}_{field}"] / hours, 4) if hours else ""
        result[key] = {
            "ci_months": len(entry["months"]),
            "ci_replicas_avg_office": average("office", "replicas_avg"),
            "ci_replicas_avg_offhours": average("offhours", "replicas_avg"),
            "ci_replicas_max": entry["replicas_max"],
            "ci_replicas_max_office": entry["replicas_max_office"] if entry["office_hours"] else "",
            "ci_cpu_used_per_replica": average("office", "cpu_used_per_replica"),
            "ci_cpu_request_per_replica": average("office", "cpu_request_per_replica"),
            "ci_mem_used_gib_per_replica": average("office", "mem_used_gib_per_replica"),
            "ci_mem_request_gib_per_replica": average("office", "mem_request_gib_per_replica"),
        }
    return result


def node_sizes(usage, ids_by_name):
    """Approximate allocatable vCPU and GiB per node of each cluster, from the last month of aks_usage_monthly."""
    last = max((row["month"] for row in usage), default="")
    nodes, capacity = defaultdict(float), {}
    for row in usage:
        if row["month"] != last:
            continue
        key = cluster_key(dict(row, cluster_id=aa.cluster_id_of(row, ids_by_name)))
        if row.get("mode") == "ALLOCATABLE":
            capacity[key] = (to_float(row.get("cpu_allocatable_avg")), to_float(row.get("mem_allocatable_avg_gib")))
        else:
            nodes[key] += to_float(row.get("nodes_avg")) or 0
    sizes = {}
    for key, (cpu, mem) in capacity.items():
        if nodes.get(key) and cpu and mem:
            sizes[key] = (round(cpu / nodes[key], 2), round(mem / nodes[key], 2))
    return sizes


def environment_of(row, inventory):
    cluster = inventory.get(row["cluster_id"])
    if cluster:
        return cluster["environment"], "YES"
    return ENVIRONMENTS.get(str(row.get("environment_export") or "").strip().lower(), "UNKNOWN"), "NO"


def hpa_rows(rows, history, inventory, taken_on):
    out = []
    for row in rows:
        environment, in_inventory = environment_of(row, inventory)
        low, high = to_float(row["hpa_min"]), to_float(row["hpa_max"])
        now = to_float(row["replicas_now"])
        trend = parse_trend(row.pop("trend"))
        trend_max = max(trend) if trend else to_float(row["trend_max_export"])
        last_scale = azure.parse_datetime(row["last_scale_time"])
        days = (taken_on - last_scale.date()).days if last_scale else ""
        ci = history.get((cluster_key(row), row["namespace"], row["deployment"]), {})
        fixed = low is not None and high is not None and low >= high
        office_max = ci.get("ci_replicas_max_office", "")
        needed = replicas_needed(ci)
        min_above_need = needed is not None and low is not None and low > needed
        behaviour_ci = behaviour_from_average(ci.get("ci_replicas_avg_office"), low, high)
        record = {
            "cluster": row["cluster"], "cluster_id": row["cluster_id"], "environment": environment,
            "in_inventory": in_inventory, "namespace": row["namespace"], "hpa": row["hpa"], "deployment": row["deployment"],
            "platform_component": yes(aa.is_platform_namespace(row["namespace"])),
            "microservice": yes(aa.is_microservice(row["namespace"], row["deployment"])),
            "last_scale_time": azure.iso(last_scale), "days_since_last_scale": days,
            "last_scale_bucket": scale_bucket(days),
            "replicas_now": now, "hpa_min": low, "hpa_max": high, "trend_points": len(trend), "trend_max": trend_max,
            "range_used": round((trend_max - low) / (high - low), 4)
            if None not in (trend_max, low, high) and high > low else "",
            **{key: ci.get(key, "") for key in ("ci_months", "ci_replicas_avg_office", "ci_replicas_avg_offhours",
                                                "ci_replicas_max", "ci_replicas_max_office")},
            "office_max_reached_hpa_max": yes(not fixed and high is not None and office_max != "" and office_max >= high),
            "fixed_min_equals_max": yes(fixed),
            "at_max_now": yes(None not in (now, high) and now >= high),
            "trend_reached_max": yes(not fixed and None not in (trend_max, high) and trend_max >= high),
            "trend_never_above_min": yes(not fixed and None not in (trend_max, low) and trend_max <= low),
            "min_two_or_more": yes(low is not None and low >= 2),
            "ci_matched": yes(bool(ci)),
            "behaviour_export": behaviour(trend_max, low, high),
            "behaviour_ci": behaviour_ci,
            "no_scale_cause": no_scale_cause(fixed, behaviour_ci, min_above_need),
            "replicas_needed_by_cpu": "" if needed is None else needed,
            "min_above_need": yes(min_above_need),
            "min_above_need_beyond_ha": yes(needed is not None and low is not None and low > max(HA_REPLICAS, needed)),
        }
        out.append(record)
    return out


def deployment_rows(rows, history, inventory, hpas, clusters_with_ci, nodes):
    by_deployment = {(cluster_key(row), row["namespace"], row["deployment"]): row for row in hpas}
    out = []
    for row in rows:
        environment, in_inventory = environment_of(row, inventory)
        key = (cluster_key(row), row["namespace"], row["deployment"])
        ci, hpa = history.get(key, {}), by_deployment.get(key)
        cpu_request, cpu_limit = to_float(row["cpu_request_m"]), to_float(row["cpu_limit_m"])
        mem_request, mem_limit = to_float(row["mem_request_mb"]), to_float(row["mem_limit_mb"])
        request_vcpu = cpu_request / 1000 if cpu_request is not None else None
        request_gib = round(mem_request * MB_TO_GIB, 4) if mem_request is not None else None
        used = ci.get("ci_cpu_used_per_replica")
        ci_request = ci.get("ci_cpu_request_per_replica")
        node_vcpu, node_gib = nodes.get(cluster_key(row), ("", ""))
        record = {
            "cluster": row["cluster"], "cluster_id": row["cluster_id"], "environment": environment,
            "in_inventory": in_inventory, "namespace": row["namespace"], "deployment": row["deployment"],
            "service": row["service"],
            "platform_component": yes(aa.is_platform_namespace(row["namespace"])),
            "microservice": yes(aa.is_microservice(row["namespace"], row["deployment"])),
            "layer": layer_of(row["deployment"]),
            "bia": normalize_bia(row.get("bia")),
            "capabilities": ";".join(capabilities_of(row["deployment"])),
            "cpu_request_vcpu": request_vcpu, "cpu_limit_vcpu": cpu_limit / 1000 if cpu_limit is not None else None,
            "mem_request_gib": request_gib,
            "mem_limit_gib": round(mem_limit * MB_TO_GIB, 4) if mem_limit is not None else None,
            "has_hpa": yes(hpa), "hpa_min": hpa["hpa_min"] if hpa else "", "hpa_max": hpa["hpa_max"] if hpa else "",
            **{name: ci.get(name, "") for name in ("ci_months", "ci_replicas_avg_office", "ci_replicas_max",
                                                    "ci_cpu_used_per_replica", "ci_mem_used_gib_per_replica",
                                                    "ci_cpu_request_per_replica")},
            # Container Insights sums the requests of every container of the pod (Istio sidecar included); a
            # positive gap means the export only holds the application container.
            "ci_request_minus_export_vcpu": round(ci_request - request_vcpu, 4)
            if isinstance(ci_request, (int, float)) and request_vcpu is not None else "",
            "cpu_request_to_used": round(request_vcpu / used, 2) if request_vcpu and used else "",
            "cluster_has_ci": yes(cluster_key(row) in clusters_with_ci),
            "node_vcpu": node_vcpu, "node_mem_gib": node_gib,
        }
        record["no_usage_in_ci"] = yes(record["cluster_has_ci"] == "YES" and not ci)
        record["request_over_3x_used"] = yes(bool(used) and bool(request_vcpu) and request_vcpu > OVER_REQUEST_FACTOR * used)
        record["used_over_request"] = yes(bool(used) and request_vcpu is not None and used > request_vcpu)
        record["limit_10x_request"] = yes(bool(cpu_request) and cpu_limit is not None
                                          and cpu_limit >= HIGH_LIMIT_FACTOR * cpu_request)
        # Empty when the export has no request: unknown size is not "too big for Container Apps".
        record["fits_aca_by_request"] = "" if request_vcpu is None or request_gib is None else yes(
            request_vcpu <= ACA_MAX_VCPU and request_gib <= ACA_MAX_GIB)
        record["pod_too_big_for_node"] = yes(node_vcpu != "" and (
            (request_vcpu or 0) > NODE_SHARE_TOO_BIG * node_vcpu or (request_gib or 0) > NODE_SHARE_TOO_BIG * node_gib))
        out.append(record)
    return out


def ci_workloads_by_cluster(workloads, ids_by_name):
    """Per cluster key: application, tool and own-gateway workloads, microservice StatefulSets and batch Jobs, seen
    by Container Insights."""
    apps, tools, gateways, stateful, jobs = (defaultdict(set) for _ in range(5))
    for row in workloads:
        if aa.is_platform_namespace(row["namespace"]):
            continue
        key = cluster_key(dict(row, cluster_id=aa.cluster_id_of(row, ids_by_name)))
        name = (row["namespace"], row["workload"])
        if row.get("kind") == "Job" and aa.is_microservice(row["namespace"], row["workload"]):
            jobs[key].add(name)
        if row.get("kind") not in APP_KINDS:
            continue
        if aa.is_tool_workload(row["workload"]):
            (gateways if aa.is_app_gateway(row["workload"]) else tools)[key].add(name)
            continue
        apps[key].add(name)
        if row["kind"] == "StatefulSet":
            stateful[key].add(name)
    return apps, tools, gateways, stateful, jobs


def read_overrides(folder):
    """cluster_overrides.csv: cluster_id or cluster -> (category, reason)."""
    out = {}
    for row in aa.read_csv(os.path.join(folder, OVERRIDES_FILE)):
        category = str(row.get("category") or "").strip().upper()
        if category not in (aa.APPLICATION, aa.INFRASTRUCTURE, aa.NO_APPLICATIONS, aa.NO_DATA):
            continue
        key = str(row.get("cluster_id") or "").strip().lower() or "name:" + str(row.get("cluster") or "").strip().lower()
        out[key] = (category, row.get("reason") or "Set by hand in " + OVERRIDES_FILE)
    return out


def category_of(entry, ci_apps, ci_tools, ci_gateways, overrides):
    """(category, reason) of a cluster entry; see the module docstring."""
    by_hand = overrides.get(entry["key"]) or overrides.get("name:" + entry["cluster"].lower())
    if by_hand:
        return by_hand
    if entry["cod_app"] in aa.INFRASTRUCTURE_APPS:
        return aa.INFRASTRUCTURE, aa.INFRASTRUCTURE_APPS[entry["cod_app"]]
    if entry["microservices"]:
        return aa.APPLICATION, f"{entry['microservices']} microservices in the export"
    if ci_apps:
        return aa.APPLICATION, f"{ci_apps} application workloads in Container Insights, none in the export"
    if entry["gateways"] or ci_gateways:
        return aa.APPLICATION, "Only the application's own gateway, no microservice"
    if entry["tools"] or ci_tools:
        return aa.INFRASTRUCTURE, "Only tool workloads"
    if entry["in_export"] == "YES" or entry["has_ci"]:
        return aa.NO_APPLICATIONS, "No microservice in the export nor in Container Insights"
    return aa.NO_DATA, "Not in the export and no Container Insights data"


def cluster_rows(deployments, hpas, inventory, workloads, ids_by_name, clusters_with_ci, overrides):
    """One row per cluster: category, microservices, domains, HPAs, BIA and Container Apps fit."""
    def empty(key, name, cluster_id, environment, in_inventory, cod_app):
        return {"key": key, "cluster": name, "cluster_id": cluster_id, "environment": environment,
                "in_inventory": in_inventory, "cod_app": cod_app, "in_export": "NO", "microservices": 0, "tools": 0, "gateways": 0, "too_big": 0,
                "namespaces": set(), "hpas": 0, "levels": set(), "over_aca_size": 0}

    clusters = {cluster_id: empty(cluster_id, row["name"], cluster_id, row["environment"], "YES", row.get("cod_app", ""))
                for cluster_id, row in inventory.items()}
    for row in deployments:
        key = cluster_key(row)
        entry = clusters.setdefault(key, empty(key, row["cluster"], row["cluster_id"], row["environment"], "NO",
                                               azure.cod_app_from_name(row["cluster"])))
        entry["in_export"] = "YES"
        if row["microservice"] == "YES":
            entry["microservices"] += 1
            entry["namespaces"].add(row["namespace"])
            entry["levels"].add(row["bia"])
            entry["over_aca_size"] += row["fits_aca_by_request"] == "NO"
            entry["too_big"] += row["pod_too_big_for_node"] == "YES"
        elif row["platform_component"] == "NO":
            entry["gateways" if aa.is_app_gateway(row["deployment"]) else "tools"] += 1
    for row in hpas:
        entry = clusters.get(cluster_key(row))
        if entry and row["microservice"] == "YES":
            entry["hpas"] += 1
    ci_apps, ci_tools, ci_gateways, stateful, jobs = ci_workloads_by_cluster(workloads, ids_by_name)
    out = []
    for key, entry in clusters.items():
        entry["has_ci"] = key in clusters_with_ci
        category, reason = category_of(entry, len(ci_apps.get(key, ())), len(ci_tools.get(key, ())),
                                       len(ci_gateways.get(key, ())), overrides)
        statefulsets = len(stateful.get(key, ()))
        fit = 1 <= entry["microservices"] <= ADT_MAX_MICROSERVICES and not entry["over_aca_size"] and not statefulsets
        out.append({"cluster": entry["cluster"], "cluster_id": entry["cluster_id"], "cod_app": entry["cod_app"],
                    "environment": entry["environment"], "category": category, "category_reason": reason,
                    "in_inventory": entry["in_inventory"], "in_export": entry["in_export"],
                    "microservices": entry["microservices"], "tool_deployments": entry["tools"],
                    "ci_app_workloads": len(ci_apps.get(key, ())), "namespaces": len(entry["namespaces"]),
                    "hpas": entry["hpas"], "bia": most_critical(entry["levels"]),
                    "microservices_over_aca_size": entry["over_aca_size"], "statefulsets": statefulsets,
                    "batch_jobs": len(jobs.get(key, ())), "pods_too_big_for_node": entry["too_big"],
                    "aca_technical_fit": yes(fit and category == aa.APPLICATION)})
    return sorted(out, key=lambda value: (value["environment"], value["category"], -value["microservices"],
                                          value["cluster"]))


def namespace_codes(workloads, inventory, ids_by_name):
    """PROD namespace -> app code of the clusters that run it (Container Insights), when they share one code."""
    codes = defaultdict(set)
    for row in workloads:
        if row["environment"] != "PROD" or aa.is_platform_namespace(row["namespace"]):
            continue
        cluster = inventory.get(aa.cluster_id_of(row, ids_by_name), {})
        codes[row["namespace"].lower()].add(cluster.get("cod_app") or azure.cod_app_from_name(row["cluster"]))
    return {namespace: next(iter(values)) for namespace, values in codes.items() if len(values) == 1 and "" not in values}


def api_code(row, by_namespace):
    """App code of an API: the cluster that runs the namespace of a .svc.cluster.local backend, else the inventory's."""
    host = str(row.get("backend_host") or "").lower()
    if host.endswith(".svc.cluster.local"):
        namespace = host.split(".")[1]
        if by_namespace.get(namespace):
            return by_namespace[namespace]
    return row.get("backend_cod_app") or ""


def app_traffic(apis, monthly, profile, by_namespace):
    """Per app code, last month, from one APIM layer: the first of aa_common.TRAFFIC_LAYERS (BS, UX, Open Banking,
    Yape, Core) with traffic that month. Layers are never added: a call goes on from one gateway to the next and
    would count twice. Only APIs whose backend runs on AKS count.

    Returns (traffic, presence): traffic per app code with traffic (APIs with traffic, requests, office share, peak TPS
    as the sum of each API's 5-minute peak (an upper bound), busiest-hour TPS (the APIs added hour by hour), weekday
    hours with traffic, off-hours active share and the layer used); presence is app code -> layers where it has
    AKS-backed APIs, and the layers the metrics cover, for traffic_status."""
    if apis and "apim_layer" not in apis[0]:
        raise SystemExit("apim_apis.csv has no apim_layer column: run aa_inventory.py again")
    api_info = {(row["apim"], row["api_id"]): (api_code(row, by_namespace), row["apim_layer"])
                for row in apis if (row.get("backend_platform") or aa.backend_platform(row.get("backend_host"))) == "AKS"}
    presence = defaultdict(set)
    for code, layer in api_info.values():
        if code:
            presence[code].add(layer)
    measured_layers = {api_info[key][1] for key in ((row["apim"], row["api_id"]) for row in monthly) if key in api_info}
    last = max((row["month"] for row in monthly), default="")
    traffic = defaultdict(lambda: {"apis": 0, "requests": 0.0, "office": 0.0, "peak_5min": 0.0, "offhours_active": 0.0})
    for row in monthly:
        code, layer = api_info.get((row["apim"], row["api_id"]), ("", ""))
        if row["month"] != last or not code or layer not in aa.TRAFFIC_LAYERS:
            continue
        entry = traffic[(code, layer)]
        entry["apis"] += 1
        entry["requests"] += to_float(row.get("requests")) or 0
        entry["office"] += to_float(row.get("office_requests")) or 0
        entry["peak_5min"] = entry["peak_5min"] + (to_float(row.get("peak_5min_requests")) or 0)
        # Off-hours intervals of the month: every 5-minute interval minus the office ones. The app's share is the
        # busiest API's, a floor: two APIs active in the same interval are still one interval.
        offhours_buckets = monthrange(int(last[:4]), int(last[5:7]))[1] * 288 - \
            (to_float(row.get("office_hours_in_window")) or 0) * 12
        if offhours_buckets > 0:
            entry["offhours_active"] = max(entry["offhours_active"],
                                           (to_float(row.get("offhours_active_buckets")) or 0) / offhours_buckets)
    chosen = {}
    for code, layer in sorted(traffic, key=lambda key: aa.TRAFFIC_LAYERS.index(key[1])):
        if code not in chosen and traffic[(code, layer)]["requests"]:
            chosen[code] = layer
    hourly = defaultdict(float)
    for row in profile:
        code, layer = api_info.get((row["apim"], row["api_id"]), ("", ""))
        if code and chosen.get(code) == layer:
            hourly[(code, row["day_type"], int(row["hour"]))] += to_float(row.get("avg_requests_per_hour")) or 0
    out = {}
    for code, layer in chosen.items():
        entry = traffic[(code, layer)]
        if not entry["requests"]:
            continue
        out[code] = dict(traffic_fields(code, entry, hourly), month=last, traffic_layer=layer)
    return out, {"presence": presence, "measured_layers": measured_layers}


def traffic_fields(code, entry, hourly):
    weekday = [hourly.get((code, "weekday", hour), 0.0) for hour in range(24)]
    busiest = max(weekday + [hourly.get((code, "weekend", hour), 0.0) for hour in range(24)])
    office = {("weekday", hour): weekday[hour] for hour in range(aa.OFFICE_START_HOUR, aa.OFFICE_END_HOUR)}
    offhours = {(day_type, hour): hourly.get((code, day_type, hour), 0.0)
                for day_type in ("weekday", "weekend") for hour in range(24)
                if (day_type, hour) not in office}
    top = max({**office, **offhours}.items(), key=lambda item: item[1])[0]
    return {"apis_with_traffic": entry["apis"], "requests_month": int(entry["requests"]),
            "office_share": round(entry["office"] / entry["requests"], 3) if entry["requests"] else "",
            "peak_tps_upper_bound": round(entry["peak_5min"] / 300, 1),
            "busiest_hour_tps": round(busiest / 3600, 3),
            "busiest_hour": f"{top[0]} {top[1]:02d}:00" if busiest else "",
            "busiest_hour_in_office": ("YES" if top in office else "NO") if busiest else "",
            "busiest_office_hour_tps": round(max(office.values()) / 3600, 3),
            "busiest_offhours_hour_tps": round(max(offhours.values()) / 3600, 3),
            "weekday_hours_with_traffic": sum(value >= ACTIVE_HOUR_SHARE * busiest for value in weekday)
            if busiest else 0,
            "offhours_active_share": round(entry["offhours_active"], 3)}


def traffic_status(code, traffic, coverage):
    """MEASURED; ZERO_IN_MONTH (APIs on a measured APIM, no requests); LAYER_NOT_MEASURED (APIs only on layers the
    metrics did not cover); NO_APIM_API (no AKS-backed API on any APIM)."""
    if code in traffic:
        return "MEASURED"
    layers = coverage["presence"].get(code, set())
    if layers & coverage["measured_layers"] & set(aa.TRAFFIC_LAYERS):
        return "ZERO_IN_MONTH"
    return "LAYER_NOT_MEASURED" if layers else "NO_APIM_API"


def app_cost(cost, application_ids):
    """AKS amortized cost of the last month per cluster id."""
    last = max((row["month"] for row in cost if row.get("owner_type") == "AKS"), default="")
    total = defaultdict(float)
    for row in cost:
        if row.get("owner_type") == "AKS" and row["month"] == last and row["owner_id"].lower() in application_ids:
            total[row["owner_id"].lower()] += to_float(row.get("cost_usd")) or 0
    return last, total


def adt_runtime(app):
    """Runtime the ADT_019 tree gives an application, with the reasons; see docs/adt-019-runtime.md."""
    if app["busiest_hour_tps"] != "" and app["busiest_hour_tps"] > ADT_MAX_TPS:
        return "AKS", f"Busiest hour above {ADT_MAX_TPS} TPS"
    expected = max(app["microservices"], app["microservices_projected"] or 0)
    if (app["weekday_hours_with_traffic"] != "" and app["weekday_hours_with_traffic"] > ADT_MAX_ACTIVE_HOURS
            and expected > ADT_MAX_MICROSERVICES):
        grows = "" if app["microservices"] > ADT_MAX_MICROSERVICES else f" (expected {expected} with DESA)"
        return "AKS", (f"Traffic more than {ADT_MAX_ACTIVE_HOURS} h a day and more than {ADT_MAX_MICROSERVICES} "
                       f"microservices{grows}")
    if app["offhours_active_share"] != "" and app["offhours_active_share"] > OFFHOURS_ACTIVE_MAX:
        return "AKS", (f"Traffic in {app['offhours_active_share']:.0%} of the off-hours time: cannot scale to zero "
                       "(never warm 24/7)")
    if app["microservices_over_aca_size"]:
        return "AKS", f"{app['microservices_over_aca_size']} microservices request more than 4 vCPU / 8 GiB"
    # StatefulSets and batch jobs are not criteria for now (no PROD application has them); the columns stay.
    reasons = []
    if app["bia"] in CRITICAL_BIA:
        reasons.append("Critical (BIA ALTA or MUY ALTA)")
    if app["busiest_hour_tps"] == "":
        reasons.append("No APIM traffic mapped to the application")
    if expected > ADT_MAX_MICROSERVICES >= app["microservices"]:
        reasons.append(f"Grows to {expected} microservices with DESA: review AKS")
    return ("ACA_WITH_REVIEW", "; ".join(reasons)) if reasons else ("ACA", "Within every ADT_019 threshold")


def base_name(deployment):
    """Deployment name without its version suffix: channel-x-v1 and channel-x-v2 are one microservice."""
    return re.sub(r"(-v\d+)+$", "", str(deployment or "").lower())


def microservice_names(deployments, cluster_table, environment):
    """App code -> base names of the microservices in that environment's non-infrastructure clusters of the export."""
    code_of = {row["cluster_id"] or "name:" + row["cluster"].lower(): row["cod_app"] or row["cluster"]
               for row in cluster_table if row["environment"] == environment and row["category"] != aa.INFRASTRUCTURE}
    out = defaultdict(set)
    for row in deployments:
        code = code_of.get(cluster_key(row))
        if code and row["environment"] == environment and row["microservice"] == "YES":
            out[code].add(base_name(row["deployment"]))
    return out


NEW_CLUSTER_DAYS = 120
# Every application cluster has an ingress load balancer (APIM reaches it that way); it only explains the missing
# traffic of an application without any API on APIM.
BALANCER_NOTE = "Own ingress load balancer and no API on APIM: callers reach it directly"


def traffic_notes(cluster_table, inventory, services, prod_names):
    """App code -> what the data shows about a PROD application without traffic: a new cluster, only event
    consumers, its own ingress load balancer (callers reach it without APIM), or no microservice running."""
    balancers = {key for key, kinds in services.items() if "LoadBalancer" in kinds}
    notes = defaultdict(list)
    for row in cluster_table:
        if row["environment"] != "PROD" or row["category"] != aa.APPLICATION:
            continue
        code = row["cod_app"] or row["cluster"]
        cluster = inventory.get(row["cluster_id"], {})
        created, month = cluster.get("created_at", ""), cluster.get("measurement_month", "")
        if created and month and (date.fromisoformat(month + "-01") - date.fromisoformat(created)).days < NEW_CLUSTER_DAYS:
            notes[code].append(f"New cluster (created {created})")
        if row["cluster_id"] in balancers:
            notes[code].append(BALANCER_NOTE)
    for code, names in prod_names.items():
        if names and all(name.startswith("async") for name in names):
            notes[code].append("Only event consumers (async-*): load comes from queues, not APIM")
    return notes


def note_of(notes, status, microservices):
    kept = [note for note in notes if note != BALANCER_NOTE or status == "NO_APIM_API"]
    if not microservices:
        kept.append("No microservice running")
    return "; ".join(dict.fromkeys(kept))


def services_by_cluster(rows, ids_by_name):
    """Cluster id -> service types seen in Container Insights (aks_services.csv)."""
    out = defaultdict(set)
    for row in rows:
        out[aa.cluster_id_of(row, ids_by_name)].add(row.get("service_type", ""))
    return out


def projection(prod, desa):
    """DESA microservices and the projection for PROD: PROD plus the DESA microservices not in PROD yet (by base
    name). Empty when the export has no DESA cluster for the application."""
    if desa is None:
        return {"microservices_desa": "", "microservices_new_in_desa": "", "microservices_projected": ""}
    return {"microservices_desa": len(desa), "microservices_new_in_desa": len(desa - prod),
            "microservices_projected": len(prod | desa)}


def application_rows(cluster_table, hpas, deployments, traffic, coverage, cost_by_cluster, inventory, services):
    """One row per PROD application (app code with APPLICATION clusters)."""
    groups = defaultdict(list)
    for row in cluster_table:
        if row["environment"] == "PROD" and row["category"] == aa.APPLICATION:
            groups[row["cod_app"] or row["cluster"]].append(row)
    prod_names = microservice_names(deployments, cluster_table, "PROD")
    desa_names = microservice_names(deployments, cluster_table, "DESA")
    notes = traffic_notes(cluster_table, inventory, services, prod_names)
    hpa_causes = defaultdict(list)
    for row in hpas:
        if row["environment"] == "PROD" and row["microservice"] == "YES":
            hpa_causes[cluster_key(row)].append(row["no_scale_cause"])
    out = []
    for code, rows in sorted(groups.items()):
        keys = [row["cluster_id"] or "name:" + row["cluster"].lower() for row in rows]
        causes = [cause for key in keys for cause in hpa_causes.get(key, []) if cause != "NO_DATA"]
        app = {"cod_app": code, "clusters": len(rows), "cluster_names": ";".join(row["cluster"] for row in rows),
               "cost_month_usd": round(sum(cost_by_cluster.get(key, 0) for key in keys), 2),
               "bia": most_critical({row["bia"] for row in rows}),
               **{field: sum(row[field] for row in rows) for field in (
                   "microservices", "namespaces", "hpas", "microservices_over_aca_size", "statefulsets", "batch_jobs",
                   "pods_too_big_for_node")},
               "hpas_not_scaling_share": round(sum(cause != "SCALES" for cause in causes) / len(causes), 3) if causes else "",
               "possible_monolith": yes(sum(row["microservices"] for row in rows) == 1),
               **projection(prod_names.get(code, set()), desa_names.get(code)),
               "traffic_status": traffic_status(code, traffic, coverage),
               "traffic_note": "" if code in traffic else note_of(notes.get(code, []), traffic_status(code, traffic, coverage),
                                                                  sum(row["microservices"] for row in rows)),
               **{field: traffic.get(code, {}).get(field, "") for field in (
                   "traffic_layer", "apis_with_traffic", "requests_month", "office_share", "peak_tps_upper_bound", "busiest_hour_tps",
                   "busiest_hour", "busiest_hour_in_office", "busiest_office_hour_tps", "busiest_offhours_hour_tps",
                   "weekday_hours_with_traffic", "offhours_active_share")}}
        app["adt_runtime"], app["adt_reason"] = adt_runtime(app)
        out.append(app)
    return out


def capability_rows(deployments, application_keys):
    """PROD namespaces (domains) and microservices of APPLICATION clusters per capability keyword."""
    out = []
    for capability in CAPABILITIES:
        matches = [row for row in deployments if row["environment"] == "PROD" and row["microservice"] == "YES"
                   and cluster_key(row) in application_keys and capability in row["capabilities"].split(";")]
        per_namespace = defaultdict(int)
        for row in matches:
            per_namespace[row["namespace"]] += 1
        out.append({"capability": capability, "namespaces_prod": len(per_namespace), "deployments_prod": len(matches),
                    "namespaces": ", ".join(f"{name}({count})" for name, count in sorted(per_namespace.items()))})
    return out


def with_cod_app(rows, inventory):
    codes = {key: row.get("cod_app", "") for key, row in inventory.items()}
    return aa.with_cod_app(rows, "cluster_id", lambda row: codes.get(row["cluster_id"]) or azure.cod_app_from_name(row["cluster"]))


def prod_application(rows, application_keys):
    return [row for row in rows if row["environment"] == "PROD" and cluster_key(row) in application_keys
            and row["microservice"] == "YES"]


def flag_counts(rows, flags):
    return {"rows": len(rows), **{flag: sum(row[flag] == "YES" for row in rows) for flag in flags}}


def counts(rows, field):
    return dict(sorted(Counter(row[field] for row in rows).items()))


def write_export_history(folder, delimiter, taken_on, cluster_table, deployments, application_keys):
    """Add (or replace) the export day in export_history.csv: PROD clusters by category and their microservices."""
    path = os.path.join(folder, EXPORT_HISTORY_FILE)
    rows = [row for row in aa.read_csv(path) if row["snapshot_date"] != taken_on.isoformat()]
    prod = [row for row in cluster_table if row["environment"] == "PROD"]
    by_category = Counter(row["category"] for row in prod)
    rows.append({"snapshot_date": taken_on.isoformat(),
                 "prod_application_clusters": by_category.get(aa.APPLICATION, 0),
                 "prod_microservices": len(prod_application(deployments, application_keys)),
                 "prod_deployments_in_export": sum(row["environment"] == "PROD" for row in deployments),
                 **{f"prod_{category.lower()}_clusters": by_category.get(category, 0)
                    for category in (aa.INFRASTRUCTURE, aa.NO_APPLICATIONS, aa.NO_DATA)}})
    azure.Output(folder, delimiter).csv(EXPORT_HISTORY_FILE, sorted(rows, key=lambda row: row["snapshot_date"]))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--inputs", default=DEFAULT_INPUTS, help="Folder with the reportHpa and reportDeployment exports")
    parser.add_argument("--snapshot-date", default=None,
                        help="Day the exports were taken (YYYY-MM-DD). Default: modification date of the HPA file")
    parser.add_argument("--inventory", default=aa.results_path("aa-inventory"), help="Folder with the aa_inventory.py output")
    parser.add_argument("--container-insights", default=aa.results_path("aa-container-insights"),
                        help="Folder with the aa_container_insights.py output")
    parser.add_argument("--metrics", default=aa.results_path("aa-metrics"), help="Folder with the aa_metrics.py output")
    parser.add_argument("--cost", default=aa.results_path("aa-cost"), help="Folder with the aa_cost.py output")
    parser.add_argument("--output-dir", default=aa.results_path("aa-k8s-snapshot"))
    parser.add_argument("--history-dir", default=aa.results_path("aa-history"),
                        help="Folder of the growing history; this run adds or replaces its export day")
    parser.add_argument("--delimiter", default=",", help="CSV delimiter of the outputs")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)

    hpa_path, deployment_path = find_input(args.inputs, "reportHpa"), find_input(args.inputs, "reportDeployment")
    taken_on = snapshot_date(hpa_path, args.snapshot_date)
    clusters = aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))
    ids_by_name = aa.cluster_ids_by_name(clusters)
    inventory = {row["id"].lower(): row for row in clusters}
    workloads = aa.read_csv(os.path.join(args.container_insights, "aks_workloads_monthly.csv"))
    if not workloads:
        azure.LOG.warning("No Container Insights results in %s: the history columns stay empty", args.container_insights)
    history = ci_history(workloads, ids_by_name)
    clusters_with_ci = {cluster_key(dict(row, cluster_id=aa.cluster_id_of(row, ids_by_name))) for row in workloads}
    nodes = node_sizes(aa.read_csv(os.path.join(args.metrics, "aks_usage_monthly.csv")), ids_by_name)
    shared_names = sorted(name for name, ids in ids_by_name.items() if len(ids) > 1)

    hpas = hpa_rows(select(read_rows(hpa_path), HPA_COLUMNS, hpa_path, ids_by_name), history, inventory, taken_on)
    deployments = deployment_rows(select(read_rows(deployment_path), DEPLOYMENT_COLUMNS, deployment_path, ids_by_name),
                                  history, inventory, hpas, clusters_with_ci, nodes)
    for rows in (hpas, deployments):
        for row in rows:
            row.pop("trend_max_export", None)
            row.pop("environment_export", None)
    cluster_table = cluster_rows(deployments, hpas, inventory, workloads, ids_by_name, clusters_with_ci,
                                 read_overrides(args.inputs))
    application = {row["cluster_id"] or "name:" + row["cluster"].lower() for row in cluster_table
                   if row["category"] == aa.APPLICATION}

    export_keys = {(cluster_key(row), row["namespace"], row["deployment"]) for row in deployments}
    ci_only = {key for key in history if key[0] in application and aa.is_microservice(key[1], key[2])
               and key not in export_keys}
    app_hpas, app_deployments = prod_application(hpas, application), prod_application(deployments, application)
    gaps = [row["ci_request_minus_export_vcpu"] for row in app_deployments if row["ci_request_minus_export_vcpu"] != ""]

    write_export_history(args.history_dir, args.delimiter, taken_on, cluster_table, deployments, application)
    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("hpa_snapshot.csv", with_cod_app(hpas, inventory))
    output.csv("deployment_snapshot.csv", with_cod_app(deployments, inventory))
    output.csv("cluster_snapshot.csv", cluster_table)
    output.csv("capability_domains.csv", capability_rows(deployments, application))
    traffic, coverage = app_traffic(aa.read_csv(os.path.join(args.inventory, "apim_apis.csv")),
                          aa.read_csv(os.path.join(args.metrics, "apim_api_monthly.csv")),
                          aa.read_csv(os.path.join(args.metrics, "apim_hourly_profile.csv")),
                          namespace_codes(workloads, inventory, ids_by_name))
    cost_month, cost_by_cluster = app_cost(aa.read_csv(os.path.join(args.cost, "cost_by_resource.csv")), application)
    services = services_by_cluster(aa.read_csv(os.path.join(args.container_insights, "aks_services.csv")), ids_by_name)
    applications = application_rows(cluster_table, hpas, deployments, traffic, coverage, cost_by_cluster, inventory,
                                    services)
    output.csv("application_snapshot.csv", applications)
    prod_clusters = [row for row in cluster_table if row["environment"] == "PROD"]
    output.json("snapshot_summary.json", {
        "generated_at": azure.iso(azure.utc_now()),
        "snapshot_date": taken_on.isoformat(),
        "inputs": [{"file": os.path.basename(path), "rows": count,
                    "modified": datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="minutes")}
                   for path, count in ((hpa_path, len(hpas)), (deployment_path, len(deployments)))],
        "units": {"cpu": "millicores in the export, vCPU in the outputs",
                  "memory": "MB (10^6 bytes) in the export, GiB in the outputs"},
        "container_insights_months": sorted({row["month"] for row in workloads}),
        "cluster_names_shared_by_several_clusters": shared_names,
        "prod_clusters_by_category": counts(prod_clusters, "category"),
        "prod_applications": len(applications), "cost_month": cost_month,
        "prod_applications_by_adt_runtime": counts(applications, "adt_runtime"),
        "prod_applications_by_traffic_status": counts(applications, "traffic_status"),
        "prod_applications_by_traffic_layer": counts(applications, "traffic_layer"),
        "prod_applications_with_desa": sum(row["microservices_desa"] != "" for row in applications),
        "prod_applications_over_15_only_with_desa": sum(
            row["microservices_projected"] != "" and row["microservices_projected"] > ADT_MAX_MICROSERVICES
            >= row["microservices"] for row in applications),
        "clusters_not_in_inventory": sorted({row["cluster"] for row in hpas + deployments if row["in_inventory"] == "NO"}),
        "prod_application": {
            "microservices": len(app_deployments),
            "export_microservices_without_ci_history": sum(row["ci_months"] == "" for row in app_deployments),
            "ci_microservices_missing_from_export": len(ci_only),
            "ci_request_minus_export_vcpu_median": round(statistics.median(gaps), 4) if gaps else "",
            "hpa": flag_counts(app_hpas, ["fixed_min_equals_max", "office_max_reached_hpa_max", "min_two_or_more",
                                          "ci_matched"]),
            "hpa_behaviour_ci": counts(app_hpas, "behaviour_ci"),
            "hpa_no_scale_cause": counts(app_hpas, "no_scale_cause"),
            "hpa_last_scale_bucket": counts(app_hpas, "last_scale_bucket"),
            "deployment": flag_counts(app_deployments, ["has_hpa", "no_usage_in_ci", "request_over_3x_used",
                                                        "used_over_request", "limit_10x_request", "fits_aca_by_request",
                                                        "pod_too_big_for_node"]),
        },
    })


if __name__ == "__main__":
    main()
