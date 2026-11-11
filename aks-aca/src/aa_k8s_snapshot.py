#!/usr/bin/env python3
"""HPA and deployment configuration of one day, joined with the replica and usage history of Container Insights.

Reads two exports from the platform team, kept outside git in --inputs (.local/aks-aca-inputs by default):
  reportHpa*.csv|xlsx         one row per HPA: min, max, current replicas, last scale time, replica trend
  reportDeployment*.csv|xlsx  one row per deployment: CPU and memory requests and limits

The exports are a snapshot (current state, no history). The history comes from aa_container_insights.py
(replicas and usage per workload and month), so run it first. No Azure call is made here.

Only the columns listed in HPA_COLUMNS and DEPLOYMENT_COLUMNS are read; the rest of each export (subscription,
application, owner and provisioning columns) is never copied to the results. `bia` is the business impact
level of the application (MUY ALTA to BAJA); a cluster takes the highest level among its microservices.

A cluster is a technical fit for Container Apps when it has 1 to 15 microservices, every one fits a Container
Apps replica by its request (4 vCPU / 8 GiB) and Container Insights shows no StatefulSet outside platform
namespaces. Requests per second per application are not part of the rule (APIs are not mapped to clusters).

Replicas needed by CPU = office CPU used by all replicas / (request per replica x 70%, a typical HPA target).
A minimum above it is more than the load asks for; above max(2, needed) it is more even keeping two replicas
for availability.

HPA behaviour over 3 months uses the office-hours average of replicas, not the hourly maximum: Container
Insights counts the pods seen in each hour, so every rolling deployment briefly adds the new pods to the old
ones and the hourly maximum overstates scaling. At the minimum = average <= 110% of minReplicas; near the
maximum = average >= 90% of maxReplicas.

Units as exported: CPU in millicores; memory in MB (10^6 bytes, so 2Gi shows as 2147.483648).

Outputs (in --output-dir):
  hpa_snapshot.csv         one row per HPA with its Container Insights history and flags
  deployment_snapshot.csv  one row per deployment with requests vs measured usage and flags
  cluster_snapshot.csv     one row per cluster: microservices, namespaces and HPAs (inventory clusters
                           missing from the export get 0)
  capability_domains.csv   per capability keyword (notification, token, login...): PROD namespaces and
                           deployments whose name carries it
  snapshot_summary.json    input files, snapshot date, rows without a match and counts per flag (PROD)

A microservice is a deployment outside platform namespaces (PLATFORM_NAMESPACES, aks-istio*, istio-*,
*ingress*). A domain is a namespace. The layer is the first word of the deployment name.
"""

import argparse
import csv
import glob
import io
import math
import os
from collections import defaultdict
from datetime import date, datetime

import aa_common as aa

azure = aa.azure

DEFAULT_INPUTS = os.path.join(os.path.dirname(aa.TOPIC_DIR), ".local", "aks-aca-inputs")
MB_TO_GIB = 1e6 / 1024 ** 3
ACA_MAX_VCPU, ACA_MAX_GIB = 4, 8
OVER_REQUEST_FACTOR = 3
HIGH_LIMIT_FACTOR = 10

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
PLATFORM_NAMESPACES = set(aa.SYSTEM_NAMESPACES) | {"flux-system", "argocd", "cert-manager", "keda", "monitoring"}
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


def normalize_bia(value):
    text = " ".join(str(value or "").upper().split())
    return text if text and text != "-" else NO_BIA


def most_critical(levels):
    """Highest BIA among the given levels; unknown labels rank below the known ones, SIN BIA last."""
    known = [level for level in levels if level != NO_BIA]
    if not known:
        return NO_BIA
    return min(known, key=lambda level: BIA_ORDER.index(level) if level in BIA_ORDER else len(BIA_ORDER))


def is_platform(namespace):
    name = str(namespace or "").lower()
    return name in PLATFORM_NAMESPACES or name.startswith(("aks-istio", "istio-")) or "ingress" in name


def layer_of(deployment):
    first = str(deployment or "").lower().split("-")[0]
    return first if first in LAYERS else "otros"


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
        return "SIN_DATO"
    if low >= high:
        return "FIJO"
    if average in (None, ""):
        return "SIN_DATO"
    if average <= low * AT_MIN_FACTOR:
        return "NUNCA_ESCALA"
    return "LLEGA_AL_MAX" if average >= high * NEAR_MAX_FACTOR else "ESCALA"


def replicas_needed(ci):
    """Replicas the measured office CPU needs at CPU_TARGET of the request, or None without data."""
    used, request, replicas = (ci.get(key) for key in ("ci_cpu_used_per_replica", "ci_cpu_request_per_replica",
                                                       "ci_replicas_avg_office"))
    if not all(isinstance(value, (int, float)) for value in (used, request, replicas)) or request <= 0:
        return None
    return max(1, math.ceil(used * replicas / (request * CPU_TARGET) - 1e-9))


def behaviour(peak, low, high):
    """FIJO (min = max), NUNCA_ESCALA (peak <= min), LLEGA_AL_MAX (peak >= max), ESCALA, or SIN_DATO."""
    if None in (low, high):
        return "SIN_DATO"
    if low >= high:
        return "FIJO"
    if peak in (None, ""):
        return "SIN_DATO"
    if peak <= low:
        return "NUNCA_ESCALA"
    return "LLEGA_AL_MAX" if peak >= high else "ESCALA"


def yes(value):
    return "SÍ" if value else "NO"


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


def select(rows, columns, path):
    """Keep only the wanted columns, renamed; stop with the list of missing ones."""
    present = {key.strip().lower(): key for key in (rows[0] if rows else {})}
    missing = [source for source in columns.values() if source.lower() not in present]
    if missing:
        raise SystemExit(f"{os.path.basename(path)} lacks the columns {', '.join(missing)}. "
                         f"Columns found: {', '.join(present.values())}")
    out = []
    for row in rows:
        record = {name: row[present[source.lower()]] for name, source in columns.items()}
        if not record.get("cluster"):
            record["cluster"] = str(record.get("cluster_id") or "").rstrip("/").split("/")[-1]
        record.pop("cluster_id", None)
        out.append(record)
    return out


def snapshot_date(path, override):
    if override:
        return date.fromisoformat(override)
    return datetime.fromtimestamp(os.path.getmtime(path)).date()


def ci_history(rows):
    """Container Insights rows grouped by (cluster, namespace, workload): replicas and per-replica usage."""
    history = defaultdict(lambda: {"months": set(), "office_hours": 0.0, "offhours_hours": 0.0, "replicas_max": 0.0,
                                   "sums": defaultdict(float)})
    weighted = ("replicas_avg", "cpu_used_per_replica", "cpu_request_per_replica", "mem_used_gib_per_replica",
                "mem_request_gib_per_replica")
    for row in rows:
        key = (row["cluster"].lower(), row["namespace"], row["workload"])
        entry = history[key]
        hours = to_float(row.get("hours")) or 0
        entry["months"].add(row["month"])
        entry["replicas_max"] = max(entry["replicas_max"], to_float(row.get("replicas_max")) or 0)
        period = "office" if row["period"] == "OFFICE" else "offhours"
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
            "ci_cpu_used_per_replica": average("office", "cpu_used_per_replica"),
            "ci_cpu_request_per_replica": average("office", "cpu_request_per_replica"),
            "ci_mem_used_gib_per_replica": average("office", "mem_used_gib_per_replica"),
        }
    return result


def environment_of(row, inventory):
    cluster = inventory.get(row["cluster"].lower())
    if cluster:
        return cluster["environment"], "SÍ"
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
        ci = history.get((row["cluster"].lower(), row["namespace"], row["deployment"]), {})
        fixed = low is not None and high is not None and low >= high
        record = {
            "cluster": row["cluster"], "environment": environment, "in_inventory": in_inventory,
            "namespace": row["namespace"], "hpa": row["hpa"], "deployment": row["deployment"],
            "platform_component": yes(is_platform(row["namespace"])),
            "last_scale_time": azure.iso(last_scale), "days_since_last_scale": (taken_on - last_scale.date()).days
            if last_scale else "",
            "replicas_now": now, "hpa_min": low, "hpa_max": high, "trend_points": len(trend), "trend_max": trend_max,
            "range_used": round((trend_max - low) / (high - low), 4)
            if None not in (trend_max, low, high) and high > low else "",
            **{key: ci.get(key, "") for key in ("ci_months", "ci_replicas_avg_office", "ci_replicas_avg_offhours",
                                                "ci_replicas_max")},
            "fixed_min_equals_max": yes(fixed),
            "at_max_now": yes(None not in (now, high) and now >= high),
            "trend_reached_max": yes(not fixed and None not in (trend_max, high) and trend_max >= high),
            "trend_never_above_min": yes(not fixed and None not in (trend_max, low) and trend_max <= low),
            "min_two_or_more": yes(low is not None and low >= 2),
            "ci_matched": yes(bool(ci)),
            "behaviour_export": behaviour(trend_max, low, high),
            "behaviour_ci": behaviour_from_average(ci.get("ci_replicas_avg_office"), low, high),
            "replicas_needed_by_cpu": "" if (needed := replicas_needed(ci)) is None else needed,
            "min_above_need": yes(needed is not None and low is not None and low > needed),
            "min_above_need_beyond_ha": yes(needed is not None and low is not None and low > max(HA_REPLICAS, needed)),
        }
        out.append(record)
    return out


def deployment_rows(rows, history, inventory, hpas, clusters_with_ci):
    by_deployment = {(row["cluster"].lower(), row["namespace"], row["deployment"]): row for row in hpas}
    out = []
    for row in rows:
        environment, in_inventory = environment_of(row, inventory)
        key = (row["cluster"].lower(), row["namespace"], row["deployment"])
        ci, hpa = history.get(key, {}), by_deployment.get(key)
        cpu_request, cpu_limit = to_float(row["cpu_request_m"]), to_float(row["cpu_limit_m"])
        mem_request, mem_limit = to_float(row["mem_request_mb"]), to_float(row["mem_limit_mb"])
        request_vcpu = cpu_request / 1000 if cpu_request is not None else None
        used = ci.get("ci_cpu_used_per_replica")
        record = {
            "cluster": row["cluster"], "environment": environment, "in_inventory": in_inventory,
            "namespace": row["namespace"], "deployment": row["deployment"], "service": row["service"],
            "platform_component": yes(is_platform(row["namespace"])), "layer": layer_of(row["deployment"]),
            "bia": normalize_bia(row.get("bia")),
            "capabilities": ";".join(capabilities_of(row["deployment"])),
            "cpu_request_vcpu": request_vcpu, "cpu_limit_vcpu": cpu_limit / 1000 if cpu_limit is not None else None,
            "mem_request_gib": round(mem_request * MB_TO_GIB, 4) if mem_request is not None else None,
            "mem_limit_gib": round(mem_limit * MB_TO_GIB, 4) if mem_limit is not None else None,
            "has_hpa": yes(hpa), "hpa_min": hpa["hpa_min"] if hpa else "", "hpa_max": hpa["hpa_max"] if hpa else "",
            **{name: ci.get(name, "") for name in ("ci_months", "ci_replicas_avg_office", "ci_replicas_max",
                                                    "ci_cpu_used_per_replica", "ci_mem_used_gib_per_replica")},
            "cpu_request_to_used": round(request_vcpu / used, 2) if request_vcpu and used else "",
            "cluster_has_ci": yes(row["cluster"].lower() in clusters_with_ci),
        }
        record["no_usage_in_ci"] = yes(record["cluster_has_ci"] == "SÍ" and not ci)
        record["request_over_3x_used"] = yes(bool(used) and bool(request_vcpu) and request_vcpu > OVER_REQUEST_FACTOR * used)
        record["used_over_request"] = yes(bool(used) and request_vcpu is not None and used > request_vcpu)
        record["limit_10x_request"] = yes(bool(cpu_request) and cpu_limit is not None
                                          and cpu_limit >= HIGH_LIMIT_FACTOR * cpu_request)
        record["fits_aca_by_request"] = yes(request_vcpu is not None and request_vcpu <= ACA_MAX_VCPU and
                                            record["mem_request_gib"] is not None and record["mem_request_gib"] <= ACA_MAX_GIB)
        out.append(record)
    return out


def statefulsets_by_cluster(workloads):
    """StatefulSets outside platform namespaces seen by Container Insights, per cluster."""
    names = defaultdict(set)
    for row in workloads:
        if row.get("kind") == "StatefulSet" and not is_platform(row["namespace"]):
            names[row["cluster"].lower()].add((row["namespace"], row["workload"]))
    return {cluster: len(items) for cluster, items in names.items()}


def cluster_rows(deployments, hpas, inventory, statefulsets):
    """Microservices, domains, HPAs, BIA and Container Apps fit per cluster; inventory clusters missing from the
    export get 0."""
    def empty(name, environment, in_inventory):
        return {"cluster": name, "environment": environment, "in_inventory": in_inventory, "in_export": "NO",
                "microservices": 0, "namespaces": set(), "hpas": 0, "levels": set(), "over_aca_size": 0}

    clusters = {key: empty(cluster["name"], cluster["environment"], "SÍ") for key, cluster in inventory.items()}
    for row in deployments:
        entry = clusters.setdefault(row["cluster"].lower(), empty(row["cluster"], row["environment"], "NO"))
        entry["in_export"] = "SÍ"
        if row["platform_component"] == "NO":
            entry["microservices"] += 1
            entry["namespaces"].add(row["namespace"])
            entry["levels"].add(row["bia"])
            entry["over_aca_size"] += row["fits_aca_by_request"] == "NO"
    for row in hpas:
        entry = clusters.get(row["cluster"].lower())
        if entry and row["platform_component"] == "NO":
            entry["hpas"] += 1
    out = []
    for key, entry in clusters.items():
        stateful = statefulsets.get(key, 0)
        fit = 1 <= entry["microservices"] <= ADT_MAX_MICROSERVICES and not entry["over_aca_size"] and not stateful
        out.append({"cluster": entry["cluster"], "environment": entry["environment"], "in_inventory": entry["in_inventory"],
                    "in_export": entry["in_export"], "microservices": entry["microservices"],
                    "namespaces": len(entry["namespaces"]), "hpas": entry["hpas"], "bia": most_critical(entry["levels"]),
                    "microservices_over_aca_size": entry["over_aca_size"], "statefulsets": stateful,
                    "aca_technical_fit": yes(fit)})
    return sorted(out, key=lambda value: (value["environment"], -value["microservices"], value["cluster"]))


def capability_rows(deployments):
    """PROD namespaces (domains) and deployments per capability keyword."""
    out = []
    for capability in CAPABILITIES:
        matches = [row for row in deployments if row["environment"] == "PROD" and row["platform_component"] == "NO"
                   and capability in row["capabilities"].split(";")]
        per_namespace = defaultdict(int)
        for row in matches:
            per_namespace[row["namespace"]] += 1
        out.append({"capability": capability, "namespaces_prod": len(per_namespace), "deployments_prod": len(matches),
                    "namespaces": ", ".join(f"{name}({count})" for name, count in sorted(per_namespace.items()))})
    return out


def flag_counts(rows, flags):
    prod = [row for row in rows if row["environment"] == "PROD"]
    return {"prod_rows": len(prod), **{flag: sum(row[flag] == "SÍ" for row in prod) for flag in flags}}


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--inputs", default=DEFAULT_INPUTS, help="Folder with the reportHpa and reportDeployment exports")
    parser.add_argument("--snapshot-date", default=None,
                        help="Day the exports were taken (YYYY-MM-DD). Default: modification date of the HPA file")
    parser.add_argument("--inventory", default=aa.results_path("aa-inventory"), help="Folder with the aa_inventory.py output")
    parser.add_argument("--container-insights", default=aa.results_path("aa-container-insights"),
                        help="Folder with the aa_container_insights.py output")
    parser.add_argument("--output-dir", default=aa.results_path("aa-k8s-snapshot"))
    parser.add_argument("--delimiter", default=",", help="CSV delimiter of the outputs")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)

    hpa_path, deployment_path = find_input(args.inputs, "reportHpa"), find_input(args.inputs, "reportDeployment")
    taken_on = snapshot_date(hpa_path, args.snapshot_date)
    inventory = {row["name"].lower(): row for row in aa.read_csv(os.path.join(args.inventory, "aks_clusters.csv"))}
    workloads = aa.read_csv(os.path.join(args.container_insights, "aks_workloads_monthly.csv"))
    if not workloads:
        azure.LOG.warning("No Container Insights results in %s: the history columns stay empty", args.container_insights)
    history = ci_history(workloads)
    clusters_with_ci = {row["cluster"].lower() for row in workloads}

    hpas = hpa_rows(select(read_rows(hpa_path), HPA_COLUMNS, hpa_path), history, inventory, taken_on)
    deployments = deployment_rows(select(read_rows(deployment_path), DEPLOYMENT_COLUMNS, deployment_path),
                                  history, inventory, hpas, clusters_with_ci)
    for rows in (hpas, deployments):
        for row in rows:
            row.pop("trend_max_export", None)
            row.pop("environment_export", None)

    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("hpa_snapshot.csv", hpas)
    output.csv("deployment_snapshot.csv", deployments)
    output.csv("cluster_snapshot.csv", cluster_rows(deployments, hpas, inventory, statefulsets_by_cluster(workloads)))
    output.csv("capability_domains.csv", capability_rows(deployments))
    output.json("snapshot_summary.json", {
        "generated_at": azure.iso(azure.utc_now()),
        "snapshot_date": taken_on.isoformat(),
        "inputs": [{"file": os.path.basename(path), "rows": count,
                    "modified": datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="minutes")}
                   for path, count in ((hpa_path, len(hpas)), (deployment_path, len(deployments)))],
        "units": {"cpu": "millicores in the export, vCPU in the outputs",
                  "memory": "MB (10^6 bytes) in the export, GiB in the outputs"},
        "container_insights_months": sorted({row["month"] for row in workloads}),
        "clusters_not_in_inventory": sorted({row["cluster"] for row in hpas + deployments if row["in_inventory"] == "NO"}),
        "hpa_without_ci_history": sum(row["ci_matched"] == "NO" for row in hpas),
        "hpa_prod": flag_counts(hpas, ["fixed_min_equals_max", "at_max_now", "trend_reached_max", "trend_never_above_min",
                                       "min_two_or_more"]),
        "hpa_prod_behaviour_ci": {code: sum(row["behaviour_ci"] == code for row in hpas if row["environment"] == "PROD")
                                  for code in ("NUNCA_ESCALA", "ESCALA", "LLEGA_AL_MAX", "FIJO", "SIN_DATO")},
        "deployment_prod": flag_counts(deployments, ["has_hpa", "no_usage_in_ci", "request_over_3x_used",
                                                     "used_over_request", "limit_10x_request", "fits_aca_by_request"]),
    })


if __name__ == "__main__":
    main()
