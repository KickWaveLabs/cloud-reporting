#!/usr/bin/env python3
"""Criteria v2 inputs: one row per PROD application cluster and the hourly traffic of each cluster and application.

Reads only results/ (no Azure call): the snapshot (aa_k8s_snapshot.py), the inventory's APIM APIs, the metrics'
hourly series (apim_api_hourly.csv, entry_points_hourly.csv), Container Insights (replicas by hour of day) and cost.

Each API goes to a cluster by its backend:
  azingress<code><region><nn>.lima.bcp.com.pe   -> aksv<region><code>p<nn> (azingressbadieu211 -> aksveu2badip11)
  <service>.<namespace>.svc.cluster.local         -> the only PROD application cluster that runs that namespace
  anything else (or an entry point)               -> the application's cluster when it has only one
  azingress<code>.lima.bcp.com.pe (no number)     -> aksveu2<code>p01, the application's first cluster
  an entry point named ...<region><code>[int]p<nn> -> aksv<region><code>p<nn> (apgweu2badip11 -> aksveu2badip11)
An API that matches none stays at application level only (attribution NOT_ATTRIBUTED in api_cluster_map.csv).

A cluster takes one traffic layer, as an application does in v1: BS when it has BS traffic in the month, else the
busiest of the other layers. Layers are never added (a call goes on from one gateway to the next).

Hourly profile per scope (cluster or application) and day type (weekday Monday to Friday, weekend) and Lima hour:
  tps_min / tps_avg / tps_max_hour  lowest, average and highest requests of that hour over the month's days / 3600
                                    (a day without traffic in that hour counts as 0)
  tps_peak                          busiest 5 minutes of that hour in the month / 300, adding each API's busiest
                                    5 minutes of the hour (an upper bound: APIs do not peak at the same minute)
  replicas_today_avg / _max         Container Insights replicas per "micro"-service of the scope at that hour
                                    (average over its "micro"-services), when aks_workloads_hourly_profile.csv exists

Outputs (in --output-dir, results/aa-cluster-snapshot by default):
  cluster_criteria_input.csv   PROD application cluster: application, BIA, "micro"-services (cluster and
                               application), traffic layer and figures, cost, the inputs of every v2 criterion
  traffic_hourly_profile.csv   scope x day type x hour, as above
  api_cluster_map.csv          PROD AKS API -> cluster and how it was attributed
  cluster_snapshot_summary.json
"""

import argparse
import calendar
import json
import os
import re
from collections import defaultdict
from datetime import date, datetime

import aa_common as aa
import aa_k8s_snapshot as snap

azure = aa.azure
INGRESS_HOST = re.compile(r"^azingress(?P<code>[a-z0-9]+?)(?P<region>eu2|cu1|sc1)(?P<number>\d{2})\.")
INGRESS_PLAIN = re.compile(r"^azingress(?P<code>[a-z0-9]{4})\.")
# Entry points carry the cluster in their name: apgweu2badip11 -> aksveu2badip11, AGTWEU2BADIINTP01 -> aksveu2badip01.
ENTRY_NAME = re.compile(r"(?P<region>eu2|cu1|sc1)(?P<code>[a-z0-9]{4})(?:int)?p(?P<number>\d{2})")
DAY_TYPES = ("weekday", "weekend")


def day_counts(month):
    """Days of each day type in a YYYY-MM month."""
    year, number = int(month[:4]), int(month[5:7])
    days = [date(year, number, day) for day in range(1, calendar.monthrange(year, number)[1] + 1)]
    return {"weekday": sum(day.weekday() < 5 for day in days), "weekend": sum(day.weekday() >= 5 for day in days)}


def is_office(day_type, hour):
    return day_type == "weekday" and aa.OFFICE_START_HOUR <= hour < aa.OFFICE_END_HOUR


def attribute(row, clusters_by_name, clusters_by_code, namespace_clusters):
    """(cluster id, how) for a PROD AKS API, or ("", "NOT_ATTRIBUTED")."""
    host = str(row.get("backend_host") or "").lower()
    code = (row.get("cod_app") or "").upper()
    match = INGRESS_HOST.match(host)
    if match:
        name = f"aksv{match['region']}{match['code']}p{match['number']}"
        if name in clusters_by_name:
            return clusters_by_name[name], "INGRESS_HOST"
    if host.endswith(".svc.cluster.local"):
        running = namespace_clusters.get(host.split(".")[1], set())
        if len(running) == 1:
            return next(iter(running)), "NAMESPACE"
    plain = INGRESS_PLAIN.match(host)
    if plain:
        # The ingress host without a number belongs to the application's first cluster (p01): where an application
        # has several, the numbered hosts point to the others and p01 has no numbered host of its own.
        name = f"aksveu2{plain['code']}p01"
        if name in clusters_by_name:
            return clusters_by_name[name], "INGRESS_HOST_P01"
        code = code or plain["code"].upper()
    own = clusters_by_code.get(code, [])
    if len(own) == 1:
        return own[0], "ONLY_CLUSTER"
    return "", "NOT_ATTRIBUTED"


def choose_layer(requests_by_layer):
    """BS when it has traffic, else the busiest fallback layer; '' without traffic."""
    if requests_by_layer.get(aa.PRIMARY_LAYER):
        return aa.PRIMARY_LAYER
    busiest = max(aa.FALLBACK_LAYERS, key=lambda layer: requests_by_layer.get(layer, 0))
    return busiest if requests_by_layer.get(busiest) else ""


def profile_rows(scope, scope_type, series, days, replicas):
    """24 x 2 rows for one scope from its hourly series {(date, hour): (requests, peak_5min)}."""
    by_slot = defaultdict(list)
    for (day, hour), (requests, peak) in series.items():
        by_slot[("weekday" if day.weekday() < 5 else "weekend", hour)].append((requests, peak))
    out = []
    for day_type in DAY_TYPES:
        for hour in range(24):
            values = by_slot.get((day_type, hour), [])
            count = days[day_type]
            hourly = [requests for requests, _ in values] + [0] * max(count - len(values), 0)
            today = replicas.get((day_type, hour), ("", ""))
            out.append({"scope": scope, "scope_type": scope_type, "day_type": day_type, "hour": hour,
                        "office": 1 if is_office(day_type, hour) else 0, "days": count,
                        "tps_min": round(min(hourly) / 3600, 3) if hourly else 0,
                        "tps_avg": round(sum(hourly) / count / 3600, 3) if count else 0,
                        "tps_max_hour": round(max(hourly) / 3600, 3) if hourly else 0,
                        "tps_peak": round(max((peak for _, peak in values), default=0) / 300, 3),
                        "replicas_today_avg": today[0], "replicas_today_max": today[1]})
    return out


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output-dir", default=aa.results_path("aa-cluster-snapshot"))
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)
    read = lambda folder, name: aa.read_csv(os.path.join(args.results, folder, name))

    clusters = [row for row in read("aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["category"] == aa.APPLICATION]
    if not clusters:
        raise SystemExit("No PROD application cluster in aa-k8s-snapshot/cluster_snapshot.csv: run aa_k8s_snapshot.py")
    applications = {row["cod_app"]: row for row in read("aa-k8s-snapshot", "application_snapshot.csv")}
    deployments = [row for row in read("aa-k8s-snapshot", "deployment_snapshot.csv")
                   if row["environment"] == "PROD" and row["microservice"] == "YES"]
    by_id = {row["cluster_id"]: row for row in clusters}
    clusters_by_name = {row["cluster"].lower(): row["cluster_id"] for row in clusters}
    clusters_by_code = defaultdict(list)
    for row in clusters:
        clusters_by_code[row["cod_app"].upper()].append(row["cluster_id"])
    inventory = read("aa-inventory", "aks_clusters.csv")
    ids_by_name = aa.cluster_ids_by_name(inventory)
    namespace_clusters = defaultdict(set)
    for row in read("aa-container-insights", "aks_workloads_monthly.csv"):
        cluster_id = aa.cluster_id_of(row, ids_by_name)
        if cluster_id in by_id and aa.is_microservice(row["namespace"], row["workload"]):
            namespace_clusters[row["namespace"].lower()].add(cluster_id)

    # API -> (cluster, application, layer)
    api_map, map_rows = {}, []
    for row in read("aa-inventory", "apim_apis.csv"):
        if row.get("apim_environment", "PROD") != "PROD" or (
                row.get("backend_platform") or aa.backend_platform(row.get("backend_host"))) != "AKS":
            continue
        code = snap.api_code(row, {})
        cluster_id, how = attribute(dict(row, cod_app=code), clusters_by_name, clusters_by_code, namespace_clusters)
        if not code and cluster_id:
            code = by_id[cluster_id]["cod_app"]
        api_map[(row["apim"], row["api_id"])] = (cluster_id, code, row["apim_layer"])
        map_rows.append({"apim": row["apim"], "apim_layer": row["apim_layer"], "api_id": row["api_id"],
                         "backend_host": row.get("backend_host", ""), "cod_app": code,
                         "cluster": by_id[cluster_id]["cluster"] if cluster_id else "", "attribution": how})
    monthly = read("aa-metrics", "apim_api_monthly.csv") + read("aa-metrics", "entry_points_monthly.csv")
    for row in monthly:
        if row.get("apim_layer") in aa.GATEWAY_LAYERS + aa.QUEUE_LAYERS and (row["apim"], row["api_id"]) not in api_map:
            code = (row.get("cod_app") or "").upper()
            own = clusters_by_code.get(code, [])
            named = ENTRY_NAME.search(row["apim"].lower())
            name = f"aksv{named['region']}{named['code']}p{named['number']}" if named else ""
            cluster_id = clusters_by_name.get(name) or (own[0] if len(own) == 1 else "")
            api_map[(row["apim"], row["api_id"])] = (cluster_id, code, row["apim_layer"])
            map_rows.append({"apim": row["apim"], "apim_layer": row["apim_layer"], "api_id": row["api_id"],
                             "backend_host": "", "cod_app": code,
                             "cluster": by_id[cluster_id]["cluster"] if cluster_id else "",
                             "attribution": ("ENTRY_NAME" if name in clusters_by_name else "ONLY_CLUSTER")
                             if cluster_id else "NOT_ATTRIBUTED"})
    month = max((row["month"] for row in monthly), default="")
    days = day_counts(month)

    # Requests of the month per (cluster, layer) and (application, layer); one layer each.
    cluster_layers, app_layers = defaultdict(lambda: defaultdict(float)), defaultdict(lambda: defaultdict(float))
    for row in monthly:
        cluster_id, code, layer = api_map.get((row["apim"], row["api_id"]), ("", "", ""))
        if row["month"] != month or layer not in aa.APPLICATION_TRAFFIC_ORDER:
            continue
        requests = snap.to_float(row.get("requests")) or 0
        if cluster_id:
            cluster_layers[cluster_id][layer] += requests
        if code:
            app_layers[code][layer] += requests
    cluster_layer = {key: choose_layer(value) for key, value in cluster_layers.items()}
    app_layer = {code: (row.get("traffic_layer") or "") for code, row in applications.items()}

    # Hourly series per scope.
    series = defaultdict(dict)
    for row in read("aa-metrics", "apim_api_hourly.csv") + read("aa-metrics", "entry_points_hourly.csv"):
        if row["month"] != month:
            continue
        cluster_id, code, layer = api_map.get((row["apim"], row["api_id"]), ("", "", ""))
        moment = datetime.strptime(row["hour_start"], "%Y-%m-%dT%H:%M")
        key = (moment.date(), moment.hour)
        values = (snap.to_float(row["requests"]) or 0, snap.to_float(row["peak_5min_requests"]) or 0)
        for scope, chosen in ((cluster_id, cluster_layer.get(cluster_id)), (code, app_layer.get(code))):
            if scope and layer and layer == chosen:
                previous = series[scope].get(key, (0, 0))
                series[scope][key] = (previous[0] + values[0], previous[1] + values[1])

    # Replicas per "micro"-service by hour of day (Container Insights), averaged over the scope's "micro"-services.
    service_scope = {}
    for row in deployments:
        if row["cluster_id"] in by_id:
            service_scope[(row["cluster_id"], row["namespace"], row["deployment"])] = (
                row["cluster_id"], by_id[row["cluster_id"]]["cod_app"])
    hourly_path = os.path.join(args.results, "aa-container-insights", "aks_workloads_hourly_profile.csv")
    replicas = defaultdict(lambda: defaultdict(list))
    for row in aa.read_csv(hourly_path) if os.path.exists(hourly_path) else []:
        if row.get("month") != month:
            continue
        scopes = service_scope.get((aa.cluster_id_of(row, ids_by_name), row["namespace"], row["workload"]))
        for scope in scopes or ():
            replicas[scope][(row["day_type"], int(row["hour"]))].append(
                (snap.to_float(row["replicas_avg"]) or 0, snap.to_float(row["replicas_max"]) or 0))
    replicas_today = {scope: {slot: (round(sum(a for a, _ in values) / len(values), 2),
                                     round(sum(m for _, m in values) / len(values), 2))
                              for slot, values in slots.items()} for scope, slots in replicas.items()}

    profile = []
    for cluster_id in sorted(by_id, key=lambda key: by_id[key]["cluster"]):
        profile += profile_rows(by_id[cluster_id]["cluster"], "CLUSTER", series.get(cluster_id, {}), days,
                                replicas_today.get(cluster_id, {}))
    for code in sorted({row["cod_app"] for row in clusters}):
        profile += profile_rows(code, "APP", series.get(code, {}), days, replicas_today.get(code, {}))
    busiest = defaultdict(float)
    peak = defaultdict(float)
    for row in profile:
        busiest[row["scope"]] = max(busiest[row["scope"]], row["tps_avg"])
        peak[row["scope"]] = max(peak[row["scope"]], row["tps_peak"])

    cost_month, cost = snap.app_cost(read("aa-cost", "cost_by_resource.csv"), set(by_id))
    sizes = defaultdict(int)
    for row in deployments:
        if row["cluster_id"] in by_id and row["fits_aca_by_request"] == "NO":
            sizes[row["cluster_id"]] += 1
    out = []
    for cluster_id, row in sorted(by_id.items(), key=lambda item: item[1]["cluster"]):
        app = applications.get(row["cod_app"], {})
        layer = cluster_layer.get(cluster_id, "")
        requests = cluster_layers.get(cluster_id, {}).get(layer, 0)
        microservices_app = int(snap.to_float(app.get("microservices")) or 0)
        projected = snap.to_float(app.get("microservices_projected"))
        if layer:
            status, reason = "MEASURED", ""
        elif app.get("traffic_status") == "MEASURED":
            status, reason = "NOT_ATTRIBUTED", "IN_USE_NOT_MEASURED"
        else:
            status, reason = app.get("traffic_status", "NO_APIM_API"), app.get("no_traffic_reason", "")
        out.append({
            "cluster": row["cluster"], "cluster_id": cluster_id, "cod_app": row["cod_app"],
            "clusters_of_app": len(clusters_by_code.get(row["cod_app"].upper(), [])), "bia": row["bia"],
            "microservices": int(snap.to_float(row["microservices"]) or 0), "microservices_app": microservices_app,
            "microservices_app_projected": int(projected) if projected is not None else "",
            "microservices_over_aca_size": sizes.get(cluster_id, 0),
            "traffic_status": status, "traffic_layer": layer, "requests_month": int(requests),
            "busiest_hour_tps": round(busiest.get(row["cluster"], 0), 3) if layer else "",
            "peak_tps": round(peak.get(row["cluster"], 0), 3) if layer else "",
            "event_driven": "YES" if aa.event_driven(layer) else "NO",
            "cost_month_usd": round(cost.get(cluster_id, 0), 2), "month": cost_month,
            "no_traffic_reason": reason})
    output = azure.Output(args.output_dir, ",")
    output.csv("cluster_criteria_input.csv", out)
    output.csv("traffic_hourly_profile.csv", profile)
    output.csv("api_cluster_map.csv", map_rows)
    attribution = defaultdict(int)
    for row in map_rows:
        attribution[row["attribution"]] += 1
    output.json("cluster_snapshot_summary.json", {
        "generated_at": azure.iso(azure.utc_now()), "traffic_month": month, "cost_month": cost_month,
        "day_counts": days, "clusters": len(out),
        "clusters_by_traffic_status": {status: sum(row["traffic_status"] == status for row in out)
                                       for status in sorted({row["traffic_status"] for row in out})},
        "api_attribution": dict(attribution),
        "replicas_by_hour_available": os.path.exists(hourly_path),
    })


if __name__ == "__main__":
    main()
