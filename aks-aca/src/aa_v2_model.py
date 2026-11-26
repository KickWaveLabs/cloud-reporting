"""Criteria v2 model: the decision and the Container Apps cost of each PROD application cluster.

One source for the v2 and v3 deck (aa_deck_exec.py); the workbook (aa_criteria_workbook.py) writes the same rules
as Excel formulas over the same inputs, so both give the same figures with the default variables.

Decision of a cluster, in order (the first that applies):
  NO_TRAFFIC        no traffic measured in the month on its own layer (grouped by no_traffic_reason), or
                    event-driven: its load is Event Hubs / Service Bus messages, not measured yet (group IN_USE)
  AKS               a "micro"-service requests more than a Container Apps replica allows (4 vCPU / 8 GiB): implicit
  AKS               busiest hour above MAX_TPS (average TPS of the busiest hour of the day, as v1)
  AKS               more than MAX_SERVICES "micro"-services in the cluster
  AKS               BIA at AKS_FROM_BIA or above (MUY ALTA > ALTA > MEDIA > BAJA; NONE by default: no cluster)
  AKS               shared: its traffic is on the Core layer, or its application is in SHARED_APPS (authentication
                    and identity, used by several applications)
  AKS               high, constant load: Container Apps would run, on average over the month, at least
                    MAX_REPLICAS_VS_AKS of the replicas it runs today on AKS (aca_replicas / replicas_today_total)
  ACA_WITH_REVIEW   critical (BIA ALTA or MUY ALTA) or growing past MAX_SERVICES with DESA
  ACA               otherwise
The saving margin is not a criterion: at the end, a cluster that meets the criteria moves to Container Apps only when
it saves at least MIN_SAVING of its AKS cost; otherwise it stays (group BELOW_MARGIN).

Container Apps cost of a cluster (US$ a month, list price):
  replicas(hour)  = max(minimum, ceil(TPS(hour) x TRAFFIC_SHARE / CONCURRENT_REQUESTS)) per "micro"-service, for each
                    hour of the day and day type; TPS is the hour's peak (busiest 5 minutes) or its average
                    (SIZING_BASIS); minimum = MIN_REPLICAS_OFFICE in office hours, MIN_REPLICAS_OFFHOURS otherwise.
                    BUSIEST_HOUR instead replays, 5 minutes at a time, that hour's busiest occurrence in the month
                    (the day it had the most requests): the hour's replicas are the average of its 12 buckets', and a
                    bucket keeps the replicas of the one before (5 minutes before Container Apps scales in).
                    A minimum of 0 lets Container Apps scale to zero, so the hour is billed only for its 5-minute
                    buckets that cannot be at zero: x (12 - off buckets) / 12, off buckets of the day with the
                    fewest (PEAK) or their average (AVERAGE) (aa_cluster_snapshot.py says how they count).
                    Container Apps adds a replica each CONCURRENT_REQUESTS requests per second (HTTP scale rule,
                    default 10: requests of the last 15 s / 15).
  compute         = sum over hours of days x 3600 x replicas(hour) x sum over "micro"-services of
                    size x (price per vCPU-s + GIB_PER_VCPU x price per GiB-s); size as v1 (aca_size)
  requests        = requests of the month x REQUESTS_BUFFER / 1e6 x price per million (0 when event-driven)
  environment     = ENVIRONMENTS_PER_CLUSTER x HOURS_MONTH x Standard load balancer hour
  (+ ENVIRONMENT_FEE: ENVIRONMENTS_PER_CLUSTER x HOURS_MONTH x Dedicated plan management hour, charged when the
   environment has a private endpoint or planned maintenance)
Criteria v3 is this model with the PRESETS["v3"] values (use("v3")): each hour's busiest occurrence replayed every 5
minutes; replicas sized by what one replica handles today on AKS (CAPACITY_FROM_TODAY: TPS per replica at the peak
hour x CAPACITY_SAFETY, never below CONCURRENT_REQUESTS); each "micro"-service's size from its measured use
(SIZE_FROM_USAGE: CPU and memory used per replica x USAGE_HEADROOM, never above its request). v2 keeps them all off;
ENVIRONMENT_FEE is off in both (a switch for when a private endpoint is confirmed).
With the bank's discount and IGV in soles: x (1 - BANK_DISCOUNT) x (1 + IGV) x EXCHANGE_RATE.
"""

import math
import os
from collections import defaultdict

import aa_common as aa
import aa_deck_exec_v1 as v1

DEFAULTS = {
    "max_tps": 100, "max_services": 15, "min_saving": 0.10,
    "concurrent_requests": 10, "traffic_share": 1.0, "sizing_basis": "BUSIEST_HOUR",
    "min_replicas_office": 1, "min_replicas_offhours": 0,
    "aks_from_bia": "NONE", "max_replicas_vs_aks": 0.8,
    "shared_apps": "AUTH,ASOB,B2CC,CB2C,CIAM,TOKE,TOTP,TPWL,TRSA,BIOM",
    "capacity_from_today": "NO", "capacity_safety": 0.7, "size_from_usage": "NO", "usage_headroom": 2.0,
    "environment_fee": "NO",
    "requests_buffer": v1.REQUESTS_BUFFER, "environments_per_cluster": 1, "hours_month": v1.HOURS_MONTH,
    "min_vcpu": v1.ACA_MIN_VCPU, "vcpu_step": 0.25, "gib_per_vcpu": 2,
    "bank_discount": v1.BANK_DISCOUNT, "igv": v1.IGV, "exchange_rate": v1.EXCHANGE_RATE,
}
BASE_DEFAULTS = dict(DEFAULTS)
CRITICAL_BIA = ("MUY ALTA", "ALTA")
BIA_LEVELS = ["BAJA", "MEDIA", "ALTA", "MUY ALTA"]  # lowest first; SIN BIA is below them all
PRICE_KEYS = {"price_vcpu": "aca_vcpu_active_s", "price_gib": "aca_mem_active_gib_s",
              "price_requests": "aca_requests_million", "price_lb_hour": "lb_standard_rules_hour",
              "price_env_hour": "aca_env_management_hour"}
PRESETS = {"v2": {}, "v3": {"sizing_basis": "BUSIEST_HOUR", "capacity_from_today": "YES", "size_from_usage": "YES"}}


def use(criteria):
    """Make DEFAULTS those of criteria v2 or v3 (the deck and the workbook read DEFAULTS)."""
    DEFAULTS.update(BASE_DEFAULTS)
    DEFAULTS.update(PRESETS[criteria])
    return criteria
GROUP_OF_REASON = {"IN_DEVELOPMENT": "IN_DEVELOPMENT", "IN_USE_NOT_MEASURED": "IN_USE"}


def to_float(value):
    return v1.to_float(value)


def size_vcpu(cpu, mem_gib, variables, cpu_used=None, mem_used=None):
    """vCPU of the smallest Container Apps replica that covers the request (as v1, with the variables); with
    size_from_usage, the request comes down to the measured use x usage_headroom when that is lower."""
    if variables["size_from_usage"] == "YES":
        if cpu_used:
            cpu = min(cpu or 0, cpu_used * variables["usage_headroom"])
        if mem_used:
            mem_gib = min(mem_gib or 0, mem_used * variables["usage_headroom"])
    step = variables["vcpu_step"]
    need = max(cpu or 0, (mem_gib or 0) / variables["gib_per_vcpu"], variables["min_vcpu"])
    return math.ceil(need / step - 1e-9) * step


def load(results):
    """Inputs of the model from results/: clusters, hourly profile, services and prices."""
    read = lambda folder, name: aa.read_csv(os.path.join(results, folder, name))
    clusters = read("aa-cluster-snapshot", "cluster_criteria_input.csv")
    if not clusters:
        raise SystemExit("No aa-cluster-snapshot/cluster_criteria_input.csv: run aa_cluster_snapshot.py")
    names = {row["cluster_id"]: row["cluster"] for row in clusters}
    services = [row for row in read("aa-k8s-snapshot", "deployment_snapshot.csv")
                if row["microservice"] == "YES" and row["cluster_id"] in names]
    prices = v1.read_prices(results)
    return {"clusters": clusters, "profile": read("aa-cluster-snapshot", "traffic_hourly_profile.csv"),
            "services": services, "names": names,
            "prices": {name: prices.get((key, "eastus2"), 0) for name, key in PRICE_KEYS.items()}}


BUCKETS = 12  # 5-minute buckets in an hour
BUSIEST_COLUMNS = [f"busiest_hour_5min_{minute:02d}" for minute in range(0, 60, 5)]
SIZING_BASES = ("PEAK", "BUSIEST_HOUR", "AVERAGE")


def capacity(row, variables):
    """TPS per replica before Container Apps adds one: concurrent_requests, or with capacity_from_today what one
    replica handles today on AKS x capacity_safety, never below concurrent_requests."""
    today = to_float(row.get("tps_per_replica_today"))
    if variables["capacity_from_today"] == "YES" and today:
        return max(variables["concurrent_requests"], today * variables["capacity_safety"])
    return variables["concurrent_requests"]


def need(tps, variables, per_replica):
    return math.ceil(tps * variables["traffic_share"] / per_replica - 1e-9)


def replicas(row, variables):
    """Simulated replicas per "micro"-service at one hour of the profile (an average over the hour when the
    5-minute buckets at zero count)."""
    minimum = variables["min_replicas_office"] if row["office"] in ("1", 1) else variables["min_replicas_offhours"]
    if variables["sizing_basis"] == "BUSIEST_HOUR":  # zeros are in the buckets themselves
        tps = [to_float(row.get(name)) or 0 for name in BUSIEST_COLUMNS]
        return sum(max(minimum, need(max(now, before), variables, capacity(row, variables)))
                   for now, before in zip(tps, tps[:1] + tps[:-1])) / BUCKETS
    peak = variables["sizing_basis"] == "PEAK"
    tps = to_float(row["tps_peak"] if peak else row["tps_avg"]) or 0
    count = max(minimum, need(tps, variables, capacity(row, variables)))
    if minimum == 0:  # scale to zero allowed: the buckets without traffic are not billed
        off = to_float(row.get("off_5min_min" if peak else "off_5min_avg")) or 0
        return count * (BUCKETS - off) / BUCKETS
    return count


def decide(row, variables):
    """(decision, reason) of one cluster row of cluster_criteria_input.csv."""
    if row["event_driven"] == "YES":
        return "NO_TRAFFIC", "Event-driven: its load is messages, not measured yet"
    if row["traffic_status"] != "MEASURED" or row["busiest_hour_tps"] == "":
        return "NO_TRAFFIC", f"No traffic of its own in the month ({row['traffic_status']})"
    if int(to_float(row["microservices_over_aca_size"]) or 0):
        return "AKS", "A \"micro\"-service requests more than 4 vCPU / 8 GiB"
    if (to_float(row["busiest_hour_tps"]) or 0) > variables["max_tps"]:
        return "AKS", f"Busiest hour above {variables['max_tps']} TPS"
    services, app = int(to_float(row["microservices"]) or 0), int(to_float(row["microservices_app"]) or 0)
    if services > variables["max_services"]:
        return "AKS", f"More than {variables['max_services']} \"micro\"-services"
    threshold = variables["aks_from_bia"]
    if threshold in BIA_LEVELS and row["bia"] in BIA_LEVELS and \
            BIA_LEVELS.index(row["bia"]) >= BIA_LEVELS.index(threshold):
        return "AKS", f"BIA {threshold} or above"
    shared = {code.strip().upper() for code in variables["shared_apps"].split(",") if code.strip()}
    if row["traffic_layer"] == "CORE" or row["cod_app"].upper() in shared:
        return "AKS", "Shared: authentication or core, used by several applications"
    today, aca = to_float(row.get("replicas_today_total")), to_float(row.get("aca_replicas"))
    if today and aca is not None and aca / today >= variables["max_replicas_vs_aks"]:
        return "AKS", f"High, constant load: Container Apps needs {aca / today:.0%} of today's AKS replicas"
    reasons = []
    if row["bia"] in CRITICAL_BIA:
        reasons.append("Critical (BIA)")
    projected = to_float(row["microservices_app_projected"])
    if projected is not None and projected > variables["max_services"] >= app:
        reasons.append("Grows with DESA")
    return ("ACA_WITH_REVIEW", "; ".join(reasons)) if reasons else ("ACA", "Meets every criterion")


def evaluate(data, variables=None):
    """One dict per cluster: inputs, decision, cost on AKS and Container Apps (soles with IGV) and the group."""
    variables = dict(DEFAULTS, **(variables or {}))
    prices = data["prices"]
    unit = defaultdict(float)  # cluster -> US$ per second of one replica of every "micro"-service
    for row in data["services"]:
        size = size_vcpu(to_float(row["cpu_request_vcpu"]), to_float(row["mem_request_gib"]), variables,
                         to_float(row.get("ci_cpu_used_per_replica")), to_float(row.get("ci_mem_used_gib_per_replica")))
        unit[data["names"][row["cluster_id"]]] += size * (prices["price_vcpu"] + variables["gib_per_vcpu"] *
                                                          prices["price_gib"])
    replica_seconds, seconds = defaultdict(float), defaultdict(float)
    for row in data["profile"]:
        if row["scope_type"] == "CLUSTER":
            replica_seconds[row["scope"]] += int(row["days"]) * 3600 * replicas(row, variables)
            seconds[row["scope"]] += int(row["days"]) * 3600
    soles = (1 + variables["igv"]) * variables["exchange_rate"]
    out = []
    for row in data["clusters"]:
        # replicas of all its "micro"-services on Container Apps, on average over the month
        aca_replicas = replica_seconds[row["cluster"]] / seconds[row["cluster"]] * \
            (to_float(row["microservices"]) or 0) if seconds[row["cluster"]] else None
        row = dict(row, aca_replicas=aca_replicas)
        decision, reason = decide(row, variables)
        requests = 0 if row["event_driven"] == "YES" else (to_float(row["requests_month"]) or 0)
        compute = replica_seconds[row["cluster"]] * unit[row["cluster"]]
        aca_list = (compute + requests * variables["requests_buffer"] / 1e6 * prices["price_requests"]
                    + variables["environments_per_cluster"] * variables["hours_month"] * (
                        prices["price_lb_hour"] + (prices["price_env_hour"] if variables["environment_fee"] == "YES"
                                                   else 0)))
        aca = aca_list * (1 - variables["bank_discount"]) * soles
        exact = to_float(row.get("exact_replica_seconds"))  # reference, with the default variables
        aca_exact = None if exact is None else \
            (aca_list - compute + exact * unit[row["cluster"]]) * (1 - variables["bank_discount"]) * soles
        aks = (to_float(row["cost_month_usd"]) or 0) * soles
        saving = aks - aca
        share = saving / aks if aks else 0
        movable = decision in ("ACA", "ACA_WITH_REVIEW")
        if decision == "NO_TRAFFIC":
            group = "IN_USE" if row["event_driven"] == "YES" else \
                GROUP_OF_REASON.get(row["no_traffic_reason"], "NOT_IN_USE")
        elif movable and share < variables["min_saving"]:
            group = "BELOW_MARGIN"
        else:
            group = decision
        out.append(dict(row, decision=decision, reason=reason, replica_seconds=replica_seconds[row["cluster"]],
                        unit_usd_per_second=unit[row["cluster"]], aca_list_usd=aca_list, aca_month_pen=aca,
                        aks_month_pen=aks, saving_month_pen=saving, aca_exact_month_pen=aca_exact, saving_share=share, group=group,
                        goes_to_aca=group in ("ACA", "ACA_WITH_REVIEW")))
    return out
