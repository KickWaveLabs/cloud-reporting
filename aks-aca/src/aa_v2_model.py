"""Criteria v2 model: the decision and the Container Apps cost of each PROD application cluster.

One source for the v2 deck (aa_deck_exec_v2.py); the v2 workbook (aa_criteria_workbook_v2.py) writes the same rules
as Excel formulas over the same inputs, so both give the same figures with the default variables.

Decision of a cluster, in order (the first that applies):
  NO_TRAFFIC        no traffic measured in the month on its own layer (grouped by no_traffic_reason)
  AKS               event-driven (Event Hubs / Service Bus messages, not TPS)
  AKS               a "micro"-service requests more than a Container Apps replica allows (4 vCPU / 8 GiB): implicit
  AKS               busiest hour above MAX_TPS (average TPS of the busiest hour of the day, as v1)
  AKS               more than MAX_SERVICES "micro"-services in the cluster
  ACA_WITH_REVIEW   critical (BIA ALTA or MUY ALTA), growing past MAX_SERVICES with DESA, or its application has more
                    than MAX_SERVICES "micro"-services across its clusters
  ACA               otherwise
The saving margin is not a criterion: at the end, a cluster that meets the criteria moves to Container Apps only when
it saves at least MIN_SAVING of its AKS cost; otherwise it stays (group BELOW_MARGIN).

Container Apps cost of a cluster (US$ a month, list price):
  replicas(hour)  = max(minimum, ceil(TPS(hour) x TRAFFIC_SHARE / CONCURRENT_REQUESTS)) per "micro"-service, for each
                    hour of the day and day type; TPS is the hour's peak (busiest 5 minutes) or its average
                    (SIZING_BASIS); minimum = MIN_REPLICAS_OFFICE in office hours, MIN_REPLICAS_OFFHOURS otherwise.
                    Container Apps adds a replica each CONCURRENT_REQUESTS requests per second (HTTP scale rule,
                    default 10: requests of the last 15 s / 15).
  compute         = sum over hours of days x 3600 x replicas(hour) x sum over "micro"-services of
                    size x (price per vCPU-s + GIB_PER_VCPU x price per GiB-s); size as v1 (aca_size)
  requests        = requests of the month x REQUESTS_BUFFER / 1e6 x price per million (0 when event-driven)
  environment     = ENVIRONMENTS_PER_CLUSTER x HOURS_MONTH x Standard load balancer hour
With the bank's discount and IGV in soles: x (1 - BANK_DISCOUNT) x (1 + IGV) x EXCHANGE_RATE.
"""

import math
import os
from collections import defaultdict

import aa_common as aa
import aa_deck_exec as v1

DEFAULTS = {
    "max_tps": 100, "max_services": 15, "min_saving": 0.10,
    "concurrent_requests": 10, "traffic_share": 1.0, "sizing_basis": "PEAK",
    "min_replicas_office": 1, "min_replicas_offhours": 1,
    "requests_buffer": v1.REQUESTS_BUFFER, "environments_per_cluster": 1, "hours_month": v1.HOURS_MONTH,
    "min_vcpu": v1.ACA_MIN_VCPU, "vcpu_step": 0.25, "gib_per_vcpu": 2,
    "bank_discount": v1.BANK_DISCOUNT, "igv": v1.IGV, "exchange_rate": v1.EXCHANGE_RATE,
}
CRITICAL_BIA = ("MUY ALTA", "ALTA")
PRICE_KEYS = {"price_vcpu": "aca_vcpu_active_s", "price_gib": "aca_mem_active_gib_s",
              "price_requests": "aca_requests_million", "price_lb_hour": "lb_standard_rules_hour"}
GROUP_OF_REASON = {"IN_DEVELOPMENT": "IN_DEVELOPMENT", "IN_USE_NOT_MEASURED": "IN_USE"}


def to_float(value):
    return v1.to_float(value)


def size_vcpu(cpu, mem_gib, variables):
    """vCPU of the smallest Container Apps replica that covers the request (as v1, with the variables)."""
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


def replicas(row, variables):
    """Simulated replicas per "micro"-service at one hour of the profile."""
    tps = to_float(row["tps_peak"] if variables["sizing_basis"] == "PEAK" else row["tps_avg"]) or 0
    minimum = variables["min_replicas_office"] if row["office"] in ("1", 1) else variables["min_replicas_offhours"]
    return max(minimum, math.ceil(tps * variables["traffic_share"] / variables["concurrent_requests"] - 1e-9))


def decide(row, variables):
    """(decision, reason) of one cluster row of cluster_criteria_input.csv."""
    if row["traffic_status"] != "MEASURED" or row["busiest_hour_tps"] == "":
        return "NO_TRAFFIC", f"No traffic of its own in the month ({row['traffic_status']})"
    if row["event_driven"] == "YES":
        return "AKS", "Event-driven: messages, not TPS"
    if int(to_float(row["microservices_over_aca_size"]) or 0):
        return "AKS", "A \"micro\"-service requests more than 4 vCPU / 8 GiB"
    if (to_float(row["busiest_hour_tps"]) or 0) > variables["max_tps"]:
        return "AKS", f"Busiest hour above {variables['max_tps']} TPS"
    services, app = int(to_float(row["microservices"]) or 0), int(to_float(row["microservices_app"]) or 0)
    if services > variables["max_services"]:
        return "AKS", f"More than {variables['max_services']} \"micro\"-services"
    reasons = []
    if row["bia"] in CRITICAL_BIA:
        reasons.append("Critical (BIA)")
    projected = to_float(row["microservices_app_projected"])
    if projected is not None and projected > variables["max_services"] >= app:
        reasons.append("Grows with DESA")
    if app > variables["max_services"]:
        reasons.append(f"Its application has {app} \"micro\"-services")
    return ("ACA_WITH_REVIEW", "; ".join(reasons)) if reasons else ("ACA", "Meets every criterion")


def evaluate(data, variables=None):
    """One dict per cluster: inputs, decision, cost on AKS and Container Apps (soles with IGV) and the group."""
    variables = dict(DEFAULTS, **(variables or {}))
    prices = data["prices"]
    unit = defaultdict(float)  # cluster -> US$ per second of one replica of every "micro"-service
    for row in data["services"]:
        size = size_vcpu(to_float(row["cpu_request_vcpu"]), to_float(row["mem_request_gib"]), variables)
        unit[data["names"][row["cluster_id"]]] += size * (prices["price_vcpu"] + variables["gib_per_vcpu"] *
                                                          prices["price_gib"])
    replica_seconds = defaultdict(float)
    for row in data["profile"]:
        if row["scope_type"] == "CLUSTER":
            replica_seconds[row["scope"]] += int(row["days"]) * 3600 * replicas(row, variables)
    soles = (1 + variables["igv"]) * variables["exchange_rate"]
    out = []
    for row in data["clusters"]:
        decision, reason = decide(row, variables)
        requests = 0 if row["event_driven"] == "YES" else (to_float(row["requests_month"]) or 0)
        compute = replica_seconds[row["cluster"]] * unit[row["cluster"]]
        aca_list = (compute + requests * variables["requests_buffer"] / 1e6 * prices["price_requests"]
                    + variables["environments_per_cluster"] * variables["hours_month"] * prices["price_lb_hour"])
        aca = aca_list * (1 - variables["bank_discount"]) * soles
        aks = (to_float(row["cost_month_usd"]) or 0) * soles
        saving = aks - aca
        share = saving / aks if aks else 0
        movable = decision in ("ACA", "ACA_WITH_REVIEW")
        if decision == "NO_TRAFFIC":
            group = GROUP_OF_REASON.get(row["no_traffic_reason"], "NOT_IN_USE")
        elif movable and share < variables["min_saving"]:
            group = "BELOW_MARGIN"
        else:
            group = decision
        out.append(dict(row, decision=decision, reason=reason, replica_seconds=replica_seconds[row["cluster"]],
                        unit_usd_per_second=unit[row["cluster"]], aca_list_usd=aca_list, aca_month_pen=aca,
                        aks_month_pen=aks, saving_month_pen=saving, saving_share=share, group=group,
                        goes_to_aca=group in ("ACA", "ACA_WITH_REVIEW")))
    return out
