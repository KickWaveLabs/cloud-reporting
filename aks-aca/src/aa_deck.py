#!/usr/bin/env python3
"""Build the full AKS vs Container Apps deck from the saved workbook and the signals CSVs.

Reads the values Excel calculated in aks_aca_comparison.xlsx, so open the workbook in Excel and save it once
before running this script. The capacity health slide also reads results/aa-signals (aa_signals.py); without
it the slide says what to run. No Azure call is made; nothing leaves the machine.

Slides:
   1. Cover and agenda
  Diagnosis
   2. Summary: cost, clusters, CPU and memory use, cost of a minimum cluster
   3. Pattern 1: microservices by layer and capabilities repeated across domains
   4. Pattern 2: clusters and monthly cost by number of microservices
   5. Pattern 4: current replicas, HPA behaviour, time without scaling and inflated minimums
   6. Capacity health: unschedulable pods, OOM kills, CPU vs memory
   7. Why CPU is low and memory high: what uses node memory, requests vs use, memory per replica, hypotheses
   8. Configuration drift and clusters under the 4-node minimum
  Technical criteria
   9. Criteria AKS vs Container Apps with the PROD figure
  10. Cost under the ADT_019 rules
  Strategic front
  11. Application, configuration and releases
  12. Security
  13. Resilience
  14. Multicloud and key message
  Proposal
  15. Per-cluster comparison with measured replicas
  16. Traffic profile per API
  17. Quick wins and next steps
  Annex
  18. Business criticality (BIA), information only
  19. Sources and method

Only APPLICATION clusters enter the PROD figures, as in the workbook; the CSVs read here (signals, node and
workload usage) are filtered the same way with aa_workbook.cluster_categories.

Output: aks_aca_deck.pptx next to the workbook (or --output). Charts are native and editable in PowerPoint.
Every figure is an example to check against the workbook and the CSVs before presenting.
"""

import argparse
import csv
import glob
import os
import re
import statistics
from datetime import date

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

import aa_common as aa
from aa_workbook import cluster_categories

# Palette of the team's deck: royal blue header ending in a rust tip, navy cards, royal blue highlight boxes.
NAVY = RGBColor(0x1E, 0x3A, 0x78)
ROYAL = RGBColor(0x1F, 0x41, 0x8E)
RUST = RGBColor(0xB0, 0x45, 0x2C)
RUST_DARK = RGBColor(0x7A, 0x2C, 0x1F)
CARD = RGBColor(0x1F, 0x33, 0x55)
ORANGE = RGBColor(0xF2, 0xA1, 0x5A)
STEEL = RGBColor(0x4A, 0x8F, 0xB8)
SKY = RGBColor(0x5D, 0xAD, 0xE2)
GREEN = RGBColor(0x4C, 0xB0, 0x7A)
RED = RGBColor(0xE8, 0x57, 0x4A)
GREY = RGBColor(0x5B, 0x67, 0x70)
LIGHT = RGBColor(0xEE, 0xF1, 0xF5)
SOFT = RGBColor(0xC5, 0xCD, 0xD6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"
LAYER_LABELS = {"Layer business-*": "business", "Layer channel-*": "channel", "Layer app-*": "app",
                "Layer async-*": "async", "Layer support-*": "support", "Layer core-*": "core", "Other layers": "otras"}
# Row label in the Patterns sheet, label on the slide, colour.
BEHAVIOURS = [("Stays at its minimum", "Se queda en su mínimo", ORANGE),
              ("Scales between minimum and maximum", "Escala", STEEL),
              ("Reaches or nears its maximum", "Cerca del máximo", RUST), ("Minimum = maximum", "Mínimo = máximo", NAVY),
              ("No data", "Sin datos", SOFT)]
SMALL_BUCKETS = {"1", "2 to 3", "4 to 5", "6 to 10", "11 to 15"}
# Workbook codes and labels shown on the slides in Spanish.
TRAFFIC_CLASSES = {"OFFICE": "OFICINA", "EXTENDED": "EXTENDIDO", "24x7": "24/7", "NO_TRAFFIC": "SIN TRÁFICO"}
STACKS = {"Unidentified": "Sin identificar"}
MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
          "noviembre", "diciembre"]
SUMMARY_LABELS = {
    "months": "Months in the data", "clusters": "PROD AKS clusters (all categories)",
    "application": "PROD clusters: APPLICATION (in the analysis)",
    "no_applications": "PROD clusters: NO_APPLICATIONS (created without applications)",
    "infrastructure": "PROD clusters: INFRASTRUCTURE (tools only, excluded)",
    "no_data": "PROD clusters: NO_DATA (not in the platform export, excluded)",
    "no_applications_cost": "Monthly cost of PROD clusters without applications (finding)",
    "non_prod": "Non-production AKS clusters (Dev/QA)", "le4": "PROD application clusters with 4 nodes or fewer",
    "cost_avg": "Average monthly cost per PROD application cluster",
    "cost_total": "Total monthly cost of PROD application clusters",
    "cpu": "CPU use in office hours, PROD application clusters",
    "mem": "Memory use in office hours, PROD application clusters",
    "cpu_p95": "CPU use, hourly 95th percentile (whole day)",
    "cpu_max": "CPU use, maximum (upper bound: peaks of each cluster added up)",
    "mem_p95": "Memory use, hourly 95th percentile (whole day)",
    "mem_max": "Memory use, maximum (upper bound: peaks of each cluster added up)",
    "aks_min_list": "Minimum AKS cluster (4 nodes) at list price",
    "aks_min_effective": "Minimum AKS cluster with the measured discount",
    "aca_per_ms": "ACA per microservice (ADT schedule)",
}
MEASURED_LABELS = {"lt4_min": "PROD application clusters with a configured minimum below 4 nodes",
                   "lt4_now": "PROD application clusters with fewer than 4 nodes at inventory time"}
# Sources sheet values shown on the annex slide.
SOURCE_TEXT = {
    "Measured": "Medido", "Derived": "Derivado", "Estimated": "Estimado", "Vendor data": "Dato del proveedor",
    "Measured (actual billed cost)": "Medido (costo real facturado)", "Derived from actual cost": "Derivado de costo real",
    "Internal guideline": "Lineamiento interno", "Assumption": "Supuesto", "Measured (one day)": "Medido (un día)",
    "Inventory of clusters, node pools and add-ons": "Inventario de clusters, node pools y add-ons",
    "Container Apps apps and environments": "Apps y ambientes de Container Apps",
    "APIM APIs and their backend": "APIs de APIM y su backend",
    "Node CPU and memory use": "Uso de CPU y memoria de nodos", "Requests per API": "Peticiones por API",
    "Wake-ups and active seconds outside office hours": "Despertares y segundos activos fuera de horario",
    "Replicas, requests and use per workload": "Réplicas, requests y uso por workload",
    "Actual cost of AKS and Container Apps": "Costo real de AKS y Container Apps",
    "Actual node discount": "Descuento real de nodos", "List prices": "Precios de lista",
    "Container Apps minimum replica, schedule and limits": "Réplica mínima, horario y límites de Container Apps",
    "Load and capacity assumptions": "Supuestos de carga y capacidad",
    "ARM (read)": "ARM (lectura)", "Azure Monitor (platform metrics)": "Azure Monitor (métricas de plataforma)",
    "Architecture criterion": "Criterio de arquitectura", "Platform team export": "Export del equipo de plataforma",
    "Azure Retail Prices API (public)": "Azure Retail Prices API (pública)",
}
BIA_ORDER = ["MUY ALTA", "ALTA", "MEDIA", "BAJA", "SIN BIA"]
PROPERTY_LABELS = {
    "kubernetes_version": "Versión de Kubernetes", "node_os_upgrade_channel": "Actualización del SO de nodos",
    "workload_identity": "Workload identity", "oidc_issuer": "OIDC issuer",
    "disable_local_accounts": "Cuentas locales deshabilitadas", "service_mesh": "Service mesh (Istio)",
    "network_policy": "Network policy", "zones": "Zonas de disponibilidad", "azure_policy": "Azure Policy",
    "defender": "Defender for Containers", "gitops_flux": "GitOps (Flux)", "upgrade_channel": "Upgrade automático",
    "container_insights": "Container Insights", "keda": "KEDA", "sku_tier": "Tier del cluster",
    "network_plugin_mode": "Modo del plugin de red", "image_cleaner": "Image cleaner", "key_vault_csi": "Key Vault CSI",
}


def latest_workbook():
    paths = glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*", "aks_aca_comparison.xlsx"))
    if not paths:
        raise SystemExit("No reports/*/aks_aca_comparison.xlsx found. Run aa_workbook.py or pass --workbook.")
    return max(paths, key=os.path.getmtime)


# ---------------------------------------------------------------------------
# Reading the workbook and the signals
# ---------------------------------------------------------------------------


def sheet_rows(book, name):
    """Rows of a sheet as lists, trailing empty cells removed; [] when the sheet is missing."""
    if name not in book.sheetnames:
        return []
    rows = []
    for row in book[name].iter_rows(values_only=True):
        values = list(row)
        while values and values[-1] is None:
            values.pop()
        rows.append(values)
    return rows


def label_of(row):
    return row[0].strip() if row and isinstance(row[0], str) else ""


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def read_summary(book):
    values = {}
    for row in sheet_rows(book, "Summary"):
        if label_of(row) and len(row) > 1:
            values.setdefault(label_of(row), row[1])
    return values


def read_bia(book):
    rows, inside = {}, False
    for row in sheet_rows(book, "Summary"):
        label = label_of(row)
        if label == "BIA":
            inside = True
        elif inside and label and len(row) > 2 and is_number(row[1]):
            rows[label] = {"clusters": row[1], "cost": row[2]}
        elif inside and not label:
            inside = False
    return rows


def read_patterns(book):
    """Split the Patterns sheet by section and read each table by its labels."""
    data = {"layers": [], "capabilities": [], "clusters": [], "replicas": [], "behaviour": {}, "title": ""}
    section, part = "", ""
    behaviours = [label for label, _, _ in BEHAVIOURS]
    for row in sheet_rows(book, "Patterns"):
        label = label_of(row)
        if not label:
            continue
        row = row + [None] * (5 - len(row))
        if not data["title"]:
            data["title"] = label
        if label.startswith("Pattern 1"):
            section, part = "p1", ""
        elif label.startswith("Pattern 2"):
            section, part = "p2", ""
        elif label.startswith("Pattern 4"):
            section, part = "p4", ""
        elif section == "p1":
            if label == "PROD microservices":
                data["microservices"] = row[1]
            elif label in LAYER_LABELS:
                data["layers"].append((LAYER_LABELS[label], row[1], row[2]))
            elif label.startswith("Capability repeated"):
                part = "capabilities"
            elif part == "capabilities" and is_number(row[1]):
                data["capabilities"].append((label, row[1], row[2]))
        elif section == "p2":
            if label.startswith("PROD clusters with 15"):
                data["clusters_le15"], data["clusters_le15_share"] = row[1], row[2]
            elif label.startswith("Monthly cost of PROD clusters with 3"):
                data["cost_le3"] = row[3]
            elif is_number(row[1]):
                data["clusters"].append((label, row[1], row[2], row[3]))
        elif section == "p4":
            if label.startswith("Current replicas"):
                part = "replicas"
            elif label == "HPA behaviour":
                part = "behaviour"
            elif label.startswith("Time since the last scale"):
                part = "last_scale"
            elif label in behaviours:
                data["behaviour"][label] = {"export": row[1], "ci": row[3]}
            elif label.startswith("Does not scale"):
                data["no_scale"] = {"export": row[1], "ci": row[3]}
            elif label.startswith("Its minimum is above that even keeping"):
                data["inflated_beyond_ha"], data["inflated_beyond_ha_count"] = row[2], row[1]
            elif label.startswith("Its minimum is above"):
                data["inflated"], data["inflated_count"] = row[2], row[1]
            elif label.startswith("HPAs without a scale event for more than 3 months"):
                data["stale_3m"] = (row[1], row[2])
            elif label.startswith("HPAs without a scale event for more than 1 year"):
                data["stale_1y"] = (row[1], row[2])
            elif label.startswith("HPAs whose office maximum"):
                data["office_max_reached"] = (row[1], row[2])
            elif part == "replicas" and is_number(row[1]):
                data["replicas"].append((label, row[1], row[2]))
    match = re.search(r"\(configuration on ([0-9-]+)\)", data["title"])
    data["snapshot_date"] = match.group(1) if match else ""
    return data


def read_measured(book):
    """Measured_AKS: the PROD summary rows by label and the PROD cluster rows."""
    summary, clusters, header = {}, [], None
    for row in sheet_rows(book, "Measured_AKS"):
        label = label_of(row)
        if label == "Cluster":
            header = row
        elif header and len(row) > 1 and row[1] == "PROD":
            clusters.append({"name": label, "cost": row[5] if len(row) > 5 else None,
                             "cpu_use": row[8] if len(row) > 8 else None,
                             "workloads": row[11] if len(row) > 11 else None})
        elif label and len(row) > 1 and not (len(row) > 2 and row[2] is not None):
            summary.setdefault(label, row[1])
    return summary, clusters


def read_criteria(book):
    rows, criteria, started = sheet_rows(book, "Technical_criteria"), [], False
    keda = None
    for row in rows:
        row = row + [None] * (8 - len(row))
        label = label_of(row)
        if label == "Criterion":
            started = True
            continue
        if not started:
            continue
        if label.startswith("Separate scenario"):
            started = False
            continue
        if not label and row[3]:
            criteria[-1]["extra"].append((row[3], row[4], row[5]))
            continue
        if label:
            criteria.append({"name": label, "aca": row[1] or "", "aks": row[2] or "", "measure": row[3] or "",
                             "value": row[4], "total": row[5], "note": row[7] or "", "extra": []})
    for row in rows:
        if label_of(row) == "Status at the bank":
            row = row + [None] * (8 - len(row))
            keda = {"text": row[2], "value": row[4], "total": row[5]}
    return criteria, keda


def read_assumptions(book):
    values = {}
    for row in sheet_rows(book, "Assumptions"):
        if label_of(row) and len(row) > 2:
            values[label_of(row)] = row[2]
    return values


def read_model(book):
    return [row[:5] for row in sheet_rows(book, "Model") if row and is_number(row[0]) and len(row) >= 5]


def read_equilibrium(book):
    values, rules, in_rules = {}, [], False
    for row in sheet_rows(book, "Break_even"):
        label = label_of(row)
        if label == "ADT_019 rule":
            in_rules = True
            continue
        if in_rules:
            if not label or label.startswith("Other ADT_019 rules"):
                in_rules = False
                continue
            rules.append((row + [None] * 4)[:4])
        elif label and len(row) > 1:
            values.setdefault(label, row[1])
    return values, rules


def read_per_cluster(book):
    rows, header, total = [], None, None
    for row in sheet_rows(book, "ACA_by_cluster"):
        label = label_of(row)
        if label == "Cluster":
            header = row
        elif label.startswith("Total of the clusters"):
            total = row
        elif header and label and len(row) > 10:
            row = row + [None] * (21 - len(row))
            rows.append({"name": label, "workloads": row[1], "aks": row[5], "aca_use": row[7], "aca_adt": row[8],
                         "difference": row[9], "cheaper": row[10], "nodes": row[12], "nodes_rs": row[13],
                         "aks_rs": row[14], "cheaper_rs": row[16], "nodes_hpa": row[17], "aks_hpa": row[18],
                         "aca_hpa": row[19], "cheaper_hpa": row[20]})
    return rows, total


def read_profile(book):
    classes, inside = [], False
    for row in sheet_rows(book, "Traffic_profile"):
        label = label_of(row)
        if label == "Class":
            inside = True
        elif inside and label in TRAFFIC_CLASSES:
            classes.append((label, row[1], row[2]))
        elif inside and label == "APIM":
            break
    settings = {label_of(row): row[1] for row in sheet_rows(book, "Traffic_profile")[:9] if label_of(row) and len(row) > 1}
    return classes, settings


def read_config(book):
    properties = []
    for row in sheet_rows(book, "Config_AKS"):
        row = row + [None] * (7 - len(row))
        if label_of(row) and is_number(row[3]) and is_number(row[4]):
            properties.append({"name": label_of(row), "common": row[2], "other": row[3], "total": row[4]})
    return properties


def read_sources(book):
    rows, inside = [], False
    for row in sheet_rows(book, "Sources"):
        label = label_of(row)
        if label == "Data":
            inside = True
        elif inside and label:
            rows.append((row + [None] * 6)[:6])
    return rows


def read_csv(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8-sig", newline="") as handle:
        sample = handle.readline()
        handle.seek(0)
        return list(csv.DictReader(handle, delimiter=";" if sample.count(";") > sample.count(",") else ","))


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def application_clusters(results):
    """(ids of the PROD APPLICATION clusters, cluster name -> ids), with the workbook's categories; None without the
    inventory."""
    clusters = read_csv(os.path.join(results, "aa-inventory", "aks_clusters.csv"))
    if not clusters:
        return None
    snapshot = read_csv(os.path.join(results, "aa-k8s-snapshot", "cluster_snapshot.csv")) or []
    categories, _ = cluster_categories(clusters, snapshot)
    ids = {row["id"].lower() for row in clusters
           if row["environment"] == "PROD" and categories[row["id"].lower()][0] == aa.APPLICATION}
    return ids, aa.cluster_ids_by_name(clusters)


def in_applications(row, applications):
    """The row's cluster is a PROD APPLICATION cluster (any PROD cluster when the inventory is missing)."""
    if applications is None:
        return row.get("environment") == "PROD"
    ids, ids_by_name = applications
    return aa.cluster_id_of(row, ids_by_name) in ids


def read_signals(folder, applications=None):
    """Figures of aa_signals.py for PROD APPLICATION clusters; None for each part whose CSV is missing."""
    signals = {"autoscaler": None, "oom": None, "latency": None}
    autoscaler = read_csv(os.path.join(folder, "aks_autoscaler_monthly.csv"))
    if autoscaler is not None:
        prod = [row for row in autoscaler if in_applications(row, applications)]
        with_data = {row["cluster"] for row in prod if to_float(row["hours_with_data"]) > 0}
        short = {row["cluster"] for row in prod if to_float(row["hours_with_unschedulable"]) > 0}
        hours = sum(to_float(row["hours_with_data"]) for row in prod)
        waiting = sum(to_float(row["hours_with_unschedulable"]) for row in prod)
        signals["autoscaler"] = {"clusters": len(with_data), "short": len(short),
                                 "hours_share": waiting / hours if hours else None}
    oom = read_csv(os.path.join(folder, "aks_oomkilled_monthly.csv"))
    if oom is not None:
        apps = [row for row in oom if in_applications(row, applications) and row["system_namespace"] == "NO"
                and aa.is_microservice(row["namespace"], row["workload"])]
        by_workload = {}
        for row in apps:
            key = (row["cluster"], row["workload"])
            by_workload[key] = by_workload.get(key, 0) + to_float(row["oom_kills"])
        top = sorted(by_workload.items(), key=lambda item: -item[1])[:5]
        signals["oom"] = {"kills": sum(by_workload.values()), "workloads": len(by_workload),
                          "clusters": len({key[0] for key in by_workload}), "top": top}
    latency = read_csv(os.path.join(folder, "apim_latency_monthly.csv"))
    if latency:
        current = date.today().strftime("%Y-%m")
        months = sorted({row["month"] for row in latency if row["month"] < current}) or sorted({row["month"] for row in latency})
        month = months[-1]
        values = [to_float(row.get("backend_duration_avg_ms")) for row in latency
                  if row["month"] == month and row.get("backend_duration_avg_ms") not in (None, "")]
        if values:
            signals["latency"] = {"month": month, "apis": len(values), "median": statistics.median(values),
                                  "under_100": sum(value < 100 for value in values) / len(values),
                                  "over_1000": sum(value >= 1000 for value in values)}
    return signals


def read_footprint(results):
    """Node memory and CPU of PROD APPLICATION clusters in office hours, split by what uses it, for the last full month.

    Node usage is the platform metric (everything on the node); Container Insights gives the part of apps,
    DaemonSets and platform pods. Only clusters with Container Insights are compared, so both sides match.
    """
    clusters = read_csv(os.path.join(results, "aa-inventory", "aks_clusters.csv"))
    usage = read_csv(os.path.join(results, "aa-metrics", "aks_usage_monthly.csv"))
    workloads = read_csv(os.path.join(results, "aa-container-insights", "aks_workloads_monthly.csv"))
    if not (clusters and usage and workloads):
        return None
    applications = application_clusters(results)
    ids_by_name = applications[1]
    current = date.today().strftime("%Y-%m")
    months = sorted({row["month"] for row in workloads if row["month"] < current})
    if not months:
        return None
    month = months[-1]
    parts = {"apps": [0.0, 0.0, 0.0, 0.0], "daemonsets": [0.0, 0.0, 0.0, 0.0], "platform": [0.0, 0.0, 0.0, 0.0]}
    per_replica, with_ci = [], set()
    for row in workloads:
        if row["month"] != month or row["period"] != "OFFICE" or not in_applications(row, applications):
            continue
        with_ci.add(aa.cluster_id_of(row, ids_by_name))
        if row["kind"] == "DaemonSet":
            key = "daemonsets"
        elif row["system_namespace"] == "YES" or not aa.is_microservice(row["namespace"], row["workload"]):
            key = "platform"
        else:
            key = "apps"
            if row["kind"] == "ReplicaSet" and row.get("mem_used_gib_per_replica"):
                per_replica.append((to_float(row["mem_used_gib_per_replica"]), to_float(row.get("cpu_used_per_replica"))))
        for index, field in enumerate(("mem_used_gib_avg", "cpu_used_avg", "mem_request_gib_avg", "cpu_request_avg")):
            parts[key][index] += to_float(row.get(field))
    node = {"mem_used": 0.0, "mem_alloc": 0.0, "cpu_used": 0.0, "cpu_alloc": 0.0}
    for row in usage:
        if row["month"] != month or aa.cluster_id_of(row, ids_by_name) not in with_ci:
            continue
        if row["mode"] == "ALLOCATABLE":
            node["mem_alloc"] += to_float(row["mem_allocatable_avg_gib"])
            node["cpu_alloc"] += to_float(row["cpu_allocatable_avg"])
        else:
            node["mem_used"] += to_float(row["mem_used_office_avg_gib"])
            node["cpu_used"] += to_float(row["cpu_used_office_avg"])
    if not node["mem_used"] or not per_replica:
        return None
    buckets = [("< 128 MiB", 0, 0.125), ("128–512 MiB", 0.125, 0.5), ("512 MiB–1 GiB", 0.5, 1.0), ("≥ 1 GiB", 1.0, 10 ** 6)]
    memory = [value for value, _ in per_replica]
    apps = parts["apps"]
    stack = read_csv(os.path.join(results, "aa-signals", "aks_stack_memory.csv")) or []
    stack = [row for row in stack if in_applications(row, applications) and row["kind"] == "ReplicaSet"
             and row["system_namespace"] == "NO" and aa.is_microservice(row["namespace"], row["workload"])]
    app_memory = sum(to_float(row["app_mem_gib_per_replica"]) * to_float(row["replicas_avg"]) for row in stack)
    sidecar_memory = sum(to_float(row["sidecar_mem_gib_per_replica"]) * to_float(row["replicas_avg"]) for row in stack)
    stacks = {}
    for row in stack:
        name = STACKS.get(row["stack"], row["stack"])
        stacks[name] = stacks.get(name, 0) + 1
    return {
        "stack_workloads": len(stack), "stacks": sorted(stacks.items(), key=lambda item: -item[1]),
        "sidecar_share": share_of(sidecar_memory, app_memory + sidecar_memory) if stack else None,
        "with_sidecar": sum(1 for row in stack if row["sidecar_mem_gib_per_replica"] not in ("", None)),
        "month": month, "clusters": len(with_ci), "node": node, "parts": parts,
        "rest_mem": node["mem_used"] - sum(part[0] for part in parts.values()),
        "buckets": [(label, sum(low <= value < high for value in memory)) for label, low, high in buckets],
        "median_mem": statistics.median(memory), "median_cpu": statistics.median([cpu for _, cpu in per_replica]),
        "workloads": len(memory),
        "cpu_request_over_use": apps[3] / apps[1] if apps[1] else None,
        "mem_use_over_request": apps[0] / apps[2] if apps[2] else None,
        "ratio_used": apps[0] / apps[1] if apps[1] else None,
        "ratio_requested": apps[2] / apps[3] if apps[3] else None,
        "ratio_node": node["mem_alloc"] / node["cpu_alloc"] if node["cpu_alloc"] else None,
    }


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------


def money(value):
    return f"US${value:,.0f}".replace(",", ".") if is_number(value) else "-"


def number(value):
    return f"{value:,.0f}".replace(",", ".") if is_number(value) else "-"


def percent(value):
    if not is_number(value):
        return "-"
    return f"{value:.1%}".replace(".", ",") if 0 < value < 0.01 else f"{value:.0%}"


def spanish_period(months):
    """'octubre 2026' for "2026-10"; 'ago–oct 2026' or 'dic 2025–ene 2026' for "2026-08, 2026-09, 2026-10"."""
    months = sorted(re.findall(r"\d{4}-\d{2}", str(months or "")))
    if not months:
        return "el periodo medido"
    first, last = months[0], months[-1]
    if first == last:
        return f"{MONTHS[int(first[5:]) - 1]} {first[:4]}"
    short = {month: MONTHS[int(month[5:]) - 1][:3] for month in (first, last)}
    if first[:4] == last[:4]:
        return f"{short[first]}–{short[last]} {last[:4]}"
    return f"{short[first]} {first[:4]}–{short[last]} {last[:4]}"


def spanish_bucket(label):
    """Bucket labels of the workbook ("2 to 3", "More than 50") in Spanish."""
    return (str(label).replace(" to ", " a ").replace("More than", "Más de")
            .replace("(possible zombie)", "(posible zombi)").replace("(no microservice in the export)", "(sin microservicios)"))


def share_of(value, total):
    return value / total if is_number(value) and is_number(total) and total else None


def measured(value, total):
    """'104 de 140 (74%)', or a percentage alone when the value is a share."""
    if is_number(value) and is_number(total) and total:
        return f"{number(value)} de {number(total)} ({percent(value / total)})"
    if is_number(value) and value <= 1:
        return percent(value)
    return number(value) if is_number(value) else "Límite documentado"


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------


def text_box(slide, left, top, width, height, text, size=12, bold=False, color=GREY, align=PP_ALIGN.LEFT,
             italic=False, anchor=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = line
        run.font.name, run.font.size, run.font.bold, run.font.italic = FONT, Pt(size), bold, italic
        run.font.color.rgb = color
    return shape


def rounded(slide, left, top, width, height, fill, radius=0.12):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.shadow.inherit = False
    box.adjustments[0] = radius
    return box


def gradient(shape, angle=0):
    shape.line.fill.background()
    shape.shadow.inherit = False
    shape.fill.gradient()
    shape.fill.gradient_angle = angle
    stops = shape.fill.gradient_stops
    stops[0].color.rgb, stops[0].position = NAVY, 0.0
    stops[1].color.rgb, stops[1].position = RUST, 1.0


def banner(slide, deck, title):
    """Title band as in the team's deck: a royal blue arrow over a rust band that darkens to the right."""
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.slide_width, Inches(1.05))
    band.line.fill.background()
    band.shadow.inherit = False
    band.fill.gradient()
    band.fill.gradient_angle = 0
    stops = band.fill.gradient_stops
    stops[0].color.rgb, stops[0].position = RUST, 0.7
    stops[1].color.rgb, stops[1].position = RUST_DARK, 1.0
    arrow = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, 0, 0, int(deck.slide_width * 0.78), Inches(1.05))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = NAVY
    arrow.line.fill.background()
    arrow.shadow.inherit = False
    arrow.adjustments[0] = 0.45
    text_box(slide, Inches(0.45), 0, Inches(12.2), Inches(1.05), title.upper(), size=19 if len(title) <= 68 else 16,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def base_slide(deck, title, source, section=""):
    slide = deck.slides.add_slide(deck.slide_layouts[6])  # blank layout
    banner(slide, deck, title)
    text_box(slide, Inches(0.45), Inches(6.98), Inches(10.6), Inches(0.25), source, size=8, italic=True)
    if section:
        text_box(slide, Inches(11.2), Inches(7.05), Inches(1.2), Inches(0.35), section, size=8, color=RUST,
                 bold=True, align=PP_ALIGN.RIGHT)
    text_box(slide, Inches(12.4), Inches(7.05), Inches(0.5), Inches(0.35), str(len(deck.slides)), size=9,
             align=PP_ALIGN.RIGHT)
    return slide


def heading(slide, left, top, width, text, size=13):
    text_box(slide, left, top, width, Inches(0.35), text, size=size, bold=True, color=NAVY)


def kpi(slide, left, top, width, height, value, label, value_color=ORANGE, value_size=30, label_size=11):
    rounded(slide, left, top, width, height, CARD)
    text_box(slide, left, top + Inches(0.08), width, height * 0.55, value, size=value_size, bold=True,
             color=value_color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, left + Inches(0.12), top + height * 0.58, width - Inches(0.24), height * 0.4, label,
             size=label_size, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def callout(slide, left, top, width, height, text, size=15, fill=CARD):
    rounded(slide, left, top, width, height, fill, radius=0.18)
    text_box(slide, left + Inches(0.3), top, width - Inches(0.6), height, text, size=size, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def note_box(slide, left, top, width, height, text, size=10):
    """Light box for method notes and pending items."""
    box = rounded(slide, left, top, width, height, LIGHT, radius=0.1)
    box.line.color.rgb = SOFT
    text_box(slide, left + Inches(0.2), top, width - Inches(0.4), height, text, size=size, color=GREY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def bar_chart(slide, left, top, width, height, categories, values, kind, colors, number_format="#,##0"):
    data = CategoryChartData()
    data.categories = categories
    data.add_series("Serie", values)
    chart = slide.shapes.add_chart(kind, left, top, width, height, data).chart
    chart.font.name, chart.font.size = FONT, Pt(11)
    chart.font.color.rgb = GREY
    chart.has_legend, chart.has_title = False, False
    plot = chart.plots[0]
    plot.gap_width = 55
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format, labels.number_format_is_linked = number_format, False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    labels.font.bold = True
    for index, color in enumerate(colors):
        point = plot.series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False
    chart.category_axis.format.line.color.rgb = SOFT
    return chart


def line_chart(slide, left, top, width, height, categories, series):
    data = CategoryChartData()
    data.categories = categories
    for name, values, _ in series:
        data.add_series(name, values)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE, left, top, width, height, data).chart
    chart.font.name, chart.font.size = FONT, Pt(10)
    chart.font.color.rgb = GREY
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = LIGHT
    chart.value_axis.tick_labels.number_format, chart.value_axis.tick_labels.number_format_is_linked = '"$"#,##0', False
    chart.category_axis.format.line.color.rgb = SOFT
    for plot_series, (_, _, color) in zip(chart.plots[0].series, series):
        plot_series.format.line.color.rgb = color
        plot_series.format.line.width = Pt(2.5)
        plot_series.smooth = False
    return chart


def table(slide, left, top, width, rows, widths, highlight_column=None, row_height=0.34, size=11, align_first=True):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), left, top, width, Inches(row_height * len(rows)))
    grid = shape.table
    for column, share in enumerate(widths):
        grid.columns[column].width = Emu(int(width * share))
    for r, values in enumerate(rows):
        grid.rows[r].height = Inches(row_height)
        for c, value in enumerate(values):
            cell = grid.cell(r, c)
            cell.text = str(value) or " "  # an empty cell has no run to style
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD if r == 0 else (LIGHT if r % 2 else WHITE)
            for paragraph in cell.text_frame.paragraphs:  # a value with line breaks has one paragraph per line
                paragraph.alignment = PP_ALIGN.LEFT if (c == 0 or not align_first) else PP_ALIGN.CENTER
                for run in paragraph.runs:
                    font = run.font
                    font.name, font.size = FONT, Pt(size)
                    if r == 0:
                        font.bold, font.color.rgb = True, WHITE
                    else:
                        font.color.rgb = RUST if c == highlight_column else GREY
                        font.bold = c == highlight_column or (c == 0 and not align_first)
    return shape


def legend_item(slide, left, top, color, text, width=Inches(3.4)):
    swatch = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + Inches(0.05), Inches(0.18), Inches(0.18))
    swatch.fill.solid()
    swatch.fill.fore_color.rgb = color
    swatch.line.fill.background()
    swatch.shadow.inherit = False
    text_box(slide, left + Inches(0.25), top, width, Inches(0.3), text, size=10)


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def slide_cover(deck, taken_on, period="el periodo medido"):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.slide_width, deck.slide_height)
    gradient(background, angle=0)
    text_box(slide, Inches(0.8), Inches(1.4), Inches(11.5), Inches(1.2), "MICROSERVICIOS EN AKS", size=40, bold=True,
             color=WHITE)
    text_box(slide, Inches(0.8), Inches(2.45), Inches(11.5), Inches(0.8),
             "Diagnóstico de producción y criterios frente a Azure Container Apps", size=22, color=WHITE)
    text_box(slide, Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.5),
             f"Costo amortizado y métricas de {period} · configuración al {taken_on or 'día del export'}",
             size=13, color=SOFT)
    sections = [("1", "Diagnóstico", "Qué hay hoy en PROD"), ("2", "Criterios técnicos", "Qué decide la plataforma"),
                ("3", "Frente estratégico", "Seguridad, resiliencia, multicloud"),
                ("4", "Propuesta", "Comparación real y siguientes pasos")]
    for index, (number_, name, detail) in enumerate(sections):
        left = Inches(0.8 + index * 3.0)
        rounded(slide, left, Inches(4.6), Inches(2.8), Inches(1.55), CARD)
        text_box(slide, left, Inches(4.7), Inches(2.8), Inches(0.6), number_, size=26, bold=True, color=ORANGE,
                 align=PP_ALIGN.CENTER)
        text_box(slide, left, Inches(5.25), Inches(2.8), Inches(0.4), name, size=15, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        text_box(slide, left + Inches(0.1), Inches(5.65), Inches(2.6), Inches(0.4), detail, size=10, color=SOFT,
                 align=PP_ALIGN.CENTER)
    text_box(slide, Inches(0.8), Inches(6.7), Inches(11.5), Inches(0.4),
             "Cifras generadas desde el workbook: verificar antes de presentar", size=10,
             italic=True, color=SOFT)


def slide_summary(deck, s, p, measured_summary, clusters, source):
    total = s.get(SUMMARY_LABELS["cost_total"])
    cpu, mem = s.get(SUMMARY_LABELS["cpu"]), s.get(SUMMARY_LABELS["mem"])
    slide = base_slide(deck, "AKS PROD combina alto costo fijo con baja utilización de CPU", source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.2), Inches(8), "Estado actual y señales principales (clusters con aplicaciones)")
    tiles = [(number(s.get(SUMMARY_LABELS["application"])), "clusters PROD con aplicaciones", SKY),
             (money(total), "costo mensual", ORANGE),
             (money(s.get(SUMMARY_LABELS["cost_avg"])), "promedio por cluster", ORANGE),
             (number(s.get(SUMMARY_LABELS["le4"])), "clusters con ≤4 nodos", SKY),
             (percent(cpu), "CPU en oficina", RED),
             (percent(mem), "memoria en oficina", GREEN),
             (percent(p.get("clusters_le15_share")), "clusters con ≤15 MS", GREEN),
             (percent(no_scale_counts(p)[0]), "HPA no escala", RED)]
    for index, (value, label, color) in enumerate(tiles):
        row, column = divmod(index, 4)
        kpi(slide, Inches(0.45 + column * 3.13), Inches(1.6 + row * 1.35), Inches(2.95), Inches(1.2), value, label,
            value_color=color, value_size=26, label_size=10)

    heading(slide, Inches(0.45), Inches(4.4), Inches(6.5), "Uso de la capacidad asignable (clusters PROD con aplicaciones)",
            size=12)
    rows = [["", "Promedio en oficina", "p95 por hora", "Máximo*"],
            ["CPU", percent(cpu), percent(s.get(SUMMARY_LABELS["cpu_p95"])), percent(s.get(SUMMARY_LABELS["cpu_max"]))],
            ["Memoria", percent(mem), percent(s.get(SUMMARY_LABELS["mem_p95"])), percent(s.get(SUMMARY_LABELS["mem_max"]))]]
    table(slide, Inches(0.45), Inches(4.75), Inches(6.4), rows, [0.22, 0.28, 0.25, 0.25], row_height=0.36, size=10)
    text_box(slide, Inches(0.45), Inches(5.85), Inches(6.4), Inches(0.3),
             "* Suma del pico de cada cluster: cota superior, porque los picos no ocurren a la misma hora.", size=8,
             italic=True)
    note_box(slide, Inches(0.45), Inches(6.2), Inches(6.4), Inches(0.55),
             f"PARQUE · {number(s.get(SUMMARY_LABELS['clusters']))} clusters PROD: "
             f"{number(s.get(SUMMARY_LABELS['application']))} con aplicaciones, "
             f"{number(s.get(SUMMARY_LABELS['no_applications']))} sin aplicaciones "
             f"({money(s.get(SUMMARY_LABELS['no_applications_cost']))} al mes), "
             f"{number(s.get(SUMMARY_LABELS['infrastructure']))} de infraestructura y "
             f"{number(s.get(SUMMARY_LABELS['no_data']))} sin dato (excluidos)"
             + (f" · {number(s.get(SUMMARY_LABELS['non_prod']))} Dev/QA"
                if is_number(s.get(SUMMARY_LABELS["non_prod"])) else ""), size=9)

    heading(slide, Inches(7.3), Inches(4.4), Inches(5.6), "Lo que cuesta el modelo de hoy", size=12)
    kpi(slide, Inches(7.3), Inches(4.75), Inches(2.7), Inches(1.3), money(s.get(SUMMARY_LABELS["aks_min_effective"])),
        "cluster AKS mínimo (4 nodos) al mes", value_size=24, label_size=10)
    kpi(slide, Inches(10.2), Inches(4.75), Inches(2.7), Inches(1.3), money(s.get(SUMMARY_LABELS["aca_per_ms"])),
        "un microservicio en ACA al mes (horario ADT)", value_size=24, label_size=10)
    callout(slide, Inches(7.3), Inches(6.2), Inches(5.6), Inches(0.68),
            "Cada aplicación tiene su propio cluster: el costo lo manda la capacidad reservada, no el consumo.",
            size=11, fill=ROYAL)


def slide_pattern1(deck, p, source):
    capabilities = sorted(p["capabilities"], key=lambda item: -item[1])
    most = capabilities[0][1] if capabilities else 0
    slide = base_slide(deck, f"Patrón 1: {number(p.get('microservices'))} microservicios; capacidades repetidas en "
                             f"hasta {number(most)} dominios", source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.3), Inches(6.3), "Microservicios por capa")
    layers = sorted(p["layers"], key=lambda item: item[1])
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(6.5), Inches(4.0),
              [f"{name} {percent(share)}" for name, _, share in layers], [item[1] for item in layers],
              XL_CHART_TYPE.BAR_CLUSTERED, [ORANGE if item[0] in ("business", "channel") else STEEL for item in layers])
    two = sum(item[2] or 0 for item in layers if item[0] in ("business", "channel"))
    callout(slide, Inches(0.45), Inches(5.8), Inches(6.3), Inches(0.85),
            f"business y channel concentran el {percent(two)} de los microservicios", size=13)

    heading(slide, Inches(7.2), Inches(1.3), Inches(5.7), "Capacidades repetidas en varios dominios")
    rows = [["Capacidad", "Dominios", "Microservicios"]] + [[name, number(domains), number(count)]
                                                           for name, domains, count in capabilities[:9]]
    table(slide, Inches(7.2), Inches(1.7), Inches(5.7), rows, [0.5, 0.22, 0.28], highlight_column=1)
    total_dup = sum(item[2] or 0 for item in capabilities)
    callout(slide, Inches(7.2), Inches(5.8), Inches(5.7), Inches(0.85),
            f"{number(total_dup)} microservicios repiten {len(capabilities)} capacidades: candidatas a servicio "
            "compartido", size=13)
    text_box(slide, Inches(7.2), Inches(1.7) + Inches(0.34 * len(rows)) + Inches(0.1), Inches(5.7), Inches(0.6),
             "Detectadas por el nombre del deployment; cada dominio debe confirmarlas.", size=9, italic=True)


def slide_pattern2(deck, p, source):
    buckets = [item for item in p["clusters"] if not item[0].startswith("0")]
    empty = next((item for item in p["clusters"] if item[0].startswith("0")), None)
    total = sum(item[1] for item in buckets)
    slide = base_slide(deck, f"Patrón 2: {number(p.get('clusters_le15'))} de {number(total)} clusters tienen 15 "
                             "microservicios o menos", source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.3), Inches(8), "Clusters PROD con aplicaciones por número de microservicios")
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(8.2), Inches(4.3), [spanish_bucket(item[0]) for item in buckets],
              [item[1] for item in buckets], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0] in SMALL_BUCKETS else STEEL for item in buckets])
    legend_item(slide, Inches(0.6), Inches(6.05), ORANGE, "15 microservicios o menos (umbral ADT_019)")
    legend_item(slide, Inches(4.3), Inches(6.05), STEEL, "Más de 15")
    if empty and empty[1]:
        text_box(slide, Inches(0.6), Inches(6.4), Inches(7.8), Inches(0.4),
                 f"{number(empty[1])} clusters con aplicaciones no tienen microservicios en el export, solo su gateway o "
                 f"cargas vistas por Container Insights ({money(empty[3])} al mes): por revisar.", size=10, italic=True)

    kpi(slide, Inches(8.85), Inches(1.35), Inches(1.95), Inches(1.45), percent(p.get("clusters_le15_share")),
        "bajo el umbral de 15", value_size=26)
    kpi(slide, Inches(10.95), Inches(1.35), Inches(1.95), Inches(1.45), money(p.get("cost_le3")),
        "al mes, clusters con 3 o menos", value_size=20)
    rows = [["Microservicios", "Clusters", "Costo/mes"]] + [[spanish_bucket(item[0]), number(item[1]), money(item[3])]
                                                             for item in buckets]
    table(slide, Inches(8.85), Inches(3.05), Inches(4.05), rows, [0.38, 0.24, 0.38], row_height=0.33, size=10)
    callout(slide, Inches(8.85), Inches(6.2), Inches(4.05), Inches(0.65),
            "La fragmentación multiplica mínimos de nodos, operación y componentes comunes.", size=10, fill=ROYAL)


def no_scale_counts(p):
    """(share of HPAs that do not scale, of them at their minimum, of them with min = max), from Container Insights."""
    behaviour = p["behaviour"]
    counts = [behaviour.get(label, {}).get("ci") or 0 for label, _, _ in BEHAVIOURS]
    total = sum(counts)
    never, fixed = counts[0], counts[3]
    no_scale = (p.get("no_scale") or {}).get("ci")
    no_scale = no_scale if is_number(no_scale) else never + fixed
    return share_of(no_scale, total), share_of(never, total), share_of(fixed, total)


def slide_pattern4(deck, p, source, period="el periodo"):
    behaviour = p["behaviour"]
    counts = [behaviour.get(label, {}).get("ci") or 0 for label, _, _ in BEHAVIOURS]
    total = sum(counts)
    no_scale, never, fixed = no_scale_counts(p)
    slide = base_slide(deck, f"Patrón 4: {percent(no_scale)} de los HPA no escala en {period}", source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.3), Inches(6.5), "HPA por réplicas actuales (día de la foto)")
    replicas = p["replicas"]
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(6.6), Inches(4.2), [spanish_bucket(item[0]) for item in replicas],
              [item[1] for item in replicas], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0].startswith("2 to 3") else STEEL for item in replicas])
    peak = max(replicas, key=lambda item: item[1]) if replicas else ("", 0, 0)
    callout(slide, Inches(0.45), Inches(5.95), Inches(6.4), Inches(0.75),
            f"El {percent(peak[2])} de los HPA corre con {spanish_bucket(peak[0])} réplicas", size=13)

    heading(slide, Inches(7.2), Inches(1.3), Inches(5.7), f"Comportamiento en {period} (promedio en oficina)")
    data = CategoryChartData()
    data.categories = [label for _, label, _ in BEHAVIOURS]
    data.add_series("Comportamiento", counts)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(7.1), Inches(1.6), Inches(3.6), Inches(3.5),
                                   data).chart
    chart.font.name, chart.font.size = FONT, Pt(11)
    chart.has_title, chart.has_legend = False, False
    plot = chart.plots[0]
    for index, (_, _, color) in enumerate(BEHAVIOURS):
        point = plot.series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
    text_box(slide, Inches(7.1), Inches(2.95), Inches(3.6), Inches(0.8), percent(no_scale), size=28, bold=True,
             color=ORANGE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for index, ((_, label, color), count) in enumerate(zip(BEHAVIOURS, counts)):
        legend_item(slide, Inches(10.75), Inches(1.8 + index * 0.5), color,
                    f"{label}: {percent(count / total if total else None)}", width=Inches(2.3))
    lines = [f"No escala: {percent(no_scale)} = {percent(never)} se queda en su mínimo + {percent(fixed)} con "
             "mínimo = máximo"]
    stale_3m, stale_1y = p.get("stale_3m"), p.get("stale_1y")
    if stale_3m and stale_1y:
        lines.append(f"Sin escalar hace más de 3 meses: {percent(stale_3m[1])} de los HPA; más de un año: "
                     f"{percent(stale_1y[1])}")
    text_box(slide, Inches(7.2), Inches(5.15), Inches(5.7), Inches(0.75), lines, size=10, bold=True, color=NAVY)
    message = "Los mínimos de réplicas son capacidad que se paga 24/7"
    if is_number(p.get("inflated")):
        message = (f"En el {percent(p['inflated'])} de los que se quedan en su mínimo, el mínimo supera lo que pide su "
                   f"CPU ({percent(p.get('inflated_beyond_ha'))} aun dejando 2 réplicas por alta disponibilidad)")
    callout(slide, Inches(7.2), Inches(5.95), Inches(5.7), Inches(0.75), message, size=11)


def slide_capacity_health(deck, s, signals, source, period="el periodo"):
    cpu, mem = s.get(SUMMARY_LABELS["cpu"]), s.get(SUMMARY_LABELS["mem"])
    autoscaler, oom = signals["autoscaler"], signals["oom"]
    hours = autoscaler["hours_share"] if autoscaler else None
    slide = base_slide(deck, "El bajo uso no es falta de capacidad ni de memoria", source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.2), Inches(12), "¿Falta capacidad o memoria? Lo que dicen el autoscaler y "
                                                          "los contenedores")
    tiles = [
        (f"{number(autoscaler['short'])} de {number(autoscaler['clusters'])}" if autoscaler else "-",
         "clusters con autoscaler tuvieron pods esperando nodo", SKY),
        (percent(hours) if autoscaler else "-", "de las horas con pods esperando nodo", GREEN),
        (number(oom["kills"]) if oom else "-", f"contenedores OOMKilled (apps, {period})", RED),
        (number(oom["workloads"]) if oom else "-",
         f"workloads con OOMKilled en {number(oom['clusters'])} clusters" if oom else "workloads con OOMKilled", ORANGE)]
    for index, (value, label, color) in enumerate(tiles):
        kpi(slide, Inches(0.45 + index * 3.13), Inches(1.6), Inches(2.95), Inches(1.3), value, label,
            value_color=color, value_size=24, label_size=10)

    heading(slide, Inches(0.45), Inches(3.15), Inches(6), "Uso en oficina sobre la capacidad asignable", size=12)
    bar_chart(slide, Inches(0.3), Inches(3.45), Inches(6.0), Inches(2.6), ["Memoria", "CPU"],
              [mem if is_number(mem) else 0, cpu if is_number(cpu) else 0], XL_CHART_TYPE.BAR_CLUSTERED,
              [GREEN, RED], number_format="0%")
    heading(slide, Inches(6.9), Inches(3.15), Inches(6), "Workloads con más OOMKilled", size=12)
    if oom and oom["top"]:
        rows = [["Cluster", "Workload", "OOMKilled"]] + [[cluster, workload, number(kills)]
                                                        for (cluster, workload), kills in oom["top"]]
        table(slide, Inches(6.9), Inches(3.5), Inches(6.0), rows, [0.3, 0.55, 0.15], row_height=0.36, size=9,
              highlight_column=2)
    else:
        note_box(slide, Inches(6.9), Inches(3.5), Inches(6.0), Inches(1.2),
                 "Sin datos de OOMKilled: correr aa_signals.py" if not oom else "Ningún OOMKilled en aplicaciones")
    if not autoscaler:
        note_box(slide, Inches(0.45), Inches(6.1), Inches(12.45), Inches(0.75),
                 "Para completar esta lámina: python aks-aca/src/aa_signals.py (unschedulable pods, OOMKilled, latencia)")
        return
    message = (f"Se usa {percent(cpu)} de la CPU y la memoria ({percent(mem)}) es lo que dimensiona los nodos. "
               "Los OOMKilled se concentran en pocos workloads: es memoria mal ajustada en esos servicios, no falta de nodos.")
    callout(slide, Inches(0.45), Inches(6.15), Inches(12.45), Inches(0.75), message, size=12)


def slide_footprint(deck, s, footprint, source):
    cpu, mem = s.get(SUMMARY_LABELS["cpu"]), s.get(SUMMARY_LABELS["mem"])
    slide = base_slide(deck, f"Por qué la CPU está en {percent(cpu)} y la memoria en {percent(mem)}: hipótesis", source,
                       "Diagnóstico")
    if not footprint:
        note_box(slide, Inches(0.45), Inches(3.0), Inches(12.45), Inches(1.0),
                 "Faltan aks_workloads_monthly.csv o aks_usage_monthly.csv en results/: correr aa_metrics.py y "
                 "aa_container_insights.py")
        return
    node, parts = footprint["node"], footprint["parts"]
    heading(slide, Inches(0.45), Inches(1.2), Inches(6.2), f"Qué ocupa la memoria usada de los nodos ({footprint['month']})",
            size=12)
    items = [("SO, kubelet y caché", footprint["rest_mem"], SOFT), ("Plataforma", parts["platform"][0], NAVY),
             ("DaemonSets", parts["daemonsets"][0], STEEL), ("Apps (con sidecar y agente)", parts["apps"][0], ORANGE)]
    bar_chart(slide, Inches(0.3), Inches(1.5), Inches(6.2), Inches(2.3), [label for label, _, _ in items],
              [share_of(value, node["mem_used"]) or 0 for _, value, _ in items], XL_CHART_TYPE.BAR_CLUSTERED,
              [color for _, _, color in items], number_format="0%")
    heading(slide, Inches(0.45), Inches(3.9), Inches(6.2), "Memoria usada por réplica de los microservicios", size=12)
    bar_chart(slide, Inches(0.3), Inches(4.2), Inches(6.2), Inches(1.75), [label for label, _ in footprint["buckets"]],
              [share_of(count, footprint["workloads"]) or 0 for _, count in footprint["buckets"]],
              XL_CHART_TYPE.COLUMN_CLUSTERED, [GREEN, STEEL, ORANGE, RUST], number_format="0%")

    tiles = [(f"{footprint['cpu_request_over_use']:.1f}×".replace(".", ",") if footprint["cpu_request_over_use"] else "-",
              "CPU reservada (request) sobre la usada", RED),
             (f"{footprint['mem_use_over_request']:.1f}×".replace(".", ",") if footprint["mem_use_over_request"] else "-",
              "memoria usada sobre la reservada", GREEN),
             (f"{footprint['median_mem']:.2f} GiB".replace(".", ","), "memoria mediana por réplica", ORANGE),
             (f"{footprint['median_cpu'] * 1000:.0f} m", "CPU mediana por réplica (milicores)", SKY)]
    for index, (value, label, color) in enumerate(tiles):
        row, column = divmod(index, 2)
        kpi(slide, Inches(6.9 + column * 3.05), Inches(1.25 + row * 1.15), Inches(2.9), Inches(1.05), value, label,
            value_color=color, value_size=20, label_size=9)
    ratio = (f"Proporción vCPU : GiB de las apps · reservada 1 : {footprint['ratio_requested']:.1f} · usada 1 : "
             f"{footprint['ratio_used']:.0f} · forma del nodo 1 : {footprint['ratio_node']:.1f}").replace(".", ",")
    note_box(slide, Inches(6.9), Inches(3.6), Inches(5.95), Inches(0.5), ratio, size=9)
    heading(slide, Inches(6.9), Inches(4.2), Inches(5.9), "Hipótesis (a validar con los equipos)", size=12)
    stack_line = "• JVM: reserva su heap aunque no haya tráfico; casi nada corre en modo nativo (Quarkus nativo < 128 MiB)."
    if footprint["stack_workloads"]:
        detected = ", ".join(f"{name} {percent(share_of(count, footprint['stack_workloads']))}"
                             for name, count in footprint["stacks"][:3])
        stack_line = f"• Stack detectado: {detected}. La JVM reserva su heap aunque no haya tráfico."
    sidecar_line = "• Cada pod suma el sidecar de Istio (si está en la malla) y un agente Java inyectado (javaagent)."
    if footprint["sidecar_share"] is not None:
        sidecar_line = (f"• El sidecar de Istio es el {percent(footprint['sidecar_share'])} de la memoria de los pods "
                        f"({number(footprint['with_sidecar'])} de {number(footprint['stack_workloads'])} workloads con "
                        "sidecar); además, un agente Java inyectado (javaagent).")
    text_box(slide, Inches(6.9), Inches(4.5), Inches(5.95), Inches(1.6), [
        stack_line, sidecar_line,
        "• Kubernetes ubica los pods por lo que reservan (request), no por lo que usan: esa CPU queda apartada en el "
        "nodo aunque nadie la use, y el autoscaler suma nodos por ella.",
        "• La parte del SO, kubelet y caché no es de las apps."], size=9)
    callout(slide, Inches(0.45), Inches(6.1), Inches(12.45), Inches(0.8),
            "Las apps se configuran 1:4 pero se comportan como cargas de memoria: bajar requests de CPU, ajustar el heap "
            "y evaluar nodos de memoria (serie E) antes de sumar nodos.", size=11)


def slide_configuration(deck, measured_summary, properties, criteria, source):
    upgrade = next((item for item in criteria if item["measure"].startswith("Clusters without automatic")), None)
    under = measured_summary.get(MEASURED_LABELS["lt4_min"])
    slide = base_slide(deck, f"La configuración se fragmenta y {number(under)} clusters están bajo el mínimo de 4 nodos",
                       source, "Diagnóstico")
    heading(slide, Inches(0.45), Inches(1.3), Inches(7), "Clusters PROD con aplicaciones con un valor distinto al más común")
    ranked = sorted([item for item in properties if item["name"] != "minor_versions_behind" and item["other"]],
                    key=lambda item: -item["other"])[:8]
    ranked.reverse()
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(7.2), Inches(4.4),
              [PROPERTY_LABELS.get(item["name"], item["name"]) for item in ranked], [item["other"] for item in ranked],
              XL_CHART_TYPE.BAR_CLUSTERED, [STEEL] * len(ranked))
    text_box(slide, Inches(0.45), Inches(6.1), Inches(7), Inches(0.4),
             "Dispersión respecto del valor más frecuente en PROD; el cumplimiento de la LBS lo valida el equipo de "
             "lineamientos.", size=9, italic=True)
    version = next((item for item in properties if item["name"] == "kubernetes_version"), None)
    tiles = [(number(under), "clusters con mínimo configurado bajo 4 nodos (2 sistema + 2 user)", RED),
             (number(measured_summary.get(MEASURED_LABELS["lt4_now"])),
              "clusters corriendo hoy con menos de 4 nodos", ORANGE),
             (number(version["other"]) if version else "-", "clusters fuera de la versión de Kubernetes más común", SKY),
             (f"{number(upgrade['value'])} de {number(upgrade['total'])}" if upgrade else "-",
              "clusters con upgrade de versión manual", SKY)]
    for index, (value, label, color) in enumerate(tiles):
        row, column = divmod(index, 2)
        kpi(slide, Inches(7.9 + column * 2.55), Inches(1.4 + row * 1.6), Inches(2.4), Inches(1.45), value, label,
            value_color=color, value_size=24, label_size=9)
    note_box(slide, Inches(7.9), Inches(4.75), Inches(4.95), Inches(0.75),
             "Nivel oficial de incumplimiento de la LBS: [dato del equipo de lineamientos]", size=10)
    callout(slide, Inches(7.9), Inches(5.65), Inches(4.95), Inches(1.0),
            "Cada cluster nace con la línea base del momento y se queda ahí: mantenerlos al día cuesta horas-hombre "
            "a cada equipo.", size=11, fill=ROYAL)


# Short wording for the figure of each criterion, keyed by the start of the measure in Technical_criteria.
MEASURE_LABELS = [("Clusters with 15 microservices", "clusters con ≤15 microservicios"),
                  ("APIs with a peak above 100 tps", "APIs con pico de más de 100 tps"),
                  ("Requests outside office hours", "de las peticiones llega fuera de oficina"),
                  ("HPAs that do not scale", "HPA no escala"),
                  ("Microservices that fit", "microservicios caben en una réplica de ACA"),
                  ("StatefulSet workloads", "workloads son StatefulSet"),
                  ("Clusters with managed Istio", "clusters con Istio"), ("Clusters with GitOps", "clusters con GitOps (Flux)"),
                  ("Clusters without automatic", "clusters con upgrade manual")]


def figure_text(measure, value, total):
    """'75% · 99 de 132 clusters con ≤15 microservicios', or '44% de las peticiones ...' for a share."""
    label = next((short for start, short in MEASURE_LABELS if measure.startswith(start)), measure)
    if is_number(value) and is_number(total) and total:
        return f"{percent(value / total)} · {number(value)} de {number(total)} {label}"
    if is_number(value) and value <= 1:
        return f"{percent(value)} {label}"
    return f"{number(value)} {label}" if is_number(value) else label


# Decision rules of ADT_019 shown on the slide, with the measure of Technical_criteria that backs each one.
DECISION_RULES = [("15 microservicios o menos por aplicación", "Clusters with 15 microservices"),
                  ("Menos de 100 tps de carga", "APIs with a peak above 100 tps"),
                  ("Uso concentrado en horario, no constante más de 10 h al día", "Requests outside office hours"),
                  ("Réplicas de hasta 4 vCPU / 8 GiB", "Microservices that fit"),
                  ("Sin StatefulSet, batch largo ni contenedores privilegiados", "StatefulSet workloads")]
# ADT_019 rules of the Break_even sheet, in its order, as shown on the cost slide.
ADT_RULES = ["Más de 15 microservicios con uso constante → AKS", "Más de 10 h al día de uso constante → AKS",
             "Encendido 24/7 cuesta más que AKS", "Más de 100 tps de carga total → AKS"]


def slide_criteria(deck, criteria, signals, bia, source):
    slide = base_slide(deck, "Cómo decidimos si un AKS puede pasar a Container Apps", source, "Criterios técnicos")
    by_measure = {}
    for item in criteria:
        by_measure[item["measure"]] = item
        for measure, value, total in item["extra"]:
            by_measure[measure] = {"measure": measure, "value": value, "total": total}
    rows = [["#", "Regla de decisión (ADT_019)", "Hoy en producción"]]
    for rule, measure in DECISION_RULES:
        item = next((value for key, value in by_measure.items() if key.startswith(measure)), None)
        rows.append([str(len(rows)), rule, figure_text(item["measure"], item["value"], item["total"]) if item else "-"])
    latency = signals.get("latency")
    rows.append([str(len(rows)), "El consumidor tolera ~30 s en el primer request tras escalar a 0",
                 f"{latency['median']:.0f} ms de latencia media del backend; {percent(latency['under_100'])} de las APIs "
                 "bajo 100 ms" if latency else "Validar con cada equipo"])
    high = sum((bia.get(level) or {}).get("clusters") or 0 for level in ("MUY ALTA", "ALTA"))
    rows.append([str(len(rows)), "Si el servicio es crítico, lo revisa la unidad de gobierno de arquitectura de soluciones",
                 f"{number(high)} clusters con BIA alta o muy alta" if high else "-"])
    table(slide, Inches(0.45), Inches(1.3), Inches(12.45), rows, [0.05, 0.5, 0.45], row_height=0.5, size=11,
          align_first=False, highlight_column=2)
    callout(slide, Inches(0.45), Inches(5.95), Inches(12.45), Inches(0.9),
            "Un cluster pasa a Container Apps solo si cumple todas las reglas; el costo (siguiente lámina) decide el orden.",
            size=13)
    text_box(slide, Inches(0.45), Inches(5.5), Inches(12.45), Inches(0.35),
             "Detalle de cada criterio, qué favorece a cada plataforma y sus fuentes: hoja Technical_criteria del workbook.",
             size=9, italic=True)


def rule_reading(index, value):
    """Spanish reading of the ADT_019 rule at `index` in Break_even, from the model result (same logic as the sheet)."""
    if index == 0:
        if not is_number(value):
            return "ACA más barato en todo el rango"
        if value > 15:
            return f"El costo respalda el umbral: ACA sigue siendo más barato hasta {number(value - 1)}"
        return f"ACA cuesta más que AKS desde {number(value)} microservicios, antes de 15"
    if index == 1:
        if not is_number(value):
            return "-"
        hours = f"{value:.1f}".replace(".", ",")
        if value > 24:
            return "Con n_ref, ACA es más barato aunque la réplica esté caliente todo el día"
        if value > 10:
            return f"El umbral de 10 h es conservador: el cruce está en {hours} h"
        return f"ACA cuesta más que AKS antes de 10 h: {hours} h"
    if index == 2:
        if is_number(value):
            return f"Con réplicas calientes 24/7, ACA supera a AKS lista desde {number(value)} microservicios"
        return f"Aun encendido 24/7, ACA es más barato hasta {str(value or '').lstrip('>')} microservicios"
    if index == 3 and is_number(value):
        return ("Con 15 microservicios y 100 tps en oficina, ACA aún es más barato" if value < 0
                else "Con 15 microservicios y 100 tps, AKS es más barato")
    return "-"


def slide_cost(deck, s, model, equilibrium, rules, assumptions, source):
    crossing = equilibrium.get("Microservices at which ACA (ADT schedule) exceeds AKS effective")
    if crossing in (">500", None, ""):
        title = "Reglas del ADT_019: ACA cuesta menos que un cluster AKS dedicado"
    else:
        title = f"Reglas del ADT_019: ACA cuesta menos que AKS hasta {crossing} microservicios"
    slide = base_slide(deck, title, source, "Criterios técnicos")
    heading(slide, Inches(0.45), Inches(1.2), Inches(7.6), "Costo mensual por aplicación según microservicios (estimado)")
    points = [row for row in model if row[0] <= 30]
    line_chart(slide, Inches(0.3), Inches(1.5), Inches(7.7), Inches(4.4), [str(row[0]) for row in points], [
        ("ACA horario ADT", [row[1] for row in points], ORANGE), ("ACA 24/7", [row[2] for row in points], RUST),
        ("AKS lista", [row[3] for row in points], NAVY), ("AKS con descuento real", [row[4] for row in points], STEEL)])
    tiles = [(money(s.get(SUMMARY_LABELS["aks_min_list"])), "cluster AKS mínimo a precio de lista", SKY),
             (money(s.get(SUMMARY_LABELS["aks_min_effective"])), "cluster AKS mínimo con descuento real", SKY),
             (money(assumptions.get("aca_env_month")), "ACA: costo fijo del ambiente al mes", ORANGE),
             (money(s.get(SUMMARY_LABELS["aca_per_ms"])), "ACA: por microservicio al mes (8–20)", ORANGE)]
    for index, (value, label, color) in enumerate(tiles):
        row, column = divmod(index, 2)
        kpi(slide, Inches(8.3 + column * 2.35), Inches(1.25 + row * 1.2), Inches(2.2), Inches(1.1), value, label,
            value_color=color, value_size=20, label_size=9)
    if rules:
        heading(slide, Inches(8.3), Inches(3.75), Inches(4.6), "Umbrales del ADT_019 frente al modelo", size=11)
        rows = [["Regla del ADT", "Lectura del modelo"]] + [
            [ADT_RULES[index] if index < len(ADT_RULES) else str(rule[0]), rule_reading(index, rule[2])]
            for index, rule in enumerate(rules)]
        table(slide, Inches(8.3), Inches(4.05), Inches(4.6), rows, [0.45, 0.55], row_height=0.4, size=8,
              align_first=False)
    callout(slide, Inches(0.45), Inches(6.15), Inches(12.45), Inches(0.75),
            "El modelo supone 1 réplica de 8 a 20 por microservicio. El resultado real depende de las réplicas que "
            "cada aplicación necesita: ver la comparación por cluster en la propuesta.", size=11, fill=ROYAL)


def strategic_table(deck, title, header, rows, widths, message, source):
    slide = base_slide(deck, title, source, "Estratégico")
    table(slide, Inches(0.35), Inches(1.3), Inches(12.6), [header] + rows, widths, row_height=0.5, size=9.5,
          align_first=False)
    callout(slide, Inches(0.35), Inches(6.05), Inches(12.6), Inches(0.85), message, size=12)
    return slide


def slide_application(deck, source):
    strategic_table(deck, "La aplicación es la misma; cambian configuración, releases y GitOps",
                    ["Tema", "AKS", "Container Apps", "Esfuerzo AKS → ACA"], [
                        ["Imagen y pipeline", "Misma imagen OCI y pipeline hasta ACR", "Igual", "Ninguno"],
                        ["Exposición", "Service + Ingress detrás de APIM", "Ingress interno de la app detrás de APIM",
                         "Cambia la URL de backend en APIM"],
                        ["Productores y consumidores", "Pods siempre arriba o KEDA", "Reglas KEDA; escala a 0 fuera de horario",
                         "Tolerar arranque en frío y lag a las 8:00"],
                        ["Arranque en frío", "No (pods 24/7)", "Hasta ~30 s tras escalar a 0 (ADT_019)", "Reintentos en quien llama"],
                        ["Escalado", "HPA/KEDA sobre nodos que dimensiona el equipo", "Reglas KEDA (HTTP, cola, cron)",
                         "Reescribir el HPA como reglas"],
                        ["Configuración", "ConfigMaps, Secrets, CSI de Key Vault", "Variables en el YAML; secretos vía Key Vault",
                         "Cada ConfigMap pasa a variables o Key Vault"],
                        ["Releases", "Canary por peso o header, rollback automático (Istio, Argo Rollouts)",
                         "Canary solo por peso entre revisiones", "Sin canary por header ni rollback automático"],
                        ["GitOps", "Argo CD / Flux reconcilian desde Git", "Sin reconciliador: el pipeline aplica YAML o Bicep",
                         "El drift no se corrige solo"]],
                    [0.18, 0.27, 0.28, 0.27],
                    "Para el negocio la aplicación funciona igual; el trabajo está en configuración, releases y GitOps.",
                    source)


def slide_security(deck, source):
    strategic_table(deck, "Seguridad: Istio da identidad y autorización por servicio; ACA no",
                    ["Capacidad", "AKS + add-on de Istio", "Container Apps"], [
                        ["mTLS entre servicios", "PeerAuthentication STRICT por namespace o workload",
                         "Cifrado peer-to-peer: un ajuste para todo el ambiente, apagado por defecto; autentica pero no autoriza"],
                        ["Identidad en el certificado", "SPIFFE ID por service account", "No expuesta"],
                        ["Certificados", "CA de istiod o CA propia desde Key Vault; rotación por la malla",
                         "Los gestiona la plataforma; sin control de la CA"],
                        ["Autorización entre servicios", "AuthorizationPolicy por identidad, path y método",
                         "No a ese nivel: políticas de APIM y reglas de red"],
                        ["Control de salida", "Egress gateway, ServiceEntry", "Red del ambiente (UDR, firewall)"]],
                    [0.22, 0.36, 0.42],
                    "En ACA la autorización entre servicios recae en APIM y la red; Istio la resuelve por identidad.",
                    source)


def slide_resilience(deck, source):
    strategic_table(deck, "Resiliencia: Istio lo cubre por ruta; ACA lo tiene en preview o en APIM",
                    ["Capacidad", "AKS + Istio", "Container Apps"], [
                        ["Balanceo", "Round robin, least request, hash consistente", "Balanceo de la plataforma"],
                        ["Timeout y retry", "VirtualService por ruta", "Política de resiliencia (preview, una por app) o APIM"],
                        ["Circuit breaker interno", "DestinationRule (connectionPool + outlierDetection)",
                         "Circuit breaker de la política de resiliencia (preview)"],
                        ["Circuit breaker externo", "ServiceEntry + DestinationRule", "Circuit breaker de backend en APIM, o en código"],
                        ["Outlier detection", "Expulsa endpoints que fallan", "Cubierto por el mismo circuit breaker (preview)"],
                        ["Inyección de fallas", "Sí", "No"]],
                    [0.22, 0.36, 0.42],
                    "Como todo pasa por APIM, timeout, retry y circuit breaker pueden unificarse en APIM en ambas "
                    "plataformas; la malla agrega el salto interno.", source)


def slide_multicloud(deck, source):
    slide = base_slide(deck, "Para el negocio son equivalentes; para TI, doble esfuerzo", source, "Estratégico")
    columns = [("AKS (Kubernetes estándar)", STEEL,
                ["Los mismos manifiestos, Helm, Istio y Argo CD corren en EKS o GKE.",
                 "El modelo de operación y los equipos se llevan a otra nube.",
                 "El costo es mantener clusters: versiones, nodos, malla."]),
               ("Container Apps (API propia de Azure)", ORANGE,
                ["Microsoft.App no existe en otras nubes.",
                 "Llevar una app a otra nube es pasarla antes a Kubernetes.",
                 "Menos operación hoy; migración más larga si cambia la estrategia."])]
    for index, (name, color, lines) in enumerate(columns):
        left = Inches(0.45 + index * 6.3)
        rounded(slide, left, Inches(1.35), Inches(6.1), Inches(2.6), CARD)
        text_box(slide, left + Inches(0.3), Inches(1.5), Inches(5.5), Inches(0.5), name, size=17, bold=True, color=color)
        text_box(slide, left + Inches(0.3), Inches(2.15), Inches(5.5), Inches(1.7), [f"• {line}" for line in lines],
                 size=14, color=WHITE)
    heading(slide, Inches(0.45), Inches(4.15), Inches(12), "Migración ACA → AKS (si una app crece o cambia la estrategia)",
            size=12)
    text_box(slide, Inches(0.45), Inches(4.5), Inches(12.4), Inches(0.7),
             "Misma imagen. Se reescribe el YAML de la app como Deployment y Service, las reglas de escala como HPA o "
             "KEDA, y las variables como ConfigMaps. Esfuerzo por aplicación y ambiente.", size=11)
    note_box(slide, Inches(0.45), Inches(5.25), Inches(12.45), Inches(0.6),
             "Por completar con el equipo de plataforma: modelo de delivery y esfuerzo de migración en horas-hombre.",
             size=10)
    callout(slide, Inches(0.45), Inches(6.0), Inches(12.45), Inches(0.9),
            "Dos plataformas = dos formas de configurar, liberar, asegurar y dar resiliencia, y dos caminos de "
            "migración multicloud.", size=13)


def slide_per_cluster(deck, rows, source):
    with_data = [row for row in rows if row["cheaper"] in ("ACA", "AKS")]
    levels = [("cheaper", "AKS hoy", NAVY), ("cheaper_rs", "AKS con right sizing", STEEL),
              ("cheaper_hpa", "ambos con HPA ajustado", ORANGE)]
    levels = [level for level in levels if any(row[level[0]] in ("ACA", "AKS") for row in with_data)]
    wins = {key: sum(row[key] == "ACA" for row in with_data) for key, _, _ in levels}
    last = levels[-1][0]
    slide = base_slide(deck, f"ACA gana en los clusters chicos aun con AKS ajustado: {number(wins[last])} de "
                             f"{number(len(with_data))}", source, "Propuesta")
    buckets = [("1 a 5", 1, 5), ("6 a 15", 6, 15), ("16 a 30", 16, 30), ("Más de 30", 31, 10 ** 9)]
    heading(slide, Inches(0.45), Inches(1.2), Inches(6.5), "Clusters donde ACA cuesta menos, por workloads", size=12)
    data = CategoryChartData()
    data.categories = [label for label, _, _ in buckets]
    groups = [[row for row in with_data if is_number(row["workloads"]) and low <= row["workloads"] <= high]
              for _, low, high in buckets]
    for key, name, _ in levels:
        data.add_series(f"vs {name}", [sum(row[key] == "ACA" for row in group) for group in groups])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.3), Inches(1.5), Inches(6.5), Inches(3.6),
                                   data).chart
    chart.font.name, chart.font.size = FONT, Pt(10)
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    plot.data_labels.font.bold = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    for series, (_, _, color) in zip(plot.series, levels):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = color

    def total(key, group):
        return sum(row[key] for row in group if is_number(row[key]))

    heading(slide, Inches(7.1), Inches(1.2), Inches(5.8), "Costo mensual por tamaño del cluster", size=12)
    table_rows = [["Workloads", "Clusters", "AKS hoy", "AKS ajustado", "+ HPA", "ACA (HPA aj.)"]]
    for (label, _, _), group in zip(buckets, groups):
        table_rows.append([label, number(len(group)), money(total("aks", group)), money(total("aks_rs", group)),
                           money(total("aks_hpa", group)), money(total("aca_hpa", group) or total("aca_use", group))])
    table(slide, Inches(7.1), Inches(1.5), Inches(5.8), table_rows, [0.17, 0.13, 0.17, 0.19, 0.16, 0.18],
          row_height=0.42, size=9)
    tiles = [(money(total("aks", with_data)), "AKS hoy al mes", SKY),
             (money(total("aks_rs", with_data)), "AKS con right sizing de requests y nodos", ORANGE),
             (money(total("aks_hpa", with_data)), "y además mínimos de HPA ajustados", GREEN)]
    for index, (value, label, color) in enumerate(tiles):
        kpi(slide, Inches(7.1 + index * 1.97), Inches(3.75), Inches(1.85), Inches(1.3), value, label, value_color=color,
            value_size=17, label_size=8)
    text_box(slide, Inches(0.45), Inches(5.25), Inches(12.4), Inches(0.75), [
        "Right sizing: requests al uso + 50%, en nodos del tamaño actual, con piso de 4 nodos (2 sistema + 2 user); VMs y "
        "discos escalan con los nodos. HPA ajustado: los que se quedan en su mínimo bajan a lo que pide su CPU (mínimo 2), "
        "en AKS y en ACA. ACA: réplicas medidas, tamaño por uso + 50%, 24/7, precio de lista. Estimado: verificar por cluster."],
        size=9, italic=True)
    callout(slide, Inches(0.45), Inches(6.05), Inches(12.45), Inches(0.82),
            "El right sizing baja AKS sobre todo en los clusters grandes; en los chicos el piso de 4 nodos pesa más y "
            "ACA sigue ganando.", size=12)


def slide_traffic(deck, classes, settings, source):
    total_apis = sum(item[1] or 0 for item in classes)
    office = next((item for item in classes if item[0] == "OFFICE"), ("OFFICE", 0, 0))
    slide = base_slide(deck, f"Solo {percent(share_of(office[1], total_apis))} de las APIs es de oficina; la mayoría "
                             "es extendido o 24/7", source, "Propuesta")
    heading(slide, Inches(0.45), Inches(1.25), Inches(6.5), "APIs por perfil de tráfico")
    colors = {"OFFICE": GREEN, "EXTENDED": ORANGE, "24x7": RUST, "NO_TRAFFIC": SOFT}
    data = CategoryChartData()
    data.categories = [TRAFFIC_CLASSES.get(item[0], item[0]) for item in classes]
    data.add_series("APIs", [item[1] or 0 for item in classes])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.4), Inches(1.6), Inches(3.8), Inches(3.8), data).chart
    chart.has_title, chart.has_legend = False, False
    for index, item in enumerate(classes):
        point = chart.plots[0].series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors.get(item[0], STEEL)
    text_box(slide, Inches(0.4), Inches(3.1), Inches(3.8), Inches(0.8), number(total_apis), size=26, bold=True,
             color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for index, item in enumerate(classes):
        legend_item(slide, Inches(4.4), Inches(2.1 + index * 0.6), colors.get(item[0], STEEL),
                    f"{TRAFFIC_CLASSES.get(item[0], item[0])}: {number(item[1])} APIs", width=Inches(2.5))
    heading(slide, Inches(7.1), Inches(1.25), Inches(5.8), "Cómo se clasifica (celdas editables)")
    office_window = f"{settings.get('Office: start hour (Monday to Friday)', 8)}–{settings.get('Office: end hour', 20)}"
    extended_window = f"{settings.get('Extended: start hour (every day)', 6)}–{settings.get('Extended: end hour', 22)}"
    threshold = percent(settings.get("Classification threshold", 0.9))
    rows = [["Clase", "Regla", "APIs", "% peticiones"],
            ["OFFICE", f"≥ {threshold} lun–vie {office_window}"],
            ["EXTENDED", f"≥ {threshold} entre {extended_window}, todos los días"],
            ["24x7", "Repartido en el día y la noche"], ["NO_TRAFFIC", "Sin peticiones"]]
    by_name = {item[0]: item for item in classes}
    for row in rows[1:]:
        item = by_name.get(row[0], (row[0], None, None))
        row += [number(item[1]), percent(item[2])]
        row[0] = TRAFFIC_CLASSES[row[0]]
    table(slide, Inches(7.1), Inches(1.6), Inches(5.8), rows, [0.22, 0.46, 0.14, 0.18], row_height=0.5, size=9)
    note_box(slide, Inches(7.1), Inches(4.35), Inches(5.8), Inches(1.0),
             "Siguiente paso: mapear cada API a su deployment para clasificar aplicaciones, no APIs, y aplicar el "
             "modelo de ACA según su perfil real.", size=10)
    callout(slide, Inches(0.45), Inches(5.95), Inches(12.45), Inches(0.9),
            "Propuesta: decidir por perfil de tráfico. Oficina → ACA con escala a 0; extendido → ACA con horario "
            "ampliado; 24/7 con carga alta → AKS.", size=12)


def slide_quick_wins(deck, p, measured_summary, clusters, per_cluster, signals, source, summary=None):
    slide = base_slide(deck, "Quick wins y siguientes pasos", source, "Propuesta")
    summary = summary or {}
    empty_count = summary.get(SUMMARY_LABELS["no_applications"])
    empty_cost = summary.get(SUMMARY_LABELS["no_applications_cost"])
    small = [item for item in p["clusters"] if item[0] in ("1", "2 to 3")]
    aca = [row for row in per_cluster if row["cheaper"] == "ACA" and is_number(row["workloads"]) and row["workloads"] <= 15]
    oom = signals.get("oom")

    def total(key):
        return sum(row[key] for row in per_cluster if is_number(row.get(key)) and is_number(row.get("aks_rs")))

    saving_rs, saving_hpa = total("aks") - total("aks_rs"), total("aks_rs") - total("aks_hpa")
    nodes_now, nodes_rs = total("nodes"), total("nodes_rs")
    rows = [["Acción", "Alcance", "Efecto", "Dónde se ve"],
            ["Right sizing: requests al uso + 50% y menos nodos",
             f"{money(saving_rs)} al mes ({number(nodes_now)} → {number(nodes_rs)} nodos)" if saving_rs else "-",
             "Ahorro sin migrar", "ACA_by_cluster"],
            ["Bajar mínimos de HPA inflados",
             f"{number(p.get('inflated_beyond_ha_count'))} HPA con mínimo sobre lo que pide su CPU, aun dejando 2"
             + (f"; {money(saving_hpa)} más al mes" if saving_hpa else ""), "Ahorro sin migrar", "Patrón 4"],
            ["Revisar clusters sin aplicaciones", f"{number(empty_count)} clusters, {money(empty_cost)} al mes",
             "Ahorro", "Clusters"],
            ["Consolidar apps de 1 a 3 microservicios", f"{number(sum(item[1] for item in small))} clusters, "
             f"{money(p.get('cost_le3'))} al mes", "Ahorro (por evaluar)", "Patrón 2"],
            ["Mover a ACA los clusters chicos donde gana", f"{number(len(aca))} clusters con ≤15 workloads, "
             f"{money(sum(row['difference'] for row in aca if is_number(row['difference'])))} al mes", "Ahorro con ADT_019",
             "ACA_by_cluster"],
            ["Llevar a 4 nodos los clusters bajo el mínimo",
             f"{number(measured_summary.get(MEASURED_LABELS['lt4_min']))} clusters",
             "Resiliencia (sube el costo)", "Measured_AKS"],
            ["Ajustar memoria de workloads con OOMKilled",
             f"{number(oom['workloads'])} workloads" if oom else "Correr aa_signals.py", "Estabilidad", "aa-signals"],
            ["Mapear API → aplicación", "Para clasificar apps por perfil de tráfico", "Habilita la decisión por app",
             "Traffic_profile"]]
    table(slide, Inches(0.35), Inches(1.3), Inches(12.6), rows, [0.27, 0.38, 0.18, 0.17], row_height=0.5, size=10,
          align_first=False)
    callout(slide, Inches(0.35), Inches(6.0), Inches(12.6), Inches(0.9),
            "Decisiones que pedimos: aprobar el ajuste de mínimos y la revisión de clusters vacíos, y un piloto de ACA "
            "con 2 o 3 aplicaciones chicas de perfil oficina.", size=12)


def slide_bia(deck, bia, source):
    slide = base_slide(deck, "Anexo BIA: la criticidad prioriza, no decide la plataforma", source, "Anexo")
    heading(slide, Inches(0.45), Inches(1.25), Inches(6.8), "Costo mensual de AKS PROD por criticidad de negocio")
    levels = [level for level in BIA_ORDER if level in bia]
    total = bia.get("Total", {}).get("cost") or sum(bia[level]["cost"] or 0 for level in levels)
    chart_levels = list(reversed(levels))
    bar_chart(slide, Inches(0.3), Inches(1.6), Inches(7.0), Inches(4.3),
              [level if level != "SIN BIA" else "SIN BIA (pendiente)" for level in chart_levels],
              [bia[level]["cost"] or 0 for level in chart_levels], XL_CHART_TYPE.BAR_CLUSTERED,
              [SKY if level == "SIN BIA" else ORANGE for level in chart_levels], number_format='"US$"#,##0')
    text_box(slide, Inches(0.45), Inches(6.0), Inches(6.8), Inches(0.4),
             "Método: cada cluster toma el BIA más alto entre sus microservicios.", size=9, italic=True)
    top = bia.get("MUY ALTA", {})
    heading(slide, Inches(7.6), Inches(1.25), Inches(5.3), "Lectura")
    kpi(slide, Inches(7.6), Inches(1.6), Inches(5.3), Inches(1.25), money(top.get("cost")),
        f"costo mensual en BIA MUY ALTA · {percent(share_of(top.get('cost'), total))} del costo total", value_size=26,
        label_size=10)
    high = sum(bia.get(level, {}).get("clusters") or 0 for level in ("MUY ALTA", "ALTA"))
    tiles = [(number(high), "clusters BIA alta o muy alta", RED),
             (number(bia.get("SIN BIA", {}).get("clusters")), "clusters sin BIA en el export", SKY),
             (number(bia.get("Total", {}).get("clusters")), "clusters PROD", GREEN)]
    for index, (value, label, color) in enumerate(tiles):
        kpi(slide, Inches(7.6 + index * 1.8), Inches(3.0), Inches(1.7), Inches(1.3), value, label, value_color=color,
            value_size=22, label_size=9)
    callout(slide, Inches(7.6), Inches(4.5), Inches(2.6), Inches(0.9), ["1 PRIORIZAR", "Gobierno y secuencia"], size=10,
            fill=RUST)
    callout(slide, Inches(10.3), Inches(4.5), Inches(2.6), Inches(0.9), ["2 DECIDIR TÉCNICAMENTE", "Carga, réplica, operación"],
            size=10, fill=ROYAL)
    callout(slide, Inches(7.6), Inches(5.55), Inches(5.3), Inches(0.75),
            "El BIA orienta la prioridad; no determina por sí solo AKS o Container Apps.", size=11)


def slide_sources(deck, sources, source):
    slide = base_slide(deck, "Anexo: de dónde sale cada cifra", source, "Anexo")
    rows = [["Dato", "Tipo", "Fuente", "Archivo"]] + [[SOURCE_TEXT.get(str(row[0] or ""), str(row[0] or "")),
                                                       SOURCE_TEXT.get(str(row[1] or ""), str(row[1] or "")),
                                                       SOURCE_TEXT.get(str(row[2] or ""), str(row[2] or "")),
                                                       str(row[4] or "")] for row in sources[:11]]
    rows.append(["Pods esperando nodo, OOMKilled, latencia de APIM", "Medido", "Azure Monitor; Container Insights",
                 "aks_autoscaler_monthly.csv, aks_oomkilled_monthly.csv, apim_latency_monthly.csv"])
    table(slide, Inches(0.35), Inches(1.25), Inches(12.6), rows, [0.32, 0.14, 0.24, 0.30], row_height=0.36, size=8.5,
          align_first=False)
    note_box(slide, Inches(0.35), Inches(6.4), Inches(12.6), Inches(0.55),
             "Medido = leído de Azure. Derivado = calculado de lo medido. Estimado = modelo con precios de lista y "
             "supuestos editables (hoja Assumptions). Consultas exactas en la hoja Queries.", size=9)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--workbook", default=None,
                        help="aks_aca_comparison.xlsx, opened and saved in Excel. Default: the newest one in reports/")
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None, help="Output .pptx. Default: aks_aca_deck.pptx next to the workbook")
    args = parser.parse_args()
    path = args.workbook or latest_workbook()
    book = load_workbook(path, data_only=True)
    summary, patterns = read_summary(book), read_patterns(book)
    if summary.get(SUMMARY_LABELS["cost_total"]) in (None, ""):
        raise SystemExit("The workbook has no calculated values. Open it in Excel, save it once and run this again.")
    if not patterns["layers"]:
        raise SystemExit("The Patterns sheet is empty. Run aa_k8s_snapshot.py, then aa_workbook.py, and save in Excel.")
    measured_summary, clusters = read_measured(book)
    criteria, _ = read_criteria(book)
    equilibrium, rules = read_equilibrium(book)
    per_cluster, _ = read_per_cluster(book)
    classes, settings = read_profile(book)
    signals = read_signals(os.path.join(args.results, "aa-signals"), application_clusters(args.results))
    period = spanish_period(summary.get(SUMMARY_LABELS["months"]))

    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    taken_on = patterns["snapshot_date"]
    snapshot = f"; HPA y deployments al {taken_on}" if taken_on else ""
    measured_source = ("Fuente: Azure Cost Management (costo amortizado), Azure Monitor y Container Insights "
                       f"({period}){snapshot}. Solo clusters con aplicaciones.")
    model_source = ("Estimado: precios de lista (Azure Retail Prices), descuento real de nodos (Cost Management) y "
                    "supuestos de la hoja Assumptions; reglas del ADT_019.")
    strategic_source = "Fuente: Microsoft Learn (revisado nov-2026) y ADT_019."

    slide_cover(deck, taken_on, period)
    slide_summary(deck, summary, patterns, measured_summary, clusters, measured_source)
    slide_pattern1(deck, patterns, measured_source)
    slide_pattern2(deck, patterns, measured_source)
    slide_pattern4(deck, patterns, measured_source, period)
    slide_capacity_health(deck, summary, signals, measured_source.replace("Container Insights", "Container Insights "
                                                                           "(KubePodInventory), métricas del autoscaler", 1),
                          period)
    slide_footprint(deck, summary, read_footprint(args.results), measured_source)
    slide_configuration(deck, measured_summary, read_config(book), criteria,
                        "Fuente: Azure Resource Graph (configuración de clusters y node pools).")
    slide_criteria(deck, criteria, signals, read_bia(book), measured_source)
    slide_cost(deck, summary, read_model(book), equilibrium, rules, read_assumptions(book), model_source)
    slide_application(deck, strategic_source)
    slide_security(deck, strategic_source)
    slide_resilience(deck, strategic_source)
    slide_multicloud(deck, strategic_source)
    cost_month = spanish_period(max(re.findall(r"\d{4}-\d{2}", str(summary.get(SUMMARY_LABELS["months"]) or "")),
                                    default=""))
    slide_per_cluster(deck, per_cluster, f"Costo AKS real (Cost Management, {cost_month}); ACA estimado con "
                                         "réplicas medidas por Container Insights.")
    slide_traffic(deck, classes, settings, "Fuente: Azure Monitor, métrica Requests de APIM por API y hora.")
    slide_quick_wins(deck, patterns, measured_summary, clusters, per_cluster, signals, measured_source, summary)
    slide_bia(deck, read_bia(book), measured_source)
    slide_sources(deck, read_sources(book), "Hojas Sources y Queries del workbook.")
    output = args.output or os.path.join(os.path.dirname(path), "aks_aca_deck.pptx")
    deck.save(output)
    print(output)


if __name__ == "__main__":
    main()
