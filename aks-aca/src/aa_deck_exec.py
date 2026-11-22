#!/usr/bin/env python3
"""Build the corporate AKS vs Container Apps deck (slides 1 to 6) for a non-technical forum.

Same look as aa_deck.py (its drawing helpers are reused), plain figures only: TPS, vCPU, GB of RAM, instances,
hours with traffic, money per month and per year. Slide text is in Spanish for the audience. Every amount is in
soles with IGV (18 %) at S/ 3.41333 per US$ (soles()), as the forum reads the bank's numbers.

Reads only results/ (inventory, metrics, cost, prices, snapshot), so it needs no saved workbook. The cost slide
uses a simplified ADT_019 model (see aca_estimate); the workbook's ACA_by_cluster sheet has the detail per
cluster. No Azure call is made; nothing leaves the machine.

Slides (the story: what we build, where it runs, why we challenge it, the rules, the result, the money):
  1. Ways to build an application: modular monolith, miniservices, microservices (and asynchronous processes when
     SHOW_FUNCTIONS), each drawn and with its usual runtime (FAT_001 and ADT_019), and a legend of the acronyms
  2. AKS in PROD: clusters, applications, "micro"-services, monthly and annual cost (PROD plus non-PROD), the
     clusters by category (only the clusters with applications enter the analysis) and non-PROD in small print
  3. Hypothesis: three signs, each with a sentence the presenter can say, the motivation (economic first, then
     operation and complexity) and the smallest AKS cluster against one Container Apps "micro"-service
  4. The two conditions to move to Container Apps (peak hour traffic, size) and the off-hours rule (it may scale to
     zero, but the cost counts one replica on 24 h), plus what is validated before moving (critical, growing)
  5. Result: where each application goes, the conditions each group meets, its AKS cost today, and how many
     applications stay on AKS and move to Container Apps (the funnel slide was dropped on 2026-11-22;
     aa_criteria_workbook.py keeps the funnel and an editable chart)
  6. Possible saving: monthly and annual, with and without discount
ADT_019 is the reference, not the rule (asked on 2026-11-22): the conditions are the ones on slide 4 (criteria()),
and every Container Apps estimate keeps one replica per "micro"-service on outside office hours.
aa_criteria_workbook.py writes the same conditions and costs per application as an Excel workbook built on Power
Query, where the off-hours replicas of each application can be changed.
A last slide (effort and migration estimate) is the architecture team's and is not built here.

Runtime counts only take resources that existed in the measured month (existed_at_measurement in the inventory):
run the scripts after the month ends and the deck shows that month, whatever day it is built.

Output: aks_aca_corporate.pptx in the newest reports/<date>-aks-aca folder (or --output), and next to it
aca_estimate_by_app.csv: the Container Apps cost of each application with every input and step (aca_breakdown). Every figure is a draft to check against
the workbook and the CSVs before presenting.
"""

import argparse
import glob
import json
import math
import os
from collections import Counter, defaultdict
from datetime import date

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches

import aa_common as aa
import aa_deck as d

MONTHS_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
             "noviembre", "diciembre")
COUNT_WORDS = {3: "Tres", 4: "Cuatro", 5: "Cinco", 6: "Seis", 7: "Siete"}
# Every amount on the slides: soles with IGV, at the exchange rate the forum uses.
EXCHANGE_RATE, IGV = 3.41333, 0.18
MONEY_NOTE = "Montos en soles con IGV (18 %); tipo de cambio S/ 3,41333 por US$."
# Smallest PROD AKS cluster (architecture team: 2 system + 2 user nodes of 8 vCPU / 32 GB) and the Container Apps
# "micro"-service of the model (0.5 vCPU / 1 GiB, one replica on 24 h, 5 million requests a month).
MIN_AKS_NODES, MIN_AKS_VM, HOURS_MONTH = 4, "Standard_D8s_v4", 730
# The bank's discount on Azure list prices: for the minimum comparison and for Container Apps, which has no
# reservations (the measured node discount includes them).
BANK_DISCOUNT = 0.30
# Azure Functions and asynchronous processes are left out of slides 1 and 2 for now (asked on 2026-11-19); True
# shows them again.
SHOW_FUNCTIONS = False
# Official Azure icons (Azure Architecture Icons, as packaged by draw.io), rasterised to PNG.
ICONS_DIR = os.path.join(aa.TOPIC_DIR, "assets", "icons")
ACA_VCPU, ACA_GIB, ACA_REQUESTS_MILLION = 0.5, 1.0, 5
# The bank's deployed services are written "micro"-services on the slides (asked on 2026-11-22): many are not small.
MICRO, MICRO_ONE, MICRO_CAP = "“micro”-servicios", "“micro”-servicio", "“Micro”-servicios"
ACRONYMS = "AKS: Azure Kubernetes Service  ·  ACA: Azure Container Apps"
YELLOW = d.RGBColor(0xE3, 0xC2, 0x1B)
# The conditions to move to Container Apps (slide 4) and the off-hours replicas of the cost model. ADT_019 is the
# reference: no standard says when an application may scale to zero, so every estimate keeps OFFHOURS_REPLICAS per
# "micro"-service on outside office hours (Microsoft: minReplicas >= 1 for production that must always answer), billed
# at the active rate. aa_criteria_workbook.py lets each application change it.
MAX_TPS, MAX_MICROSERVICES, OFFHOURS_REPLICAS = 100, 15, 1
CRITICAL_BIA = ("MUY ALTA", "ALTA")
# The two minimums on slide 3 are computed from the list prices in prices.csv (Azure Retail Prices API, the same
# prices as the Azure pricing calculator), set up as one would enter them in the calculator (East US 2):
#   AKS: Standard tier, 4 x D8s v4 Linux, 730 h, 3-year reservation, one P10 managed OS disk per node.
#   Container Apps: Consumption, 1 replica of 0.5 vCPU / 1 GiB on for 730 h, 5 M requests (the calculator shows it
#   minus the monthly free grant, which belongs to the subscription, not to one application).
# On 2026-11-19 the calculator gave AKS 426.11 + 71.68 + 73.00 = 570.79 US$; Container Apps for 730 h is
# 31.54 + 7.88 + 2.00 = 41.42 US$.
MIN_AKS_DISK, MIN_AKS_RESERVATION = "P10", "3y"


def soles(usd):
    """US$ amount -> 'S/ 1.234' in soles with IGV."""
    return f"S/ {usd * EXCHANGE_RATE * (1 + IGV):,.0f}".replace(",", ".") if d.is_number(usd) else "-"


def month_label(month):
    """'2026-10' -> 'octubre 2026'."""
    try:
        year, number_ = month.split("-")
        return f"{MONTHS_ES[int(number_) - 1]} {year}"
    except (ValueError, IndexError):
        return month or "el mes"


def deployed_microservices(apps):
    """"Micro"-services running in the clusters: each cluster counts its own, so one deployed in two clusters of an
    application counts twice (microservices_deployed). The 15 rule uses microservices: distinct per application."""
    return sum(int(to_float(row.get("microservices_deployed") or row.get("microservices")) or 0) for row in apps)


def results_csv(results, folder, name):
    return aa.read_csv(os.path.join(results, folder, name))


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------


def existed(row):
    """The resource existed in the measured month (inventory column existed_at_measurement; kept when unknown)."""
    return row.get("existed_at_measurement", "") != "NO"


def prod_inventory(results, name):
    """PROD rows of an inventory CSV that existed in the measured month: the inventory is current, the figures are
    of a past month."""
    return [row for row in results_csv(results, "aa-inventory", name) if row["environment"] == "PROD" and existed(row)]


def read_estate(results):
    """Figures of the runtime estate in PROD, by cluster category, for the measured month."""
    in_month = {row["id"].lower() for row in prod_inventory(results, "aks_clusters.csv")}
    clusters = [row for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["in_inventory"] == "YES" and row["cluster_id"] in in_month]
    cost_rows = results_csv(results, "aa-cost", "cost_by_resource.csv")
    month = max((row["month"] for row in cost_rows if row["owner_type"] == "AKS"), default="")
    cost = defaultdict(float)
    for row in cost_rows:
        if row["owner_type"] == "AKS" and row["month"] == month:
            cost[row["owner_id"].lower()] += to_float(row["cost_usd"]) or 0
    by_category = defaultdict(lambda: {"clusters": 0, "cost": 0.0, "microservices": 0})
    for row in clusters:
        entry = by_category[row["category"]]
        entry["clusters"] += 1
        entry["cost"] += cost.get(row["cluster_id"], 0)
        entry["microservices"] += int(row["microservices"] or 0)
    application_ids = {row["cluster_id"] for row in clusters if row["category"] == aa.APPLICATION}
    ids_by_name = aa.cluster_ids_by_name(results_csv(results, "aa-inventory", "aks_clusters.csv"))
    usage = [row for row in results_csv(results, "aa-metrics", "aks_usage_monthly.csv")
             if row["month"] == month and aa.cluster_id_of(row, ids_by_name) in application_ids]
    cpu_used = sum(to_float(row["cpu_used_office_avg"]) or 0 for row in usage if row["mode"] != "ALLOCATABLE")
    mem_used = sum(to_float(row["mem_used_office_avg_gib"]) or 0 for row in usage if row["mode"] != "ALLOCATABLE")
    # Peak: each pool's busiest hour of the month, added up. The hours differ, so the estate's real peak is lower.
    cpu_peak = sum(to_float(row["cpu_used_max"]) or 0 for row in usage if row["mode"] != "ALLOCATABLE")
    mem_peak = sum(to_float(row["mem_used_max_gib"]) or 0 for row in usage if row["mode"] != "ALLOCATABLE")
    cpu_total = sum(to_float(row["cpu_allocatable_avg"]) or 0 for row in usage if row["mode"] == "ALLOCATABLE")
    mem_total = sum(to_float(row["mem_allocatable_avg_gib"]) or 0 for row in usage if row["mode"] == "ALLOCATABLE")
    sites = prod_inventory(results, "app_services.csv")
    plans = prod_inventory(results, "app_service_plans.csv")
    prod = {
        "aks": len(in_month),
        "aca": len(prod_inventory(results, "aca_apps.csv")),
        "aca_environments": len(prod_inventory(results, "aca_environments.csv")),
        "web": sum(row.get("type") == "WEB" for row in sites),
        "functions": sum(row.get("type") == "FUNCTION" for row in sites),
        "plans": Counter(row.get("runtime") or "APP_SERVICE" for row in plans),
    }
    prod_cost = defaultdict(float)
    for row in cost_rows:
        if row["month"] == month and row["environment"] == "PROD":
            prod_cost[row["owner_type"]] += to_float(row["cost_usd"]) or 0
    # Non-PROD AKS (certification, development and clusters without an environment), in small print on slide 2.
    nonprod_clusters = sum(row["environment"] != "PROD" and existed(row)
                           for row in results_csv(results, "aa-inventory", "aks_clusters.csv"))
    nonprod_cost = sum(to_float(row["cost_usd"]) or 0 for row in cost_rows
                       if row["owner_type"] == "AKS" and row["month"] == month and row["environment"] != "PROD")
    summary_path = os.path.join(results, "aa-cost", "cost_summary.json")
    summary = {}
    if os.path.exists(summary_path):
        with open(summary_path, encoding="utf-8") as handle:
            summary = json.load(handle)
    web_scope = summary.get("app_service_environments", [])
    web_measured = web_scope == "ALL" or "PROD" in web_scope
    return {"month": month, "by_category": by_category, "total_clusters": len(clusters),
            "cpu_used": cpu_used, "cpu_total": cpu_total, "mem_used": mem_used, "mem_total": mem_total,
            "cpu_peak": cpu_peak, "mem_peak": mem_peak,
            "prod": prod, "prod_cost": prod_cost, "web_measured": web_measured,
            "nonprod_clusters": nonprod_clusters, "nonprod_cost": nonprod_cost}


def read_scaling(results):
    """Share of microservice HPAs of APPLICATION clusters that do not scale, and time without scaling."""
    application = {row["cluster_id"] for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                   if row["category"] == aa.APPLICATION}
    hpas = [row for row in results_csv(results, "aa-k8s-snapshot", "hpa_snapshot.csv")
            if row["environment"] == "PROD" and row["microservice"] == "YES" and row["cluster_id"] in application]
    known = [row for row in hpas if row["behaviour_ci"] != "NO_DATA"]
    behaviour = Counter(row["behaviour_ci"] for row in known)
    dated = [row for row in hpas if row["last_scale_bucket"] != "NO_DATA"]
    buckets = Counter(row["last_scale_bucket"] for row in dated)
    return {"hpas": len(hpas), "known": len(known), "never": behaviour["NEVER_SCALES"], "fixed": behaviour["FIXED"],
            "dated": len(dated), "over_3_months": buckets["3_TO_12_MONTHS"] + buckets["OVER_1_YEAR"],
            "over_1_year": buckets["OVER_1_YEAR"]}


def read_applications(results):
    return results_csv(results, "aa-k8s-snapshot", "application_snapshot.csv")


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def box(slide, left, top, width, height, fill, radius=0.12):
    shape = d.rounded(slide, left, top, width, height, fill, radius=radius)
    shape.line.fill.background()
    return shape


def shape(slide, kind, left, top, width, height, fill):
    item = slide.shapes.add_shape(kind, left, top, width, height)
    item.fill.solid()
    item.fill.fore_color.rgb = fill
    item.line.fill.background()
    return item


def draw_monolith(slide, left, top, width, accent):
    """One block with modules inside."""
    box(slide, left + width / 2 - Inches(0.85), top, Inches(1.7), Inches(1.15), accent)
    for row in range(2):
        for column in range(2):
            box(slide, left + width / 2 - Inches(0.7) + column * Inches(0.75), top + Inches(0.15) + row * Inches(0.5),
                Inches(0.65), Inches(0.38), d.WHITE, radius=0.2)


def draw_pieces(slide, left, top, width, accent, columns, rows, size):
    """A grid of separate pieces: few and medium (miniservices) or many and small (microservices)."""
    gap = Inches(0.1)
    total = columns * size + (columns - 1) * gap
    for row in range(rows):
        for column in range(columns):
            box(slide, left + (width - total) / 2 + column * (size + gap), top + row * (size + gap), size, size,
                accent, radius=0.2)


def draw_async(slide, left, top, width, accent):
    """Messages waiting in a queue, then the process that takes them."""
    start = left + width / 2 - Inches(1.25)
    box(slide, start, top + Inches(0.3), Inches(1.1), Inches(0.55), d.SOFT, radius=0.15)
    for index in range(4):
        box(slide, start + Inches(0.1) + index * Inches(0.25), top + Inches(0.38), Inches(0.17), Inches(0.39), accent,
            radius=0.3)
    shape(slide, MSO_SHAPE.RIGHT_ARROW, start + Inches(1.2), top + Inches(0.42), Inches(0.4), Inches(0.3), d.GREY)
    box(slide, start + Inches(1.7), top + Inches(0.2), Inches(0.8), Inches(0.75), accent)


def type_card(slide, left, width, title, drawing, lines, runtime, accent):
    top, height = Inches(1.35), Inches(4.35)
    box(slide, left, top, width, height, d.LIGHT)
    d.text_box(slide, left, top + Inches(0.15), width, Inches(0.4), title, size=16, bold=True, color=accent,
               align=d.PP_ALIGN.CENTER)
    drawing(slide, left, top + Inches(0.7), width, accent)
    d.text_box(slide, left + Inches(0.2), top + Inches(2.05), width - Inches(0.4), Inches(1.4), lines, size=12,
               color=d.GREY, align=d.PP_ALIGN.CENTER)
    box(slide, left + Inches(0.2), top + height - Inches(0.8), width - Inches(0.4), Inches(0.6), accent)
    d.text_box(slide, left + Inches(0.2), top + height - Inches(0.8), width - Inches(0.4), Inches(0.6), runtime,
               size=12, bold=True, color=d.WHITE, align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)


def slide_context(deck, apps, source):
    cards = [("Monolito modular", draw_monolith,
              ["Una sola aplicación, ordenada en módulos por negocio.", "Se despliega y escala completa."],
              "App Service", d.SKY),
             ("Miniservicios", lambda sl, le, to, wi, ac: draw_pieces(sl, le, to, wi, ac, 2, 1, Inches(0.95)),
              ["Pocos servicios medianos: cada uno agrupa funciones de un mismo negocio."],
              "ACA o AKS", d.GREEN),
             ("Microservicios", lambda sl, le, to, wi, ac: draw_pieces(sl, le, to, wi, ac, 4, 3, Inches(0.33)),
              ["Muchos servicios pequeños: cada uno se despliega y escala solo."],
              "AKS o ACA, según la carga", d.ORANGE)]
    if SHOW_FUNCTIONS:
        cards.append(("Procesos asíncronos", draw_async,
                      ["Atienden mensajes o eventos que esperan en una cola: nadie espera la respuesta."],
                      "Functions, ACA o AKS", d.STEEL))
    slide = d.base_slide(deck, f"{COUNT_WORDS[len(cards)]} formas de construir una aplicación, cada una con su "
                               "runtime", source)
    gap = Inches(0.12)
    width = int((Inches(12.15) - gap * (len(cards) - 1)) / len(cards))
    for index, (title, drawing, lines, runtime, accent) in enumerate(cards):
        type_card(slide, Inches(0.6) + index * (width + gap), width, title, drawing, lines, runtime, accent)
    microservices = deployed_microservices(apps)
    d.callout(slide, Inches(0.6), Inches(5.9), Inches(12.15), Inches(0.75),
              f"Hoy casi todo corre en AKS: {d.number(sum(int(to_float(row.get('clusters')) or 1) for row in apps))} "
              f"clusters y {d.number(microservices)} {MICRO}, "
              "en el runtime pensado para carga alta y constante.", size=15, fill=d.RUST)
    d.text_box(slide, Inches(0.6), Inches(6.68), Inches(12.15), Inches(0.28), ACRONYMS, size=10, color=d.GREY,
               align=d.PP_ALIGN.CENTER)


def accent_card(slide, left, top, width, height, label, value, detail, accent, value_size=28):
    """A light card with a coloured edge: small label, big figure, one line of detail (as the bank's summaries)."""
    d.rounded(slide, left, top, width, height, d.LIGHT, radius=0.08)
    shape(slide, MSO_SHAPE.RECTANGLE, left, top + Inches(0.1), Inches(0.08), height - Inches(0.2), accent)
    d.text_box(slide, left + Inches(0.25), top + Inches(0.1), width - Inches(0.35), Inches(0.3), label.upper(), size=10,
               bold=True, color=d.GREY)
    d.text_box(slide, left + Inches(0.25), top + Inches(0.4), width - Inches(0.35), Inches(0.65), value,
               size=value_size, bold=True, color=d.NAVY, anchor=d.MSO_ANCHOR.MIDDLE)
    d.text_box(slide, left + Inches(0.25), top + height - Inches(0.5), width - Inches(0.35), Inches(0.4), detail,
               size=10, color=d.GREY)


def slide_estate(deck, estate, apps, source):
    """AKS only: the clusters, what runs in them and what they cost (PROD plus non-PROD), with non-PROD in small print."""
    month = month_label(estate["month"])
    category = estate["by_category"]
    prod = estate["prod"]
    prod_cost, nonprod_cost = estate["prod_cost"]["AKS"], estate["nonprod_cost"]
    microservices = deployed_microservices(apps)
    slide = d.base_slide(deck, f"Producción: {prod['aks']} clusters AKS concentran los {MICRO}", source)
    left, top, width, height = Inches(0.6), Inches(1.3), Inches(3.4), Inches(3.45)
    d.rounded(slide, left, top, width, height, d.CARD)
    d.text_box(slide, left, top + Inches(0.1), width, Inches(0.4), "AKS", size=16, bold=True, color=d.WHITE,
               align=d.PP_ALIGN.CENTER)
    icon = Inches(1.0)
    slide.shapes.add_picture(os.path.join(ICONS_DIR, "aks.png"), left + int((width - icon) / 2), top + Inches(0.6),
                             height=icon)
    d.text_box(slide, left, top + Inches(1.7), width, Inches(0.35), "Servidores encendidos 24/7", size=11,
               color=d.SOFT, align=d.PP_ALIGN.CENTER)
    d.text_box(slide, left, top + Inches(1.95), width, Inches(0.7), d.number(prod["aks"]), size=40, bold=True,
               color=d.ORANGE, align=d.PP_ALIGN.CENTER)
    d.text_box(slide, left, top + Inches(2.62), width, Inches(0.3), "clusters productivos", size=12, color=d.SOFT,
               align=d.PP_ALIGN.CENTER)
    d.text_box(slide, left, top + Inches(2.95), width, Inches(0.35),
               f"+ {d.number(estate['nonprod_clusters'])} no productivos (certificación y desarrollo)", size=10,
               color=d.SOFT, align=d.PP_ALIGN.CENTER)
    grid_left, gap = Inches(4.25), Inches(0.2)
    card_width = int((Inches(12.75) - grid_left - gap) / 2)
    app_clusters = category.get(aa.APPLICATION, {}).get("clusters", 0)
    cards = [("Clusters con aplicaciones", d.number(app_clusters), "base del análisis", d.STEEL),
             (MICRO, d.number(microservices), "servicios desplegados en esos clusters", d.GREEN),
             ("Costo mensual", soles(prod_cost + nonprod_cost),
              f"producción {soles(prod_cost)} + no productivo {soles(nonprod_cost)} ({month})", d.ORANGE),
             ("Costo anual", soles((prod_cost + nonprod_cost) * 12), "costo mensual × 12", d.RUST)]
    for index, (label, value, detail, accent) in enumerate(cards):
        accent_card(slide, grid_left + (index % 2) * (card_width + gap), top + (index // 2) * Inches(1.75), card_width,
                    Inches(1.6), label, value, detail, accent, value_size=30 if index >= 2 else 28)
    rows = [["Clusters AKS en producción", "Clusters", f"Costo {month}"]]
    for code, label in ((aa.APPLICATION, "Con aplicaciones (base del análisis)"),
                        (aa.NO_APPLICATIONS, "Creados sin aplicaciones"),
                        (aa.INFRASTRUCTURE, "Infraestructura (CI/CD, observabilidad, gateways)"),
                        (aa.NO_DATA, "Sin datos")):
        entry = category.get(code)
        if entry and entry["clusters"]:
            rows.append([label, d.number(entry["clusters"]), soles(entry["cost"])])
    d.table(slide, Inches(0.6), Inches(4.95), Inches(7.9), rows, [0.6, 0.15, 0.25], highlight_column=None,
            row_height=0.32, size=11)
    d.note_box(slide, Inches(8.85), Inches(4.95), Inches(3.9), Inches(0.85),
               f"{MICRO_CAP} = servicios de una aplicación desplegados en AKS.", size=11)


def minimum_costs(month, prices, region="eastus2"):
    """(smallest AKS cluster, one Container Apps "micro"-service with one replica on 24 h) per month in US$ after
    discount, and the list-price lines as the calculator shows them. The AKS nodes already carry the 3-year reservation; their disks and tier, and
    Container Apps (no reservations), take the bank's discount (October 2026 bills: 51.60 US$ of tier per cluster
    against 73 at list price)."""
    price = lambda key: prices.get((key, region), 0)
    term_hours = {"1y": 8760, "3y": 26280}[MIN_AKS_RESERVATION]  # a reservation is quoted for its whole term
    lines = {
        "aks_nodes": MIN_AKS_NODES * HOURS_MONTH * price(f"vm_{MIN_AKS_VM}_{MIN_AKS_RESERVATION}") / term_hours,
        "aks_disks": MIN_AKS_NODES * price(f"disk_{MIN_AKS_DISK}_month"),
        "aks_tier": HOURS_MONTH * price("aks_standard_tier_hour"),
        "aca": HOURS_MONTH * 3600 * (ACA_VCPU * price("aca_vcpu_active_s") + ACA_GIB * price("aca_mem_active_gib_s"))
               + ACA_REQUESTS_MILLION * price("aca_requests_million"),
    }
    aks = lines["aks_nodes"] + (lines["aks_disks"] + lines["aks_tier"]) * (1 - BANK_DISCOUNT)
    return aks, lines["aca"] * (1 - BANK_DISCOUNT), lines


def measured_minimum_aks(results, month):
    """(median cost in US$, clusters) of the PROD clusters that ran exactly MIN_AKS_NODES nodes of MIN_AKS_VM the whole
    month: the bill of a real smallest cluster, with nodes, disks, tier, load balancer and every discount in it.
    (None, 0) when no cluster qualifies."""
    clusters = results_csv(results, "aa-inventory", "aks_clusters.csv")
    ids_by_name = aa.cluster_ids_by_name(clusters)
    prod = {row["id"].lower() for row in clusters if row["environment"] == "PROD"}
    sizes = defaultdict(set)
    for row in results_csv(results, "aa-inventory", "aks_node_pools.csv"):
        sizes[aa.cluster_id_of(row, ids_by_name)].add(row["vm_size"])
    year, number_ = (int(part) for part in month.split("-"))
    month_hours = ((date(year + number_ // 12, number_ % 12 + 1, 1) - date(year, number_, 1)).days) * 24
    nodes, hours = defaultdict(float), {}
    for row in results_csv(results, "aa-metrics", "aks_usage_monthly.csv"):
        if row["month"] == month and row["mode"] != "ALLOCATABLE":
            cluster_id = aa.cluster_id_of(row, ids_by_name)
            nodes[cluster_id] += to_float(row["nodes_avg"]) or 0
            hours[cluster_id] = max(hours.get(cluster_id, 0), to_float(row["hours"]) or 0)
    cost = defaultdict(float)
    for row in results_csv(results, "aa-cost", "cost_by_resource.csv"):
        if row["owner_type"] == "AKS" and row["month"] == month:
            cost[row["owner_id"].lower()] += to_float(row["cost_usd"]) or 0
    bills = sorted(cost[cluster_id] for cluster_id in prod
                   if sizes.get(cluster_id) == {MIN_AKS_VM} and abs(nodes.get(cluster_id, 0) - MIN_AKS_NODES) <= 0.1
                   and hours.get(cluster_id, 0) >= month_hours - 1 and cost.get(cluster_id))
    if not bills:
        return None, 0
    middle = len(bills) // 2
    return (bills[middle] if len(bills) % 2 else (bills[middle - 1] + bills[middle]) / 2), len(bills)


PEAK_COLOR, TRACK_COLOR = d.RGBColor(0xF7, 0xD3, 0xB0), d.RGBColor(0x3A, 0x50, 0x75)


def sign_card(slide, left, top, width, height, sentence, value, label, bar=None):
    """One sign of the hypothesis: a sentence the presenter can say, the figure and what it counts. bar = (average,
    peak) draws the whole capacity with the average strong and the peak light."""
    d.rounded(slide, left, top, width, height, d.CARD)
    d.text_box(slide, left + Inches(0.12), top + Inches(0.1), width - Inches(0.24), Inches(0.6), sentence, size=12,
               bold=True, color=d.WHITE, align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
    shape(slide, MSO_SHAPE.RECTANGLE, left + Inches(0.5), top + Inches(0.75), width - Inches(1.0), Inches(0.02), d.ORANGE)
    d.text_box(slide, left, top + Inches(0.82), width, Inches(0.6), value, size=28, bold=True, color=d.ORANGE,
               align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
    d.text_box(slide, left + Inches(0.12), top + Inches(1.45), width - Inches(0.24), Inches(1.0), label, size=10.5,
               color=d.SOFT, align=d.PP_ALIGN.CENTER)
    if bar:
        average, peak = bar
        bar_left, bar_width = left + Inches(0.25), width - Inches(0.5)
        bar_top, bar_height = top + height - Inches(0.3), Inches(0.12)
        for share, colour in ((1.0, TRACK_COLOR), (peak or 0, PEAK_COLOR), (average or 0, d.ORANGE)):
            if share > 0:
                d.rounded(slide, bar_left, bar_top, max(int(bar_width * min(share, 1.0)), Inches(0.04)), bar_height,
                          colour, radius=0.5)


def slide_hypothesis(deck, estate, apps, scaling, results, source):
    month = estate["month"]
    aks_min, aca_min, lines = minimum_costs(month, read_prices(results))
    aa.azure.LOG.info("Minimum AKS at list price (as in the calculator): nodes %.2f + disks %.2f + tier %.2f = %.2f US$; "
                      "Container Apps %.2f US$ (1 replica, 730 h)", lines["aks_nodes"], lines["aks_disks"],
                      lines["aks_tier"], lines["aks_nodes"] + lines["aks_disks"] + lines["aks_tier"], lines["aca"])
    # Cross-check only: what the real smallest clusters cost in the month.
    measured_min, measured_clusters = measured_minimum_aks(results, month)
    if measured_min:
        aa.azure.LOG.info("Minimum AKS after discount: %.2f US$; real bills %.2f US$ (median of %d clusters, %s)",
                          aks_min, measured_min, measured_clusters, month)
    aks_label = (f"el cluster AKS más chico ({MIN_AKS_NODES} servidores con reserva de 3 años, discos y tier "
                 f"Standard), aunque la aplicación sea pequeña")
    app_clusters = [row for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                    if row["environment"] == "PROD" and row["category"] == aa.APPLICATION and row["in_inventory"] == "YES"]
    small = sum(0 < int(row["microservices"] or 0) <= MAX_MICROSERVICES for row in app_clusters)
    cpu_average = d.share_of(estate["cpu_used"], estate["cpu_total"])
    cpu_peak = d.share_of(estate["cpu_peak"], estate["cpu_total"])
    stuck = d.share_of(scaling["over_3_months"], scaling["dated"]) if scaling["dated"] else None
    slide = d.base_slide(deck, "Hipótesis: un cluster AKS por aplicación paga más de lo que usa", source)
    d.heading(slide, Inches(0.6), Inches(1.3), Inches(7.2), "Indicios", size=14)
    signs = [("Pagamos CPU que casi no usamos", d.percent(cpu_average),
              f"de la CPU de los clusters se usa en horario de oficina (pico del mes: {d.percent(cpu_peak)})",
              (cpu_average, cpu_peak)),
             ("El autoescalado casi no se mueve", d.percent(stuck),
              f"de los {MICRO} con autoescalado no sube ni baja de réplicas hace más de 3 meses", None),
             ("Clusters enteros para pocas cargas", f"{d.number(small)} de {d.number(len(app_clusters))}",
              f"clusters con aplicaciones tienen {MAX_MICROSERVICES} {MICRO} o menos", None)]
    width, gap = Inches(2.3), Inches(0.15)
    for index, (sentence, value, label, bar) in enumerate(signs):
        sign_card(slide, Inches(0.6) + index * (width + gap), Inches(1.75), width, Inches(2.8), sentence, value, label,
                  bar)
    d.heading(slide, Inches(8.1), Inches(1.3), Inches(4.6), "Motivación", size=14)
    d.callout(slide, Inches(8.1), Inches(1.75), Inches(4.6), Inches(1.2),
              "Económica: pagar por uso, no por servidores encendidos 24/7", size=15, fill=d.RUST)
    for index, text in enumerate(("Operación: menos clusters que actualizar y vigilar",
                                  "Complejidad: sin servidores ni versiones de Kubernetes que administrar")):
        d.callout(slide, Inches(8.1), Inches(3.05) + index * (Inches(0.75)), Inches(4.6), Inches(0.65), text, size=12,
                  fill=d.CARD)
    d.heading(slide, Inches(0.6), Inches(4.75), Inches(12.1), f"Lo mínimo que se paga al mes, según la calculadora "
              f"de Azure (con el {d.percent(BANK_DISCOUNT)} de descuento del banco)", size=14)
    d.kpi(slide, Inches(0.6), Inches(5.15), Inches(5.9), Inches(1.3), soles(aks_min), aks_label,
          value_color=d.RUST, value_size=28, label_size=11)
    d.kpi(slide, Inches(6.8), Inches(5.15), Inches(5.9), Inches(1.3), soles(aca_min),
          f"un {MICRO_ONE} en Container Apps, con 1 réplica encendida las 24 horas", value_color=d.GREEN,
          value_size=28, label_size=11)
    # Both minimums leave out the load balancer every AKS cluster and every Container Apps environment pays (the
    # saving on slide 7 includes it on both sides).
    d.text_box(slide, Inches(0.6), Inches(6.5), Inches(12.1), Inches(0.25), "No incluye el balanceador.", size=9,
               color=d.GREY, italic=True)


def apps_label(count):
    return f"{d.number(count)} {'aplicación' if count == 1 else 'aplicaciones'}"


def criteria(row):
    """The conditions of slide 4 for one application_snapshot row, and where it goes:
    NO_TRAFFIC (no PROD traffic measured), AKS (fails condition 1 or 2), ACA_WITH_REVIEW (meets both, but is critical or
    grows past MAX_MICROSERVICES with DESA) or ACA. aa_criteria_workbook.py writes the same tests as Excel formulas."""
    measured = row.get("traffic_status") == "MEASURED" and row.get("busiest_hour_tps", "") != ""
    microservices = int(to_float(row.get("microservices")) or 0)
    projected = int(to_float(row.get("microservices_projected")) or 0)
    out = {"has_traffic": measured,
           "tps_ok": measured and (to_float(row["busiest_hour_tps"]) or 0) <= MAX_TPS,
           "size_ok": microservices <= MAX_MICROSERVICES,
           "critical": row.get("bia") in CRITICAL_BIA,
           "growing": projected > MAX_MICROSERVICES >= microservices}
    if not measured:
        out["decision"] = "NO_TRAFFIC"
    elif not (out["tps_ok"] and out["size_ok"]):
        out["decision"] = "AKS"
    else:
        out["decision"] = "ACA_WITH_REVIEW" if out["critical"] or out["growing"] else "ACA"
    return out


def decision(row):
    return criteria(row)["decision"]


MOVABLE = ("ACA", "ACA_WITH_REVIEW")
# The two conditions to move to Container Apps: (funnel label, title, plain meaning, criteria() key).
ACA_CONDITIONS = [
    ("Hasta 100 TPS", "Tráfico moderado", "Hasta 100 transacciones por segundo (TPS) en su hora pico.", "tps_ok"),
    (f"Hasta 15 {MICRO}", f"Hasta 15 {MICRO}", "Si tiene más, se queda en AKS.", "size_ok"),
]
OFFHOURS_RULE = ("Tiende a cero fuera de horario (de 20 a 8 h y fines de semana)",
                 ["Se puede apagar: la primera transacción la despierta en hasta ~30 s. Validar el timeout de quien la "
                  "llama.",
                  f"En el costo se cuenta 1 réplica encendida las 24 horas por {MICRO_ONE}."])
CHECKS_BEFORE_MOVING = ["Si es crítica para el banco, se valida con gobierno de arquitectura.",
                        f"Si está creciendo y pronto pasará de 15 {MICRO}, se confirma antes de mover."]


def condition_card(slide, left, top, width, height, badge, colour, title, lines, oval=True):
    box(slide, left, top, width, height, d.LIGHT)
    shape(slide, MSO_SHAPE.OVAL if oval else MSO_SHAPE.RECTANGLE, left + Inches(0.3), top + Inches(0.4), Inches(0.6),
          Inches(0.6), colour)
    d.text_box(slide, left + Inches(0.3), top + Inches(0.4), Inches(0.6), Inches(0.6), badge, size=18, bold=True,
               color=d.WHITE, align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
    d.text_box(slide, left + Inches(1.15), top + Inches(0.2), width - Inches(1.4), Inches(0.45), title, size=18,
               bold=True, color=d.NAVY)
    d.text_box(slide, left + Inches(1.15), top + Inches(0.68), width - Inches(1.4), height - Inches(0.75), lines,
               size=14, color=d.GREY)


def slide_criteria(deck, source):
    slide = d.base_slide(deck, "Para pasar a Container Apps: dos condiciones y una regla para la noche", source)
    width, height, gap = Inches(5.9), Inches(1.35), Inches(0.3)
    for index, (_, title, meaning, _) in enumerate(ACA_CONDITIONS):
        condition_card(slide, Inches(0.6) + index * (width + gap), Inches(1.35), width, height, str(index + 1), d.GREEN,
                       title, meaning)
    title, lines = OFFHOURS_RULE
    condition_card(slide, Inches(0.6), Inches(2.9), Inches(12.1), Inches(1.45), "3", YELLOW, title, lines, oval=False)
    top = Inches(4.55)
    box(slide, Inches(0.6), top, Inches(12.1), Inches(1.35), d.LIGHT)
    d.text_box(slide, Inches(0.9), top + Inches(0.12), Inches(11.5), Inches(0.4), "Antes de mover, se valida:", size=16,
               bold=True, color=d.ORANGE)
    d.text_box(slide, Inches(0.9), top + Inches(0.52), Inches(11.5), Inches(0.8),
               [f"• {line}" for line in CHECKS_BEFORE_MOVING], size=14, color=d.GREY)
    d.text_box(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(0.4),
               "Si no cumple la 1 o la 2, se queda en AKS.", size=14, bold=True, color=d.NAVY)


# Groups of the distribution slide, in plain words for a non-technical audience: (label, what it means, colour).
DISTRIBUTION_GROUPS = {
    "AKS": ("Se queda en AKS", f"Más de 100 TPS o más de 15 {MICRO}", d.NAVY),
    "ACA_WITH_REVIEW": ("Container Apps, con revisión", "Cumple, pero es crítica o está creciendo", d.ORANGE),
    "ACA": ("Candidata a Container Apps", f"Hasta 100 TPS y hasta 15 {MICRO}", d.GREEN),
    "IN_USE": ("En uso, sin medición", "Tiene uso, pero su tráfico no se ve", d.STEEL),
    "NOT_IN_USE": ("Sin uso", "Por revisar", d.GREY),
}


def distribution_group(row):
    """Slide group of an application; None for one still in development (shown in the note, not in the chart)."""
    key = decision(row)
    if key != "NO_TRAFFIC":
        return key
    reason = row.get("no_traffic_reason")
    if reason == "IN_DEVELOPMENT":
        return None
    return "IN_USE" if reason == "IN_USE_NOT_MEASURED" else "NOT_IN_USE"


HAIRLINE = d.RGBColor(0xD5, 0xDB, 0xE3)


def hairline(slide, left, top, width, color=HAIRLINE, weight=0.75):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left, top, left + width, top)
    line.line.color.rgb = color
    line.line.width = d.Pt(weight)
    return line


def badge(slide, centre_x, centre_y, fails, known=True):
    """A condition in the result table: a green tick when the whole group meets it, a rust circle with the number of
    applications that do not, a dash when there is no traffic to test."""
    size = Inches(0.3)
    if not known:
        d.text_box(slide, centre_x - Inches(0.3), centre_y - Inches(0.15), Inches(0.6), Inches(0.3), "—", size=11,
                   color=d.SOFT, align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
        return
    circle = shape(slide, MSO_SHAPE.OVAL, centre_x - size / 2, centre_y - size / 2, size, size,
                   d.GREEN if not fails else d.RUST)
    circle.shadow.inherit = False
    text = slide.shapes.add_textbox(centre_x - size / 2, centre_y - size / 2, size, size).text_frame
    text.margin_left = text.margin_right = text.margin_top = text.margin_bottom = 0
    text.vertical_anchor = d.MSO_ANCHOR.MIDDLE
    paragraph = text.paragraphs[0]
    paragraph.alignment = d.PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = "✓" if not fails else d.number(fails)
    run.font.name, run.font.size, run.font.bold = d.FONT, d.Pt(10 if fails else 12), True
    run.font.color.rgb = d.WHITE


def stat_card(slide, left, top, width, height, label, value, detail, value_color, fill=d.LIGHT, text_color=d.NAVY):
    card = d.rounded(slide, left, top, width, height, fill, radius=0.08)
    card.shadow.inherit = False
    d.text_box(slide, left + Inches(0.25), top + Inches(0.1), width - Inches(0.5), Inches(0.3), label.upper(), size=10,
               bold=True, color=d.GREY if fill == d.LIGHT else d.SOFT)
    d.text_box(slide, left + Inches(0.25), top + Inches(0.38), width - Inches(0.5), Inches(0.55), value, size=26,
               bold=True, color=value_color, anchor=d.MSO_ANCHOR.MIDDLE)
    d.text_box(slide, left + Inches(0.25), top + Inches(0.92), width - Inches(0.5), Inches(0.3), detail, size=10,
               color=text_color)


def slide_distribution(deck, apps, estate, results, source):
    """Where each application goes, without the new cost (that is slide 6): a chart of the groups; a table with the
    conditions each group meets, where it goes and what it costs on AKS today; and the comparison in applications
    (today all on AKS, then how many stay and how many go to Container Apps). Replaces the funnel slide."""
    month = month_label(estate["month"])
    groups = defaultdict(lambda: {"apps": [], "clusters": 0, "today": 0.0, "fails_tps": 0, "fails_size": 0})
    developing = []
    for row in apps:
        key = distribution_group(row)
        if key is None:
            developing.append(row["cod_app"])
            continue
        test, today = criteria(row), to_float(row["cost_month_usd"]) or 0
        entry = groups[key]
        entry["apps"].append(row)
        entry["clusters"] += int(to_float(row.get("clusters")) or 1)
        entry["today"] += today
        clusters = int(to_float(row.get("clusters")) or 1)
        entry["fails_tps"] += clusters if test["has_traffic"] and not test["tps_ok"] else 0
        entry["fails_size"] += clusters if not test["size_ok"] else 0
    movable = sum(len(groups[key]["apps"]) for key in MOVABLE)
    movable_clusters = sum(groups[key]["clusters"] for key in MOVABLE)
    order = [key for key in DISTRIBUTION_GROUPS if groups[key]["apps"]]
    counted_clusters = sum(groups[key]["clusters"] for key in order)
    slide = d.base_slide(deck, f"Resultado: {movable_clusters} de {counted_clusters} clusters podrían pasar a Container "
                               "Apps", source)
    short = {"AKS": "AKS", "ACA_WITH_REVIEW": "ACA con\nrevisión", "ACA": "ACA", "IN_USE": "Sin\nmedición",
             "NOT_IN_USE": "Sin uso"}
    d.bar_chart(slide, Inches(0.45), Inches(1.3), Inches(4.1), Inches(3.45), [short[key] for key in order],
                [groups[key]["clusters"] for key in order], XL_CHART_TYPE.COLUMN_CLUSTERED,
                [DISTRIBUTION_GROUPS[key][2] for key in order])

    # Table drawn with shapes: header, one row per group with hairlines, a tinted total row.
    left, width, top, row_height = Inches(4.85), Inches(8.0), Inches(1.35), Inches(0.46)
    columns = [("Destino", 2.6, "left"), ("≤ 100 TPS", 1.0, "center"), (f"≤ 15 {MICRO}", 1.3, "center"),
               ("Clusters", 1.0, "right"), ("AKS hoy / mes", 2.1, "right")]
    table_labels = {"AKS": "Se queda en AKS", "ACA_WITH_REVIEW": "ACA, con revisión", "ACA": "Candidata a ACA",
                    "IN_USE": "En uso, sin medición", "NOT_IN_USE": "Sin uso"}
    edges, x = [], left
    for _, share, _ in columns:
        edges.append((x, Inches(share)))
        x += Inches(share)
    align = {"left": d.PP_ALIGN.LEFT, "center": d.PP_ALIGN.CENTER, "right": d.PP_ALIGN.RIGHT}

    def cell(index, y, text, size=11, bold=False, color=d.GREY):
        x0, w = edges[index]
        box_ = d.text_box(slide, x0 + Inches(0.05), y, w - Inches(0.1), row_height, text, size=size, bold=bold,
                          color=color, align=align[columns[index][2]], anchor=d.MSO_ANCHOR.MIDDLE)
        box_.text_frame.margin_left = box_.text_frame.margin_right = 0
    for index, (label, _, _) in enumerate(columns):
        cell(index, top, label, size=9.5, bold=True, color=d.NAVY)
    hairline(slide, left, top + row_height, width, color=d.NAVY, weight=1.25)
    y = top + row_height
    for key in order:
        label, colour = table_labels[key], DISTRIBUTION_GROUPS[key][2]
        entry = groups[key]
        dot = Inches(0.16)
        shape(slide, MSO_SHAPE.OVAL, left + Inches(0.05), y + (row_height - dot) / 2, dot, dot, colour)
        d.text_box(slide, left + Inches(0.28), y, Inches(2.15), row_height, label, size=11, bold=True, color=d.NAVY,
                   anchor=d.MSO_ANCHOR.MIDDLE)
        known = key not in ("IN_USE", "NOT_IN_USE")
        for index, fails in ((1, entry["fails_tps"]), (2, entry["fails_size"])):
            badge(slide, edges[index][0] + edges[index][1] / 2, y + row_height / 2, fails, known)
        cell(3, y, d.number(entry["clusters"]), bold=True, color=d.NAVY)
        cell(4, y, soles(entry["today"]))
        y += row_height
        hairline(slide, left, y, width)
    today = sum(entry["today"] for entry in groups.values())
    counted = sum(len(groups[key]["apps"]) for key in order)
    total = d.rounded(slide, left, y + Inches(0.04), width, row_height, d.LIGHT, radius=0.1)
    total.shadow.inherit = False
    d.text_box(slide, left + Inches(0.28), y + Inches(0.04), Inches(2.15), row_height, "Total", size=11, bold=True,
               color=d.NAVY, anchor=d.MSO_ANCHOR.MIDDLE)
    cell(3, y + Inches(0.04), d.number(counted_clusters), bold=True, color=d.NAVY)
    cell(4, y + Inches(0.04), soles(today), bold=True, color=d.NAVY)

    # The comparison in clusters (what is paid and operated): today all on AKS; after, how many stay and how many are
    # turned off because their application moves to Container Apps.
    cards_top, card_h, gap = Inches(5.0), Inches(1.25), Inches(0.3)
    card_w = int((Inches(12.85) - Inches(0.45) - 2 * gap) / 3)
    stay = counted_clusters - movable_clusters
    stat_card(slide, Inches(0.45), cards_top, card_w, card_h, "Hoy", f"{d.number(counted_clusters)} clusters",
              "todos en AKS", d.NAVY)
    stat_card(slide, Inches(0.45) + card_w + gap, cards_top, card_w, card_h, "Nuevo: AKS", f"{d.number(stay)} clusters",
              f"se quedan ({d.percent(d.share_of(stay, counted_clusters))})", d.NAVY)
    stat_card(slide, Inches(0.45) + 2 * (card_w + gap), cards_top, card_w, card_h, "Nuevo: ACA",
              f"{d.number(movable_clusters)} clusters",
              f"se apagan y pasan a Container Apps ({d.percent(d.share_of(movable_clusters, counted_clusters))})",
              d.WHITE, fill=d.CARD, text_color=d.SOFT)
    notes = [f"Número en rojo: clusters que no cumplen la condición. Costo AKS de {month}, con IGV."]
    single = sum(int(to_float(row.get("clusters")) or 1) for row in apps if row["possible_monolith"] == "YES")
    extra = []
    if single:
        extra.append(f"{d.number(single)} clusters con un solo {MICRO_ONE} podrían ir a App Service")
    developing_clusters = sum(int(to_float(row.get("clusters")) or 1) for row in apps if row["cod_app"] in developing)
    if developing_clusters:
        extra.append(f"{d.number(developing_clusters)} en desarrollo quedan fuera del total")
    if extra:
        text = "; ".join(extra)
        notes.append(text[0].upper() + text[1:] + ".")
    d.text_box(slide, Inches(0.45), Inches(6.35), Inches(12.4), Inches(0.55), notes, size=9, color=d.GREY)


def office_seconds(month):
    """Seconds of office hours (Monday to Friday, 8:00 to 20:00) in a YYYY-MM month: the ADT_019 warm window."""
    year, number_ = (int(part) for part in month.split("-"))
    days = (date(year + number_ // 12, number_ % 12 + 1, 1) - date(year, number_, 1)).days
    weekdays = sum(date(year, number_, day).weekday() < 5 for day in range(1, days + 1))
    return weekdays * (aa.OFFICE_END_HOUR - aa.OFFICE_START_HOUR) * 3600


def read_prices(results):
    """List prices by key and region (retail_price), from prices.csv."""
    return {(row["key"], row["region"]): to_float(row["retail_price"]) or 0
            for row in results_csv(results, "aa-prices", "prices.csv") if row["key"] != "aca_other"}


def measured_discount(results, month, prices):
    """Amortized cost of PROD node VMs / their node-hours at list price, in the month. 1 when it cannot be measured."""
    clusters = results_csv(results, "aa-inventory", "aks_clusters.csv")
    ids_by_name = aa.cluster_ids_by_name(clusters)
    regions = {row["id"].lower(): row["location"] for row in clusters}
    sizes = {(row["cluster_id"].lower(), row["pool"].lower()): row["vm_size"]
             for row in results_csv(results, "aa-inventory", "aks_node_pools.csv")}
    listed = 0.0
    for row in results_csv(results, "aa-metrics", "aks_usage_monthly.csv"):
        if row["month"] != month or row["environment"] != "PROD" or row["mode"] == "ALLOCATABLE":
            continue
        cluster_id = aa.cluster_id_of(row, ids_by_name)
        price = prices.get((f"vm_{sizes.get((cluster_id, row['pool'].lower()), '')}_payg", regions.get(cluster_id, "")))
        listed += (to_float(row["nodes_avg"]) or 0) * (to_float(row["hours"]) or 0) * (price or 0)
    paid = sum(to_float(row["cost_usd"]) or 0 for row in results_csv(results, "aa-cost", "cost_by_resource.csv")
               if row["owner_type"] == "AKS" and row["environment"] == "PROD" and row["month"] == month
               and row["meter_category"] == "Virtual Machines")
    return min(1.0, paid / listed) if listed and paid else 1.0


# Headroom on the requests each application received in the month, so the saving is not overstated.
REQUESTS_BUFFER = 1.25
# Container Apps environments per application: each one bills its internal Standard load balancer.
ENVIRONMENTS_PER_APP = 1


def month_seconds(month):
    year, number_ = (int(part) for part in month.split("-"))
    return (date(year + number_ // 12, number_ % 12 + 1, 1) - date(year, number_, 1)).days * 86400


def aca_size(row):
    """(vCPU, GiB) of the smallest Consumption size that covers a deployment's request: 0.25 vCPU steps, 2 GiB per
    vCPU."""
    cpu, mem = to_float(row["cpu_request_vcpu"]) or 0, to_float(row["mem_request_gib"]) or 0
    vcpu = math.ceil(max(cpu, mem / 2, 0.25) * 4 - 1e-9) / 4
    return vcpu, 2 * vcpu


def office_replicas(row):
    """Office-hours replicas of a deployment: the Container Insights average, else the HPA minimum; at least 1."""
    return max(1.0, to_float(row["ci_replicas_avg_office"]) or to_float(row["hpa_min"]) or 1)


def aca_breakdown(results, apps, month, prices, ratio, region="eastus2", offhours_replicas=OFFHOURS_REPLICAS):
    """One row per application: the inputs of its monthly Container Apps cost and every step of the formula, so it
    can be rebuilt by hand, in the Azure pricing calculator or in the criteria workbook (same formula).

    Per "micro"-service: allocation = aca_size(); office hours (Monday to Friday 8:00-20:00) run office_replicas();
    the rest of the month runs offhours_replicas, billed at the active rate (an idle replica may bill less). Then:
        vcpu_s = (office replicas x office seconds + off-hours replicas x off-hours seconds) x vCPU
        gib_s  = the same with GiB
        requests = requests of the month x REQUESTS_BUFFER
        list = vcpu_s x vCPU-s price + gib_s x GiB-s price + requests / 1e6 x price per million
               + ENVIRONMENTS_PER_APP x 730 h x Standard load balancer hour price
        with discount = list x ratio
    The free grant is left out (one per subscription)."""
    price = lambda key: prices.get((key, region), 0)
    office = office_seconds(month)
    offhours = month_seconds(month) - office
    application = {row["cluster_id"]: row["cod_app"]
                   for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                   if row["category"] == aa.APPLICATION and row["environment"] == "PROD"}
    per_app = defaultdict(lambda: {"microservices": 0, "replicas": 0.0, "vcpu_s": 0.0, "gib_s": 0.0,
                                   "offhours_vcpu_s": 0.0, "offhours_gib_s": 0.0})
    for row in results_csv(results, "aa-k8s-snapshot", "deployment_snapshot.csv"):
        code = application.get(row["cluster_id"])
        if not code or row["microservice"] != "YES":
            continue
        vcpu, gib = aca_size(row)
        replicas = office_replicas(row)
        entry = per_app[code]
        entry["microservices"] += 1
        entry["replicas"] += replicas
        entry["vcpu_s"] += replicas * vcpu * office
        entry["gib_s"] += replicas * gib * office
        entry["offhours_vcpu_s"] += offhours_replicas * vcpu * offhours
        entry["offhours_gib_s"] += offhours_replicas * gib * offhours
    # A Consumption environment inside the VNet bills only its internal Standard load balancer (October 2026: 13.02 US$
    # per environment, 744 h x 0.025 x 0.70, in all 24 of the bank's environments). The 0.10 US$/h management charge
    # applies only with dedicated profiles, private endpoints or planned maintenance, which they do not use.
    environment_hour = price("lb_standard_rules_hour")
    out = []
    for app in apps:
        entry = per_app[app["cod_app"]]
        test = criteria(app)
        requests = (to_float(app["requests_month"]) or 0) * REQUESTS_BUFFER
        compute_cost = (entry["vcpu_s"] + entry["offhours_vcpu_s"]) * price("aca_vcpu_active_s")
        memory_cost = (entry["gib_s"] + entry["offhours_gib_s"]) * price("aca_mem_active_gib_s")
        offhours_cost = (entry["offhours_vcpu_s"] * price("aca_vcpu_active_s")
                         + entry["offhours_gib_s"] * price("aca_mem_active_gib_s"))
        requests_cost = requests / 1e6 * price("aca_requests_million")
        environment_cost = ENVIRONMENTS_PER_APP * HOURS_MONTH * environment_hour
        total = compute_cost + memory_cost + requests_cost + environment_cost
        aks = to_float(app["cost_month_usd"]) or 0
        out.append({"cod_app": app["cod_app"], "decision": test["decision"],
                    "peak_hour_tps_ok": yes_no(test["tps_ok"], test["has_traffic"]),
                    "microservices_ok": yes_no(test["size_ok"]), "critical": yes_no(test["critical"]),
                    "growing": yes_no(test["growing"]), "adt_runtime_reference": app["adt_runtime"], "month": month,
                    "microservices": entry["microservices"], "replicas_office": round(entry["replicas"], 2),
                    "offhours_replicas_per_microservice": offhours_replicas,
                    "office_seconds": office, "offhours_seconds": offhours,
                    "vcpu_seconds": round(entry["vcpu_s"] + entry["offhours_vcpu_s"]),
                    "gib_seconds": round(entry["gib_s"] + entry["offhours_gib_s"]),
                    "requests_month": int(to_float(app["requests_month"]) or 0),
                    "requests_with_buffer": int(requests), "price_vcpu_s": price("aca_vcpu_active_s"),
                    "price_gib_s": price("aca_mem_active_gib_s"), "price_requests_million": price("aca_requests_million"),
                    "price_environment_hour": environment_hour, "environments": ENVIRONMENTS_PER_APP,
                    "compute_usd": round(compute_cost, 2), "memory_usd": round(memory_cost, 2),
                    "offhours_part_usd": round(offhours_cost, 2),
                    "requests_usd": round(requests_cost, 2), "environment_usd": round(environment_cost, 2),
                    "aca_list_usd": round(total, 2), "discount_ratio": round(ratio, 4),
                    "aca_with_discount_usd": round(total * ratio, 2),
                    "aks_cost_usd": aks,
                    "aca_with_discount_pen_igv": round(total * ratio * EXCHANGE_RATE * (1 + IGV), 2),
                    "aks_cost_pen_igv": round(aks * EXCHANGE_RATE * (1 + IGV), 2)})
    return out


def yes_no(value, known=True):
    return ("YES" if value else "NO") if known else "NO_DATA"


def aca_estimate(results, apps, month, prices, region="eastus2"):
    """Monthly Container Apps cost per app code at list price (aca_breakdown, without discount)."""
    return {row["cod_app"]: row["aca_list_usd"] for row in aca_breakdown(results, apps, month, prices, 1.0, region)}


def slide_cost(deck, apps, estate, results, source, ratio):
    month = estate["month"]
    prices = read_prices(results)
    aca = aca_estimate(results, apps, month, prices)
    movable = [row for row in apps if decision(row) in MOVABLE]
    aks_paid = sum(to_float(row["cost_month_usd"]) or 0 for row in movable)
    aca_list = sum(aca[row["cod_app"]] for row in movable)
    # AKS: what the bank paid (reservations included); Container Apps in consumption has no reservations, so only
    # the bank's discount on list prices applies to it.
    scenarios = [("Precio de lista (sin descuento)", aks_paid / ratio, aca_list),
                 ("Con descuentos (AKS con reservas)", aks_paid, aca_list * (1 - BANK_DISCOUNT))]
    saving = scenarios[1][1] - scenarios[1][2]
    aa.azure.LOG.info("Saving with discounts: %d applications, AKS %.2f - ACA %.2f = %.2f US$ a month", len(movable),
                      aks_paid, aca_list * (1 - BANK_DISCOUNT), saving)
    clusters = sum(int(to_float(row.get("clusters")) or 1) for row in movable)
    slide = d.base_slide(deck, f"Apagar {clusters} clusters ahorraría ~{soles(saving)} al mes ({soles(saving * 12)} al "
                               "año)", source)
    rows = [["Escenario (con IGV)", "AKS hoy / mes", "Container Apps / mes", "Ahorro / mes", "Ahorro / año"]]
    for label, aks_cost, aca_cost in scenarios:
        rows.append([label, soles(aks_cost), soles(aca_cost), soles(aks_cost - aca_cost),
                     soles((aks_cost - aca_cost) * 12)])
    d.table(slide, Inches(0.6), Inches(1.4), Inches(12.1), rows, [0.3, 0.17, 0.19, 0.17, 0.17], highlight_column=3,
            row_height=0.5, size=13)
    rows = [["Cómo se paga", "AKS", "Container Apps"],
            ["Cobro", "Nodos completos (VM) las 24 horas", "vCPU-segundo y GB-segundo en uso, y transacciones"],
            ["Fuera de horario", "Se paga igual", f"1 réplica encendida por {MICRO_ONE} (podría bajar a cero)"],
            ["Descuento", f"{d.percent(1 - ratio)} de descuento por reservas en las VM",
             f"{d.percent(BANK_DISCOUNT)} del banco (sin reservas)"]]
    d.table(slide, Inches(0.6), Inches(3.25), Inches(12.1), rows, [0.2, 0.38, 0.42], row_height=0.45, size=12)
    d.note_box(slide, Inches(0.6), Inches(5.35), Inches(12.1), Inches(0.65),
               f"Estimado con las réplicas medidas en {month_label(month)} en horario de oficina, cada una del tamaño que "
               f"pide hoy, y 1 réplica encendida fuera de horario. Incluye los clusters \"con revisión\": el ahorro "
               f"firme es menor. Detalle en el Excel de criterios.", size=11)
    d.note_box(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(0.6),
               "Beneficio adicional sin monto: menos clusters que actualizar y mantener (horas-hombre de los equipos).",
               size=11)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: aks_aca_corporate.pptx in the newest reports/<date>-aks-aca folder")
    parser.add_argument("--aks-discount", type=float, default=None,
                        help="Discount on the list price of AKS node VMs, as a fraction (0.60 = 60%%). Default: measured, "
                             "the amortized cost of PROD node VMs in the month over their node-hours at list price")
    args = parser.parse_args()
    aa.azure.setup_logging(False)

    estate, apps = read_estate(args.results), read_applications(args.results)
    if not apps:
        raise SystemExit("No application_snapshot.csv in " + os.path.join(args.results, "aa-k8s-snapshot")
                         + ": run aa_k8s_snapshot.py first.")
    scaling = read_scaling(args.results)
    inventory_months = {row.get("measurement_month", "") for row in results_csv(args.results, "aa-inventory",
                                                                                "aks_clusters.csv")} - {""}
    if inventory_months and inventory_months != {estate["month"]}:
        aa.azure.LOG.warning("The inventory was filtered for %s but cost and metrics are of %s: rerun aa_inventory.py "
                             "with --measurement-month %s", ", ".join(sorted(inventory_months)), estate["month"],
                             estate["month"])
    period = month_label(estate["month"])
    measured = f"Fuente: inventario de Azure, Cost Management (costo amortizado, {period}) y métricas de {period}."
    traffic_source = (f"Fuente: tráfico de APIM, Application Gateway y Event Hubs, Container Insights y configuración de "
                      f"autoescalado, {period}. Referencia: lineamiento ADT_019.")

    prices = read_prices(args.results)
    if args.aks_discount is not None:
        aks_ratio = 1 - args.aks_discount
    else:
        aks_ratio = measured_discount(args.results, estate["month"], prices)
    aa.azure.LOG.info("AKS node VMs: %.0f%% of the list price is paid (%s)", aks_ratio * 100,
                      "--aks-discount" if args.aks_discount is not None else "measured")
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    slide_context(deck, apps, "Fuente: lineamientos de arquitectura del banco (FAT_001, ADT_019).")
    slide_estate(deck, estate, apps, f"{measured} {MONEY_NOTE}")
    slide_hypothesis(deck, estate, apps, scaling, args.results,
                     f"{measured} Mínimos: calculadora de precios de Azure (East US 2, precios de lista). {MONEY_NOTE}")
    slide_criteria(deck, "Referencia: lineamiento ADT_019 del banco; arranque en frío según Microsoft Learn (Container "
                         "Apps).")
    slide_distribution(deck, apps, estate, args.results, f"{traffic_source} {MONEY_NOTE}")
    slide_cost(deck, apps, estate, args.results,
               f"{measured} Precios de lista: Azure Retail Prices API. {MONEY_NOTE}", aks_ratio)
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, "aks_aca_corporate.pptx")
    deck.save(output)
    breakdown = aca_breakdown(args.results, apps, estate["month"], prices, 1 - BANK_DISCOUNT)
    aa.azure.Output(os.path.dirname(os.path.abspath(output)), ",").csv("aca_estimate_by_app.csv", breakdown)
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


if __name__ == "__main__":
    main()
