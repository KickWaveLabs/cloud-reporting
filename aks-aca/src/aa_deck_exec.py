#!/usr/bin/env python3
"""Build the corporate AKS vs Container Apps deck (slides 1 to 7) for a non-technical forum.

Same look as aa_deck.py (its drawing helpers are reused), plain figures only: TPS, vCPU, GB of RAM, instances,
hours with traffic, money per month and per year. Slide text is in Spanish for the audience. Every amount is in
soles with IGV (18 %) at S/ 3.41333 per US$ (soles()), as the forum reads the bank's numbers.

Reads only results/ (inventory, metrics, cost, prices, snapshot), so it needs no saved workbook. The cost slide
uses a simplified ADT_019 model (see aca_estimate); the workbook's ACA_by_cluster sheet has the detail per
cluster. No Azure call is made; nothing leaves the machine.

Slides (the story: what we build, where it runs, why we challenge it, the rules, the result, the money):
  1. Ways to build an application: modular monolith, miniservices, microservices, asynchronous processes, each
     drawn and with its usual runtime (FAT_001 and ADT_019)
  2. The PROD runtime estate: AKS, Container Apps, App Service and Functions, each drawn as it works inside, and
     the AKS clusters by category (only the clusters with applications enter the analysis)
  3. Hypothesis: signs that one AKS cluster per application costs more than it uses, the motivation (economic
     first, then operation and complexity) and the smallest AKS cluster against one Container Apps microservice
  4. The four conditions to move to Container Apps (the ADT_019 tree read the other way round: meeting all four
     is the same as meeting none of the AKS conditions), plus what is validated before moving (critical, growing)
  5. The conditions applied as a funnel: applications with measured traffic, then those left after each one
  6. Result: where each application goes
  7. Possible saving: monthly and annual, with and without discount
An eighth slide (effort and migration estimate) is the architecture team's and is not built here.

Runtime counts only take resources that existed in the measured month (existed_at_measurement in the inventory):
run the scripts after the month ends and the deck shows that month, whatever day it is built.

Output: aks_aca_corporate.pptx in the newest reports/<date>-aks-aca folder (or --output). Every figure is a draft to check against
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
from pptx.enum.shapes import MSO_SHAPE
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
# microservice of the ADT_019 model (0.5 vCPU / 1 GiB warm in office hours, 5 million requests a month).
MIN_AKS_NODES, MIN_AKS_VM, HOURS_MONTH = 4, "Standard_D8s_v4", 730
# The bank's discount on Azure list prices for the minimum comparison (the measured node discount includes
# reservations that a new, small cluster would not have).
BANK_DISCOUNT = 0.30
ACA_VCPU, ACA_GIB, ACA_REQUESTS_MILLION = 0.5, 1.0, 5


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
    summary_path = os.path.join(results, "aa-cost", "cost_summary.json")
    summary = {}
    if os.path.exists(summary_path):
        with open(summary_path, encoding="utf-8") as handle:
            summary = json.load(handle)
    web_scope = summary.get("app_service_environments", [])
    web_measured = web_scope == "ALL" or "PROD" in web_scope
    return {"month": month, "by_category": by_category, "total_clusters": len(clusters),
            "cpu_used": cpu_used, "cpu_total": cpu_total, "mem_used": mem_used, "mem_total": mem_total,
            "prod": prod, "prod_cost": prod_cost, "web_measured": web_measured}


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
    slide = d.base_slide(deck, "Cuatro formas de construir una aplicación, cada una con su runtime", source)
    width, gap = Inches(2.95), Inches(0.12)
    cards = [("Monolito modular", draw_monolith,
              ["Una sola aplicación, ordenada en módulos por negocio.", "Se despliega y escala completa."],
              "App Service", d.SKY),
             ("Miniservicios", lambda sl, le, to, wi, ac: draw_pieces(sl, le, to, wi, ac, 2, 1, Inches(0.95)),
              ["Pocas piezas medianas: cada una agrupa funciones de un mismo negocio."],
              "Container Apps o AKS", d.GREEN),
             ("Microservicios", lambda sl, le, to, wi, ac: draw_pieces(sl, le, to, wi, ac, 4, 3, Inches(0.33)),
              ["Muchas piezas pequeñas: cada una se despliega y escala sola."],
              "AKS o Container Apps, según la carga", d.ORANGE),
             ("Procesos asíncronos", draw_async,
              ["Atienden mensajes o eventos que esperan en una cola: nadie espera la respuesta."],
              "Functions, Container Apps o AKS", d.STEEL)]
    for index, (title, drawing, lines, runtime, accent) in enumerate(cards):
        type_card(slide, Inches(0.6) + index * (width + gap), width, title, drawing, lines, runtime, accent)
    microservices = sum(int(row["microservices"] or 0) for row in apps)
    d.callout(slide, Inches(0.6), Inches(5.9), Inches(12.15), Inches(0.75),
              f"Hoy casi todo corre en AKS: {d.number(len(apps))} aplicaciones y {d.number(microservices)} "
              "microservicios, en el runtime pensado para carga alta y constante.", size=15, fill=d.RUST)


def draw_aks(slide, left, top, width, accent):
    """A cluster: servers that are always on, with the services packed inside."""
    outer = Inches(2.3)
    start = left + (width - outer) / 2
    box(slide, start, top, outer, Inches(1.0), d.ROYAL)
    for index in range(4):
        box(slide, start + Inches(0.12) + index * Inches(0.54), top + Inches(0.15), Inches(0.46), Inches(0.7), accent,
            radius=0.15)
        for slot in range(2):
            box(slide, start + Inches(0.2) + index * Inches(0.54), top + Inches(0.25) + slot * Inches(0.28),
                Inches(0.3), Inches(0.2), d.WHITE, radius=0.3)


def draw_aca(slide, left, top, width, accent):
    """Copies that turn on with use and off without it."""
    outer = Inches(2.3)
    start = left + (width - outer) / 2
    box(slide, start, top, outer, Inches(1.0), d.ROYAL)
    for index, fill in enumerate((accent, accent, d.GREY, d.GREY)):
        shape(slide, MSO_SHAPE.OVAL, start + Inches(0.15) + index * Inches(0.53), top + Inches(0.25), Inches(0.45),
              Inches(0.45), fill)


def draw_app_service(slide, left, top, width, accent):
    """A plan: fixed servers shared by several web apps."""
    outer = Inches(2.3)
    start = left + (width - outer) / 2
    box(slide, start, top, outer, Inches(1.0), d.ROYAL)
    for index in range(2):
        box(slide, start + Inches(0.2) + index * Inches(1.05), top + Inches(0.15), Inches(0.9), Inches(0.7), accent,
            radius=0.15)
        box(slide, start + Inches(0.32) + index * Inches(1.05), top + Inches(0.32), Inches(0.66), Inches(0.36),
            d.WHITE, radius=0.3)


def draw_functions(slide, left, top, width, accent):
    """Code that runs only when an event arrives."""
    start = left + width / 2 - Inches(0.9)
    shape(slide, MSO_SHAPE.LIGHTNING_BOLT, start, top + Inches(0.1), Inches(0.6), Inches(0.8), d.ORANGE)
    shape(slide, MSO_SHAPE.RIGHT_ARROW, start + Inches(0.7), top + Inches(0.35), Inches(0.35), Inches(0.3), d.SOFT)
    box(slide, start + Inches(1.15), top + Inches(0.2), Inches(0.65), Inches(0.6), accent, radius=0.2)


def slide_estate(deck, estate, apps, source):
    month = month_label(estate["month"])
    category = estate["by_category"]
    prod = estate["prod"]
    app = category.get(aa.APPLICATION, {"clusters": 0, "cost": 0.0, "microservices": 0})

    def web_cost(owner_type):
        return f"{soles(estate['prod_cost'][owner_type])} al mes" if estate["web_measured"] else "Costo no medido"
    slide = d.base_slide(deck, f"Producción: {prod['aks']} clusters AKS concentran los microservicios", source)
    cards = [("AKS", draw_aks, "Servidores encendidos 24/7", prod["aks"], "clusters",
              [f"{d.number(app['microservices'])} microservicios", f"{soles(estate['prod_cost']['AKS'])} al mes"],
              d.ORANGE),
             ("Container Apps", draw_aca, "Copias que se prenden y apagan con el uso", prod["aca"], "apps",
              [f"{d.number(prod['aca_environments'])} environments", f"{soles(estate['prod_cost']['ACA'])} al mes"],
              d.GREEN),
             ("App Service", draw_app_service, "Servidores fijos que comparten apps web", prod["web"], "web apps",
              [f"{d.number(prod['plans']['APP_SERVICE'])} planes", web_cost("APP_SERVICE")], d.SKY),
             ("Azure Functions", draw_functions, "Código que corre solo cuando llega un evento", prod["functions"],
              "function apps", [f"{d.number(prod['plans']['FUNCTIONS'])} planes", web_cost("FUNCTIONS")], d.SKY)]
    for index, (name, drawing, meaning, count, unit, lines, accent) in enumerate(cards):
        left = Inches(0.6 + index * 3.07)
        d.rounded(slide, left, Inches(1.3), Inches(2.9), Inches(3.55), d.CARD)
        d.text_box(slide, left, Inches(1.38), Inches(2.9), Inches(0.4), name, size=14, bold=True, color=d.WHITE,
                   align=d.PP_ALIGN.CENTER)
        drawing(slide, left, Inches(1.85), Inches(2.9), accent)
        d.text_box(slide, left + Inches(0.1), Inches(2.9), Inches(2.7), Inches(0.35), meaning, size=10, color=d.SOFT,
                   align=d.PP_ALIGN.CENTER)
        d.text_box(slide, left, Inches(3.2), Inches(2.9), Inches(0.6), d.number(count), size=30, bold=True,
                   color=accent, align=d.PP_ALIGN.CENTER)
        d.text_box(slide, left, Inches(3.75), Inches(2.9), Inches(0.3), unit, size=11, color=d.SOFT,
                   align=d.PP_ALIGN.CENTER)
        d.text_box(slide, left + Inches(0.15), Inches(4.05), Inches(2.6), Inches(0.75), lines, size=11, bold=True,
                   color=d.WHITE, align=d.PP_ALIGN.CENTER)
    rows = [["Clusters AKS en producción", "Clusters", f"Costo {month}"]]
    for code, label in ((aa.APPLICATION, "Con aplicaciones (base del análisis)"),
                        (aa.NO_APPLICATIONS, "Creados sin aplicaciones"),
                        (aa.INFRASTRUCTURE, "Infraestructura (CI/CD, observabilidad, gateways)"),
                        (aa.NO_DATA, "Sin datos")):
        entry = category.get(code)
        if entry and entry["clusters"]:
            rows.append([label, d.number(entry["clusters"]), soles(entry["cost"])])
    d.table(slide, Inches(0.6), Inches(5.0), Inches(7.9), rows, [0.6, 0.15, 0.25], highlight_column=None,
            row_height=0.32, size=11)
    d.note_box(slide, Inches(8.85), Inches(5.0), Inches(3.85), Inches(1.3),
               "Microservicios = piezas de aplicación desplegadas en AKS.", size=11)


def minimum_costs(month, prices, ratio=1 - BANK_DISCOUNT, region="eastus2"):
    """(smallest AKS cluster, one Container Apps microservice) per month in US$, both with the bank's discount."""
    price = lambda key: prices.get((key, region), 0)
    aks = MIN_AKS_NODES * HOURS_MONTH * price(f"vm_{MIN_AKS_VM}_payg") * ratio
    aca = (office_seconds(month) * (ACA_VCPU * price("aca_vcpu_active_s") + ACA_GIB * price("aca_mem_active_gib_s"))
           + ACA_REQUESTS_MILLION * price("aca_requests_million")) * ratio
    return aks, aca


def slide_hypothesis(deck, estate, apps, scaling, results, source):
    month = estate["month"]
    prices = read_prices(results)
    aks_min, aca_min = minimum_costs(month, prices)
    small = sum(0 < int(row["microservices"] or 0) <= 15 for row in apps)
    slide = d.base_slide(deck, "Hipótesis: un cluster AKS por aplicación paga más de lo que usa", source)
    d.heading(slide, Inches(0.6), Inches(1.3), Inches(7.2), "Indicios", size=14)
    tiles = [(d.percent(d.share_of(estate["cpu_used"], estate["cpu_total"])),
              "de la CPU pagada se usa en horario de oficina"),
             (d.percent(d.share_of(estate["mem_used"], estate["mem_total"])),
              "de la memoria pagada se usa en horario de oficina"),
             (d.percent(d.share_of(scaling["over_3_months"], scaling["dated"])) if scaling["dated"] else "-",
              "de los autoescalados no cambia de tamaño hace más de 3 meses"),
             (f"{d.number(small)} de {d.number(len(apps))}",
              "aplicaciones tienen 15 microservicios o menos y aun así un cluster propio")]
    for index, (value, label) in enumerate(tiles):
        d.kpi(slide, Inches(0.6) + (index % 2) * Inches(3.65), Inches(1.75) + (index // 2) * Inches(1.45), Inches(3.5),
              Inches(1.3), value, label, value_size=26, label_size=11)
    d.heading(slide, Inches(8.1), Inches(1.3), Inches(4.6), "Motivación", size=14)
    d.callout(slide, Inches(8.1), Inches(1.75), Inches(4.6), Inches(1.2),
              "Económica: pagar por uso, no por servidores encendidos 24/7", size=15, fill=d.RUST)
    for index, text in enumerate(("Operación: menos clusters que actualizar y vigilar",
                                  "Complejidad: sin servidores ni versiones de Kubernetes que administrar")):
        d.callout(slide, Inches(8.1), Inches(3.05) + index * Inches(0.75), Inches(4.6), Inches(0.65), text, size=12,
                  fill=d.CARD)
    d.heading(slide, Inches(0.6), Inches(4.75), Inches(12.1), "Lo mínimo que se paga al mes (con el 30 % de descuento "
              "del banco)", size=14)
    d.kpi(slide, Inches(0.6), Inches(5.15), Inches(5.9), Inches(1.3), soles(aks_min),
          "el cluster AKS más chico (4 servidores), aunque la aplicación sea pequeña", value_color=d.RUST,
          value_size=28, label_size=11)
    d.kpi(slide, Inches(6.8), Inches(5.15), Inches(5.9), Inches(1.3), soles(aca_min),
          "un microservicio en Container Apps, encendido en horario de oficina", value_color=d.GREEN,
          value_size=28, label_size=11)


def apps_label(count):
    return f"{d.number(count)} {'aplicación' if count == 1 else 'aplicaciones'}"


def expected_microservices(row):
    """Microservices in PROD, or the projection with DESA when larger (as the ADT_019 rule in the snapshot)."""
    return max(int(row["microservices"] or 0), int(to_float(row.get("microservices_projected")) or 0))


# The ADT_019 tree read the other way round: an application goes to Container Apps when it meets all four
# conditions (the same as meeting none of the AKS ones). (short label, condition, plain meaning, test on a row).
ACA_CONDITIONS = [
    ("Tráfico moderado", "Tráfico moderado", "Menos de 100 pedidos por segundo en su hora más cargada.",
     lambda row: (to_float(row["busiest_hour_tps"]) or 0) <= 100),
    ("Pocas piezas o por horas", "Pocas piezas o uso por horas",
     "Hasta 15 microservicios (contando los que vienen de desarrollo), o menos de 10 horas de uso al día.",
     lambda row: not ((to_float(row["weekday_hours_with_traffic"]) or 0) > 10 and expected_microservices(row) > 15)),
    ("Se apaga de noche", "Se puede apagar de noche", "Casi no se usa de noche ni los fines de semana.",
     lambda row: (to_float(row.get("offhours_active_share")) or 0) <= 0.5),
    ("Piezas livianas", "Piezas livianas", "Cada pieza necesita hasta 4 vCPU y 8 GB de memoria.",
     lambda row: int(row["microservices_over_aca_size"] or 0) == 0),
]
CHECKS_BEFORE_MOVING = ["Si es crítica para el banco, se valida con gobierno de arquitectura.",
                        "Si va a pasar de 15 microservicios, se confirma que siga cumpliendo."]


def slide_criteria(deck, source):
    slide = d.base_slide(deck, "Para pasar a Container Apps, una aplicación cumple estas cuatro condiciones", source)
    width, height, gap = Inches(5.9), Inches(1.45), Inches(0.3)
    for index, (_, title, meaning, _) in enumerate(ACA_CONDITIONS):
        left = Inches(0.6) + (index % 2) * (width + gap)
        top = Inches(1.4) + (index // 2) * (height + gap)
        box(slide, left, top, width, height, d.LIGHT)
        shape(slide, MSO_SHAPE.OVAL, left + Inches(0.3), top + Inches(0.4), Inches(0.6), Inches(0.6), d.GREEN)
        d.text_box(slide, left + Inches(0.3), top + Inches(0.4), Inches(0.6), Inches(0.6), str(index + 1), size=18,
                   bold=True, color=d.WHITE, align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
        d.text_box(slide, left + Inches(1.15), top + Inches(0.25), width - Inches(1.4), Inches(0.45), title, size=18,
                   bold=True, color=d.NAVY)
        d.text_box(slide, left + Inches(1.15), top + Inches(0.75), width - Inches(1.4), Inches(0.5), meaning, size=14,
                   color=d.GREY)
    top = Inches(1.4) + 2 * (height + gap)
    box(slide, Inches(0.6), top, Inches(12.1), Inches(1.45), d.LIGHT)
    d.text_box(slide, Inches(0.9), top + Inches(0.15), Inches(11.5), Inches(0.4), "Antes de mover, se valida:", size=16,
               bold=True, color=d.ORANGE)
    d.text_box(slide, Inches(0.9), top + Inches(0.55), Inches(11.5), Inches(0.8), [f"• {line}" for line in
                                                                                    CHECKS_BEFORE_MOVING],
               size=14, color=d.GREY)
    d.text_box(slide, Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.4),
               "Si no cumple alguna de las cuatro, se queda en AKS.", size=14, bold=True, color=d.NAVY)


def slide_criteria_applied(deck, apps, source):
    """A funnel: applications with measured traffic, then those that still qualify after each condition."""
    remaining = [row for row in apps if row["busiest_hour_tps"] != ""]
    labels, values = ["Con tráfico medido"], [len(remaining)]
    for label, _, _, test in ACA_CONDITIONS:
        remaining = [row for row in remaining if test(row)]
        labels.append(f"+ {label.lower()}")
        values.append(len(remaining))
    review = sum(row["adt_runtime"] == "ACA_WITH_REVIEW" for row in remaining)
    slide = d.base_slide(deck, f"Al aplicar las condiciones, {len(remaining)} aplicaciones pueden pasar a "
                               "Container Apps", source)
    colours = [d.GREY] + [d.STEEL] * (len(values) - 2) + [d.GREEN]
    d.bar_chart(slide, Inches(0.6), Inches(1.4), Inches(8.4), Inches(4.9), labels[::-1], values[::-1],
                XL_CHART_TYPE.BAR_CLUSTERED, colours[::-1])
    lines = [f"De {apps_label(len(apps))}, {d.number(values[0])} tienen tráfico medido; las demás se ven aparte.", "",
             "Cada barra suma una condición: quedan las que cumplen todas las anteriores."]
    if review:
        lines += ["", f"{d.number(review)} de las {d.number(len(remaining))} se validan antes de mover (críticas o en "
                      "crecimiento)."]
    d.text_box(slide, Inches(9.3), Inches(1.6), Inches(3.4), Inches(4.0), lines, size=14, color=d.GREY)


# Groups of the distribution slide, in plain words for a non-technical audience: (label, what it means, colour).
DISTRIBUTION_GROUPS = {
    "AKS": ("Se queda en AKS", "Mucho tráfico o uso 24/7", d.NAVY),
    "ACA_WITH_REVIEW": ("Container Apps, con revisión", "Poco tráfico, pero crítica o creciendo", d.ORANGE),
    "ACA": ("Candidata a Container Apps", "Poco tráfico, solo en horario de oficina", d.GREEN),
    "IN_USE": ("En uso, sin medición", "Tiene uso, pero su tráfico no se ve", d.STEEL),
    "NOT_IN_USE": ("Sin uso", "Por revisar", d.GREY),
}


def distribution_group(row):
    """Slide group of an application; None for one still in development (shown in the note, not in the chart)."""
    if row["adt_runtime"] != "INSUFFICIENT_DATA":
        return row["adt_runtime"]
    reason = row.get("no_traffic_reason")
    if reason == "IN_DEVELOPMENT":
        return None
    return "IN_USE" if reason == "IN_USE_NOT_MEASURED" else "NOT_IN_USE"


def slide_distribution(deck, apps, estate, source):
    month = month_label(estate["month"])
    groups = defaultdict(lambda: {"apps": [], "cost": 0.0})
    developing = []
    for row in apps:
        key = distribution_group(row)
        if key is None:
            developing.append(row["cod_app"])
            continue
        groups[key]["apps"].append(row)
        groups[key]["cost"] += to_float(row["cost_month_usd"]) or 0
    movable = len(groups["ACA"]["apps"]) + len(groups["ACA_WITH_REVIEW"]["apps"])
    slide = d.base_slide(deck, f"Resultado: {movable} de {len(apps)} aplicaciones podrían ir a Container Apps", source)
    order = [key for key in DISTRIBUTION_GROUPS if groups[key]["apps"]]
    d.bar_chart(slide, Inches(0.5), Inches(1.4), Inches(5.6), Inches(4.4), [DISTRIBUTION_GROUPS[key][0] for key in order],
                [len(groups[key]["apps"]) for key in order], XL_CHART_TYPE.COLUMN_CLUSTERED,
                [DISTRIBUTION_GROUPS[key][2] for key in order])
    rows = [["Destino", "Apps", f"Costo AKS al mes\n({month}, con IGV)"]]
    for key in order:
        label, meaning, _ = DISTRIBUTION_GROUPS[key]
        rows.append([f"{label}\n{meaning}", d.number(len(groups[key]["apps"])), soles(groups[key]["cost"])])
    d.table(slide, Inches(6.3), Inches(1.5), Inches(6.6), rows, [0.6, 0.13, 0.27], highlight_column=2, row_height=0.62,
            size=10)
    notes = []
    single = sum(row["possible_monolith"] == "YES" for row in apps)
    if single:
        notes.append(f"• {d.number(single)} aplicaciones tienen un solo microservicio: podrían ir a App Service.")
    if developing:
        notes.append(f"• {d.number(len(developing))} aplicaciones están en desarrollo y aún no salen a producción.")
    if notes:
        box = d.rounded(slide, Inches(6.3), Inches(4.75), Inches(6.6), Inches(0.8), d.LIGHT, radius=0.1)
        box.line.color.rgb = d.SOFT
        d.text_box(slide, Inches(6.5), Inches(4.75), Inches(6.2), Inches(0.8), notes, size=11,
                   anchor=d.MSO_ANCHOR.MIDDLE)


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


def aca_estimate(results, apps, month, prices, region="eastus2"):
    """Monthly Container Apps cost per app code at list price, with the ADT_019 model: the replicas Container
    Insights measured in office hours, warm from 8:00 to 20:00 on weekdays and at zero otherwise; each replica sized
    to the smallest Consumption allocation that covers its request; plus requests and one environment per app."""
    price = lambda key: prices.get((key, region), 0)
    seconds = office_seconds(month)
    per_replica = defaultdict(float)
    application = {row["cluster_id"]: row["cod_app"]
                   for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                   if row["category"] == aa.APPLICATION and row["environment"] == "PROD"}
    for row in results_csv(results, "aa-k8s-snapshot", "deployment_snapshot.csv"):
        code = application.get(row["cluster_id"])
        if not code or row["microservice"] != "YES":
            continue
        cpu, mem = to_float(row["cpu_request_vcpu"]) or 0, to_float(row["mem_request_gib"]) or 0
        vcpu = math.ceil(max(cpu, mem / 2, 0.25) * 4 - 1e-9) / 4
        replicas = max(1.0, to_float(row["ci_replicas_avg_office"]) or to_float(row["hpa_min"]) or 1)
        per_replica[code] += replicas * seconds * (vcpu * price("aca_vcpu_active_s") + 2 * vcpu * price("aca_mem_active_gib_s"))
    environment = 730 * (price("aca_env_management_hour") + price("aca_env_private_endpoint_hour"))
    return {row["cod_app"]: per_replica.get(row["cod_app"], 0) + environment
            + (to_float(row["requests_month"]) or 0) / 1e6 * price("aca_requests_million") for row in apps}


def slide_cost(deck, apps, estate, results, source):
    month = estate["month"]
    prices = read_prices(results)
    ratio = measured_discount(results, month, prices)
    aca = aca_estimate(results, apps, month, prices)
    movable = [row for row in apps if row["adt_runtime"] in ("ACA", "ACA_WITH_REVIEW")]
    aks_paid = sum(to_float(row["cost_month_usd"]) or 0 for row in movable)
    aca_list = sum(aca[row["cod_app"]] for row in movable)
    scenarios = [("Precio de lista (sin descuento)", aks_paid / ratio, aca_list),
                 ("Con el descuento del banco", aks_paid, aca_list * ratio)]
    saving = scenarios[1][1] - scenarios[1][2]
    slide = d.base_slide(deck, f"Mover las {len(movable)} aplicaciones candidatas ahorraría ~{soles(saving)} al mes "
                               f"({soles(saving * 12)} al año)", source)
    rows = [["Escenario (con IGV)", "AKS hoy / mes", "Container Apps / mes", "Ahorro / mes", "Ahorro / año"]]
    for label, aks_cost, aca_cost in scenarios:
        rows.append([label, soles(aks_cost), soles(aca_cost), soles(aks_cost - aca_cost),
                     soles((aks_cost - aca_cost) * 12)])
    d.table(slide, Inches(0.6), Inches(1.4), Inches(12.1), rows, [0.3, 0.17, 0.19, 0.17, 0.17], highlight_column=3,
            row_height=0.5, size=13)
    rows = [["Cómo se paga", "AKS", "Container Apps"],
            ["Cobro", "Nodos completos (VM) las 24 horas", "vCPU-segundo y GB-segundo en uso, y peticiones"],
            ["Fuera de horario", "Se paga igual", "Escala a cero (ADT_019: instancias calientes de 8 a 20 h)"],
            ["Descuento", f"Medido: se paga {d.percent(ratio)} del precio de lista de las VM",
             "Se asume el mismo porcentaje"]]
    d.table(slide, Inches(0.6), Inches(3.25), Inches(12.1), rows, [0.2, 0.38, 0.42], row_height=0.45, size=12)
    d.note_box(slide, Inches(0.6), Inches(5.35), Inches(12.1), Inches(0.65),
               f"Estimado con las réplicas medidas en {month_label(month)}, cada una del tamaño que pide hoy. Incluye las "
               f"aplicaciones \"con revisión\": el ahorro firme es menor. Detalle por cluster en el workbook.", size=11)
    d.note_box(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(0.6),
               "Beneficio adicional sin monto: menos clusters que actualizar y mantener (horas-hombre de los equipos).",
               size=11)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: aks_aca_corporate.pptx in the newest reports/<date>-aks-aca folder")
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
    criteria = (f"Fuente: lineamiento ADT_019; tráfico de APIM, Application Gateway y Event Hubs, Container Insights y configuración de autoescalado, "
                f"{period}.")

    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    slide_context(deck, apps, "Fuente: lineamientos de arquitectura del banco (FAT_001, ADT_019).")
    slide_estate(deck, estate, apps, f"{measured} {MONEY_NOTE}")
    slide_hypothesis(deck, estate, apps, scaling, args.results,
                     f"{measured} Precios de lista: Azure Retail Prices API. {MONEY_NOTE}")
    slide_criteria(deck, "Fuente: lineamiento ADT_019 del banco.")
    slide_criteria_applied(deck, apps, criteria)
    slide_distribution(deck, apps, estate, f"{criteria} {MONEY_NOTE}")
    slide_cost(deck, apps, estate, args.results,
               f"{measured} Precios de lista: Azure Retail Prices API. {MONEY_NOTE}")
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, "aks_aca_corporate.pptx")
    deck.save(output)
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


if __name__ == "__main__":
    main()
