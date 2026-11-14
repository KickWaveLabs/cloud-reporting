#!/usr/bin/env python3
"""Build the corporate AKS vs Container Apps deck (slides 1 to 5) for a non-technical forum.

Same look as aa_deck.py (its drawing helpers are reused), plain figures only: TPS, vCPU, GB of RAM, instances,
hours with traffic, USD per month and per year. Slide text is in Spanish for the audience.

Reads only results/ (inventory, metrics, cost, prices, snapshot), so it needs no saved workbook. The cost slide
uses a simplified ADT_019 model (see aca_estimate); the workbook's ACA_by_cluster sheet has the detail per
cluster. No Azure call is made; nothing leaves the machine.

Slides:
  1. Context: workloads and the runtime for each
  2. The runtime estate today (only APPLICATION clusters count; the rest are named as excluded)
  3. Criteria to sort the estate (ADT_019), each with the PROD figure
  4. Theoretical distribution after applying the criteria
  5. Cost optimization: monthly and annual, with and without discount

Output: aks_aca_corporate.pptx in the newest reports/<date>-aks-aca folder (or --output). Every figure is a draft to check against
the workbook and the CSVs before presenting.
"""

import argparse
import glob
import math
import os
from collections import Counter, defaultdict
from datetime import date

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

import aa_common as aa
import aa_deck as d

MONTHS_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
             "noviembre", "diciembre")
RUNTIME_LABELS = {"AKS": "Se queda en AKS", "ACA": "Candidata a Container Apps",
                  "ACA_WITH_REVIEW": "Container Apps, con revisión"}
RUNTIME_COLORS = {"AKS": d.NAVY, "ACA": d.GREEN, "ACA_WITH_REVIEW": d.ORANGE}


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


def read_estate(results):
    """Figures of the runtime estate in PROD, by cluster category."""
    clusters = [row for row in results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["in_inventory"] == "YES"]
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
    aca = [row for row in results_csv(results, "aa-inventory", "aca_apps.csv") if row["environment"] == "PROD"]
    sites = [row for row in results_csv(results, "aa-inventory", "app_services.csv") if row["environment"] == "PROD"]
    return {"month": month, "by_category": by_category, "total_clusters": len(clusters),
            "cpu_used": cpu_used, "cpu_total": cpu_total, "mem_used": mem_used, "mem_total": mem_total,
            "aca_apps": len(aca), "web_apps": sum(row.get("type") == "WEB" for row in sites),
            "function_apps": sum(row.get("type") == "FUNCTION" for row in sites)}


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


def slide_context(deck, source):
    slide = d.base_slide(deck, "Cada tipo de carga tiene un runtime: hoy casi todo corre en AKS", source)
    rows = [["Runtime", "Para qué carga", "Cómo se paga"],
            ["AKS (Kubernetes)", "Mucha carga constante, muchos microservicios, plataforma compartida",
             "Nodos encendidos 24/7, se usen o no"],
            ["Azure Container Apps", "Microservicios con carga variable o solo en horario de oficina",
             "Por segundo de uso; escala a cero fuera de horario"],
            ["App Service", "Aplicaciones monolíticas internas", "Plan fijo por instancia"],
            ["Azure Functions", "Lógica que responde a eventos de Azure", "Por ejecución"]]
    d.table(slide, Inches(0.6), Inches(1.45), Inches(12.1), rows, [0.22, 0.46, 0.32], row_height=0.62, size=13)
    d.callout(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(0.95),
              "El lineamiento de arquitectura (ADT_019) define qué runtime corresponde según carga, uso y tamaño.",
              size=16)
    d.note_box(slide, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.7),
               "El mismo lineamiento pide re-evaluar la plataforma cuando el uso real cambia: eso es lo que mide esta "
               "propuesta.", size=11)


def slide_estate(deck, estate, apps, source):
    month = month_label(estate["month"])
    category = estate["by_category"]
    app = category.get(aa.APPLICATION, {"clusters": 0, "cost": 0.0, "microservices": 0})
    slide = d.base_slide(deck, f"Hoy: {app['clusters']} clusters AKS de aplicación cuestan "
                               f"{d.money(app['cost'])} al mes", source)
    figures = [(d.number(app["clusters"]), "clusters AKS de aplicación"),
               (d.number(len(apps)), "aplicaciones"),
               (d.number(app["microservices"]), "microservicios"),
               (d.money(app["cost"] * 12), "costo anual estimado"),
               (d.percent(d.share_of(estate["cpu_used"], estate["cpu_total"])), "de la CPU usada en horario de oficina"),
               (d.percent(d.share_of(estate["mem_used"], estate["mem_total"])), "de la RAM usada en horario de oficina")]
    for index, (value, label) in enumerate(figures):
        left = Inches(0.6 + (index % 3) * 4.1)
        top = Inches(1.4 + (index // 3) * 1.65)
        d.kpi(slide, left, top, Inches(3.8), Inches(1.45), value, label, value_size=28)
    rows = [["Otros clusters en PROD (fuera del análisis)", "Clusters", f"Costo {month}"]]
    for code, label in ((aa.NO_APPLICATIONS, "Creados sin aplicaciones"),
                        (aa.INFRASTRUCTURE, "Infraestructura (CI/CD, observabilidad, gateways)"),
                        (aa.NO_DATA, "Sin datos")):
        entry = category.get(code)
        if entry and entry["clusters"]:
            rows.append([label, d.number(entry["clusters"]), d.money(entry["cost"])])
    d.table(slide, Inches(0.6), Inches(4.85), Inches(7.6), rows, [0.6, 0.15, 0.25], row_height=0.36, size=11)
    d.note_box(slide, Inches(8.5), Inches(4.85), Inches(4.2), Inches(1.45),
               f"Otros runtimes en PROD: {estate['aca_apps']} Container Apps, {estate['web_apps']} App Services web y "
               f"{estate['function_apps']} Functions.", size=11)


def adt_counts(apps):
    with_traffic = [row for row in apps if row["busiest_hour_tps"] != ""]
    return {
        "with_traffic": len(with_traffic),
        "over_tps": sum(to_float(row["busiest_hour_tps"]) > 100 for row in with_traffic),
        "constant_large": sum((to_float(row["weekday_hours_with_traffic"]) or 0) > 10
                              and int(row["microservices"]) > 15 for row in with_traffic),
        "over_size": sum(int(row["microservices_over_aca_size"]) > 0 for row in apps),
        "batch_state": sum(int(row["batch_jobs"]) > 0 or int(row["statefulsets"]) > 0 for row in apps),
        "critical": sum(row["bia"] in ("MUY ALTA", "ALTA") for row in apps),
    }


def slide_criteria(deck, apps, scaling, source):
    counts = adt_counts(apps)
    total = len(apps)
    not_scaling = scaling["never"] + scaling["fixed"]
    slide = d.base_slide(deck, "Seis criterios ordenan el parque; cada uno con su cifra de producción", source)
    rows = [["Criterio", "Va a AKS si…", "Producción hoy"],
            ["Carga (TPS en la hora pico)", "Más de 100 TPS",
             f"{d.number(counts['over_tps'])} de {d.number(counts['with_traffic'])} aplicaciones con tráfico medido"],
            ["Uso constante y tamaño", "Más de 10 h de tráfico al día y más de 15 microservicios",
             f"{d.number(counts['constant_large'])} de {d.number(total)} aplicaciones"],
            ["Tamaño por instancia", "Más de 4 vCPU / 8 GB de RAM",
             f"{d.number(counts['over_size'])} aplicaciones con algún microservicio más grande"],
            ["¿Escala con la demanda?", "Escalar a cero es la ventaja de Container Apps",
             f"{d.percent(d.share_of(not_scaling, scaling['known']))} de los autoescalados no escala "
             f"({d.percent(d.share_of(scaling['never'], scaling['known']))} en su mínimo, "
             f"{d.percent(d.share_of(scaling['fixed'], scaling['known']))} con mínimo = máximo)"],
            ["Procesos batch o con estado", "Batch largo o StatefulSets: revisar con arquitectura",
             f"{d.number(counts['batch_state'])} aplicaciones"],
            ["Criticidad del negocio", "BIA ALTA o MUY ALTA: revisar con gobierno de arquitectura",
             f"{d.number(counts['critical'])} aplicaciones"]]
    d.table(slide, Inches(0.6), Inches(1.4), Inches(12.1), rows, [0.24, 0.36, 0.40], row_height=0.6, size=12)
    if scaling["dated"]:
        d.note_box(slide, Inches(0.6), Inches(5.75), Inches(12.1), Inches(0.75),
                   f"{d.percent(d.share_of(scaling['over_3_months'], scaling['dated']))} de los autoescalados no escala "
                   f"hace más de 3 meses y {d.percent(d.share_of(scaling['over_1_year'], scaling['dated']))} hace más "
                   f"de un año. Container Apps arranca en frío en ~30 s: los consumidores deben reintentar.", size=11)


def slide_distribution(deck, apps, estate, source):
    month = month_label(estate["month"])
    groups = defaultdict(lambda: {"apps": 0, "cost": 0.0, "microservices": 0})
    for row in apps:
        entry = groups[row["adt_runtime"]]
        entry["apps"] += 1
        entry["cost"] += to_float(row["cost_month_usd"]) or 0
        entry["microservices"] += int(row["microservices"] or 0)
    movable = groups["ACA"]["apps"] + groups["ACA_WITH_REVIEW"]["apps"]
    slide = d.base_slide(deck, f"Aplicando los criterios, {movable} de {len(apps)} aplicaciones podrían ir a "
                               f"Container Apps", source)
    order = [key for key in ("AKS", "ACA_WITH_REVIEW", "ACA") if groups[key]["apps"]]
    d.bar_chart(slide, Inches(0.5), Inches(1.4), Inches(6.2), Inches(4.4), [RUNTIME_LABELS[key] for key in order],
                [groups[key]["apps"] for key in order], XL_CHART_TYPE.COLUMN_CLUSTERED,
                [RUNTIME_COLORS[key] for key in order])
    rows = [["Destino", "Apps", "Micro-\nservicios", f"Costo AKS\n{month}"]]
    for key in order:
        rows.append([RUNTIME_LABELS[key], d.number(groups[key]["apps"]), d.number(groups[key]["microservices"]),
                     d.money(groups[key]["cost"])])
    d.table(slide, Inches(7.0), Inches(1.5), Inches(5.8), rows, [0.4, 0.14, 0.2, 0.26], row_height=0.5, size=11)
    monoliths = sum(row["possible_monolith"] == "YES" for row in apps)
    d.note_box(slide, Inches(7.0), Inches(3.6), Inches(5.8), Inches(1.0),
               f"{d.number(monoliths)} aplicaciones tienen un solo microservicio: si son monolitos, el lineamiento las "
               f"llevaría a App Service. Se confirma con cada equipo.", size=11)
    d.note_box(slide, Inches(7.0), Inches(4.8), Inches(5.8), Inches(1.0),
               "\"Con revisión\": batch, estado, criticidad alta o tráfico sin medir. Cada caso se valida con "
               "arquitectura antes de migrar.", size=11)


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
    slide = d.base_slide(deck, f"Mover las {len(movable)} aplicaciones candidatas ahorraría ~{d.money(saving)} al mes "
                               f"({d.money(saving * 12)} al año)", source)
    rows = [["Escenario", "AKS hoy / mes", "Container Apps / mes", "Ahorro / mes", "Ahorro / año"]]
    for label, aks_cost, aca_cost in scenarios:
        rows.append([label, d.money(aks_cost), d.money(aca_cost), d.money(aks_cost - aca_cost),
                     d.money((aks_cost - aca_cost) * 12)])
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
    period = month_label(estate["month"])
    measured = f"Fuente: inventario de Azure, Cost Management (costo amortizado, {period}) y métricas de {period}."
    criteria = (f"Fuente: lineamiento ADT_019; tráfico de APIM, Container Insights y configuración de autoescalado, "
                f"{period}.")

    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    slide_context(deck, "Fuente: lineamiento ADT_019 (runtime de ejecución).")
    slide_estate(deck, estate, apps, measured)
    slide_criteria(deck, apps, scaling, criteria)
    slide_distribution(deck, apps, estate, criteria)
    slide_cost(deck, apps, estate, args.results, measured + " Precios de lista: Azure Retail Prices API.")
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, "aks_aca_corporate.pptx")
    deck.save(output)
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


if __name__ == "__main__":
    main()
