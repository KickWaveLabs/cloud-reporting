#!/usr/bin/env python3
"""Build the corporate deck of criteria v1, v2 or v3 (--criteria; v2 by default).

v1 runs aa_deck_exec_v1.py (per application, frozen at tag v1). v2 and v3 come from the criteria model
(aa_v2_model.py; v3 is v2 with its PRESETS["v3"] defaults): the decision per cluster and the Container Apps cost from
each hour's traffic.

Slides 1 and 2 are v1's (aa_deck_exec_v1.py). Then:
  3  hypothesis, with the real example costed by the v2 model
  4  criteria per cluster: four conditions, the rules that keep a cluster on AKS, the saving margin, how it is costed
  5  result per cluster (groups of the v2 workbook's Results sheet)
  6  saving: what moving the clusters of the ACA groups saves a month and a year, against their cost and the total
  7  operating cost of moving: what each application must meet (Quarkus, pipelines, no data in the cluster, network)
Reads only results/ (run aa_cluster_snapshot.py first). Output: aks_aca_corporate_<criteria>.pptx in the newest
reports/<date>-aks-aca folder (or --output); next to it, cluster_decisions_<criteria>.csv with every cluster's figures.
"""

import argparse
import glob
import math
import os
import re
import sys
from collections import defaultdict
from datetime import date

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

import aa_common as aa
import aa_deck as d
import aa_deck_exec_v1 as v1
import aa_v2_model as model

MICRO, MICRO_ONE = v1.MICRO, v1.MICRO_ONE


CRITERIA = {"name": "v2"}  # set by main (--criteria)


def sizing(where):
    """How replicas are sized, in the slides' words: v2's sentences as presented, v3's for its sizing."""
    per = model.DEFAULTS["concurrent_requests"]
    office = f"mínimo 1 por {MICRO_ONE} en horario de oficina y 0 fuera de él"
    band = ("Costo en Container Apps: las réplicas suben y bajan con el tráfico, y se apagan fuera del horario de "
            "oficina cuando no hay uso.")
    texts = {
        "v2": {"example": f"cada 5 minutos según el tráfico, 1 por cada {per} TPS; {office}.",
               "criteria": f"1 por cada {per} TPS, cada 5 minutos de la hora más cargada del mes (regla de escalado de "
                           f"Azure); {office}.",
               "cost": f"Según el tráfico cada 5 minutos: 1 por cada {per} TPS; se apagan fuera de horario sin uso",
               "band": band},
        "v3": {"example": f"cada 5 minutos con la capacidad por réplica medida hoy; {office}.",
               "criteria": f"Cada 5 minutos de la hora más cargada del mes, con lo que hoy atiende una réplica en AKS "
                           f"(al menos {per} TPS, regla de escalado de Azure); {office}.",
               "cost": "Cada 5 min según el tráfico y la capacidad medida; se apagan fuera de horario sin uso",
               "band": band}}
    return texts[CRITERIA["name"]][where]


def pen(value):
    """Soles already with IGV -> 'S/ 1.234'."""
    return f"S/ {value:,.0f}".replace(",", ".")


def slide_estate(deck, estate, apps, results, source):
    """(Costs in US$, shown in soles.) Left, the whole PROD universe and its split; right, the business application
    clusters, the base of the analysis, in large print; below, what stays out of the analysis and its monthly cost."""
    month = v1.month_label(estate["month"])
    category = estate["by_category"]
    application = category.get(aa.APPLICATION, {"clusters": 0, "cost": 0})
    total = estate["prod_cost"]["AKS"] + estate["nonprod_cost"]
    slide = d.base_slide(deck, f"Producción: {application['clusters']} clusters AKS de aplicaciones de negocio, la "
                               "base del análisis", source)
    left, top, width, height = Inches(0.6), Inches(1.3), Inches(3.4), Inches(4.75)
    d.rounded(slide, left, top, width, height, d.CARD)
    icon = Inches(0.8)
    slide.shapes.add_picture(os.path.join(v1.ICONS_DIR, "aks.png"), left + int((width - icon) / 2),
                             top + Inches(0.45), height=icon)
    d.text_box(slide, left, top + Inches(1.3), width, Inches(0.7), d.number(estate["prod"]["aks"]), size=40,
               bold=True, color=d.ORANGE, align=d.PP_ALIGN.CENTER)
    d.text_box(slide, left, top + Inches(2.0), width, Inches(0.3), "clusters AKS en producción", size=12,
               color=d.SOFT, align=d.PP_ALIGN.CENTER)
    split = [(aa.APPLICATION, "aplicaciones de negocio"), (aa.INFRASTRUCTURE, "infraestructura"),
             (aa.NO_APPLICATIONS, "vacíos, aún en desarrollo"), (aa.NO_DATA, "sin datos")]
    row = top + Inches(2.55)
    for code, label in split:
        entry = category.get(code)
        if not entry or not entry["clusters"]:
            continue
        base = code == aa.APPLICATION
        d.text_box(slide, left + Inches(0.3), row, Inches(0.75), Inches(0.35), d.number(entry["clusters"]),
                   size=16 if base else 13, bold=True, color=d.ORANGE if base else d.WHITE, align=d.PP_ALIGN.RIGHT)
        d.text_box(slide, left + Inches(1.15), row + Inches(0.03), width - Inches(1.3), Inches(0.35), label,
                   size=11, bold=base, color=d.WHITE if base else d.SOFT)
        row += Inches(0.42)
    grid_left, gap = Inches(4.25), Inches(0.2)
    card_width = int((Inches(12.75) - grid_left - gap) / 2)
    cards = [(MICRO, d.number(v1.deployed_microservices(apps)), "desplegados en esos clusters", d.GREEN),
             ("Su peso en AKS", d.percent(d.share_of(application["cost"], total)),
              f"de todo AKS, productivo y no productivo ({month})", d.STEEL),
             ("Costo mensual", v1.soles(application["cost"]), f"de {v1.soles(total)}: todo AKS, productivo y no "
                                                              "productivo", d.ORANGE),
             ("Costo anual", v1.soles(application["cost"] * 12), "costo mensual × 12", d.RUST)]
    for index, (label, value, detail, accent) in enumerate(cards):
        v1.accent_card(slide, grid_left + (index % 2) * (card_width + gap), top + (index // 2) * Inches(1.65),
                       card_width, Inches(1.5), label, value, detail, accent, value_size=30 if index >= 2 else 28)
    # Context, not the analysis: below a rule, apart and in smaller print.
    box_top, box_height = top + Inches(3.8), Inches(0.95)
    d.rounded(slide, grid_left, top + Inches(3.45), Inches(12.75) - grid_left, Inches(0.025), d.SOFT, radius=0)
    halves = int((Inches(12.75) - grid_left - gap) / 2)
    outside = [f"{label[0].upper() + label[1:]}: {v1.soles(category[code]['cost'])}" for code, label in split[1:]
               if category.get(code, {}).get("clusters")]
    nonprod = business_nonprod(results, estate["month"])
    production = sum(category[code]["clusters"] for code, _ in split[1:] if category.get(code, {}).get("clusters"))
    for index, (title, lines) in enumerate(((f"Fuera del análisis ({production} clusters), al mes", outside),
                                            (f"No productivos de negocio ({d.number(nonprod[0])} clusters), al mes",
                                             [v1.soles(nonprod[1]), f"Todos los no productivos: "
                                                                    f"{d.number(estate['nonprod_clusters'])} clusters, "
                                                                    f"{v1.soles(estate['nonprod_cost'])}"]))):
        left_box = grid_left + index * (halves + gap)
        d.note_box(slide, left_box, box_top, halves, box_height, "", size=11)
        d.text_box(slide, left_box + Inches(0.2), box_top + Inches(0.06), halves - Inches(0.4), Inches(0.35), title,
                   size=10, bold=True, color=d.GREY)
        d.text_box(slide, left_box + Inches(0.2), box_top + Inches(0.32), halves - Inches(0.4), Inches(0.6),
                   lines if not index else lines[:1], size=14 if index else 9.5, bold=bool(index), color=d.GREY)
        if index:
            d.text_box(slide, left_box + Inches(0.2), box_top + Inches(0.62), halves - Inches(0.4), Inches(0.3),
                       lines[1], size=9.5, color=d.GREY)


THOUSANDS = re.compile(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+")


def english_numbers(deck):
    """1,656,229 and 3.41333 instead of 1.656.229 and 3,41333 on every text and table of the deck."""
    swap = str.maketrans(".,", ",.")
    frames = []
    for slide in deck.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                frames.append(shape.text_frame)
            if shape.has_table:
                frames += [cell.text_frame for row in shape.table.rows for cell in row.cells]
    for frame in frames:
        for paragraph in frame.paragraphs:
            for run in paragraph.runs:
                run.text = THOUSANDS.sub(lambda match: match.group(0).translate(swap), run.text)


def example_cluster(rows):
    """The v2 counterpart of v1's example: the cluster of v1.EXAMPLE_APP when it moves, else the single-cluster
    application with the largest saving share among those with at least v1.EXAMPLE_MIN_MICROSERVICES."""
    options = [row for row in rows if row["group"] == "ACA" and int(row["clusters_of_app"]) == 1
               and int(row["microservices"]) >= v1.EXAMPLE_MIN_MICROSERVICES]
    chosen = [row for row in options if row["cod_app"] == v1.EXAMPLE_APP]
    return (chosen or sorted(options, key=lambda row: -row["saving_share"]))[0]


def business_nonprod(results, month):
    """(clusters, US$ in the month) of the CERT and DESA clusters whose application has a PROD business application
    cluster (matched by the application code in the cluster name)."""
    business = {row["cod_app"].upper() for row in v1.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["category"] == aa.APPLICATION}
    ids = {row["id"].lower() for row in v1.results_csv(results, "aa-inventory", "aks_clusters.csv")
           if row["environment"] in ("CERT", "DESA") and row["name"].lower().startswith("aksveu2")
           and row["name"][7:11].upper() in business}
    cost = defaultdict(float)
    for row in v1.results_csv(results, "aa-cost", "cost_by_resource.csv"):
        if row["month"] == month and row["owner_id"].lower() in ids:
            cost[row["owner_id"].lower()] += v1.to_float(row["cost_usd"]) or 0
    return sum(1 for value in cost.values() if value > 0), sum(cost.values())


def empty_clusters_time(results, month):
    """(average months on, soles paid since created at this month's cost) of the PROD clusters with no application
    that cost something in the month: how long the infrastructure has waited for its application."""
    soles = (1 + v1.IGV) * v1.EXCHANGE_RATE
    ids = {row["cluster_id"].lower() for row in v1.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
           if row["environment"] == "PROD" and row["category"] == aa.NO_APPLICATIONS}
    cost = defaultdict(float)
    for row in v1.results_csv(results, "aa-cost", "cost_by_resource.csv"):
        if row["month"] == month and row["owner_id"].lower() in ids:
            cost[row["owner_id"].lower()] += v1.to_float(row["cost_usd"]) or 0
    year, number = (int(part) for part in month.split("-"))
    end = date(year + number // 12, number % 12 + 1, 1)
    months, paid = [], 0.0
    for row in v1.results_csv(results, "aa-inventory", "aks_clusters.csv"):
        key = row["id"].lower()
        if key in cost and cost[key] > 0 and row.get("created_at"):
            age = max((end - date.fromisoformat(row["created_at"][:10])).days / 30.44, 0)
            months.append(age)
            paid += age * cost[key] * soles
    return (sum(months) / len(months) if months else 0), paid


def slide_hypothesis(deck, estate, apps, scaling, results, example, source):
    """The hypothesis: four signs in a square (CPU paid and not used, autoscaling that does not move, few
    "micro"-services per cluster, PROD clusters created empty while their application is still in development), the
    motivation beside them and the example costed by the v2 model below."""
    month = estate["month"]
    aks_min, aca_min, _ = v1.minimum_costs(month, v1.read_prices(results))
    app_clusters = [row for row in v1.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                    if row["environment"] == "PROD" and row["category"] == aa.APPLICATION and row["in_inventory"] == "YES"]
    small = sum(0 < int(row["microservices"] or 0) <= v1.MAX_MICROSERVICES for row in app_clusters)
    stuck = d.share_of(scaling["over_3_months"], scaling["dated"]) if scaling["dated"] else None
    empty = estate["by_category"].get(aa.NO_APPLICATIONS, {"clusters": 0, "cost": 0})
    waited, paid = empty_clusters_time(results, month)
    slide = d.base_slide(deck, "Hipótesis: un cluster AKS por aplicación paga más de lo que usa", source)
    cpu_average = d.share_of(estate["cpu_used"], estate["cpu_total"])
    cpu_peak = d.share_of(estate["cpu_peak"], estate["cpu_total"])
    signs = [(d.percent(cpu_average), "Pagamos CPU que casi no usamos",
              f"Se usa en promedio; en el pico del mes, {d.percent(cpu_peak)}."),
             (d.percent(stuck), "El autoescalado casi no se mueve", "No sube ni baja de réplicas hace más de 3 meses."),
             (f"{d.number(small)} de {d.number(len(app_clusters))}", f"Un cluster para pocos {MICRO}",
              f"Clusters con {v1.MAX_MICROSERVICES} o menos."),
             (d.number(empty["clusters"]), "Clusters creados antes de tiempo",
              f"Vacíos {waited:.0f} meses en promedio: {v1.soles(empty['cost'])} al mes, ~{pen(paid)} ya pagados.")]
    d.heading(slide, Inches(0.6), Inches(1.2), Inches(7.6), "Indicios", size=14)
    # One block split in four equal parts.
    width, height, figure = Inches(3.825), Inches(1.5), Inches(1.65)
    d.rounded(slide, Inches(0.6), Inches(1.6), 2 * width, 2 * height, d.CARD, radius=0.06)
    v1.shape(slide, v1.MSO_SHAPE.RECTANGLE, Inches(0.6) + width - Inches(0.01), Inches(1.75), Inches(0.02),
             2 * height - Inches(0.3), v1.TRACK_COLOR)
    v1.shape(slide, v1.MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(1.6) + height - Inches(0.01), 2 * width - Inches(0.3),
             Inches(0.02), v1.TRACK_COLOR)
    for index, (value, title, detail) in enumerate(signs):
        left = Inches(0.6) + (index % 2) * width
        top = Inches(1.6) + (index // 2) * height
        bar = index == 0  # the CPU paid: average strong, peak light with a tick, as v1 drew it
        d.text_box(slide, left + Inches(0.05), top + (Inches(0.2) if bar else 0), figure,
                   Inches(0.6) if bar else height, value, size=18 if " de " in value else 22, bold=True, color=d.ORANGE,
                   align=d.PP_ALIGN.CENTER, anchor=d.MSO_ANCHOR.MIDDLE)
        if bar:
            bar_left, bar_width, bar_top, bar_height = left + Inches(0.2), figure - Inches(0.3), top + Inches(0.95), Inches(0.16)
            for share, colour in ((1.0, v1.TRACK_COLOR), (cpu_peak or 0, v1.PEAK_COLOR), (cpu_average or 0, d.ORANGE)):
                if share > 0:
                    d.rounded(slide, bar_left, bar_top, max(int(bar_width * min(share, 1.0)), Inches(0.04)),
                              bar_height, colour, radius=0.5)
            if cpu_peak:
                tick = bar_left + int(bar_width * min(cpu_peak, 1.0))
                v1.shape(slide, v1.MSO_SHAPE.RECTANGLE, tick - Inches(0.012), bar_top - Inches(0.04), Inches(0.024),
                         bar_height + Inches(0.08), d.WHITE)
        d.text_box(slide, left + figure, top + Inches(0.12), width - figure - Inches(0.1), Inches(0.5), title, size=12,
                   bold=True, color=d.WHITE)
        d.text_box(slide, left + figure, top + Inches(0.68), width - figure - Inches(0.1), Inches(0.72), detail,
                   size=10.5, color=d.SOFT)
    d.heading(slide, Inches(8.5), Inches(1.2), Inches(4.2), "Motivación", size=14)
    d.callout(slide, Inches(8.5), Inches(1.6), Inches(4.2), Inches(1.2),
              "Económica: pagar por uso, no por servidores encendidos 24/7", size=14, fill=d.RUST)
    for index, text in enumerate(("Operación: menos clusters que actualizar y vigilar",
                                  "Complejidad: sin servidores ni Kubernetes que administrar")):
        d.callout(slide, Inches(8.5), Inches(2.95) + index * Inches(0.75), Inches(4.2), Inches(0.6), text, size=12,
                  fill=d.CARD)
    d.heading(slide, Inches(0.6), Inches(4.75), Inches(12.1),
              f"Ejemplo real: {example['cod_app']}, {int(example['microservices'])} {MICRO} en un cluster. En "
              f"Container Apps costaría {d.percent(example['saving_share'])} menos al mes", size=14)
    d.kpi(slide, Inches(0.6), Inches(5.15), Inches(5.9), Inches(1.15), pen(example["aks_month_pen"]),
          f"hoy en AKS: lo que se pagó en {v1.month_label(month)} por su cluster, con reservas", value_color=d.RUST,
          value_size=26, label_size=11)
    d.kpi(slide, Inches(6.8), Inches(5.15), Inches(5.9), Inches(1.15), pen(example["aca_month_pen"]),
          f"en Container Apps: los mismos {int(example['microservices'])} {MICRO}, con el "
          f"{d.percent(v1.BANK_DISCOUNT)} de descuento del banco", value_color=d.GREEN, value_size=26, label_size=11)
    d.text_box(slide, Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.25),
               f"Con IGV. Mínimo de referencia: {v1.soles(aks_min)} un cluster AKS de {v1.MIN_AKS_NODES} servidores; "
               f"{v1.soles(aca_min)} un {MICRO_ONE} en Container Apps encendido todo el mes.", size=9, color=d.GREY,
               italic=True)


def slide_criteria(deck, source):
    """Four numbered conditions, the rules that always keep a cluster on AKS, the saving margin, and at the bottom
    (not a condition) how Container Apps is costed; the sizing figures go in the footnote."""
    conditions = [
        (f"No más de {model.DEFAULTS['max_tps']} TPS", "Promedio de su hora pico, en los días del mes."),
        (f"No más de {model.DEFAULTS['max_services']} {MICRO}", "En el cluster."),
        ("No es compartida", "No es autenticación ni Core.") if model.DEFAULTS["rule_shared"] == "YES" else None,
        ("Sin carga alta ni constante", f"Hasta {model.DEFAULTS['max_avg_tps'] * 86400 / 1e6:.1f} millones de solicitudes al día "
                                        f"({model.DEFAULTS['max_avg_tps']} TPS).") if model.DEFAULTS["rule_high_load"] == "YES" else None,
        ("Sin carga constante", f"Usaría menos del {d.percent(model.DEFAULTS['max_replicas_vs_aks'])} de "
                                "las réplicas de hoy.") if model.DEFAULTS["rule_constant_load"] == "YES" else None]
    conditions = [condition for condition in conditions if condition]  # the rules switched off in the variables
    words = {2: "dos", 3: "tres", 4: "cuatro", 5: "cinco"}
    slide = d.base_slide(deck, f"Para pasar a Container Apps: {words[len(conditions)]} criterios por cluster", source)
    width, height, gap = Inches(5.9), Inches(1.1), Inches(0.3)
    for index, (title, meaning) in enumerate(conditions):
        alone = index == len(conditions) - 1 and index % 2 == 0  # an odd last card takes the whole row
        v1.condition_card(slide, Inches(0.6) + (index % 2) * (width + gap), Inches(1.25) + (index // 2) * Inches(1.25),
                          2 * width + gap if alone else width, height, str(index + 1), d.GREEN, title, meaning)
    v1.box(slide, Inches(0.6), Inches(3.85), Inches(12.1), Inches(0.75), d.LIGHT)
    d.text_box(slide, Inches(0.9), Inches(3.85), Inches(11.5), Inches(0.75),
               f"Siempre se queda en AKS si un {MICRO_ONE} demanda más CPU o memoria de la que soporta Container Apps.",
               size=14, bold=True, color=d.NAVY, anchor=d.MSO_ANCHOR.MIDDLE)
    v1.box(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(0.75), d.LIGHT)
    d.text_box(slide, Inches(0.9), Inches(4.85), Inches(11.5), Inches(0.75), sizing("band"), size=14, color=d.NAVY,
               anchor=d.MSO_ANCHOR.MIDDLE)


def slide_requirement(deck, rows, source):
    """The operating cost of moving: what each application must meet before it moves, after the saving. Quarkus for
    now; more requirements are added as they are confirmed."""
    slide = d.base_slide(deck, "Costo operativo de migrar: lo que cada aplicación debe cumplir", source)
    left, top, width, height = Inches(0.6), Inches(1.4), Inches(12.1), Inches(2.2)
    v1.box(slide, left, top, width, height, d.LIGHT)
    v1.box(slide, left, top, Inches(0.14), height, d.GREEN, radius=0.02)
    d.text_box(slide, left + Inches(0.45), top + Inches(0.25), width - Inches(0.8), Inches(0.5),
               f"Sus {MICRO} en Quarkus", size=22, bold=True, color=d.NAVY)
    d.text_box(slide, left + Inches(0.45), top + Inches(0.9), width - Inches(0.8), Inches(1.2),
               ["• Arranca en segundos: puede apagarse fuera de horario y volver rápido.",
                "• Usa menos memoria: réplicas más chicas y más baratas.",
                f"• Los {MICRO} que no están en Quarkus se migran primero: ese es el costo de adaptarse."],
               size=14, color=d.GREY)
    d.text_box(slide, Inches(0.6), Inches(3.85), Inches(12.1), Inches(0.4),
               "Nota: los pipelines para desplegar Quarkus en Container Apps ya están disponibles.", size=12,
               italic=True, color=d.GREY)


# The result slide's groups, at most four: (label, what it means, colour, groups of the model).
DECK_GROUPS = [("Se queda en AKS", "No cumple un criterio o no ahorra", d.NAVY, ("AKS", "BELOW_MARGIN")),
               ("Container Apps, con revisión", "Crítica o crece", d.ORANGE, ("ACA_WITH_REVIEW",)),
               ("Candidato a Container Apps", "Cumple todo y ahorra", d.GREEN, ("ACA",)),
               ("Sin tráfico medido", "En uso sin medición, sin uso o en desarrollo", d.SOFT,
                ("IN_USE", "NOT_IN_USE", "IN_DEVELOPMENT"))]


def slide_distribution(deck, rows, month, source):
    shown = []
    for label, meaning, colour, keys in DECK_GROUPS:
        members = [row for row in rows if row["group"] in keys]
        if members:
            shown.append((label, meaning, colour, len(members), sum(row["aks_month_pen"] for row in members), keys))
    total = len(rows)
    moving = sum(1 for row in rows if row["goes_to_aca"])
    unmeasured = sum(entry[3] for entry in shown if "IN_USE" in entry[5])
    slide = d.base_slide(deck, f"Resultado: {moving} de {total} clusters podrían pasar a Container Apps", source)
    d.bar_chart(slide, Inches(0.45), Inches(1.3), Inches(4.3), Inches(3.5),
                [entry[0].replace(", ", ",\n") for entry in shown], [entry[3] for entry in shown],
                XL_CHART_TYPE.COLUMN_CLUSTERED, [entry[2] for entry in shown])
    table = [["Destino", "Qué significa", "Clusters", "AKS hoy / mes"]]
    for label, meaning, _, count, cost, _ in shown:
        table.append([label, meaning, d.number(count), pen(cost)])
    table.append(["Total", "", d.number(total), pen(sum(entry[4] for entry in shown))])
    d.table(slide, Inches(5.0), Inches(1.3), Inches(7.8), table, [0.3, 0.38, 0.12, 0.2], row_height=0.45, size=11)
    card_w, gap, top = int((Inches(12.4) - 2 * Inches(0.3)) / 3), Inches(0.3), Inches(5.0)
    v1.stat_card(slide, Inches(0.45), top, card_w, Inches(1.25), "Hoy", f"{d.number(total)} clusters", "todos en AKS",
                 d.NAVY)
    v1.stat_card(slide, Inches(0.45) + card_w + gap, top, card_w, Inches(1.25), "Nuevo: AKS",
                 f"{d.number(total - moving)} clusters", f"se quedan ({d.percent(d.share_of(total - moving, total))}), "
                 f"{d.number(unmeasured)} aún sin tráfico medido", d.NAVY)
    v1.stat_card(slide, Inches(0.45) + 2 * (card_w + gap), top, card_w, Inches(1.25), "Nuevo: ACA",
                 f"{d.number(moving)} clusters", f"pasan a Container Apps ({d.percent(d.share_of(moving, total))})",
                 d.WHITE, fill=d.CARD, text_color=d.SOFT)
    d.text_box(slide, Inches(0.45), Inches(6.35), Inches(12.4), Inches(0.4),
               f"Costo AKS de {v1.month_label(month)}, con IGV. Una aplicación puede tener clusters en grupos distintos.",
               size=9, color=d.GREY)


def slide_cost(deck, rows, source):
    moving = [row for row in rows if row["goes_to_aca"]]
    aks = sum(row["aks_month_pen"] for row in moving)
    aca = sum(row["aca_month_pen"] for row in moving)
    saving = aks - aca
    share = d.percent(d.share_of(saving, aks))
    total = sum(row["aks_month_pen"] for row in rows)  # every business application cluster of the analysis
    slide = d.base_slide(deck, f"Pasar {len(moving)} clusters a Container Apps ahorra ~{pen(saving)} al mes: {share} "
                               f"de su costo, {d.percent(d.share_of(saving, total))} del total", source)
    table = [["Con IGV, al mes", "Hoy en AKS", "Después", "Ahorro / mes", "Baja", "Ahorro / año"],
             [f"Los {len(moving)} clusters que pasan", pen(aks), pen(aca), pen(saving), share, pen(saving * 12)],
             [f"Total de negocio ({len(rows)} clusters)", pen(total), pen(total - saving), pen(saving),
              d.percent(d.share_of(saving, total)), pen(saving * 12)]]
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.3, 0.14, 0.14, 0.14, 0.1, 0.18], highlight_column=4,
            row_height=0.45, size=13)
    how = [["Cómo se paga", "AKS", "Container Apps"],
           ["Cobro", "Servidores completos (VM) las 24 horas", "vCPU-segundo y GB-segundo de las réplicas encendidas"],
           ["Réplicas", "Las que se configuran, haya o no tráfico",
            sizing("cost")],
           ["Descuento", "Reservas de 3 años en las VM", f"{d.percent(model.DEFAULTS['bank_discount'])} del banco"]]
    d.table(slide, Inches(0.6), Inches(2.85), Inches(12.1), how, [0.18, 0.37, 0.45], row_height=0.45, size=12)
    d.note_box(slide, Inches(0.6), Inches(5.0), Inches(12.1), Inches(0.75),
               "Incluye los clusters \"con revisión\": el ahorro firme es el de los candidatos. "
               + ("Incluye el tamaño de cada réplica según su uso real. " if CRITERIA["name"] == "v3" else "")
               + ("Incluye el cargo por entorno con endpoint privado. "
                  if model.DEFAULTS["environment_fee"] == "YES" else ""), size=11)
    d.note_box(slide, Inches(0.6), Inches(5.75), Inches(12.1), Inches(0.6),
               "Beneficio adicional sin monto: menos clusters que actualizar y mantener (horas-hombre de los equipos).",
               size=11)


def run_v1(args, rest, name, v1_main):
    """Run a v1 script's main with the same --results and --output (default: name in the newest report folder)."""
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, name)
    sys.argv = [sys.argv[0], "--results", args.results, "--output", output] + rest
    v1_main()


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--criteria", choices=["v1"] + sorted(model.PRESETS), default="v2",
                        help="Criteria v1 (aa_deck_exec_v1.py), v2 or v3 (aa_v2_model.PRESETS)")
    parser.add_argument("--set", action="append", default=[], metavar="NAME=VALUE",
                        help="Change a variable's default, e.g. --set rule_shared=NO (repeatable; give the workbook "
                             "and the deck the same ones)")
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: aks_aca_corporate_<criteria>.pptx in the newest reports/<date>-aks-aca")
    args, rest = parser.parse_known_args()
    if args.criteria == "v1":
        run_v1(args, rest, "aks_aca_corporate_v1.pptx", v1.main)
        return
    if rest:
        parser.error("unrecognized arguments: " + " ".join(rest))
    aa.azure.setup_logging(False)
    CRITERIA["name"] = model.use(args.criteria, args.set)
    data = model.load(args.results)
    rows = model.evaluate(data)
    estate, apps, scaling = v1.read_estate(args.results), v1.read_applications(args.results), v1.read_scaling(args.results)
    period = v1.month_label(estate["month"])
    measured = f"Fuente: inventario de Azure, Cost Management (costo amortizado, {period}) y métricas de {period}."
    traffic = (f"Fuente: tráfico de APIM, Application Gateway y Event Hubs por hora, Container Insights y configuración "
               f"de autoescalado, {period}. Referencia: lineamiento ADT_019; escalado HTTP de Container Apps "
               "(Microsoft Learn).")
    example = example_cluster(rows)
    aa.azure.LOG.info("Example: %s (%s), AKS S/ %.0f vs Container Apps S/ %.0f (%.0f%% less)", example["cod_app"],
                      example["cluster"], example["aks_month_pen"], example["aca_month_pen"],
                      100 * example["saving_share"])
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    v1.slide_context(deck, apps, "Fuente: lineamientos de arquitectura del banco (FAT_001, ADT_019).")
    slide_estate(deck, estate, apps, args.results, f"{measured} {v1.MONEY_NOTE}")
    slide_hypothesis(deck, estate, apps, scaling, args.results, example,
                     f"{measured} Mínimos: calculadora de precios de Azure (East US 2, precios de lista). {v1.MONEY_NOTE}")
    slide_criteria(deck, f"Referencia: lineamiento ADT_019 del banco. Hora pico: la hora del día con más tráfico, "
                         "promediada en los días del mes. Carga alta: más de "
                         f"{model.DEFAULTS['max_avg_tps']} TPS de promedio en el día; constante: desde {model.DEFAULTS['min_constant_tps']} TPS de promedio, "
                         f"con su hora más tranquila en al menos {d.percent(model.DEFAULTS['min_quiet_share'])} del "
                         f"promedio. Réplicas: {sizing('criteria')} "
                         "Fuente: Microsoft Learn.")
    slide_distribution(deck, rows, estate["month"], f"{traffic} {v1.MONEY_NOTE}")
    slide_cost(deck, rows, f"{measured} Precios de lista: Azure Retail Prices API. {v1.MONEY_NOTE}")
    slide_requirement(deck, rows, "")
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR,
                                         f"aks_aca_corporate_{args.criteria}.pptx")
    english_numbers(deck)
    aa.keep_version(output)
    deck.save(output)
    columns = ["cluster", "cod_app", "group", "decision", "reason", "microservices", "traffic_layer",
               "busiest_hour_tps", "peak_tps", "aks_month_pen", "aca_month_pen", "saving_month_pen", "saving_share"]
    aa.azure.Output(os.path.dirname(os.path.abspath(output)), ",").csv(
        f"cluster_decisions_{args.criteria}.csv",
        [{key: (round(row[key], 4) if isinstance(row[key], float) else row[key]) for key in columns} for row in rows])
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


if __name__ == "__main__":
    main()
