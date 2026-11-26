#!/usr/bin/env python3
"""Build the corporate deck of criteria v1, v2 or v3 (--criteria; v2 by default).

v1 runs aa_deck_exec_v1.py (per application, frozen at tag v1). v2 and v3 come from the criteria model
(aa_v2_model.py; v3 is v2 with its PRESETS["v3"] defaults): the decision per cluster and the Container Apps cost from
each hour's traffic.

Slides 1 and 2 are v1's (aa_deck_exec_v1.py). Then:
  3  hypothesis, with the real example costed by the v2 model
  4  criteria per cluster: four conditions, the rules that keep a cluster on AKS, the saving margin, how it is costed
  5  result per cluster (groups of the v2 workbook's Results sheet)
  6  saving: what moving the clusters of the ACA groups saves a month and a year
Reads only results/ (run aa_cluster_snapshot.py first). Output: aks_aca_corporate_<criteria>.pptx in the newest
reports/<date>-aks-aca folder (or --output); next to it, cluster_decisions_<criteria>.csv with every cluster's figures.
"""

import argparse
import glob
import math
import os
import sys
from collections import defaultdict

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

import aa_common as aa
import aa_deck as d
import aa_deck_exec_v1 as v1
import aa_v2_model as model

MICRO, MICRO_ONE = v1.MICRO, v1.MICRO_ONE
# Groups of the v2 result, in slide order: (label, what it means, colour).
GROUPS = {
    "AKS": ("Se queda en AKS", "No cumple una condición o una regla de AKS", d.NAVY),
    "BELOW_MARGIN": ("Cumple, pero ahorra poco", "Ahorra menos del margen: se queda en AKS", d.STEEL),
    "ACA_WITH_REVIEW": ("Container Apps, con revisión", "Crítica o crece", d.ORANGE),
    "ACA": ("Candidato a Container Apps", "Cumple todo y ahorra", d.GREEN),
    "IN_USE": ("En uso, sin medición", "Tiene uso, pero su tráfico no se ve", d.SOFT),
    "NOT_IN_USE": ("Sin uso", "Por revisar", d.GREY),
    "IN_DEVELOPMENT": ("En desarrollo", "Aún no sale a producción", d.SOFT),
}
MOVABLE = ("ACA", "ACA_WITH_REVIEW")


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


def example_cluster(rows):
    """The v2 counterpart of v1's example: the cluster of v1.EXAMPLE_APP when it moves, else the single-cluster
    application with the largest saving share among those with at least v1.EXAMPLE_MIN_MICROSERVICES."""
    options = [row for row in rows if row["group"] == "ACA" and int(row["clusters_of_app"]) == 1
               and int(row["microservices"]) >= v1.EXAMPLE_MIN_MICROSERVICES]
    chosen = [row for row in options if row["cod_app"] == v1.EXAMPLE_APP]
    return (chosen or sorted(options, key=lambda row: -row["saving_share"]))[0]


def slide_hypothesis(deck, estate, apps, scaling, results, example, source):
    """v1's slide 3 with the example costed by the v2 model."""
    original = v1.example_app
    soles = (1 + v1.IGV) * v1.EXCHANGE_RATE
    v1.example_app = lambda *_: {"code": example["cod_app"], "cluster": example["cluster"],
                                 "microservices": int(example["microservices"]),
                                 "aks": example["aks_month_pen"] / soles, "aca": example["aca_month_pen"] / soles,
                                 "saving": example["saving_share"]}
    try:
        v1.slide_hypothesis(deck, estate, apps, scaling, results, source)
    finally:
        v1.example_app = original
    slide = deck.slides[-1]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.startswith("Con IGV."):
            shape.text_frame.paragraphs[0].runs[0].text = (
                f"Con IGV. Container Apps con réplicas según su tráfico: {sizing('example')} Mínimo de referencia: S/ 2.124 un cluster AKS de 4 "
                f"servidores; S/ 117 un {MICRO_ONE} en Container Apps encendido todo el mes.")
            for run in shape.text_frame.paragraphs[0].runs[1:]:
                run.text = ""


def slide_criteria(deck, source):
    """Four numbered conditions, the rules that always keep a cluster on AKS, the saving margin, and at the bottom
    (not a condition) how Container Apps is costed; the sizing figures go in the footnote."""
    conditions = [
        (f"Hasta {model.DEFAULTS['max_tps']} TPS", "En su hora más cargada del día."),
        (f"Hasta {model.DEFAULTS['max_services']} {MICRO}", "En el cluster."),
        ("No es compartida", "No es autenticación ni Core.") if model.DEFAULTS["rule_shared"] == "YES" else None,
        ("Sin carga alta y constante", f"Usaría menos del {d.percent(model.DEFAULTS['max_replicas_vs_aks'])} de "
                                       "las réplicas de hoy.") if model.DEFAULTS["rule_constant_load"] == "YES" else None]
    conditions = [condition for condition in conditions if condition]  # the rules switched off in the variables
    words = {2: "dos", 3: "tres", 4: "cuatro"}
    slide = d.base_slide(deck, f"Para pasar a Container Apps: {words[len(conditions)]} condiciones por cluster", source)
    width, height, gap = Inches(5.9), Inches(1.1), Inches(0.3)
    for index, (title, meaning) in enumerate(conditions):
        alone = index == len(conditions) - 1 and index % 2 == 0  # an odd last card takes the whole row
        v1.condition_card(slide, Inches(0.6) + (index % 2) * (width + gap), Inches(1.25) + (index // 2) * Inches(1.25),
                          2 * width + gap if alone else width, height, str(index + 1), d.GREEN, title, meaning)
    v1.box(slide, Inches(0.6), Inches(3.85), Inches(12.1), Inches(0.75), d.LIGHT)
    d.text_box(slide, Inches(0.9), Inches(3.85), Inches(11.5), Inches(0.75),
               f"Siempre se queda en AKS si un {MICRO_ONE} demanda más CPU o memoria de la que soporta Container Apps.",
               size=14, bold=True, color=d.NAVY, anchor=d.MSO_ANCHOR.MIDDLE)
    d.text_box(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(0.4),
               f"Solo se considera para migrar si ahorra al menos {d.percent(model.DEFAULTS['min_saving'])}.",
               size=15, bold=True, color=d.NAVY)
    v1.box(slide, Inches(0.6), Inches(5.45), Inches(12.1), Inches(0.75), d.LIGHT)
    d.text_box(slide, Inches(0.9), Inches(5.45), Inches(11.5), Inches(0.75), sizing("band"), size=14, color=d.NAVY,
               anchor=d.MSO_ANCHOR.MIDDLE)


def slide_distribution(deck, rows, month, source):
    groups = defaultdict(lambda: {"clusters": 0, "today": 0.0, "after": 0.0})
    for row in rows:
        entry = groups[row["group"]]
        entry["clusters"] += 1
        entry["today"] += row["aks_month_pen"]
        entry["after"] += row["aca_month_pen"] if row["goes_to_aca"] else row["aks_month_pen"]
    order = [key for key in GROUPS if groups[key]["clusters"]]
    total = sum(groups[key]["clusters"] for key in order)
    moving = sum(groups[key]["clusters"] for key in MOVABLE)
    slide = d.base_slide(deck, f"Resultado: {moving} de {total} clusters podrían pasar a Container Apps", source)
    d.bar_chart(slide, Inches(0.45), Inches(1.3), Inches(4.3), Inches(3.5),
                [GROUPS[key][0].replace(", ", ",\n") for key in order], [groups[key]["clusters"] for key in order],
                XL_CHART_TYPE.COLUMN_CLUSTERED, [GROUPS[key][2] for key in order])
    table = [["Destino", "Qué significa", "Clusters", "AKS hoy / mes"]]
    for key in order:
        table.append([GROUPS[key][0], GROUPS[key][1], d.number(groups[key]["clusters"]), pen(groups[key]["today"])])
    table.append(["Total", "", d.number(total), pen(sum(groups[key]["today"] for key in order))])
    d.table(slide, Inches(5.0), Inches(1.3), Inches(7.8), table, [0.3, 0.38, 0.12, 0.2], row_height=0.4, size=11)
    card_w, gap, top = int((Inches(12.4) - 2 * Inches(0.3)) / 3), Inches(0.3), Inches(5.0)
    v1.stat_card(slide, Inches(0.45), top, card_w, Inches(1.25), "Hoy", f"{d.number(total)} clusters", "todos en AKS",
                 d.NAVY)
    v1.stat_card(slide, Inches(0.45) + card_w + gap, top, card_w, Inches(1.25), "Nuevo: AKS",
                 f"{d.number(total - moving)} clusters", f"se quedan ({d.percent(d.share_of(total - moving, total))})",
                 d.NAVY)
    v1.stat_card(slide, Inches(0.45) + 2 * (card_w + gap), top, card_w, Inches(1.25), "Nuevo: ACA",
                 f"{d.number(moving)} clusters", f"pasan a Container Apps ({d.percent(d.share_of(moving, total))})",
                 d.WHITE, fill=d.CARD, text_color=d.SOFT)
    below = groups["BELOW_MARGIN"]["clusters"]
    note = (f"{d.number(below)} {'cluster cumple' if below == 1 else 'clusters cumplen'} las condiciones pero "
            f"{'ahorra' if below == 1 else 'ahorran'} menos del {d.percent(model.DEFAULTS['min_saving'])}: "
            f"{'se queda' if below == 1 else 'se quedan'} en AKS. " if below else "")
    d.text_box(slide, Inches(0.45), Inches(6.35), Inches(12.4), Inches(0.4),
               f"{note}Costo AKS de {v1.month_label(month)}, con IGV. Una aplicación puede tener clusters en grupos "
               "distintos.", size=10 if below else 9, bold=bool(below), color=d.NAVY if below else d.GREY)


def slide_cost(deck, rows, source):
    moving = [row for row in rows if row["goes_to_aca"]]
    aks = sum(row["aks_month_pen"] for row in moving)
    aca = sum(row["aca_month_pen"] for row in moving)
    saving = aks - aca
    slide = d.base_slide(deck, f"Apagar {len(moving)} clusters ahorraría ~{pen(saving)} al mes ({pen(saving * 12)} al "
                               "año)", source)
    table = [["Con IGV", "AKS hoy / mes", "Container Apps / mes", "Ahorro / mes", "Ahorro / año"],
             [f"{len(moving)} clusters que pasan", pen(aks), pen(aca), pen(saving), pen(saving * 12)]]
    d.table(slide, Inches(0.6), Inches(1.4), Inches(12.1), table, [0.28, 0.18, 0.2, 0.17, 0.17], highlight_column=3,
            row_height=0.5, size=13)
    how = [["Cómo se paga", "AKS", "Container Apps"],
           ["Cobro", "Servidores completos (VM) las 24 horas", "vCPU-segundo y GB-segundo de las réplicas encendidas"],
           ["Réplicas", "Las que se configuran, haya o no tráfico",
            sizing("cost")],
           ["Descuento", "Reservas de 3 años en las VM", f"{d.percent(model.DEFAULTS['bank_discount'])} del banco"]]
    d.table(slide, Inches(0.6), Inches(2.6), Inches(12.1), how, [0.18, 0.37, 0.45], row_height=0.45, size=12)
    d.note_box(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(0.75),
               f"Incluye los clusters \"con revisión\": el ahorro firme es el de los candidatos. Los que cumplen pero "
               f"ahorran menos de {d.percent(model.DEFAULTS['min_saving'])} se quedan en AKS. "
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
    v1.slide_estate(deck, estate, apps, f"{measured} {v1.MONEY_NOTE}")
    slide_hypothesis(deck, estate, apps, scaling, args.results, example,
                     f"{measured} Mínimos: calculadora de precios de Azure (East US 2, precios de lista). {v1.MONEY_NOTE}")
    slide_criteria(deck, f"Referencia: lineamiento ADT_019 del banco. Réplicas: {sizing('criteria')} "
                         "Fuente: Microsoft Learn.")
    slide_distribution(deck, rows, estate["month"], f"{traffic} {v1.MONEY_NOTE}")
    slide_cost(deck, rows, f"{measured} Precios de lista: Azure Retail Prices API. {v1.MONEY_NOTE}")
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR,
                                         f"aks_aca_corporate_{args.criteria}.pptx")
    deck.save(output)
    columns = ["cluster", "cod_app", "group", "decision", "reason", "microservices", "traffic_layer",
               "busiest_hour_tps", "peak_tps", "aks_month_pen", "aca_month_pen", "saving_month_pen", "saving_share"]
    aa.azure.Output(os.path.dirname(os.path.abspath(output)), ",").csv(
        f"cluster_decisions_{args.criteria}.csv",
        [{key: (round(row[key], 4) if isinstance(row[key], float) else row[key]) for key in columns} for row in rows])
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


if __name__ == "__main__":
    main()
