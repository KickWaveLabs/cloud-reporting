#!/usr/bin/env python3
"""Build the four-slide AKS deck (summary and patterns 1, 2 and 4) from the saved workbook.

Reads the values Excel calculated in aks_aca_comparison.xlsx, so open the workbook in Excel and save it once
before running this script. No Azure call is made; nothing leaves the machine.

Slides:
  1. Summary: cost, clusters, office CPU use, share of clusters with 15 microservices or fewer
  2. Pattern 1: microservices by layer and capabilities repeated across domains
  3. Pattern 2: clusters by number of microservices and the cost of the smallest
  4. Pattern 4: current replicas and HPA behaviour over the measured months

Output: aks_patrones.pptx next to the workbook (or --output). Charts are native and editable in PowerPoint.
"""

import argparse
import glob
import os
import re

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

import aa_common as aa

NAVY = RGBColor(0x0B, 0x2A, 0x5B)
ORANGE = RGBColor(0xF7, 0x94, 0x1D)
GREY = RGBColor(0x5B, 0x67, 0x70)
LIGHT = RGBColor(0xEE, 0xF1, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xB8, 0xC2, 0xCC)
FONT = "Arial"
LAYER_LABELS = {"Capa business-*": "business", "Capa channel-*": "channel", "Capa app-*": "app",
                "Capa async-*": "async", "Capa support-*": "support", "Capa core-*": "core", "Otras capas": "otras"}
BEHAVIOUR_ORDER = ["Se queda en su mínimo", "Escala entre mínimo y máximo", "Llega o se acerca a su máximo",
                   "Mínimo = máximo"]
SMALL_BUCKETS = {"1", "2 a 3", "4 a 5", "6 a 10", "11 a 15"}


def latest_workbook():
    paths = glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*", "aks_aca_comparison.xlsx"))
    if not paths:
        raise SystemExit("No reports/*/aks_aca_comparison.xlsx found. Run aa_workbook.py or pass --workbook.")
    return max(paths, key=os.path.getmtime)


def read_summary(book):
    values = {}
    for row in book["Resumen"].iter_rows(values_only=True):
        if row and isinstance(row[0], str) and len(row) > 1:
            values.setdefault(row[0].strip(), row[1])
    return values


def read_patterns(book):
    """Split the Patrones sheet by section and read each table by its labels."""
    data = {"layers": [], "capabilities": [], "clusters": [], "replicas": [], "behaviour": {}, "title": ""}
    section, in_capabilities = "", False
    for row in book["Patrones"].iter_rows(values_only=True):
        label = row[0] if row else None
        if not isinstance(label, str):
            continue
        label = label.strip()
        if not data["title"]:
            data["title"] = label
        if label.startswith("Patrón 1"):
            section = "p1"
        elif label.startswith("Patrón 2"):
            section, in_capabilities = "p2", False
        elif label.startswith("Patrón 4"):
            section = "p4"
        elif section == "p1":
            if label == "Microservicios PROD":
                data["microservices"] = row[1]
            elif label in LAYER_LABELS:
                data["layers"].append((LAYER_LABELS[label], row[1], row[2]))
            elif label.startswith("Capacidad repetida"):
                in_capabilities = True
            elif in_capabilities and isinstance(row[1], (int, float)):
                data["capabilities"].append((label, row[1], row[2]))
        elif section == "p2":
            if label.startswith("Clusters PROD con 15"):
                data["clusters_le15"], data["clusters_le15_share"] = row[1], row[2]
            elif label.startswith("Costo mensual de clusters PROD con 3"):
                data["cost_le3"] = row[3]
            elif isinstance(row[1], (int, float)):
                data["clusters"].append((label, row[1], row[2], row[3]))
        elif section == "p4":
            if label in BEHAVIOUR_ORDER or label == "Sin datos":
                data["behaviour"][label] = {"export": row[1], "ci": row[3]}
            elif isinstance(row[1], (int, float)):
                data["replicas"].append((label, row[1], row[2]))
    match = re.search(r"\(configuración al ([0-9-]+)\)", data["title"])
    data["snapshot_date"] = match.group(1) if match else ""
    return data


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------


def text_box(slide, left, top, width, height, text, size=12, bold=False, color=GREY, align=PP_ALIGN.LEFT,
             italic=False):
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name, run.font.size, run.font.bold, run.font.italic = FONT, Pt(size), bold, italic
    run.font.color.rgb = color
    return shape


def base_slide(deck, title, subtitle, source):
    slide = deck.slides.add_slide(deck.slide_layouts[6])  # blank layout
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.slide_width, Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()
    text_box(slide, Inches(0.5), Inches(0.25), Inches(12.3), Inches(0.9), title, size=22, bold=True, color=NAVY)
    text_box(slide, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.5), subtitle, size=13)
    text_box(slide, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.35), source, size=9, italic=True)
    return slide


def bar_chart(slide, left, top, width, height, categories, values, kind, colors):
    data = CategoryChartData()
    data.categories = categories
    data.add_series("Serie", values)
    chart = slide.shapes.add_chart(kind, left, top, width, height, data).chart
    chart.font.name, chart.font.size = FONT, Pt(11)
    chart.font.color.rgb = GREY
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format, labels.number_format_is_linked = "#,##0", False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    for index, color in enumerate(colors):
        point = plot.series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False
    chart.category_axis.format.line.color.rgb = SOFT
    return chart


def kpi(slide, left, top, value, label, width=Inches(2.95), height=Inches(1.9)):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT
    box.line.fill.background()
    box.adjustments[0] = 0.08
    text_box(slide, left, top + Inches(0.25), width, Inches(0.9), value, size=32, bold=True, color=ORANGE,
             align=PP_ALIGN.CENTER)
    text_box(slide, left + Inches(0.15), top + Inches(1.15), width - Inches(0.3), Inches(0.7), label, size=12,
             align=PP_ALIGN.CENTER)


def table(slide, left, top, width, rows, widths):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), left, top, width, Inches(0.32 * len(rows)))
    grid = shape.table
    for column, share in enumerate(widths):
        grid.columns[column].width = Emu(int(width * share))
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            cell = grid.cell(r, c)
            cell.text = str(value)
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            font = paragraph.runs[0].font
            font.name, font.size = FONT, Pt(11)
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = NAVY
                font.bold, font.color.rgb = True, WHITE
            else:
                cell.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
                font.color.rgb = GREY


def money(value):
    return f"US${value:,.0f}".replace(",", ".") if isinstance(value, (int, float)) else "-"


def number(value):
    return f"{value:,.0f}".replace(",", ".") if isinstance(value, (int, float)) else "-"


def percent(value):
    return f"{value:.0%}" if isinstance(value, (int, float)) else "-"


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def slide_summary(deck, s, p, source):
    total, cpu = s.get("Costo mensual total AKS PROD"), s.get("Uso de CPU en horario de oficina, PROD")
    slide = base_slide(deck, f"AKS en producción: pagamos {money(total)} al mes y usamos el {percent(cpu)} de la CPU",
                       "Cada aplicación tiene su propio cluster, con un mínimo de nodos que se paga aunque no se use.",
                       source)
    tiles = [(number(s.get("Clusters AKS PROD")), "clusters AKS en producción"),
             (money(total), "al mes (promedio ago–oct, costo amortizado)"),
             (percent(cpu), "de la CPU pagada se usa en horario de oficina"),
             (percent(p.get("clusters_le15_share")), "de los clusters tiene 15 microservicios o menos")]
    for index, (value, label) in enumerate(tiles):
        kpi(slide, Inches(0.5 + index * 3.12), Inches(2.0), value, label)
    text_box(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(1.6),
             f"Un cluster mínimo de 4 nodos cuesta {money(s.get('Cluster AKS mínimo con el descuento medido'))} al mes "
             f"con el descuento del banco, aunque la aplicación use una fracción. Con el horario del ADT_019, un "
             f"microservicio en Container Apps cuesta {money(s.get('ACA por microservicio (horario ADT)'))} al mes.",
             size=14, color=NAVY)


def slide_pattern1(deck, p, source):
    capabilities = sorted(p["capabilities"], key=lambda item: -item[1])
    most = capabilities[0][1] if capabilities else 0
    slide = base_slide(deck, f"Patrón 1: {number(p.get('microservices'))} microservicios; las mismas capacidades se "
                             f"repiten en hasta {number(most)} dominios",
                       "Dos capas concentran casi dos tercios de los microservicios; notificación, OTP, token o login se "
                       "construyen una y otra vez por dominio.", source)
    layers = sorted(p["layers"], key=lambda item: item[1])
    bar_chart(slide, Inches(0.5), Inches(1.7), Inches(6.2), Inches(5.1), [item[0] for item in layers],
              [item[1] for item in layers], XL_CHART_TYPE.BAR_CLUSTERED,
              [ORANGE if item[0] in ("business", "channel") else NAVY for item in layers])
    rows = [["Capacidad", "Dominios", "Microservicios"]] + [[name, number(domains), number(count)]
                                                           for name, domains, count in capabilities]
    table(slide, Inches(7.1), Inches(1.9), Inches(5.7), rows, [0.5, 0.25, 0.25])
    text_box(slide, Inches(7.1), Inches(2.1) + Inches(0.32 * len(rows)), Inches(5.7), Inches(0.9),
             "Candidatas a servicio compartido de plataforma. Detectadas por el nombre del deployment: cada dominio "
             "debe confirmarlas.", size=11, italic=True)


def slide_pattern2(deck, p, source):
    buckets = [item for item in p["clusters"] if not item[0].startswith("0")]
    empty = next((item for item in p["clusters"] if item[0].startswith("0")), None)
    total = sum(item[1] for item in p["clusters"])
    slide = base_slide(deck, f"Patrón 2: {number(p.get('clusters_le15'))} de {number(total)} clusters PROD tienen 15 "
                             "microservicios o menos",
                       "Un cluster por aplicación multiplica los costos fijos: nodos de sistema, balanceador y SLA en "
                       "cada uno.", source)
    bar_chart(slide, Inches(0.5), Inches(1.7), Inches(8.2), Inches(5.1), [item[0] for item in buckets],
              [item[1] for item in buckets], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0] in SMALL_BUCKETS else SOFT for item in buckets])
    kpi(slide, Inches(9.2), Inches(1.9), money(p.get("cost_le3")),
        "al mes cuestan los clusters con 3 microservicios o menos", width=Inches(3.6))
    kpi(slide, Inches(9.2), Inches(4.0), percent(p.get("clusters_le15_share")),
        "de los clusters, bajo el umbral de 15 microservicios del ADT_019", width=Inches(3.6))
    if empty:
        text_box(slide, Inches(9.2), Inches(6.1), Inches(3.6), Inches(0.8),
                 f"{number(empty[1])} clusters no aparecen en el export de plataforma: por revisar.", size=10,
                 italic=True)


def slide_pattern4(deck, p, source):
    behaviour = p["behaviour"]
    counts = [behaviour.get(name, {}).get("ci") or 0 for name in BEHAVIOUR_ORDER]
    total = sum(counts) + (behaviour.get("Sin datos", {}).get("ci") or 0)
    slide = base_slide(deck, f"Patrón 4: {percent(counts[0] / total if total else None)} de los HPA se queda en su "
                             "mínimo durante 3 meses",
                       "Los mínimos de réplicas son capacidad que se paga 24/7: en AKS no bajan de noche.", source)
    text_box(slide, Inches(0.5), Inches(1.55), Inches(6.6), Inches(0.4), "HPA por réplicas actuales (día de la foto)",
             size=12, bold=True, color=NAVY)
    replicas = p["replicas"]
    bar_chart(slide, Inches(0.5), Inches(1.9), Inches(6.6), Inches(4.9), [item[0] for item in replicas],
              [item[1] for item in replicas], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0].startswith("2 a 3") else NAVY for item in replicas])
    text_box(slide, Inches(7.4), Inches(1.55), Inches(5.4), Inches(0.4),
             "Comportamiento en 3 meses (promedio en oficina)", size=12, bold=True, color=NAVY)
    data = CategoryChartData()
    data.categories = BEHAVIOUR_ORDER
    data.add_series("Comportamiento", counts)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(7.4), Inches(1.9), Inches(5.4), Inches(4.9),
                                   data).chart
    chart.font.name, chart.font.size = FONT, Pt(11)
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.show_percentage, plot.data_labels.show_value = True, False
    plot.data_labels.number_format, plot.data_labels.number_format_is_linked = "0%", False
    plot.data_labels.font.size, plot.data_labels.font.color.rgb = Pt(12), WHITE
    for index, color in enumerate([ORANGE, NAVY, GREY, SOFT]):
        point = plot.series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--workbook", default=None,
                        help="aks_aca_comparison.xlsx, opened and saved in Excel. Default: the newest one in reports/")
    parser.add_argument("--output", default=None, help="Output .pptx. Default: aks_patrones.pptx next to the workbook")
    args = parser.parse_args()
    path = args.workbook or latest_workbook()
    book = load_workbook(path, data_only=True)
    summary, patterns = read_summary(book), read_patterns(book)
    if summary.get("Costo mensual total AKS PROD") in (None, ""):
        raise SystemExit("The workbook has no calculated values. Open it in Excel, save it once and run this again.")
    if not patterns["layers"]:
        raise SystemExit("The Patrones sheet is empty. Run aa_k8s_snapshot.py, then aa_workbook.py, and save in Excel.")

    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    snapshot = f"; configuración de HPA y deployments al {patterns['snapshot_date']}" if patterns["snapshot_date"] else ""
    source = ("Fuente: Azure Cost Management (costo amortizado), Azure Monitor y Container Insights (ago–nov 2026)"
              f"{snapshot}. Datos elaborados por BCP para uso interno.")
    slide_summary(deck, summary, patterns, source)
    slide_pattern1(deck, patterns, source)
    slide_pattern2(deck, patterns, source)
    slide_pattern4(deck, patterns, source)
    output = args.output or os.path.join(os.path.dirname(path), "aks_patrones.pptx")
    deck.save(output)
    print(output)


if __name__ == "__main__":
    main()
