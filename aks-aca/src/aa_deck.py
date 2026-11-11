#!/usr/bin/env python3
"""Build the four-slide AKS deck (summary and patterns 1, 2 and 4) from the saved workbook.

Reads the values Excel calculated in aks_aca_comparison.xlsx, so open the workbook in Excel and save it once
before running this script. No Azure call is made; nothing leaves the machine.

Slides:
  1. Summary: cost, clusters, CPU and memory use (average, p95, peak), cost of a minimum cluster
  2. Pattern 1: microservices by layer and capabilities repeated across domains
  3. Pattern 2: clusters and monthly cost by number of microservices
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
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

import aa_common as aa

NAVY = RGBColor(0x1B, 0x2F, 0x5C)
RUST = RGBColor(0xB5, 0x45, 0x2D)
CARD = RGBColor(0x1F, 0x2B, 0x3E)
ORANGE = RGBColor(0xF7, 0x94, 0x1D)
STEEL = RGBColor(0x4A, 0x8F, 0xB8)
GREY = RGBColor(0x5B, 0x67, 0x70)
LIGHT = RGBColor(0xEE, 0xF1, 0xF5)
SOFT = RGBColor(0xC5, 0xCD, 0xD6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"
LAYER_LABELS = {"Capa business-*": "business", "Capa channel-*": "channel", "Capa app-*": "app",
                "Capa async-*": "async", "Capa support-*": "support", "Capa core-*": "core", "Otras capas": "otras"}
BEHAVIOURS = [("Se queda en su mínimo", "Se queda en su mínimo", ORANGE), ("Escala entre mínimo y máximo", "Escala", STEEL),
              ("Llega o se acerca a su máximo", "Cerca del máximo", RUST), ("Mínimo = máximo", "Mínimo = máximo", NAVY),
              ("Sin datos", "Sin datos", SOFT)]
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
            if label in [code for code, _, _ in BEHAVIOURS]:
                data["behaviour"][label] = {"export": row[1], "ci": row[3]}
            elif label.startswith("Su mínimo supera eso aun dejando"):
                data["inflated_beyond_ha"] = row[2]
            elif label.startswith("Su mínimo supera"):
                data["inflated"] = row[2]
            elif isinstance(row[1], (int, float)):
                data["replicas"].append((label, row[1], row[2]))
    match = re.search(r"\(configuración al ([0-9-]+)\)", data["title"])
    data["snapshot_date"] = match.group(1) if match else ""
    return data


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------


def money(value):
    return f"US${value:,.0f}".replace(",", ".") if isinstance(value, (int, float)) else "-"


def number(value):
    return f"{value:,.0f}".replace(",", ".") if isinstance(value, (int, float)) else "-"


def percent(value):
    return f"{value:.0%}" if isinstance(value, (int, float)) else "-"


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


def banner(slide, deck, title):
    """Title band: navy fading into rust, like the bank's report decks."""
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.slide_width, Inches(1.05))
    band.line.fill.background()
    band.shadow.inherit = False
    band.fill.gradient()
    band.fill.gradient_angle = 0
    stops = band.fill.gradient_stops
    stops[0].color.rgb, stops[0].position = NAVY, 0.0
    stops[1].color.rgb, stops[1].position = RUST, 1.0
    tip = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, deck.slide_width - Inches(0.55), Inches(0.18), Inches(0.35),
                                 Inches(0.7))
    tip.fill.solid()
    tip.fill.fore_color.rgb = WHITE
    tip.line.fill.background()
    tip.shadow.inherit = False
    text_box(slide, Inches(0.45), 0, Inches(12.2), Inches(1.05), title.upper(), size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)


def base_slide(deck, title, source, number_):
    slide = deck.slides.add_slide(deck.slide_layouts[6])  # blank layout
    banner(slide, deck, title)
    text_box(slide, Inches(0.45), Inches(7.05), Inches(11.5), Inches(0.35), source, size=8, italic=True)
    text_box(slide, Inches(12.4), Inches(7.05), Inches(0.5), Inches(0.35), str(number_), size=9, align=PP_ALIGN.RIGHT)
    return slide


def heading(slide, left, top, width, text):
    text_box(slide, left, top, width, Inches(0.35), text, size=13, bold=True, color=NAVY)


def kpi(slide, left, top, width, height, value, label, value_color=ORANGE, value_size=30):
    rounded(slide, left, top, width, height, CARD)
    text_box(slide, left, top + Inches(0.15), width, height * 0.55, value, size=value_size, bold=True,
             color=value_color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, left + Inches(0.15), top + height * 0.58, width - Inches(0.3), height * 0.38, label, size=11,
             bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def callout(slide, left, top, width, height, text, size=15):
    rounded(slide, left, top, width, height, CARD, radius=0.18)
    text_box(slide, left + Inches(0.3), top, width - Inches(0.6), height, text, size=size, bold=True, color=WHITE,
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


def table(slide, left, top, width, rows, widths, highlight_column=None, row_height=0.34):
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
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            font = paragraph.runs[0].font
            font.name, font.size = FONT, Pt(11)
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = CARD
                font.bold, font.color.rgb = True, WHITE
            else:
                cell.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
                font.color.rgb = RUST if c == highlight_column else GREY
                font.bold = c == highlight_column


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


def slide_summary(deck, s, p, source):
    total = s.get("Costo mensual total AKS PROD")
    cpu, mem = s.get("Uso de CPU en horario de oficina, PROD"), s.get("Uso de memoria en horario de oficina, PROD")
    title = f"AKS en producción: pagamos {money(total)} al mes; usamos el {percent(cpu)} de la CPU"
    title += f" y el {percent(mem)} de la memoria" if mem not in (None, "") else ""
    slide = base_slide(deck, title, source, 1)
    tiles = [(number(s.get("Clusters AKS PROD")), "clusters AKS en producción"),
             (money(total), "al mes, promedio ago–oct (costo amortizado)"),
             (money(s.get("Costo mensual promedio por cluster PROD")), "costo mensual promedio por cluster"),
             (percent(p.get("clusters_le15_share")), "de los clusters tiene 15 microservicios o menos")]
    for index, (value, label) in enumerate(tiles):
        kpi(slide, Inches(0.45 + index * 3.13), Inches(1.35), Inches(2.95), Inches(1.65), value, label)

    heading(slide, Inches(0.45), Inches(3.25), Inches(6.5), "Uso de la capacidad asignable (todos los clusters PROD)")
    rows = [["", "Promedio en oficina", "p95 por hora", "Máximo*"],
            ["CPU", percent(cpu), percent(s.get("Uso de CPU, percentil 95 por hora (todo el día)")),
             percent(s.get("Uso de CPU, máximo (cota superior: picos de cada cluster sumados)"))],
            ["Memoria", percent(mem), percent(s.get("Uso de memoria, percentil 95 por hora (todo el día)")),
             percent(s.get("Uso de memoria, máximo (cota superior: picos de cada cluster sumados)"))]]
    table(slide, Inches(0.45), Inches(3.7), Inches(6.4), rows, [0.22, 0.28, 0.25, 0.25], row_height=0.42)
    text_box(slide, Inches(0.45), Inches(5.05), Inches(6.4), Inches(0.5),
             "* Suma del pico de cada cluster: cota superior, porque los picos no ocurren a la misma hora.", size=9,
             italic=True)

    heading(slide, Inches(7.3), Inches(3.25), Inches(5.6), "Lo que cuesta el modelo de hoy")
    kpi(slide, Inches(7.3), Inches(3.7), Inches(2.7), Inches(1.45), money(s.get("Cluster AKS mínimo con el descuento medido")),
        "cluster AKS mínimo (4 nodos) al mes", value_size=24)
    kpi(slide, Inches(10.2), Inches(3.7), Inches(2.7), Inches(1.45), money(s.get("ACA por microservicio (horario ADT)")),
        "un microservicio en ACA al mes (horario ADT)", value_size=24)
    callout(slide, Inches(0.45), Inches(5.65), Inches(12.45), Inches(1.0),
            "Cada aplicación tiene su propio cluster, con un mínimo de nodos que se paga aunque no se use: el costo lo "
            "manda la capacidad reservada, no el consumo.", size=14)


def slide_pattern1(deck, p, source):
    capabilities = sorted(p["capabilities"], key=lambda item: -item[1])
    most = capabilities[0][1] if capabilities else 0
    slide = base_slide(deck, f"Patrón 1: {number(p.get('microservices'))} microservicios; capacidades repetidas en "
                             f"hasta {number(most)} dominios", source, 2)
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
                                                           for name, domains, count in capabilities]
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
    total = sum(item[1] for item in p["clusters"])
    slide = base_slide(deck, f"Patrón 2: {number(p.get('clusters_le15'))} de {number(total)} clusters tienen 15 "
                             "microservicios o menos", source, 3)
    heading(slide, Inches(0.45), Inches(1.3), Inches(8), "Clusters PROD por número de microservicios")
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(8.2), Inches(4.3), [item[0] for item in buckets],
              [item[1] for item in buckets], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0] in SMALL_BUCKETS else STEEL for item in buckets])
    legend_item(slide, Inches(0.6), Inches(6.05), ORANGE, "15 microservicios o menos (umbral ADT_019)")
    legend_item(slide, Inches(4.3), Inches(6.05), STEEL, "Más de 15")
    if empty:
        text_box(slide, Inches(0.6), Inches(6.4), Inches(7.8), Inches(0.4),
                 f"{number(empty[1])} clusters PROD no aparecen en el export de plataforma ({money(empty[3])} al mes): "
                 "por revisar.", size=10, italic=True)

    kpi(slide, Inches(8.85), Inches(1.35), Inches(1.95), Inches(1.45), percent(p.get("clusters_le15_share")),
        "bajo el umbral de 15", value_size=26)
    kpi(slide, Inches(10.95), Inches(1.35), Inches(1.95), Inches(1.45), money(p.get("cost_le3")),
        "al mes, clusters con 3 o menos", value_size=20)
    rows = [["Microservicios", "Clusters", "Costo/mes"]] + [[item[0], number(item[1]), money(item[3])]
                                                             for item in buckets]
    table(slide, Inches(8.85), Inches(3.05), Inches(4.05), rows, [0.38, 0.24, 0.38], row_height=0.36)


def slide_pattern4(deck, p, source):
    behaviour = p["behaviour"]
    counts = [behaviour.get(code, {}).get("ci") or 0 for code, _, _ in BEHAVIOURS]
    total = sum(counts)
    share = counts[0] / total if total else None
    slide = base_slide(deck, f"Patrón 4: {percent(share)} de los HPA se queda en su mínimo durante 3 meses", source, 4)
    heading(slide, Inches(0.45), Inches(1.3), Inches(6.5), "HPA por réplicas actuales (día de la foto)")
    replicas = p["replicas"]
    bar_chart(slide, Inches(0.3), Inches(1.65), Inches(6.6), Inches(4.2), [item[0] for item in replicas],
              [item[1] for item in replicas], XL_CHART_TYPE.COLUMN_CLUSTERED,
              [ORANGE if item[0].startswith("2 a 3") else STEEL for item in replicas])
    peak = max(replicas, key=lambda item: item[1]) if replicas else ("", 0, 0)
    callout(slide, Inches(0.45), Inches(5.95), Inches(6.4), Inches(0.75),
            f"El {percent(peak[2])} de los HPA corre con {peak[0]} réplicas", size=13)

    heading(slide, Inches(7.2), Inches(1.3), Inches(5.7), "Comportamiento en 3 meses (promedio en oficina)")
    data = CategoryChartData()
    data.categories = [label for _, label, _ in BEHAVIOURS]
    data.add_series("Comportamiento", counts)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(7.1), Inches(1.6), Inches(3.6), Inches(3.9),
                                   data).chart
    chart.font.name, chart.font.size = FONT, Pt(11)
    chart.has_title, chart.has_legend = False, False
    plot = chart.plots[0]
    for index, (_, _, color) in enumerate(BEHAVIOURS):
        point = plot.series[0].points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
    text_box(slide, Inches(7.1), Inches(3.15), Inches(3.6), Inches(0.8), percent(share), size=28, bold=True,
             color=ORANGE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for index, ((_, label, color), count) in enumerate(zip(BEHAVIOURS, counts)):
        legend_item(slide, Inches(10.75), Inches(2.0 + index * 0.55), color,
                    f"{label}: {percent(count / total if total else None)}", width=Inches(2.3))
    message = "Los mínimos de réplicas son capacidad que se paga 24/7"
    if isinstance(p.get("inflated"), (int, float)):
        message = (f"En el {percent(p['inflated'])} de esos HPA el mínimo supera lo que pide su CPU "
                   f"({percent(p.get('inflated_beyond_ha'))} aun dejando 2 réplicas por alta disponibilidad)")
    callout(slide, Inches(7.2), Inches(5.95), Inches(5.7), Inches(0.75), message, size=12)


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
