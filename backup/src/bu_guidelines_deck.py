#!/usr/bin/env python3
"""Build the backup guidelines deck: retention rules, price per TB against the vendors, and references.

Slides:
  1. Guidelines, platform-neutral (cloud or on-premise, any database engine, servers and file shares): the base
     short- and long-term rule, and where it changes by the data a system holds
  2. Decision tree (docs/diagrams/backup-decision-tree.drawio, exported to PNG) that gives each application a
     backup plan code, P0 to P4, and what each plan keeps
  3. How the base plan is configured in Azure (VM, SQL Database, Cosmos DB) and how VM snapshots are kept
  4. Price per TB-month in Central US for LRS and ZRS backups, net of the bank's discount, against the vendors
  5. References: the regulation behind slide 1, the examples each norm gives and the minimum time it asks to
     keep the data
     (Microsoft and pricing sources are in docs/)

Prices come from the public Azure Retail Prices API (no sign-in, no subscription data). Retention limits and
regulatory articles are fixed in this file, checked against the official texts in November 2026. The regulation
is listed on slide 5; every Microsoft and pricing source is in docs/guidelines-references.md. Vendor tariffs are
the ones quoted in the backup guidelines review (unit to be confirmed).

Usage:
  pip install python-pptx requests
  python backup/src/bu_guidelines_deck.py
  python backup/src/bu_guidelines_deck.py --output lineamientos_backup.pptx

Output: backup/reports/<date>-lineamientos/lineamientos_backup.pptx (or --output).
"""

import argparse
import math
import os
import sys
import time
from datetime import date

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

TOPIC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRICES_API = "https://prices.azure.com/api/retail/prices"
REGIONS = [("eastus2", "East US 2"), ("centralus", "Central US")]
GB_PER_TB = 1024

# Vendor tariffs in US$ per TB-month, as quoted. The unit (protected or stored TB) is still to be confirmed.
VENDORS = [("D3", 56.01, 7.92), ("Opción 1", 49.90, 12.84)]
# The bank's discount on Azure list prices, applied to every Azure price on slide 4 (vendors quote net prices).
DISCOUNT = 0.30

# (key, serviceName, productName, meterName) of every meter the deck prices. The deck shows LRS and ZRS only;
# the geo meters stay because bu_guidelines_cost.py prices the redundancy each database has today.
METERS = [
    ("vault_lrs", "Backup", "Backup", "Standard LRS Data Stored"),
    ("vault_zrs", "Backup", "Backup", "Standard ZRS Data Stored"),
    ("vault_grs", "Backup", "Backup", "Standard GRS Data Stored"),
    ("vault_ragrs", "Backup", "Backup", "Standard RA-GRS Data Stored"),
    ("files_vault_lrs", "Backup", "Backup", "Azure Files Vaulted LRS Data Stored"),
    ("files_vault_zrs", "Backup", "Backup", "Azure Files Vaulted ZRS Data Stored"),
    ("files_vault_grs", "Backup", "Backup", "Azure Files Vaulted GRS Data Stored"),
    ("files_vault_ragrs", "Backup", "Backup", "Azure Files Vaulted RA-GRS Data Stored"),
    ("archive_lrs", "Backup", "Backup", "Archive LRS Data Stored"),
    ("archive_grs", "Backup", "Backup", "Archive GRS Data Stored"),
    ("vm_instance", "Backup", "Backup", "Azure VM Protected Instance"),
    ("pg_instance", "Backup", "Backup", "PostgreSQL Protected Instance"),
    ("cosmos_instance", "Backup", "Backup", "Cosmos DB Protected Instance"),
    ("sqldb_pitr_lrs", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage", "LRS Data Stored"),
    ("sqldb_pitr_zrs", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage", "ZRS Data Stored"),
    ("sqldb_pitr_geo", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage", "RA-GRS Data Stored"),
    ("sqldb_pitr_ragzrs", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage",
     "RA-GZRS Data Stored"),
    ("sqldb_ltr_lrs", "SQL Database", "SQL Database - LTR Backup Storage", "Backup LRS Data Stored"),
    ("sqldb_ltr_zrs", "SQL Database", "SQL Database - LTR Backup Storage", "Backup ZRS Data Stored"),
    ("sqldb_ltr_geo", "SQL Database", "SQL Database - LTR Backup Storage", "Backup RA-GRS Data Stored"),
    ("sqldb_ltr_ragzrs", "SQL Database", "SQL Database - LTR Backup Storage", "Backup RA-GZRS Data Stored"),
    ("sqlmi_pitr_lrs", "SQL Managed Instance", "SQL Managed Instance PITR Backup Storage", "LRS Data Stored"),
    ("sqlmi_pitr_zrs", "SQL Managed Instance", "SQL Managed Instance PITR Backup Storage", "ZRS Data Stored"),
    ("sqlmi_pitr_geo", "SQL Managed Instance", "SQL Managed Instance PITR Backup Storage", "RA-GRS Data Stored"),
    ("sqlmi_ltr_lrs", "SQL Managed Instance", "SQL Managed Instance - LTR Backup Storage", "LTR Backup LRS Data Stored"),
    ("sqlmi_ltr_zrs", "SQL Managed Instance", "SQL Managed Instance - LTR Backup Storage", "LTR Backup ZRS Data Stored"),
    ("sqlmi_ltr_geo", "SQL Managed Instance", "SQL Managed Instance - LTR Backup Storage",
     "LTR Backup RA-GRS Data Stored"),
    ("pg_pitr", "Azure Database for PostgreSQL", "Azure Database for PostgreSQL Flexible Server Backup Storage",
     "Backup Storage LRS Data Stored"),
    ("mysql_pitr", "Azure Database for MySQL", "Azure Database for MySQL Flexible Server Backup Storage",
     "Backup Storage LRS Data Stored"),
    ("mysql_pitr_zrs", "Azure Database for MySQL", "Azure Database for MySQL Flexible Server Backup Storage",
     "Backup Storage ZRS Data Stored"),
    ("hs_pitr_zrs", "SQL Database", "SQL Database SingleDB Hyperscale - Storage", "Backup - ZRS Data Stored"),
    ("hs_ltr_zrs", "SQL Database", "SQL Database SingleDB Hyperscale - LTR Backup Storage", "Backup ZRS Data Stored"),
    ("cosmos_30d", "Azure Cosmos DB", "Azure Cosmos DB - PITR", "Cont 30D Bckp Continuous Backup"),
    ("cosmos_restore", "Azure Cosmos DB", "Azure Cosmos DB - PITR", "Cont 30D Bckp Data Restore"),
    ("cosmos_periodic_lrs", "Azure Cosmos DB", "Azure Cosmos DB Snapshot", "Periodic Backup - LRS Data Stored"),
    ("cosmos_periodic_zrs", "Azure Cosmos DB", "Azure Cosmos DB Snapshot", "Periodic Backup - ZRS Data Stored"),
    ("cosmos_periodic_ragrs", "Azure Cosmos DB", "Azure Cosmos DB Snapshot", "Periodic Backup - RA-GRS Data Stored"),
]
# Protected instance fees are charged per block of source data (Azure Backup pricing page).
INSTANCE_BLOCK_GB = {"vm_instance": 500, "pg_instance": 250, "cosmos_instance": 500}


# Palette of the team's deck: royal blue header ending in a rust tip, navy cards.
NAVY = RGBColor(0x1E, 0x3A, 0x78)
RUST = RGBColor(0xB0, 0x45, 0x2C)
RUST_DARK = RGBColor(0x7A, 0x2C, 0x1F)
CARD = RGBColor(0x1F, 0x33, 0x55)
ORANGE = RGBColor(0xF2, 0xA1, 0x5A)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
GREY = RGBColor(0x5B, 0x67, 0x70)
LIGHT = RGBColor(0xEE, 0xF1, 0xF5)
SOFT = RGBColor(0xC5, 0xCD, 0xD6)
SECTION = RGBColor(0xDD, 0xE3, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"


# (norm, what the deck uses, examples the norm itself gives, minimum period, official source, link). The examples are
# taken from the official texts: art. 183 lists none; SBS 2660-2015 art. 49.1-49.3, 55 and 58; SBS 504-2021
# art. 12.3 f) and h); BCRP 0011-2026 Anexo V sec. VI and the definitions in art. 2.
REGULATION_REFERENCES = [
    ("Ley 26702, Ley General del Sistema Financiero", "Art. 183: libros y documentos",
     "• No da ejemplos: aplica a todos los libros y documentos\n"
     "• Si hay juicio, los documentos del caso se guardan mientras dure\n"
     "• Permite microfilm u otros medios análogos",
     "10 años\n(más si hay juicio)", "SPIJ, texto con notas a mayo 2026",
     "https://spij.minjus.gob.pe/spij-ext-web/#/detallenorma/H773064"),
    ("Res. SBS 2660-2015, Gestión de Riesgos de LA/FT", "Arts. 49 y 54: registro de operaciones y copia trimestral",
     "• Retiros de fondos, cobro de cheques, cheques de gerencia\n"
     "• Desembolsos de préstamos y financiamiento por descuento\n"
     "• Transferencias nacionales e internacionales\n"
     "• Compra y venta de divisas en efectivo desde US$5 mil",
     "Registro: 10 años\nCopia trimestral: 5 años (20 copias)", "Portal de normas SBS, texto consolidado v5",
     "https://intranet2.sbs.gob.pe/dv_int_cn/1540/v5.0/Adjuntos/2660-2015.r.pdf"),
    ("Res. SBS 2660-2015, Gestión de Riesgos de LA/FT", "Arts. 55 y 58: KYC, contrapartes e inusuales",
     "• Vinculación del cliente y debida diligencia (KYC)\n"
     "• Vinculación y operaciones con contrapartes y proveedores\n"
     "• Operaciones inusuales y el análisis que las sustenta\n"
     "• Políticas y procedimientos de prevención LA/FT",
     "10 años", "Portal de normas SBS, texto consolidado v5",
     "https://intranet2.sbs.gob.pe/dv_int_cn/1540/v5.0/Adjuntos/2660-2015.r.pdf"),
    ("Res. SBS 504-2021, Seguridad de la Información y Ciberseguridad",
     "Art. 12, num. 3, lit. f) y h): copias de respaldo y pruebas",
     "• Copias de información, software y lo necesario para restablecerlos\n"
     "• Recuperación ante malware\n"
     "• Incidentes internos o externos que afecten el ambiente productivo",
     "No fija plazo", "Portal de normas SBS, texto consolidado v5",
     "https://intranet2.sbs.gob.pe/dv_int_cn/2046/v5.0/Adjuntos/504-2021.R.pdf"),
    ("Circular BCRP 0011-2026, Reglamento del Sistema LBTR", "Anexo V, sección VI: respaldo diario y prueba anual",
     "• Información de la operación diaria del LBTR\n"
     "• Órdenes de transferencia de fondos y cartas orden\n"
     "• Aplicación Participante LBTR del banco",
     "Respaldos: 10 años", "Circulares vigentes del BCRP (vigente desde junio 2026)",
     "https://www.bcrp.gob.pe/docs/Transparencia/Normas-Legales/Circulares/2026/circular-0011-2026-bcrp.pdf"),
]


# ---------------------------------------------------------------------------
# Prices
# ---------------------------------------------------------------------------


def get_items(price_filter):
    """Every price row that matches an OData filter, following the pages."""
    url, params, items = PRICES_API, {"$filter": price_filter}, []
    while url:
        for attempt in range(6):  # the public API throttles bursts with 429
            response = requests.get(url, params=params, timeout=60)
            if response.status_code != 429:
                break
            time.sleep(min(float(response.headers.get("Retry-After", 2 ** attempt * 5)), 60))
        response.raise_for_status()
        body = response.json()
        items += body.get("Items", [])
        url, params = body.get("NextPageLink"), None
    return items


def meter_price(items, region, meter):
    """Retail price of one meter among the rows of its product, or None when there is no row for it."""
    prices = {item["retailPrice"] for item in items if item["meterName"] == meter and not item.get("tierMinimumUnits")}
    if len(prices) > 1:
        raise SystemExit(f"{meter} in {region} has several prices {sorted(prices)}; check the meter definition.")
    return prices.pop() if prices else None


def per_tb(prices):
    """US$ per TB-month for storage meters, and US$ per protected TB for protected instance fees."""
    values = {}
    for key, price in prices.items():
        block = INSTANCE_BLOCK_GB.get(key)
        values[key] = price / block * GB_PER_TB if block else price * GB_PER_TB
    return values


def load_prices():
    """One request per product and region (not per meter), so the public API does not throttle the run."""
    prices, missing, products = {}, [], {}
    for region, _ in REGIONS:
        prices[region] = {}
        for key, service, product, meter in METERS:
            if (region, service, product) not in products:
                products[region, service, product] = get_items(
                    f"serviceName eq '{service}' and productName eq '{product}' and armRegionName eq '{region}' "
                    "and priceType eq 'Consumption'")
            value = meter_price(products[region, service, product], region, meter)
            if value is None:
                missing.append(f"{meter} ({region})")
            prices[region][key] = value
    if missing:
        raise SystemExit("The price API returned no price for: " + "; ".join(missing))
    return {region: per_tb(values) for region, values in prices.items()}, prices


# ---------------------------------------------------------------------------
# Formatting and drawing helpers
# ---------------------------------------------------------------------------


def usd(value, decimals=1):
    text = f"{value:,.{decimals}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"US${text}"


def amount(value):
    return usd(value)[3:]


def text_box(slide, left, top, width, height, text, size=12, bold=False, color=GREY, align=PP_ALIGN.LEFT,
             italic=False, anchor=MSO_ANCHOR.TOP):
    """Text box; text is a string, a list of lines, or a list of lines made of (text, bold, color) runs."""
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = Inches(0.04)
    for index, line in enumerate(text if isinstance(text, list) else [text]):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        for part in line if isinstance(line, list) else [(line, bold, color)]:
            run = paragraph.add_run()
            run.text = part[0]
            run.font.name, run.font.size, run.font.bold, run.font.italic = FONT, Pt(size), part[1], italic
            run.font.color.rgb = part[2]
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
    """Title band as in the team's deck: a navy arrow over a rust band that darkens to the right."""
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.slide_width, Inches(0.95))
    band.line.fill.background()
    band.shadow.inherit = False
    band.fill.gradient()
    band.fill.gradient_angle = 0
    stops = band.fill.gradient_stops
    stops[0].color.rgb, stops[0].position = RUST, 0.7
    stops[1].color.rgb, stops[1].position = RUST_DARK, 1.0
    arrow = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, 0, 0, int(deck.slide_width * 0.8), Inches(0.95))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = NAVY
    arrow.line.fill.background()
    arrow.shadow.inherit = False
    arrow.adjustments[0] = 0.45
    text_box(slide, Inches(0.45), 0, Inches(10.2), Inches(0.95), title.upper(), size=19, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)


def base_slide(deck, title, source):
    slide = deck.slides.add_slide(deck.slide_layouts[6])  # blank layout
    banner(slide, deck, title)
    text_box(slide, Inches(0.35), Inches(7.08), Inches(11.8), Inches(0.3), source, size=8, italic=True)
    text_box(slide, Inches(12.4), Inches(7.08), Inches(0.55), Inches(0.3), str(len(deck.slides)), size=9,
             align=PP_ALIGN.RIGHT)
    return slide


def heading(slide, left, top, width, text, size=13, color=NAVY):
    text_box(slide, left, top, width, Inches(0.32), text, size=size, bold=True, color=color)


def style_cell(cell, text, size, bold=False, color=GREY, fill=WHITE, align=PP_ALIGN.CENTER):
    cell.text = text or " "  # an empty cell has no run to style
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.05)
    cell.margin_top = cell.margin_bottom = Inches(0.015)
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    for paragraph in cell.text_frame.paragraphs:  # a value with line breaks has one paragraph per line
        paragraph.alignment = align
        for run in paragraph.runs:
            run.font.name, run.font.size, run.font.bold = FONT, Pt(size), bold
            run.font.color.rgb = color


def table(slide, left, top, width, header, rows, widths, row_height=0.3, size=10, header_fill=CARD):
    """Rows are lists of cell values or of (text, options) pairs; a row of one value is a section row, and an
    option "span" merges a cell over the next columns."""
    shape = slide.shapes.add_table(len(rows) + 1, len(header), left, top, width, Inches(row_height * (len(rows) + 1)))
    grid = shape.table
    for column, share in enumerate(widths):
        grid.columns[column].width = Emu(int(width * share))
    for c, value in enumerate(header):
        style_cell(grid.cell(0, c), value, size, bold=True, color=WHITE, fill=header_fill,
                   align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    for r, values in enumerate(rows, start=1):
        grid.rows[r].height = Inches(row_height)
        if len(values) == 1:
            grid.cell(r, 0).merge(grid.cell(r, len(header) - 1))
            style_cell(grid.cell(r, 0), values[0], size, bold=True, color=NAVY, fill=SECTION, align=PP_ALIGN.LEFT)
            continue
        c = 0
        for value in values:
            text, options = value if isinstance(value, tuple) else (value, {})
            span = options.get("span", 1)
            if span > 1:
                grid.cell(r, c).merge(grid.cell(r, c + span - 1))
            style_cell(grid.cell(r, c), text, options.get("size", size), bold=options.get("bold", c == 0),
                       color=options.get("color", NAVY if c == 0 else GREY),
                       fill=options.get("fill", LIGHT if r % 2 else WHITE),
                       align=PP_ALIGN.LEFT if c == 0 or options.get("left") else PP_ALIGN.CENTER)
            c += span
    return grid


def note_box(slide, left, top, width, height, text, size=10, fill=LIGHT, color=GREY):
    box = rounded(slide, left, top, width, height, fill, radius=0.1)
    if fill == LIGHT:
        box.line.color.rgb = SOFT
    text_box(slide, left + Inches(0.15), top, width - Inches(0.3), height, text, size=size, color=color,
             anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def section_bar(slide, left, top, width, text, fill):
    """Coloured heading bar: navy for the base rule, rust for the regulated scopes, so the two never mix."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.36))
    bar.fill.solid()
    bar.fill.fore_color.rgb = fill
    bar.line.fill.background()
    bar.shadow.inherit = False
    text_box(slide, left + Inches(0.1), top, width - Inches(0.2), Inches(0.36), text, size=12, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def divider(slide, left, top, height):
    line = slide.shapes.add_connector(1, left, top, left, top + height)  # 1 = straight connector
    line.line.color.rgb = SOFT
    line.line.width = Pt(1.5)


def slide_guidelines(deck):
    """Platform-neutral rule: the same retention for cloud and on-premise, databases, servers and file shares."""
    slide = base_slide(deck, "Lineamientos de backup",
                       "Fuente: Ley 26702; Res. SBS 2660-2015 y 504-2021; Circular BCRP 0011-2026 (revisadas nov-2026). "
                       "Plazos legales a validar por Cumplimiento.")

    # Left: the rule every production system follows, whatever the platform.
    left, width = Inches(0.35), Inches(4.75)
    section_bar(slide, left, Inches(1.08), width, "Lineamiento base: todo PROD", NAVY)
    rows = [
        ["Bases de datos\n(cualquier motor)", "Volver a cualquier\nmomento de los\núltimos 7 días",
         "4 semanales\n12 mensuales"],
        ["Máquinas virtuales\ny servidores", "1 copia diaria,\nse guardan 7", "4 semanales\n12 mensuales"],
        ["Archivos\ncompartidos", "1 copia diaria,\nse guardan 7", "4 semanales\n12 mensuales"],
    ]
    grid = table(slide, left, Inches(1.55), width, ["Qué se respalda", "Corto plazo", "Largo plazo"], rows,
                 [0.34, 0.36, 0.30], row_height=0.78, size=10)
    grid.rows[0].height = Inches(0.4)
    note_box(slide, left, Inches(4.55), width, Inches(2.38), [
        [("Corto plazo: ", True, NAVY),
         ("recupera errores recientes (borrado, corrupción, ransomware). En bases de datos se llama "
          "recuperación a un punto en el tiempo: se vuelve al minuto anterior al error.", False, GREY)],
        [("", False, GREY)],
        [("Largo plazo: ", True, NAVY),
         ("copias completas guardadas meses o años. Cubren la conservación que pide la norma y los errores "
          "que se detectan tarde.", False, GREY)],
        [("", False, GREY)],
        [("Aplica igual en nube y on-premise. ", True, NAVY),
         ("La lámina 3 muestra cómo se configura en Azure.", False, GREY)]], size=9.5)

    divider(slide, Inches(5.3), Inches(1.08), Inches(5.85))

    # Right: where the rule changes, by the data the system holds. Rust, not navy, so it reads as the regulation.
    left, width = Inches(5.5), Inches(7.48)
    section_bar(slide, left, Inches(1.08), width, "Cuándo varía: según el dato que guarda el sistema", RUST)
    plain = {"fill": RGBColor(0xF6, 0xE9, 0xE4)}
    action = {"left": True, "size": 8.5}
    keep_data = ("• Si la aplicación guarda el dato 10 años: basta el lineamiento base\n"
                 "• Si lo borra antes: agregar 1 backup anual que se guarde 10 años")
    scopes = [
        [("Cualquier otro dato", plain), ("SBS 504-2021,\nart. 12", plain),
         ("Estrategia de respaldo y pruebas de restauración", plain),
         ("Lineamiento base", {**plain, **action})],
        ["Libros y documentos contables", "Ley 26702,\nart. 183",
         "Conservar el dato 10 años (más si hay juicio)", (keep_data, action)],
        ["Registro de operaciones\n(≥ US$10 mil; múltiples ≥ US$50 mil/mes)", "SBS 2660-2015,\narts. 49 y 54",
         "Copia al cierre de cada trimestre, guardada 5 años; entrega en 48 h",
         ("• Agregar 1 copia por trimestre, guardada 5 años (20 copias), inmutable\n"
          "• Si la aplicación borra el dato antes de 10 años: 1 backup anual por 10 años", action)],
        ["KYC, contrapartes y operaciones inusuales", "SBS 2660-2015,\narts. 55 y 58",
         "Conservar el dato 10 años", (keep_data, action)],
        ["Sistemas del LBTR", "BCRP 0011-2026,\nAnexo V, sec. VI",
         "Respaldo diario, prueba anual y respaldos guardados 10 años",
         ("• Agregar 1 backup anual que se guarde 10 años\n• Probar la restauración 1 vez al año y dejarla "
          "documentada", action)],
    ]
    grid = table(slide, left, Inches(1.55), width,
                 ["Ámbito (dato)", "Norma", "Qué exige", "Qué se agrega al lineamiento base"],
                 scopes, [0.2, 0.14, 0.24, 0.42], row_height=0.8, size=9, header_fill=RUST_DARK)
    grid.rows[0].height = Inches(0.4)
    note_box(slide, left, Inches(6.18), width, Inches(0.75), [
        [("Recomendado en los ámbitos regulados: ", True, RUST),
         ("copia inmutable (nadie puede borrarla ni cambiarla antes de su vencimiento) y prueba de restauración "
          "documentada.", False, GREY)]], size=9.5)
    slide.notes_slide.notes_text_frame.text = (
        "El lineamiento es el mismo para cualquier plataforma (Azure, otra nube u on-premise) y cualquier motor de "
        "base de datos. Corto plazo: en bases de datos, backup completo más los backups del log de transacciones, "
        "que permiten restaurar a cualquier minuto de los últimos 7 días; en Azure se llama PITR. En servidores y "
        "archivos, 1 copia diaria que se guarda 7 días. "
        "Largo plazo: cantidad de copias completas que se guardan (4 semanales = las últimas 4 semanas; "
        "12 mensuales = el último año). En Azure SQL se llama LTR; los proveedores lo llaman Archive. "
        "La regla base aplica a todo PROD; un sistema solo cambia si guarda uno de los datos de la tabla de ámbitos. "
        "La Ley 26702 art. 183 y la Res. SBS 2660-2015 arts. 55 y 58 obligan a conservar el dato, no el backup: si la "
        "aplicación no borra esos registros antes de 10 años, el backup se queda en la regla base. "
        "Solo dos normas piden backups: la Res. SBS 2660-2015 art. 54 (copia al final de cada trimestre, agrupada por "
        "5 años: 20 copias) y la Circular BCRP 0011-2026 Anexo V sección VI (conservar los respaldos 10 años). "
        "Cumplimiento debe confirmar con qué frecuencia se guardan los respaldos de 10 años del LBTR (la circular no "
        "la fija) y qué aplicaciones guardan el registro de operaciones y el LBTR. "
        "Ninguna norma fija un plazo para el corto plazo; 7 días aplica a todo. "
        "El LSEG_008 (Lineamiento de Respaldo y Retención de Información) no fija plazos de retención; este "
        "lineamiento los fija.")


TREE_PNG = os.path.join(TOPIC_DIR, "docs", "diagrams", "backup-decision-tree.png")
# Backup plans that the decision tree assigns. (code, name, short term, long term, also, fill). The rules come from
# slide 1; the tree itself lives in docs/diagrams/backup-decision-tree.drawio and is exported to PNG.
PLANS = [
    ("P0", "Sin backup", "–", "–", "No guarda datos: se reinstala. Los logs tienen su propio plazo", GREY),
    ("PN", "No productivo", "7 días", "–", "–", SOFT),
    ("P1", "Base", "7 días", "4 semanales + 12 mensuales", "–", NAVY),
    ("P2", "Base + 10 años", "7 días", "Base + 1 anual por 10 años",
     "Solo si la aplicación borra libros, KYC u operaciones inusuales antes de 10 años", RUST),
    ("P3", "Registro de operaciones", "7 días", "Base + 1 por trimestre por 5 años (20 copias)",
     "Copia inmutable. Si la aplicación borra el dato antes de 10 años: + 1 anual por 10 años", RUST),
    ("P4", "LBTR", "7 días", "Base + 1 anual por 10 años", "Prueba de restauración 1 vez al año, documentada", RUST),
]


def slide_layers(deck):
    """Decision tree that gives each application its backup plan, and what each plan keeps."""
    slide = base_slide(deck, "Qué plan de backup le toca a cada aplicación",
                       "El árbol aplica la lámina 1. Fuente del árbol: docs/diagrams/backup-decision-tree.drawio.")
    width = Inches(12.63)
    picture = slide.shapes.add_picture(TREE_PNG, Inches(0.35), Inches(1.05), width=width)
    top = picture.top + picture.height + Inches(0.15)
    rows = []
    for code, name, short, long_term, also, fill in PLANS:
        text = NAVY if fill == SOFT else WHITE
        rows.append([(code, {"fill": fill, "color": text, "bold": True}), (name, {"bold": True, "color": NAVY}),
                     short, long_term, (also, {"left": True, "size": 8.5})])
    grid = table(slide, Inches(0.35), top, width, ["Código", "Plan", "Corto plazo", "Largo plazo", "Además"], rows,
                 [0.07, 0.17, 0.1, 0.28, 0.38], row_height=0.36, size=9)
    grid.rows[0].height = Inches(0.3)
    text_box(slide, Inches(0.35), Inches(6.72), width, Inches(0.3), [[
        ("Si aplica más de una respuesta, se toma el plan más alto (P4 > P3 > P2 > P1). ", True, NAVY),
        ("El dato regulado suele estar en las bases de BS, Core o mainframe.", False, GREY)]], size=9)
    slide.notes_slide.notes_text_frame.text = (
        "El árbol se recorre por aplicación y da un código de plan. Primero: si la aplicación no guarda datos propios "
        "(canal e integración solo pasan la operación; las plataformas de seguridad guardan logs), no lleva backup: "
        "se reinstala, y la traza de cada operación queda en los logs con su propia retención. Las bases que guardan "
        "datos regulados suelen estar en los servicios de negocio (BS), Core o mainframe. "
        "Ambientes no productivos: solo corto plazo. En producción, el plan depende del dato que guarda la base: "
        "LBTR (Circular BCRP 0011-2026), registro de operaciones (Res. SBS 2660-2015 arts. 49 y 54), libros "
        "contables, KYC u operaciones inusuales (Ley 26702 art. 183; Res. SBS 2660-2015 arts. 55 y 58), o cualquier "
        "otro dato (plan base). Si aplica más de una respuesta, se toma el plan más alto. "
        "La ley pide conservar libros, KYC e inusuales 10 años: si la aplicación ya los guarda ese tiempo, basta el "
        "plan base; si los borra antes, el backup anual por 10 años los cubre. "
        "El LSEG_008 no fija plazos de retención; este árbol los fija por aplicación.")



def slide_azure(deck):
    """How the base rule is configured in each Azure service, and how VM snapshots of Azure Backup are kept."""
    slide = base_slide(deck, "Cómo se configura en Azure",
                       "Fuente: Microsoft Learn (revisado nov-2026): retención, snapshots, redundancia y tiempos de "
                       "restauración de cada servicio.")

    terms = [
        ("PITR = recuperación a un punto en el tiempo",
         "Azure guarda la base y sus cambios cada pocos minutos. Permite volver a cualquier minuto de los "
         "últimos 7 días. Es el corto plazo."),
        ("LTR = retención de largo plazo",
         "Copias completas que Azure guarda aparte de la base: semanales, mensuales o anuales, hasta 10 años. "
         "Es el largo plazo."),
        ("Bóveda (vault) de Azure Backup",
         "Donde se guardan las copias de VMs y Cosmos DB. Inmutable: nadie la borra antes de su vencimiento."),
    ]
    for index, (title, text) in enumerate(terms):
        left = Inches(0.35 + index * 4.25)
        rounded(slide, left, Inches(1.08), Inches(4.13), Inches(0.95), CARD)
        text_box(slide, left + Inches(0.12), Inches(1.08), Inches(3.89), Inches(0.95),
                 [[(title, True, ORANGE)], [(text, False, WHITE)]], size=9, anchor=MSO_ANCHOR.MIDDLE)

    # Left: the base rule per Azure service.
    left, width = Inches(0.35), Inches(6.7)
    section_bar(slide, left, Inches(2.2), width, "Plan base por servicio (P1)", NAVY)
    rows = [
        ["Azure Backup: VM", "1 diario, se guardan 7", "4 semanales, 12 mensuales (bóveda)"],
        ["Azure SQL Database", "PITR 7 días", "LTR: 4 semanales, 12 mensuales"],
        ["Cosmos DB", "Continuo 7 días (sin costo)", "4 semanales, 12 mensuales (bóveda)"],
    ]
    grid = table(slide, left, Inches(2.65), width, ["Servicio", "Corto plazo", "Largo plazo"], rows,
                 [0.32, 0.3, 0.38], row_height=0.62, size=9.5)
    grid.rows[0].height = Inches(0.36)
    text_box(slide, left, Inches(4.95), width, Inches(0.5), [[
        ("Los otros planes ", True, NAVY),
        ("(P2 a P4) agregan copias de largo plazo en el mismo LTR o bóveda: anuales por 10 años o mensuales por 5 "
         "años.", False, GREY)]], size=9)

    divider(slide, Inches(7.25), Inches(2.2), Inches(3.75))

    # Right: VM snapshots that Azure Backup creates; the only snapshots this guideline covers.
    left, width = Inches(7.45), Inches(5.53)
    section_bar(slide, left, Inches(2.2), width, "Snapshots de VM (Azure Backup)", NAVY)
    action = {"left": True, "size": 9}
    rows = [
        ["Recuperar rápido", "2 días", ("Vuelve en minutos; para días anteriores se usa la bóveda", action)],
        ["Antes de un cambio", "Hasta 7 días\ndespués", ("Siempre con fecha de borrado y un responsable", action)],
    ]
    grid = table(slide, left, Inches(2.65), width, ["Para qué", "Cuánto se guarda", "Regla"], rows,
                 [0.3, 0.22, 0.48], row_height=0.62, size=9.5)
    grid.rows[0].height = Inches(0.36)
    text_box(slide, left, Inches(4.35), width, Inches(0.9), [[
        ("Se pagan cada mes: ", True, NAVY),
        ("cada snapshot guarda lo que cambió desde el anterior. Más días de snapshot cuestan más y no reemplazan "
         "la copia de la bóveda.", False, GREY)]], size=9)

    note_box(slide, Inches(0.35), Inches(6.12), Inches(12.63), Inches(0.82), [
        [("Redundancia del backup: ", True, NAVY),
         ("PROD: ZRS (copias en 3 zonas de la región). No productivo: LRS. Archive solo existe en bóvedas LRS.",
          False, GREY)],
        [("Tiempos: ", True, NAVY),
         ("SQL restaura normalmente en < 12 h; Cosmos DB, 30 a 90 min por TB; Archive tarda hasta 15 h en "
          "rehidratar, dentro de las 48 h hábiles que da la SBS.", False, GREY)]], size=9.5)
    slide.notes_slide.notes_text_frame.text = (
        "PITR (point-in-time restore): Azure toma backups completos, diferenciales y del log cada 5 a 10 minutos en "
        "SQL y con eso restaura la base a cualquier minuto dentro de los días de retención. Es el corto plazo. "
        "LTR (long-term retention): copias completas que se guardan aparte, con su propia retención (semanal, "
        "mensual o anual, hasta 10 años). Es el largo plazo. SQL LTR no tiene frecuencia trimestral: el plan P3 usa "
        "60 mensuales para cubrir los 5 años. "
        "Cosmos DB: continuo 7 días es gratis (solo se cobra restaurar); el largo plazo va a un Backup vault. Hoy 30 "
        "de las 41 cuentas Cosmos de producción ya usan continuo 7 días y 11 siguen en periódico. "
        "Snapshots de VM: Azure Backup guarda snapshots locales para restaurar rápido, de 1 a 5 días con la política "
        "Standard (2 por defecto) y de 1 a 30 con la Enhanced (7 por defecto). Con Enhanced el primer snapshot es una "
        "copia completa del disco y los siguientes guardan solo los cambios (normalmente 2% a 7% al día). "
        "Redundancia: LRS guarda 3 copias en un datacenter; ZRS, 3 copias en zonas distintas de la misma región. En "
        "SQL, ZRS cuesta lo mismo que LRS. Las bóvedas ZRS no admiten Archive.")


def slide_prices(deck, tb, retrieved, discount):
    prices = tb["centralus"]
    small = {"left": True, "size": 8}
    transfer = f"~{GB_PER_TB * 8 / 3600:.1f} h/TB a 1 Gbps".replace(".", ",")
    vault_hours = f"~{GB_PER_TB * 1024 / 30 / 3600:.0f} h/TB"

    def priced(label, keys, rpo, rto, extra, best=False):
        """keys = meter keys for LRS and ZRS; None when the option doesn't exist."""
        options = {"color": GREEN, "bold": True} if best else {}
        cells = []
        for key in keys:
            if key is None:
                cells.append(("–", {}))
            else:
                cells.append((amount(prices[key]), options))
        return [label, *cells, (rpo, {"size": 8}), (rto, small), (extra, small)]

    def vendor(name, value, tier, rto):
        return [(f"{name} · {tier}", {"color": RUST}),
                (f"{amount(value)}  (según su servicio)", {"color": RUST, "bold": True, "span": 2}),
                ("Según su política", {"size": 8}), (rto, small), ("Confirmar TB protegido o almacenado", small)]

    sql_label = "SQL Database"
    rows = [
        ["Corto plazo: recuperación operativa"],
        priced("Azure Backup · VM (bóveda)", ("vault_lrs", "vault_zrs"),
               "24 h (4 h Enhanced)", f"Snapshot: más rápido; bóveda: {vault_hours}",
               f"+ {usd(prices['vm_instance'])}/TB protegido"),
        priced(f"{sql_label} · PITR", ("sqldb_pitr_lrs", "sqldb_pitr_zrs"),
               "5 a 10 min", "Normalmente < 12 h", "Gratis hasta el tamaño de la base"),
        ["Cosmos DB · continuo 7 días", ("Sin costo", {"span": 2, "color": GREEN, "bold": True}),
         ("~100 s", {"size": 8}), ("~30 a 90 min por TB", small), ("Solo se cobra restaurar", small)],
        *[vendor(name, active, "Active", f"Restaurar + traer: {transfer}") for name, active, _ in VENDORS],
        ["Largo plazo: retención"],
        priced("Azure Backup · Archive (VM)", ("archive_lrs", None), "1 mes",
               f"Rehidratar ≤ 15 h + {vault_hours}", "Puntos > 6 meses; mínimo 180 días", best=True),
        priced(f"{sql_label} · LTR", ("sqldb_ltr_lrs", "sqldb_ltr_zrs"),
               "1 semana", "Normalmente < 12 h", "Solo backups completos"),
        priced("Cosmos DB · Backup vault", ("vault_lrs", "vault_zrs"),
               "1 día", "~30 a 90 min por TB", f"+ {usd(prices['cosmos_instance'])}/TB protegido"),
        *[vendor(name, archive, "Archive", f"Su Archive + {transfer}") for name, _, archive in VENDORS],
    ]

    slide = base_slide(deck, "Costo por TB al mes: Azure vs proveedores",
                       f"Fuente: Azure Retail Prices API (Central US), precio de lista consultado el {retrieved}, menos "
                       f"{discount:.0%} de descuento del banco; 1 TB = 1.024 GB. Proveedores: tarifas cotizadas.")
    text_box(slide, Inches(0.35), Inches(1.0), Inches(12.63), Inches(0.28), [[
        ("Unidad: US$ por TB almacenado al mes", True, NAVY),
        (f", Central US, Azure con {discount:.0%} de descuento. – = el servicio no ofrece esa redundancia.",
         False, GREY)]], size=9)
    grid = table(slide, Inches(0.35), Inches(1.3), Inches(12.63),
                 ["Opción", "LRS", "ZRS", "RPO", "RTO aprox. (restaurar)", "Qué más se cobra"], rows,
                 [0.25, 0.08, 0.08, 0.11, 0.2, 0.28], row_height=0.28, size=8.5)
    grid.rows[0].height = Inches(0.36)
    for r, values in enumerate(rows, start=1):
        if len(values) == 1:
            grid.rows[r].height = Inches(0.26)

    vault_protected = prices["vault_zrs"] + prices["vm_instance"]
    actives = [active for _, active, _ in VENDORS]
    archives = [archive for _, _, archive in VENDORS]
    takeaways = [
        ("Largo plazo", f"Archive LRS de Azure: {usd(prices['archive_lrs'])} por TB, "
                        f"{'más barato' if prices['archive_lrs'] < min(archives) else 'frente a'} que el Archive de "
                        f"los proveedores ({' y '.join(amount(value) for value in archives)})."),
        ("PITR vs LTR", f"En SQL, el PITR ZRS cuesta {prices['sqldb_pitr_zrs'] / prices['sqldb_ltr_zrs']:.0f} veces "
                        f"el LTR ZRS por TB ({amount(prices['sqldb_pitr_zrs'])} vs "
                        f"{amount(prices['sqldb_ltr_zrs'])}): PITR corto y LTR para retener."),
        ("Corto plazo", f"Bóveda ZRS + instancia VM = {usd(vault_protected)} por TB protegido, "
                        f"{'más barato' if vault_protected < min(actives) else 'frente a'} que los proveedores "
                        f"({amount(min(actives))}–{amount(max(actives))}), que además suman la transferencia "
                        f"({transfer})."),
    ]
    for index, (title, text) in enumerate(takeaways):
        left = Inches(0.35 + index * 4.25)
        rounded(slide, left, Inches(6.0), Inches(4.13), Inches(0.8), CARD)
        text_box(slide, left + Inches(0.12), Inches(6.0), Inches(3.89), Inches(0.8),
                 [[(title + ": ", True, ORANGE), (text, False, WHITE)]], size=9, anchor=MSO_ANCHOR.MIDDLE)
    meters = "; ".join(f"{meter} ({product})" for _, _, product, meter in METERS)
    slide.notes_slide.notes_text_frame.text = (
        f"Precios de Azure en Central US: lista de la API menos {discount:.0%} de descuento del banco. East US 2 cuesta "
        "lo mismo en bóvedas y entre 14% y 23% menos en bases de datos. "
        "Cada precio por GB-mes se multiplica por 1.024. Las instancias protegidas se cobran por bloque de datos de "
        "origen: VM US$10 por 500 GB, Cosmos DB US$25 por 500 GB (precio de lista). "
        "Redundancia: solo LRS y ZRS, las que usa el banco. Cosmos DB continuo 7 días no cobra almacenamiento de "
        "backup, solo cada restauración. Archive no admite bóvedas ZRS. "
        "El PITR de SQL solo cobra lo que excede el almacenamiento gratis (el tamaño de la base). "
        "Azure Backup guarda incrementales: los TB almacenados son menos que los TB protegidos. "
        f"RTO de VM desde la bóveda: Microsoft considera óptimos 30 MB/s y da el ejemplo de 2 TB en hasta 20 h; a "
        f"30 MB/s, 1 TB toma {vault_hours}. Cosmos DB: restaurar 1 TB toma típicamente 30 a 90 min (FAQ de "
        "restauración en la misma cuenta). "
        "RPO y RTO: SQL Database (log cada 5 a 10 min; restaurar normalmente toma menos de 12 h, una base grande o "
        "muy activa puede tomar varias horas); Cosmos DB (cambios respaldados en 100 s); Archive "
        "(rehidratar hasta 15 h en prioridad Standard, menos de 1 h en High para objetos menores a 10 GB). "
        f"Proveedores: el dato vuelve por la red; {transfer} es el mínimo teórico (1.024 GB x 8 bits / 1 Gbps), sin "
        "latencia ni protocolo, y el ancho de banda real del enlace del banco lo reduce o lo amplía. "
        f"Medidores consultados en {PRICES_API} (serviceName, productName, meterName, armRegionName, "
        f"priceType Consumption): {meters}.")


WEEKS_PER_MONTH = 365.25 / 12 / 7


def kept(policy, now, taken):
    """True when the weekly LTR backup taken in week `taken` is still kept in week `now` under (W, M, Y): the
    monthly copy is the month's first weekly backup and the yearly copy the year's first."""
    weekly, monthly, yearly = policy
    age = now - taken
    first_of = lambda span: math.floor(taken / span) != math.floor((taken - 1) / span)
    return (age < weekly or (monthly and first_of(WEEKS_PER_MONTH) and age < monthly * WEEKS_PER_MONTH)
            or (yearly and first_of(WEEKS_PER_MONTH * 12) and age < yearly * 12 * WEEKS_PER_MONTH))


def ltr_copies_per_month(policy, months, base=(4, 12, 0)):
    """Average number of full LTR copies kept in each month. The database starts on the base rule: copies taken
    before week 0 exist only if the base rule kept them, and from week 0 on the new rule decides."""
    weeks = int(months * WEEKS_PER_MONTH)
    history = range(-int(max(policy[1], 12 * policy[2], base[1]) * WEEKS_PER_MONTH) - 1, 0)
    existing = [week for week in history if kept(base, 0, week)]
    per_month = [[] for _ in range(months)]
    for now in range(weeks):
        count = sum(kept(policy, now, week) for week in existing + list(range(now + 1)))
        per_month[min(months - 1, math.floor(now / WEEKS_PER_MONTH))].append(count)
    return [sum(values) / len(values) for values in per_month]


def slide_references(deck):
    slide = base_slide(deck, "Referencias: regulación",
                       "Textos oficiales revisados en noviembre de 2026. Clic en el número de la norma para abrir su texto.")
    rows = [[norm.replace(", ", "\n", 1), used, (examples, {"left": True, "size": 9}), (period, {"bold": True, "color": NAVY}), source]
            for norm, used, examples, period, source, _ in REGULATION_REFERENCES]
    grid = table(slide, Inches(0.35), Inches(1.15), Inches(12.63),
                 ["Norma", "Qué se usa", "Ejemplos que da la norma", "Plazo mínimo de conservación",
                  "Fuente oficial"], rows, [0.2, 0.2, 0.34, 0.14, 0.12], row_height=1.0, size=9.5)
    grid.rows[0].height = Inches(0.45)
    # The norm's number (first line of the cell) opens its official text; the title stays plain below it.
    for r, reference in enumerate(REGULATION_REFERENCES, start=1):
        for run in grid.cell(r, 0).text_frame.paragraphs[0].runs:
            run.hyperlink.address = reference[-1]


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: backup/reports/<date>-lineamientos/lineamientos_backup.pptx")
    parser.add_argument("--discount", type=float, default=DISCOUNT,
                        help="Bank discount on Azure list prices, as a fraction (default 0.30)")
    args = parser.parse_args()
    today = date.today()
    output = args.output or os.path.join(TOPIC_DIR, "reports", f"{today:%Y-%m-%d}-lineamientos",
                                         "lineamientos_backup.pptx")

    try:
        tb, _ = load_prices()
    except requests.RequestException as error:
        sys.exit(f"Could not read the Azure Retail Prices API: {error}")

    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    slide_guidelines(deck)
    slide_layers(deck)
    slide_azure(deck)
    net = {region: {key: value * (1 - args.discount) for key, value in values.items()} for region, values in tb.items()}
    slide_prices(deck, net, f"{today:%d-%m-%Y}", args.discount)
    slide_references(deck)
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    deck.save(output)
    print(output)


if __name__ == "__main__":
    main()
