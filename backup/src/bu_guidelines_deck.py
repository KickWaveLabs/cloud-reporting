#!/usr/bin/env python3
"""Build the backup guidelines deck: retention rules, price per TB against the vendors, and references.

Slides:
  1. Guidelines: which data the regulation covers, and PITR / LTR per service for regular and regulated data
  2. List price per TB-month in East US 2 and Central US against the vendors' quoted tariffs
  3. References: regulation, Microsoft Learn and pricing pages, each with its link

Prices come from the public Azure Retail Prices API (no sign-in, no subscription data). Retention limits and
regulatory articles are fixed in this file, checked against the official texts in November 2026; each one is
listed on slide 3. Vendor tariffs are the ones quoted in the backup guidelines review (unit to be confirmed).

Usage:
  pip install python-pptx requests
  python backup/src/bu_guidelines_deck.py
  python backup/src/bu_guidelines_deck.py --output lineamientos_backup.pptx

Output: backup/reports/<date>-lineamientos/lineamientos_backup.pptx (or --output).
"""

import argparse
import os
import sys
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

# (key, serviceName, productName, meterName) of every meter the deck prices.
METERS = [
    ("vault_lrs", "Backup", "Backup", "Standard LRS Data Stored"),
    ("vault_grs", "Backup", "Backup", "Standard GRS Data Stored"),
    ("files_vault_lrs", "Backup", "Backup", "Azure Files Vaulted LRS Data Stored"),
    ("files_vault_grs", "Backup", "Backup", "Azure Files Vaulted GRS Data Stored"),
    ("archive_lrs", "Backup", "Backup", "Archive LRS Data Stored"),
    ("archive_grs", "Backup", "Backup", "Archive GRS Data Stored"),
    ("vm_instance", "Backup", "Backup", "Azure VM Protected Instance"),
    ("pg_instance", "Backup", "Backup", "PostgreSQL Protected Instance"),
    ("cosmos_instance", "Backup", "Backup", "Cosmos DB Protected Instance"),
    ("sqldb_pitr_lrs", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage", "LRS Data Stored"),
    ("sqldb_pitr_geo", "SQL Database", "SQL Database Single/Elastic Pool PITR Backup Storage", "RA-GRS Data Stored"),
    ("sqldb_ltr_lrs", "SQL Database", "SQL Database - LTR Backup Storage", "Backup LRS Data Stored"),
    ("sqldb_ltr_geo", "SQL Database", "SQL Database - LTR Backup Storage", "Backup RA-GRS Data Stored"),
    ("sqlmi_pitr_lrs", "SQL Managed Instance", "SQL Managed Instance PITR Backup Storage", "LRS Data Stored"),
    ("sqlmi_pitr_geo", "SQL Managed Instance", "SQL Managed Instance PITR Backup Storage", "RA-GRS Data Stored"),
    ("sqlmi_ltr_lrs", "SQL Managed Instance", "SQL Managed Instance - LTR Backup Storage", "LTR Backup LRS Data Stored"),
    ("sqlmi_ltr_geo", "SQL Managed Instance", "SQL Managed Instance - LTR Backup Storage",
     "LTR Backup RA-GRS Data Stored"),
    ("pg_pitr", "Azure Database for PostgreSQL", "Azure Database for PostgreSQL Flexible Server Backup Storage",
     "Backup Storage LRS Data Stored"),
    ("mysql_pitr", "Azure Database for MySQL", "Azure Database for MySQL Flexible Server Backup Storage",
     "Backup Storage LRS Data Stored"),
    ("cosmos_30d", "Azure Cosmos DB", "Azure Cosmos DB - PITR", "Cont 30D Bckp Continuous Backup"),
    ("cosmos_restore", "Azure Cosmos DB", "Azure Cosmos DB - PITR", "Cont 30D Bckp Data Restore"),
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

SBS_2660 = "https://intranet2.sbs.gob.pe/dv_int_cn/1540/v5.0/Adjuntos/2660-2015.r.pdf"
SBS_504 = "https://intranet2.sbs.gob.pe/dv_int_cn/2046/v5.0/Adjuntos/504-2021.R.pdf"
SBS_814 = "https://intranet2.sbs.gob.pe/dv_int_cn/2483/v1.0/Adjuntos/814-2025.R.pdf"
BCRP_0011 = "https://www.bcrp.gob.pe/docs/Transparencia/Normas-Legales/Circulares/2026/circular-0011-2026-bcrp.pdf"
SPIJ = "https://spij.minjus.gob.pe/"
LEARN = "https://learn.microsoft.com/en-us/azure/"
PRICING = "https://azure.microsoft.com/en-us/pricing/"

REGULATION_REFERENCES = [
    ("Ley 26702, Ley General del Sistema Financiero, art. 183",
     "Libros y documentos: 10 años. Texto actualizado: buscar «Ley 26702» en el SPIJ.", SPIJ),
    ("Res. SBS 2660-2015, Gestión de Riesgos de LA/FT (texto consolidado)",
     "Arts. 49 y 54: registro de operaciones y copia trimestral. Arts. 55 y 58: KYC, contrapartes, inusuales.",
     SBS_2660),
    ("Res. SBS 504-2021, Seguridad de la Información y Ciberseguridad",
     "Art. 12, num. 3, lit. f) y h): copias de respaldo y pruebas. Arts. 23 y 25: nube y exterior.", SBS_504),
    ("Res. SBS 00814-2025, modifica Res. SBS 877-2020 (Continuidad)",
     "Productos pasivos: POR igual a cero y TOR de hasta 3 horas.", SBS_814),
    ("Circular BCRP 0011-2026, Reglamento del Sistema LBTR",
     "Art. 46.5: 2 horas. Anexo V, sección VI: respaldo diario, prueba anual, 10 años.", BCRP_0011),
]
MICROSOFT_REFERENCES = [
    ("Azure Backup para VMs: frecuencia y retención", LEARN + "backup/backup-support-matrix-iaas"),
    ("Azure Backup para Files: snapshot y bóveda", LEARN + "backup/azure-file-share-backup-overview"),
    ("Azure SQL Database: PITR 1–35 días, LTR 10 años", LEARN + "azure-sql/database/automated-backups-overview"),
    ("Azure SQL Managed Instance: PITR 1–35 días, LTR 10 años",
     LEARN + "azure-sql/managed-instance/automated-backups-overview"),
    ("PostgreSQL Flexible: PITR 7–35 días", LEARN + "postgresql/backup-restore/concepts-backup-restore"),
    ("PostgreSQL Flexible: LTR en bóveda hasta 10 años", LEARN + "backup/backup-azure-database-postgresql-flex-overview"),
    ("MySQL Flexible: PITR 1–35 días; LTR en pausa", LEARN + "mysql/flexible-server/concepts-backup-restore"),
    ("Cosmos DB: backup continuo 7, 30 o 35 días", LEARN + "cosmos-db/continuous-backup-restore-introduction"),
    ("Cosmos DB: LTR en bóveda hasta 10 años", LEARN + "backup/backup-azure-cosmos-db-overview"),
    ("SQL Database: RTO y RPO de cada opción de recuperación",
     LEARN + "azure-sql/database/business-continuity-high-availability-disaster-recover-hadr-overview"),
    ("Archive: rehidratación hasta 15 h (Standard)", LEARN + "storage/blobs/archive-rehydrate-overview"),
]
PRICING_REFERENCES = [
    ("Calculadora de precios de Azure", PRICING + "calculator/"),
    ("Azure Backup (bóveda, Archive, instancias protegidas)", PRICING + "details/backup/"),
    ("Azure SQL Database (PITR y LTR)", PRICING + "details/azure-sql-database/single/"),
    ("Azure SQL Managed Instance (PITR y LTR)", PRICING + "details/azure-sql-managed-instance/single/"),
    ("PostgreSQL Flexible Server (almacenamiento de backup)", PRICING + "details/postgresql/flexible-server/"),
    ("MySQL Flexible Server (almacenamiento de backup)", PRICING + "details/mysql/flexible-server/"),
    ("Cosmos DB (backup continuo)", PRICING + "details/cosmos-db/autoscale-provisioned/"),
    ("API pública de precios (Azure Retail Prices)", PRICES_API),
]


# ---------------------------------------------------------------------------
# Prices
# ---------------------------------------------------------------------------


def price_filter(region, service, product, meter):
    return (f"serviceName eq '{service}' and productName eq '{product}' and meterName eq '{meter}' "
            f"and armRegionName eq '{region}' and priceType eq 'Consumption'")


def fetch_price(region, service, product, meter):
    """Retail price of one meter in one region, or None when the API has no row for it."""
    url, params, items = PRICES_API, {"$filter": price_filter(region, service, product, meter)}, []
    while url:
        response = requests.get(url, params=params, timeout=60)
        response.raise_for_status()
        body = response.json()
        items += body.get("Items", [])
        url, params = body.get("NextPageLink"), None
    prices = {item["retailPrice"] for item in items if not item.get("tierMinimumUnits")}
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
    prices, missing = {}, []
    for region, _ in REGIONS:
        prices[region] = {}
        for key, service, product, meter in METERS:
            value = fetch_price(region, service, product, meter)
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
    """Text box; text is a string, a list of lines, or a list of lines made of (text, bold, color, link) runs."""
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = Inches(0.04)
    for index, line in enumerate(text if isinstance(text, list) else [text]):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        for part in line if isinstance(line, list) else [(line, bold, color, None)]:
            run = paragraph.add_run()
            run.text = part[0]
            run.font.name, run.font.size, run.font.bold, run.font.italic = FONT, Pt(size), part[1], italic
            run.font.color.rgb = part[2]
            if part[3]:
                run.hyperlink.address = part[3]
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
    """Rows are lists of cell values or of (text, options) pairs; a row of one value is a section row."""
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
        for c, value in enumerate(values):
            text, options = value if isinstance(value, tuple) else (value, {})
            style_cell(grid.cell(r, c), text, options.get("size", size), bold=options.get("bold", c == 0),
                       color=options.get("color", NAVY if c == 0 else GREY),
                       fill=options.get("fill", LIGHT if r % 2 else WHITE),
                       align=PP_ALIGN.LEFT if c == 0 or options.get("left") else PP_ALIGN.CENTER)
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


def slide_guidelines(deck):
    slide = base_slide(deck, "Lineamientos de backup y retención en Azure",
                       "Fuente: Ley 26702; Res. SBS 2660-2015, 504-2021 y 00814-2025; Circular BCRP 0011-2026; "
                       "Microsoft Learn (revisado nov-2026). Plazos legales a validar por Cumplimiento.")

    # Left: which data the regulation covers.
    left, width = Inches(0.35), Inches(3.95)
    heading(slide, left, Inches(1.08), width, "Datos regulados: 10 años como mínimo")
    regulated = [
        ("Libros y documentos de la empresa", "Ley 26702, art. 183",
         "Se conservan 10 años; si hay un juicio, mientras dure. Se permiten medios análogos (backup, archivo)."),
        ("Registro de operaciones", "Res. SBS 2660-2015, arts. 49 y 54",
         "Operaciones ≥ US$10 mil; múltiples ≥ US$50 mil al mes. Copia al cierre de cada trimestre, "
         "agrupada por 5 años, entregable en 48 h hábiles."),
        ("KYC, contrapartes e inusuales", "Res. SBS 2660-2015, arts. 55 y 58",
         "Debida diligencia de clientes, contrapartes y proveedores, y el análisis de operaciones inusuales."),
        ("Información del LBTR", "Circular BCRP 0011-2026, Anexo V",
         "Respaldo diario, prueba de restauración al menos una vez al año, conservación de 10 años."),
    ]
    top = Inches(1.45)
    for name, norm, detail in regulated:
        rounded(slide, left, top, width, Inches(1.08), LIGHT, radius=0.08)
        rail = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.07), Inches(1.08))
        rail.fill.solid()
        rail.fill.fore_color.rgb = RUST
        rail.line.fill.background()
        text_box(slide, left + Inches(0.15), top + Inches(0.02), width - Inches(0.2), Inches(0.25), name, size=10,
                 bold=True, color=NAVY)
        text_box(slide, left + Inches(0.15), top + Inches(0.23), width - Inches(0.2), Inches(0.22), norm, size=8.5,
                 bold=True, color=RUST)
        text_box(slide, left + Inches(0.15), top + Inches(0.44), width - Inches(0.2), Inches(0.62), detail, size=9)
        top += Inches(1.16)

    # Right: retention per service.
    left, width = Inches(4.5), Inches(8.48)
    heading(slide, left, Inches(1.08), width, "Retención por servicio (PROD)")
    base, regulated_fill = {"fill": WHITE}, {"fill": RGBColor(0xFB, 0xEE, 0xE9)}
    rows = [
        ["Azure Backup\n(VM y Files)", "Diario\n7 diarios", "4 semanales\n12 mensuales",
         ("Diario\n14 diarios", regulated_fill), ("4 sem · 60 mens · 10 anuales\nVM: Archive desde 6 meses",
                                                  regulated_fill)],
        ["Azure SQL Database", "PITR 7 días", "LTR S4 · M12", ("PITR 14 días", regulated_fill),
         ("LTR S4 · M60 · A10\ninmutable", regulated_fill)],
        ["Azure SQL Managed\nInstance", "PITR 7 días", "LTR S4 · M12", ("PITR 14 días", regulated_fill),
         ("LTR S4 · M60 · A10\nsin inmutable: lock y RBAC", regulated_fill)],
        ["PostgreSQL Flexible", "PITR 7 días", "Backup vault\nS4 · M12", ("PITR 14 días", regulated_fill),
         ("Backup vault S4 · M60 · A10\nhasta 1 TiB por servidor", regulated_fill)],
        ["MySQL Flexible", "PITR 7 días", "Sin LTR de Azure\n(en pausa)", ("PITR 14 días", regulated_fill),
         ("Export mensual a\nBlob inmutable, 10 años", regulated_fill)],
        ["Cosmos DB", "Continuo 7 días\n(sin costo)", "Backup vault\nS4 · M12", ("Continuo 30 días", regulated_fill),
         ("Backup vault S4 · M60 · A10", regulated_fill)],
    ]
    for row in rows:
        for index in (1, 2):
            row[index] = (row[index], base)
    grid = table(slide, left, Inches(1.45), width,
                 ["Servicio", "Base: corto plazo", "Base: largo plazo", "Regulado: corto plazo",
                  "Regulado: largo plazo"], rows, [0.19, 0.16, 0.18, 0.19, 0.28], row_height=0.56, size=9.5)
    style_cell(grid.cell(0, 3), "Regulado: corto plazo", 9.5, bold=True, color=WHITE, fill=RUST)
    style_cell(grid.cell(0, 4), "Regulado: largo plazo", 9.5, bold=True, color=WHITE, fill=RUST)

    text_box(slide, left, Inches(5.47), width, Inches(0.62), [
        [("Controles para datos regulados: ", True, RUST, None),
         ("copia geo-redundante a Central US (región par, a más de 500 km); bóveda inmutable con borrado temporal y "
          "autorización multiusuario; prueba de restauración documentada al menos una vez al año. ", False, GREY, None),
         ("No productivo: ", True, NAVY, None), ("PITR 7 días, sin LTR.", False, GREY, None)]], size=9.5)
    note_box(slide, Inches(0.35), Inches(6.18), Inches(12.63), Inches(0.82), [
        [("PITR", True, NAVY, None),
         (" recupera errores recientes: borrado, corrupción, ransomware (RPO de 5 a 10 minutos).   ", False, GREY, None),
         ("LTR", True, NAVY, None), (" cumple la conservación y los errores detectados tarde.", False, GREY, None)],
        [("RPO cero de productos pasivos (SBS 00814-2025) y recuperación del LBTR en 2 horas (BCRP 0011-2026): ",
          True, RUST, None),
         ("se cumplen con replicación y failover, no con backup.", False, GREY, None)]], size=10)
    slide.notes_slide.notes_text_frame.text = (
        "S = semanales, M = mensuales, A = anuales (cantidad de puntos que se guardan). "
        "M60 cubre la copia trimestral agrupada por 5 años del art. 54 de la Res. SBS 2660-2015; A10 cubre los 10 años "
        "de la Ley 26702 art. 183, la Res. SBS 2660-2015 arts. 54-55 y la Circular BCRP 0011-2026. "
        "La obligación es sobre el dato: si el dato regulado ya se conserva 10 años en otro repositorio inmutable, el "
        "backup de ese sistema puede quedarse en el lineamiento base. "
        "PITR regulado de 14 días: ninguna norma fija un plazo de PITR, y el PITR cuesta unas 4 veces el LTR por TB "
        "(slide 2); los S4 del LTR cubren hasta 4 semanas. Cosmos DB no tiene tier de 14 días (solo 7, 30 o 35). "
        "Azure Backup cubre VMs y Azure Files (Recovery Services vault). El LTR de PostgreSQL y Cosmos DB va a un "
        "Backup vault de Azure Backup, aparte de la base; su PITR nativo queda en el servicio. "
        "MySQL Flexible: el LTR con Azure Backup está en pausa por Microsoft; el export (mysqldump) a Blob con "
        "inmutabilidad lo reemplaza. PostgreSQL LTR: semanal, mensual o anual, máximo 1 TiB, restaura como archivos. "
        "SQL Managed Instance no admite LTR inmutable. Archive: puntos mensuales y anuales con más de 6 meses; "
        "rehidratar toma horas, menos que las 48 h hábiles de la SBS.")


def slide_prices(deck, tb, retrieved):
    east, central = tb["eastus2"], tb["centralus"]
    regions = (east, central)
    small = {"left": True, "size": 8}
    transfer = f"~{GB_PER_TB * 8 / 3600:.1f} h/TB a 1 Gbps".replace(".", ",")

    def pair(local, geo):
        return [amount(region[local]) for region in regions], [amount(region[geo]) if geo else "–" for region in regions]

    def priced(label, local, geo, rpo, rto, extra, best=False):
        locals_, geos = pair(local, geo)
        options = {"color": GREEN, "bold": True} if best else {}
        return [label, (locals_[0], options), (geos[0], options), (locals_[1], options), (geos[1], options),
                (rpo, {"size": 8}), (rto, small), (extra, small)]

    def vendor(name, value, tier, rto):
        options = {"color": RUST, "bold": True}
        return [(f"{name} · {tier}", {"color": RUST}), *[(amount(value), options)] * 4,
                ("Según su política", {"size": 8}), (rto, small),
                ("Cotizada; confirmar TB protegido o almacenado", small)]

    sql_label = "SQL Database / Managed Instance"
    sql_equal = all(region[f"sqldb_{kind}"] == region[f"sqlmi_{kind}"]
                    for region in regions for kind in ("pitr_lrs", "pitr_geo", "ltr_lrs", "ltr_geo"))
    if not sql_equal:
        raise SystemExit("SQL Database and SQL Managed Instance backup prices differ; split their rows.")
    rows = [
        ["Corto plazo: recuperación operativa"],
        priced("Azure Backup · VM (bóveda)", "vault_lrs", "vault_grs", "24 h (Enhanced: 4 h)",
               "Sin cifra publicada; snapshot local más rápido", f"+ {usd(east['vm_instance'])} por TB protegido"),
        priced("Azure Backup · Files (bóveda)", "files_vault_lrs", "files_vault_grs", "24 h (diario)",
               "Archivos sueltos al instante desde snapshot", "+ instancia protegida (ver calculadora)"),
        priced(f"{sql_label} · PITR", "sqldb_pitr_lrs", "sqldb_pitr_geo", "5 a 10 min",
               "Normalmente < 12 h", "Gratis hasta el tamaño máximo de la base"),
        priced("PostgreSQL / MySQL · PITR", "pg_pitr", None, "≤ 5 min", "Minutos a pocas horas",
               "Gratis hasta el storage aprovisionado; geo x2"),
        priced("Cosmos DB · continuo 30 días", "cosmos_30d", None, "~100 s", "Sin cifra publicada",
               f"Por región; restaurar {usd(east['cosmos_restore'])} por TB"),
        *[vendor(name, active, "Active", f"Restaurar + traer: {transfer}") for name, active, _ in VENDORS],
        ["Largo plazo: retención"],
        priced("Azure Backup · Archive (VM)", "archive_lrs", "archive_grs", "1 mes",
               "+ rehidratar: hasta 15 h", "Puntos > 6 meses; mínimo 180 días", best=True),
        priced(f"{sql_label} · LTR", "sqldb_ltr_lrs", "sqldb_ltr_geo", "1 semana", "Normalmente < 12 h",
               "Solo backups completos"),
        priced("PostgreSQL / Cosmos · Backup vault", "vault_lrs", "vault_grs", "1 sem. / 1 día",
               "Sin cifra; PostgreSQL vuelve como archivos",
               f"+ PG {usd(east['pg_instance'])} · Cosmos {usd(east['cosmos_instance'])} por TB protegido"),
        *[vendor(name, archive, "Archive", f"Su Archive (a confirmar) + {transfer}") for name, _, archive in VENDORS],
    ]
    if any(east[key] != central[key] for key in ("vm_instance", "pg_instance", "cosmos_instance", "cosmos_restore")):
        print("Note: protected instance or restore fees differ between regions; the slide shows East US 2.")

    slide = base_slide(deck, "Costo por TB al mes: Azure vs proveedores",
                       f"Fuente: Azure Retail Prices API, precio de lista consultado el {retrieved}, sin descuento "
                       "del banco; 1 TB = 1.024 GB. Proveedores: tarifas cotizadas.")
    text_box(slide, Inches(0.35), Inches(1.0), Inches(12.63), Inches(0.28), [[
        ("Unidad: US$ por TB almacenado al mes", True, NAVY, None),
        (" (1 TB = 1.024 GB), precio de lista de la API de Azure y tarifas cotizadas. RPO y RTO: Microsoft Learn; "
         "en proveedores, estimado.", False, GREY, None)]], size=9)
    grid = table(slide, Inches(0.35), Inches(1.3), Inches(12.63),
                 ["Opción", "East US 2\nLRS", "East US 2\nGRS", "Central US\nLRS", "Central US\nGRS", "RPO",
                  "RTO (restaurar)", "Qué más se cobra"], rows,
                 [0.215, 0.06, 0.06, 0.06, 0.06, 0.095, 0.2, 0.25], row_height=0.3, size=8.5)
    grid.rows[0].height = Inches(0.42)
    for r, values in enumerate(rows, start=1):
        if len(values) == 1:
            grid.rows[r].height = Inches(0.26)

    vault_protected = east["vault_grs"] + east["vm_instance"]
    actives = [active for _, active, _ in VENDORS]
    archives = [archive for _, _, archive in VENDORS]
    takeaways = [
        ("Largo plazo", f"Archive GRS de Azure: {usd(east['archive_grs'])} por TB, más barato que el Archive de los "
                        f"proveedores ({' y '.join(amount(value) for value in archives)})."),
        ("PITR vs LTR", f"El PITR geo cuesta {east['sqldb_pitr_geo'] / east['sqldb_ltr_geo']:.0f} veces el LTR por TB "
                        f"({amount(east['sqldb_pitr_geo'])} vs {amount(east['sqldb_ltr_geo'])}): PITR corto y LTR "
                        "para retener."),
        ("Corto plazo", f"Bóveda GRS + instancia VM = {usd(vault_protected)} por TB protegido vs proveedores "
                        f"{amount(min(actives))}–{amount(max(actives))}; su RTO suma la transferencia ({transfer})."),
    ]
    for index, (title, text) in enumerate(takeaways):
        left = Inches(0.35 + index * 4.25)
        rounded(slide, left, Inches(6.02), Inches(4.13), Inches(0.68), CARD)
        text_box(slide, left + Inches(0.12), Inches(6.02), Inches(3.89), Inches(0.68),
                 [[(title + ": ", True, ORANGE, None), (text, False, WHITE, None)]], size=9, anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, Inches(0.35), Inches(6.76), Inches(12.6), Inches(0.3), [[
        ("Calcula cada caso en la ", False, GREY, None), ("calculadora de Azure", True, NAVY, PRICING + "calculator/"),
        ("; precios por servicio en ", False, GREY, None), ("Azure Backup", True, NAVY, PRICING + "details/backup/"),
        (" · ", False, GREY, None), ("SQL Database", True, NAVY, PRICING + "details/azure-sql-database/single/"),
        (" · ", False, GREY, None), ("SQL MI", True, NAVY, PRICING + "details/azure-sql-managed-instance/single/"),
        (" · ", False, GREY, None), ("PostgreSQL", True, NAVY, PRICING + "details/postgresql/flexible-server/"),
        (" · ", False, GREY, None), ("MySQL", True, NAVY, PRICING + "details/mysql/flexible-server/"),
        (" · ", False, GREY, None), ("Cosmos DB", True, NAVY, PRICING + "details/cosmos-db/autoscale-provisioned/"),
    ]], size=9)
    meters = "; ".join(f"{meter} ({product})" for _, _, product, meter in METERS)
    slide.notes_slide.notes_text_frame.text = (
        "Cada precio por GB-mes se multiplica por 1.024. Las instancias protegidas se cobran por bloque de datos de "
        "origen: VM US$10 por 500 GB, PostgreSQL US$7,5 por 250 GB, Cosmos DB US$25 por 500 GB. "
        "El PITR de SQL, PostgreSQL y MySQL solo cobra lo que excede el almacenamiento gratis. "
        "Azure Backup guarda incrementales: los TB almacenados son menos que los TB protegidos. "
        "RPO y RTO: SQL Database y Managed Instance (log cada 5 a 10 min; restaurar normalmente toma menos de 12 h, "
        "una base grande o muy activa puede tomar varias horas); PostgreSQL (RPO hasta 5 min; recuperar toma de "
        "minutos a pocas horas); MySQL (log cada 5 min); Cosmos DB (cambios respaldados en 100 s); Archive "
        "(rehidratar hasta 15 h en prioridad Standard, menos de 1 h en High para objetos menores a 10 GB). "
        f"Proveedores: el dato vuelve por la red; {transfer} es el mínimo teórico (1.024 GB x 8 bits / 1 Gbps), sin "
        "latencia ni protocolo, y el ancho de banda real del enlace del banco lo reduce o lo amplía. "
        f"Medidores consultados en {PRICES_API} (serviceName, productName, meterName, armRegionName, "
        f"priceType Consumption): {meters}.")


def slide_references(deck):
    slide = base_slide(deck, "Referencias", "Todos los enlaces son públicos y oficiales; revisados en noviembre de 2026.")
    columns = [
        ("Regulación (Perú)", Inches(0.35), Inches(4.6)),
        ("Microsoft Learn: límites de cada servicio", Inches(5.1), Inches(4.0)),
        ("Precios", Inches(9.25), Inches(3.73)),
    ]
    for title, left, width in columns:
        rounded(slide, left, Inches(1.12), width, Inches(0.42), CARD, radius=0.2)
        text_box(slide, left + Inches(0.1), Inches(1.12), width - Inches(0.2), Inches(0.42), title, size=12, bold=True,
                 color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    left, width = columns[0][1], columns[0][2]
    top = Inches(1.66)
    for name, detail, url in REGULATION_REFERENCES:
        text_box(slide, left, top, width, Inches(0.98), [
            [(name, True, NAVY, url)], [(detail, False, GREY, None)], [(url, False, RUST, url)]], size=9)
        top += Inches(1.0)

    for column, references in ((columns[1], MICROSOFT_REFERENCES), (columns[2], PRICING_REFERENCES)):
        left, width = column[1], column[2]
        top = Inches(1.66)
        step = Inches(5.0 / len(references))
        for name, url in references:
            text_box(slide, left, top, width, step, [[(name, True, NAVY, url)], [(url, False, RUST, url)]], size=8.5)
            top += step


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: backup/reports/<date>-lineamientos/lineamientos_backup.pptx")
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
    slide_prices(deck, tb, f"{today:%d-%m-%Y}")
    slide_references(deck)
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    deck.save(output)
    print(output)


if __name__ == "__main__":
    main()
