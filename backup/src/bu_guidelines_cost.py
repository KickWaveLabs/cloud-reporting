#!/usr/bin/env python3
"""Estimate what the backup guidelines cost per database, today and under the proposed rule.

A working file for the architect, not for the deck. Reads the database inventory (Resource Graph and ARM,
written by bu_database_inventory.py) and, when present, the used size of each SQL database
(bu_database_used_storage.py). Prices come from the public Azure Retail Prices API through
bu_guidelines_deck.py. No other Azure call is made.

Sheets:
  Resumen          totals per service and environment, today vs proposed (formulas)
  Bases            one row per database: size, PITR and LTR today, scope (drop-down), proposed rule and cost
  Hoy              how VMs, Files and databases are configured today, by retention setting
  Supuestos        prices per GB and per protected instance, the LTR size factor, scope rules
  Como_se_calcula  the method in plain Spanish

The estimate is list price before the bank's discount. LTR storage = size x copies kept x factor; the factor
(0.43) is the ratio between billed LTR GB and used size x copies measured on the October 2026 bill. Open the
file in Excel and save it once so the formulas calculate.

Usage:
  python backup/src/bu_guidelines_cost.py
Output: backup/reports/<date>-lineamientos/costo_lineamiento_bases.xlsx (or --output).
"""

import argparse
import csv
import math
import os
import re
from collections import Counter
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import bu_guidelines_deck as deck

RESULTS = os.path.join(deck.TOPIC_DIR, "results")
LTR_FACTOR = 0.43
SCOPES = ["Base", "Registro de operaciones", "LBTR", "Libros o KYC que purgan"]
SERVICE_LABELS = {"AZURE_SQL_DB": "Azure SQL Database", "POSTGRES_FLEX": "PostgreSQL Flexible",
                  "MYSQL_FLEX": "MySQL Flexible", "COSMOS_DB": "Cosmos DB"}
GEO = {"Geo", "GeoZone", "Geo/Region"}
HEADER = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3355")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def read_csv(*parts):
    path = os.path.join(RESULTS, *parts)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def iso_days(value):
    """Days in an ISO 8601 duration such as P4W, P1M, P12M or P5Y; 0 when empty."""
    match = re.fullmatch(r"P(\d+)([DWMY])", (value or "").strip())
    if not match:
        return 0
    return int(match.group(1)) * {"D": 1, "W": 7, "M": 30.4, "Y": 365}[match.group(2)]


def copies(weekly, monthly, yearly):
    """LTR copies kept by a SQL policy: weeks, months and years it retains."""
    return (round(iso_days(weekly) / 7), round(iso_days(monthly) / 30.4), round(iso_days(yearly) / 365))


def environment(row):
    value = (row.get("environment") or "").upper()
    return "PROD" if value.startswith("P") else value or "SIN TAG"


def number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def database_rows():
    used = {row["id"].lower(): number(row.get("used_gb")) for row in read_csv("backup-database",
                                                                              "database_used_storage.csv")
            if row.get("id")}
    rows = []
    for row in read_csv("backup-database", "database_inventory.csv"):
        service = row["service_type"]
        if service not in SERVICE_LABELS or row.get("backup_method") == "NOT_APPLICABLE_REPLICA":
            continue
        size, source = used.get(row["id"].lower()), "usado (Azure Monitor)"
        if size is None and number(row.get("max_size_gb")):
            size, source = number(row["max_size_gb"]), "máximo de la base (sobreestima)"
        if size is None and number(row.get("storage_gb")):
            size, source = number(row["storage_gb"]), "aprovisionado (sobreestima)"
        if size is None:
            source = "sin dato: completar"
        weeks, months, years = copies(row["ltr_weekly"], row["ltr_monthly"], row["ltr_yearly"])
        rows.append({
            "service": SERVICE_LABELS[service], "service_type": service, "name": row["name"],
            "server": row["server_name"], "subscription": row["subscription"], "cod_app": row["cod_app"],
            "environment": environment(row), "location": row["location"],
            "redundancy": row["native_backup_redundancy"], "size": size, "size_source": source,
            "pitr": row["native_retention_days"], "ltr_text": " · ".join(
                f"{label}{value}" for label, value in (("S", weeks), ("M", months), ("A", years)) if value) or "Sin LTR",
            "copies": weeks + months + years,
        })
    return sorted(rows, key=lambda r: (r["environment"] != "PROD", r["service"], r["cod_app"], r["name"]))


def price_for(prices, location, kind, redundancy):
    """US$ per GB-month of a backup meter in the row's region and redundancy; East US 2 when the region was not
    priced."""
    region = location if location in prices else "eastus2"
    tier = "geo" if redundancy in GEO else "zrs" if "Zone" in (redundancy or "") else "lrs"
    keys = {"sqldb_ltr": {"geo": "sqldb_ltr_geo", "zrs": "sqldb_ltr_zrs", "lrs": "sqldb_ltr_lrs"},
            "vault": {"geo": "vault_grs", "zrs": "vault_zrs", "lrs": "vault_lrs"}}[kind]
    return prices[region][keys[tier]], region


def header(sheet, values, widths):
    sheet.append(values)
    for column, width in enumerate(widths, start=1):
        cell = sheet.cell(row=sheet.max_row, column=column)
        cell.font, cell.fill = HEADER, HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.column_dimensions[get_column_letter(column)].width = width
    sheet.freeze_panes = f"C{sheet.max_row + 1}"


def build_assumptions(book, raw, retrieved):
    sheet = book.create_sheet("Supuestos")
    header(sheet, ["Supuesto", "Valor", "Unidad", "Fuente"], [52, 12, 22, 70])
    rows = [("Factor de tamaño LTR", LTR_FACTOR, "GB facturado / (usado x copias)",
             "Factura real oct-2026 de SQL LTR: 91.492 GB-mes facturados / suma(usado x copias) = 0,43")]
    for region, label in deck.REGIONS:
        for key, text in (("vm_instance", "Instancia protegida VM, por 500 GB"),
                          ("pg_instance", "Instancia protegida PostgreSQL, por 250 GB"),
                          ("cosmos_instance", "Instancia protegida Cosmos DB, por 500 GB")):
            rows.append((f"{text} ({label})", raw[region][key], "US$ al mes", deck.PRICES_API))
    rows.append(("Precios consultados", retrieved, "", "Precio de lista, sin descuento del banco"))
    for row in rows:
        sheet.append(list(row))
    sheet.append([])
    sheet.append(["Regla propuesta por ámbito (solo PROD)", "S", "M", "A"])
    for scope, rule in (("Base", (4, 12, 0)), ("Registro de operaciones", (4, 60, 0)), ("LBTR", (4, 12, 10)),
                        ("Libros o KYC que purgan", (4, 12, 10))):
        sheet.append([scope, *rule])
    sheet.append(["No productivo", 0, 0, 0])
    return "Supuestos!$B$2"


def build_databases(book, rows, prices, raw, factor):
    sheet = book.create_sheet("Bases")
    header(sheet, ["Servicio", "Base de datos", "Servidor", "Suscripción", "codApp", "Ambiente", "Región",
                   "Redundancia", "Tamaño GB", "Origen del tamaño", "PITR hoy (días)", "LTR hoy",
                   "Copias LTR hoy", "Costo LTR hoy US$/mes", "Ámbito", "PITR propuesto", "S", "M", "A",
                   "Copias propuestas", "Costo propuesto US$/mes", "Diferencia US$/mes", "US$ por GB-mes",
                   "Instancia US$/bloque", "Bloque GB", "Nota"],
                 [20, 28, 22, 22, 10, 10, 13, 11, 10, 26, 10, 14, 10, 12, 22, 10, 6, 6, 6, 10, 12, 12, 10, 10, 8, 50])
    scope_list = DataValidation(type="list", formula1='"' + ",".join(SCOPES) + '"', allow_blank=False)
    sheet.add_data_validation(scope_list)
    for index, row in enumerate(rows, start=2):
        service = row["service_type"]
        kind = "sqldb_ltr" if service == "AZURE_SQL_DB" else "vault"
        unit_price, region = price_for(prices, row["location"], kind, row["redundancy"])
        instance = block = None
        if service == "POSTGRES_FLEX":
            instance, block = raw.get(region, raw["eastus2"])["pg_instance"], 250
        elif service == "COSMOS_DB":
            instance, block = raw.get(region, raw["eastus2"])["cosmos_instance"], 500
        note = []
        if region != row["location"]:
            note.append(f"Región {row['location']} sin precio: se usa East US 2.")
        if service == "MYSQL_FLEX":
            note.append("Sin LTR de Azure (en pausa): si un ámbito lo exige, export a Blob inmutable.")
        if service == "COSMOS_DB" and row["size"] is None:
            note.append("Completar el tamaño de datos de la cuenta (métrica DataUsage).")
        size = f"I{index}"
        today = f"=IF(OR({size}=\"\",M{index}=0),0,{size}*M{index}*{factor}*W{index})"
        prod = f'F{index}="PROD"'
        proposed_copies = (f"=IF(B{index}=\"\",0,IF(A{index}=\"MySQL Flexible\",0,Q{index}+R{index}+S{index}))")
        storage = f"{size}*T{index}*{factor}*W{index}"
        instance_fee = f"IF(T{index}=0,0,CEILING({size}/Y{index},1)*X{index})" if instance else "0"
        sheet.append([
            row["service"], row["name"], row["server"], row["subscription"], row["cod_app"], row["environment"],
            row["location"], row["redundancy"], row["size"], row["size_source"], number(row["pitr"]),
            row["ltr_text"], row["copies"], today if service == "AZURE_SQL_DB" else 0, "Base", 7,
            f"=IF({prod},4,0)", f'=IF({prod},IF(O{index}="Registro de operaciones",60,12),0)',
            f'=IF(AND({prod},OR(O{index}="LBTR",O{index}="Libros o KYC que purgan")),10,0)', proposed_copies,
            f"=IF(OR({size}=\"\",T{index}=0),0,{storage}+{instance_fee})", f"=U{index}-N{index}",
            unit_price, instance, block, " ".join(note),
        ])
        sheet.cell(row=index, column=15).fill = INPUT_FILL
        sheet.cell(row=index, column=9).fill = INPUT_FILL
        for column in (14, 21, 22, 23):
            sheet.cell(row=index, column=column).number_format = "#,##0.00"
    if rows:
        scope_list.add(f"O2:O{len(rows) + 1}")
    sheet.auto_filter.ref = sheet.dimensions
    return len(rows) + 1


def build_summary(book, last_row, rows):
    sheet = book.create_sheet("Resumen", 0)
    sheet["A1"] = "Costo de LTR de bases de datos: hoy vs lineamiento propuesto (US$ al mes, precio de lista)"
    sheet["A1"].font = Font(bold=True, size=13)
    sheet["A2"] = ("Cambia el Ámbito (columna O de Bases) o el tamaño (columna I) y guarda en Excel: los totales "
                   "se recalculan. PITR de 7 días no se suma: entra en el espacio de backup gratis de cada servicio. "
                   "Con 'Tamaño no medido' > 0 el costo sale inflado (tamaño máximo o aprovisionado): corre "
                   "bu_database_used_storage.py y vuelve a generar este archivo.")
    sheet.append([])
    header(sheet, ["Servicio", "Ambiente", "Bases", "Con LTR hoy", "Tamaño no medido", "Costo hoy",
                   "Costo propuesto", "Diferencia"], [24, 12, 10, 12, 16, 14, 16, 14])
    groups = sorted({(r["service"], r["environment"]) for r in rows}, key=lambda g: (g[1] != "PROD", g[0], g[1]))
    for service, env in groups:
        crit = f'Bases!$A$2:$A${last_row},"{service}",Bases!$F$2:$F${last_row},"{env}"'
        row = sheet.max_row + 1
        sheet.append([service, env, f"=COUNTIFS({crit})", f'=COUNTIFS({crit},Bases!$M$2:$M${last_row},">0")',
                      f'=COUNTIFS({crit},Bases!$J$2:$J${last_row},"<>usado (Azure Monitor)")',
                      f"=SUMIFS(Bases!$N$2:$N${last_row},{crit})", f"=SUMIFS(Bases!$U$2:$U${last_row},{crit})",
                      f"=G{row}-F{row}"])
    total = sheet.max_row + 1
    sheet.append(["Total", "", f"=SUM(C5:C{total - 1})", f"=SUM(D5:D{total - 1})", f"=SUM(E5:E{total - 1})",
                  f"=SUM(F5:F{total - 1})", f"=SUM(G5:G{total - 1})", f"=G{total}-F{total}"])
    for cell in sheet[total]:
        cell.font = Font(bold=True)
    for row in sheet.iter_rows(min_row=5, max_row=total, min_col=6, max_col=8):
        for cell in row:
            cell.number_format = "#,##0"


def build_today(book):
    sheet = book.create_sheet("Hoy")
    header(sheet, ["Servicio", "Ambiente", "Configuración de hoy", "Cantidad"], [24, 12, 60, 10])
    counts = Counter()
    for row in read_csv("backup-database", "database_inventory.csv"):
        if row["service_type"] not in SERVICE_LABELS:
            continue
        weeks, months, years = copies(row["ltr_weekly"], row["ltr_monthly"], row["ltr_yearly"])
        ltr = " · ".join(f"{label}{value}" for label, value in (("S", weeks), ("M", months), ("A", years)) if value)
        counts[(SERVICE_LABELS[row["service_type"]], environment(row),
                f"PITR {row['native_retention_days'] or '?'} días; {('LTR ' + ltr) if ltr else 'sin LTR'}; "
                f"backup {row['native_backup_redundancy'] or '?'}")] += 1
    for row in read_csv("backup-compute", "compute_vm_inventory.csv"):
        if row.get("backup_enabled") != "YES":
            setting = "Sin backup de Azure Backup (incluye nodos de AKS, Databricks y otras plataformas)"
        else:
            setting = (f"{row['daily_retention'] or 0} diarios · {row['weekly_retention'] or 0} semanales · "
                       f"{row['monthly_retention'] or 0} mensuales · {row['yearly_retention'] or 0} anuales")
        counts[("Azure Backup VM", environment(row), setting)] += 1
    for row in read_csv("backup-storage", "storage_backup_policies.csv"):
        if row["policy_workload"] == "AzureFileShare":
            counts[("Azure Backup Files", "Todos", f"{row['daily_retention'] or 0} diarios · "
                    f"{row['weekly_retention'] or 0} semanales · {row['monthly_retention'] or 0} mensuales · "
                    f"{row['yearly_retention'] or 0} anuales (por política)")] += int(row["protected_items_count"] or 0)
    for (service, env, setting), count in sorted(counts.items(), key=lambda item: (item[0][0], item[0][1] != "PROD",
                                                                                   -item[1])):
        sheet.append([service, env, setting, count])


def build_method(book):
    sheet = book.create_sheet("Como_se_calcula")
    sheet.column_dimensions["A"].width = 120
    for line in [
        "Cómo se calcula cada cifra",
        "",
        "1. Tamaño: espacio usado de la base (Azure Monitor) si existe database_used_storage.csv; si no, el tamaño "
        "máximo (SQL) o el aprovisionado (PostgreSQL, MySQL). Esos dos sobreestiman el costo. Cosmos DB no trae "
        "tamaño: hay que completarlo.",
        "2. Copias LTR: semanas, meses y años que guarda la política. S4 · M12 = 4 + 12 = 16 copias.",
        "3. GB de LTR al mes = tamaño x copias x factor 0,43. El factor sale de la factura real de octubre 2026: lo "
        "facturado por LTR fue el 43% de (tamaño usado x copias), porque los backups van comprimidos.",
        "4. Costo = GB de LTR x precio por GB de la región y redundancia (Supuestos y columna W).",
        "5. PostgreSQL y Cosmos DB guardan el LTR en un Backup vault: se suma la instancia protegida "
        "(US$7,5 por cada 250 GB en PostgreSQL, US$25 por cada 500 GB en Cosmos DB).",
        "6. PITR de 7 días: no se suma. SQL vCore, PostgreSQL y MySQL dan gratis backup hasta el tamaño de la base; "
        "en SQL DTU el PITR va incluido; Cosmos continuo 7 días no cobra.",
        "7. Ámbito: elige en la columna O. Base = S4 · M12; Registro de operaciones = S4 · M60; LBTR y Libros o KYC "
        "que purgan = S4 · M12 · A10. No productivo = sin LTR.",
        "8. Todo es precio de lista; el costo real lleva el descuento del banco (para SQL LTR, ~0,0309 US$/GB-mes real "
        "frente a 0,05 de lista RA-GRS).",
    ]:
        sheet.append([line])
    sheet["A1"].font = Font(bold=True, size=13)
    for row in sheet.iter_rows(min_row=3):
        row[0].alignment = Alignment(wrap_text=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: backup/reports/<date>-lineamientos/costo_lineamiento_bases.xlsx")
    args = parser.parse_args()
    today = date.today()
    output = args.output or os.path.join(deck.TOPIC_DIR, "reports", f"{today:%Y-%m-%d}-lineamientos",
                                         "costo_lineamiento_bases.xlsx")
    rows = database_rows()
    if not rows:
        raise SystemExit("No databases found. Run bu_database_inventory.py first.")
    prices_tb, raw = deck.load_prices()
    prices = {region: {key: value / deck.GB_PER_TB for key, value in values.items()}
              for region, values in prices_tb.items()}

    book = Workbook()
    book.remove(book.active)
    factor = build_assumptions(book, raw, f"{today:%d-%m-%Y}")
    last_row = build_databases(book, rows, prices, raw, factor)
    build_summary(book, last_row, rows)
    build_today(book)
    build_method(book)
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    book.save(output)
    print(output)


if __name__ == "__main__":
    main()
