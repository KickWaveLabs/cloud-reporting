#!/usr/bin/env python3
"""Build the backup insights workbook and a metrics JSON from the bu_*_inventory.py outputs.

Reads the CSVs written by the four inventory scripts (and, if present, the cost CSVs and
cached vault cost queries), then writes:
  backup_insights.xlsx  Summary, opportunities, risks, long-term restore risks, minimum
                        policy, assumptions and the detail sheets the formulas count over.
  metrics.json          Figures used by the slide decks.

Every count in the summary sheets is a COUNTIFS/SUMIFS formula over the detail sheets, and
every saving is a formula over the assumptions sheet, so the workbook recalculates when an
assumption is edited.
"""

import argparse
import csv
import glob
import json
import os
from collections import Counter, defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
TITLE_FONT = Font(name=FONT, bold=True, size=14)
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
USD = '"$"#,##0;("$"#,##0);"-"'
WRAP = Alignment(wrap_text=True, vertical="top")


def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


# ---------------------------------------------------------------------------
# Detail sheets
# ---------------------------------------------------------------------------

DETAIL_SHEETS = {
    "D_VMs": ("compute", "compute_vm_inventory.csv", [
        "subscription", "resource_group", "name", "environment", "cod_app", "operation_normalized",
        "infrastructure_class", "backup_scope", "vm_size", "os_type", "power_state", "total_disk_gb",
        "backup_method", "backup_enabled", "vault_name", "vault_storage_redundancy", "policy_name",
        "policy_type", "schedule_frequency", "retention_max_days", "recovery_point_count",
        "last_backup_age_hours", "archive_enabled", "finding", "severity", "id"]),
    "D_Huerfanos": ("compute", "compute_orphan_backup_items.csv", [
        "orphan_reason", "vault_name", "vault_subscription_id", "vault_storage_redundancy",
        "protected_item_name", "workload_type", "protection_state", "last_recovery_point",
        "days_since_last_recovery_point", "recovery_point_count", "configured_max_retention_days",
        "policy_name", "source_resource_id"]),
    "D_K8s": ("compute", "compute_k8s_inventory.csv", [
        "service_type", "subscription", "resource_group", "name", "environment", "kubernetes_version",
        "sku_tier", "power_state", "node_count", "gitops_flux", "backup_extension_installed",
        "backup_enabled", "backup_method", "finding", "severity", "id"]),
    "D_Bovedas": ("compute", "compute_backup_vaults.csv", [
        "vault_kind", "vault_name", "subscription_id", "resource_group", "location", "storage_redundancy",
        "cross_region_restore", "soft_delete", "immutability", "item_count", "active_item_count", "workloads",
        "vault_id"]),
    "D_BD": ("database", "database_inventory.csv", [
        "service_type", "subscription", "resource_group", "name", "environment", "cod_app", "engine_version",
        "sku_name", "sku_tier", "max_size_gb", "storage_gb", "native_retention_days", "native_backup_redundancy",
        "native_geo_backup", "ltr_configured", "ltr_max_days", "backup_method", "finding", "severity", "id"]),
    "D_Storage": ("storage", "storage_inventory.csv", [
        "service_type", "subscription", "resource_group", "name", "environment", "cod_app", "infrastructure_class",
        "backup_scope", "kind", "sku", "redundancy", "access_tier", "capacity_gb", "used_capacity_gb", "lifecycle_policy",
        "disk_tier", "disk_state", "unattached_days",
        "snapshot_count", "snapshot_total_gb", "blob_soft_delete_days", "versioning", "pitr_days",
        "share_soft_delete_days", "data_role_hint", "backup_method", "backup_enabled", "finding", "severity",
        "all_findings", "id"]),
    "D_Snapshots": ("storage", "storage_snapshots_by_source.csv", [
        "source_type", "source_exists", "source_disk_state", "snapshot_count", "snapshot_full_count",
        "snapshot_incremental_count", "snapshot_total_gb", "oldest_snapshot_age_days", "newest_snapshot",
        "snapshot_resource_groups", "source_resource_id"]),
    "D_Nativo": ("native", "native_protection_inventory.csv", [
        "service_type", "subscription", "resource_group", "name", "environment", "resource_kind", "sku", "tier",
        "zone_redundant", "multi_region", "native_backup_capability", "native_backup_configured", "soft_delete_days",
        "purge_protection", "has_delete_lock", "iac_managed", "backup_method", "infrastructure_class", "finding",
        "severity", "id"]),
    "D_Planes": ("native", "native_app_service_plans.csv", [
        "subscription", "resource_group", "name", "environment", "sku", "tier", "instance_count", "zone_redundant",
        "hosted_sites", "finding", "severity", "id"]),
}


class Detail:
    """A detail sheet: rows plus a column-letter lookup for formulas."""

    def __init__(self, name, fields, rows):
        self.name = name
        self.fields = fields
        self.rows = rows

    def col(self, field):
        letter = get_column_letter(self.fields.index(field) + 1)
        return f"{self.name}!${letter}:${letter}"


def write_detail(workbook, detail):
    sheet = workbook.create_sheet(detail.name)

    sheet.append(detail.fields)
    for cell in sheet[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for row in detail.rows:
        values = []
        for field in detail.fields:
            value = row.get(field, "")
            if field in NUMERIC_FIELDS and value not in ("", None):
                try:
                    value = float(value)
                except ValueError:
                    pass
            values.append(value)
        sheet.append(values)
    if detail.name == "D_Storage":
        meter_col = get_column_letter(detail.fields.index("disk_meter") + 1)
        size_col = get_column_letter(detail.fields.index("capacity_gb") + 1)
        price_col = detail.fields.index("disk_monthly_usd") + 1
        for row_index in range(2, sheet.max_row + 1):
            sheet.cell(row=row_index, column=price_col, value=(
                f'=IF(A{row_index}<>"MANAGED_DISK",0,IFERROR(INDEX(Gasto_medido!$E:$E,MATCH({meter_col}{row_index},'
                f'Gasto_medido!$B:$B,0)),{size_col}{row_index}*Supuestos!$C$DISKGB))'))
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, field in enumerate(detail.fields, 1):
        sheet.column_dimensions[get_column_letter(index)].width = min(45, max(12, len(field) + 2))
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = NORMAL
    return sheet


NUMERIC_FIELDS = {
    "total_disk_gb", "retention_max_days", "recovery_point_count", "last_backup_age_hours",
    "days_since_last_recovery_point", "configured_max_retention_days", "node_count", "item_count",
    "active_item_count", "max_size_gb", "storage_gb", "native_retention_days", "ltr_max_days",
    "soft_delete_days", "instance_count", "hosted_sites", "capacity_gb", "unattached_days", "snapshot_count",
    "snapshot_total_gb", "blob_soft_delete_days", "pitr_days", "share_soft_delete_days", "snapshot_full_count",
    "used_capacity_gb", "monthly_cost_usd",
    "snapshot_incremental_count", "oldest_snapshot_age_days",
}


# ---------------------------------------------------------------------------
# Cost calibration
# ---------------------------------------------------------------------------


def load_vault_costs(cost_dir, cache_glob):
    """Per-vault monthly cost from cost_by_vault.csv, or from cached query JSON files."""
    costs = defaultdict(lambda: defaultdict(float))
    quantities = defaultdict(float)
    rows = read_csv(os.path.join(cost_dir, "cost_by_vault.csv")) if cost_dir else []
    for row in rows:
        costs[row["vault_id"].lower()][row["meter"]] += float(row["cost_usd"] or 0)
        quantities[row["meter"]] += float(row["usage_quantity"] or 0)
    if not rows:
        for path in glob.glob(cache_glob):
            with open(path, encoding="utf-8") as handle:
                body = json.load(handle)
            columns = [column["name"] for column in body["properties"]["columns"]]
            for values in body["properties"]["rows"]:
                record = dict(zip(columns, values))
                costs[record["ResourceId"].lower()][record["Meter"]] += record["CostUSD"]
                quantities[record["Meter"]] += record["UsageQuantity"]
    return costs, quantities


BACKUP_METER_KEYWORDS = ("backup", "snapshot", "protected instance", "replicated to azure", "point-in-time", "pitr", "ltr ")


def load_meter_costs(cost_dir):
    """Estate cost by meter category and meter: cost_by_meter.csv, or the per-subscription cache."""
    totals = defaultdict(lambda: [0.0, 0.0])
    rows = read_csv(os.path.join(cost_dir, "cost_by_meter.csv")) if cost_dir else []
    if not rows:
        for path in glob.glob(os.path.join(cost_dir, "_cache", "*", "*.json")):
            with open(path, encoding="utf-8") as handle:
                rows.extend(json.load(handle).get("meters", []))
    subscriptions = {row["subscription_id"] for row in rows}
    for row in rows:
        key = (row["meter_category"] or "(sin categoría)", row["meter"] or "(sin medidor)")
        totals[key][0] += float(row["cost_usd"] or 0)
        totals[key][1] += float(row["usage_quantity"] or 0)
    return totals, len(subscriptions)


def load_account_costs(pattern, csv_path=""):
    """Monthly cost per resource id: cost_by_storage_account.csv, or query JSON files grouped by ResourceId."""
    costs = defaultdict(float)
    for row in read_csv(csv_path) if csv_path else []:
        costs[row["resource_id"].lower()] += float(row["cost_usd"] or 0)
    if costs:
        return {key: round(value, 2) for key, value in costs.items()}
    for path in glob.glob(pattern) if pattern else []:
        try:
            with open(path, encoding="utf-8") as handle:
                body = json.load(handle)
        except (ValueError, OSError):
            continue
        columns = [column["name"] for column in body["properties"]["columns"]]
        for values in body["properties"]["rows"]:
            record = dict(zip(columns, values))
            costs[str(record["ResourceId"]).lower()] += record["CostUSD"]
    return {key: round(value, 2) for key, value in costs.items()}


def disk_meter(row):
    """Billing meter name of a managed disk, e.g. 'P30 LRS Disk', derived from SKU and size."""
    sku = row.get("sku") or ""
    redundancy = "ZRS" if "ZRS" in sku.upper() else "LRS"
    tier = row.get("disk_tier") or ""
    if not tier:
        prefix = {"standard_lrs": "S", "standard_zrs": "S", "standardssd_lrs": "E", "standardssd_zrs": "E"}.get(sku.lower())
        if not prefix:
            return ""
        size = float(row.get("capacity_gb") or 0)
        for limit, number in ((4, 1), (8, 2), (16, 3), (32, 4), (64, 6), (128, 10), (256, 15), (512, 20), (1024, 30),
                              (2048, 40), (4096, 50), (8192, 60), (16384, 70), (32767, 80)):
            if size <= limit:
                tier = f"{prefix}{number}"
                break
    return f"{tier} {redundancy} Disk" if tier else ""


def build_meter_sheet(workbook, meters, subscription_count):
    sheet = workbook.create_sheet("Gasto_medido")
    headers = ["Categoría", "Medidor", "Costo oct-2026 (USD)", "Cantidad", "Precio unitario efectivo", "Relacionado a backup"]
    sheet.append(headers)
    style_header(sheet, 1, len(headers))
    for (category, meter), (cost, quantity) in sorted(meters.items(), key=lambda pair: -pair[1][0]):
        if cost < 1:
            continue
        related = "SI" if any(word in meter.lower() for word in BACKUP_METER_KEYWORDS) or category == "Backup" else "NO"
        sheet.append([category, meter, round(cost, 2), round(quantity, 4), None, related])
        row = sheet.max_row
        sheet.cell(row=row, column=5, value=f"=IF(D{row}=0,0,C{row}/D{row})")
        sheet.cell(row=row, column=3).number_format = USD
        sheet.cell(row=row, column=5).number_format = "#,##0.0000"
    note = sheet.max_row + 2
    sheet.cell(row=note, column=1, value=(
        f"Fuente: Cost Management, ActualCost oct-2026, {subscription_count} suscripciones, agrupado por tipo de recurso, "
        "categoría y medidor. 'Relacionado a backup' marca medidores de bóvedas, snapshots y backup de bases de datos."))
    set_widths(sheet, [26, 48, 18, 16, 18, 14])
    return sheet


def calibration_rows(vault_costs, vaults, orphans):
    orphan_count = Counter(row["vault_name"].lower() for row in orphans)
    by_name = {row["vault_id"].lower(): row for row in vaults}
    rows = []
    for vault_id, meters in sorted(vault_costs.items(), key=lambda pair: -sum(pair[1].values())):
        vault = by_name.get(vault_id, {})
        name = vault.get("vault_name") or vault_id.split("/")[-1]
        rows.append({
            "vault_name": name,
            "storage_redundancy": vault.get("storage_redundancy", ""),
            "item_count": int(vault.get("item_count") or 0),
            "orphan_items": min(orphan_count.get(name.lower(), 0), int(vault.get("item_count") or 0)),
            "storage_usd": round(sum(value for key, value in meters.items() if "data stored" in key.lower()), 2),
            "protected_instance_usd": round(sum(value for key, value in meters.items() if "protected instance" in key.lower()), 2),
            "other_usd": round(sum(value for key, value in meters.items()
                                   if "data stored" not in key.lower() and "protected instance" not in key.lower()), 2),
        })
    return rows


# ---------------------------------------------------------------------------
# Workbook sections
# ---------------------------------------------------------------------------


def style_header(sheet, row, columns):
    for column in range(1, columns + 1):
        cell = sheet.cell(row=row, column=column)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def set_widths(sheet, widths):
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def countifs(detail, *pairs):
    """COUNTIFS over a detail sheet: countifs(d, ("environment", "PROD"), ...)."""
    if detail is None:
        return 0
    parts = []
    for field, criteria in pairs:
        parts.append(detail.col(field))
        parts.append(f'"{criteria}"')
    return "=COUNTIFS(" + ",".join(parts) + ")"


def build_assumptions(workbook, calibration, quantities):
    sheet = workbook.create_sheet("Supuestos", 1)
    sheet["A1"] = "Supuestos y tarifas (celdas amarillas: editables)"
    sheet["A1"].font = TITLE_FONT
    storage_usd = sum(row["storage_usd"] for row in calibration)
    pi_usd = sum(row["protected_instance_usd"] for row in calibration)
    storage_gb = sum(value for key, value in quantities.items() if "data stored" in key.lower())
    pi_units = sum(value for key, value in quantities.items() if "protected instance" in key.lower())
    rows = [
        ("pi_usd", "Azure Backup: costo por instancia protegida (USD/mes)",
         round(pi_usd / pi_units, 2) if pi_units else 7.0,
         "Tarifa efectiva medida en Cost Management (oct-2026) en las suscripciones calibradas: "
         "costo 'Azure VM Protected Instances' / cantidad facturada. Incluye descuento corporativo."),
        ("grs_gb_usd", "Azure Backup: almacenamiento GRS (USD/GB-mes)",
         round(storage_usd / storage_gb, 4) if storage_gb else 0.0314,
         "Tarifa efectiva medida: costo 'GRS Data Stored' / GB-mes facturados (oct-2026)."),
        ("lrs_factor", "Factor de costo LRS respecto a GRS", 0.5,
         "Supuesto: LRS cuesta aproximadamente la mitad que GRS en Azure Backup. Validar con la hoja de precios del contrato."),
        ("retention_reduction", "Reducción de almacenamiento al pasar de 360 diarios a esquema GFS (30 diarios + 12 mensuales)",
         0.3, "Supuesto conservador. Depende de la tasa de cambio diario de cada VM; medir por bóveda tras el cambio."),
        ("live_item_storage_usd", "Almacenamiento de backup por VM viva (USD/mes)",
         None, "Fórmula: almacenamiento medido no atribuible a huérfanos / ítems no huérfanos en bóvedas calibradas."),
        ("redis_premium_delta", "Diferencia mensual Redis Premium P1 vs Standard de memoria equivalente (USD/mes)",
         77.0, "Supuesto a validar: diferencia de precio lista referencial entre P1 (6 GB) y C3 (6 GB). "
               "Reemplazar por la tarifa del contrato antes de presentar."),
        ("asp_empty_usd", "Costo promedio de un App Service Plan vacío (USD/mes)", 40.0,
         "Supuesto a validar: mezcla observada de planes vacíos (mayoría Basic, algunos P0v3/P1v3)."),
        ("disk_gb_usd", "Disco sin tier conocido (Premium v2/Ultra): costo (USD/GB-mes)", 0.08,
         "Supuesto a validar, solo para discos cuyo medidor no aparece en Gasto_medido. El resto usa el precio efectivo por tier."),
        ("snapshot_spend_usd", "Gasto mensual medido en snapshots de disco (USD)", None,
         "Fórmula: suma de medidores 'Snapshots' en Gasto_medido (oct-2026)."),
        ("snapshot_removable", "Fracción eliminable de snapshots antiguos (>1 año) tras política de expiración", 0.5,
         "Supuesto: se conserva la mitad (último snapshot por disco y los exigidos). Ajustar tras revisión con dueños."),
        ("geo_not_required", "Fracción del gasto de geo-replicación de storage sin requisito de DR", 0.25,
         "Supuesto a validar por cuenta: no todas las cuentas GRS/RA-GRS necesitan copia en otra región."),
        ("calibrated_items", "Ítems de backup en las bóvedas calibradas", sum(row["item_count"] for row in calibration),
         "Conteo desde compute_backup_vaults.csv para las bóvedas con costo medido."),
        ("total_items", "Ítems de backup totales (RSV + Backup vault)", None, "Fórmula sobre D_Bovedas."),
        ("calibrated_storage_usd", "Almacenamiento medido en bóvedas calibradas (USD/mes)", round(storage_usd, 2),
         "Suma de 'GRS Data Stored' en Costo_calibrado."),
    ]
    sheet.append([])
    sheet.append(["Clave", "Supuesto", "Valor", "Fuente / nota"])
    style_header(sheet, 3, 4)
    names = {}
    for key, label, value, note in rows:
        sheet.append([key, label, value, note])
        row = sheet.max_row
        names[key] = f"Supuestos!$C${row}"
        cell = sheet.cell(row=row, column=3)
        if value is not None:
            cell.font = BLUE
            cell.fill = INPUT_FILL
        for column in (1, 2, 4):
            sheet.cell(row=row, column=column).alignment = WRAP
    set_widths(sheet, [22, 55, 14, 90])
    return sheet, names


def fill_formula_assumptions(sheet, names, details):
    vaults = details.get("D_Bovedas")
    total_cell = names["total_items"].split("!")[1].replace("$", "")
    sheet[total_cell] = f"=SUM({vaults.col('item_count')})" if vaults else 0
    live_cell = names["live_item_storage_usd"].split("!")[1].replace("$", "")
    sheet[live_cell] = "=IF(Costo_calibrado!C2-Costo_calibrado!D2<=0,0,(Costo_calibrado!E2-Costo_calibrado!H2)/(Costo_calibrado!C2-Costo_calibrado!D2))"
    sheet[live_cell].number_format = '"$"#,##0.00'
    snapshot_cell = names["snapshot_spend_usd"].split("!")[1].replace("$", "")
    sheet[snapshot_cell] = '=SUMIFS(Gasto_medido!$C:$C,Gasto_medido!$B:$B,"*Snapshots*")'
    sheet[snapshot_cell].number_format = USD
    for key in ("pi_usd", "redis_premium_delta", "asp_empty_usd", "calibrated_storage_usd"):
        cell = names[key].split("!")[1].replace("$", "")
        sheet[cell].number_format = '"$"#,##0.00'


CA = "Cost Management > Cost analysis, ámbito suscripción, Costo real, oct-2026"
OPPORTUNITY_SOURCES = {
    "O1": ("Resource Graph: ítems protegidos cuyo recurso origen ya no existe (compute_orphan_backup_items.csv)",
           "Cost Management: costo por bóveda y medidor (cost_by_vault.csv) + tarifa efectiva por instancia",
           "Backup center > Backup instances (estado del origen) y " + CA + ", filtro Tipo de recurso = Recovery Services vaults, agrupar por Recurso"),
    "O2": ("Resource Graph: políticas con 360 puntos diarios (compute_vm_inventory.csv)",
           "Cost Management: almacenamiento GRS por VM viva (derivado) x % de reducción supuesto",
           "Bóveda > Backup policies > retención; " + CA + ", medidor 'GRS Data Stored'"),
    "O3": ("Resource Graph: VMs PROD apagadas con backup activo", "Tarifa efectiva por instancia protegida (Cost Management)",
           "Virtual machines (estado Deallocated) y bóveda > Backup items"),
    "O4": ("Resource Graph: VMs CERT/DESA con backup", "Tarifa efectiva por instancia protegida (Cost Management)",
           "Bóveda > Backup items, filtrar por VM"),
    "O5": ("Resource Graph: bases SQL con redundancia Geo en CERT/DESA", "Cost Management: medidores 'Backup ... Data Stored' de SQL Database",
           CA + ", Categoría del medidor = SQL Database, medidor contiene 'Backup'"),
    "O6": ("Inventario nativo: Redis Premium en CERT/DESA", "Supuesto: diferencia de precio (a validar con la hoja de precios del contrato)",
           "Azure Pricing Calculator o hoja de precios del contrato; " + CA + ", Tipo de recurso = Redis"),
    "O7": ("Inventario nativo: App Service Plans sin apps", "Supuesto: costo promedio por plan (a validar)",
           "App Service plans > Apps = 0; " + CA + ", agrupar por Recurso"),
    "O8": ("Inventario nativo: planes no productivos con zone redundancy", "Por medir", CA + ", agrupar por Recurso"),
    "O9": ("Resource Graph: bóvedas sin ítems y políticas sin uso", "Sin costo directo", "Bóveda > Overview > Backup items = 0"),
    "O10": ("Resource Graph: snapshots cuyo disco origen no existe (storage_snapshots_by_source.csv)",
            "Cost Management: gasto en medidores 'Snapshots' prorrateado por GB",
            "Snapshots (columna Source disk) y " + CA + ", medidor contiene 'Snapshots'"),
    "O11": ("Resource Graph: discos con estado Unattached (storage_inventory.csv)",
            "Cost Management: precio efectivo por tier de disco (medidor '<tier> LRS Disk' costo / cantidad)",
            "Disks, filtro Disk state = Unattached; " + CA + ", medidor '<tier> LRS Disk'"),
    "O12": ("Resource Graph: snapshots de más de 1 año de discos vivos", "Cost Management: gasto 'Snapshots' x % eliminable supuesto",
            "Snapshots, ordenar por Time created; " + CA + ", medidor contiene 'Snapshots'"),
    "O13": ("Resource Graph: cuentas con redundancia GRS/RA-GRS", "Cost Management: medidor 'Geo-Replication v2 Data transfer' x % supuesto",
            "Storage account > Redundancy; " + CA + ", medidor 'Geo-Replication v2 Data transfer'"),
}


def build_opportunities(workbook, details, names):
    sheet = workbook.create_sheet("Oportunidades", 1)
    sheet["A1"] = "Oportunidades de ahorro y racionalización de backup"
    sheet["A1"].font = TITLE_FONT
    sheet["A2"] = ("Ahorro mensual = recursos x tarifa (hoja Supuestos). 'Piso' = solo instancia protegida; "
                   "el almacenamiento liberado se suma donde hay medición.")
    sheet["A2"].alignment = WRAP
    headers = ["ID", "Frente", "Oportunidad", "Recursos", "Ahorro mensual estimado (USD)", "Ahorro anual (USD)",
               "Confianza", "Qué cambiar", "Riesgo / condición", "Fuente de la cantidad", "Fuente del costo",
               "Cómo verificarlo en el portal"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    vm = details.get("D_VMs")
    orphans = details.get("D_Huerfanos")
    db = details.get("D_BD")
    native = details.get("D_Nativo")
    plans = details.get("D_Planes")
    pi, grs = names["pi_usd"], names["grs_gb_usd"]
    items = [
        ("O1", "Compute", "Ítems de backup huérfanos (VM eliminada o sin backup >30 días)",
         f"=COUNTA({orphans.col('orphan_reason')})-1" if orphans else 0,
         "=D{r}*" + pi + "+Costo_calibrado!$H$2",
         "Media: piso medido en hoja Costo_calibrado (bóvedas 100% huérfanas); resto prorrateado",
         "Validar con el dueño de la app, detener protección y eliminar datos de backup (respetando retención regulatoria).",
         "Si existe requisito legal de conservar el último punto, mover a retención mínima y no borrar."),
        ("O2", "Compute", "Retención de 360 puntos diarios en VMs (pasar a esquema GFS: 30 diarios + 12 mensuales)",
         countifs(vm, ("backup_enabled", "YES"), ("retention_max_days", "360")),
         "=D{r}*" + names["live_item_storage_usd"] + "*" + names["retention_reduction"],
         "Media (supuesto de reducción)",
         "Cambiar política: diario 30 días + mensual 12 meses. Los puntos existentes expiran según la política anterior.",
         "Confirmar RPO/RTO y requisito regulatorio por aplicación antes de recortar."),
        ("O3", "Compute", "VMs PROD apagadas (deallocated) con backup activo",
         countifs(vm, ("finding", "DEALLOCATED_VM_WITH_BACKUP")), "=D{r}*" + pi, "Alta (instancia)",
         "Si la VM está en desuso: detener backup con retención o eliminar VM y backup.",
         "Una VM apagada puede ser de contingencia: validar con el dueño."),
        ("O4", "Compute", "Backup en VMs no productivas (CERT/DESA)",
         countifs(vm, ("finding", "REVIEW_NONPROD_BACKUP")), "=D{r}*" + pi, "Alta (instancia)",
         "Quitar backup si la VM se reconstruye por IaC; si no, política mínima de 7 días en bóveda LRS.",
         "Confirmar que no hay datos únicos en la VM."),
        ("O5", "Database", "Backup SQL geo-redundante en no productivo (Geo → Local)",
         countifs(db, ("finding", "NONPROD_GEO_BACKUP_REDUNDANCY")), 0,
         "Medida: bajo",
         "Cambiar backup storage redundancy a Local/Zone en CERT/DESA.",
         "Ahorro casi nulo: el storage de PITR está mayormente dentro del tramo gratuito (US$3/mes en la suscripción con más bases SQL). Prioridad baja."),
        ("O6", "Nativo", "Redis Premium en CERT/DESA",
         (f"=COUNTIFS({native.col('service_type')},\"REDIS\",{native.col('finding')},\"NONPROD_PREMIUM_SKU\")"
          if native else 0),
         "=D{r}*" + names["redis_premium_delta"], "Baja (precio a validar)",
         "Pasar a Standard de memoria equivalente donde no se requiera VNet injection, clustering ni persistencia.",
         "Premium se usa a veces por red privada: verificar dependencia antes de cambiar."),
        ("O7", "Nativo", "App Service Plans vacíos",
         (f"=COUNTIFS({plans.col('finding')},\"EMPTY_APP_SERVICE_PLAN\")" if plans else 0),
         "=D{r}*" + names["asp_empty_usd"], "Media (precio promedio supuesto)",
         "Eliminar planes sin apps o consolidar.", "Ninguno si el plan está realmente vacío."),
        ("O8", "Nativo", "Planes no productivos con redundancia de zona",
         (f"=COUNTIFS({plans.col('finding')},\"NONPROD_ZONE_REDUNDANT_PLAN\")" if plans else 0), 0,
         "Por medir", "Desactivar zone redundancy en CERT/DESA (requiere recrear el plan).",
         "Zone redundancy exige mínimo 3 instancias: medir costo real por plan."),
        ("O10", "Storage", "Snapshots de discos que ya no existen",
         (f"=SUMIFS({details['D_Snapshots'].col('snapshot_count')},{details['D_Snapshots'].col('source_exists')},\"NOT_IN_SCOPE_OR_DELETED\")"
          if details.get("D_Snapshots") else 0),
         ((f"={names['snapshot_spend_usd']}*SUMIFS({details['D_Snapshots'].col('snapshot_total_gb')},"
           f"{details['D_Snapshots'].col('source_exists')},\"NOT_IN_SCOPE_OR_DELETED\")/MAX(1,SUM({details['D_Snapshots'].col('snapshot_total_gb')}))")
          if details.get("D_Snapshots") else 0),
         "Media (gasto medido, prorrateo por GB)",
         "Eliminar snapshots cuyo disco origen ya no existe, salvo los retenidos por requisito.",
         "Algunos pueden estar en suscripciones no visibles: confirmar antes de borrar."),
        ("O12", "Storage", "Snapshots completos de más de 1 año (pasar a incrementales y expirar)",
         (f"=COUNTIFS({details['D_Snapshots'].col('source_exists')},\"YES\",{details['D_Snapshots'].col('oldest_snapshot_age_days')},\">365\")"
          if details.get("D_Snapshots") else 0),
         ((f"={names['snapshot_spend_usd']}*SUMIFS({details['D_Snapshots'].col('snapshot_total_gb')},"
           f"{details['D_Snapshots'].col('source_exists')},\"YES\",{details['D_Snapshots'].col('oldest_snapshot_age_days')},\">365\")"
           f"/MAX(1,SUM({details['D_Snapshots'].col('snapshot_total_gb')}))*{names['snapshot_removable']}")
          if details.get("D_Snapshots") else 0),
         "Media (gasto medido; fracción eliminable supuesta)",
         "Snapshots incrementales y política de expiración; el 95% de los snapshots actuales son completos.",
         "Puede ser el único respaldo del disco: cruzar con discos PROD sin backup."),
        ("O13", "Storage", "Geo-replicación de storage (GRS/RA-GRS) donde no se requiere DR geográfico",
         (f"=COUNTIFS({details['D_Storage'].col('redundancy')},\"*GRS*\")" if details.get("D_Storage") else 0),
         '=SUMIFS(Gasto_medido!$C:$C,Gasto_medido!$B:$B,"Geo-Replication*")*' + names["geo_not_required"],
         "Baja (gasto medido; fracción sin requisito supuesta)",
         "Validar requisito de DR por cuenta; pasar a ZRS donde no aplique (elimina la transferencia de geo-replicación).",
         "La copia geográfica puede ser parte del plan de DR: decidir por cuenta con el dueño."),
        ("O11", "Storage", "Discos administrados sin adjuntar",
         (f"=COUNTIFS({details['D_Storage'].col('all_findings')},\"*UNATTACHED_DISK*\",{details['D_Storage'].col('backup_scope')},\"<>EXCLUDED_PLATFORM\")"
          if details.get("D_Storage") else 0),
         ((f"=SUMIFS({details['D_Storage'].col('disk_monthly_usd')},{details['D_Storage'].col('all_findings')},\"*UNATTACHED_DISK*\","
           f"{details['D_Storage'].col('backup_scope')},\"<>EXCLUDED_PLATFORM\")")
          if details.get("D_Storage") else 0),
         "Alta (precio efectivo por tier medido)",
         "Tomar snapshot final si aplica y eliminar el disco.",
         "Confirmar que no es un disco de DR o de una VM en reconstrucción."),
        ("O9", "Gobierno", "Bóvedas vacías y políticas sin uso",
         (f"=COUNTIFS({details['D_Bovedas'].col('item_count')},0)" if details.get("D_Bovedas") else 0), 0,
         "Sin costo directo", "Eliminar bóvedas sin ítems y políticas por defecto no usadas para simplificar gobierno.",
         "Bóvedas vacías no facturan; el beneficio es operativo y de auditoría."),
    ]
    first = sheet.max_row + 1
    for item in items:
        row = sheet.max_row + 1
        identifier, front, title, count, saving, confidence, change, risk = item
        saving_value = saving.format(r=row) if isinstance(saving, str) else saving
        quantity_source, cost_source, check = OPPORTUNITY_SOURCES.get(identifier, ("", "", ""))
        sheet.append([identifier, front, title, count, saving_value, f"=E{row}*12", confidence, change, risk,
                      quantity_source, cost_source, check])
        for column in range(1, 13):
            sheet.cell(row=row, column=column).alignment = WRAP
            sheet.cell(row=row, column=column).font = NORMAL
        sheet.cell(row=row, column=5).number_format = USD
        sheet.cell(row=row, column=6).number_format = USD
    last = sheet.max_row
    total_row = last + 1
    sheet.cell(row=total_row, column=3, value="Total estimado").font = BOLD
    sheet.cell(row=total_row, column=5, value=f"=SUM(E{first}:E{last})").number_format = USD
    sheet.cell(row=total_row, column=6, value=f"=SUM(F{first}:F{last})").number_format = USD
    sheet.cell(row=total_row, column=5).font = BOLD
    sheet.cell(row=total_row, column=6).font = BOLD
    set_widths(sheet, [6, 11, 48, 11, 16, 16, 18, 55, 55, 45, 45, 50])
    return sheet


def build_risks(workbook, details):
    sheet = workbook.create_sheet("Riesgos", 2)
    sheet["A1"] = "Brechas de protección y riesgos (priorizados)"
    sheet["A1"].font = TITLE_FONT
    headers = ["ID", "Frente", "Riesgo", "Recursos", "Severidad", "Acción recomendada"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    vm, db, k8s = details.get("D_VMs"), details.get("D_BD"), details.get("D_K8s")
    native, vaults = details.get("D_Nativo"), details.get("D_Bovedas")
    rows = [
        ("R1", "Compute", "VMs PROD tradicionales sin backup", countifs(vm, ("finding", "PROD_WITHOUT_BACKUP")), "ALTA",
         "Confirmar si tienen datos únicos; activar política PROD o documentar que se reconstruyen por IaC."),
        ("R2", "Compute", "VMs PROD con backup fallido o sin punto reciente (>48 h)",
         countifs(vm, ("finding", "PROD_BACKUP_UNHEALTHY")), "ALTA", "Revisar jobs fallidos y agente; alertar por Azure Monitor."),
        ("R3", "Database", "SQL Server en VM PROD sin backup de base de datos (solo snapshot de VM o ninguno)",
         countifs(db, ("finding", "PROD_DB_WORKLOAD_WITHOUT_BACKUP")), "ALTA",
         "Habilitar backup de workload SQL en Recovery Services (log backups) para restauración a punto en el tiempo."),
        ("R4", "Database", "Bases PROD sin backup geo-redundante (incluye todos los PostgreSQL/MySQL)",
         countifs(db, ("finding", "PROD_BACKUP_NOT_GEO_REDUNDANT")), "MEDIA",
         "Para PostgreSQL/MySQL Flexible el geo-backup solo se habilita al crear el servidor: planificar recreación o réplica."),
        ("R5", "Database", "Cosmos DB PROD con retención por debajo del mínimo (modo periódico)",
         countifs(db, ("finding", "PROD_RETENTION_BELOW_MIN")), "MEDIA",
         "Migrar a backup continuo de 7 días (restauración self-service; el periódico requiere ticket a soporte)."),
        ("R6", "Compute", "Clusters AKS PROD sin backup (estado de PV desconocido)",
         countifs(k8s, ("finding", "PROD_CLUSTER_WITHOUT_BACKUP")), "MEDIA",
         "Identificar namespaces con volúmenes persistentes; AKS Backup solo para esos. El resto se reconstruye por GitOps."),
        ("R7", "Gobierno", "Bóvedas sin inmutabilidad", (f"=COUNTA({vaults.col('vault_name')})-1-COUNTIFS({vaults.col('immutability')},\"Enabled\")-COUNTIFS({vaults.col('immutability')},\"Locked\")" if vaults else 0),
         "MEDIA", "Habilitar inmutabilidad y MUA en bóvedas PROD (protección ante ransomware y borrado malicioso)."),
        ("R8", "Gobierno", "Bóvedas GRS sin Cross Region Restore",
         (f"=COUNTIFS({vaults.col('storage_redundancy')},\"GeoRedundant\",{vaults.col('cross_region_restore')},\"Disabled\")"
          if vaults else 0), "MEDIA",
         "Sin CRR la copia secundaria solo se usa si Microsoft declara desastre regional: activar CRR en PROD crítico o bajar a LRS/ZRS."),
        ("R9", "Nativo", "Zonas DNS PROD sin delete lock", countifs(native, ("finding", "PROD_DNS_ZONE_WITHOUT_LOCK")), "ALTA",
         "Aplicar CanNotDelete en zonas DNS (borrar una zona privada rompe la resolución de los private endpoints)."),
        ("R10", "Nativo", "Red core PROD sin delete lock", countifs(native, ("finding", "PROD_NETWORK_CORE_WITHOUT_LOCK")), "MEDIA",
         "Locks en VNets, gateways, ExpressRoute, App Gateway y Front Door de PROD."),
        ("R11", "Nativo", "Key Vault PROD sin purge protection",
         countifs(native, ("finding", "PROD_KEYVAULT_WITHOUT_PURGE_PROTECTION")), "ALTA",
         "Habilitar purge protection (irreversible, sin costo)."),
        ("R13", "Storage", "Cuentas y discos PROD sin ningún mecanismo de recuperación",
         countifs(details.get("D_Storage"), ("finding", "PROD_WITHOUT_RECOVERY")), "ALTA",
         "Habilitar soft delete (blob/contenedor) y versioning en cuentas con dato final; backup vaulted si es crítico."),
        ("R14", "Storage", "File shares PROD sin Azure Backup",
         countifs(details.get("D_Storage"), ("finding", "PROD_SHARE_WITHOUT_BACKUP")), "MEDIA",
         "Azure Backup para shares con dato de negocio; el soft delete solo cubre borrado del share completo."),
        ("R15", "Storage", "Storage PROD LRS sin backup (una sola copia en un datacenter)",
         countifs(details.get("D_Storage"), ("finding", "PROD_LRS_NO_BACKUP")), "MEDIA",
         "ZRS para dato final de PROD o backup en bóveda; LRS solo para dato reproducible."),
        ("R12", "Todos", "Recursos sin ambiente identificable (tag ausente)",
         "=" + "+".join(f"COUNTIFS({detail.col('finding')},\"REQUIRES_ENVIRONMENT_CLASSIFICATION\")"
                        for detail in (vm, db, native, details.get("D_Storage")) if detail), "BAJA",
         "Corregir tag environment: sin él no se puede aplicar política por ambiente."),
    ]
    for identifier, front, title, count, severity, action in rows:
        sheet.append([identifier, front, title, count, severity, action])
        row = sheet.max_row
        for column in range(1, 7):
            sheet.cell(row=row, column=column).alignment = WRAP
            sheet.cell(row=row, column=column).font = NORMAL
    set_widths(sheet, [6, 11, 60, 11, 11, 80])


def build_long_term(workbook, details, versions):
    sheet = workbook.create_sheet("Restauracion_LP", 3)
    sheet["A1"] = "Riesgos de restauración a largo plazo (compatibilidad de versiones y dependencias)"
    sheet["A1"].font = TITLE_FONT
    headers = ["ID", "Escenario", "Recursos", "Por qué falla la restauración", "Mitigación"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    vm, db, orphans = details.get("D_VMs"), details.get("D_BD"), details.get("D_Huerfanos")
    rows = [
        ("L1", "Backups de VM con retención de 4 a 12 años",
         (f"=COUNTIFS({vm.col('backup_enabled')},\"YES\",{vm.col('retention_max_days')},\">=1460\")" if vm else 0),
         "Restaura la imagen completa: SO, agentes y librerías de ese momento. En 2036 un Windows Server 2019 o un RHEL 8 "
         "estarán fuera de soporte, sin parches y posiblemente sin tamaños de VM compatibles.",
         "Retener a largo plazo el DATO (LTR de base de datos, exportes en formato abierto en blob inmutable/Archive), "
         "no la imagen de VM. VM: 30 días + 12 mensuales como máximo."),
        ("L2", "Ítems huérfanos con retención de 10 años",
         (f"=COUNTIFS({orphans.col('configured_max_retention_days')},\">=3650\")" if orphans else 0),
         "La VM origen ya no existe: no hay documentación viva de su configuración, red ni dependencias para restaurarla.",
         "Si hay requisito regulatorio, extraer el dato y documentar; si no, eliminar."),
        ("L3", "VMs Gen1 (BIOS) con backup",
         versions.get("gen1_with_backup", 0),
         "Las familias de VM nuevas solo soportan Gen2. Restaurar un backup Gen1 dentro de años puede no tener tamaño disponible.",
         "Plan de migración Gen1→Gen2/Trusted Launch para VMs con retención larga."),
        ("L4", "VMs con disco cifrado con clave del cliente (Disk Encryption Set)",
         versions.get("cmk_vms_with_backup", 0),
         "La restauración requiere la versión de clave usada al momento del backup. Si se rota y elimina la clave, el backup no se puede restaurar.",
         "Purge protection en el Key Vault y conservar versiones de clave al menos tanto como la retención del backup."),
        ("L5", "SQL Database con LTR de 5 a 10 años",
         (f"=COUNTIFS({db.col('ltr_max_days')},\">=1825\")" if db else 0),
         "Azure SQL restaura sobre el motor actual (sin versión fija): el riesgo es la aplicación, no el motor. "
         "Un esquema de hace 5 años puede no ser legible por la versión actual de la app.",
         "Prueba de restauración anual de una muestra y versionar el esquema junto con el backup."),
        ("L6", "PostgreSQL / MySQL en versiones fuera de soporte comunitario",
         versions.get("eol_db_engines", 0),
         "El backup solo restaura a la misma versión mayor. Cuando Azure retira la versión, no hay dónde restaurar y se paga extended support.",
         "Actualizar a versión soportada (PG 16+/MySQL 8.4) y no depender del backup como archivo histórico."),
        ("L7", "Clusters AKS en versiones fuera de soporte",
         versions.get("aks_unsupported", 0),
         "Los manifiestos en Git pueden usar APIs deprecadas: la reconstrucción por GitOps falla en una versión nueva.",
         "Prueba de reconstrucción en versión actual y lint de APIs deprecadas en el pipeline."),
        ("L8", "Cosmos DB en modo periódico",
         (f"=COUNTIFS({db.col('service_type')},\"COSMOS_DB\",{db.col('native_retention_days')},\"<1\")" if db else 0),
         "La restauración en modo periódico requiere ticket a soporte y solo hay 8 h de ventana: RTO no controlado.",
         "Backup continuo de 7 días (gratuito) para PROD."),
    ]
    for identifier, title, count, why, mitigation in rows:
        sheet.append([identifier, title, count, why, mitigation])
        row = sheet.max_row
        for column in range(1, 6):
            sheet.cell(row=row, column=column).alignment = WRAP
            sheet.cell(row=row, column=column).font = NORMAL
    set_widths(sheet, [6, 40, 11, 70, 70])


MINIMUM_POLICY = [
    ("VM con estado (IaaS)", "Diario; 30 días + 12 mensuales; bóveda GRS con CRR, inmutable",
     "Diario; 30 días; bóveda ZRS/LRS", "7 días o sin backup si se reconstruye por IaC; LRS",
     "Sin backup (IaC); excepción documentada"),
    ("VM sin estado / reconstruible", "Sin backup: imagen + IaC", "Sin backup", "Sin backup", "Sin backup"),
    ("SQL Database", "PITR 14 días; LTR solo si hay requisito (W4/M12/Y10); Geo", "PITR 7 días; Zone",
     "PITR 7 días; Local", "PITR 1-7 días; Local"),
    ("PostgreSQL / MySQL Flexible", "35 días + geo-backup (al crear)", "14 días", "7 días", "7 días"),
    ("Cosmos DB", "Continuo 7 días", "Continuo 7 días", "Periódico, Local", "Periódico, Local"),
    ("SQL Server en VM", "Backup de workload SQL (full + log 15 min)", "Workload SQL diario", "Snapshot VM 7 días", "Sin backup"),
    ("Blob / ADLS", "Soft delete 14 d + versioning + PITR o backup vaulted si es dato final",
     "Soft delete 14 d + versioning", "Soft delete 7 d", "Soft delete 7 d"),
    ("Azure Files", "Azure Backup 30 días + soft delete", "Snapshot 14 días", "Soft delete 7 d", "Soft delete 7 d"),
    ("AKS", "GitOps + AKS Backup solo en namespaces con PV", "GitOps + backup de PV críticos", "GitOps", "GitOps"),
    ("Key Vault", "Soft delete 90 + purge protection", "Igual", "Igual", "Igual"),
    ("Red / DNS", "IaC + delete locks", "IaC + delete locks", "IaC", "IaC"),
    ("Redis", "Sin backup si es caché; si es almacén de datos, persistencia Premium", "Caché sin backup", "Standard", "Standard/Basic"),
]


def build_policy(workbook):
    sheet = workbook.create_sheet("Politica_minima", 4)
    sheet["A1"] = "Política mínima propuesta por servicio y ambiente (punto de partida para validar con Riesgos y Cumplimiento)"
    sheet["A1"].font = TITLE_FONT
    headers = ["Servicio", "PROD crítico", "PROD no crítico", "CERT", "DESA"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    for row in MINIMUM_POLICY:
        sheet.append(list(row))
        for column in range(1, 6):
            sheet.cell(row=sheet.max_row, column=column).alignment = WRAP
            sheet.cell(row=sheet.max_row, column=column).font = NORMAL
    set_widths(sheet, [28, 45, 32, 32, 28])


def build_calibration(workbook, calibration):
    sheet = workbook.create_sheet("Costo_calibrado")
    headers = ["Bóveda", "Redundancia", "Ítems", "Ítems huérfanos", "Almacenamiento (USD)", "Instancias (USD)", "Otros (USD)",
               "Almacenamiento atribuible a huérfanos (USD)", "Proporción huérfanos",
               "Costo de bóvedas 100% huérfanas (USD, medido)"]
    sheet.append(headers)
    style_header(sheet, 1, len(headers))
    # Row 2 holds the totals the opportunity formulas reference.
    sheet.append(["TOTAL", "", "=SUM(C3:C1000)", "=SUM(D3:D1000)", "=SUM(E3:E1000)", "=SUM(F3:F1000)", "=SUM(G3:G1000)",
                  "=SUM(H3:H1000)", "=IF(C2=0,0,D2/C2)", "=SUM(J3:J1000)"])
    for cell in sheet[2]:
        cell.font = BOLD
    for row in calibration:
        sheet.append([row["vault_name"], row["storage_redundancy"], row["item_count"], row["orphan_items"],
                      row["storage_usd"], row["protected_instance_usd"], row["other_usd"]])
        current = sheet.max_row
        sheet.cell(row=current, column=8, value=f"=IF(C{current}=0,0,E{current}*D{current}/C{current})")
        sheet.cell(row=current, column=9, value=f"=IF(C{current}=0,0,D{current}/C{current})")
        sheet.cell(row=current, column=10, value=f"=IF(AND(C{current}>0,D{current}=C{current}),E{current}+F{current}+G{current},0)")
    for row in sheet.iter_rows(min_row=2):
        for index in (5, 6, 7, 8, 10):
            row[index - 1].number_format = USD
        row[8].number_format = "0.0%"
    note_row = sheet.max_row + 2
    sheet.cell(row=note_row, column=1, value=(
        "Fuente: Cost Management, ActualCost oct-2026, agrupado por bóveda y medidor, en las suscripciones con más "
        "ítems huérfanos. El almacenamiento atribuible a huérfanos se prorratea por cantidad de ítems en cada bóveda."))
    set_widths(sheet, [30, 16, 8, 12, 16, 14, 12, 20, 12, 20])


THIRD_PARTY_INPUTS = [
    ("d3_active", "Tarifa D3 Active", 56.01, "Tarifa compartida por el proveedor (reunión de revisión de lineamientos de backup, 05-nov). "
     "Supuesto de unidad: USD por TB-mes. Confirmar si es TB protegido (front-end) o TB almacenado (back-end)."),
    ("d3_archive", "Tarifa D3 Archive", 7.92, "Idem. Archive: puntos de largo plazo (mensuales/anuales)."),
    ("o1_active", "Tarifa Opción 1 Active", 49.90, "Idem."),
    ("o1_archive", "Tarifa Opción 1 Archive", 12.84, "Idem."),
    ("archive_share", "Fracción del volumen que iría a Archive", 0.5,
     "Supuesto: puntos mensuales/anuales. Con el esquema GFS propuesto, la mayor parte del volumen es largo plazo."),
    ("include_snapshots", "Incluir snapshots manuales de disco en el alcance (1 = sí, 0 = no)", 0,
     "Los snapshots manuales no son Azure Backup; incluirlos solo si el tercero los reemplazaría."),
    ("egress_gb", "Costo de salida de datos si el repositorio está fuera de Azure (USD/GB)", 0.0,
     "0 si el repositorio está en Azure. Si sale por ExpressRoute o Internet, usar la tarifa del contrato."),
    ("daily_change", "Tasa de cambio diaria del dato protegido", 0.02, "Supuesto: volumen incremental que se transfiere cada día."),
    ("link_gbps", "Ancho de banda disponible para restaurar desde el repositorio (Gbps)", 10, "Supuesto a validar."),
    ("link_efficiency", "Eficiencia efectiva del enlace", 0.7, "Supuesto: protocolo, deduplicación y concurrencia."),
    ("restore_tb_small", "Restauración puntual (TB)", 1, "Ejemplo: una VM grande."),
    ("restore_tb_large", "Restauración masiva (TB)", 50, "Ejemplo: recuperación de una aplicación completa."),
]

RPO_RTO_ROWS = [
    ("VM (IaaS)", "Azure Backup: 1 a 6 puntos/día (Enhanced) o diario", "Minutos desde snapshot (1-30 días); horas desde bóveda",
     "Hasta 99 años; Archive para mensual/anual", "Sí, es el caso principal", "Comparar RPO, restore instantáneo y consistencia de aplicación."),
    ("SQL Database / SQL MI", "Nativo: log cada 5-10 min (PITR)", "Minutos a horas según tamaño", "LTR nativo hasta 10 años",
     "No el PITR; a lo más el LTR", "Un tercero suele exportar (bacpac): RPO de horas, no minutos."),
    ("PostgreSQL / MySQL Flexible", "Nativo: continuo, 7-35 días", "Minutos a horas", "Sin LTR nativo (vaulted en PG)",
     "Parcial: LTR", "Validar soporte del proveedor y versión del motor."),
    ("Cosmos DB", "Nativo continuo: restauración a cualquier segundo", "Minutos a horas", "7 o 30 días",
     "No", "El backup continuo de 7 días es gratuito; no hay ahorro que capturar."),
    ("SQL Server / SAP HANA en VM", "Workload backup: log cada 15 min", "Horas", "Hasta 99 años",
     "Sí", "Hoy no hay workload backup en Azure: brecha que el tercero podría cubrir."),
    ("Blob / ADLS", "Nativo: soft delete, versioning, PITR", "Inmediato (nativo)", "Vaulted hasta 10 años",
     "Parcial: copia aislada y LTR", "Útil como copia fuera de la cuenta (ransomware)."),
    ("Azure Files", "Snapshot diario + vaulted", "Minutos (snapshot)", "Hasta 10 años", "Sí", ""),
    ("AKS", "AKS Backup (no usado hoy)", "Horas", "Según política", "Sí", "La mayoría se reconstruye por GitOps."),
    ("Key Vault, red, mensajería", "Nativo / IaC", "Reconstrucción", "No aplica", "No", "No son objeto de backup de terceros."),
]

VENDOR_QUESTIONS = [
    "¿La tarifa es por TB protegido (front-end) o por TB almacenado (back-end), y antes o después de deduplicación?",
    "¿Incluye licencia por instancia/VM, infraestructura del repositorio, egress y costo de restauración?",
    "¿Dónde vive el repositorio (Azure, otra nube, on-premise) y por qué enlace se restaura?",
    "RPO mínimo por workload (VM, SQL, PostgreSQL, Files, Blob) y si hay restore instantáneo.",
    "Tiempo de recuperación desde Archive (rehidratación) y costo por GB recuperado.",
    "Inmutabilidad, aislamiento (copia fuera del tenant) y soporte de cifrado con claves propias.",
    "Soporte de retención de 10 años y garantía de restaurabilidad de formatos antiguos.",
]


def build_third_party(workbook):
    """Cost and RPO/RTO comparison between Azure Backup and the proposed third-party backup tiers."""
    sheet = workbook.create_sheet("Comparacion_terceros", 5)
    sheet["A1"] = "Azure Backup vs alternativa de backup de terceros (D3 / Opción 1)"
    sheet["A1"].font = TITLE_FONT
    sheet["A2"] = ("Alcance comparable: lo que hoy está en bóvedas (y opcionalmente los snapshots manuales). "
                   "Los backups nativos de PaaS (PITR de SQL, Cosmos continuo, soft delete) no se reemplazan.")
    sheet["A2"].alignment = WRAP
    sheet.merge_cells("A2:F2")
    sheet.row_dimensions[2].height = 30
    sheet.append([])
    sheet.append(["Clave", "Dato de entrada", "Valor", "Fuente / nota"])
    style_header(sheet, 4, 4)
    cells = {}
    for key, label, value, note in THIRD_PARTY_INPUTS:
        sheet.append([key, label, value, note])
        row = sheet.max_row
        cells[key] = f"$C${row}"
        sheet.cell(row=row, column=3).font = BLUE
        sheet.cell(row=row, column=3).fill = INPUT_FILL
        for column in (2, 4):
            sheet.cell(row=row, column=column).alignment = WRAP
    for key in ("d3_active", "d3_archive", "o1_active", "o1_archive"):
        sheet[cells[key].replace("$", "")].number_format = '"$"#,##0.00'
    for key in ("archive_share", "daily_change", "link_efficiency"):
        sheet[cells[key].replace("$", "")].number_format = "0%"

    sheet.append([])
    sheet.append(["Clave", "Medido en Azure (oct-2026)", "Valor", "Cálculo"])
    style_header(sheet, sheet.max_row, 4)
    measured = [
        ("vault_tb", "Volumen almacenado en bóvedas (TB)",
         '=SUMIFS(Gasto_medido!$D:$D,Gasto_medido!$A:$A,"Backup",Gasto_medido!$B:$B,"*Data Stored")/1024',
         "GB-mes facturados de 'Data Stored' en la categoría Backup / 1024"),
        ("vault_cost", "Costo mensual de las bóvedas (USD)", "=Costo_calibrado!$E$2+Costo_calibrado!$F$2+Costo_calibrado!$G$2",
         "Almacenamiento + instancias protegidas + otros, por bóveda"),
        ("vault_opt_cost", "Costo mensual de bóvedas tras optimizar (USD)",
         '=MAX(0,{vault_cost}-SUMIFS(Oportunidades!$E:$E,Oportunidades!$A:$A,"O1")-SUMIFS(Oportunidades!$E:$E,Oportunidades!$A:$A,"O2"))',
         "Menos huérfanos (O1) y retención GFS (O2)"),
        ("vault_opt_tb", "Volumen tras optimizar (TB)", "=IF({vault_cost}=0,0,{vault_tb}*{vault_opt_cost}/{vault_cost})",
         "Proporcional al costo optimizado"),
        ("snap_tb", "Snapshots manuales de disco (TB)", '=SUMIFS(Gasto_medido!$D:$D,Gasto_medido!$B:$B,"*Snapshots*")/1024',
         "GB-mes facturados de snapshots / 1024"),
        ("snap_cost", "Costo mensual de snapshots (USD)", '=SUMIFS(Gasto_medido!$C:$C,Gasto_medido!$B:$B,"*Snapshots*")', ""),
        ("base_tb", "Volumen a respaldar con el tercero (TB)", "={vault_opt_tb}+{include_snapshots}*{snap_tb}", ""),
        ("azure_base", "Costo Azure comparable optimizado (USD/mes)", "={vault_opt_cost}+{include_snapshots}*{snap_cost}", ""),
        ("egress", "Salida de datos mensual del tercero (USD/mes)", "={base_tb}*1024*{daily_change}*30*{egress_gb}",
         "Incrementales diarios transferidos fuera de Azure"),
    ]
    for key, label, formula, note in measured:
        sheet.append([key, label, None, note])
        cells[key] = f"$C${sheet.max_row}"
    for key, label, formula, note in measured:
        cell = cells[key].replace("$", "")
        sheet[cell] = formula.format(**cells)
        sheet[cell].number_format = '#,##0.0' if key.endswith("_tb") else USD
        sheet.cell(row=sheet[cell].row, column=2).alignment = WRAP

    sheet.append([])
    sheet.append(["Escenario", "TB", "Costo mensual (USD)", "Costo anual (USD)", "Diferencia mensual vs Azure optimizado",
                  "USD por TB-mes", "Fracción Archive de equilibrio"])
    style_header(sheet, sheet.max_row, 7)
    scenario_start = sheet.max_row + 1
    scenarios = [
        ("Azure Backup actual", "={vault_tb}+{include_snapshots}*{snap_tb}", "={vault_cost}+{include_snapshots}*{snap_cost}", None),
        ("Azure Backup optimizado", "={base_tb}", "={azure_base}", None),
        ("D3 todo en Active", "={base_tb}", "={base_tb}*{d3_active}+{egress}", None),
        ("D3 con fracción en Archive", "={base_tb}", "={base_tb}*((1-{archive_share})*{d3_active}+{archive_share}*{d3_archive})+{egress}",
         "=IF({d3_active}={d3_archive},\"-\",MAX(0,MIN(1,({base_tb}*{d3_active}+{egress}-{azure_base})/MAX(0.0001,{base_tb}*({d3_active}-{d3_archive})))))"),
        ("Opción 1 todo en Active", "={base_tb}", "={base_tb}*{o1_active}+{egress}", None),
        ("Opción 1 con fracción en Archive", "={base_tb}", "={base_tb}*((1-{archive_share})*{o1_active}+{archive_share}*{o1_archive})+{egress}",
         "=IF({o1_active}={o1_archive},\"-\",MAX(0,MIN(1,({base_tb}*{o1_active}+{egress}-{azure_base})/MAX(0.0001,{base_tb}*({o1_active}-{o1_archive})))))"),
    ]
    for name, tb, cost, breakeven in scenarios:
        sheet.append([name])
        row = sheet.max_row
        sheet.cell(row=row, column=2, value=tb.format(**cells)).number_format = "#,##0.0"
        sheet.cell(row=row, column=3, value=cost.format(**cells)).number_format = USD
        sheet.cell(row=row, column=4, value=f"=C{row}*12").number_format = USD
        sheet.cell(row=row, column=5, value=f"=C{row}-{cells['azure_base']}").number_format = '"$"#,##0;[Red]-"$"#,##0;"-"'
        sheet.cell(row=row, column=6, value=f"=IF(B{row}=0,0,C{row}/B{row})").number_format = '"$"#,##0.00'
        if breakeven:
            sheet.cell(row=row, column=7, value=breakeven.format(**cells)).number_format = "0%"
    cells["scenario_start"] = scenario_start
    sheet.cell(row=sheet.max_row + 1, column=1, value=(
        "Fracción Archive de equilibrio: porcentaje del volumen que tendría que ir a Archive para igualar el costo de "
        "Azure optimizado. 100% indica que no alcanza el equilibrio."))
    sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
    sheet.merge_cells(start_row=sheet.max_row, start_column=1, end_row=sheet.max_row, end_column=7)
    sheet.row_dimensions[sheet.max_row].height = 30

    sheet.append([])
    sheet.append(["Restauración desde el repositorio del tercero", "TB", "Horas estimadas", "Referencia Azure Backup"])
    style_header(sheet, sheet.max_row, 4)
    for label, key, reference in (
            ("Restauración puntual", "restore_tb_small", "Minutos si el punto está en el snapshot tier (instant restore)"),
            ("Restauración masiva", "restore_tb_large", "Horas: restauración paralela dentro de la región, sin salir de Azure")):
        sheet.append([label, f"={cells[key]}", None, reference])
        row = sheet.max_row
        sheet.cell(row=row, column=3, value=f"=B{row}*1024*8/({cells['link_gbps']}*{cells['link_efficiency']}*3600)").number_format = "0.0"
        sheet.cell(row=row, column=4).alignment = WRAP
    sheet.append(["Desde Archive se suma la rehidratación (horas) en ambas opciones."])

    sheet.append([])
    sheet.append(["Workload", "RPO hoy (nativo / Azure Backup)", "RTO hoy", "LTR hoy", "¿Lo reemplaza un tercero?", "Nota"])
    style_header(sheet, sheet.max_row, 6)
    for values in RPO_RTO_ROWS:
        sheet.append(list(values))
        for column in range(1, 7):
            sheet.cell(row=sheet.max_row, column=column).alignment = WRAP

    sheet.append([])
    sheet.append(["Preguntas para el proveedor antes de decidir"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    for question in VENDOR_QUESTIONS:
        sheet.append([question])
        sheet.merge_cells(start_row=sheet.max_row, start_column=1, end_row=sheet.max_row, end_column=6)
    set_widths(sheet, [34, 40, 18, 40, 22, 40, 16])
    return cells


SOURCE_ROWS = [
    # (indicator label in Resumen or None, indicator, type, source, how it is obtained, output file, how to check)
    ("VMs inventariadas", "VMs inventariadas", "Medido", "Azure Resource Graph",
     "resources | where type =~ 'microsoft.compute/virtualmachines'", "compute_vm_inventory.csv",
     "Resource Graph Explorer con la misma consulta, o la vista Virtual machines"),
    ("  con backup", "VMs PROD con backup", "Medido", "Azure Resource Graph",
     "recoveryservicesresources | where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'"
     " (properties.sourceResourceId = id de la VM)", "compute_vm_inventory.csv",
     "Backup center > Backup instances, o bóveda > Backup items"),
    ("VMs de plataforma excluidas (Databricks, ORCA, ARO, AKS)", "Clasificación de plataforma", "Derivado", "Azure Resource Graph",
     "Tags (vendor, orca, kubernetes.io_cluster) y managedBy del resource group (resourcecontainers)", "compute_vm_inventory.csv",
     "Resource group > Properties > Managed by"),
    ("Ítems de backup huérfanos", "Ítems de backup huérfanos", "Derivado", "Azure Resource Graph",
     "Ítems protegidos cuyo sourceResourceId no existe en resources, o sin punto de recuperación en 30 días",
     "compute_orphan_backup_items.csv", "Backup center > Backup instances; abrir el ítem y el recurso origen"),
    ("Bóvedas (Recovery Services + Backup vault)", "Bóvedas y su configuración", "Medido", "Azure Resource Graph",
     "resources | where type in~ ('microsoft.recoveryservices/vaults','microsoft.dataprotection/backupvaults') "
     "(redundancySettings, securitySettings)", "compute_backup_vaults.csv", "Bóveda > Properties (Backup configuration, Security settings)"),
    (None, "Políticas y retención", "Medido", "Azure Resource Graph",
     "recoveryservicesresources | where type in~ ('microsoft.recoveryservices/vaults/backuppolicies',"
     " 'microsoft.dataprotection/backupvaults/backuppolicies')", "compute_backup_policies.csv", "Bóveda > Backup policies"),
    ("Bases de datos / servidores", "Bases de datos y su backup nativo", "Medido", "Azure Resource Graph + ARM (lectura)",
     "Resource Graph para SQL, Cosmos, PostgreSQL, MySQL; ARM GET .../backupShortTermRetentionPolicies/default y "
     ".../backupLongTermRetentionPolicies/default para retención SQL", "database_inventory.csv",
     "SQL server > Backups > Retention policies; Cosmos > Backup & Restore; servidor flexible > Backup and restore"),
    ("Unidades inventariadas (cuentas, shares, discos)", "Storage y protección nativa", "Medido", "Azure Resource Graph + ARM (lectura)",
     "Resource Graph para cuentas, discos y snapshots; ARM GET .../blobServices/default, .../fileServices/default y "
     ".../fileServices/default/shares", "storage_inventory.csv", "Storage account > Data protection; Disks; Snapshots"),
    ("Recursos con protección nativa o reconstruibles", "Protección nativa y locks", "Medido", "Azure Resource Graph + ARM (lectura)",
     "Resource Graph por servicio; ARM GET /subscriptions/{sub}/providers/Microsoft.Authorization/locks",
     "native_protection_inventory.csv", "Recurso > Locks; Key Vault > Properties; Redis > Data persistence"),
    (None, "Versiones de SO, Gen1, claves de disco, motores", "Medido", "Azure Resource Graph",
     "properties.extended.instanceView (osName, osVersion, hyperVGeneration), storageProfile.osDisk.managedDisk.diskEncryptionSet, "
     "properties.version", "restore_vm_versions.csv, restore_engine_versions.csv", "VM > Overview (Operating system, VM generation)"),
    ("Gasto cloud total oct-2026 (USD)", "Gasto cloud total", "Medido (costo real facturado)", "Cost Management Query API",
     "POST /subscriptions/{sub}/providers/Microsoft.CostManagement/query, type ActualCost, oct-2026, "
     "agrupado por ResourceType + MeterCategory + Meter", "cost_by_meter.csv", CA + ", agrupar por Medidor"),
    ("Gasto relacionado a backup oct-2026 (USD)", "Gasto relacionado a backup", "Medido (costo real facturado)",
     "Cost Management Query API", "Suma de medidores de bóvedas, snapshots y backup de bases de datos (columna "
     "'Relacionado a backup' de Gasto_medido)", "cost_by_meter.csv", CA + ", filtro Medidor contiene 'Backup' o 'Snapshots'"),
    ("  de ello en bóvedas (Recovery Services / Backup vault)", "Costo por bóveda", "Medido (costo real facturado)",
     "Cost Management Query API", "Agrupado por ResourceId + Meter, filtro ResourceType = microsoft.recoveryservices/vaults, "
     "microsoft.dataprotection/backupvaults", "cost_by_vault.csv", CA + ", Tipo de recurso = Recovery Services vaults, agrupar por Recurso"),
    ("  de ello en snapshots de disco", "Costo de snapshots", "Medido (costo real facturado)", "Cost Management Query API",
     "Medidores 'ZRS Snapshots' y 'LRS Snapshots'", "cost_by_meter.csv", CA + ", medidor contiene 'Snapshots'"),
    (None, "Tarifa efectiva por instancia protegida y por GB GRS", "Derivado de costo real", "Cost Management Query API",
     "Costo / cantidad facturada de 'Azure VM Protected Instances' y 'GRS Data Stored' (incluye descuentos del contrato)",
     "Gasto_medido", CA + ", ver Costo y Cantidad por medidor"),
    (None, "Precio de disco por tier", "Derivado de costo real", "Cost Management Query API",
     "Costo / cantidad del medidor '<tier> LRS Disk' (por ejemplo 'P30 LRS Disk')", "Gasto_medido", CA + ", medidor '<tier> LRS Disk'"),
    (None, "Costo por cuenta de storage", "Medido (costo real facturado)", "Cost Management Query API",
     "Agrupado por ResourceId + Meter, filtro ResourceType = microsoft.storage/storageaccounts, en las suscripciones con más gasto de storage",
     "cost_by_storage_account.csv", CA + ", Tipo de recurso = Storage accounts, agrupar por Recurso y Medidor"),
    (None, "Diferencias de precio Redis y App Service Plan, disco sin tier", "Supuesto", "Precio lista referencial",
     "Valores en la hoja Supuestos (celdas amarillas)", "Supuestos", "Azure Pricing Calculator o hoja de precios del contrato"),
    (None, "% de reducción por retención, % eliminable de snapshots, % geo sin requisito", "Supuesto", "Criterio de arquitectura",
     "Valores en la hoja Supuestos (celdas amarillas)", "Supuestos", "Validar con dueños de aplicación y Cumplimiento"),
    (None, "Tarifas D3 / Opción 1", "Dato del proveedor", "Reunión de revisión de lineamientos de backup",
     "Valores en la hoja Comparacion_terceros; unidad supuesta USD por TB-mes", "Comparacion_terceros",
     "Confirmar unidad y alcance con el proveedor"),
]


def build_sources(workbook):
    sheet = workbook.create_sheet("Fuentes", 1)
    sheet["A1"] = "De dónde sale cada cifra"
    sheet["A1"].font = TITLE_FONT
    sheet["A2"] = ("Medido = dato leído de Azure en la corrida. Derivado = cálculo sobre datos medidos. Supuesto = valor "
                   "editable a validar. El costo viene de la API de Cost Management (ActualCost): es el mismo dato que muestra "
                   "Cost analysis en el portal, con los descuentos del contrato; no es una calculadora de precios.")
    sheet["A2"].alignment = WRAP
    sheet.merge_cells("A2:G2")
    sheet.row_dimensions[2].height = 45
    headers = ["Indicador", "Valor", "Tipo de dato", "Fuente", "Cómo se obtiene", "Archivo de salida", "Cómo verificarlo"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    for summary_label, label, kind, source, how, output_file, check in SOURCE_ROWS:
        value = f'=IFERROR(INDEX(Resumen!$C:$C,MATCH("{summary_label}",Resumen!$B:$B,0)),"")' if summary_label else ""
        sheet.append([label, value, kind, source, how, output_file, check])
        row = sheet.max_row
        for column in range(1, 8):
            sheet.cell(row=row, column=column).alignment = WRAP
            sheet.cell(row=row, column=column).font = NORMAL
        if kind == "Supuesto":
            sheet.cell(row=row, column=3).fill = INPUT_FILL
        if "USD" in str(summary_label):
            sheet.cell(row=row, column=2).number_format = USD
    set_widths(sheet, [36, 14, 18, 24, 60, 28, 50])


def build_queries(workbook, results):
    """Exact queries and calls recorded by each script run (evidence block of *_summary.json)."""
    sheet = workbook.create_sheet("Consultas")
    sheet["A1"] = "Consultas ejecutadas (evidencia registrada por los scripts)"
    sheet["A1"].font = TITLE_FONT
    headers = ["Script", "Fecha de la corrida", "Tipo", "Etiqueta / endpoint", "Consulta, cuerpo o api-version",
               "Filas / llamadas", "Errores"]
    sheet.append([])
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    missing = []
    for folder, summary_name in (("backup-compute", "compute_summary.json"), ("backup-database", "database_summary.json"),
                                 ("backup-storage", "storage_summary.json"), ("backup-native", "native_summary.json"),
                                 ("backup-restore", "restore_summary.json"), ("backup-cost", "cost_summary.json")):
        path = os.path.join(results, folder, summary_name)
        if not os.path.exists(path):
            missing.append(summary_name)
            continue
        with open(path, encoding="utf-8") as handle:
            summary = json.load(handle)
        evidence = summary.get("evidence") or (summary.get("run") or {}).get("evidence")
        when = summary.get("generated_at") or (summary.get("run") or {}).get("generated_at", "")
        if not evidence:
            missing.append(summary_name)
            continue
        for query in evidence.get("resource_graph_queries", []):
            sheet.append([summary_name, when, "Azure Resource Graph", query["label"], query["query"], query["rows"], ""])
        for call in evidence.get("arm_calls", []):
            kind = "Cost Management" if "costmanagement" in call["pattern"].lower() else "ARM (lectura)"
            sheet.append([summary_name, when, kind, f"{call['method']} {call['pattern']}", f"api-version={call['api_version']}",
                          call["calls"], call["errors"]])
        for body in evidence.get("request_bodies", []):
            sheet.append([summary_name, when, "Cuerpo de consulta", f"POST {body['endpoint']}",
                          json.dumps(body["body"], ensure_ascii=False), "", ""])
    if missing:
        sheet.append([])
        sheet.append([("Sin evidencia registrada en: " + ", ".join(missing) + ". Esas corridas son anteriores al registro de "
                       "consultas; al volver a ejecutar los scripts esta hoja se completa con las consultas exactas.")])
        sheet.merge_cells(start_row=sheet.max_row, start_column=1, end_row=sheet.max_row, end_column=7)
        sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
        sheet.row_dimensions[sheet.max_row].height = 30
    for row in sheet.iter_rows(min_row=4):
        for cell in row:
            cell.alignment = WRAP
            cell.font = NORMAL
    set_widths(sheet, [24, 20, 20, 50, 90, 12, 10])


def build_categories(workbook):
    """Spend by meter category with formulas over Gasto_medido."""
    sheet = workbook.create_sheet("Gasto_por_categoria")
    sheet.append(["Categoría del medidor", "Costo oct-2026 (USD)", "% del total", "De ello relacionado a backup (USD)"])
    style_header(sheet, 1, 4)
    categories = sorted({row[0] for row in workbook["Gasto_medido"].iter_rows(min_row=2, values_only=True)
                         if row[0] and row[2] is not None and not str(row[0]).startswith("Fuente")},
                        key=lambda name: -sum(r[2] or 0 for r in workbook["Gasto_medido"].iter_rows(min_row=2, values_only=True)
                                              if r[0] == name and isinstance(r[2], (int, float))))
    for category in categories:
        sheet.append([category])
        row = sheet.max_row
        sheet.cell(row=row, column=2, value=f'=SUMIFS(Gasto_medido!$C:$C,Gasto_medido!$A:$A,A{row})').number_format = USD
        sheet.cell(row=row, column=3, value=f"=IF(SUM($B$2:$B$1000)=0,0,B{row}/SUM($B$2:$B$1000))").number_format = "0.0%"
        sheet.cell(row=row, column=4, value=f'=SUMIFS(Gasto_medido!$C:$C,Gasto_medido!$A:$A,A{row},Gasto_medido!$F:$F,"SI")').number_format = USD
    set_widths(sheet, [34, 18, 12, 24])


def build_summary(workbook, details, metrics):
    sheet = workbook.create_sheet("Resumen", 0)
    sheet["A1"] = "Backup en Azure: situación actual e insights"
    sheet["A1"].font = Font(name=FONT, bold=True, size=16)
    sheet["A2"] = (f"Fecha: {metrics['generated']}. Alcance: {metrics['subscriptions']} suscripciones visibles con la cuenta "
                   "de reporte (Reader). Fuente: Azure Resource Graph, ARM (lectura) y Cost Management.")
    sheet["A2"].alignment = WRAP
    sheet.merge_cells("A2:D2")
    sheet.row_dimensions[2].height = 32
    sheet.append([])
    sheet.append(["Frente", "Indicador", "Valor", "Nota"])
    style_header(sheet, 4, 4)
    vm, db, k8s = details.get("D_VMs"), details.get("D_BD"), details.get("D_K8s")
    native, orphans, vaults = details.get("D_Nativo"), details.get("D_Huerfanos"), details.get("D_Bovedas")
    storage = details.get("D_Storage")
    rows = [
        ("Compute", "VMs inventariadas", f"=COUNTA({vm.col('name')})-1" if vm else 0, "Incluye plataformas gestionadas"),
        ("Compute", "VMs de plataforma excluidas (Databricks, ORCA, ARO, AKS)",
         countifs(vm, ("backup_scope", "EXCLUDED_PLATFORM")), "No requieren backup de VM"),
        ("Compute", "VMs PROD tradicionales", countifs(vm, ("infrastructure_class", "TRADITIONAL"), ("environment", "PROD")), ""),
        ("Compute", "  con backup", countifs(vm, ("infrastructure_class", "TRADITIONAL"), ("environment", "PROD"),
                                             ("backup_enabled", "YES")), ""),
        ("Compute", "Cobertura PROD", "=IF(C7=0,0,C8/C7)", "VMs PROD tradicionales con backup activo"),
        ("Compute", "Ítems de backup huérfanos", f"=COUNTA({orphans.col('orphan_reason')})-1" if orphans else 0,
         "Origen eliminado o sin backup >30 días"),
        ("Compute", "Clusters AKS/ARO", f"=COUNTA({k8s.col('name')})-1" if k8s else 0, "Ninguno con backup; 93% con GitOps (Flux)"),
        ("Database", "Bases de datos / servidores", f"=COUNTA({db.col('name')})-1" if db else 0, ""),
        ("Database", "  con backup nativo", countifs(db, ("backup_method", "NATIVE")), "SQL, Cosmos, PostgreSQL, MySQL"),
        ("Database", "  en Recovery Services / Backup vault", "=COUNTIFS(" + db.col("backup_method") + ",\"*VAULT*\")" if db else 0,
         "Ninguna base usa bóveda"),
        ("Storage", "Unidades inventariadas (cuentas, shares, discos)",
         f"=COUNTA({storage.col('name')})-1" if storage else "Pendiente", "Corrida de storage pendiente si figura 'Pendiente'"),
        ("Nativo", "Recursos con protección nativa o reconstruibles", f"=COUNTA({native.col('name')})-1" if native else 0, ""),
        ("Nativo", "  con evidencia de IaC (Terraform)", countifs(native, ("iac_managed", "YES")), "Tags provisionedBy / managed=TFC"),
        ("Gobierno", "Bóvedas (Recovery Services + Backup vault)", f"=COUNTA({vaults.col('vault_name')})-1" if vaults else 0, ""),
        ("Gobierno", "  sin ítems", f"=COUNTIFS({vaults.col('item_count')},0)" if vaults else 0, ""),
        ("Gobierno", "  GRS", countifs(vaults, ("storage_redundancy", "GeoRedundant")), ""),
        ("Costo", "Gasto cloud total oct-2026 (USD)", "=SUM(Gasto_medido!C:C)", "Suscripciones con costo consultado"),
        ("Costo", "Gasto relacionado a backup oct-2026 (USD)", '=SUMIFS(Gasto_medido!C:C,Gasto_medido!F:F,"SI")',
         "Bóvedas, snapshots y backup de bases de datos"),
        ("Costo", "  de ello en bóvedas (Recovery Services / Backup vault)", "=SUM(Costo_calibrado!E2:G2)", "Por bóveda, hoja Costo_calibrado"),
        ("Costo", "  de ello en snapshots de disco", "=" + "Supuestos!$C$SNAP", "Medidores 'Snapshots'"),
        ("Costo", "Ahorro mensual estimado identificado (USD)", "=Oportunidades!E{total}", "Ver hoja Oportunidades"),
    ]
    for front, label, value, note in rows:
        sheet.append([front, label, value, note])
        row = sheet.max_row
        for column in range(1, 5):
            sheet.cell(row=row, column=column).font = NORMAL
            sheet.cell(row=row, column=column).alignment = WRAP
    sheet["C9"].number_format = "0.0%"
    for row in sheet.iter_rows(min_row=5):
        if row[0].value == "Costo":
            row[2].number_format = USD
    set_widths(sheet, [12, 52, 14, 60])
    return sheet


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default="./results", help="Folder with backup-<front> result folders")
    parser.add_argument("--cost-cache", default="", help="Glob of cached vault cost query JSON files (optional)")
    parser.add_argument("--versions", default="", help="Optional JSON with restore-compatibility counts")
    parser.add_argument("--account-costs", default="", help="Glob of per-storage-account cost query JSON files (optional)")
    parser.add_argument("--output-dir", default="./reports")
    args = parser.parse_args()

    folders = {"compute": "backup-compute", "database": "backup-database", "storage": "backup-storage",
               "native": "backup-native"}
    details = {}
    account_costs = load_account_costs(args.account_costs,
                                       os.path.join(args.results, "backup-cost", "cost_by_storage_account.csv"))
    for sheet_name, (front, file_name, fields) in DETAIL_SHEETS.items():
        rows = read_csv(os.path.join(args.results, folders[front], file_name))
        if sheet_name == "D_Storage":
            for row in rows:
                row["monthly_cost_usd"] = account_costs.get(row.get("id", "").lower(), "")
            fields = fields[:-1] + ["monthly_cost_usd", "id"]
        if rows:
            details[sheet_name] = Detail(sheet_name, [field for field in fields if field in rows[0]], rows)
            if sheet_name == "D_Storage":
                details[sheet_name].fields += ["disk_meter", "disk_monthly_usd"]
                for row in rows:
                    row["disk_meter"] = disk_meter(row) if row.get("service_type") == "MANAGED_DISK" else ""

    vault_costs, quantities = load_vault_costs(os.path.join(args.results, "backup-cost"), args.cost_cache)
    meters, meter_subscriptions = load_meter_costs(os.path.join(args.results, "backup-cost"))
    calibration = calibration_rows(vault_costs, details["D_Bovedas"].rows if "D_Bovedas" in details else [],
                                   details["D_Huerfanos"].rows if "D_Huerfanos" in details else [])
    versions_path = args.versions or os.path.join(args.results, "backup-restore", "versions.json")
    versions = json.load(open(versions_path, encoding="utf-8")) if os.path.exists(versions_path) else {}
    summary = json.load(open(os.path.join(args.results, "backup-compute", "compute_summary.json"), encoding="utf-8"))

    metrics = {"generated": date.today().isoformat(),
               "subscriptions": summary.get("run", summary).get("subscriptions_in_scope", "")}

    workbook = Workbook()
    workbook.remove(workbook.active)
    # Excel computes every formula when the file is opened; no separate recalculation step is needed.
    workbook.calculation.fullCalcOnLoad = True
    for detail in details.values():
        write_detail(workbook, detail)
    build_calibration(workbook, calibration)
    build_meter_sheet(workbook, meters, meter_subscriptions)
    assumptions, names = build_assumptions(workbook, calibration, quantities)
    disk_row = names["disk_gb_usd"].split("$")[-1]
    if "D_Storage" in workbook.sheetnames:
        for row in workbook["D_Storage"].iter_rows(min_row=2):
            cell = row[-1]
            if isinstance(cell.value, str):
                cell.value = cell.value.replace("$C$DISKGB", f"$C${disk_row}")
    opportunities = build_opportunities(workbook, details, names)
    fill_formula_assumptions(assumptions, names, details)
    build_risks(workbook, details)
    build_long_term(workbook, details, versions)
    build_policy(workbook)
    build_third_party(workbook)
    build_sources(workbook)
    build_queries(workbook, args.results)
    build_categories(workbook)
    summary_sheet = build_summary(workbook, details, metrics)
    snap_row = names["snapshot_spend_usd"].split("$")[-1]
    for row in summary_sheet.iter_rows(min_row=5):
        if isinstance(row[2].value, str) and "$C$SNAP" in row[2].value:
            row[2].value = row[2].value.replace("$C$SNAP", f"$C${snap_row}")
    total_row = opportunities.max_row
    for row in summary_sheet.iter_rows(min_row=5):
        if isinstance(row[2].value, str) and "{total}" in row[2].value:
            row[2].value = row[2].value.replace("{total}", str(total_row))
    order = ["Resumen", "Fuentes", "Oportunidades", "Riesgos", "Restauracion_LP", "Politica_minima", "Comparacion_terceros",
             "Supuestos", "Gasto_por_categoria", "Gasto_medido",
             "Costo_calibrado", "Consultas"]
    workbook._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order else len(order))
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.font is None or cell.font.name != FONT:
                    cell.font = Font(name=FONT, bold=cell.font.bold if cell.font else False,
                                     color=cell.font.color if cell.font else None,
                                     size=cell.font.size if cell.font else 11)

    os.makedirs(args.output_dir, exist_ok=True)
    path = os.path.join(args.output_dir, "backup_insights.xlsx")
    workbook.save(path)
    with open(os.path.join(args.output_dir, "calibration.json"), "w", encoding="utf-8") as handle:
        json.dump({"calibration": calibration, "quantities": quantities}, handle, indent=2)
    print(path)


if __name__ == "__main__":
    main()
