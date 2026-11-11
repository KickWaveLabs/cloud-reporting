#!/usr/bin/env python3
"""Build the AKS vs Container Apps workbook from the aa_*.py outputs.

Reads results/aa-prices, aa-inventory, aa-metrics, aa-container-insights and aa-cost (any of them may
be missing; the sheets that need it say so) and writes aks_aca_comparison.xlsx:

  Resumen        key figures, all formulas
  Supuestos      editable inputs (yellow) with the measured reference next to each one
  Modelo         monthly cost for 1 to 30 microservices: ACA (ADT office-hours model and 24/7) and
                 AKS (list, effective with the measured discount, 1-year reservation), with a chart
  Equilibrio     break-even points and the ADT_019 thresholds tested against them
  Mapa           microservices x requests per month: ACA minus AKS (negative = ACA is cheaper)
  Medido_AKS     per cluster: nodes, monthly cost, CPU/memory used vs allocatable, workloads
  Trafico_APIM   per API: requests, office share, peaks, off-hours wake-ups
  Workloads      per workload: replicas and per-replica CPU/memory (Container Insights)
  Config_AKS     configuration dispersion across clusters
  Precios, D_*   list prices and the detail rows every formula reads
  Fuentes        where each figure comes from and how to check it
  Consultas      exact queries and calls recorded by each script

Every model figure is a formula over Supuestos and Precios, so editing a yellow cell recalculates
the curve, the break-even points and the map.
"""

import argparse
import math
import json
import os
from calendar import monthrange
from collections import Counter, defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import aa_common as aa

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FONT = Font(name=FONT, bold=True, size=14)
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
USD = '"$"#,##0;("$"#,##0);"-"'
USD2 = '"$"#,##0.00'
WRAP = Alignment(wrap_text=True, vertical="top")
# One by one up to 30, then wide steps to show there is no crossing further out.
MODEL_POINTS = list(range(1, 31)) + [50, 100, 150, 200, 300, 500]
MAX_MICROSERVICES = MODEL_POINTS[-1]
ACA_MAX_VCPU = 4
ACA_MIN_VCPU = 0.25  # smallest Consumption allocation
USE_HEADROOM = 1.5
MAP_MAX_MICROSERVICES = 30  # the heat map stays readable
REQUEST_LEVELS = [1_000_000, 5_000_000, 10_000_000, 25_000_000, 50_000_000, 95_040_000, 200_000_000]


def style_header(sheet, row, columns):
    for column in range(1, columns + 1):
        cell = sheet.cell(row=row, column=column)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def set_widths(sheet, widths):
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def title(sheet, text, note=None, span=8):
    sheet["A1"] = text
    sheet["A1"].font = TITLE_FONT
    if note:
        sheet["A2"] = note
        sheet["A2"].alignment = WRAP
        sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        sheet.row_dimensions[2].height = 45


def number(value):
    """CSV text to int/float when it is a number, so formulas can sum it."""
    if value is None or value == "":
        return value
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return value
    return int(as_float) if as_float.is_integer() and "." not in str(value) and "e" not in str(value).lower() else as_float


# ---------------------------------------------------------------------------
# Detail sheets
# ---------------------------------------------------------------------------


class Detail:
    """A detail sheet written from CSV rows; col() gives the whole-column reference of a field."""

    def __init__(self, workbook, name, rows, fields, extra=None, complete=False):
        self.name = name
        present = set(fields) if complete or not rows else set(rows[0])
        self.fields = [field for field in fields if field in present] + list(extra or [])
        self.rows = rows
        sheet = workbook.create_sheet(name)
        sheet.append(self.fields)
        style_header(sheet, 1, len(self.fields))
        for row in rows:
            sheet.append([number(row.get(field)) for field in self.fields])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(max(1, len(self.fields)))}{max(2, len(rows) + 1)}"
        self.sheet = sheet

    def col(self, field):
        letter = get_column_letter(self.fields.index(field) + 1)
        return f"{self.name}!${letter}:${letter}"

    def letter(self, field):
        return get_column_letter(self.fields.index(field) + 1)


def all_fields(rows):
    fields = []
    for row in rows:
        fields.extend(key for key in row if key not in fields)
    return fields


def load(results, folder, name):
    return aa.read_csv(os.path.join(results, folder, name))


# ---------------------------------------------------------------------------
# Supuestos
# ---------------------------------------------------------------------------

# (key, label, value or None for a formula, note). Sections are rows with key None.
ASSUMPTIONS = [
    (None, "General", None, None),
    ("region", "Región de precios", "eastus2", "Región ARM. East US 2 es la referencia; Central US está en Precios para sensibilidad."),
    ("hours_month", "Horas por mes", 730, "Convención de Azure para precios mensuales."),
    ("office_days", "Días hábiles por mes", 21.75, "261 días hábiles al año / 12. No descuenta feriados."),
    ("office_hours", "Horas de oficina por día", 12, "ADT_019: lunes a viernes de 8:00 a 20:00."),
    ("office_seconds", "Segundos de oficina por mes", None, "Fórmula: días hábiles × horas × 3600."),
    ("n_ref", "Microservicios del caso de referencia", 10, "Se usa en los umbrales de horas, réplicas y peticiones."),
    (None, "Microservicio en Container Apps (Consumption)", None, None),
    ("cpu_rep", "vCPU por réplica", 0.5, "Supuesto del modelo. ACA permite de 0.25 vCPU / 0.5 GiB a 4 vCPU / 8 GiB; el ADT_019 solo fija el máximo: más de 4 vCPU / 8 GB va a AKS."),
    ("mem_rep", "Memoria por réplica (GiB)", 1.0, "En ACA la memoria es 2 GiB por vCPU."),
    ("min_rep_office", "Réplicas mínimas en horario de oficina", 1, "ADT_019: 1 réplica de 8 a 20 en días hábiles; fuera de horario escala a cero."),
    ("req_month", "Peticiones por microservicio al mes", 1_000_000, "Referencia: mediana de peticiones mensuales por API en APIM."),
    ("office_req_share", "Fracción de peticiones en horario de oficina", 0.9, "Referencia: peticiones en oficina / total en APIM."),
    ("rps_rep", "Peticiones por segundo que atiende una réplica", 50, "Supuesto: validar con una prueba de carga de un microservicio típico."),
    ("peak_factor", "Pico horario / promedio de oficina", 3, "Referencia: pico horario / promedio por hora de oficina en APIM."),
    ("office_active_share", "Fracción del tiempo de oficina con tráfico (tarifa activa)", 0.3,
     "Referencia: intervalos de 5 minutos con peticiones / intervalos de oficina en APIM. El resto se cobra a tarifa idle."),
    ("off_active_s", "Segundos activos fuera de horario por microservicio al mes", 36_000,
     "Referencia: intervalos de 5 minutos con tráfico fuera de horario × 300 s (APIM)."),
    ("off_idle_s", "Segundos idle fuera de horario por microservicio al mes (cooldown)", 12_000,
     "Referencia: despertares fuera de horario × cooldown de 300 s (APIM)."),
    ("env_count", "Ambientes de Container Apps por aplicación", 1, "Un ambiente por aplicación."),
    ("env_mgmt_apply", "Cobrar 'Environment Management Hour' (1 = sí)", 1,
     "Meter vigente desde 2026-09; Microsoft Learn no dice a qué ambientes aplica. En sep-oct 2026 los ambientes del "
     "banco (workload profiles sin perfil dedicado) no lo pagaron: 1 es el caso conservador."),
    ("pe_apply", "Cobrar 'Environment Private Endpoint' (1 = sí)", 1,
     "ADT_019: red privada. Sin cargos en los ambientes del banco en sep-oct 2026. Poner 0 si el ambiente usa VNet "
     "integrada sin private endpoint."),
    ("grant_apply", "Aplicar la concesión gratuita mensual (1 = sí)", 1,
     "Por suscripción y mes; se comparte entre todas las apps de la suscripción. Poner 0 si la suscripción ya la consume."),
    ("grant_vcpu_s", "Concesión gratuita: vCPU-segundos", 180_000, "Precios de Container Apps (Consumption)."),
    ("grant_mem_s", "Concesión gratuita: GiB-segundos", 360_000, "Precios de Container Apps (Consumption)."),
    ("grant_requests", "Concesión gratuita: peticiones", 2_000_000, "Precios de Container Apps (Consumption)."),
    (None, "Cluster AKS mínimo de producción", None, None),
    ("system_nodes", "Nodos del pool de sistema", 2, "Mínimo de producción: 4 nodos de 8 vCPU / 32 GB (2 sistema + 2 user)."),
    ("user_nodes_min", "Nodos mínimos del pool user", 2, ""),
    ("node_sku", "Tamaño de VM de los nodos", "Standard_D8s_v5", "Referencia: tamaño más frecuente en los node pools PROD."),
    ("node_vcpu", "vCPU por nodo", 8, ""),
    ("node_mem", "Memoria por nodo (GiB)", 32, ""),
    ("os_disk_tier", "Disco de SO (tier, o 'ephemeral')", "P10", "Managed disk de 128 GB = P10. Con disco efímero no hay costo de disco."),
    ("aks_tier_key", "Tier del cluster (clave de Precios)", "aks_standard_tier_hour",
     "aks_standard_tier_hour (Standard, SLA) o aks_lts_hour (Premium/LTS). Free = cualquier otro texto."),
    ("lb_count", "Load balancers Standard", 1, "El de salida del cluster."),
    ("ip_count", "IPs públicas Standard", 1, "IP de salida. 0 si la salida es por firewall (outboundType userDefinedRouting)."),
    ("defender_apply", "Defender for Containers por vCore (1 = sí)", 1,
     "Solo cuenta si está habilitado en AKS; en Container Apps no cobra por vCore."),
    ("alloc_cpu_share", "Fracción de CPU asignable por nodo", 0.85, "Kube-reserved; comparar con vCPU asignables en Medido_AKS."),
    ("alloc_mem_share", "Fracción de memoria asignable por nodo", 0.8, "Kube-reserved y umbral de desalojo."),
    ("daemon_cpu", "vCPU por nodo en daemonsets (monitoreo, CSI, CNI)", 0.3, "Supuesto."),
    ("daemon_mem", "GiB por nodo en daemonsets", 0.75, "Supuesto."),
    ("sidecar_cpu", "vCPU de request del sidecar de Istio por réplica", 0.1, "Request por defecto del sidecar del add-on de Istio."),
    ("sidecar_mem", "GiB de request del sidecar de Istio por réplica", 0.125, "Request por defecto del sidecar del add-on de Istio."),
    ("aks_min_rep", "Réplicas mínimas por microservicio en AKS (24/7)", 2, "Alta disponibilidad en producción."),
    ("discount", "Costo efectivo / precio lista de los nodos", None,
     "Fórmula: costo amortizado de VMs de nodos PROD / (horas-nodo medidas × precio lista), en los meses completos "
     "de métricas y costo. 1 si no hay costo medido."),
    (None, "Calculados (no editar)", None, None),
    ("p_vcpu_active", "ACA: USD por vCPU-segundo activo", None, "Precios."),
    ("p_vcpu_idle", "ACA: USD por vCPU-segundo idle", None, "Precios."),
    ("p_mem_active", "ACA: USD por GiB-segundo activo", None, "Precios."),
    ("p_mem_idle", "ACA: USD por GiB-segundo idle", None, "Precios."),
    ("p_requests", "ACA: USD por millón de peticiones", None, "Precios."),
    ("rate_active", "ACA: USD por segundo de réplica activa", None, "vCPU × tarifa + GiB × tarifa."),
    ("rate_idle", "ACA: USD por segundo de réplica idle", None, "vCPU × tarifa + GiB × tarifa."),
    ("aca_env_month", "ACA: costo fijo del ambiente al mes", None, "Gestión del ambiente + private endpoint."),
    ("aca_grant", "ACA: valor de la concesión gratuita al mes", None, "A tarifa activa."),
    ("p_vm_hour", "AKS: USD por hora de nodo (lista)", None, "Precios, Linux pay-as-you-go."),
    ("p_vm_1y_hour", "AKS: USD por hora de nodo (reserva 1 año)", None, "Precios, reserva 1 año / 8760 h."),
    ("disk_month", "AKS: disco de SO por nodo al mes", None, "Precios."),
    ("defender_node_month", "AKS: Defender por nodo al mes", None, "vCPU × USD por vCore-hora × horas."),
    ("aks_fixed_month", "AKS: costo fijo del cluster al mes", None, "Tier + load balancer + IPs públicas."),
    ("user_cpu_node", "AKS: vCPU disponibles por nodo user", None, "vCPU × fracción asignable − daemonsets."),
    ("user_mem_node", "AKS: GiB disponibles por nodo user", None, "GiB × fracción asignable − daemonsets."),
]


def price(key_expr, column="H"):
    """SUMIFS lookup in Precios by key (text or expression) and the selected region."""
    return f'SUMIFS(Precios!${column}:${column},Precios!$A:$A,{key_expr},Precios!$B:$B,{{region}})'


def build_assumptions(workbook, refs, apis_sheet, workloads_sheet, cost, usage=None, discount_months=()):
    sheet = workbook.create_sheet("Supuestos")
    title(sheet, "Supuestos (celdas amarillas: editables)",
          "Columna C: valor que usa el modelo. Columna E: referencia medida en el banco cuando existe; copiarla a C "
          "si se quiere modelar con el dato real. Precios en USD de lista; la columna 'discount' trae el descuento real.", 5)
    sheet.append([])
    sheet.append(["Clave", "Supuesto", "Valor", "Fuente / nota", "Referencia medida"])
    style_header(sheet, 4, 5)
    for key, label, value, note in ASSUMPTIONS:
        if key is None:
            sheet.append(["", label])
            for column in range(1, 6):
                sheet.cell(row=sheet.max_row, column=column).fill = SECTION_FILL
            sheet.cell(row=sheet.max_row, column=2).font = BOLD
            continue
        sheet.append([key, label, value, note])
        row = sheet.max_row
        refs[key] = f"Supuestos!$C${row}"
        if value is not None:
            sheet.cell(row=row, column=3).font = BLUE
            sheet.cell(row=row, column=3).fill = INPUT_FILL
        for column in (2, 4):
            sheet.cell(row=row, column=column).alignment = WRAP
    r = {key: value.split("!")[1] for key, value in refs.items()}

    def set_formula(key, formula, fmt=None):
        cell = sheet[r[key].replace("$", "")]
        cell.value = "=" + formula.format(**r).replace("{region}", r["region"])
        if fmt:
            cell.number_format = fmt

    set_formula("office_seconds", "{office_days}*{office_hours}*3600", "#,##0")
    for key, meter in (("p_vcpu_active", "aca_vcpu_active_s"), ("p_vcpu_idle", "aca_vcpu_idle_s"),
                       ("p_mem_active", "aca_mem_active_gib_s"), ("p_mem_idle", "aca_mem_idle_gib_s"),
                       ("p_requests", "aca_requests_million")):
        set_formula(key, price(f'"{meter}"'), "0.000000")
    set_formula("rate_active", "{cpu_rep}*{p_vcpu_active}+{mem_rep}*{p_mem_active}", "0.0000000")
    set_formula("rate_idle", "{cpu_rep}*{p_vcpu_idle}+{mem_rep}*{p_mem_idle}", "0.0000000")
    set_formula("aca_env_month", "{env_count}*{hours_month}*({env_mgmt_apply}*" + price('"aca_env_management_hour"')
                + "+{pe_apply}*" + price('"aca_env_private_endpoint_hour"') + ")", USD2)
    set_formula("aca_grant", "{grant_apply}*({grant_vcpu_s}*{p_vcpu_active}+{grant_mem_s}*{p_mem_active}"
                "+{grant_requests}/1000000*{p_requests})", USD2)
    set_formula("p_vm_hour", price('"vm_"&{node_sku}&"_payg"'), "0.0000")
    set_formula("p_vm_1y_hour", price('"vm_"&{node_sku}&"_1y"', "I"), "0.0000")
    set_formula("disk_month", 'IF(LOWER({os_disk_tier})="ephemeral",0,' + price('"disk_"&{os_disk_tier}&"_month"') + ")", USD2)
    set_formula("defender_node_month", "{defender_apply}*{node_vcpu}*{hours_month}*"
                + price('"defender_containers_vcore_hour"'), USD2)
    set_formula("aks_fixed_month", "{hours_month}*(" + price("{aks_tier_key}") + "+{lb_count}*"
                + price('"lb_standard_rules_hour"') + "+{ip_count}*" + price('"public_ip_standard_hour"') + ")", USD2)
    set_formula("user_cpu_node", "{node_vcpu}*{alloc_cpu_share}-{daemon_cpu}", "0.00")
    set_formula("user_mem_node", "{node_mem}*{alloc_mem_share}-{daemon_mem}", "0.00")
    if cost and usage and discount_months:
        # Billed VM hours are not usable here: savings-plan rows report about twice the real node-hours.
        months = "{{" + ",".join(f'"{month}"' for month in discount_months) + "}}"  # set_formula runs str.format
        set_formula("discount", f'IFERROR(SUM(SUMIFS({cost.col("cost_usd")},{cost.col("meter_category")},"Virtual Machines",'
                    f'{cost.col("owner_type")},"AKS",{cost.col("environment")},"PROD",{cost.col("month")},{months}))/'
                    f'SUM(SUMIFS({usage.col("list_cost_usd")},{usage.col("environment")},"PROD",'
                    f'{usage.col("month")},{months})),1)', "0.00")
    else:
        sheet[r["discount"].replace("$", "")] = 1
    sheet[r["discount"].replace("$", "")].fill = INPUT_FILL

    # Measured references (column E).
    def reference(key, formula, fmt=None):
        cell = sheet[r[key].replace("$", "").replace("C", "E", 1)]
        cell.value = formula
        if fmt:
            cell.number_format = fmt

    if apis_sheet:
        reference("req_month", f"=IFERROR(MEDIAN(Trafico_APIM!$D$5:$D${apis_sheet}),\"\")", "#,##0")
        reference("office_req_share", f"=IFERROR(SUM(Trafico_APIM!$E$5:$E${apis_sheet})/SUM(Trafico_APIM!$D$5:$D${apis_sheet}),\"\")", "0%")
        reference("peak_factor", f"=IFERROR(MEDIAN(Trafico_APIM!$J$5:$J${apis_sheet}),\"\")", "0.0")
        reference("office_active_share", f"=IFERROR(MEDIAN(Trafico_APIM!$K$5:$K${apis_sheet}),\"\")", "0%")
        reference("off_active_s", f"=IFERROR(MEDIAN(Trafico_APIM!$L$5:$L${apis_sheet}),\"\")", "#,##0")
        reference("off_idle_s", f"=IFERROR(MEDIAN(Trafico_APIM!$M$5:$M${apis_sheet}),\"\")", "#,##0")
    if workloads_sheet:
        reference("cpu_rep", f"=IFERROR(MEDIAN(Workloads!$G$5:$G${workloads_sheet}),\"\")", "0.00")
        reference("mem_rep", f"=IFERROR(MEDIAN(Workloads!$I$5:$I${workloads_sheet}),\"\")", "0.00")
    set_widths(sheet, [20, 58, 18, 90, 18])
    return r


# ---------------------------------------------------------------------------
# Model formulas
# ---------------------------------------------------------------------------


def aca_formula(r, n, req, warm="ADT", rps_rep=None):
    """ACA monthly cost for n microservices with req requests each (Excel expressions)."""
    os_ = f"Supuestos!{r['office_seconds']}"
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    rps = f"({req}*{s['office_req_share']}/{os_})"
    replicas = f"MAX({s['min_rep_office']},CEILING({rps}/{rps_rep or s['rps_rep']},1))"
    extra = f"({replicas}-{s['min_rep_office']})*{os_}*{s['rate_active']}"
    if warm == "ADT":
        warm_cost = (f"{s['min_rep_office']}*{os_}*({s['office_active_share']}*{s['rate_active']}"
                     f"+(1-{s['office_active_share']})*{s['rate_idle']})")
        off = f"{s['off_active_s']}*{s['rate_active']}+{s['off_idle_s']}*{s['rate_idle']}"
    else:  # warm 24/7: the minimum replica never scales to zero
        warm_cost = (f"{s['min_rep_office']}*{s['hours_month']}*3600*({s['office_active_share']}*{s['rate_active']}"
                     f"+(1-{s['office_active_share']})*{s['rate_idle']})")
        off = "0"
    per_ms = f"({warm_cost}+{extra}+{off}+{req}/1000000*{s['p_requests']})"
    return f"MAX(0,{n}*{per_ms}-{s['aca_grant']})+{s['aca_env_month']}"


def aks_nodes_formula(r, n, req, rps_rep=None):
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    rps = f"({req}*{s['office_req_share']}/{s['office_seconds']})"
    peak = f"MAX({s['aks_min_rep']},CEILING({rps}*{s['peak_factor']}/{rps_rep or s['rps_rep']},1))"
    return (f"MAX({s['user_nodes_min']},CEILING({n}*{peak}*({s['cpu_rep']}+{s['sidecar_cpu']})/{s['user_cpu_node']},1),"
            f"CEILING({n}*{peak}*({s['mem_rep']}+{s['sidecar_mem']})/{s['user_mem_node']},1))")


def aks_formula(r, nodes, mode="list"):
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    vm = {"list": s["p_vm_hour"], "effective": f"{s['p_vm_hour']}*{s['discount']}", "1y": s["p_vm_1y_hour"]}[mode]
    return (f"({s['system_nodes']}+{nodes})*({vm}*{s['hours_month']}+{s['disk_month']}+{s['defender_node_month']})"
            f"+{s['aks_fixed_month']}")


def build_model(workbook, r):
    sheet = workbook.create_sheet("Modelo")
    title(sheet, "Costo mensual por aplicación: Container Apps vs AKS (USD, precios de lista salvo 'efectivo')",
          "ACA horario ADT: 1 réplica de 8 a 20 en días hábiles y escala a cero fuera de horario. ACA 24/7: la réplica "
          "mínima nunca baja a cero. AKS: cluster dedicado a la aplicación (como hoy), 2 nodos de sistema + nodos user "
          "que crecen si las réplicas no caben. Efectivo: precio de nodo × descuento medido.", 10)
    sheet.append([])
    headers = ["Microservicios", "ACA horario ADT", "ACA 24/7", "AKS lista", "AKS efectivo", "AKS reserva 1 año",
               "Nodos user AKS", "ACA ADT − AKS lista", "Más barato (lista)", "Más barato (efectivo)"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    req = f"Supuestos!{r['req_month']}"
    for n in MODEL_POINTS:
        row = sheet.max_row + 1
        sheet.cell(row=row, column=1, value=n)
        sheet.cell(row=row, column=2, value="=" + aca_formula(r, f"$A{row}", req))
        sheet.cell(row=row, column=3, value="=" + aca_formula(r, f"$A{row}", req, warm="24x7"))
        sheet.cell(row=row, column=7, value="=" + aks_nodes_formula(r, f"$A{row}", req))
        sheet.cell(row=row, column=4, value="=" + aks_formula(r, f"$G{row}", "list"))
        sheet.cell(row=row, column=5, value="=" + aks_formula(r, f"$G{row}", "effective"))
        sheet.cell(row=row, column=6, value="=" + aks_formula(r, f"$G{row}", "1y"))
        sheet.cell(row=row, column=8, value=f"=B{row}-D{row}")
        sheet.cell(row=row, column=9, value=f'=IF(B{row}<=D{row},"ACA","AKS")')
        sheet.cell(row=row, column=10, value=f'=IF(B{row}<=E{row},"ACA","AKS")')
        for column in range(2, 7):
            sheet.cell(row=row, column=column).number_format = USD
        sheet.cell(row=row, column=8).number_format = '"$"#,##0;[Red]-"$"#,##0'
    last = sheet.max_row
    chart = LineChart()
    chart.title = "Costo mensual por aplicación (USD)"
    chart.y_axis.title = "USD / mes"
    chart.x_axis.title = "Microservicios"
    chart.height, chart.width = 9, 18
    chart.add_data(Reference(sheet, min_col=2, max_col=5, min_row=4, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=5, max_row=last))
    sheet.add_chart(chart, "L4")
    set_widths(sheet, [14, 16, 14, 14, 14, 16, 12, 18, 16, 18])
    sheet.freeze_panes = "B5"
    return {"first": 5, "last": last}


def build_breakeven(workbook, r, model, map_info):
    sheet = workbook.create_sheet("Equilibrio")
    title(sheet, "Punto de equilibrio y umbrales del ADT_019",
          "Primer número de microservicios en que Container Apps cuesta más que el cluster AKS dedicado, con los "
          "supuestos de la hoja Supuestos. Los umbrales de horas y réplicas usan el caso de referencia (n_ref).", 5)
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    first, last = model["first"], model["last"]

    def first_n(aca_col, aks_col):
        return (f'=IFERROR(INDEX(Modelo!$A${first}:$A${last},MATCH(1,INDEX((Modelo!${aca_col}${first}:${aca_col}${last}'
                f'>Modelo!${aks_col}${first}:${aks_col}${last})*1,0),0)),">{MAX_MICROSERVICES}")')

    n_ref = s["n_ref"]
    req = s["req_month"]
    aks_ref = aks_formula(r, aks_nodes_formula(r, n_ref, req), "list")
    aks_ref_eff = aks_formula(r, aks_nodes_formula(r, n_ref, req), "effective")
    warm_rate = f"({s['office_active_share']}*{s['rate_active']}+(1-{s['office_active_share']})*{s['rate_idle']})"
    rps = f"({req}*{s['office_req_share']}/{s['office_seconds']})"
    replicas = f"MAX({s['min_rep_office']},CEILING({rps}/{s['rps_rep']},1))"
    not_warm = (f"(({replicas}-{s['min_rep_office']})*{s['office_seconds']}*{s['rate_active']}"
                f"+{s['off_active_s']}*{s['rate_active']}+{s['off_idle_s']}*{s['rate_idle']}+{req}/1000000*{s['p_requests']})")
    base = f"({s['off_active_s']}*{s['rate_active']}+{s['off_idle_s']}*{s['rate_idle']}+{req}/1000000*{s['p_requests']})"
    budget = f"(({aks_ref}-{s['aca_env_month']}+{s['aca_grant']})/{n_ref})"
    rows = [
        ("Microservicios en que ACA (horario ADT) supera a AKS lista", first_n("B", "D"), "0"),
        ("Microservicios en que ACA (horario ADT) supera a AKS efectivo", first_n("B", "E"), "0"),
        ("Microservicios en que ACA (horario ADT) supera a AKS con reserva 1 año", first_n("B", "F"), "0"),
        ("Microservicios en que ACA 24/7 supera a AKS lista", first_n("C", "D"), "0"),
        ("Costo ACA por microservicio al mes (horario ADT)", f"=Modelo!B{first + 1}-Modelo!B{first}", USD2),
        ("Costo del cluster AKS mínimo al mes (lista)", f"=Modelo!D{first}", USD),
        ("Costo del cluster AKS mínimo al mes (efectivo)", f"=Modelo!E{first}", USD),
        ("Horas de réplica caliente por día hábil en que ACA iguala a AKS lista (n_ref)",
         f"=IFERROR(({budget}-{not_warm})/({s['min_rep_office']}*{s['office_days']}*3600*{warm_rate}),\"\")", "0.0"),
        ("Réplicas calientes promedio por microservicio en que ACA iguala a AKS lista (n_ref, horario de oficina)",
         f"=IFERROR(({budget}-{base})/({s['office_seconds']}*{warm_rate}),\"\")", "0.0"),
        ("Costo AKS lista del caso de referencia", "=" + aks_ref, USD),
        ("Costo AKS efectivo del caso de referencia", "=" + aks_ref_eff, USD),
    ]
    sheet.append([])
    sheet.append(["Indicador", "Valor"])
    style_header(sheet, 4, 2)
    cells = {}
    for label, formula, fmt in rows:
        sheet.append([label, formula])
        sheet.cell(row=sheet.max_row, column=2).number_format = fmt
        sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
        cells[label] = f"Equilibrio!$B${sheet.max_row}"

    sheet.append([])
    sheet.append(["Regla del ADT_019", "Umbral ADT", "Resultado del modelo", "Lectura"])
    style_header(sheet, sheet.max_row, 4)
    n_star = cells[rows[0][0]]
    hours_star = cells[rows[7][0]]
    tps_cell = f"Mapa!${get_column_letter(map_info['tps_col'])}${map_info['row_for_15']}"
    adt = [
        ("Más de 15 microservicios con uso constante → AKS", 15, f"={n_star}",
         f'=IF(ISNUMBER({n_star}),IF({n_star}>15,"El costo respalda el umbral: ACA sigue siendo más barato hasta "&{n_star}-1,'
         f'"ACA cuesta más que AKS desde "&{n_star}&" microservicios, antes de 15"),"ACA más barato en todo el rango")'),
        ("Más de 10 h al día de uso constante → AKS", 10, f"={hours_star}",
         f'=IF({hours_star}="","",IF({hours_star}>24,"Con n_ref, ACA es más barato aunque la réplica esté caliente todo el día",'
         f'IF({hours_star}>10,"El umbral de 10 h es conservador: el cruce está en "&TEXT({hours_star},"0.0")&" h",'
         f'"ACA cuesta más que AKS antes de 10 h: "&TEXT({hours_star},"0.0")&" h")))'),
        ("Encendido 24/7 cuesta más que AKS", "", f"={cells[rows[3][0]]}",
         f'=IF(ISNUMBER({cells[rows[3][0]]}),"Con réplicas calientes 24/7, ACA supera a AKS lista desde "&{cells[rows[3][0]]}'
         f'&" microservicios","Aun encendido 24/7, ACA es más barato hasta {MAX_MICROSERVICES} microservicios")'),
        ("Más de 100 tps de carga total → AKS", 100, f"={tps_cell}",
         f'=IF({tps_cell}<0,"Con 15 microservicios y 100 tps en oficina, ACA aún es más barato","Con 15 microservicios y 100 tps, '
         f'AKS es más barato")'),
    ]
    for label, threshold, result, reading in adt:
        sheet.append([label, threshold, result, reading])
        row = sheet.max_row
        for column in (1, 4):
            sheet.cell(row=row, column=column).alignment = WRAP
    sheet.cell(row=sheet.max_row, column=3).number_format = '"$"#,##0;[Red]-"$"#,##0'
    sheet.append([])
    sheet.append(["Otras reglas del ADT_019 que no son de costo y descartan ACA aunque sea más barato: réplica de más de "
                  "4 vCPU / 8 GiB, procesos batch o ETL, consumidores que no toleran cold start (hasta ~30 s), ambiente "
                  "dedicado (requiere aprobación de arquitectura)."])
    sheet.merge_cells(start_row=sheet.max_row, start_column=1, end_row=sheet.max_row, end_column=4)
    sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
    sheet.row_dimensions[sheet.max_row].height = 45
    set_widths(sheet, [62, 16, 22, 70])
    return cells


def build_map(workbook, r):
    sheet = workbook.create_sheet("Mapa")
    title(sheet, "ACA (horario ADT) − AKS lista, USD/mes: negativo = ACA más barato",
          "Filas: microservicios de la aplicación. Columnas: peticiones totales de la aplicación al mes (editables), "
          "repartidas por igual entre microservicios. 95.040.000 ≈ 100 tps sostenidos en horario de oficina.", 9)
    sheet.append([])
    sheet.append(["Microservicios \\ peticiones/mes"] + REQUEST_LEVELS)
    style_header(sheet, 4, len(REQUEST_LEVELS) + 1)
    for column in range(2, len(REQUEST_LEVELS) + 2):
        cell = sheet.cell(row=4, column=column)
        cell.number_format = "#,##0"
        cell.fill = INPUT_FILL
        cell.font = Font(name=FONT, bold=True, color="0000FF")
    row_for_15 = None
    for n in range(1, MAP_MAX_MICROSERVICES + 1):
        row = sheet.max_row + 1
        sheet.cell(row=row, column=1, value=n)
        if n == 15:
            row_for_15 = row
        for column in range(2, len(REQUEST_LEVELS) + 2):
            letter = get_column_letter(column)
            req = f"({letter}$4/$A{row})"
            formula = aca_formula(r, f"$A{row}", req) + "-(" + aks_formula(r, aks_nodes_formula(r, f"$A{row}", req)) + ")"
            sheet.cell(row=row, column=column, value="=" + formula).number_format = '"$"#,##0;[Red]-"$"#,##0'
    area = f"B5:{get_column_letter(len(REQUEST_LEVELS) + 1)}{sheet.max_row}"
    sheet.conditional_formatting.add(area, CellIsRule(operator="lessThan", formula=["0"], fill=GREEN_FILL))
    sheet.conditional_formatting.add(area, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=RED_FILL))
    set_widths(sheet, [30] + [14] * len(REQUEST_LEVELS))
    sheet.freeze_panes = "B5"
    return {"tps_col": REQUEST_LEVELS.index(95_040_000) + 2, "row_for_15": row_for_15}


# ---------------------------------------------------------------------------
# Measured sheets
# ---------------------------------------------------------------------------


def build_measured(workbook, clusters, usage, cost, workloads, services, months):
    sheet = workbook.create_sheet("Medido_AKS")
    title(sheet, "AKS medido: costo, capacidad y uso por cluster",
          "Costo: promedio mensual de Cost Management (AmortizedCost) del cluster + su grupo de recursos de nodos. Uso: "
          "promedio de horario de oficina (lunes a viernes 8-20, Lima) de métricas de plataforma. Cada cluster atiende "
          "una sola aplicación.", 13)
    sheet.append([])
    headers = ["Cluster", "Ambiente", "Nodos sistema", "Nodos user", "Tamaños de VM", "Costo mensual (USD)",
               "vCPU asignables", "vCPU usados oficina", "Uso CPU oficina", "vCPU pico", "GiB usados oficina",
               "Workloads (sin sistema)", "Services", "Malla", "Id"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    cost_months = max(1, months.get("cost", 1))
    usage_months = max(1, months.get("usage", 1))
    for row_data in sorted(clusters.rows, key=lambda row: (row["environment"] != "PROD", row["name"])):
        name = row_data["name"]
        sheet.append([name, row_data["environment"], number(row_data["system_nodes"]), number(row_data["user_nodes"]),
                      row_data["vm_sizes"]])
        row = sheet.max_row
        if cost:
            # By resource id: two subscriptions can hold clusters with the same name.
            sheet.cell(row=row, column=6, value=f'=SUMIFS({cost.col("cost_usd")},{cost.col("owner_id")},O{row},'
                                                f'{cost.col("owner_type")},"AKS")/{cost_months}').number_format = USD
        if usage:
            sheet.cell(row=row, column=7, value=f'=SUMIFS({usage.col("cpu_allocatable_avg")},{usage.col("cluster")},A{row})'
                                                f'/{usage_months}').number_format = "0.0"
            sheet.cell(row=row, column=8, value=f'=SUMIFS({usage.col("cpu_used_office_avg")},{usage.col("cluster")},A{row})'
                                                f'/{usage_months}').number_format = "0.0"
            sheet.cell(row=row, column=9, value=f"=IFERROR(H{row}/G{row},\"\")").number_format = "0%"
            sheet.cell(row=row, column=10, value=f'=_xlfn.MAXIFS({usage.col("cpu_used_max")},{usage.col("cluster")},A{row})'
                                                 ).number_format = "0.0"
            sheet.cell(row=row, column=11, value=f'=SUMIFS({usage.col("mem_used_office_avg_gib")},{usage.col("cluster")},'
                                                 f'A{row})/{usage_months}').number_format = "0.0"
        if workloads:
            sheet.cell(row=row, column=12, value=f'=COUNTIFS(Workloads!$A:$A,A{row})')
        if services:
            sheet.cell(row=row, column=13, value=f'=COUNTIFS({services.col("cluster")},A{row})')
        sheet.cell(row=row, column=14, value=row_data.get("service_mesh", ""))
        sheet.cell(row=row, column=15, value=row_data.get("id", ""))
    last = sheet.max_row
    sheet.append([])
    sheet.append(["Resumen PROD"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    summary = [
        ("Clusters PROD", f'=COUNTIFS($B$5:$B${last},"PROD")', "0"),
        ("Clusters PROD con 4 nodos o menos", f'=SUMPRODUCT(($B$5:$B${last}="PROD")*(($C$5:$C${last}+$D$5:$D${last})<=4))', "0"),
        ("Costo mensual promedio por cluster PROD (USD)", f'=IFERROR(AVERAGEIFS($F$5:$F${last},$B$5:$B${last},"PROD"),"")', USD),
        ("Costo mensual total AKS PROD (USD)", f'=SUMIFS($F$5:$F${last},$B$5:$B${last},"PROD")', USD),
        ("Uso de CPU promedio en oficina, PROD", f'=IFERROR(SUMIFS($H$5:$H${last},$B$5:$B${last},"PROD")/'
                                                  f'SUMIFS($G$5:$G${last},$B$5:$B${last},"PROD"),"")', "0%"),
        ("vCPU asignables PROD sin usar en oficina", f'=SUMIFS($G$5:$G${last},$B$5:$B${last},"PROD")-'
                                                    f'SUMIFS($H$5:$H${last},$B$5:$B${last},"PROD")', "0"),
    ]
    if usage:
        prod = f'{usage.col("environment")},"PROD"'
        def share(used, allocatable):
            return f'=IFERROR(SUMIFS({usage.col(used)},{prod})/SUMIFS({usage.col(allocatable)},{prod}),"")'

        # p95 and maximum are per pool and month; summed over clusters they are an upper bound,
        # because the peaks of different clusters do not happen at the same hour.
        summary += [
            ("Uso de memoria promedio en oficina, PROD", share("mem_used_office_avg_gib", "mem_allocatable_avg_gib"), "0%"),
            ("Uso de CPU percentil 95 por hora, PROD", share("cpu_used_p95", "cpu_allocatable_avg"), "0%"),
            ("Uso de CPU máximo (pico de cada cluster sumado), PROD", share("cpu_used_max", "cpu_allocatable_avg"), "0%"),
            ("Uso de memoria percentil 95 por hora, PROD", share("mem_used_p95_gib", "mem_allocatable_avg_gib"), "0%"),
            ("Uso de memoria máximo (pico de cada cluster sumado), PROD", share("mem_used_max_gib", "mem_allocatable_avg_gib"),
             "0%"),
        ]
    refs = {}
    for label, formula, fmt in summary:
        sheet.append([label, formula])
        sheet.cell(row=sheet.max_row, column=2).number_format = fmt
        refs[label] = f"Medido_AKS!$B${sheet.max_row}"
    if cost:
        sheet.append([])
        sheet.append(["Container Apps existentes (costo mensual promedio por ambiente)"])
        sheet.cell(row=sheet.max_row, column=1).font = BOLD
        for env_name in sorted({row["owner"] for row in cost.rows if row["owner_type"] == "ACA"}):
            sheet.append([env_name, f'=SUMIFS({cost.col("cost_usd")},{cost.col("owner")},A{sheet.max_row + 1},'
                                    f'{cost.col("owner_type")},"ACA")/{cost_months}'])
            sheet.cell(row=sheet.max_row, column=2).number_format = USD
    set_widths(sheet, [34, 12, 10, 10, 22, 16, 12, 14, 12, 10, 12, 14, 10, 10, 40])
    sheet.freeze_panes = "B5"
    return refs, last


def build_traffic(workbook, traffic, apis, months):
    sheet = workbook.create_sheet("Trafico_APIM")
    title(sheet, "Tráfico por API en APIM (promedio mensual)",
          "Modelo north-south: toda llamada, incluso entre microservicios, pasa por APIM, así que estas peticiones son "
          "las que Container Apps cobraría. Fuera de horario: cada racha de intervalos de 5 minutos con tráfico despierta "
          "una réplica, que sigue encendida 300 s (cooldown) después de la última petición.", 13)
    sheet.append([])
    headers = ["APIM", "API", "Backend", "Peticiones/mes", "En oficina/mes", "Fuera de horario/mes", "Pico 5 min",
               "Pico hora", "Promedio por hora de oficina", "Pico / promedio", "Fracción oficina con tráfico",
               "Seg. activos fuera de horario/mes", "Seg. idle (cooldown)/mes", "Despertares/mes"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    n = max(1, months)
    backends = {(row["apim"], row["api_id"]): row.get("backend_host", "") for row in (apis or [])}
    keys = sorted({(row["apim"], row["api_id"]) for row in traffic.rows})
    t = traffic

    def sumifs(field, row):
        return f'SUMIFS({t.col(field)},{t.col("apim")},$A{row},{t.col("api_id")},$B{row})'

    for apim, api in keys:
        sheet.append([apim, api, backends.get((apim, api), "")])
        row = sheet.max_row
        sheet.cell(row=row, column=4, value=f"={sumifs('requests', row)}/{n}").number_format = "#,##0"
        sheet.cell(row=row, column=5, value=f"={sumifs('office_requests', row)}/{n}").number_format = "#,##0"
        sheet.cell(row=row, column=6, value=f"={sumifs('offhours_requests', row)}/{n}").number_format = "#,##0"
        sheet.cell(row=row, column=7, value=f'=_xlfn.MAXIFS({t.col("peak_5min_requests")},{t.col("apim")},$A{row},'
                                            f'{t.col("api_id")},$B{row})').number_format = "#,##0"
        sheet.cell(row=row, column=8, value=f'=_xlfn.MAXIFS({t.col("peak_hour_requests")},{t.col("apim")},$A{row},'
                                            f'{t.col("api_id")},$B{row})').number_format = "#,##0"
        sheet.cell(row=row, column=9, value=f"=IFERROR({sumifs('office_requests', row)}/{sumifs('office_hours_in_window', row)},0)"
                   ).number_format = "#,##0"
        sheet.cell(row=row, column=10, value=f"=IFERROR(H{row}/I{row},\"\")").number_format = "0.0"
        sheet.cell(row=row, column=11, value=f'=IFERROR(AVERAGEIFS({t.col("office_active_share")},{t.col("apim")},$A{row},'
                                             f'{t.col("api_id")},$B{row}),"")').number_format = "0%"
        sheet.cell(row=row, column=12, value=f"={sumifs('offhours_active_seconds', row)}/{n}").number_format = "#,##0"
        sheet.cell(row=row, column=13, value=f"={sumifs('offhours_idle_seconds', row)}/{n}").number_format = "#,##0"
        sheet.cell(row=row, column=14, value=f"={sumifs('offhours_wakeups', row)}/{n}").number_format = "#,##0"
    set_widths(sheet, [18, 30, 30, 14, 14, 14, 10, 10, 14, 10, 12, 16, 14, 12])
    sheet.freeze_panes = "C5"
    return sheet.max_row if keys else None


def build_workloads(workbook, workloads):
    sheet = workbook.create_sheet("Workloads")
    title(sheet, "Workloads por cluster (Container Insights, horario de oficina, sin namespaces de sistema)",
          "Requests y uso por réplica: promedio de los meses medidos. 'Cabe en ACA' aplica el límite del ADT_019 "
          "(4 vCPU / 8 GiB por réplica).", 11)
    sheet.append([])
    headers = ["Cluster", "Ambiente", "Namespace", "Workload", "Tipo", "Réplicas promedio", "vCPU request por réplica",
               "vCPU usados por réplica", "GiB request por réplica", "GiB usados por réplica", "Cabe en ACA"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    w = workloads
    keys = sorted({(row["cluster"], row["environment"], row["namespace"], row["workload"], row["kind"])
                   for row in w.rows if row.get("system_namespace") == "NO" and row.get("period") == "OFFICE"})

    def avg(field, row):
        return (f'=IFERROR(AVERAGEIFS({w.col(field)},{w.col("cluster")},$A{row},{w.col("namespace")},$C{row},'
                f'{w.col("workload")},$D{row},{w.col("period")},"OFFICE"),"")')

    for key in keys:
        sheet.append(list(key))
        row = sheet.max_row
        for column, field, fmt in ((6, "replicas_avg", "0.0"), (7, "cpu_request_per_replica", "0.00"),
                                   (8, "cpu_used_per_replica", "0.00"), (9, "mem_request_gib_per_replica", "0.00"),
                                   (10, "mem_used_gib_per_replica", "0.00")):
            if field in w.fields:
                sheet.cell(row=row, column=column, value=avg(field, row)).number_format = fmt
        sheet.cell(row=row, column=11, value=f'=IF(OR(N(G{row})>4,N(I{row})>8),"NO","SÍ")')
    set_widths(sheet, [30, 10, 24, 34, 12, 12, 14, 14, 14, 14, 10])
    sheet.freeze_panes = "E5"
    return sheet.max_row if keys else None


CONFIG_PROPERTIES = [
    "kubernetes_version", "minor_versions_behind", "sku_tier", "private_cluster", "authorized_ip_ranges",
    "disable_local_accounts", "entra_integration", "azure_rbac", "network_plugin", "network_plugin_mode",
    "network_policy", "outbound_type", "upgrade_channel", "node_os_upgrade_channel", "defender", "workload_identity",
    "oidc_issuer", "image_cleaner", "disk_encryption_set", "encryption_at_host", "azure_policy", "key_vault_csi",
    "container_insights", "managed_prometheus", "service_mesh", "keda", "gitops_flux", "zones",
]


# Flags that Resource Graph omits while the feature is off, so an empty value means NO and is not drift.
OFF_WHEN_EMPTY = ("private_cluster", "disable_local_accounts", "defender", "workload_identity", "oidc_issuer",
                  "image_cleaner", "azure_policy", "key_vault_csi", "container_insights", "managed_prometheus", "keda")


def normalize_booleans(rows):
    """'True'/'False' from Resource Graph become SÍ/NO, so COUNTIFS compares text and not Excel booleans."""
    for row in rows:
        for key, value in row.items():
            if value in ("True", "true"):
                row[key] = "SÍ"
            elif value in ("False", "false") or (value == "" and key in OFF_WHEN_EMPTY):
                row[key] = "NO"


def add_version_lag(rows):
    def minor(version):
        parts = str(version or "").split(".")
        return (int(parts[0]), int(parts[1])) if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit() else None

    newest = max((minor(row.get("kubernetes_version")) for row in rows if minor(row.get("kubernetes_version"))),
                 default=None)
    for row in rows:
        current = minor(row.get("kubernetes_version"))
        row["minor_versions_behind"] = (newest[1] - current[1]) if newest and current and newest[0] == current[0] else ""


def build_config(workbook, clusters):
    sheet = workbook.create_sheet("Config_AKS")
    title(sheet, "Dispersión de configuración entre clusters AKS",
          "Para cada propiedad: el valor más frecuente en PROD y cuántos clusters PROD tienen otro valor. No juzga "
          "cumplimiento: el equipo de lineamientos marca qué valor exige la LBS (columna F, editable) y la columna G "
          "cuenta los que no lo cumplen.", 7)
    sheet.append([])
    headers = ["Propiedad", "Valores distintos (todos)", "Valor más frecuente PROD", "Clusters PROD con otro valor",
               "Clusters PROD", "Valor exigido por LBS", "Clusters PROD que no cumplen"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    c = clusters
    prod = [row for row in c.rows if row["environment"] == "PROD"]
    for prop in CONFIG_PROPERTIES:
        if prop not in c.fields:
            continue
        values = Counter(str(row.get(prop, "")) for row in prod)
        mode = values.most_common(1)[0][0] if values else ""
        distinct = Counter(str(row.get(prop, "")) for row in c.rows)
        sheet.append([prop, "; ".join(f"{value or '(vacío)'}: {count}" for value, count in distinct.most_common()), mode])
        row = sheet.max_row
        criteria = f'"<>"&C{row}' if mode != "" else '"<>"'
        sheet.cell(row=row, column=4, value=f'=COUNTIFS({c.col("environment")},"PROD",{c.col(prop)},{criteria})')
        sheet.cell(row=row, column=5, value=f'=COUNTIFS({c.col("environment")},"PROD")')
        sheet.cell(row=row, column=6).fill = INPUT_FILL
        sheet.cell(row=row, column=7, value=f'=IF(F{row}="","",COUNTIFS({c.col("environment")},"PROD",'
                                            f'{c.col(prop)},"<>"&F{row}))')
        sheet.cell(row=row, column=2).alignment = WRAP
    set_widths(sheet, [26, 60, 22, 14, 10, 18, 14])


HPA_FINDINGS = [
    ("HPA que no puede escalar (mín. = máx.)", "fixed_min_equals_max", "SÍ", "Mínimo y máximo iguales en el export"),
    ("HPA con mínimo de 2 o más réplicas", "min_two_or_more", "SÍ",
     "Capacidad que se paga 24/7; en Container Apps fuera de horario el mínimo es 0"),
    ("HPA al máximo el día de la foto", "at_max_now", "SÍ", "Réplicas actuales = máximo del HPA"),
    ("HPA que se queda en su mínimo en los meses medidos", "behaviour_ci", "NUNCA_ESCALA",
     "Promedio de réplicas en oficina (Container Insights) ≤ 110% del mínimo del HPA"),
    ("HPA cerca de su máximo en los meses medidos", "behaviour_ci", "LLEGA_AL_MAX",
     "Promedio de réplicas en oficina (Container Insights) ≥ 90% del máximo del HPA"),
    ("HPA sin historia en Container Insights", "ci_matched", "NO",
     "Cluster sin el add-on o workload sin pods en los meses medidos"),
]
DEPLOYMENT_FINDINGS = [
    ("Deployments sin HPA", "has_hpa", "NO", "Sin fila en el export de HPA"),
    ("Deployments sin pods en los meses medidos (posibles zombis)", "no_usage_in_ci", "SÍ",
     "El cluster tiene Container Insights pero el workload no aparece"),
    ("Request de CPU más de 3 veces el uso real", "request_over_3x_used", "SÍ",
     "Request del export / vCPU usados por réplica en horario de oficina"),
    ("Uso real de CPU por encima del request", "used_over_request", "SÍ",
     "El scheduler reserva menos de lo que usa: riesgo de contención en el nodo"),
    ("Límite de CPU 10 veces el request o más", "limit_10x_request", "SÍ", "Límite / request del export"),
    ("Cabe en una réplica de Container Apps según su request", "fits_aca_by_request", "SÍ",
     "Request ≤ 4 vCPU y ≤ 8 GiB, máximo por réplica del ADT_019"),
]


def build_k8s(workbook, hpa, deployments, taken_on):
    sheet = workbook.create_sheet("Config_K8s")
    title(sheet, f"HPA y deployments PROD: configuración al {taken_on} frente a los meses medidos",
          "La configuración es una foto de un día (export del equipo de plataforma, sin historia). La historia de "
          "réplicas y uso sale de Container Insights. Detalle por fila en D_HPA y D_Deploy.", 5)
    sheet.append([])
    headers = ["Hallazgo", "PROD", "Total PROD", "%", "Cómo se calcula"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    for detail, findings in ((hpa, HPA_FINDINGS), (deployments, DEPLOYMENT_FINDINGS)):
        if detail is None:
            continue
        for label, field, value, how in findings:
            if field not in detail.fields:
                continue
            sheet.append([label])
            row = sheet.max_row
            sheet.cell(row=row, column=2, value=f'=COUNTIFS({detail.col("environment")},"PROD",{detail.col(field)},"{value}")')
            sheet.cell(row=row, column=3, value=f'=COUNTIFS({detail.col("environment")},"PROD")')
            sheet.cell(row=row, column=4, value=f'=IFERROR(B{row}/C{row},"")').number_format = "0%"
            sheet.cell(row=row, column=5, value=how).alignment = WRAP
        sheet.append([])
    set_widths(sheet, [56, 10, 12, 8, 70])


MICROSERVICE_BUCKETS = [("0 (sin deployments en el export)", 0, 0), ("1", 1, 1), ("2 a 3", 2, 3), ("4 a 5", 4, 5),
                        ("6 a 10", 6, 10), ("11 a 15", 11, 15), ("16 a 30", 16, 30), ("31 a 100", 31, 100),
                        ("Más de 100", 101, 10 ** 9)]
REPLICA_BUCKETS = [("0 (posible zombi)", 0, 0), ("1", 1, 1), ("2 a 3", 2, 3), ("4 a 5", 4, 5), ("6 a 10", 6, 10),
                   ("11 a 20", 11, 20), ("21 a 50", 21, 50), ("Más de 50", 51, 10 ** 9)]
LAYER_ORDER = ["business", "channel", "app", "async", "support", "core", "otros"]
BEHAVIOURS = [("NUNCA_ESCALA", "Se queda en su mínimo"), ("ESCALA", "Escala entre mínimo y máximo"),
              ("LLEGA_AL_MAX", "Llega o se acerca a su máximo"), ("FIJO", "Mínimo = máximo"), ("SIN_DATO", "Sin datos")]


def section(sheet, text, span=5):
    sheet.append([])
    sheet.append([text])
    row = sheet.max_row
    for column in range(1, span + 1):
        sheet.cell(row=row, column=column).fill = SECTION_FILL
    sheet.cell(row=row, column=1).font = BOLD


def header(sheet, values):
    sheet.append(values)
    style_header(sheet, sheet.max_row, len(values))


def build_patterns(workbook, deploy, hpa, kclusters, capabilities, taken_on):
    sheet = workbook.create_sheet("Patrones")
    title(sheet, f"Patrones de uso de AKS en PROD (configuración al {taken_on})",
          "Microservicio = deployment fuera de namespaces de plataforma. Dominio = namespace. Capa = primera palabra "
          "del nombre. Réplicas y HPA del export de un día; la historia de 3 meses sale de Container Insights.", 5)
    prod_ms = [deploy.col("environment"), '"PROD"', deploy.col("platform_component"), '"NO"']

    section(sheet, "Patrón 1. Granularidad de microservicios")
    header(sheet, ["Dato", "Valor", "% del total", "", "Cómo se calcula"])
    sheet.append(["Microservicios PROD", f"=COUNTIFS({','.join(prod_ms)})", "", "", "Filas de D_Deploy en PROD sin plataforma"])
    total_row = sheet.max_row
    for layer in LAYER_ORDER:
        sheet.append([f"Capa {layer}-*" if layer != "otros" else "Otras capas"])
        row = sheet.max_row
        sheet.cell(row=row, column=2, value=f'=COUNTIFS({",".join(prod_ms)},{deploy.col("layer")},"{layer}")')
        sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/$B${total_row},"")').number_format = "0%"
    if capabilities:
        sheet.append([])
        header(sheet, ["Capacidad repetida en varios dominios", "Dominios PROD", "Deployments PROD", "",
                       "Dominios (deployments)"])
        for row in capabilities:
            sheet.append([row["capability"], number(row["namespaces_prod"]), number(row["deployments_prod"]), "",
                          row["namespaces"]])
            sheet.cell(row=sheet.max_row, column=5).alignment = WRAP

    section(sheet, "Patrón 2. Clusters por número de microservicios")
    header(sheet, ["Microservicios por cluster", "Clusters PROD", "% de clusters", "Costo mensual (USD)", "Cómo se calcula"])
    k = kclusters
    base = [k.col("environment"), '"PROD"', k.col("in_inventory"), '"SÍ"']
    first = sheet.max_row + 1
    for label, low, high in MICROSERVICE_BUCKETS:
        criteria = ",".join(base + [k.col("microservices"), f'">={low}"', k.col("microservices"), f'"<={high}"'])
        sheet.append([label, f"=COUNTIFS({criteria})"])
        row = sheet.max_row
        sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${first + len(MICROSERVICE_BUCKETS) - 1}),"")'
                   ).number_format = "0%"
        sheet.cell(row=row, column=4, value=f"=SUMIFS({k.col('cost_month_usd')},{criteria})").number_format = USD
    last = sheet.max_row
    sheet.append(["Clusters PROD con 15 microservicios o menos",
                  f'=COUNTIFS({",".join(base + [k.col("microservices"), chr(34) + "<=15" + chr(34)])})',
                  f'=IFERROR(B{last + 1}/SUM($B${first}:$B${last}),"")', "", "Umbral de 15 microservicios del ADT_019"])
    sheet.cell(row=sheet.max_row, column=3).number_format = "0%"
    sheet.append(["Costo mensual de clusters PROD con 3 microservicios o menos", "", "",
                  f'=SUMIFS({k.col("cost_month_usd")},{",".join(base + [k.col("microservices"), chr(34) + "<=3" + chr(34)])})',
                  "Costo amortizado real (Cost Management), último mes completo del periodo"])
    sheet.cell(row=sheet.max_row, column=4).number_format = USD

    if hpa:
        section(sheet, "Patrón 4. Réplicas y HPA")
        header(sheet, ["Réplicas actuales (día de la foto)", "HPA PROD", "% de HPA", "", "Cómo se calcula"])
        prod_hpa = [hpa.col("environment"), '"PROD"', hpa.col("platform_component"), '"NO"']
        first = sheet.max_row + 1
        for label, low, high in REPLICA_BUCKETS:
            criteria = ",".join(prod_hpa + [hpa.col("replicas_now"), f'">={low}"', hpa.col("replicas_now"), f'"<={high}"'])
            sheet.append([label, f"=COUNTIFS({criteria})"])
            row = sheet.max_row
            sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${first + len(REPLICA_BUCKETS) - 1}),"")'
                       ).number_format = "0%"
        sheet.append([])
        header(sheet, ["Comportamiento del HPA", "Según ValTrend (máximo, pocos días)", "%",
                       "Según Container Insights (promedio en oficina, 3 meses)", "%"])
        first = sheet.max_row + 1
        for code, label in BEHAVIOURS:
            sheet.append([label])
            row = sheet.max_row
            end = first + len(BEHAVIOURS) - 1
            sheet.cell(row=row, column=2, value=f'=COUNTIFS({",".join(prod_hpa)},{hpa.col("behaviour_export")},"{code}")')
            sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${end}),"")').number_format = "0%"
            sheet.cell(row=row, column=4, value=f'=COUNTIFS({",".join(prod_hpa)},{hpa.col("behaviour_ci")},"{code}")')
            sheet.cell(row=row, column=5, value=f'=IFERROR(D{row}/SUM($D${first}:$D${end}),"")').number_format = "0%"
        if "min_above_need" in hpa.fields:
            at_min = f'{",".join(prod_hpa)},{hpa.col("behaviour_ci")},"NUNCA_ESCALA"'
            sheet.append([])
            header(sheet, ["De los HPA que se quedan en su mínimo", "HPA", "% de los que se quedan en su mínimo", "",
                           "Cómo se calcula"])
            for label, field, how in (
                    ("Su mínimo supera las réplicas que pide su CPU", "min_above_need",
                     "Réplicas necesarias = CPU usada en oficina / (request por réplica × 70%)"),
                    ("Su mínimo supera eso aun dejando 2 réplicas por alta disponibilidad", "min_above_need_beyond_ha",
                     "Mínimo mayor que el máximo entre 2 y las réplicas necesarias")):
                sheet.append([label, f'=COUNTIFS({at_min},{hpa.col(field)},"SÍ")'])
                row = sheet.max_row
                sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/COUNTIFS({at_min}),"")').number_format = "0%"
                sheet.cell(row=row, column=5, value=how).alignment = WRAP
        sheet.append(["Container Insights: promedio de réplicas en oficina; en su mínimo = hasta 110% del mínimo, cerca del "
                      "máximo = desde 90% del máximo. No se usa el máximo por hora porque cada despliegue suma por unos "
                      "minutos los pods nuevos a los viejos."])
    set_widths(sheet, [52, 16, 14, 22, 70])


LEARN = "https://learn.microsoft.com/azure/"
ADT = "ADT_019 (Confluence)"
NOTE_FONT = Font(name=FONT, color="C00000", italic=True)


def build_criteria(workbook, clusters, kclusters, deploy, hpa, traffic, workloads_last, traffic_month):
    """Technical criteria to choose a platform, each with the measured PROD figure, its source and any caveat."""
    sheet = workbook.create_sheet("Criterios_tecnicos")
    title(sheet, "Criterios técnicos AKS vs Container Apps, con el dato medido en PROD",
          "Cada fila: qué favorece a cada plataforma, cuánto del parque actual cae de cada lado y de dónde sale el "
          "criterio. Las notas en rojo son puntos discutibles. Microsoft Learn revisado el 11-nov-2026.", 8)
    sheet.append([])
    header(sheet, ["Criterio", "Favorece Container Apps", "Favorece AKS", "Dato medido (PROD)", "Valor", "De un total de",
                   "Fuente", "Nota"])

    def add(criterion, aca, aks, measure, value="", total="", source="", note=""):
        sheet.append([criterion, aca, aks, measure, value, total, source, note])
        row = sheet.max_row
        for column in (1, 2, 3, 4, 7, 8):
            sheet.cell(row=row, column=column).alignment = WRAP
        sheet.cell(row=row, column=8).font = NOTE_FONT
        return row

    k_base = f'{kclusters.col("environment")},"PROD",{kclusters.col("in_inventory")},"SÍ"' if kclusters else ""
    add("Microservicios por aplicación", "15 o menos", "Más de 15, o una plataforma compartida por varias apps",
        "Clusters con 15 microservicios o menos",
        f'=COUNTIFS({k_base},{kclusters.col("microservices")},"<=15")' if kclusters else "",
        f"=COUNTIFS({k_base})" if kclusters else "", ADT,
        "El umbral de 15 es del ADT; Microsoft no publica uno.")
    if traffic and traffic_month:
        t = traffic
        month = f'{t.col("month")},"{traffic_month}"'
        add("Transacciones por segundo y patrón de tráfico", "Menos de 100 tps, o tráfico con picos y valles",
            "Más de 100 tps sostenidos", f"APIs con pico de más de 100 tps en {traffic_month}",
            f'=COUNTIFS({month},{t.col("peak_5min_requests")},">30000")', f"=COUNTIFS({month})",
            f"{ADT}; métricas de APIM (Requests)",
            "Pico por API promediado en 5 min: el pico real por segundo es algo mayor. Es por API, no por aplicación.")
        row = add("Patrón horario", "Uso concentrado en horario de oficina; escala a 0 fuera de él",
                  "Uso constante 24/7 (más de 10 h al día)", f"Peticiones fuera de horario / total, {traffic_month}",
                  f'=IFERROR(SUMIFS({t.col("offhours_requests")},{month})/SUMIFS({t.col("requests")},{month}),"")', "",
                  f"{ADT}; {LEARN}container-apps/billing")
        sheet.cell(row=row, column=5).number_format = "0%"
    if hpa:
        h = hpa
        base = f'{h.col("environment")},"PROD",{h.col("platform_component")},"NO"'
        add("Escalado y escalado a cero", "Escala a 0 y por eventos de forma nativa; hasta 1.000 réplicas por revisión",
            "Densidad de pods alta y constante; hoy en PROD no escala a 0 (ver escenario con KEDA abajo)",
            "HPA que se queda en su mínimo (promedio en oficina, 3 meses)",
            f'=COUNTIFS({base},{h.col("behaviour_ci")},"NUNCA_ESCALA")', f"=COUNTIFS({base})",
            f"{LEARN}container-apps/scale-app")
    if deploy:
        d = deploy
        base = f'{d.col("environment")},"PROD",{d.col("platform_component")},"NO"'
        add("Tamaño por réplica", "Hasta 4 vCPU / 8 GiB", "Réplicas más grandes o VMs especiales (memoria, GPU, red)",
            "Microservicios que caben por su request", f'=COUNTIFS({base},{d.col("fits_aca_by_request")},"SÍ")',
            f"=COUNTIFS({base})", f"{ADT}; {LEARN}container-apps/containers",
            "En environments solo de consumo el máximo baja a 2 vCPU / 4 GiB; los del banco usan workload profiles.")
    add("Tipo de carga y arquitectura", "HTTP, APIs, eventos y jobs cortos",
        "StatefulSet, DaemonSet propios, batch largo, contenedores privilegiados", "Workloads StatefulSet en namespaces de aplicación",
        '=COUNTIFS(Workloads!$B:$B,"PROD",Workloads!$E:$E,"StatefulSet")' if workloads_last else "",
        '=COUNTIFS(Workloads!$B:$B,"PROD")' if workloads_last else "",
        f"{LEARN}container-apps/compare-options; {LEARN}container-apps/containers",
        "Container Apps no permite contenedores privilegiados ni imágenes que no sean Linux x86-64.")
    if clusters:
        c = clusters
        base = f'{c.col("environment")},"PROD"'
        add("Acceso a la API de Kubernetes y control de la plataforma",
            "No necesita kubectl, CRDs ni operadores propios (la API está oculta)",
            "Service mesh, operadores, CRDs y políticas de red propias", "Clusters con Istio gestionado",
            f'=COUNTIFS({base},{c.col("service_mesh")},"Istio")', f"=COUNTIFS({base})",
            f"{LEARN}container-apps/compare-options; {LEARN}aks/istio-about",
            "Solo cuenta el Istio gestionado por Azure; un Istio instalado por la app no aparece.")
        add("", "", "", "Clusters con GitOps (Flux)", f'=COUNTIFS({base},{c.col("gitops_flux")},"YES")',
            f"=COUNTIFS({base})", f"{LEARN}azure-arc/kubernetes/conceptual-gitops-flux2")
        add("Curva de aprendizaje y esfuerzo operativo", "La plataforma actualiza runtime, nodos y parches",
            "Upgrades de versión y nodos a cargo de cada equipo; requiere experiencia en Kubernetes",
            "Clusters sin upgrade automático de versión", f'=COUNTIFS({base},{c.col("upgrade_channel")},"none")',
            f"=COUNTIFS({base})", f"{LEARN}aks/auto-upgrade-cluster; {LEARN}aks/istio-about",
            "El add-on de Istio no se actualiza solo: cada revisión se sube a mano. Las horas-hombre no se cuantifican.")
    add("Límites de escala", "Cores de consumo por environment (cuota ampliable a pedido); 1.000 réplicas por revisión",
        "5.000 nodos por cluster, 250 pods por nodo, 100 node pools", "Sin dato medido: límite documentado", "", "",
        f"{LEARN}container-apps/quotas; {LEARN}aks/quotas-skus-regions",
        "La página de cuotas de Container Apps (9-oct-2026) no publica un tope de 100 o 200 apps por environment.")
    add("Costo", "Pago por segundo activo", "Nodos reservados con uso alto y constante", "Ver hojas Modelo, Equilibrio y Mapa",
        "", "", "https://azure.microsoft.com/pricing/details/container-apps/; Azure Retail Prices API; Cost Management")
    sheet.append([])
    sheet.append(["Escenario aparte: AKS con KEDA (no está en PROD, no cuenta en la tabla de arriba)"])
    for column in range(1, 9):
        sheet.cell(row=sheet.max_row, column=column).fill = SECTION_FILL
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    add("Qué cambiaría", "", "Con el add-on de KEDA, AKS escala pods a 0 por eventos; los nodos mínimos del cluster se "
        "siguen pagando, así que el ahorro depende de que el node pool también baje", "", "", "",
        f"{LEARN}aks/keda-about")
    if clusters:
        c = clusters
        add("Estado en el banco", "", "Iniciativa de KEDA con autoscaler predictivo; no está en PROD porque Microsoft no "
            "puede asegurar la capacidad", "Clusters PROD con KEDA", f'=COUNTIFS({c.col("environment")},"PROD",'
            f'{c.col("keda")},"SÍ")', f'=COUNTIFS({c.col("environment")},"PROD")', "Equipo de plataforma",
            "Dato del equipo de plataforma; confirmar el estado de la iniciativa antes de presentarlo.")
    if kclusters and "bia" in kclusters.fields:
        sheet.append([])
        sheet.append(["Aparte, informativo: criticidad del negocio (BIA). No forma parte de los criterios técnicos"])
        for column in range(1, 9):
            sheet.cell(row=sheet.max_row, column=column).fill = SECTION_FILL
        sheet.cell(row=sheet.max_row, column=1).font = BOLD
        add("Criticidad (BIA)", "BIA media o baja", "BIA alta o muy alta", "Clusters con BIA alta o muy alta",
            f'=COUNTIFS({k_base},{kclusters.col("bia")},"MUY ALTA")+COUNTIFS({k_base},{kclusters.col("bia")},"ALTA")',
            f"=COUNTIFS({k_base})", "Export de deployments (bia); hoja Candidatos_BIA",
            "Solo para conocimiento; qué niveles de BIA exigen AKS lo deciden seguridad y arquitectura.")
    set_widths(sheet, [24, 34, 34, 36, 10, 12, 40, 44])


BIA_LEVELS = ["MUY ALTA", "ALTA", "MEDIA", "BAJA", "MUY BAJA", "SIN BIA"]
BIA_REQUIRES_AKS = ("MUY ALTA", "ALTA")


def build_bia(workbook, kclusters):
    sheet = workbook.create_sheet("Candidatos_BIA")
    title(sheet, "Anexo informativo: clusters PROD aptos para Container Apps, separados por criticidad (BIA)",
          "Apto técnicamente = 1 a 15 microservicios, todos caben en una réplica de Container Apps por su request "
          "(4 vCPU / 8 GiB) y sin StatefulSet de aplicación. La columna B (editable) dice qué niveles de BIA se quedan "
          "en AKS por cumplimiento. No incluye peticiones por segundo por aplicación. Detalle en D_K8s_clusters.", 8)
    sheet.append([])
    header(sheet, ["BIA", "¿Se queda en AKS por BIA?", "Clusters PROD", "Aptos técnicamente para ACA",
                   "Costo mensual actual de los aptos (USD)", "Clusters que pasarían a ACA",
                   "Microservicios que pasarían a ACA", "Aptos que se quedan en AKS por BIA"])
    k = kclusters
    base = f'{k.col("environment")},"PROD",{k.col("in_inventory")},"SÍ"'
    present = {row["bia"] for row in k.rows}
    levels = [level for level in BIA_LEVELS if level in present] + sorted(present - set(BIA_LEVELS))
    first = sheet.max_row + 1
    for level in levels:
        sheet.append([level, "SÍ" if level in BIA_REQUIRES_AKS else "NO"])
        row = sheet.max_row
        sheet.cell(row=row, column=2).fill = INPUT_FILL
        by_level = f'{base},{k.col("bia")},$A{row}'
        fit = f'{by_level},{k.col("aca_technical_fit")},"SÍ"'
        sheet.cell(row=row, column=3, value=f"=COUNTIFS({by_level})")
        sheet.cell(row=row, column=4, value=f"=COUNTIFS({fit})")
        sheet.cell(row=row, column=5, value=f"=SUMIFS({k.col('cost_month_usd')},{fit})").number_format = USD
        sheet.cell(row=row, column=6, value=f'=IF($B{row}="SÍ",0,D{row})')
        sheet.cell(row=row, column=7, value=f'=IF($B{row}="SÍ",0,SUMIFS({k.col("microservices")},{fit}))')
        sheet.cell(row=row, column=8, value=f'=IF($B{row}="SÍ",D{row},0)')
    last = sheet.max_row
    sheet.append(["Total", ""] + [f"=SUM({get_column_letter(column)}{first}:{get_column_letter(column)}{last})"
                                  for column in range(3, 9)])
    total = sheet.max_row
    sheet.cell(row=total, column=1).font = BOLD
    sheet.cell(row=total, column=5).number_format = USD
    sheet.append([])
    sheet.append(["Costo mensual actual de los clusters que pasarían a ACA (USD)", "", "", "",
                  f'=SUMPRODUCT((B{first}:B{last}<>"SÍ")*E{first}:E{last})'])
    sheet.cell(row=sheet.max_row, column=5).number_format = USD
    set_widths(sheet, [16, 16, 12, 16, 20, 16, 18, 18])


SENSITIVITY_RPS = [5, 10, 25, 50, 100]
SENSITIVITY_REQUESTS = [1_000_000, 10_000_000, 50_000_000, 200_000_000]


def build_sensitivity(workbook, r):
    """How much the peticiones-por-réplica assumption moves the result, for the reference case (n_ref)."""
    sheet = workbook["Equilibrio"]
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    n = s["n_ref"]
    sheet.append([])
    sheet.append(["Sensibilidad: peticiones por segundo que atiende una réplica (caso de referencia: n_ref microservicios)"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    sheet.append(["El valor de Supuestos es un supuesto sin prueba de carga. Columna A editable. Con pocas peticiones "
                  "casi no influye; con muchas, más réplicas encarecen ACA y también AKS."])
    for label, kind in (("ACA horario ADT (USD/mes)", "aca"), ("AKS efectivo (USD/mes)", "aks")):
        header(sheet, [label] + [f"{level / 1_000_000:g} M peticiones/mes por microservicio" for level in SENSITIVITY_REQUESTS])
        for rps in SENSITIVITY_RPS:
            sheet.append([rps])
            row = sheet.max_row
            sheet.cell(row=row, column=1).fill = INPUT_FILL
            for column, level in enumerate(SENSITIVITY_REQUESTS, 2):
                if kind == "aca":
                    formula = aca_formula(r, n, str(level), rps_rep=f"$A{row}")
                else:
                    formula = aks_formula(r, aks_nodes_formula(r, n, str(level), rps_rep=f"$A{row}"), "effective")
                sheet.cell(row=row, column=column, value="=" + formula).number_format = USD


PROFILE_CLASSES = [("OFICINA", "90% o más de las peticiones en horario de oficina"),
                   ("EXTENDIDO", "90% o más en la ventana extendida, todos los días"),
                   ("24/7", "Tráfico repartido en el día y la noche"), ("SIN TRÁFICO", "Sin peticiones en el periodo")]


def build_traffic_profile(workbook, profile, traffic):
    """Classify each API by when its traffic happens; windows and threshold are editable cells."""
    d = Detail(workbook, "D_Perfil_horario", profile, ["apim", "api_id", "day_type", "hour", "avg_requests_per_hour"])
    sheet = workbook.create_sheet("Perfil_trafico")
    title(sheet, "Propuesta: perfil de tráfico por API (oficina, horario extendido o 24/7)",
          "Peticiones promedio por hora de APIM (días hábiles y fin de semana, hora de Lima) en la ventana de métricas. "
          "Las celdas amarillas son editables: al cambiarlas se reclasifica todo sin volver a consultar Azure.", 7)
    sheet.append([])
    inputs = [("Oficina: hora de inicio (lunes a viernes)", 8), ("Oficina: hora de fin", 20),
              ("Extendido: hora de inicio (todos los días)", 6), ("Extendido: hora de fin", 22),
              ("Umbral para clasificar", 0.9)]
    cells = {}
    for label, value in inputs:
        sheet.append([label, value])
        sheet.cell(row=sheet.max_row, column=2).fill = INPUT_FILL
        cells[label] = f"$B${sheet.max_row}"
    sheet.cell(row=sheet.max_row, column=2).number_format = "0%"
    office_start, office_end, ext_start, ext_end, threshold = cells.values()
    apis = sorted({(row["apim"], row["api_id"]) for row in profile})
    sheet.append([])
    header(sheet, ["Clase", "APIs", "% de las peticiones", "", "Criterio"])
    summary_first = sheet.max_row + 1
    table_header = summary_first + len(PROFILE_CLASSES) + 1
    first, last = table_header + 1, table_header + len(apis)
    for code, meaning in PROFILE_CLASSES:
        sheet.append([code, f'=COUNTIF($F${first}:$F${last},A{sheet.max_row + 1})',
                      f'=IFERROR(SUMIF($F${first}:$F${last},A{sheet.max_row + 1},$C${first}:$C${last})/'
                      f'SUM($C${first}:$C${last}),"")', "", meaning])
        sheet.cell(row=sheet.max_row, column=3).number_format = "0%"
    sheet.append([])
    header(sheet, ["APIM", "API", "Peticiones por semana", "% en oficina", "% en ventana extendida", "Clase", "Pico tps (5 min)"])
    assert sheet.max_row == table_header
    for apim, api_id in apis:
        sheet.append([apim, api_id])
        row = sheet.max_row
        api = f'{d.col("apim")},$A{row},{d.col("api_id")},$B{row}'
        value = d.col("avg_requests_per_hour")

        def hours(day_type, start=None, end=None):
            window = f',{d.col("hour")},">="&{start},{d.col("hour")},"<"&{end}' if start else ""
            return f'SUMIFS({value},{api},{d.col("day_type")},"{day_type}"{window})'

        sheet.cell(row=row, column=3, value=f"=5*{hours('weekday')}+2*{hours('weekend')}").number_format = "#,##0"
        sheet.cell(row=row, column=4, value=f'=IFERROR(5*{hours("weekday", office_start, office_end)}/C{row},"")'
                   ).number_format = "0%"
        sheet.cell(row=row, column=5, value=f'=IFERROR((5*{hours("weekday", ext_start, ext_end)}+2*'
                                            f'{hours("weekend", ext_start, ext_end)})/C{row},"")').number_format = "0%"
        sheet.cell(row=row, column=6, value=f'=IF(N(C{row})=0,"SIN TRÁFICO",IF(D{row}>={threshold},"OFICINA",'
                                            f'IF(E{row}>={threshold},"EXTENDIDO","24/7")))')
        if traffic:
            sheet.cell(row=row, column=7, value=f'=IFERROR(_xlfn.MAXIFS({traffic.col("peak_5min_requests")},'
                                                f'{traffic.col("apim")},$A{row},{traffic.col("api_id")},$B{row})/300,"")'
                       ).number_format = "0.0"
    set_widths(sheet, [44, 40, 18, 14, 18, 14, 14])
    sheet.freeze_panes = f"C{first}"


def aca_allocation(cpu, mem_gib):
    """Smallest Consumption allocation (0.25 vCPU steps, 2 GiB per vCPU) that covers the request."""
    need = max(cpu or 0, (mem_gib or 0) / 2, ACA_MIN_VCPU)
    return math.ceil(need * 4 - 1e-9) / 4


def aca_workloads(workloads, names):
    """One row per PROD application workload: replica-hours per month (office / off-hours) and ACA size.

    Replica-hours use the full months only (the first and last month of the window are partial).
    """
    rows = [row for row in workloads if row["environment"] == "PROD" and row["system_namespace"] == "NO"
            and row["kind"] == "ReplicaSet" and row["cluster"] in names]
    months = sorted({row["month"] for row in rows})
    full = set(months[1:-1]) if len(months) >= 3 else set(months)
    grouped = defaultdict(list)
    for row in rows:
        grouped[(row["cluster"], row["namespace"], row["workload"])].append(row)
    out = []
    for (cluster, namespace, workload), items in sorted(grouped.items()):
        def replica_hours(period):
            total = sum(float(item["replicas_avg"] or 0) * float(item["hours"] or 0)
                        for item in items if item["period"] == period and item["month"] in full)
            return round(total / max(1, len(full)), 1)

        def weighted(field):
            pairs = [(float(item.get(field) or 0), float(item["hours"] or 0)) for item in items]
            hours = sum(weight for _, weight in pairs)
            return sum(value * weight for value, weight in pairs) / hours if hours else 0

        cpu_used, mem_used = weighted("cpu_used_per_replica"), weighted("mem_used_gib_per_replica")
        cpu = max(weighted("cpu_request_per_replica"), cpu_used)
        mem = max(weighted("mem_request_gib_per_replica"), mem_used)
        vcpu = aca_allocation(cpu, mem)
        # Sized by use: average use plus 50% headroom, which is what a migration would request.
        vcpu_use = min(vcpu, aca_allocation(cpu_used * USE_HEADROOM, mem_used * USE_HEADROOM))
        out.append({"cluster": cluster, "namespace": namespace, "workload": workload,
                    "office_replica_hours_month": replica_hours("OFFICE"),
                    "offhours_replica_hours_month": replica_hours("OFFHOURS"),
                    "cpu_request_or_use": round(cpu, 3), "mem_gib_request_or_use": round(mem, 3),
                    "cpu_used": round(cpu_used, 3), "mem_gib_used": round(mem_used, 3),
                    "aca_vcpu": vcpu, "aca_mem_gib": vcpu * 2, "aca_vcpu_by_use": vcpu_use,
                    "aca_mem_gib_by_use": vcpu_use * 2, "fits_aca": "SÍ" if vcpu_use <= ACA_MAX_VCPU else "NO"})
    return out, sorted(full)


def build_aca_by_cluster(workbook, r, clusters, workloads, cost, cost_month):
    names = {row["name"] for row in clusters.rows if row["environment"] == "PROD"}
    rows, full = aca_workloads(workloads, names)
    if not rows:
        return
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    d = Detail(workbook, "D_ACA_cargas", rows, list(rows[0].keys()),
               extra=["aca_same_replicas_request_usd", "aca_same_replicas_use_usd", "aca_adt_use_usd"])
    letter = {field: d.letter(field) for field in d.fields}
    for index in range(2, len(rows) + 2):
        def rate(vcpu, mem):
            return f"(${letter[vcpu]}{index}*{s['p_vcpu_active']}+${letter[mem]}{index}*{s['p_mem_active']})"

        fits = f'${letter["fits_aca"]}{index}="SÍ"'
        office = f"${letter['office_replica_hours_month']}{index}"
        all_hours = f"({office}+${letter['offhours_replica_hours_month']}{index})"
        d.sheet[f"{letter['aca_same_replicas_request_usd']}{index}"] = (
            f"=IF({fits},{all_hours}*3600*{rate('aca_vcpu', 'aca_mem_gib')},0)")
        d.sheet[f"{letter['aca_same_replicas_use_usd']}{index}"] = (
            f"=IF({fits},{all_hours}*3600*{rate('aca_vcpu_by_use', 'aca_mem_gib_by_use')},0)")
        d.sheet[f"{letter['aca_adt_use_usd']}{index}"] = (
            f"=IF({fits},{office}*3600*{rate('aca_vcpu_by_use', 'aca_mem_gib_by_use')},0)")
    sheet = workbook.create_sheet("ACA_por_cluster")
    title(sheet, "Propuesta: cada cluster PROD, costo AKS hoy vs Container Apps con sus réplicas reales",
          f"Réplicas medidas por Container Insights (promedio de {', '.join(full)}), workloads de aplicación (ReplicaSet). "
          "Tamaño ACA: la menor combinación que cubre el request o el uso. Peor caso para ACA: tarifa activa todo el "
          "tiempo, precio de lista sin descuento, sin la concesión gratuita; el request incluye los sidecars de Istio. "
          "Tamaño por uso = uso promedio + 50%. ACA va de 0,25 vCPU / 0,5 GiB a 4 vCPU / 8 GiB; el ADT_019 manda a AKS "
          "lo que pase de 4 vCPU / 8 GB, así que esos workloads quedan fuera del costo ACA.", 12)
    sheet.append([])
    headers = ["Cluster", "Workloads", "No caben en ACA", "Réplicas promedio oficina", "Réplicas promedio fuera de horario",
               f"Costo AKS {cost_month} (USD)", "ACA mismas réplicas, tamaño por request (peor caso)",
               "ACA mismas réplicas, tamaño por uso", "ACA regla ADT, tamaño por uso (fuera de horario en 0)",
               "AKS hoy − ACA por uso", "Más barato (AKS hoy vs ACA por uso)", "Id"]
    sheet.append(["Total de los clusters con datos"])
    total_row = sheet.max_row
    sheet.cell(row=total_row, column=1).font = BOLD
    header(sheet, headers)
    first = sheet.max_row + 1
    office_hours = f"({s['office_seconds']}/3600)"
    off_hours = f"({s['hours_month']}-{s['office_seconds']}/3600)"
    for cluster in sorted((row for row in clusters.rows if row["environment"] == "PROD"), key=lambda row: row["name"]):
        sheet.append([cluster["name"]])
        row = sheet.max_row
        by = f'{d.col("cluster")},$A{row}'
        sheet.cell(row=row, column=2, value=f"=COUNTIFS({by})")
        sheet.cell(row=row, column=3, value=f'=COUNTIFS({by},{d.col("fits_aca")},"NO")')
        sheet.cell(row=row, column=4, value=f'=IFERROR(SUMIFS({d.col("office_replica_hours_month")},{by})/{office_hours},"")'
                   ).number_format = "0.0"
        sheet.cell(row=row, column=5, value=f'=IFERROR(SUMIFS({d.col("offhours_replica_hours_month")},{by})/{off_hours},"")'
                   ).number_format = "0.0"
        if cost:
            sheet.cell(row=row, column=6, value=f'=SUMIFS({cost.col("cost_usd")},{cost.col("owner_id")},L{row},'
                                                f'{cost.col("owner_type")},"AKS",{cost.col("month")},"{cost_month}")'
                       ).number_format = USD
        for column, field in ((7, "aca_same_replicas_request_usd"), (8, "aca_same_replicas_use_usd"), (9, "aca_adt_use_usd")):
            sheet.cell(row=row, column=column, value=f'=SUMIFS({d.col(field)},{by})').number_format = USD
        sheet.cell(row=row, column=10, value=f'=IF(B{row}=0,"",F{row}-H{row})').number_format = '"$"#,##0;[Red]-"$"#,##0'
        sheet.cell(row=row, column=11, value=f'=IF(B{row}=0,"Sin datos",IF(H{row}<F{row},"ACA","AKS"))')
        sheet.cell(row=row, column=12, value=cluster.get("id", ""))
    last = sheet.max_row
    for column in (6, 7, 8, 9, 10):
        col = get_column_letter(column)
        sheet.cell(row=total_row, column=column, value=f'=SUMIFS({col}{first}:{col}{last},$B{first}:$B{last},">0")'
                   ).number_format = USD
    sheet.cell(row=total_row, column=2, value=f'=COUNTIFS(B{first}:B{last},">0")')
    sheet.cell(row=total_row, column=11, value=f'=COUNTIFS(K{first}:K{last},"ACA")&" de "&B{total_row}&" más baratos en ACA"')
    set_widths(sheet, [30, 11, 11, 13, 15, 16, 18, 16, 20, 16, 16, 40])
    sheet.column_dimensions["L"].hidden = True  # resource id, only used to match the cost
    sheet.freeze_panes = f"B{first}"


# ---------------------------------------------------------------------------
# Resumen, Fuentes, Consultas
# ---------------------------------------------------------------------------


def build_summary_by_bia(sheet, kclusters):
    """Clusters, cost, nodes, microservices and office CPU use per BIA level, PROD only."""
    k = kclusters
    sheet.append([])
    sheet.append(["Anexo informativo: por criticidad (BIA), PROD. No cambia las cifras de arriba"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    headers = ["BIA", "Clusters", "Costo mensual total", "Costo promedio por cluster", "Nodos promedio por cluster",
               "Microservicios promedio por cluster", "Uso de CPU en oficina"]
    sheet.append(headers)
    style_header(sheet, sheet.max_row, len(headers))
    present = {row["bia"] for row in k.rows if row["environment"] == "PROD"}
    levels = [level for level in BIA_LEVELS if level in present] + sorted(present - set(BIA_LEVELS))
    first = sheet.max_row + 1
    for level in levels + ["Total"]:
        sheet.append([level])
        row = sheet.max_row
        criteria = f'{k.col("environment")},"PROD",{k.col("in_inventory")},"SÍ"'
        if level != "Total":
            criteria += f',{k.col("bia")},$A{row}'
        sheet.cell(row=row, column=2, value=f"=COUNTIFS({criteria})")
        sheet.cell(row=row, column=3, value=f"=SUMIFS({k.col('cost_month_usd')},{criteria})").number_format = USD
        sheet.cell(row=row, column=4, value=f'=IFERROR(C{row}/B{row},"")').number_format = USD
        if "nodes_avg" in k.fields:
            sheet.cell(row=row, column=5, value=f'=IFERROR(SUMIFS({k.col("nodes_avg")},{criteria})/B{row},"")'
                       ).number_format = "0.0"
            sheet.cell(row=row, column=7, value=f'=IFERROR(SUMIFS({k.col("cpu_used_office_avg")},{criteria})/'
                                                f'SUMIFS({k.col("cpu_allocatable_avg")},{criteria}),"")').number_format = "0%"
        sheet.cell(row=row, column=6, value=f'=IFERROR(AVERAGEIFS({k.col("microservices")},{criteria}),"")'
                   ).number_format = "0.0"
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    sheet.append(["El BIA de un cluster es el más alto entre sus microservicios (export de deployments). Costo: "
                  "Cost Management, amortizado; uso: métricas de plataforma, lunes a viernes 8-20."])
    return first


def build_summary(workbook, breakeven, measured_refs, r, kclusters=None):
    sheet = workbook.create_sheet("Resumen", 0)
    title(sheet, "AKS vs Azure Container Apps: resumen", None)
    sheet.append([f"Generado: {date.today().isoformat()}. USD por mes. Tipo: medido (Azure), derivado (cálculo sobre "
                  "medidos), estimado (modelo con supuestos)."])
    sheet.append([])
    sheet.append(["Indicador", "Valor", "Tipo", "Origen"])
    style_header(sheet, 4, 4)
    s = {key: f"Supuestos!{value}" for key, value in r.items()}
    rows = [
        ("Clusters AKS PROD", f"={measured_refs['Clusters PROD']}", "Medido", "Medido_AKS", "0"),
        ("Clusters PROD con 4 nodos o menos", f"={measured_refs['Clusters PROD con 4 nodos o menos']}", "Medido", "Medido_AKS", "0"),
        ("Costo mensual promedio por cluster PROD", f"={measured_refs['Costo mensual promedio por cluster PROD (USD)']}",
         "Medido", "Medido_AKS; Cost Management", USD),
        ("Costo mensual total AKS PROD", f"={measured_refs['Costo mensual total AKS PROD (USD)']}", "Medido",
         "Medido_AKS; Cost Management", USD),
        ("Uso de CPU en horario de oficina, PROD", f"={measured_refs['Uso de CPU promedio en oficina, PROD']}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Uso de memoria en horario de oficina, PROD",
         f"={measured_refs.get('Uso de memoria promedio en oficina, PROD', chr(34) * 2)}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Uso de CPU, percentil 95 por hora (todo el día)",
         f"={measured_refs.get('Uso de CPU percentil 95 por hora, PROD', chr(34) * 2)}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Uso de CPU, máximo (cota superior: picos de cada cluster sumados)",
         f"={measured_refs.get('Uso de CPU máximo (pico de cada cluster sumado), PROD', chr(34) * 2)}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Uso de memoria, percentil 95 por hora (todo el día)",
         f"={measured_refs.get('Uso de memoria percentil 95 por hora, PROD', chr(34) * 2)}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Uso de memoria, máximo (cota superior: picos de cada cluster sumados)",
         f"={measured_refs.get('Uso de memoria máximo (pico de cada cluster sumado), PROD', chr(34) * 2)}", "Medido",
         "Medido_AKS; métricas de plataforma", "0%"),
        ("Descuento real de nodos (efectivo / lista)", f"={s['discount']}", "Derivado", "Supuestos; Cost Management", "0.00"),
        ("Cluster AKS mínimo (4 nodos) a precio de lista", f"={breakeven['Costo del cluster AKS mínimo al mes (lista)']}",
         "Estimado", "Modelo", USD),
        ("Cluster AKS mínimo con el descuento medido", f"={breakeven['Costo del cluster AKS mínimo al mes (efectivo)']}",
         "Estimado", "Modelo", USD),
        ("ACA por microservicio (horario ADT)", f"={breakeven['Costo ACA por microservicio al mes (horario ADT)']}",
         "Estimado", "Modelo", USD2),
        ("Microservicios desde los que ACA cuesta más que AKS (lista)",
         f"={breakeven['Microservicios en que ACA (horario ADT) supera a AKS lista']}", "Estimado", "Equilibrio", "0"),
        ("Microservicios desde los que ACA cuesta más que AKS (efectivo)",
         f"={breakeven['Microservicios en que ACA (horario ADT) supera a AKS efectivo']}", "Estimado", "Equilibrio", "0"),
        ("Microservicios desde los que ACA 24/7 cuesta más que AKS (lista)",
         f"={breakeven['Microservicios en que ACA 24/7 supera a AKS lista']}", "Estimado", "Equilibrio", "0"),
    ]
    for label, formula, kind, origin, fmt in rows:
        sheet.append([label, formula, kind, origin])
        sheet.cell(row=sheet.max_row, column=2).number_format = fmt
    sheet.append([f'">{MAX_MICROSERVICES}" = no hay cruce: de 1 a 30 y en 50, 100, 150, 200, 300 y {MAX_MICROSERVICES} microservicios ACA cuesta menos, porque '
                  "AKS suma nodos al crecer. AKS gana por tiempo activo o cluster compartido (hoja Equilibrio)."])
    if kclusters is not None and "bia" in kclusters.fields:
        build_summary_by_bia(sheet, kclusters)
    set_widths(sheet, [52, 16, 18, 36, 18, 22, 16])


SOURCE_ROWS = [
    ("Inventario de clusters, node pools y add-ons", "Medido", "Azure Resource Graph",
     "resources | where type =~ 'microsoft.containerservice/managedclusters'", "aks_clusters.csv, aks_node_pools.csv",
     "Resource Graph Explorer con la consulta de la hoja Consultas; o el cluster > Overview / Node pools"),
    ("Apps y ambientes de Container Apps", "Medido", "Azure Resource Graph",
     "resources | where type in~ ('microsoft.app/containerapps','microsoft.app/managedenvironments')",
     "aca_apps.csv, aca_environments.csv", "Container App > Scale; ambiente > Workload profiles"),
    ("APIs de APIM y su backend", "Medido", "ARM (lectura)", "GET {apim}/apis", "apim_apis.csv", "APIM > APIs > Settings > Web service URL"),
    ("Uso de CPU y memoria de nodos", "Medido", "Azure Monitor (métricas de plataforma)",
     "node_cpu_usage_millicores, node_memory_working_set_bytes, kube_node_status_allocatable_*; PT1H; por node y nodepool",
     "aks_usage_monthly.csv", "Cluster > Monitoring > Metrics, misma métrica, división por nodepool"),
    ("Peticiones por API", "Medido", "Azure Monitor (métricas de plataforma)", "Requests (Total), PT5M, división por ApiId",
     "apim_api_monthly.csv", "APIM > Monitoring > Metrics > Requests, aplicar división por API"),
    ("Despertares y segundos activos fuera de horario", "Derivado", "Azure Monitor", "Rachas de intervalos de 5 min con "
     "peticiones fuera de lunes-viernes 8-20 (Lima); cooldown de 300 s por racha", "apim_api_monthly.csv",
     "Recalculable con aa_metrics.py --cooldown"),
    ("Réplicas, requests y uso por workload", "Medido", "Log Analytics (Container Insights)",
     "KubePodInventory + Perf (K8SContainer); agregado por hora, mes y horario", "aks_workloads_monthly.csv",
     "Cluster > Insights > Controllers; o Logs con la consulta de la hoja Consultas"),
    ("Costo real de AKS y Container Apps", "Medido (costo real facturado)", "Cost Management Query API",
     "AmortizedCost (reservas y savings plan repartidos), mensual, agrupado por ResourceId + MeterCategory + Meter; cluster + grupo de recursos de nodos",
     "cost_by_resource.csv", "Cost analysis (vista Amortized cost) > alcance suscripción > filtro grupo de recursos MC_* y tipo managedClusters"),
    ("Descuento real de nodos", "Derivado de costo real", "Cost Management + Retail Prices",
     "Costo amortizado de VMs de nodos PROD / (horas-nodo medidas × precio lista), meses completos", "Supuestos (discount)", "Cost analysis, medidor de la VM"),
    ("Precios de lista", "Dato del proveedor", "Azure Retail Prices API (pública)",
     "Filtros en la columna 'query' de Precios", "prices.csv", "https://azure.microsoft.com/pricing/calculator"),
    ("Réplica mínima, horario y límites de Container Apps", "Lineamiento interno", "ADT_019",
     "1 réplica de 8 a 20 en días hábiles; 4 vCPU / 8 GiB; umbrales 15 microservicios, 10 h, 100 tps", "Supuestos, Equilibrio",
     "ADT_019 en Confluence"),
    ("Supuestos de carga y capacidad", "Supuesto", "Criterio de arquitectura",
     "Celdas amarillas de Supuestos (RPS por réplica, overhead de daemonsets y sidecar)", "Supuestos",
     "Validar con una prueba de carga"),
]


def build_sources(workbook, extra_rows=()):
    sheet = workbook.create_sheet("Fuentes")
    title(sheet, "De dónde sale cada cifra",
          "Medido = dato leído de Azure en la corrida. Derivado = cálculo sobre medidos. Estimado = modelo con supuestos. "
          "El costo viene de Cost Management (AmortizedCost): incluye descuentos del contrato y reparte reservas y savings plan entre las VMs que los usan.", 6)
    sheet.append([])
    headers = ["Dato", "Tipo", "Fuente", "Cómo se obtiene", "Archivo de salida", "Cómo verificarlo"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    for values in SOURCE_ROWS + list(extra_rows):
        sheet.append(list(values))
        for column in range(1, 7):
            sheet.cell(row=sheet.max_row, column=column).alignment = WRAP
    set_widths(sheet, [34, 18, 26, 60, 26, 50])


def build_queries(workbook, results):
    sheet = workbook.create_sheet("Consultas")
    title(sheet, "Consultas ejecutadas (evidencia registrada por los scripts)")
    sheet.append([])
    headers = ["Script", "Fecha de la corrida", "Tipo", "Etiqueta / endpoint", "Consulta, cuerpo o api-version",
               "Filas / llamadas", "Errores"]
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    for folder, name in (("aa-inventory", "inventory_summary.json"), ("aa-metrics", "metrics_summary.json"),
                         ("aa-container-insights", "ci_summary.json"), ("aa-cost", "cost_summary.json")):
        path = os.path.join(results, folder, name)
        if not os.path.exists(path):
            sheet.append([name, "", "Sin corrida", "", "", "", ""])
            continue
        with open(path, encoding="utf-8") as handle:
            summary = json.load(handle)
        evidence, when = summary.get("evidence", {}), summary.get("generated_at", "")
        for query in evidence.get("resource_graph_queries", []):
            sheet.append([name, when, query.get("source", "Azure Resource Graph"), query["label"], query["query"],
                          query["rows"], ""])
        for call in evidence.get("arm_calls", []):
            sheet.append([name, when, "ARM (lectura)", f"{call['method']} {call['pattern']}",
                          f"api-version={call['api_version']}", call["calls"], call["errors"]])
        for body in evidence.get("request_bodies", []):
            sheet.append([name, when, "Cuerpo de consulta", f"POST {body['endpoint']}",
                          json.dumps(body["body"], ensure_ascii=False), "", ""])
    prices = os.path.join(results, "aa-prices", "prices_summary.json")
    if os.path.exists(prices):
        with open(prices, encoding="utf-8") as handle:
            summary = json.load(handle)
        for query in summary.get("queries", []):
            sheet.append(["prices_summary.json", summary.get("generated_at", ""), "Retail Prices API", query["group"],
                          query["filter"], query["rows"], ""])
    for row in sheet.iter_rows(min_row=4):
        for cell in row:
            cell.alignment = WRAP
    set_widths(sheet, [24, 20, 24, 44, 90, 12, 10])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def add_pool_sizes(usage, pools, clusters):
    """VM size and region of each usage row, so measured node-hours can be priced at list."""
    sizes = {(pool["cluster"].lower(), pool["pool"].lower()): pool["vm_size"] for pool in pools}
    regions = {cluster["name"].lower(): cluster["location"] for cluster in clusters}
    for row in usage:
        row["vm_size"] = sizes.get((row["cluster"].lower(), row["pool"].lower()), "")
        row["location"] = regions.get(row["cluster"].lower(), "")


def full_months(usage, cost):
    """Months with cost that platform metrics cover end to end (the metrics window rarely starts on day 1)."""
    hours = defaultdict(int)
    for row in usage:
        hours[row["month"]] = max(hours[row["month"]], int(float(row.get("hours") or 0)))
    cost_months = {row["month"] for row in cost}
    return sorted(month for month, value in hours.items()
                  if month in cost_months and value >= monthrange(int(month[:4]), int(month[5:]))[1] * 24)


def distinct_months(rows):
    return len({row["month"] for row in rows if row.get("month")})


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output-dir", default=os.path.join(aa.TOPIC_DIR, "reports", f"{date.today().isoformat()}-aks-aca"))
    args = parser.parse_args()

    prices = load(args.results, "aa-prices", "prices.csv")
    if not prices:
        raise SystemExit("Run aa_prices.py first: the model needs results/aa-prices/prices.csv")
    clusters = load(args.results, "aa-inventory", "aks_clusters.csv")
    add_version_lag(clusters)
    normalize_booleans(clusters)
    pools = load(args.results, "aa-inventory", "aks_node_pools.csv")
    apis = load(args.results, "aa-inventory", "apim_apis.csv")
    apps = load(args.results, "aa-inventory", "aca_apps.csv")
    usage = load(args.results, "aa-metrics", "aks_usage_monthly.csv")
    add_pool_sizes(usage, pools, clusters)
    traffic = load(args.results, "aa-metrics", "apim_api_monthly.csv")
    workloads = load(args.results, "aa-container-insights", "aks_workloads_monthly.csv")
    services = [row for row in load(args.results, "aa-container-insights", "aks_services.csv")
                if row["namespace"] not in aa.SYSTEM_NAMESPACES and not row["namespace"].startswith("aks-istio")]
    cost = load(args.results, "aa-cost", "cost_by_resource.csv")
    profile = [row for row in load(args.results, "aa-metrics", "apim_hourly_profile.csv") if row["api_id"] != "__Empty"]
    last_cost_month = max((row["month"] for row in cost), default="")
    hpas = load(args.results, "aa-k8s-snapshot", "hpa_snapshot.csv")
    deployments = load(args.results, "aa-k8s-snapshot", "deployment_snapshot.csv")
    k8s_clusters = load(args.results, "aa-k8s-snapshot", "cluster_snapshot.csv")
    capabilities = load(args.results, "aa-k8s-snapshot", "capability_domains.csv")
    snapshot_path = os.path.join(args.results, "aa-k8s-snapshot", "snapshot_summary.json")
    snapshot = {}
    if os.path.exists(snapshot_path):
        with open(snapshot_path, encoding="utf-8") as handle:
            snapshot = json.load(handle)
    # Metrics and Container Insights may hold clusters the inventory no longer lists (for example POC ones).
    names = {row["name"] for row in clusters}
    usage, workloads, services = ([row for row in rows if row["cluster"] in names] for rows in (usage, workloads, services))

    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.calculation.fullCalcOnLoad = True

    price_fields = ["key", "region", "service", "product", "sku", "meter", "unit", "retail_price", "hourly_equivalent",
                    "price_type", "reservation_term", "tier_minimum_units", "effective_start", "meter_id", "arm_sku", "query"]
    Detail(workbook, "Precios", prices, price_fields)
    cluster_fields = ["name", "environment", "environment_evidence", "subscription", "resource_group", "location",
                      "kubernetes_version", "minor_versions_behind", "node_pools", "system_nodes", "user_nodes",
                      "user_max_nodes", "vm_sizes", "zones"] + [p for p in CONFIG_PROPERTIES if p not in
                      ("kubernetes_version", "minor_versions_behind", "zones")] + ["istio_revisions", "extension_types", "id"]
    d_clusters = Detail(workbook, "D_Clusters", clusters, cluster_fields) if clusters else None
    if pools:
        Detail(workbook, "D_Pools", pools, list(pools[0].keys()))
    if apps:
        Detail(workbook, "D_ACA_apps", apps, list(apps[0].keys()))
    # Allocatable rows lack the usage columns and vice versa, so fields are the union of every row.
    d_usage = None
    if usage:
        d_usage = Detail(workbook, "D_Uso", usage, all_fields(usage), complete=True, extra=["list_cost_usd"])
        letters = {field: d_usage.letter(field) for field in d_usage.fields}
        for index in range(2, len(usage) + 2):
            d_usage.sheet[f"{letters['list_cost_usd']}{index}"] = (
                f'=IF(OR(${letters["vm_size"]}{index}="",${letters["nodes_avg"]}{index}=""),"",'
                f'${letters["nodes_avg"]}{index}*${letters["hours"]}{index}*SUMIFS(Precios!$H:$H,'
                f'Precios!$A:$A,"vm_"&${letters["vm_size"]}{index}&"_payg",Precios!$B:$B,${letters["location"]}{index}))')
    d_traffic = Detail(workbook, "D_Trafico", traffic, list(traffic[0].keys())) if traffic else None
    d_workloads = Detail(workbook, "D_Workloads", workloads, all_fields(workloads), complete=True) if workloads else None
    d_services = Detail(workbook, "D_Services", services, list(services[0].keys())) if services else None
    d_cost = Detail(workbook, "D_Costo", cost, list(cost[0].keys())) if cost else None
    d_hpa = Detail(workbook, "D_HPA", hpas, list(hpas[0].keys())) if hpas else None
    d_deploy = Detail(workbook, "D_Deploy", deployments, list(deployments[0].keys())) if deployments else None
    d_kclusters = None
    if k8s_clusters:
        usage_extra = ["nodes_avg", "cpu_allocatable_avg", "cpu_used_office_avg"] if d_usage else []
        d_kclusters = Detail(workbook, "D_K8s_clusters", k8s_clusters, list(k8s_clusters[0].keys()),
                             extra=["cost_month_usd"] + usage_extra)
        usage_months = max(1, distinct_months(usage))
        for index in range(2, len(k8s_clusters) + 2):
            # Last full month, not the average: clusters created inside the period would look cheaper.
            d_kclusters.sheet[f"{d_kclusters.letter('cost_month_usd')}{index}"] = (
                f'=SUMIFS({d_cost.col("cost_usd")},{d_cost.col("owner")},$A{index},{d_cost.col("owner_type")},"AKS",'
                f'{d_cost.col("month")},"{last_cost_month}")' if d_cost else "")
            for field in usage_extra:
                d_kclusters.sheet[f"{d_kclusters.letter(field)}{index}"] = (
                    f'=SUMIFS({d_usage.col(field)},{d_usage.col("cluster")},$A{index})/{usage_months}')

    months = {"cost": distinct_months(cost), "usage": distinct_months(usage)}
    apis_last = build_traffic(workbook, d_traffic, apis, distinct_months(traffic)) if d_traffic else None
    workloads_last = build_workloads(workbook, d_workloads) if d_workloads else None
    refs = {}
    r = build_assumptions(workbook, refs, apis_last, workloads_last, d_cost, d_usage, full_months(usage, cost))
    model = build_model(workbook, r)
    map_info = build_map(workbook, r)
    breakeven = build_breakeven(workbook, r, model, map_info)
    build_sensitivity(workbook, r)
    if profile:
        build_traffic_profile(workbook, profile, d_traffic)
    if d_clusters and workloads:
        build_aca_by_cluster(workbook, r, d_clusters, workloads, d_cost, last_cost_month)
    if d_clusters:
        measured_refs, _ = build_measured(workbook, d_clusters, d_usage, d_cost, d_workloads, d_services, months)
        build_config(workbook, d_clusters)
    snapshot_sources = []
    if d_hpa or d_deploy:
        taken_on = snapshot.get("snapshot_date", "")
        build_k8s(workbook, d_hpa, d_deploy, taken_on)
        if d_deploy and d_kclusters:
            build_patterns(workbook, d_deploy, d_hpa, d_kclusters, capabilities, taken_on)
            if "bia" in d_kclusters.fields:
                build_bia(workbook, d_kclusters)
        snapshot_sources.append((
            f"HPA, requests y límites por deployment (foto del {taken_on})", "Medido (un día)",
            "Export del equipo de plataforma", ", ".join(item["file"] for item in snapshot.get("inputs", [])),
            "hpa_snapshot.csv, deployment_snapshot.csv",
            "kubectl get hpa / get deployment en el cluster; la historia, en Container Insights"))
    else:
        measured_refs = {key: '""' for key in ("Clusters PROD", "Clusters PROD con 4 nodos o menos",
                                               "Costo mensual promedio por cluster PROD (USD)",
                                               "Costo mensual total AKS PROD (USD)", "Uso de CPU promedio en oficina, PROD")}
    build_criteria(workbook, d_clusters, d_kclusters, d_deploy, d_hpa, d_traffic, workloads_last,
                   max((row["month"] for row in traffic if row["month"] < date.today().strftime("%Y-%m")), default=""))
    build_summary(workbook, breakeven, measured_refs, r, d_kclusters)
    build_sources(workbook, snapshot_sources)
    build_queries(workbook, args.results)

    order = ["Resumen", "Patrones", "Criterios_tecnicos", "Supuestos", "Modelo", "Equilibrio", "Mapa", "ACA_por_cluster", "Perfil_trafico", "Medido_AKS", "Trafico_APIM", "Workloads",
             "Config_AKS", "Config_K8s", "Fuentes", "Precios", "Candidatos_BIA"]
    workbook._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order
                          else (len(order) + 1 if ws.title == "Consultas" else len(order)))
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.font is None or cell.font.name != FONT:
                    cell.font = Font(name=FONT, bold=cell.font.bold if cell.font else False,
                                     color=cell.font.color if cell.font else None, size=cell.font.size if cell.font else 11)
    os.makedirs(args.output_dir, exist_ok=True)
    path = os.path.join(args.output_dir, "aks_aca_comparison.xlsx")
    workbook.save(path)
    print(path)


if __name__ == "__main__":
    main()
