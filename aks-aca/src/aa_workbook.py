#!/usr/bin/env python3
"""Build the AKS vs Container Apps workbook from the aa_*.py outputs.

Reads results/aa-prices, aa-inventory, aa-metrics, aa-container-insights, aa-cost and aa-k8s-snapshot (any of
them may be missing; the sheets that need it say so) and writes aks_aca_comparison.xlsx:

  Summary        key figures, all formulas
  Clusters       PROD clusters by category (APPLICATION, NO_APPLICATIONS, INFRASTRUCTURE, NO_DATA)
  Assumptions    editable inputs (yellow) with the measured reference next to each one
  Model          monthly cost for 1 to 30 microservices: ACA (ADT office-hours model and 24x7) and
                 AKS (list, effective with the measured discount, 1-year reservation), with a chart
  Break_even     break-even points and the ADT_019 thresholds tested against them
  Map            microservices x requests per month: ACA minus AKS (negative = ACA is cheaper)
  Measured_AKS   per cluster: nodes, monthly cost, CPU/memory used vs allocatable, workloads
  APIM_traffic   per API: requests, office share, peaks, off-hours wake-ups
  Workloads      per workload: replicas and per-replica CPU/memory (Container Insights)
  Config_AKS     configuration dispersion across clusters
  Prices, D_*    list prices and the detail rows every formula reads
  Sources        where each figure comes from and how to check it
  Queries        exact queries and calls recorded by each script

Only APPLICATION clusters enter the PROD figures. The category comes from cluster_snapshot.csv; without it,
clusters of the apps in aa_common.INFRASTRUCTURE_APPS are INFRASTRUCTURE and every other one APPLICATION.

Every model figure is a formula over Assumptions and Prices, so editing a yellow cell recalculates
the curve, the break-even points and the map.
"""

import argparse
import math
import json
import os
from calendar import month_name, monthrange
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
APP = f'"{aa.APPLICATION}"'


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


def period_text(months):
    """'October 2026' for ["2026-10"]; 'August–October 2026' or 'December 2025–January 2026' for a range."""
    months = sorted(month for month in months if month)
    if not months:
        return "the measured period"

    def name(month):
        return month_name[int(month[5:7])]

    first, last = months[0], months[-1]
    if first == last:
        return f"{name(first)} {first[:4]}"
    if first[:4] == last[:4]:
        return f"{name(first)}–{name(last)} {last[:4]}"
    return f"{name(first)} {first[:4]}–{name(last)} {last[:4]}"


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
# Cluster categories
# ---------------------------------------------------------------------------

CATEGORIES = [
    (aa.APPLICATION, "Runs at least one microservice (or only the application's own gateway): the only category "
                     "in the analysis"),
    (aa.NO_APPLICATIONS, "Created without applications: the platform export or Container Insights show only "
                         "platform components"),
    (aa.INFRASTRUCTURE, "Runs infrastructure tools only (runners, agents, observability, gateways): excluded"),
    (aa.NO_DATA, "Neither in the platform export nor with Container Insights data: excluded"),
]
FALLBACK_NOTE = ("cluster_snapshot.csv is missing, so categories come from the app code only: clusters of "
                 "infrastructure apps (" + ", ".join(sorted(aa.INFRASTRUCTURE_APPS)) + ") are INFRASTRUCTURE and "
                 "every other cluster is treated as APPLICATION; NO_APPLICATIONS and NO_DATA stay at 0.")


def cluster_categories(clusters, k8s_clusters):
    """Resource id (lowercase) -> (category, reason) for every inventory cluster, and whether the categories come
    from cluster_snapshot.csv. Without it, clusters of infrastructure apps are INFRASTRUCTURE, the rest APPLICATION."""
    if k8s_clusters:
        known = {row["cluster_id"].lower(): (row.get("category") or aa.NO_DATA, row.get("category_reason", ""))
                 for row in k8s_clusters if row.get("cluster_id")}
        return {row["id"].lower(): known.get(row["id"].lower(), (aa.NO_DATA, "Not in cluster_snapshot.csv"))
                for row in clusters}, True
    categories = {}
    for row in clusters:
        code = str(row.get("cod_app") or "").upper()
        if code in aa.INFRASTRUCTURE_APPS:
            categories[row["id"].lower()] = (aa.INFRASTRUCTURE, f"App code {code}: {aa.INFRASTRUCTURE_APPS[code]}")
        else:
            categories[row["id"].lower()] = (aa.APPLICATION, "Assumed: no cluster_snapshot.csv and the app code is "
                                                             "not an infrastructure app")
    return categories, False


def set_category(rows, categories, id_field="cluster_id"):
    """Add `category` to rows that do not carry one, from the cluster's resource id."""
    for row in rows:
        if not row.get("category"):
            row["category"] = categories.get(str(row.get(id_field) or "").lower(), ("", ""))[0]


# ---------------------------------------------------------------------------
# Assumptions
# ---------------------------------------------------------------------------

# (key, label, value or None for a formula, note). Sections are rows with key None.
ASSUMPTIONS = [
    (None, "General", None, None),
    ("region", "Pricing region", "eastus2", "ARM region. East US 2 is the reference; Central US is in Prices for sensitivity."),
    ("hours_month", "Hours per month", 730, "Azure convention for monthly prices."),
    ("office_days", "Working days per month", 21.75, "261 working days a year / 12. Public holidays are not discounted."),
    ("office_hours", "Office hours per day", 12, "ADT_019: Monday to Friday, 8:00 to 20:00."),
    ("office_seconds", "Office seconds per month", None, "Formula: working days × hours × 3600."),
    ("n_ref", "Microservices in the reference case", 10, "Used in the hours, replicas and requests thresholds."),
    (None, "Microservice on Container Apps (Consumption)", None, None),
    ("cpu_rep", "vCPU per replica", 0.5, "Model assumption. ACA allows 0.25 vCPU / 0.5 GiB to 4 vCPU / 8 GiB; ADT_019 only sets the maximum: more than 4 vCPU / 8 GB goes to AKS."),
    ("mem_rep", "Memory per replica (GiB)", 1.0, "In ACA memory is 2 GiB per vCPU."),
    ("min_rep_office", "Minimum replicas in office hours", 1, "ADT_019: 1 replica from 8 to 20 on working days; scales to zero outside office hours."),
    ("req_month", "Requests per microservice per month", 1_000_000, "Reference: median monthly requests per API in APIM."),
    ("office_req_share", "Share of requests in office hours", 0.9, "Reference: office requests / total in APIM."),
    ("rps_rep", "Requests per second one replica serves", 50, "Assumption: validate with a load test of a typical microservice."),
    ("peak_factor", "Hourly peak / office average", 3, "Reference: hourly peak / average per office hour in APIM."),
    ("office_active_share", "Share of office time with traffic (active rate)", 0.3,
     "Reference: 5-minute intervals with requests / office intervals in APIM. The rest is billed at the idle rate."),
    ("off_active_s", "Active seconds outside office hours per microservice per month", 36_000,
     "Reference: 5-minute intervals with traffic outside office hours × 300 s (APIM)."),
    ("off_idle_s", "Idle seconds outside office hours per microservice per month (cooldown)", 12_000,
     "Reference: wake-ups outside office hours × 300 s cooldown (APIM)."),
    ("env_count", "Container Apps environments per application", 1, "One environment per application."),
    ("env_mgmt_apply", "Charge 'Environment Management Hour' (1 = yes)", 1,
     "Meter effective since 2026-09; Microsoft Learn does not say which environments it applies to. In "
     "September–October 2026 the bank's environments (workload profiles without a dedicated profile) did not pay "
     "it: 1 is the conservative case."),
    ("pe_apply", "Charge 'Environment Private Endpoint' (1 = yes)", 1,
     "ADT_019: private network. No charges in the bank's environments in September–October 2026. Set 0 if the "
     "environment uses VNet integration without a private endpoint."),
    ("grant_apply", "Apply the monthly free grant (1 = yes)", 1,
     "Per subscription and month; shared by every app in the subscription. Set 0 if the subscription already uses it."),
    ("grant_vcpu_s", "Free grant: vCPU-seconds", 180_000, "Container Apps pricing (Consumption)."),
    ("grant_mem_s", "Free grant: GiB-seconds", 360_000, "Container Apps pricing (Consumption)."),
    ("grant_requests", "Free grant: requests", 2_000_000, "Container Apps pricing (Consumption)."),
    (None, "Minimum production AKS cluster", None, None),
    ("system_nodes", "System pool nodes", 2, "Production minimum: 4 nodes of 8 vCPU / 32 GB (2 system + 2 user)."),
    ("user_nodes_min", "Minimum user pool nodes", 2, ""),
    ("node_sku", "Node VM size", "Standard_D8s_v5", "Reference: most frequent size in the PROD node pools."),
    ("node_vcpu", "vCPU per node", 8, ""),
    ("node_mem", "Memory per node (GiB)", 32, ""),
    ("os_disk_tier", "OS disk (tier, or 'ephemeral')", "P10", "A 128 GB managed disk is P10. An ephemeral disk has no disk cost."),
    ("aks_tier_key", "Cluster tier (Prices key)", "aks_standard_tier_hour",
     "aks_standard_tier_hour (Standard, SLA) or aks_lts_hour (Premium/LTS). Free = any other text."),
    ("lb_count", "Standard load balancers", 1, "The cluster's outbound one."),
    ("ip_count", "Standard public IPs", 1, "Outbound IP. 0 if outbound traffic goes through a firewall (outboundType userDefinedRouting)."),
    ("defender_apply", "Defender for Containers per vCore (1 = yes)", 1,
     "Only counts if it is enabled on AKS; on Container Apps it does not charge per vCore."),
    ("alloc_cpu_share", "Allocatable share of CPU per node", 0.85, "Kube-reserved; compare with allocatable vCPU in Measured_AKS."),
    ("alloc_mem_share", "Allocatable share of memory per node", 0.8, "Kube-reserved and eviction threshold."),
    ("daemon_cpu", "vCPU per node used by DaemonSets (monitoring, CSI, CNI)", 0.3, "Assumption."),
    ("daemon_mem", "GiB per node used by DaemonSets", 0.75, "Assumption."),
    ("sidecar_cpu", "Istio sidecar vCPU request per replica", 0.1, "Default sidecar request of the Istio add-on."),
    ("sidecar_mem", "Istio sidecar GiB request per replica", 0.125, "Default sidecar request of the Istio add-on."),
    ("aks_min_rep", "Minimum replicas per microservice on AKS (24x7)", 2, "High availability in production."),
    ("rs_headroom", "AKS right sizing: headroom over measured use", 1.5,
     "Same headroom as the Container Apps size by use (use + 50%). Requires lowering the pod requests to that value."),
    ("rs_usable", "AKS right sizing: usable share of each node", 0.85, "Margin for DaemonSets, sidecars and peaks."),
    ("rs_system_nodes", "AKS right sizing: system nodes", 2, "Production minimum: 2 system + 2 user."),
    ("rs_user_min", "AKS right sizing: minimum user nodes", 2,
     "Floor of the production minimum. With each cluster's current minimum the saving would be larger."),
    ("discount", "Effective cost / list price of nodes", None,
     "Formula: amortized cost of the node VMs of PROD application clusters / (measured node-hours × list price), "
     "in the months both metrics and cost cover end to end. 1 if there is no measured cost."),
    (None, "Calculated (do not edit)", None, None),
    ("p_vcpu_active", "ACA: USD per active vCPU-second", None, "Prices."),
    ("p_vcpu_idle", "ACA: USD per idle vCPU-second", None, "Prices."),
    ("p_mem_active", "ACA: USD per active GiB-second", None, "Prices."),
    ("p_mem_idle", "ACA: USD per idle GiB-second", None, "Prices."),
    ("p_requests", "ACA: USD per million requests", None, "Prices."),
    ("rate_active", "ACA: USD per second of an active replica", None, "vCPU × rate + GiB × rate."),
    ("rate_idle", "ACA: USD per second of an idle replica", None, "vCPU × rate + GiB × rate."),
    ("aca_env_month", "ACA: fixed environment cost per month", None, "Environment management + private endpoint."),
    ("aca_grant", "ACA: value of the monthly free grant", None, "At the active rate."),
    ("p_vm_hour", "AKS: USD per node hour (list)", None, "Prices, Linux pay-as-you-go."),
    ("p_vm_1y_hour", "AKS: USD per node hour (1-year reservation)", None, "Prices, 1-year reservation / 8760 h."),
    ("disk_month", "AKS: OS disk per node per month", None, "Prices."),
    ("defender_node_month", "AKS: Defender per node per month", None, "vCPU × USD per vCore-hour × hours."),
    ("aks_fixed_month", "AKS: fixed cluster cost per month", None, "Tier + load balancer + public IPs."),
    ("user_cpu_node", "AKS: vCPU available per user node", None, "vCPU × allocatable share − DaemonSets."),
    ("user_mem_node", "AKS: GiB available per user node", None, "GiB × allocatable share − DaemonSets."),
]


def price(key_expr, column="H"):
    """SUMIFS lookup in Prices by key (text or expression) and the selected region."""
    return f'SUMIFS(Prices!${column}:${column},Prices!$A:$A,{key_expr},Prices!$B:$B,{{region}})'


def build_assumptions(workbook, refs, apis_sheet, workloads_sheet, cost, usage=None, discount_months=()):
    sheet = workbook.create_sheet("Assumptions")
    title(sheet, "Assumptions (yellow cells: editable)",
          "Column C: value the model uses. Column E: reference measured at the bank when one exists; copy it to C "
          "to model with the real figure. List prices in USD; the 'discount' row brings the actual discount.", 5)
    sheet.append([])
    sheet.append(["Key", "Assumption", "Value", "Source / note", "Measured reference"])
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
        refs[key] = f"Assumptions!$C${row}"
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
                    f'{cost.col("owner_type")},"AKS",{cost.col("environment")},"PROD",{cost.col("category")},{APP},'
                    f'{cost.col("month")},{months}))/'
                    f'SUM(SUMIFS({usage.col("list_cost_usd")},{usage.col("environment")},"PROD",'
                    f'{usage.col("category")},{APP},{usage.col("month")},{months})),1)', "0.00")
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
        reference("req_month", f"=IFERROR(MEDIAN(APIM_traffic!$D$5:$D${apis_sheet}),\"\")", "#,##0")
        reference("office_req_share", f"=IFERROR(SUM(APIM_traffic!$E$5:$E${apis_sheet})/SUM(APIM_traffic!$D$5:$D${apis_sheet}),\"\")", "0%")
        reference("peak_factor", f"=IFERROR(MEDIAN(APIM_traffic!$J$5:$J${apis_sheet}),\"\")", "0.0")
        reference("office_active_share", f"=IFERROR(MEDIAN(APIM_traffic!$K$5:$K${apis_sheet}),\"\")", "0%")
        reference("off_active_s", f"=IFERROR(MEDIAN(APIM_traffic!$L$5:$L${apis_sheet}),\"\")", "#,##0")
        reference("off_idle_s", f"=IFERROR(MEDIAN(APIM_traffic!$M$5:$M${apis_sheet}),\"\")", "#,##0")
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
    os_ = f"Assumptions!{r['office_seconds']}"
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    rps = f"({req}*{s['office_req_share']}/{os_})"
    replicas = f"MAX({s['min_rep_office']},CEILING({rps}/{rps_rep or s['rps_rep']},1))"
    extra = f"({replicas}-{s['min_rep_office']})*{os_}*{s['rate_active']}"
    if warm == "ADT":
        warm_cost = (f"{s['min_rep_office']}*{os_}*({s['office_active_share']}*{s['rate_active']}"
                     f"+(1-{s['office_active_share']})*{s['rate_idle']})")
        off = f"{s['off_active_s']}*{s['rate_active']}+{s['off_idle_s']}*{s['rate_idle']}"
    else:  # warm 24x7: the minimum replica never scales to zero
        warm_cost = (f"{s['min_rep_office']}*{s['hours_month']}*3600*({s['office_active_share']}*{s['rate_active']}"
                     f"+(1-{s['office_active_share']})*{s['rate_idle']})")
        off = "0"
    per_ms = f"({warm_cost}+{extra}+{off}+{req}/1000000*{s['p_requests']})"
    return f"MAX(0,{n}*{per_ms}-{s['aca_grant']})+{s['aca_env_month']}"


def aks_nodes_formula(r, n, req, rps_rep=None):
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    rps = f"({req}*{s['office_req_share']}/{s['office_seconds']})"
    peak = f"MAX({s['aks_min_rep']},CEILING({rps}*{s['peak_factor']}/{rps_rep or s['rps_rep']},1))"
    return (f"MAX({s['user_nodes_min']},CEILING({n}*{peak}*({s['cpu_rep']}+{s['sidecar_cpu']})/{s['user_cpu_node']},1),"
            f"CEILING({n}*{peak}*({s['mem_rep']}+{s['sidecar_mem']})/{s['user_mem_node']},1))")


def aks_formula(r, nodes, mode="list"):
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    vm = {"list": s["p_vm_hour"], "effective": f"{s['p_vm_hour']}*{s['discount']}", "1y": s["p_vm_1y_hour"]}[mode]
    return (f"({s['system_nodes']}+{nodes})*({vm}*{s['hours_month']}+{s['disk_month']}+{s['defender_node_month']})"
            f"+{s['aks_fixed_month']}")


def build_model(workbook, r):
    sheet = workbook.create_sheet("Model")
    title(sheet, "Monthly cost per application: Container Apps vs AKS (USD, list prices except 'effective')",
          "ACA ADT schedule: 1 replica from 8 to 20 on working days, scaled to zero outside office hours. ACA 24x7: the "
          "minimum replica never goes to zero. AKS: a cluster dedicated to the application (as today), 2 system nodes + "
          "user nodes that grow when the replicas do not fit. Effective: node price × measured discount.", 10)
    sheet.append([])
    headers = ["Microservices", "ACA ADT schedule", "ACA 24x7", "AKS list", "AKS effective", "AKS 1-year reservation",
               "AKS user nodes", "ACA ADT − AKS list", "Cheaper (list)", "Cheaper (effective)"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    req = f"Assumptions!{r['req_month']}"
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
    chart.title = "Monthly cost per application (USD)"
    chart.y_axis.title = "USD / month"
    chart.x_axis.title = "Microservices"
    chart.height, chart.width = 9, 18
    chart.add_data(Reference(sheet, min_col=2, max_col=5, min_row=4, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=5, max_row=last))
    sheet.add_chart(chart, "L4")
    set_widths(sheet, [14, 16, 14, 14, 14, 16, 12, 18, 16, 18])
    sheet.freeze_panes = "B5"
    return {"first": 5, "last": last}


def build_breakeven(workbook, r, model, map_info):
    sheet = workbook.create_sheet("Break_even")
    title(sheet, "Break-even point and ADT_019 thresholds",
          "First number of microservices at which Container Apps costs more than the dedicated AKS cluster, with the "
          "inputs of the Assumptions sheet. The hours and replicas thresholds use the reference case (n_ref).", 5)
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    first, last = model["first"], model["last"]

    def first_n(aca_col, aks_col):
        return (f'=IFERROR(INDEX(Model!$A${first}:$A${last},MATCH(1,INDEX((Model!${aca_col}${first}:${aca_col}${last}'
                f'>Model!${aks_col}${first}:${aks_col}${last})*1,0),0)),">{MAX_MICROSERVICES}")')

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
        ("Microservices at which ACA (ADT schedule) exceeds AKS list", first_n("B", "D"), "0"),
        ("Microservices at which ACA (ADT schedule) exceeds AKS effective", first_n("B", "E"), "0"),
        ("Microservices at which ACA (ADT schedule) exceeds AKS with a 1-year reservation", first_n("B", "F"), "0"),
        ("Microservices at which ACA 24x7 exceeds AKS list", first_n("C", "D"), "0"),
        ("ACA cost per microservice per month (ADT schedule)", f"=Model!B{first + 1}-Model!B{first}", USD2),
        ("Minimum AKS cluster cost per month (list)", f"=Model!D{first}", USD),
        ("Minimum AKS cluster cost per month (effective)", f"=Model!E{first}", USD),
        ("Warm replica hours per working day at which ACA equals AKS list (n_ref)",
         f"=IFERROR(({budget}-{not_warm})/({s['min_rep_office']}*{s['office_days']}*3600*{warm_rate}),\"\")", "0.0"),
        ("Average warm replicas per microservice at which ACA equals AKS list (n_ref, office hours)",
         f"=IFERROR(({budget}-{base})/({s['office_seconds']}*{warm_rate}),\"\")", "0.0"),
        ("AKS list cost of the reference case", "=" + aks_ref, USD),
        ("AKS effective cost of the reference case", "=" + aks_ref_eff, USD),
    ]
    sheet.append([])
    sheet.append(["Indicator", "Value"])
    style_header(sheet, 4, 2)
    cells = {}
    for label, formula, fmt in rows:
        sheet.append([label, formula])
        sheet.cell(row=sheet.max_row, column=2).number_format = fmt
        sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
        cells[label] = f"Break_even!$B${sheet.max_row}"

    sheet.append([])
    sheet.append(["ADT_019 rule", "ADT threshold", "Model result", "Reading"])
    style_header(sheet, sheet.max_row, 4)
    n_star = cells[rows[0][0]]
    hours_star = cells[rows[7][0]]
    tps_cell = f"Map!${get_column_letter(map_info['tps_col'])}${map_info['row_for_15']}"
    adt = [
        ("More than 15 microservices with constant use → AKS", 15, f"={n_star}",
         f'=IF(ISNUMBER({n_star}),IF({n_star}>15,"Cost supports the threshold: ACA stays cheaper up to "&{n_star}-1,'
         f'"ACA costs more than AKS from "&{n_star}&" microservices, before 15"),"ACA is cheaper over the whole range")'),
        ("More than 10 h a day of constant use → AKS", 10, f"={hours_star}",
         f'=IF({hours_star}="","",IF({hours_star}>24,"With n_ref, ACA is cheaper even with the replica warm all day",'
         f'IF({hours_star}>10,"The 10 h threshold is conservative: the crossing is at "&TEXT({hours_star},"0.0")&" h",'
         f'"ACA costs more than AKS before 10 h: "&TEXT({hours_star},"0.0")&" h")))'),
        ("Running 24x7 costs more than AKS", "", f"={cells[rows[3][0]]}",
         f'=IF(ISNUMBER({cells[rows[3][0]]}),"With replicas warm 24x7, ACA exceeds AKS list from "&{cells[rows[3][0]]}'
         f'&" microservices","Even running 24x7, ACA is cheaper up to {MAX_MICROSERVICES} microservices")'),
        ("More than 100 tps of total load → AKS", 100, f"={tps_cell}",
         f'=IF({tps_cell}<0,"With 15 microservices and 100 tps in office hours, ACA is still cheaper","With 15 '
         f'microservices and 100 tps, AKS is cheaper")'),
    ]
    for label, threshold, result, reading in adt:
        sheet.append([label, threshold, result, reading])
        row = sheet.max_row
        for column in (1, 4):
            sheet.cell(row=row, column=column).alignment = WRAP
    sheet.cell(row=sheet.max_row, column=3).number_format = '"$"#,##0;[Red]-"$"#,##0'
    sheet.append([])
    sheet.append(["Other ADT_019 rules that are not about cost and rule out ACA even when it is cheaper: replicas above "
                  "4 vCPU / 8 GiB, batch or ETL processes, consumers that do not tolerate a cold start (up to ~30 s), "
                  "a dedicated environment (needs architecture approval)."])
    sheet.merge_cells(start_row=sheet.max_row, start_column=1, end_row=sheet.max_row, end_column=4)
    sheet.cell(row=sheet.max_row, column=1).alignment = WRAP
    sheet.row_dimensions[sheet.max_row].height = 45
    set_widths(sheet, [62, 16, 22, 70])
    return cells


def build_map(workbook, r):
    sheet = workbook.create_sheet("Map")
    title(sheet, "ACA (ADT schedule) − AKS list, USD/month: negative = ACA is cheaper",
          "Rows: microservices of the application. Columns: total requests of the application per month (editable), "
          "split evenly across microservices. 95,040,000 ≈ 100 tps sustained in office hours.", 9)
    sheet.append([])
    sheet.append(["Microservices \\ requests/month"] + REQUEST_LEVELS)
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


def build_clusters(workbook, clusters, cost, cost_months, k8s_by_id, from_snapshot, period):
    """PROD clusters with their category, and the totals per category the Summary reads."""
    sheet = workbook.create_sheet("Clusters")
    note = (f"Monthly cost: Cost Management (AmortizedCost) of the cluster + its node resource group, average of "
            f"{period}. Only APPLICATION clusters enter the analysis; the other categories are listed so the PROD "
            "total adds up. ")
    note += FALLBACK_NOTE if not from_snapshot else ("Category from cluster_snapshot.csv: platform export, Container "
                                                     "Insights and the app code.")
    title(sheet, "PROD AKS clusters by category", note, 7)
    sheet.append([])
    header(sheet, ["Category", "PROD clusters", "Monthly cost (USD)", "Microservices", "Meaning"])
    prod = [row for row in clusters.rows if row["environment"] == "PROD"]
    order = [category for category, _ in CATEGORIES]
    prod.sort(key=lambda row: (order.index(row["category"]) if row["category"] in order else len(order), row["name"]))
    totals_first = sheet.max_row + 1
    list_first = totals_first + len(CATEGORIES) + 3  # totals, total row, blank row, list header
    list_last = max(list_first, list_first + len(prod) - 1)
    span = {column: f"${column}${list_first}:${column}${list_last}" for column in "CEF"}
    refs = {}
    for category, meaning in CATEGORIES:
        sheet.append([category, f"=COUNTIF({span['C']},A{sheet.max_row + 1})",
                      f"=SUMIF({span['C']},A{sheet.max_row + 1},{span['F']})",
                      f"=SUMIF({span['C']},A{sheet.max_row + 1},{span['E']})", meaning])
        row = sheet.max_row
        sheet.cell(row=row, column=3).number_format = USD
        sheet.cell(row=row, column=5).alignment = WRAP
        refs[category] = (f"Clusters!$B${row}", f"Clusters!$C${row}")
    last_total = sheet.max_row
    sheet.append(["Total PROD"] + [f"=SUM({letter}{totals_first}:{letter}{last_total})" for letter in "BCD"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    sheet.cell(row=sheet.max_row, column=3).number_format = USD
    refs["total"] = (f"Clusters!$B${sheet.max_row}", f"Clusters!$C${sheet.max_row}")
    sheet.append([])
    header(sheet, ["Cluster", "App code", "Category", "Category reason", "Microservices", "Monthly cost (USD)", "Id"])
    assert sheet.max_row + 1 == list_first
    for row_data in prod:
        cluster_id = row_data["id"].lower()
        services = k8s_by_id.get(cluster_id, {}).get("microservices", "")
        sheet.append([row_data["name"], row_data.get("cod_app", ""), row_data["category"],
                      row_data.get("category_reason", ""), number(services)])
        row = sheet.max_row
        if cost:
            sheet.cell(row=row, column=6, value=f'=SUMIFS({cost.col("cost_usd")},{cost.col("owner_id")},$G{row},'
                                                f'{cost.col("owner_type")},"AKS")/{cost_months}').number_format = USD
        sheet.cell(row=row, column=7, value=row_data["id"])
        sheet.cell(row=row, column=4).alignment = WRAP
    set_widths(sheet, [34, 12, 20, 60, 14, 16, 40])
    sheet.freeze_panes = f"B{list_first}"
    return refs


def build_measured(workbook, clusters, usage, cost, workloads, services, months, period):
    sheet = workbook.create_sheet("Measured_AKS")
    title(sheet, "Measured AKS: cost, capacity and use per cluster",
          f"Cost: monthly average of Cost Management (AmortizedCost) for the cluster + its node resource group, "
          f"{period}. Use: office-hours average (Monday to Friday 8-20, Lima) of platform metrics. Each cluster serves "
          "a single application. The PROD summary below counts APPLICATION clusters only (column Q).", 13)
    sheet.append([])
    headers = ["Cluster", "Environment", "System nodes", "User nodes", "VM sizes", "Monthly cost (USD)",
               "Allocatable vCPU", "vCPU used, office", "CPU use, office", "Peak vCPU", "GiB used, office",
               "Workloads (no system)", "Services", "Mesh", "Id", "Configured minimum nodes", "Category"]
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
            cluster = f'{usage.col("cluster_id")},O{row}'
            sheet.cell(row=row, column=7, value=f'=SUMIFS({usage.col("cpu_allocatable_avg")},{cluster})'
                                                f'/{usage_months}').number_format = "0.0"
            sheet.cell(row=row, column=8, value=f'=SUMIFS({usage.col("cpu_used_office_avg")},{cluster})'
                                                f'/{usage_months}').number_format = "0.0"
            sheet.cell(row=row, column=9, value=f"=IFERROR(H{row}/G{row},\"\")").number_format = "0%"
            sheet.cell(row=row, column=10, value=f'=_xlfn.MAXIFS({usage.col("cpu_used_max")},{cluster})'
                                                 ).number_format = "0.0"
            sheet.cell(row=row, column=11, value=f'=SUMIFS({usage.col("mem_used_office_avg_gib")},{cluster})'
                                                 f'/{usage_months}').number_format = "0.0"
        if workloads:
            sheet.cell(row=row, column=12, value=f'=COUNTIFS(Workloads!$L:$L,O{row})')
        if services:
            sheet.cell(row=row, column=13, value=f'=COUNTIFS({services.col("cluster_id")},O{row})')
        sheet.cell(row=row, column=14, value=row_data.get("service_mesh", ""))
        sheet.cell(row=row, column=15, value=row_data.get("id", ""))
        sheet.cell(row=row, column=16, value=number(row_data.get("min_nodes", "")))
        sheet.cell(row=row, column=17, value=row_data.get("category", ""))
    last = sheet.max_row
    sheet.append([])
    sheet.append(["PROD summary (APPLICATION clusters)"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    prod = f'$B$5:$B${last},"PROD",$Q$5:$Q${last},{APP}'
    in_prod = f'($B$5:$B${last}="PROD")*($Q$5:$Q${last}={APP})'
    summary = [
        ("prod", "PROD application clusters", f"=COUNTIFS({prod})", "0"),
        ("non_prod", "Non-production clusters (Dev/QA)",
         f'=COUNTIFS($B$5:$B${last},"DESA")+COUNTIFS($B$5:$B${last},"CERT")', "0"),
        ("all", "AKS clusters in total (PROD + Dev/QA, every category)",
         f'=COUNTIFS($B$5:$B${last},"PROD")+COUNTIFS($B$5:$B${last},"DESA")+COUNTIFS($B$5:$B${last},"CERT")', "0"),
        ("le4", "PROD application clusters with 4 nodes or fewer",
         f'=SUMPRODUCT({in_prod}*(($C$5:$C${last}+$D$5:$D${last})<=4))', "0"),
        ("lt4_now", "PROD application clusters with fewer than 4 nodes at inventory time",
         f'=SUMPRODUCT({in_prod}*(($C$5:$C${last}+$D$5:$D${last})<4))', "0"),
        ("lt4_min", "PROD application clusters with a configured minimum below 4 nodes",
         f'=SUMPRODUCT({in_prod}*($P$5:$P${last}<>"")*($P$5:$P${last}<4))', "0"),
        ("cost_avg", "Average monthly cost per PROD application cluster (USD)",
         f'=IFERROR(AVERAGEIFS($F$5:$F${last},{prod}),"")', USD),
        ("cost_total", "Total monthly cost of PROD application clusters (USD)", f'=SUMIFS($F$5:$F${last},{prod})', USD),
        ("cpu", "Average CPU use in office hours, PROD application clusters",
         f'=IFERROR(SUMIFS($H$5:$H${last},{prod})/SUMIFS($G$5:$G${last},{prod}),"")', "0%"),
        ("cpu_unused", "Allocatable vCPU unused in office hours, PROD application clusters",
         f'=SUMIFS($G$5:$G${last},{prod})-SUMIFS($H$5:$H${last},{prod})', "0"),
    ]
    if usage:
        criteria = f'{usage.col("environment")},"PROD",{usage.col("category")},{APP}'

        def share(used, allocatable):
            return f'=IFERROR(SUMIFS({usage.col(used)},{criteria})/SUMIFS({usage.col(allocatable)},{criteria}),"")'

        # p95 and maximum are per pool and month; summed over clusters they are an upper bound,
        # because the peaks of different clusters do not happen at the same hour.
        summary += [
            ("mem", "Average memory use in office hours, PROD application clusters",
             share("mem_used_office_avg_gib", "mem_allocatable_avg_gib"), "0%"),
            ("cpu_p95", "CPU use, hourly 95th percentile, PROD application clusters",
             share("cpu_used_p95", "cpu_allocatable_avg"), "0%"),
            ("cpu_max", "CPU use, maximum (peak of each cluster added up), PROD application clusters",
             share("cpu_used_max", "cpu_allocatable_avg"), "0%"),
            ("mem_p95", "Memory use, hourly 95th percentile, PROD application clusters",
             share("mem_used_p95_gib", "mem_allocatable_avg_gib"), "0%"),
            ("mem_max", "Memory use, maximum (peak of each cluster added up), PROD application clusters",
             share("mem_used_max_gib", "mem_allocatable_avg_gib"), "0%"),
        ]
    refs = {}
    for key, label, formula, fmt in summary:
        sheet.append([label, formula])
        sheet.cell(row=sheet.max_row, column=2).number_format = fmt
        refs[key] = f"Measured_AKS!$B${sheet.max_row}"
    if cost:
        sheet.append([])
        sheet.append(["Existing Container Apps (average monthly cost per environment)"])
        sheet.cell(row=sheet.max_row, column=1).font = BOLD
        for env_name in sorted({row["owner"] for row in cost.rows if row["owner_type"] == "ACA"}):
            sheet.append([env_name, f'=SUMIFS({cost.col("cost_usd")},{cost.col("owner")},A{sheet.max_row + 1},'
                                    f'{cost.col("owner_type")},"ACA")/{cost_months}'])
            sheet.cell(row=sheet.max_row, column=2).number_format = USD
    set_widths(sheet, [34, 12, 10, 10, 22, 16, 12, 14, 12, 10, 12, 14, 10, 10, 40, 14, 18])
    sheet.freeze_panes = "B5"
    return refs, last


def build_traffic(workbook, traffic, apis, months):
    sheet = workbook.create_sheet("APIM_traffic")
    title(sheet, "Traffic per API in APIM (monthly average)",
          "North-south model: every call, even between microservices, goes through APIM, so these requests are the "
          "ones Container Apps would bill. Outside office hours: each run of 5-minute intervals with traffic wakes up a "
          "replica, which stays on for 300 s (cooldown) after the last request.", 13)
    sheet.append([])
    headers = ["APIM", "API", "Backend", "Requests/month", "Office/month", "Off-hours/month", "5-min peak",
               "Hour peak", "Average per office hour", "Peak / average", "Office share with traffic",
               "Active seconds off hours/month", "Idle seconds (cooldown)/month", "Wake-ups/month"]
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


def build_workloads(workbook, workloads, period):
    sheet = workbook.create_sheet("Workloads")
    title(sheet, "Workloads of PROD application clusters (Container Insights, office hours, no system namespaces)",
          f"Requests and use per replica: average of {period}. 'Fits ACA' applies the ADT_019 limit (4 vCPU / 8 GiB "
          "per replica). Only APPLICATION clusters.", 11)
    sheet.append([])
    headers = ["Cluster", "Environment", "Namespace", "Workload", "Kind", "Average replicas", "vCPU request per replica",
               "vCPU used per replica", "GiB request per replica", "GiB used per replica", "Fits ACA", "Cluster id"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    w = workloads
    keys = sorted({(row["cluster"], row["environment"], row["namespace"], row["workload"], row["kind"], row["cluster_id"])
                   for row in w.rows if row.get("system_namespace") == "NO" and row.get("period") == "OFFICE"
                   and row.get("environment") == "PROD" and row.get("category") == aa.APPLICATION})

    def avg(field, row):
        return (f'=IFERROR(AVERAGEIFS({w.col(field)},{w.col("cluster_id")},$L{row},{w.col("namespace")},$C{row},'
                f'{w.col("workload")},$D{row},{w.col("period")},"OFFICE"),"")')

    for key in keys:
        sheet.append(list(key[:5]))
        row = sheet.max_row
        for column, field, fmt in ((6, "replicas_avg", "0.0"), (7, "cpu_request_per_replica", "0.00"),
                                   (8, "cpu_used_per_replica", "0.00"), (9, "mem_request_gib_per_replica", "0.00"),
                                   (10, "mem_used_gib_per_replica", "0.00")):
            if field in w.fields:
                sheet.cell(row=row, column=column, value=avg(field, row)).number_format = fmt
        sheet.cell(row=row, column=11, value=f'=IF(OR(N(G{row})>4,N(I{row})>8),"NO","YES")')
        sheet.cell(row=row, column=12, value=key[5])
    set_widths(sheet, [30, 10, 24, 34, 12, 12, 14, 14, 14, 14, 10, 40])
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
    """'True'/'False' from Resource Graph become YES/NO, so COUNTIFS compares text and not Excel booleans."""
    for row in rows:
        for key, value in row.items():
            if value in ("True", "true"):
                row[key] = "YES"
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
    title(sheet, "Configuration dispersion across AKS clusters",
          "For each property: the most frequent value in PROD application clusters and how many of them have another "
          "value. It does not judge compliance: the guidelines team marks which value the LBS requires (column F, "
          "editable) and column G counts the clusters that do not meet it.", 7)
    sheet.append([])
    headers = ["Property", "Distinct values (all clusters)", "Most frequent PROD value", "PROD clusters with another value",
               "PROD clusters", "Value required by the LBS", "PROD clusters not compliant"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    c = clusters
    prod = [row for row in c.rows if row["environment"] == "PROD" and row.get("category") == aa.APPLICATION]
    base = f'{c.col("environment")},"PROD",{c.col("category")},{APP}'
    for prop in CONFIG_PROPERTIES:
        if prop not in c.fields:
            continue
        values = Counter(str(row.get(prop, "")) for row in prod)
        mode = values.most_common(1)[0][0] if values else ""
        distinct = Counter(str(row.get(prop, "")) for row in c.rows)
        sheet.append([prop, "; ".join(f"{value or '(empty)'}: {count}" for value, count in distinct.most_common()), mode])
        row = sheet.max_row
        criteria = f'"<>"&C{row}' if mode != "" else '"<>"'
        sheet.cell(row=row, column=4, value=f'=COUNTIFS({base},{c.col(prop)},{criteria})')
        sheet.cell(row=row, column=5, value=f'=COUNTIFS({base})')
        sheet.cell(row=row, column=6).fill = INPUT_FILL
        sheet.cell(row=row, column=7, value=f'=IF(F{row}="","",COUNTIFS({base},{c.col(prop)},"<>"&F{row}))')
        sheet.cell(row=row, column=2).alignment = WRAP
    set_widths(sheet, [26, 60, 22, 14, 10, 18, 14])


HPA_FINDINGS = [
    ("HPA that cannot scale (min = max)", "fixed_min_equals_max", "YES", "Minimum and maximum are equal in the export"),
    ("HPA with a minimum of 2 or more replicas", "min_two_or_more", "YES",
     "Capacity paid 24x7; on Container Apps the minimum outside office hours is 0"),
    ("HPA at its maximum on the export day", "at_max_now", "YES", "Current replicas = HPA maximum"),
    ("HPA that stays at its minimum in the period", "behaviour_ci", "NEVER_SCALES",
     "Office average of replicas (Container Insights) ≤ 110% of the HPA minimum"),
    ("HPA near its maximum in the period", "behaviour_ci", "AT_MAX",
     "Office average of replicas (Container Insights) ≥ 90% of the HPA maximum"),
    ("HPA without Container Insights history", "ci_matched", "NO",
     "Cluster without the add-on, or workload without pods in the period"),
]
DEPLOYMENT_FINDINGS = [
    ("Deployments without HPA", "has_hpa", "NO", "No row in the HPA export"),
    ("Deployments without pods in the period (possible zombies)", "no_usage_in_ci", "YES",
     "The cluster has Container Insights but the workload does not show"),
    ("CPU request more than 3 times the actual use", "request_over_3x_used", "YES",
     "Export request / vCPU used per replica in office hours"),
    ("Actual CPU use above the request", "used_over_request", "YES",
     "The scheduler reserves less than it uses: risk of contention on the node"),
    ("CPU limit 10 times the request or more", "limit_10x_request", "YES", "Limit / request in the export"),
    ("Fits a Container Apps replica by its request", "fits_aca_by_request", "YES",
     "Request ≤ 4 vCPU and ≤ 8 GiB, the ADT_019 maximum per replica"),
    ("Request above half of a node of its cluster", "pod_too_big_for_node", "YES",
     "CPU or memory request above half of the allocatable of a user-pool node (approximate)"),
]


def prod_microservices(detail):
    """COUNTIFS criteria for PROD microservices of APPLICATION clusters in a snapshot detail sheet."""
    return [detail.col("environment"), '"PROD"', detail.col("category"), APP, detail.col("microservice"), '"YES"']


def build_k8s(workbook, hpa, deployments, taken_on, period):
    sheet = workbook.create_sheet("Config_K8s")
    title(sheet, f"PROD HPAs and deployments: configuration on {taken_on} against {period}",
          "The configuration is a one-day snapshot (platform team export, no history). The replica and usage history "
          "comes from Container Insights. Microservices of APPLICATION clusters only. Row detail in D_HPA and D_Deploy.", 5)
    sheet.append([])
    headers = ["Finding", "PROD", "PROD total", "%", "How it is calculated"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    for detail, findings in ((hpa, HPA_FINDINGS), (deployments, DEPLOYMENT_FINDINGS)):
        if detail is None:
            continue
        base = ",".join(prod_microservices(detail))
        for label, field, value, how in findings:
            if field not in detail.fields:
                continue
            sheet.append([label])
            row = sheet.max_row
            sheet.cell(row=row, column=2, value=f'=COUNTIFS({base},{detail.col(field)},"{value}")')
            sheet.cell(row=row, column=3, value=f'=COUNTIFS({base})')
            sheet.cell(row=row, column=4, value=f'=IFERROR(B{row}/C{row},"")').number_format = "0%"
            sheet.cell(row=row, column=5, value=how).alignment = WRAP
        sheet.append([])
    set_widths(sheet, [56, 10, 12, 8, 70])


MICROSERVICE_BUCKETS = [("0 (no microservice in the export)", 0, 0), ("1", 1, 1), ("2 to 3", 2, 3), ("4 to 5", 4, 5),
                        ("6 to 10", 6, 10), ("11 to 15", 11, 15), ("16 to 30", 16, 30), ("31 to 100", 31, 100),
                        ("More than 100", 101, 10 ** 9)]
REPLICA_BUCKETS = [("0 (possible zombie)", 0, 0), ("1", 1, 1), ("2 to 3", 2, 3), ("4 to 5", 4, 5), ("6 to 10", 6, 10),
                   ("11 to 20", 11, 20), ("21 to 50", 21, 50), ("More than 50", 51, 10 ** 9)]
LAYER_ORDER = ["business", "channel", "app", "async", "support", "core", "other"]
BEHAVIOURS = [("NEVER_SCALES", "Stays at its minimum"), ("SCALES", "Scales between minimum and maximum"),
              ("AT_MAX", "Reaches or nears its maximum"), ("FIXED", "Minimum = maximum"), ("NO_DATA", "No data")]
LAST_SCALE_BUCKETS = [("UNDER_1_MONTH", "Less than 1 month"), ("1_TO_3_MONTHS", "1 to 3 months"),
                      ("3_TO_12_MONTHS", "3 to 12 months"), ("OVER_1_YEAR", "More than 1 year"),
                      ("NO_DATA", "No date in the export")]
NO_SCALE_CAUSES = [("FIXED_MIN_MAX", "minimum = maximum, so the date is when the HPA was set"),
                   ("MIN_ABOVE_NEED", "the minimum is above what its CPU needs"),
                   ("LOAD_BELOW_TARGET", "the load never reaches the HPA target"),
                   ("SCALES", "it scales in the period"), ("NO_DATA", "no Container Insights history")]


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


def build_last_scale(sheet, hpa, prod_hpa, period):
    """Time since the last scale event (rows) by cause of not scaling (columns), and the HPAs whose office maximum
    reached their maximum."""
    if "last_scale_bucket" in hpa.fields and "no_scale_cause" in hpa.fields:
        sheet.append([])
        header(sheet, ["Time since the last scale event (export day)"] + [code for code, _ in NO_SCALE_CAUSES]
               + ["Total", "% of HPAs"])
        first = sheet.max_row + 1
        last = first + len(LAST_SCALE_BUCKETS) - 1
        total = get_column_letter(len(NO_SCALE_CAUSES) + 2)
        rows = {}
        for code, label in LAST_SCALE_BUCKETS:
            sheet.append([f"{label} ({code})"])
            row = sheet.max_row
            rows[code] = row
            for column, (cause, _) in enumerate(NO_SCALE_CAUSES, 2):
                sheet.cell(row=row, column=column, value=f'=COUNTIFS({prod_hpa},{hpa.col("last_scale_bucket")},"{code}",'
                                                         f'{hpa.col("no_scale_cause")},"{cause}")')
            sheet.cell(row=row, column=len(NO_SCALE_CAUSES) + 2,
                       value=f"=SUM(B{row}:{get_column_letter(len(NO_SCALE_CAUSES) + 1)}{row})")
            sheet.cell(row=row, column=len(NO_SCALE_CAUSES) + 3,
                       value=f'=IFERROR({total}{row}/SUM(${total}${first}:${total}${last}),"")').number_format = "0%"
        all_hpas = f"SUM(${total}${first}:${total}${last})"
        for label, codes, how in (
                ("HPAs without a scale event for more than 3 months", ("3_TO_12_MONTHS", "OVER_1_YEAR"),
                 "3 to 12 months + more than 1 year, from lastScaleTime to the export day"),
                ("HPAs without a scale event for more than 1 year", ("OVER_1_YEAR",),
                 "More than 1 year, from lastScaleTime to the export day")):
            sheet.append([label, "=" + "+".join(f"{total}{rows[code]}" for code in codes)])
            row = sheet.max_row
            sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/{all_hpas},"")').number_format = "0%"
            sheet.cell(row=row, column=5, value=how).alignment = WRAP
        sheet.append(["Causes: " + "; ".join(f"{code} = {meaning}" for code, meaning in NO_SCALE_CAUSES)
                      + ". With minimum = maximum the last scale time is when the HPA was set, not a scaling event."])
    if "office_max_reached_hpa_max" in hpa.fields:
        sheet.append([f"HPAs whose office maximum in {period} reached hpa_max",
                      f'=COUNTIFS({prod_hpa},{hpa.col("office_max_reached_hpa_max")},"YES")'])
        row = sheet.max_row
        sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/COUNTIFS({prod_hpa}),"")').number_format = "0%"
        sheet.cell(row=row, column=5, value="Hourly maximum of replicas in office hours (Container Insights) ≥ HPA "
                                            "maximum, HPAs with minimum < maximum. A rollout briefly adds the new pods "
                                            "to the old ones, so it is an upper bound.").alignment = WRAP


def build_patterns(workbook, deploy, hpa, kclusters, capabilities, taken_on, period, cost_period):
    sheet = workbook.create_sheet("Patterns")
    title(sheet, f"AKS usage patterns in PROD application clusters (configuration on {taken_on})",
          "Microservice = deployment outside platform namespaces and not a known tool. Domain = namespace. Layer = first "
          f"word of the name. Replicas and HPA from a one-day export; the history of {period} comes from Container "
          "Insights. Only APPLICATION clusters.", 5)
    prod_ms = prod_microservices(deploy)

    section(sheet, "Pattern 1. Microservice granularity")
    header(sheet, ["Figure", "Value", "% of total", "", "How it is calculated"])
    sheet.append(["PROD microservices", f"=COUNTIFS({','.join(prod_ms)})", "", "",
                  "D_Deploy rows in PROD with microservice = YES in APPLICATION clusters"])
    total_row = sheet.max_row
    for layer in LAYER_ORDER:
        sheet.append([f"Layer {layer}-*" if layer != "other" else "Other layers"])
        row = sheet.max_row
        sheet.cell(row=row, column=2, value=f'=COUNTIFS({",".join(prod_ms)},{deploy.col("layer")},"{layer}")')
        sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/$B${total_row},"")').number_format = "0%"
    if capabilities:
        sheet.append([])
        header(sheet, ["Capability repeated in several domains", "PROD domains", "PROD deployments", "",
                       "Domains (deployments)"])
        for row in capabilities:
            sheet.append([row["capability"], number(row["namespaces_prod"]), number(row["deployments_prod"]), "",
                          row["namespaces"]])
            sheet.cell(row=sheet.max_row, column=5).alignment = WRAP

    section(sheet, "Pattern 2. Clusters by number of microservices")
    header(sheet, ["Microservices per cluster", "PROD clusters", "% of clusters", "Monthly cost (USD)", "How it is calculated"])
    k = kclusters
    base = [k.col("environment"), '"PROD"', k.col("category"), APP, k.col("in_inventory"), '"YES"']
    first = sheet.max_row + 1
    for label, low, high in MICROSERVICE_BUCKETS:
        criteria = ",".join(base + [k.col("microservices"), f'">={low}"', k.col("microservices"), f'"<={high}"'])
        sheet.append([label, f"=COUNTIFS({criteria})"])
        row = sheet.max_row
        sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${first + len(MICROSERVICE_BUCKETS) - 1}),"")'
                   ).number_format = "0%"
        sheet.cell(row=row, column=4, value=f"=SUMIFS({k.col('cost_month_usd')},{criteria})").number_format = USD
    last = sheet.max_row
    # An APPLICATION cluster can show 0 microservices when it only runs the application's own gateway.
    small = base + [k.col("microservices"), '">=1"']
    sheet.append(["PROD clusters with 15 microservices or fewer",
                  f'=COUNTIFS({",".join(small + [k.col("microservices"), chr(34) + "<=15" + chr(34)])})',
                  f'=IFERROR(B{last + 1}/SUM($B${first + 1}:$B${last}),"")', "",
                  "ADT_019 threshold of 15 microservices; over the clusters with at least 1 microservice in the export"])
    sheet.cell(row=sheet.max_row, column=3).number_format = "0%"
    sheet.append(["Monthly cost of PROD clusters with 3 microservices or fewer", "", "",
                  f'=SUMIFS({k.col("cost_month_usd")},{",".join(small + [k.col("microservices"), chr(34) + "<=3" + chr(34)])})',
                  f"Actual amortized cost (Cost Management), {cost_period}; 1 to 3 microservices in the export"])
    sheet.cell(row=sheet.max_row, column=4).number_format = USD
    sheet.append(["Clusters without applications are not a pattern here: they are the NO_APPLICATIONS category of "
                  "the Clusters sheet."])

    if hpa:
        section(sheet, "Pattern 4. Replicas and HPA")
        header(sheet, ["Current replicas (export day)", "PROD HPAs", "% of HPAs", "", "How it is calculated"])
        prod_hpa = ",".join(prod_microservices(hpa))
        first = sheet.max_row + 1
        for label, low, high in REPLICA_BUCKETS:
            criteria = f'{prod_hpa},{hpa.col("replicas_now")},">={low}",{hpa.col("replicas_now")},"<={high}"'
            sheet.append([label, f"=COUNTIFS({criteria})"])
            row = sheet.max_row
            sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${first + len(REPLICA_BUCKETS) - 1}),"")'
                       ).number_format = "0%"
        sheet.append([])
        header(sheet, ["HPA behaviour", "By ValTrend (maximum, a few days)", "%",
                       f"By Container Insights (office average, {period})", "%"])
        first = sheet.max_row + 1
        end = first + len(BEHAVIOURS) - 1
        rows = {}
        for code, label in BEHAVIOURS:
            sheet.append([label])
            row = sheet.max_row
            rows[code] = row
            sheet.cell(row=row, column=2, value=f'=COUNTIFS({prod_hpa},{hpa.col("behaviour_export")},"{code}")')
            sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/SUM($B${first}:$B${end}),"")').number_format = "0%"
            sheet.cell(row=row, column=4, value=f'=COUNTIFS({prod_hpa},{hpa.col("behaviour_ci")},"{code}")')
            sheet.cell(row=row, column=5, value=f'=IFERROR(D{row}/SUM($D${first}:$D${end}),"")').number_format = "0%"
        never, fixed = rows["NEVER_SCALES"], rows["FIXED"]
        sheet.append(["Does not scale (stays at its minimum + minimum = maximum)"]
                     + [f"={letter}{never}+{letter}{fixed}" for letter in "BCDE"])
        sheet.cell(row=sheet.max_row, column=1).font = BOLD
        for column in (3, 5):
            sheet.cell(row=sheet.max_row, column=column).number_format = "0%"
        if "min_above_need" in hpa.fields:
            at_min = f'{prod_hpa},{hpa.col("behaviour_ci")},"NEVER_SCALES"'
            sheet.append([])
            header(sheet, ["Of the HPAs that stay at their minimum", "HPAs", "% of those that stay at their minimum", "",
                           "How it is calculated"])
            for label, field, how in (
                    ("Its minimum is above the replicas its CPU asks for", "min_above_need",
                     "Replicas needed = CPU used in office hours / (request per replica × 70%)"),
                    ("Its minimum is above that even keeping 2 replicas for high availability", "min_above_need_beyond_ha",
                     "Minimum greater than the larger of 2 and the replicas needed")):
                sheet.append([label, f'=COUNTIFS({at_min},{hpa.col(field)},"YES")'])
                row = sheet.max_row
                sheet.cell(row=row, column=3, value=f'=IFERROR(B{row}/COUNTIFS({at_min}),"")').number_format = "0%"
                sheet.cell(row=row, column=5, value=how).alignment = WRAP
        build_last_scale(sheet, hpa, prod_hpa, period)
        sheet.append(["Container Insights: office average of replicas; at its minimum = up to 110% of the minimum, near "
                      "the maximum = from 90% of the maximum. The hourly maximum is not used because every rollout adds "
                      "the new pods to the old ones for a few minutes."])
    set_widths(sheet, [52, 16, 14, 22, 40, 16, 12, 12])


LEARN = "https://learn.microsoft.com/azure/"
ADT = "ADT_019 (Confluence)"
NOTE_FONT = Font(name=FONT, color="C00000", italic=True)


def build_criteria(workbook, clusters, kclusters, deploy, hpa, traffic, workloads_last, traffic_month, period):
    """Technical criteria to choose a platform, each with the measured PROD figure, its source and any caveat."""
    sheet = workbook.create_sheet("Technical_criteria")
    title(sheet, "AKS vs Container Apps technical criteria, with the figure measured in PROD",
          "Each row: what favours each platform, how much of the current estate falls on each side and where the "
          "criterion comes from. PROD figures cover APPLICATION clusters only. Notes in red are debatable points. "
          "Microsoft Learn checked on 2026-11-11.", 8)
    sheet.append([])
    header(sheet, ["Criterion", "Favours Container Apps", "Favours AKS", "Measured figure (PROD)", "Value", "Out of",
                   "Source", "Note"])

    def add(criterion, aca, aks, measure, value="", total="", source="", note=""):
        sheet.append([criterion, aca, aks, measure, value, total, source, note])
        row = sheet.max_row
        for column in (1, 2, 3, 4, 7, 8):
            sheet.cell(row=row, column=column).alignment = WRAP
        sheet.cell(row=row, column=8).font = NOTE_FONT
        return row

    k_base = (f'{kclusters.col("environment")},"PROD",{kclusters.col("category")},{APP},'
              f'{kclusters.col("in_inventory")},"YES"') if kclusters else ""
    add("Microservices per application", "15 or fewer", "More than 15, or a platform shared by several apps",
        "Clusters with 15 microservices or fewer",
        f'=COUNTIFS({k_base},{kclusters.col("microservices")},">=1",{kclusters.col("microservices")},"<=15")'
        if kclusters else "",
        f'=COUNTIFS({k_base},{kclusters.col("microservices")},">=1")' if kclusters else "", ADT,
        "The threshold of 15 is the ADT's; Microsoft does not publish one.")
    if traffic and traffic_month:
        t = traffic
        month = f'{t.col("month")},"{traffic_month}"'
        add("Transactions per second and traffic pattern", "Fewer than 100 tps, or traffic with peaks and valleys",
            "More than 100 tps sustained", f"APIs with a peak above 100 tps in {period_text([traffic_month])}",
            f'=COUNTIFS({month},{t.col("peak_5min_requests")},">30000")', f"=COUNTIFS({month})",
            f"{ADT}; APIM metrics (Requests)",
            "Peak per API averaged over 5 min: the real peak per second is somewhat higher. It is per API, not per application.")
        row = add("Time-of-day pattern", "Use concentrated in office hours; scales to 0 outside them",
                  "Constant 24x7 use (more than 10 h a day)", f"Requests outside office hours / total, "
                  f"{period_text([traffic_month])}",
                  f'=IFERROR(SUMIFS({t.col("offhours_requests")},{month})/SUMIFS({t.col("requests")},{month}),"")', "",
                  f"{ADT}; {LEARN}container-apps/billing")
        sheet.cell(row=row, column=5).number_format = "0%"
    if hpa:
        h = hpa
        base = ",".join(prod_microservices(h))
        add("Scaling and scale to zero", "Scales to 0 and by events natively; up to 1,000 replicas per revision",
            "High and constant pod density; today PROD does not scale to 0 (see the KEDA scenario below)",
            f"HPAs that do not scale in {period} (stay at their minimum or minimum = maximum)",
            f'=COUNTIFS({base},{h.col("behaviour_ci")},"NEVER_SCALES")+COUNTIFS({base},{h.col("behaviour_ci")},"FIXED")',
            f"=COUNTIFS({base})", f"{LEARN}container-apps/scale-app")
    if deploy:
        d = deploy
        base = ",".join(prod_microservices(d))
        add("Size per replica", "Up to 4 vCPU / 8 GiB", "Larger replicas or special VMs (memory, GPU, network)",
            "Microservices that fit by their request", f'=COUNTIFS({base},{d.col("fits_aca_by_request")},"YES")',
            f"=COUNTIFS({base})", f"{ADT}; {LEARN}container-apps/containers",
            "In consumption-only environments the maximum drops to 2 vCPU / 4 GiB; the bank's environments use workload profiles.")
    add("Workload type and architecture", "HTTP, APIs, events and short jobs",
        "StatefulSets, own DaemonSets, long batch, privileged containers", "StatefulSet workloads in application namespaces",
        '=COUNTIFS(Workloads!$B:$B,"PROD",Workloads!$E:$E,"StatefulSet")' if workloads_last else "",
        '=COUNTIFS(Workloads!$B:$B,"PROD")' if workloads_last else "",
        f"{LEARN}container-apps/compare-options; {LEARN}container-apps/containers",
        "Container Apps does not allow privileged containers nor images that are not Linux x86-64.")
    if clusters:
        c = clusters
        base = f'{c.col("environment")},"PROD",{c.col("category")},{APP}'
        add("Access to the Kubernetes API and control of the platform",
            "Needs no kubectl, CRDs or own operators (the API is hidden)",
            "Service mesh, operators, CRDs and own network policies", "Clusters with managed Istio",
            f'=COUNTIFS({base},{c.col("service_mesh")},"Istio")', f"=COUNTIFS({base})",
            f"{LEARN}container-apps/compare-options; {LEARN}aks/istio-about",
            "Only Istio managed by Azure counts; an Istio installed by the app does not show.")
        add("", "", "", "Clusters with GitOps (Flux)", f'=COUNTIFS({base},{c.col("gitops_flux")},"YES")',
            f"=COUNTIFS({base})", f"{LEARN}azure-arc/kubernetes/conceptual-gitops-flux2")
        add("Learning curve and operating effort", "The platform updates runtime, nodes and patches",
            "Version and node upgrades are on each team; needs Kubernetes experience",
            "Clusters without automatic version upgrade", f'=COUNTIFS({base},{c.col("upgrade_channel")},"none")',
            f"=COUNTIFS({base})", f"{LEARN}aks/auto-upgrade-cluster; {LEARN}aks/istio-about",
            "The Istio add-on does not upgrade itself: each revision is moved by hand. Person-hours are not quantified.")
    add("Scale limits", "Consumption cores per environment (quota raised on request); 1,000 replicas per revision",
        "5,000 nodes per cluster, 250 pods per node, 100 node pools", "No measured figure: documented limit", "", "",
        f"{LEARN}container-apps/quotas; {LEARN}aks/quotas-skus-regions",
        "The Container Apps quotas page (2026-10-09) does not publish a cap of 100 or 200 apps per environment.")
    add("Cost", "Pay per active second", "Reserved nodes with high and constant use", "See the Model, Break_even and Map sheets",
        "", "", "https://azure.microsoft.com/pricing/details/container-apps/; Azure Retail Prices API; Cost Management")
    sheet.append([])
    sheet.append(["Separate scenario: AKS with KEDA (not in PROD, not counted in the table above)"])
    for column in range(1, 9):
        sheet.cell(row=sheet.max_row, column=column).fill = SECTION_FILL
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    add("What would change", "", "With the KEDA add-on, AKS scales pods to 0 by events; the cluster's minimum nodes are "
        "still paid, so the saving depends on the node pool going down too", "", "", "",
        f"{LEARN}aks/keda-about")
    if clusters:
        c = clusters
        base = f'{c.col("environment")},"PROD",{c.col("category")},{APP}'
        add("Status at the bank", "", "KEDA initiative with a predictive autoscaler; not in PROD because Microsoft "
            "cannot guarantee the capacity", "PROD clusters with KEDA", f'=COUNTIFS({base},{c.col("keda")},"YES")',
            f"=COUNTIFS({base})", "Platform team",
            "Figure from the platform team; confirm the status of the initiative before presenting it.")
    if kclusters and "bia" in kclusters.fields:
        sheet.append([])
        sheet.append(["Separate, for information: business criticality (BIA). Not part of the technical criteria"])
        for column in range(1, 9):
            sheet.cell(row=sheet.max_row, column=column).fill = SECTION_FILL
        sheet.cell(row=sheet.max_row, column=1).font = BOLD
        add("Criticality (BIA)", "Medium or low BIA", "High or very high BIA", "Clusters with high or very high BIA",
            f'=COUNTIFS({k_base},{kclusters.col("bia")},"MUY ALTA")+COUNTIFS({k_base},{kclusters.col("bia")},"ALTA")',
            f"=COUNTIFS({k_base})", "Deployment export (bia); BIA_candidates sheet",
            "For information only; security and architecture decide which BIA levels require AKS.")
    set_widths(sheet, [24, 34, 34, 36, 10, 12, 40, 44])


BIA_LEVELS = ["MUY ALTA", "ALTA", "MEDIA", "BAJA", "MUY BAJA", "SIN BIA"]
BIA_REQUIRES_AKS = ("MUY ALTA", "ALTA")


def build_bia(workbook, kclusters):
    sheet = workbook.create_sheet("BIA_candidates")
    title(sheet, "Information annex: PROD application clusters that fit Container Apps, by business criticality (BIA)",
          "Technical fit = 1 to 15 microservices, every one fits a Container Apps replica by its request (4 vCPU / 8 GiB) "
          "and no application StatefulSet. Column B (editable) says which BIA levels stay on AKS for compliance. It "
          "does not include requests per second per application. Detail in D_K8s_clusters.", 8)
    sheet.append([])
    header(sheet, ["BIA", "Stays on AKS because of BIA?", "PROD clusters", "Technical fit for ACA",
                   "Current monthly cost of the fits (USD)", "Clusters that would move to ACA",
                   "Microservices that would move to ACA", "Fits that stay on AKS because of BIA"])
    k = kclusters
    base = f'{k.col("environment")},"PROD",{k.col("category")},{APP},{k.col("in_inventory")},"YES"'
    present = {row["bia"] for row in k.rows}
    levels = [level for level in BIA_LEVELS if level in present] + sorted(present - set(BIA_LEVELS))
    first = sheet.max_row + 1
    for level in levels:
        sheet.append([level, "YES" if level in BIA_REQUIRES_AKS else "NO"])
        row = sheet.max_row
        sheet.cell(row=row, column=2).fill = INPUT_FILL
        by_level = f'{base},{k.col("bia")},$A{row}'
        fit = f'{by_level},{k.col("aca_technical_fit")},"YES"'
        sheet.cell(row=row, column=3, value=f"=COUNTIFS({by_level})")
        sheet.cell(row=row, column=4, value=f"=COUNTIFS({fit})")
        sheet.cell(row=row, column=5, value=f"=SUMIFS({k.col('cost_month_usd')},{fit})").number_format = USD
        sheet.cell(row=row, column=6, value=f'=IF($B{row}="YES",0,D{row})')
        sheet.cell(row=row, column=7, value=f'=IF($B{row}="YES",0,SUMIFS({k.col("microservices")},{fit}))')
        sheet.cell(row=row, column=8, value=f'=IF($B{row}="YES",D{row},0)')
    last = sheet.max_row
    sheet.append(["Total", ""] + [f"=SUM({get_column_letter(column)}{first}:{get_column_letter(column)}{last})"
                                  for column in range(3, 9)])
    total = sheet.max_row
    sheet.cell(row=total, column=1).font = BOLD
    sheet.cell(row=total, column=5).number_format = USD
    sheet.append([])
    sheet.append(["Current monthly cost of the clusters that would move to ACA (USD)", "", "", "",
                  f'=SUMPRODUCT((B{first}:B{last}<>"YES")*E{first}:E{last})'])
    sheet.cell(row=sheet.max_row, column=5).number_format = USD
    set_widths(sheet, [16, 16, 12, 16, 20, 16, 18, 18])


SENSITIVITY_RPS = [5, 10, 25, 50, 100]
SENSITIVITY_REQUESTS = [1_000_000, 10_000_000, 50_000_000, 200_000_000]


def build_sensitivity(workbook, r):
    """How much the requests-per-replica assumption moves the result, for the reference case (n_ref)."""
    sheet = workbook["Break_even"]
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    n = s["n_ref"]
    sheet.append([])
    sheet.append(["Sensitivity: requests per second one replica serves (reference case: n_ref microservices)"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    sheet.append(["The Assumptions value has no load test behind it. Column A is editable. With few requests it barely "
                  "matters; with many, more replicas make both ACA and AKS more expensive."])
    for label, kind in (("ACA ADT schedule (USD/month)", "aca"), ("AKS effective (USD/month)", "aks")):
        header(sheet, [label] + [f"{level / 1_000_000:g} M requests/month per microservice" for level in SENSITIVITY_REQUESTS])
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


PROFILE_CLASSES = [("OFFICE", "90% or more of the requests in office hours"),
                   ("EXTENDED", "90% or more in the extended window, every day"),
                   ("24x7", "Traffic spread over day and night"), ("NO_TRAFFIC", "No requests in the period")]


def build_traffic_profile(workbook, profile, traffic, period):
    """Classify each API by when its traffic happens; windows and threshold are editable cells."""
    d = Detail(workbook, "D_Hourly_profile", profile, ["apim", "api_id", "day_type", "hour", "avg_requests_per_hour"])
    sheet = workbook.create_sheet("Traffic_profile")
    title(sheet, "Proposal: traffic profile per API (office, extended hours or 24x7)",
          f"Average requests per hour in APIM (working days and weekends, Lima time) over {period}. Yellow cells are "
          "editable: changing them reclassifies every API without querying Azure again.", 7)
    sheet.append([])
    inputs = [("Office: start hour (Monday to Friday)", 8), ("Office: end hour", 20),
              ("Extended: start hour (every day)", 6), ("Extended: end hour", 22),
              ("Classification threshold", 0.9)]
    cells = {}
    for label, value in inputs:
        sheet.append([label, value])
        sheet.cell(row=sheet.max_row, column=2).fill = INPUT_FILL
        cells[label] = f"$B${sheet.max_row}"
    sheet.cell(row=sheet.max_row, column=2).number_format = "0%"
    office_start, office_end, ext_start, ext_end, threshold = cells.values()
    apis = sorted({(row["apim"], row["api_id"]) for row in profile})
    sheet.append([])
    header(sheet, ["Class", "APIs", "% of requests", "", "Rule"])
    summary_first = sheet.max_row + 1
    table_header = summary_first + len(PROFILE_CLASSES) + 1
    first, last = table_header + 1, table_header + len(apis)
    for code, meaning in PROFILE_CLASSES:
        sheet.append([code, f'=COUNTIF($F${first}:$F${last},A{sheet.max_row + 1})',
                      f'=IFERROR(SUMIF($F${first}:$F${last},A{sheet.max_row + 1},$C${first}:$C${last})/'
                      f'SUM($C${first}:$C${last}),"")', "", meaning])
        sheet.cell(row=sheet.max_row, column=3).number_format = "0%"
    sheet.append([])
    header(sheet, ["APIM", "API", "Requests per week", "% in office hours", "% in the extended window", "Class",
                   "Peak tps (5 min)"])
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
        sheet.cell(row=row, column=6, value=f'=IF(N(C{row})=0,"NO_TRAFFIC",IF(D{row}>={threshold},"OFFICE",'
                                            f'IF(E{row}>={threshold},"EXTENDED","24x7")))')
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


def hpa_replica_factors(hpas):
    """Share of today's replicas each workload keeps if its inflated HPA minimum drops to what its CPU asks for.

    Only HPAs that stay at their minimum and whose minimum is above max(2, replicas needed by CPU) change; the rest
    keep 1. Keyed by (cluster id, namespace, deployment), lower case.
    """
    factors = {}
    for row in hpas or []:
        if row.get("environment") != "PROD" or row.get("behaviour_ci") != "NEVER_SCALES" \
                or row.get("min_above_need_beyond_ha") != "YES":
            continue
        try:
            keep = max(2, float(row["replicas_needed_by_cpu"])) / float(row["ci_replicas_avg_office"])
        except (KeyError, TypeError, ValueError, ZeroDivisionError):
            continue
        key = (str(row.get("cluster_id") or "").lower(), row["namespace"].lower(), row["deployment"].lower())
        factors[key] = round(min(1.0, keep), 3)
    return factors


def aca_workloads(workloads, ids, factors=None):
    """One row per PROD application workload: replica-hours per month (office / off-hours) and ACA size.

    Replica-hours use the full months only (with three or more months, the first and last of the window are partial).
    """
    rows = [row for row in workloads if row["environment"] == "PROD" and row["system_namespace"] == "NO"
            and row["kind"] == "ReplicaSet" and row["cluster_id"] in ids]
    months = sorted({row["month"] for row in rows})
    full = set(months[1:-1]) if len(months) >= 3 else set(months)
    grouped = defaultdict(list)
    for row in rows:
        grouped[(row["cluster_id"], row["namespace"], row["workload"])].append(row)
    out = []
    for (cluster_id, namespace, workload), items in sorted(grouped.items(), key=lambda item: (item[1][0]["cluster"], item[0])):
        def replica_hours(period):
            total = sum(float(item["replicas_avg"] or 0) * float(item["hours"] or 0)
                        for item in items if item["period"] == period and item["month"] in full)
            return round(total / max(1, len(full)), 1)

        def weighted(field):
            pairs = [(float(item.get(field) or 0), float(item["hours"] or 0)) for item in items]
            hours = sum(weight for _, weight in pairs)
            return sum(value * weight for value, weight in pairs) / hours if hours else 0

        cpu_used, mem_used = weighted("cpu_used_per_replica"), weighted("mem_used_gib_per_replica")
        office_items = [item for item in items if item["period"] == "OFFICE" and item["month"] in full] or \
            [item for item in items if item["period"] == "OFFICE"]
        office_hours = sum(float(item["hours"] or 0) for item in office_items)
        replicas_office = (sum(float(item["replicas_avg"] or 0) * float(item["hours"] or 0) for item in office_items)
                           / office_hours) if office_hours else 0
        cpu = max(weighted("cpu_request_per_replica"), cpu_used)
        mem = max(weighted("mem_request_gib_per_replica"), mem_used)
        vcpu = aca_allocation(cpu, mem)
        # Sized by use: average use plus 50% headroom, which is what a migration would request.
        vcpu_use = min(vcpu, aca_allocation(cpu_used * USE_HEADROOM, mem_used * USE_HEADROOM))
        out.append({"cluster": items[0]["cluster"], "cluster_id": cluster_id, "namespace": namespace, "workload": workload,
                    "office_replica_hours_month": replica_hours("OFFICE"),
                    "offhours_replica_hours_month": replica_hours("OFFHOURS"),
                    "cpu_request_or_use": round(cpu, 3), "mem_gib_request_or_use": round(mem, 3),
                    "cpu_used": round(cpu_used, 3), "mem_gib_used": round(mem_used, 3),
                    "replicas_office": round(replicas_office, 2),
                    "cpu_used_office_total": round(replicas_office * cpu_used, 3),
                    "mem_gib_used_office_total": round(replicas_office * mem_used, 3),
                    "hpa_replica_factor": (factors or {}).get((cluster_id, namespace.lower(), workload.lower()), 1.0),
                    "aca_vcpu": vcpu, "aca_mem_gib": vcpu * 2, "aca_vcpu_by_use": vcpu_use,
                    "aca_mem_gib_by_use": vcpu_use * 2, "fits_aca": "YES" if vcpu_use <= ACA_MAX_VCPU else "NO"})
    return out, sorted(full)


def add_right_sizing(sheet, row, s, d, usage, cost, cost_month, suffix="", columns=(14, 15, 16, 17)):
    """Columns M to Q of ACA_by_cluster (R to U with suffix "_hpa", where inflated HPA minimums also drop): nodes
    and cost of the cluster if its pods asked for their use plus headroom.

    Nodes keep the cluster's current size (allocatable per node from the platform metrics of the cost month) and the
    floor of Assumptions. Node VMs and disks scale with the nodes; network, load balancer and tier stay fixed.
    """
    by = f'{d.col("cluster_id")},$L{row}'
    month = f'{usage.col("month")},"{cost_month}"'
    cluster = f'{usage.col("cluster_id")},$L{row}'
    nodes_now = f'SUMIFS({usage.col("nodes_avg")},{cluster},{month})'
    per_node = {kind: f'(SUMIFS({usage.col(field)},{cluster},{month})/{nodes_now})'
                for kind, field in (("cpu", "cpu_allocatable_avg"), ("mem", "mem_allocatable_avg_gib"))}
    need = {"cpu": f'SUMIFS({d.col("cpu_used_office_total" + suffix)},{by})',
            "mem": f'SUMIFS({d.col("mem_gib_used_office_total" + suffix)},{by})'}
    user = (f"MAX({s['rs_user_min']},"
            + ",".join(f"CEILING({need[kind]}*{s['rs_headroom']}/({per_node[kind]}*{s['rs_usable']}),1)" for kind in ("cpu", "mem"))
            + ")")
    nodes, aks, versus, cheaper = (get_column_letter(column) for column in columns)
    aca = "H" if not suffix else "T"
    if not suffix:
        sheet.cell(row=row, column=13, value=f'=IF(B{row}=0,"",{nodes_now})').number_format = "0.0"
    else:
        sheet.cell(row=row, column=columns[2], value=f'=IF(B{row}=0,"",SUMIFS({d.col("aca_use_hpa_usd")},{by}))'
                   ).number_format = USD
    sheet[f"{nodes}{row}"] = f'=IFERROR(IF(B{row}=0,"",{s["rs_system_nodes"]}+{user}),"")'
    sheet[f"{nodes}{row}"].number_format = "0"
    owner = f'{cost.col("owner_id")},L{row},{cost.col("owner_type")},"AKS",{cost.col("month")},"{cost_month}"'
    scalable = "+".join(f'SUMIFS({cost.col("cost_usd")},{owner},{cost.col("meter_category")},"{category}")'
                        for category in ("Virtual Machines", "Storage"))
    sheet[f"{aks}{row}"] = f'=IFERROR(IF({nodes}{row}="","",F{row}-({scalable})+({scalable})*{nodes}{row}/M{row}),"")'
    sheet[f"{aks}{row}"].number_format = USD
    if not suffix:
        sheet[f"{versus}{row}"] = f'=IF({aks}{row}="","",{aks}{row}-H{row})'
        sheet[f"{versus}{row}"].number_format = '"$"#,##0;[Red]-"$"#,##0'
    sheet[f"{cheaper}{row}"] = f'=IF({aks}{row}="","",IF({aca}{row}<{aks}{row},"ACA","AKS"))'


def build_aca_by_cluster(workbook, r, clusters, workloads, cost, cost_month, usage=None, hpas=None):
    prod = [row for row in clusters.rows if row["environment"] == "PROD" and row.get("category") == aa.APPLICATION]
    ids = {row["id"].lower() for row in prod}
    rows, full = aca_workloads(workloads, ids, hpa_replica_factors(hpas))
    if not rows:
        return
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    d = Detail(workbook, "D_ACA_workloads", rows, list(rows[0].keys()),
               extra=["aca_same_replicas_request_usd", "aca_same_replicas_use_usd", "aca_adt_use_usd",
                      "aca_use_hpa_usd", "cpu_used_office_total_hpa", "mem_gib_used_office_total_hpa"])
    letter = {field: d.letter(field) for field in d.fields}
    for index in range(2, len(rows) + 2):
        def rate(vcpu, mem):
            return f"(${letter[vcpu]}{index}*{s['p_vcpu_active']}+${letter[mem]}{index}*{s['p_mem_active']})"

        fits = f'${letter["fits_aca"]}{index}="YES"'
        office = f"${letter['office_replica_hours_month']}{index}"
        all_hours = f"({office}+${letter['offhours_replica_hours_month']}{index})"
        d.sheet[f"{letter['aca_same_replicas_request_usd']}{index}"] = (
            f"=IF({fits},{all_hours}*3600*{rate('aca_vcpu', 'aca_mem_gib')},0)")
        d.sheet[f"{letter['aca_same_replicas_use_usd']}{index}"] = (
            f"=IF({fits},{all_hours}*3600*{rate('aca_vcpu_by_use', 'aca_mem_gib_by_use')},0)")
        d.sheet[f"{letter['aca_adt_use_usd']}{index}"] = (
            f"=IF({fits},{office}*3600*{rate('aca_vcpu_by_use', 'aca_mem_gib_by_use')},0)")
        factor = f"${letter['hpa_replica_factor']}{index}"
        d.sheet[f"{letter['aca_use_hpa_usd']}{index}"] = f"=${letter['aca_same_replicas_use_usd']}{index}*{factor}"
        for total in ("cpu_used_office_total", "mem_gib_used_office_total"):
            d.sheet[f"{letter[total + '_hpa']}{index}"] = f"=${letter[total]}{index}*{factor}"
    sheet = workbook.create_sheet("ACA_by_cluster")
    title(sheet, "Proposal: each PROD application cluster, AKS cost today vs Container Apps with its real replicas",
          f"Replicas measured by Container Insights (average of {period_text(full)}), application workloads (ReplicaSet), "
          "APPLICATION clusters only. ACA size: the smallest combination that covers the request or the use. Worst case "
          "for ACA: active rate all the time, list price without discount, no free grant; the request includes the "
          "Istio sidecars. Size by use = average use + 50%. ACA goes from 0.25 vCPU / 0.5 GiB to 4 vCPU / 8 GiB; "
          "ADT_019 sends to AKS anything above 4 vCPU / 8 GB, so those workloads stay out of the ACA cost. AKS with "
          "right sizing (columns M to Q): the same pods by use + headroom, on nodes of the cluster's current size, with "
          "the node floor of Assumptions; VMs and disks scale with the nodes, the rest of the cost stays fixed. "
          "Columns R to U: in addition, HPAs that stay at their minimum drop to what their CPU asks for (at least 2 "
          "replicas), on AKS and on ACA.", 21)
    sheet.append([])
    headers = ["Cluster", "Workloads", "Do not fit ACA", "Average replicas, office", "Average replicas, off hours",
               f"AKS cost {cost_month} (USD)", "ACA same replicas, size by request (worst case)",
               "ACA same replicas, size by use", "ACA ADT rule, size by use (0 off hours)",
               "AKS today − ACA by use", "Cheaper (AKS today vs ACA by use)", "Id", "Nodes today",
               "Nodes with right sizing", "AKS with right sizing (USD, estimated)", "AKS with right sizing − ACA by use",
               "Cheaper (AKS with right sizing vs ACA by use)", "Nodes with right sizing and adjusted HPA minimums",
               "AKS with right sizing and adjusted HPA (USD)", "ACA by use with adjusted HPA minimums (USD)",
               "Cheaper (both with adjusted HPA)"]
    sheet.append(["Total of the clusters with data"])
    total_row = sheet.max_row
    sheet.cell(row=total_row, column=1).font = BOLD
    header(sheet, headers)
    first = sheet.max_row + 1
    office_hours = f"({s['office_seconds']}/3600)"
    off_hours = f"({s['hours_month']}-{s['office_seconds']}/3600)"
    for cluster in sorted(prod, key=lambda row: row["name"]):
        sheet.append([cluster["name"]])
        row = sheet.max_row
        by = f'{d.col("cluster_id")},$L{row}'
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
        sheet.cell(row=row, column=11, value=f'=IF(B{row}=0,"No data",IF(H{row}<F{row},"ACA","AKS"))')
        sheet.cell(row=row, column=12, value=cluster.get("id", ""))
        if usage and cost:
            add_right_sizing(sheet, row, s, d, usage, cost, cost_month)
            add_right_sizing(sheet, row, s, d, usage, cost, cost_month, "_hpa", (18, 19, 20, 21))
    last = sheet.max_row
    for column in (6, 7, 8, 9, 10) + ((13, 14, 15, 16, 18, 19, 20) if usage and cost else ()):
        col = get_column_letter(column)
        sheet.cell(row=total_row, column=column, value=f'=SUMIFS({col}{first}:{col}{last},$B{first}:$B{last},">0")'
                   ).number_format = USD
    sheet.cell(row=total_row, column=2, value=f'=COUNTIFS(B{first}:B{last},">0")')
    sheet.cell(row=total_row, column=11, value=f'=COUNTIFS(K{first}:K{last},"ACA")&" of "&B{total_row}&" cheaper on ACA"')
    if usage and cost:
        for column in ("Q", "U"):
            sheet[f"{column}{total_row}"] = (f'=COUNTIFS({column}{first}:{column}{last},"ACA")&" of "&'
                                             f'COUNTIFS({column}{first}:{column}{last},"?*")&" cheaper on ACA"')
    set_widths(sheet, [30, 11, 11, 13, 15, 16, 18, 16, 20, 16, 16, 40, 10, 12, 18, 18, 18, 14, 18, 18, 18])
    sheet.column_dimensions["L"].hidden = True  # resource id, only used to match the cost
    sheet.freeze_panes = f"B{first}"


# ---------------------------------------------------------------------------
# Summary, Sources, Queries
# ---------------------------------------------------------------------------


def build_summary_by_bia(sheet, kclusters):
    """Clusters, cost, nodes, microservices and office CPU use per BIA level, PROD application clusters only."""
    k = kclusters
    sheet.append([])
    sheet.append(["Information annex: by business criticality (BIA), PROD application clusters. It does not change "
                  "the figures above"])
    sheet.cell(row=sheet.max_row, column=1).font = BOLD
    headers = ["BIA", "Clusters", "Total monthly cost", "Average cost per cluster", "Average nodes per cluster",
               "Average microservices per cluster", "CPU use in office hours"]
    sheet.append(headers)
    style_header(sheet, sheet.max_row, len(headers))
    present = {row["bia"] for row in k.rows if row["environment"] == "PROD" and row.get("category") == aa.APPLICATION}
    levels = [level for level in BIA_LEVELS if level in present] + sorted(present - set(BIA_LEVELS))
    first = sheet.max_row + 1
    for level in levels + ["Total"]:
        sheet.append([level])
        row = sheet.max_row
        criteria = f'{k.col("environment")},"PROD",{k.col("category")},{APP},{k.col("in_inventory")},"YES"'
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
    sheet.append(["The BIA of a cluster is the highest among its microservices (deployment export). Cost: Cost "
                  "Management, amortized; use: platform metrics, Monday to Friday 8-20."])
    return first


def build_summary(workbook, breakeven, measured_refs, r, kclusters, cluster_refs, period, months, from_snapshot):
    sheet = workbook.create_sheet("Summary", 0)
    title(sheet, "AKS vs Azure Container Apps: summary", None)
    sheet.append([f"Generated: {date.today().isoformat()}. USD per month. Data period: {period}. Type: measured (Azure), "
                  "derived (calculation over measured figures), estimated (model with assumptions). PROD figures cover "
                  "APPLICATION clusters only (Clusters sheet)."])
    if not from_snapshot:
        sheet.append([FALLBACK_NOTE])
    sheet.append([])
    sheet.append(["Indicator", "Value", "Type", "Origin"])
    style_header(sheet, sheet.max_row, 4)
    s = {key: f"Assumptions!{value}" for key, value in r.items()}
    empty = chr(34) * 2

    def clusters_ref(key, index=0):
        return cluster_refs[key][index] if key in cluster_refs else empty

    rows = [
        ("Period of the data", period, "", "Cost Management months; metrics and Container Insights cover the same window", None),
        ("Months in the data", ", ".join(months), "", "", None),
        ("PROD AKS clusters (all categories)", f"={clusters_ref('total')}", "Measured", "Clusters", "0"),
        ("PROD clusters: APPLICATION (in the analysis)", f"={clusters_ref(aa.APPLICATION)}", "Measured", "Clusters", "0"),
        ("PROD clusters: NO_APPLICATIONS (created without applications)", f"={clusters_ref(aa.NO_APPLICATIONS)}",
         "Measured", "Clusters", "0"),
        ("PROD clusters: INFRASTRUCTURE (tools only, excluded)", f"={clusters_ref(aa.INFRASTRUCTURE)}", "Measured",
         "Clusters", "0"),
        ("PROD clusters: NO_DATA (not in the platform export, excluded)", f"={clusters_ref(aa.NO_DATA)}", "Measured",
         "Clusters", "0"),
        ("Monthly cost of PROD clusters without applications (finding)", f"={clusters_ref(aa.NO_APPLICATIONS, 1)}",
         "Measured", "Clusters; Cost Management", USD),
        ("Non-production AKS clusters (Dev/QA)", f"={measured_refs.get('non_prod', empty)}",
         "Measured", "Measured_AKS (DESA + CERT); context, the analysis covers PROD", "0"),
        ("PROD application clusters with 4 nodes or fewer", f"={measured_refs.get('le4', empty)}", "Measured",
         "Measured_AKS", "0"),
        ("Average monthly cost per PROD application cluster", f"={measured_refs.get('cost_avg', empty)}",
         "Measured", "Measured_AKS; Cost Management", USD),
        ("Total monthly cost of PROD application clusters", f"={measured_refs.get('cost_total', empty)}", "Measured",
         "Measured_AKS; Cost Management", USD),
        ("CPU use in office hours, PROD application clusters", f"={measured_refs.get('cpu', empty)}", "Measured",
         "Measured_AKS; platform metrics", "0%"),
        ("Memory use in office hours, PROD application clusters", f"={measured_refs.get('mem', empty)}", "Measured",
         "Measured_AKS; platform metrics", "0%"),
        ("CPU use, hourly 95th percentile (whole day)", f"={measured_refs.get('cpu_p95', empty)}", "Measured",
         "Measured_AKS; platform metrics", "0%"),
        ("CPU use, maximum (upper bound: peaks of each cluster added up)", f"={measured_refs.get('cpu_max', empty)}",
         "Measured", "Measured_AKS; platform metrics", "0%"),
        ("Memory use, hourly 95th percentile (whole day)", f"={measured_refs.get('mem_p95', empty)}", "Measured",
         "Measured_AKS; platform metrics", "0%"),
        ("Memory use, maximum (upper bound: peaks of each cluster added up)", f"={measured_refs.get('mem_max', empty)}",
         "Measured", "Measured_AKS; platform metrics", "0%"),
        ("Actual node discount (effective / list)", f"={s['discount']}", "Derived", "Assumptions; Cost Management", "0.00"),
        ("Minimum AKS cluster (4 nodes) at list price", f"={breakeven['Minimum AKS cluster cost per month (list)']}",
         "Estimated", "Model", USD),
        ("Minimum AKS cluster with the measured discount",
         f"={breakeven['Minimum AKS cluster cost per month (effective)']}", "Estimated", "Model", USD),
        ("ACA per microservice (ADT schedule)", f"={breakeven['ACA cost per microservice per month (ADT schedule)']}",
         "Estimated", "Model", USD2),
        ("Microservices from which ACA costs more than AKS (list)",
         f"={breakeven['Microservices at which ACA (ADT schedule) exceeds AKS list']}", "Estimated", "Break_even", "0"),
        ("Microservices from which ACA costs more than AKS (effective)",
         f"={breakeven['Microservices at which ACA (ADT schedule) exceeds AKS effective']}", "Estimated", "Break_even", "0"),
        ("Microservices from which ACA 24x7 costs more than AKS (list)",
         f"={breakeven['Microservices at which ACA 24x7 exceeds AKS list']}", "Estimated", "Break_even", "0"),
    ]
    for label, formula, kind, origin, fmt in rows:
        sheet.append([label, formula, kind, origin])
        if fmt:
            sheet.cell(row=sheet.max_row, column=2).number_format = fmt
    sheet.append([f'">{MAX_MICROSERVICES}" = no crossing: from 1 to 30 and at 50, 100, 150, 200, 300 and '
                  f'{MAX_MICROSERVICES} microservices ACA costs less, because AKS adds nodes as it grows. AKS wins by '
                  "active time or a shared cluster (Break_even sheet). INFRASTRUCTURE and NO_DATA clusters are only "
                  "counted here: no other figure includes them."])
    if kclusters is not None and "bia" in kclusters.fields:
        build_summary_by_bia(sheet, kclusters)
    set_widths(sheet, [52, 16, 18, 36, 18, 22, 16])


SOURCE_ROWS = [
    ("Inventory of clusters, node pools and add-ons", "Measured", "Azure Resource Graph",
     "resources | where type =~ 'microsoft.containerservice/managedclusters'", "aks_clusters.csv, aks_node_pools.csv",
     "Resource Graph Explorer with the query of the Queries sheet; or the cluster > Overview / Node pools"),
    ("Container Apps apps and environments", "Measured", "Azure Resource Graph",
     "resources | where type in~ ('microsoft.app/containerapps','microsoft.app/managedenvironments')",
     "aca_apps.csv, aca_environments.csv", "Container App > Scale; environment > Workload profiles"),
    ("APIM APIs and their backend", "Measured", "ARM (read)", "GET {apim}/apis", "apim_apis.csv",
     "APIM > APIs > Settings > Web service URL"),
    ("Node CPU and memory use", "Measured", "Azure Monitor (platform metrics)",
     "node_cpu_usage_millicores, node_memory_working_set_bytes, kube_node_status_allocatable_*; PT1H; by node and nodepool",
     "aks_usage_monthly.csv", "Cluster > Monitoring > Metrics, same metric, split by nodepool"),
    ("Requests per API", "Measured", "Azure Monitor (platform metrics)", "Requests (Total), PT5M, split by ApiId",
     "apim_api_monthly.csv", "APIM > Monitoring > Metrics > Requests, apply splitting by API"),
    ("Wake-ups and active seconds outside office hours", "Derived", "Azure Monitor", "Runs of 5-minute intervals with "
     "requests outside Monday-Friday 8-20 (Lima); 300 s cooldown per run", "apim_api_monthly.csv",
     "Recalculable with aa_metrics.py --cooldown"),
    ("Replicas, requests and use per workload", "Measured", "Log Analytics (Container Insights)",
     "KubePodInventory + Perf (K8SContainer); aggregated by hour, month and office hours", "aks_workloads_monthly.csv",
     "Cluster > Insights > Controllers; or Logs with the query of the Queries sheet"),
    ("Actual cost of AKS and Container Apps", "Measured (actual billed cost)", "Cost Management Query API",
     "AmortizedCost (reservations and savings plan spread), monthly, grouped by ResourceId + MeterCategory + Meter; "
     "cluster + node resource group",
     "cost_by_resource.csv", "Cost analysis (Amortized cost view) > subscription scope > filter resource group MC_* and type managedClusters"),
    ("Actual node discount", "Derived from actual cost", "Cost Management + Retail Prices",
     "Amortized cost of PROD application node VMs / (measured node-hours × list price), full months",
     "Assumptions (discount)", "Cost analysis, VM meter"),
    ("List prices", "Vendor data", "Azure Retail Prices API (public)",
     "Filters in the 'query' column of Prices", "prices.csv", "https://azure.microsoft.com/pricing/calculator"),
    ("Container Apps minimum replica, schedule and limits", "Internal guideline", "ADT_019",
     "1 replica from 8 to 20 on working days; 4 vCPU / 8 GiB; thresholds 15 microservices, 10 h, 100 tps",
     "Assumptions, Break_even", "ADT_019 in Confluence"),
    ("Load and capacity assumptions", "Assumption", "Architecture criterion",
     "Yellow cells of Assumptions (RPS per replica, DaemonSet and sidecar overhead)", "Assumptions",
     "Validate with a load test"),
]


def build_sources(workbook, extra_rows=()):
    sheet = workbook.create_sheet("Sources")
    title(sheet, "Where each figure comes from",
          "Measured = figure read from Azure in the run. Derived = calculation over measured figures. Estimated = model "
          "with assumptions. Cost comes from Cost Management (AmortizedCost): it includes contract discounts and spreads "
          "reservations and savings plans over the VMs that use them.", 6)
    sheet.append([])
    headers = ["Data", "Type", "Source", "How it is obtained", "Output file", "How to check it"]
    sheet.append(headers)
    style_header(sheet, 4, len(headers))
    for values in SOURCE_ROWS + list(extra_rows):
        sheet.append(list(values))
        for column in range(1, 7):
            sheet.cell(row=sheet.max_row, column=column).alignment = WRAP
    set_widths(sheet, [34, 18, 26, 60, 26, 50])


def build_queries(workbook, results):
    sheet = workbook.create_sheet("Queries")
    title(sheet, "Queries run (evidence recorded by the scripts)")
    sheet.append([])
    headers = ["Script", "Run date", "Type", "Label / endpoint", "Query, body or api-version", "Rows / calls", "Errors"]
    sheet.append(headers)
    style_header(sheet, 3, len(headers))
    for folder, name in (("aa-inventory", "inventory_summary.json"), ("aa-metrics", "metrics_summary.json"),
                         ("aa-container-insights", "ci_summary.json"), ("aa-cost", "cost_summary.json")):
        path = os.path.join(results, folder, name)
        if not os.path.exists(path):
            sheet.append([name, "", "Not run", "", "", "", ""])
            continue
        with open(path, encoding="utf-8") as handle:
            summary = json.load(handle)
        evidence, when = summary.get("evidence", {}), summary.get("generated_at", "")
        for query in evidence.get("resource_graph_queries", []):
            sheet.append([name, when, query.get("source", "Azure Resource Graph"), query["label"], query["query"],
                          query["rows"], ""])
        for call in evidence.get("arm_calls", []):
            sheet.append([name, when, "ARM (read)", f"{call['method']} {call['pattern']}",
                          f"api-version={call['api_version']}", call["calls"], call["errors"]])
        for body in evidence.get("request_bodies", []):
            sheet.append([name, when, "Query body", f"POST {body['endpoint']}",
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
    sizes = {(pool["cluster_id"].lower(), pool["pool"].lower()): pool["vm_size"] for pool in pools}
    regions = {cluster["id"].lower(): cluster["location"] for cluster in clusters}
    for row in usage:
        row["vm_size"] = sizes.get((row["cluster_id"], row["pool"].lower()), "")
        row["location"] = regions.get(row["cluster_id"], "")


def add_pool_figures(clusters, pools):
    """Configured minimum nodes per cluster (minCount of autoscaled pools, the fixed count of the others) and its
    number of user pools: 0 means the applications run on the system nodes."""
    totals, system, user_pools = defaultdict(int), defaultdict(int), defaultdict(int)
    for pool in pools:
        autoscaled = str(pool.get("autoscaling", "")).lower() == "true" and pool.get("min_count", "") != ""
        minimum = int(float(pool["min_count"] if autoscaled else pool["count"] or 0))
        key = pool["cluster_id"].lower()
        totals[key] += minimum
        if pool.get("mode") == "System":
            system[key] += minimum
        else:
            user_pools[key] += 1
    for cluster in clusters:
        key = cluster["id"].lower()
        cluster["min_nodes"] = totals.get(key, "")
        cluster["system_min_nodes"] = system.get(key, 0) if key in totals else ""
        cluster["user_pools"] = user_pools.get(key, 0) if key in totals else ""


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


def analysis_months(cost, usage):
    """Months the figures describe: the Cost Management months, else the finished months of the metrics."""
    months = sorted({row["month"] for row in cost if row.get("month")})
    if months:
        return months
    current = date.today().strftime("%Y-%m")
    finished = sorted({row["month"] for row in usage if row.get("month") and row["month"] < current})
    return finished or sorted({row["month"] for row in usage if row.get("month")})


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
    ids_by_name = aa.cluster_ids_by_name(clusters)
    pools = load(args.results, "aa-inventory", "aks_node_pools.csv")
    apis = load(args.results, "aa-inventory", "apim_apis.csv")
    apps = load(args.results, "aa-inventory", "aca_apps.csv")
    usage = load(args.results, "aa-metrics", "aks_usage_monthly.csv")
    traffic = load(args.results, "aa-metrics", "apim_api_monthly.csv")
    workloads = load(args.results, "aa-container-insights", "aks_workloads_monthly.csv")
    services = [row for row in load(args.results, "aa-container-insights", "aks_services.csv")
                if not aa.is_platform_namespace(row["namespace"])]
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

    # Join by resource id: two clusters can share a name in different subscriptions. Rows of a shared name without
    # cluster_id cannot be attributed and are left out, as are clusters the inventory no longer lists (POC ones).
    categories, from_snapshot = cluster_categories(clusters, k8s_clusters)
    for row in clusters:
        row["category"], row["category_reason"] = categories[row["id"].lower()]
    for rows in (usage, workloads, services):
        for row in rows:
            row["cluster_id"] = aa.cluster_id_of(row, ids_by_name)
    usage, workloads, services = ([row for row in rows if row["cluster_id"] in categories]
                                  for rows in (usage, workloads, services))
    for rows in (usage, workloads, hpas, deployments, k8s_clusters):
        set_category(rows, categories)
    for row in cost:
        row["category"] = categories.get(row["owner_id"].lower(), ("", ""))[0] if row["owner_type"] == "AKS" else ""
    for row in hpas + deployments:
        if not row.get("microservice"):
            row["microservice"] = "YES" if aa.is_microservice(row["namespace"], row["deployment"]) else "NO"
    add_pool_sizes(usage, pools, clusters)
    add_pool_figures(clusters, pools)
    months_list = analysis_months(cost, usage)
    period = period_text(months_list)
    k8s_by_id = {row["cluster_id"].lower(): row for row in k8s_clusters if row.get("cluster_id")}

    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.calculation.fullCalcOnLoad = True

    price_fields = ["key", "region", "service", "product", "sku", "meter", "unit", "retail_price", "hourly_equivalent",
                    "price_type", "reservation_term", "tier_minimum_units", "effective_start", "meter_id", "arm_sku", "query"]
    Detail(workbook, "Prices", prices, price_fields)
    cluster_fields = ["name", "environment", "category", "category_reason", "cod_app", "environment_evidence",
                      "subscription", "resource_group", "location", "kubernetes_version", "minor_versions_behind",
                      "node_pools", "system_nodes", "user_nodes", "user_max_nodes", "min_nodes", "vm_sizes", "zones"] + [
                      p for p in CONFIG_PROPERTIES if p not in ("kubernetes_version", "minor_versions_behind", "zones")] + [
                      "istio_revisions", "extension_types", "id"]
    d_clusters = Detail(workbook, "D_Clusters", clusters, cluster_fields) if clusters else None
    if pools:
        Detail(workbook, "D_Pools", pools, list(pools[0].keys()))
    if apps:
        Detail(workbook, "D_ACA_apps", apps, list(apps[0].keys()))
    # Allocatable rows lack the usage columns and vice versa, so fields are the union of every row.
    d_usage = None
    if usage:
        d_usage = Detail(workbook, "D_Usage", usage, all_fields(usage), complete=True, extra=["list_cost_usd"])
        letters = {field: d_usage.letter(field) for field in d_usage.fields}
        for index in range(2, len(usage) + 2):
            d_usage.sheet[f"{letters['list_cost_usd']}{index}"] = (
                f'=IF(OR(${letters["vm_size"]}{index}="",${letters["nodes_avg"]}{index}=""),"",'
                f'${letters["nodes_avg"]}{index}*${letters["hours"]}{index}*SUMIFS(Prices!$H:$H,'
                f'Prices!$A:$A,"vm_"&${letters["vm_size"]}{index}&"_payg",Prices!$B:$B,${letters["location"]}{index}))')
    d_traffic = Detail(workbook, "D_Traffic", traffic, list(traffic[0].keys())) if traffic else None
    d_workloads = Detail(workbook, "D_Workloads", workloads, all_fields(workloads), complete=True) if workloads else None
    d_services = Detail(workbook, "D_Services", services, list(services[0].keys())) if services else None
    d_cost = Detail(workbook, "D_Cost", cost, list(cost[0].keys())) if cost else None
    d_hpa = Detail(workbook, "D_HPA", hpas, all_fields(hpas), complete=True) if hpas else None
    d_deploy = Detail(workbook, "D_Deploy", deployments, all_fields(deployments), complete=True) if deployments else None
    d_kclusters = None
    if k8s_clusters:
        usage_extra = ["nodes_avg", "cpu_allocatable_avg", "cpu_used_office_avg"] if d_usage else []
        d_kclusters = Detail(workbook, "D_K8s_clusters", k8s_clusters, all_fields(k8s_clusters), complete=True,
                             extra=["cost_month_usd"] + usage_extra)
        usage_months = max(1, distinct_months(usage))
        cluster_id = d_kclusters.letter("cluster_id")
        for index in range(2, len(k8s_clusters) + 2):
            # Last full month, not the average: clusters created inside the period would look cheaper.
            d_kclusters.sheet[f"{d_kclusters.letter('cost_month_usd')}{index}"] = (
                f'=SUMIFS({d_cost.col("cost_usd")},{d_cost.col("owner_id")},${cluster_id}{index},{d_cost.col("owner_type")},"AKS",'
                f'{d_cost.col("month")},"{last_cost_month}")' if d_cost else "")
            for field in usage_extra:
                d_kclusters.sheet[f"{d_kclusters.letter(field)}{index}"] = (
                    f'=SUMIFS({d_usage.col(field)},{d_usage.col("cluster_id")},${cluster_id}{index})/{usage_months}')

    months = {"cost": distinct_months(cost), "usage": distinct_months(usage)}
    apis_last = build_traffic(workbook, d_traffic, apis, distinct_months(traffic)) if d_traffic else None
    workloads_last = build_workloads(workbook, d_workloads, period) if d_workloads else None
    refs = {}
    r = build_assumptions(workbook, refs, apis_last, workloads_last, d_cost, d_usage, full_months(usage, cost))
    model = build_model(workbook, r)
    map_info = build_map(workbook, r)
    breakeven = build_breakeven(workbook, r, model, map_info)
    build_sensitivity(workbook, r)
    if profile:
        build_traffic_profile(workbook, profile, d_traffic, period)
    if d_clusters and workloads:
        build_aca_by_cluster(workbook, r, d_clusters, workloads, d_cost, last_cost_month, d_usage, hpas)
    cluster_refs = {}
    measured_refs = {}
    if d_clusters:
        cluster_refs = build_clusters(workbook, d_clusters, d_cost, max(1, months["cost"]), k8s_by_id, from_snapshot,
                                      period)
        measured_refs, _ = build_measured(workbook, d_clusters, d_usage, d_cost, d_workloads, d_services, months, period)
        build_config(workbook, d_clusters)
    snapshot_sources = [(
        "Cluster category (APPLICATION, NO_APPLICATIONS, INFRASTRUCTURE, NO_DATA)", "Derived",
        "Platform export, Container Insights and app code" if from_snapshot else "App code (cod_app) of the inventory",
        "Microservices per cluster in the export and in Container Insights; clusters of infrastructure apps are "
        "INFRASTRUCTURE" if from_snapshot else FALLBACK_NOTE,
        "cluster_snapshot.csv" if from_snapshot else "aks_clusters.csv",
        "Clusters sheet; kubectl get deployments -A on the cluster")]
    if d_hpa or d_deploy:
        taken_on = snapshot.get("snapshot_date", "")
        build_k8s(workbook, d_hpa, d_deploy, taken_on, period)
        if d_deploy and d_kclusters:
            build_patterns(workbook, d_deploy, d_hpa, d_kclusters, capabilities, taken_on, period,
                           period_text([last_cost_month]))
            if "bia" in d_kclusters.fields:
                build_bia(workbook, d_kclusters)
        snapshot_sources.append((
            f"HPA, requests and limits per deployment (snapshot of {taken_on})", "Measured (one day)",
            "Platform team export", ", ".join(item["file"] for item in snapshot.get("inputs", [])),
            "hpa_snapshot.csv, deployment_snapshot.csv",
            "kubectl get hpa / get deployment on the cluster; the history, in Container Insights"))
    build_criteria(workbook, d_clusters, d_kclusters, d_deploy, d_hpa, d_traffic, workloads_last,
                   max((row["month"] for row in traffic if row["month"] < date.today().strftime("%Y-%m")), default=""),
                   period)
    build_summary(workbook, breakeven, measured_refs, r, d_kclusters, cluster_refs, period, months_list, from_snapshot)
    build_sources(workbook, snapshot_sources)
    build_queries(workbook, args.results)

    order = ["Summary", "Clusters", "Patterns", "Technical_criteria", "Assumptions", "Model", "Break_even", "Map",
             "ACA_by_cluster", "Traffic_profile", "Measured_AKS", "APIM_traffic", "Workloads", "Config_AKS",
             "Config_K8s", "Sources", "Prices", "BIA_candidates"]
    workbook._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order
                          else (len(order) + 1 if ws.title == "Queries" else len(order)))
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
