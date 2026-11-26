#!/usr/bin/env python3
"""Build the AKS right-sizing workbook: how well each AKS cluster uses what it pays for, and what that costs.

The question behind it: Container Apps only beats AKS where AKS is used badly. One vCPU of a reserved AKS node costs
about a quarter of one Container Apps vCPU, so a well-used AKS cluster stays cheaper; the saving comes from clusters
that pay for capacity they do not use (a dedicated cluster per application with a fixed minimum, system nodes,
requests far above use, empty clusters, non-production on 24/7).

Reads only results/ (run aa_cluster_snapshot.py first):
  aa-k8s-snapshot/cluster_snapshot.csv, hpa_snapshot.csv   category, "micro"-services, autoscaling
  aa-metrics/aks_usage_monthly.csv                          CPU and memory used against allocatable, per pool
  aa-inventory/aks_clusters.csv                             environment, nodes, VM sizes
  aa-cost/cost_by_resource.csv                              cost of the month
  the criteria v2 model (aa_v2_model.py)                    Container Apps cost of each PROD business cluster

Writes:
  results/aa-rightsizing/cluster_rightsizing.csv   one row per AKS cluster with cost in the month (any environment)
  reports/<date>-aks-aca/aks_rightsizing.xlsx      Summary, Clusters, Pain points, Breakeven, Variables; the
                                                   cluster table is a Power Query connection to the CSV

Usage: python aks-aca/src/aa_rightsizing.py [--results DIR] [--output FILE]
"""

import argparse
import glob
import os
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill

import aa_common as aa
import aa_criteria_workbook as cwb
import aa_criteria_workbook_v1 as v1wb
import aa_deck_exec_v1 as v1
import aa_powerquery as pq
import aa_v2_model as model

FOLDER, FILE = "aa-rightsizing", "cluster_rightsizing.csv"
COLUMNS = ["cluster", "cod_app", "environment", "category", "month", "cost_month_usd", "nodes_system",
           "nodes_user", "vm_sizes", "microservices", "cpu_allocatable", "mem_allocatable_gib",
           "cpu_used_office_avg", "cpu_used_offhours_avg", "cpu_used_max", "mem_used_office_avg_gib",
           "mem_used_offhours_avg_gib", "mem_used_max_gib", "cpu_used_system_office_avg",
           "mem_used_system_office_avg_gib", "hpas", "hpas_still_3_months", "aca_month_pen", "aca_group"]
TEXT = {"cluster", "cod_app", "environment", "category", "month", "vm_sizes", "aca_group"}
NUMBERS = [column for column in COLUMNS if column not in TEXT]
BOLD = Font(bold=True)
EDITABLE = PatternFill("solid", fgColor="FFF2CC")
GREEN, AMBER, RED = (PatternFill("solid", fgColor=colour) for colour in ("C6EFCE", "FFEB9C", "FFC7CE"))
USD, PEN, PCT = '"$"#,##0.00', '"S/ "#,##0', "0%"


def to_float(value):
    return v1.to_float(value)


def rows(results):
    """One row per AKS cluster with cost in the month, PROD with its use of CPU and memory."""
    read = lambda folder, name: v1.results_csv(results, folder, name)
    usage = read("aa-metrics", "aks_usage_monthly.csv")
    month = max((row["month"] for row in usage), default="")
    snapshot = {row["cluster_id"].lower(): row for row in read("aa-k8s-snapshot", "cluster_snapshot.csv")}
    # Non-production clusters take the category of their application's PROD clusters (code in the name).
    prod_category = defaultdict(set)
    for row in snapshot.values():
        if row["environment"] == "PROD":
            prod_category[row["cod_app"].upper()].add(row["category"])
    cost = defaultdict(float)
    for row in read("aa-cost", "cost_by_resource.csv"):
        if row["month"] == month:
            cost[row["owner_id"].lower()] += to_float(row["cost_usd"]) or 0
    used = defaultdict(lambda: defaultdict(float))
    for row in usage:
        if row["month"] != month:
            continue
        entry = used[row["cluster_id"].lower()]
        if row["mode"] == "ALLOCATABLE":
            entry["cpu_allocatable"] += to_float(row["cpu_allocatable_avg"]) or 0
            entry["mem_allocatable_gib"] += to_float(row["mem_allocatable_avg_gib"]) or 0
            continue
        for column in ("cpu_used_office_avg", "cpu_used_offhours_avg", "cpu_used_max", "mem_used_office_avg_gib",
                       "mem_used_offhours_avg_gib", "mem_used_max_gib"):
            entry[column] += to_float(row[column]) or 0
        if row["mode"] == "System":
            entry["cpu_used_system_office_avg"] += to_float(row["cpu_used_office_avg"]) or 0
            entry["mem_used_system_office_avg_gib"] += to_float(row["mem_used_office_avg_gib"]) or 0
    hpas = defaultdict(Counter)
    for row in read("aa-k8s-snapshot", "hpa_snapshot.csv"):
        if row["microservice"] == "YES" and row["last_scale_bucket"] != "NO_DATA":
            hpas[row["cluster_id"].lower()]["hpas"] += 1
            hpas[row["cluster_id"].lower()]["still"] += row["last_scale_bucket"] in ("3_TO_12_MONTHS", "OVER_1_YEAR")
    model.use("v2")
    aca = {row["cluster_id"].lower(): row for row in model.evaluate(model.load(results))}
    out = []
    for cluster in read("aa-inventory", "aks_clusters.csv"):
        key = cluster["id"].lower()
        if cost.get(key, 0) <= 0:
            continue
        environment = cluster["environment"]
        snap = snapshot.get(key, {})
        # PROD: the snapshot's application; CERT and DESA: the code in the name (aksveu2<code><env><nn>)
        code = (snap.get("cod_app") if environment == "PROD" and snap.get("cod_app") else cluster["name"][7:11]).upper()
        if environment == "PROD":
            category = snap.get("category") or "NO_DATA"
        else:
            categories = prod_category.get(code, set())
            category = aa.APPLICATION if aa.APPLICATION in categories else \
                (sorted(categories)[0] if categories else "NO_PROD_CLUSTER")
        entry, scaling, criteria = used.get(key, {}), hpas.get(key, Counter()), aca.get(key)
        out.append({
            "cluster": cluster["name"], "cod_app": code, "environment": environment, "category": category,
            "month": month, "cost_month_usd": round(cost[key], 2),
            "nodes_system": cluster.get("system_nodes", ""), "nodes_user": cluster.get("user_nodes", ""),
            "vm_sizes": cluster.get("vm_sizes", ""), "microservices": snap.get("microservices", ""),
            **{column: round(entry[column], 3) if entry else "" for column in
               ("cpu_allocatable", "mem_allocatable_gib", "cpu_used_office_avg", "cpu_used_offhours_avg",
                "cpu_used_max", "mem_used_office_avg_gib", "mem_used_offhours_avg_gib", "mem_used_max_gib",
                "cpu_used_system_office_avg", "mem_used_system_office_avg_gib")},
            "hpas": scaling["hpas"] or "", "hpas_still_3_months": scaling["still"] if scaling["hpas"] else "",
            "aca_month_pen": round(criteria["aca_month_pen"], 2) if criteria and environment == "PROD" else "",
            "aca_group": criteria["group"] if criteria and environment == "PROD" else ""})
    return sorted(out, key=lambda row: (row["environment"] != "PROD", -row["cost_month_usd"])), month


def summary(data, optimized=0.6, partly=0.3, minimum=None):
    """What the deck shows, computed as the workbook's formulas do with its default variables."""
    minimum = minimum or v1.MIN_AKS_NODES
    soles = (1 + v1.IGV) * v1.EXCHANGE_RATE
    figure = lambda row, column: to_float(row[column]) or 0
    out = {"classes": defaultdict(lambda: {"clusters": 0, "aks": 0.0, "idle": 0.0, "aca": 0.0, "aks_costed": 0.0}),
           "minimum": [0, 0.0], "empty": [0, 0.0], "nonprod": [0, 0.0], "hpas_still": 0, "hpas": 0,
           "system_mem": 0.0, "mem": 0.0, "rows": []}
    for row in data:
        aks = figure(row, "cost_month_usd") * soles
        business = row["environment"] == "PROD" and row["category"] == aa.APPLICATION
        if row["environment"] != "PROD":
            if row["category"] == aa.APPLICATION:
                out["nonprod"][0] += 1
                out["nonprod"][1] += aks
            continue
        if row["category"] == aa.NO_APPLICATIONS:
            out["empty"][0] += 1
            out["empty"][1] += aks
        if not business:
            continue
        cpu = figure(row, "cpu_used_office_avg") / figure(row, "cpu_allocatable") if figure(row, "cpu_allocatable") else None
        mem = figure(row, "mem_used_office_avg_gib") / figure(row, "mem_allocatable_gib") \
            if figure(row, "mem_allocatable_gib") else 0
        use = max(cpu, mem) if cpu is not None else None
        kind = "NO_DATA" if use is None else "OPTIMIZED" if use >= optimized else \
            "PARTLY_USED" if use >= partly else "OVERSIZED"
        entry = out["classes"][kind]
        entry["clusters"] += 1
        entry["aks"] += aks
        entry["idle"] += aks * (1 - min(1, use)) if use is not None else 0
        if row["aca_month_pen"] != "":
            entry["aca"] += figure(row, "aca_month_pen")
            entry["aks_costed"] += aks
        nodes = figure(row, "nodes_system") + figure(row, "nodes_user")
        if 0 < nodes <= minimum:
            out["minimum"][0] += 1
            out["minimum"][1] += aks
        out["hpas"] += int(figure(row, "hpas"))
        out["hpas_still"] += int(figure(row, "hpas_still_3_months"))
        out["system_mem"] += figure(row, "mem_used_system_office_avg_gib")
        out["mem"] += figure(row, "mem_used_office_avg_gib")
        out["rows"].append(dict(row, aks_pen=aks, use=use, kind=kind))
    return out


def query(folder):
    return v1wb.query_head(folder) + f'''    Clusters = ReadCsv("{FOLDER}", "{FILE}"),
    Columns = Table.SelectColumns(Clusters, {{{cwb.m_list(COLUMNS)}}}),
    Typed = {cwb.m_typed("Columns", NUMBERS)}
in
    Typed'''


def write_variables(book, sheet, prices):
    vm = f"vm_{v1.MIN_AKS_VM}_{v1.MIN_AKS_RESERVATION}"
    items = [("Classes", None, None),
             ("optimized_use", 0.6, "Optimized: uses on average at least this share of its CPU or memory in office hours"),
             ("partly_use", 0.3, "Partly used: at least this share; below it, oversized"),
             ("min_cluster_nodes", v1.MIN_AKS_NODES, "Minimum dedicated cluster (system + user nodes)"),
             ("Money", None, None),
             ("igv", v1.IGV, "IGV"), ("exchange_rate", v1.EXCHANGE_RATE, "Soles per US$"),
             ("bank_discount", v1.BANK_DISCOUNT, "Bank discount on Container Apps"),
             ("Break-even (list prices, East US 2)", None, None),
             ("vm_hour_reserved", round(prices.get((vm, "eastus2"), 0) / 26280, 6),
              f"{v1.MIN_AKS_VM}, 3-year reservation, US$ per hour"),
             ("vm_vcpus", 8, "vCPU of that VM"),
             ("aca_vcpu_second", prices.get(("aca_vcpu_active_s", "eastus2"), 0), "Container Apps, US$ per vCPU-second"),
             ("aca_gib_second", prices.get(("aca_mem_active_gib_s", "eastus2"), 0), "Container Apps, US$ per GiB-second"),
             ("aca_gib_per_vcpu", 2, "GiB per vCPU in Container Apps")]
    row = 1
    for label, value, comment in items:
        if value is None:
            sheet.cell(row=row, column=1, value=label).font = BOLD
        else:
            sheet.cell(row=row, column=1, value=label)
            cell = sheet.cell(row=row, column=2, value=value)
            cell.fill = EDITABLE
            if label in ("optimized_use", "partly_use", "igv", "bank_discount"):
                cell.number_format = PCT
            sheet.cell(row=row, column=3, value=comment)
            v1wb.define(book, label, f"Variables!$B${row}")
        row += 1
    for column, width in zip("ABC", (24, 12, 80)):
        sheet.column_dimensions[column].width = width


def cluster_formulas():
    c = lambda column: cwb.r("Clusters", column)
    raw = lambda column: f"INDEX(ClusterData[{column}],ROW()-1)"
    pulled = [(column, cwb.pull("ClusterData", column)) for column in COLUMNS]
    use = lambda used, alloc: f'IF(N({c(alloc)})>0,N({c(used)})/{c(alloc)},"")'
    return pulled + [
        ("aks_month_pen", f"N({c('cost_month_usd')})*(1+igv)*exchange_rate"),
        ("cpu_use", use("cpu_used_office_avg", "cpu_allocatable")),
        ("mem_use", use("mem_used_office_avg_gib", "mem_allocatable_gib")),
        # the busier of CPU and memory: what the nodes have to fit
        ("use", f'IF(ISNUMBER({c("cpu_use")}),MAX({c("cpu_use")},N({c("mem_use")})),"")'),
        ("class", f'IF(NOT(ISNUMBER({c("use")})),IF({c("environment")}="PROD","NO_DATA","NON_PROD"),'
                  f'IF({c("use")}>=optimized_use,"OPTIMIZED",IF({c("use")}>=partly_use,"PARTLY_USED","OVERSIZED")))'),
        ("idle_pen", f'IF(ISNUMBER({c("use")}),{c("aks_month_pen")}*(1-MIN(1,{c("use")})),"")'),
        ("system_mem_share", f'IF(N({c("mem_used_office_avg_gib")})>0,N({c("mem_used_system_office_avg_gib")})/'
                             f'{c("mem_used_office_avg_gib")},"")'),
        ("at_minimum", f'IF(AND({c("environment")}="PROD",N({c("nodes_system")})+N({c("nodes_user")})>0,'
                       f'N({c("nodes_system")})+N({c("nodes_user")})<=min_cluster_nodes),"YES","NO")'),
        ("aca_vs_aks", f'IF(AND(ISNUMBER({c("aca_month_pen")}),{c("aks_month_pen")}>0),'
                       f'{c("aca_month_pen")}/{c("aks_month_pen")},"")'),
    ], raw


def write_clusters(sheet, count):
    formulas, _ = cluster_formulas()
    cwb.formula_table(sheet, "Clusters", formulas, count)
    headers = [header for header, _ in formulas]
    for index, header in enumerate(headers, start=1):
        fmt = USD if header.endswith("_usd") else PEN if header.endswith("_pen") else \
            PCT if header in ("cpu_use", "mem_use", "use", "system_mem_share", "aca_vs_aks") else None
        if fmt:
            for row in range(2, count + 2):
                sheet.cell(row=row, column=index).number_format = fmt
    column = cwb.get_column_letter(headers.index("class") + 1)
    rows_ref = f"A2:{cwb.get_column_letter(len(headers))}{count + 1}"
    for value, fill in (("OPTIMIZED", GREEN), ("PARTLY_USED", AMBER), ("OVERSIZED", RED)):
        sheet.conditional_formatting.add(rows_ref, FormulaRule(formula=[f'${column}2="{value}"'], fill=fill))
    cwb.widths(sheet, headers)
    sheet.freeze_panes = "C2"


def write_summary(sheet, month):
    sheet["A1"], sheet["A1"].font = "AKS right-sizing: how well each cluster uses what it pays", Font(bold=True, size=14)
    sheet["A2"] = (f"Cost of {month} in soles with IGV. Use: average CPU or memory used in office hours over what the "
                   "nodes offer (the busier of the two). PROD business clusters unless said otherwise.")
    business = 'Clusters[environment],"PROD",Clusters[category],"APPLICATION"'
    sheet["A4"], sheet["A4"].font = "PROD business clusters by use", BOLD
    headers = ["Class", "Clusters", "AKS today (S/ a month)", "Paid and not used (S/)",
               "Container Apps (v2, S/)", "Container Apps vs AKS"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=5, column=column, value=header).font = BOLD
    classes = [("OPTIMIZED", "Optimized (use >= optimized_use)"), ("PARTLY_USED", "Partly used"),
               ("OVERSIZED", "Oversized (use < partly_use)"), ("NO_DATA", "No usage data")]
    for row, (code, label) in enumerate(classes, start=6):
        where = f'{business},Clusters[class],"{code}"'
        sheet.cell(row=row, column=1, value=label)
        sheet.cell(row=row, column=2, value=f"=COUNTIFS({where})")
        sheet.cell(row=row, column=3, value=f"=SUMIFS(Clusters[aks_month_pen],{where})").number_format = PEN
        sheet.cell(row=row, column=4, value=f"=SUMIFS(Clusters[idle_pen],{where})").number_format = PEN
        # Container Apps against AKS on the same clusters (those the v2 model costs)
        aca = f"SUMIFS(Clusters[aca_month_pen],{where})"
        aks = f'SUMIFS(Clusters[aks_month_pen],{where},Clusters[aca_month_pen],">0")'
        sheet.cell(row=row, column=5, value=f"={aca}").number_format = PEN
        sheet.cell(row=row, column=6, value=f'=IF({aks}>0,{aca}/{aks},"")').number_format = PCT
    sheet.cell(row=10, column=1, value="Total").font = BOLD
    for column, fmt in ((2, None), (3, PEN), (4, PEN), (5, PEN)):
        letter = cwb.get_column_letter(column)
        cell = sheet.cell(row=10, column=column, value=f"=SUM({letter}6:{letter}9)")
        cell.font = BOLD
        if fmt:
            cell.number_format = fmt
    sheet["A11"] = ("Container Apps vs AKS above 100%: Container Apps would cost more. It falls as the cluster is used "
                    "worse; a well-used AKS stays cheaper.")
    for column, width in zip("ABCDEF", (34, 10, 22, 22, 22, 22)):
        sheet.column_dimensions[column].width = width


def write_pains(sheet):
    sheet["A1"], sheet["A1"].font = "Where the AKS money goes (S/ a month, with IGV)", Font(bold=True, size=14)
    prod_business = 'Clusters[environment],"PROD",Clusters[category],"APPLICATION"'
    pains = [
        ("Dedicated cluster at the minimum size", "PROD business clusters with min_cluster_nodes nodes or fewer: "
         "they pay the floor whatever they run",
         f'COUNTIFS({prod_business},Clusters[at_minimum],"YES")',
         f'SUMIFS(Clusters[aks_month_pen],{prod_business},Clusters[at_minimum],"YES")'),
        ("Capacity paid and not used", "AKS cost x (1 - use), PROD business clusters",
         f"COUNTIFS({prod_business},Clusters[use],\">=0\")", f"SUMIFS(Clusters[idle_pen],{prod_business})"),
        ("Empty clusters", "PROD clusters with no application running",
         'COUNTIFS(Clusters[environment],"PROD",Clusters[category],"NO_APPLICATIONS")',
         'SUMIFS(Clusters[aks_month_pen],Clusters[environment],"PROD",Clusters[category],"NO_APPLICATIONS")'),
        ("Non-production on 24/7", "CERT and DESA clusters of business applications, nodes on all month",
         'COUNTIFS(Clusters[environment],"<>PROD",Clusters[category],"APPLICATION")',
         'SUMIFS(Clusters[aks_month_pen],Clusters[environment],"<>PROD",Clusters[category],"APPLICATION")'),
        ("Autoscaling that does not move", "Microservice HPAs of PROD business clusters with no scaling for 3 months "
         "or more (count, not money)", f"SUMIFS(Clusters[hpas_still_3_months],{prod_business})", None)]
    for column, header in enumerate(["Pain point", "What it counts", "Clusters / HPAs", "S/ a month"], start=1):
        sheet.cell(row=3, column=column, value=header).font = BOLD
    for row, (label, meaning, count, money) in enumerate(pains, start=4):
        sheet.cell(row=row, column=1, value=label)
        sheet.cell(row=row, column=2, value=meaning)
        sheet.cell(row=row, column=3, value="=" + count)
        if money:
            sheet.cell(row=row, column=4, value="=" + money).number_format = PEN
    sheet["A10"] = ("The rows overlap (a cluster at the minimum also has idle capacity): read each on its own, do not "
                    "add them up.")
    sheet["A12"], sheet["A12"].font = "System pool", BOLD
    sheet["A13"] = "Share of the memory used in PROD business clusters that the system pool takes (office hours)"
    sheet["D13"] = (f'=IF(SUMIFS(Clusters[mem_used_office_avg_gib],{prod_business})>0,SUMIFS('
                    f'Clusters[mem_used_system_office_avg_gib],{prod_business})/SUMIFS('
                    f'Clusters[mem_used_office_avg_gib],{prod_business}),"")')
    sheet["D13"].number_format = PCT
    for column, width in zip("ABCD", (38, 80, 16, 18)):
        sheet.column_dimensions[column].width = width


def write_breakeven(sheet):
    sheet["A1"], sheet["A1"].font = "Break-even: one vCPU on a reserved AKS node against Container Apps", \
        Font(bold=True, size=14)
    lines = [("AKS, US$ per vCPU-hour (reserved node)", "=vm_hour_reserved/vm_vcpus", '"$"0.0000'),
             ("Container Apps, US$ per vCPU-hour (with its GiB and the bank discount)",
              "=3600*(aca_vcpu_second+aca_gib_per_vcpu*aca_gib_second)*(1-bank_discount)", '"$"0.0000'),
             ("Container Apps costs this many times AKS", "=B4/B3", "0.0"),
             ("Container Apps is cheaper only if a vCPU is on less than this share of the time", "=B3/B4", PCT),
             ("...that is, hours a day", "=24*B3/B4", "0.0")]
    for row, (label, formula, fmt) in enumerate(lines, start=3):
        sheet.cell(row=row, column=1, value=label)
        sheet.cell(row=row, column=2, value=formula).number_format = fmt
    sheet["A9"] = ("Per vCPU actually needed. A dedicated AKS cluster also pays its minimum nodes, system nodes, disks and "
                   "tier, which is why small applications still come out cheaper on Container Apps.")
    sheet.column_dimensions["A"].width = 80
    sheet.column_dimensions["B"].width = 14


def write_calculator(sheet, default):
    """One cluster at a time: what its nodes offer, what it uses (applications and system), how many nodes would be
    enough at optimized_use, and what that costs against Container Apps."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.worksheet.datavalidation import DataValidation
    sheet["A1"], sheet["A1"].font = "Calculator: how much of its nodes a cluster uses", Font(bold=True, size=14)
    sheet["A3"], sheet["A3"].font = "Cluster", BOLD
    sheet["B3"], sheet["B3"].fill = default, EDITABLE
    validation = DataValidation(type="list", formula1="=INDIRECT(\"Clusters[cluster]\")", allow_blank=False)
    sheet.add_data_validation(validation)
    validation.add("B3")
    get = lambda column: f'IFERROR(INDEX(Clusters[{column}],MATCH($B$3,Clusters[cluster],0)),"")'
    facts = [("Application", get("cod_app"), None), ("Environment / category", f'{get("environment")}&" / "&{get("category")}', None),
             ("Nodes: system + user", f'{get("nodes_system")}&" + "&{get("nodes_user")}', None),
             ("VM sizes", get("vm_sizes"), None), ('"Micro"-services', get("microservices"), "0"),
             ("AKS cost (S/ a month, with IGV)", get("aks_month_pen"), PEN),
             ("Container Apps, criteria v2 (S/ a month)", get("aca_month_pen"), PEN),
             ("Criteria v2 decision", get("aca_group"), None),
             ("Use (the busier of CPU and memory)", get("use"), PCT), ("Class", get("class"), None)]
    for row, (label, formula, fmt) in enumerate(facts, start=5):
        sheet.cell(row=row, column=1, value=label)
        cell = sheet.cell(row=row, column=2, value="=" + formula)
        if fmt:
            cell.number_format = fmt
    top = 17
    sheet.cell(row=top, column=1, value="Office hours (weekdays 8:00-20:00)").font = BOLD
    for column, header in enumerate(["", "CPU (vCPU)", "Memory (GiB)"], start=1):
        sheet.cell(row=top + 1, column=column, value=header).font = BOLD
    lines = [("Nodes offer (allocatable)", "cpu_allocatable", "mem_allocatable_gib"),
             ("Used, peak of the month", "cpu_used_max", "mem_used_max_gib"),
             ("Used, average in office hours", "cpu_used_office_avg", "mem_used_office_avg_gib"),
             ("...of it, the system pool (Kubernetes itself)", "cpu_used_system_office_avg", "mem_used_system_office_avg_gib"),
             ("Used, average off hours", "cpu_used_offhours_avg", "mem_used_offhours_avg_gib")]
    for offset, (label, cpu, mem) in enumerate(lines, start=top + 2):
        sheet.cell(row=offset, column=1, value=label)
        sheet.cell(row=offset, column=2, value="=" + get(cpu)).number_format = "0.0"
        sheet.cell(row=offset, column=3, value="=" + get(mem)).number_format = "0.0"
    used_by_apps = top + 7
    sheet.cell(row=used_by_apps, column=1, value="Used by the applications (office average minus system)")
    sheet.cell(row=used_by_apps, column=2, value=f"=MAX(0,N(B{top + 4})-N(B{top + 5}))").number_format = "0.0"
    sheet.cell(row=used_by_apps, column=3, value=f"=MAX(0,N(C{top + 4})-N(C{top + 5}))").number_format = "0.0"
    share = top + 8
    sheet.cell(row=share, column=1, value="Share of what the nodes offer, office average")
    sheet.cell(row=share, column=2, value=f'=IF(N(B{top + 2})>0,N(B{top + 4})/B{top + 2},"")').number_format = PCT
    sheet.cell(row=share, column=3, value=f'=IF(N(C{top + 2})>0,N(C{top + 4})/C{top + 2},"")').number_format = PCT
    fit = top + 11
    sheet.cell(row=fit, column=1, value="If it were sized to optimized_use").font = BOLD
    nodes = f"(N({get('nodes_system')})+N({get('nodes_user')}))"
    rows = [("Nodes today", f"={nodes}", "0"),
            ("Nodes enough at optimized_use (never below min_cluster_nodes)",
             f'=IF({nodes}>0,MAX(min_cluster_nodes,CEILING({nodes}*N({get("use")})/optimized_use,1)),"")', "0"),
            ("AKS cost with those nodes (S/ a month)", f'=IF(N(B{fit + 1})>0,N({get("aks_month_pen")})*B{fit + 2}/B{fit + 1},"")', PEN),
            ("AKS today (S/ a month)", "=N(" + get("aks_month_pen") + ")", PEN),
            ("Container Apps, criteria v2 (S/ a month)", "=" + get("aca_month_pen"), PEN)]
    for offset, (label, formula, fmt) in enumerate(rows, start=fit + 1):
        sheet.cell(row=offset, column=1, value=label)
        sheet.cell(row=offset, column=2, value=formula).number_format = fmt
    notes = top + 18
    sheet.cell(row=notes, column=1, value="How it is measured").font = BOLD
    for offset, text in enumerate((
            "Use: CPU and memory used on the cluster's nodes (Azure Monitor node metrics of the month), averaged over "
            "office hours, over what the nodes offer to pods (allocatable). The busier of the two decides.",
            "Cost: Cost Management, amortized cost of the month for the cluster's resources (reservations spread over "
            "their term), in soles with IGV.",
            "Nodes enough: today's nodes x use / optimized_use, rounded up, never below the dedicated minimum. A cluster "
            "at the minimum cannot get cheaper inside AKS: only sharing a cluster or Container Apps removes that floor.",
            "Container Apps: the criteria v2 model (replicas from the traffic of each hour, 1 minimum in office hours, 0 "
            "outside), only for PROD business clusters."), start=1):
        sheet.cell(row=notes + offset, column=1, value=text)
    chart = BarChart()
    chart.type, chart.grouping, chart.title = "bar", "clustered", None
    chart.y_axis.title = None
    data = Reference(sheet, min_col=2, max_col=3, min_row=top + 1, max_row=top + 1 + 4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=top + 2, max_row=top + 2 + 3))
    chart.height, chart.width = 8, 16
    sheet.add_chart(chart, "E3")
    for column, width in zip("ABC", (62, 18, 16)):
        sheet.column_dimensions[column].width = width


def build(results, output):
    data, month = rows(results)
    out_dir = os.path.join(results, FOLDER)
    os.makedirs(out_dir, exist_ok=True)
    aa.azure.Output(out_dir, ",").csv(FILE, data)
    book = Workbook()
    names = ["Summary", "Calculator", "Pain points", "Clusters", "Breakeven", "Variables", "Data"]
    sheets = {name: (book.active if index == 0 else book.create_sheet(name)) for index, name in enumerate(names)}
    sheets["Summary"].title = "Summary"
    write_variables(book, sheets["Variables"], v1.read_prices(results))
    values = [[row[column] if column in TEXT else cwb.number(row[column]) for column in COLUMNS] for row in data]
    cwb.write_query_table(sheets["Data"], "ClusterData", COLUMNS, values)
    write_clusters(sheets["Clusters"], len(values))
    write_summary(sheets["Summary"], month)
    write_pains(sheets["Pain points"])
    write_breakeven(sheets["Breakeven"])
    write_calculator(sheets["Calculator"], next((row["cluster"] for row in data if row["cod_app"] == "BSEG"),
                                                data[0]["cluster"] if data else ""))
    sheets["Data"].sheet_state = "hidden"
    book.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    aa.keep_version(output)
    book.save(output)
    pq.load_queries(output, [("Data", "ClusterData", "ClusterData", query(os.path.abspath(results)), COLUMNS)])
    aa.azure.LOG.info("Wrote %s (%d clusters, Power Query connection to %s)", output, len(data), FILE)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: aks_rightsizing.xlsx in the newest reports/<date>-aks-aca")
    args = parser.parse_args()
    aa.azure.setup_logging(False)
    output = args.output
    if not output:
        folders = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
        output = os.path.join(folders[-1] if folders else os.path.join(aa.TOPIC_DIR, "reports", "aks-aca"),
                              "aks_rightsizing.xlsx")
    build(args.results, output)


if __name__ == "__main__":
    main()
