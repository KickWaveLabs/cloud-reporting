#!/usr/bin/env python3
"""Write the criteria workbook v1, v2 or v3 (--criteria; v2 by default; v1 runs aa_criteria_workbook_v1.py, per application; v3 is v2 with the PRESETS["v3"] defaults of aa_v2_model.py): one row per PROD application cluster, a summary per application, the traffic of
each hour and a calculator where the replicas of every hour can be chosen.

Same rules and cost model as aa_v2_model.py, written as Excel formulas over the same inputs (the module's docstring
explains them); with the default variables both give the same figures. Sheets:
  Results     clusters, AKS today, with Container Apps and saving per destination, and a chart
  Criteria    one row per cluster: data, criteria, decision, Container Apps cost from traffic, saving, group
  Top costs   the most expensive PROD application clusters with what each would save on Container Apps
              (highlighted against min_saving), and beside it the most expensive AKS clusters of every environment
              and category with the share PROD application clusters take of the whole AKS bill
  Apps        one row per application: its clusters, how many move, cost today and after
  Calculator  pick a cluster or an application (B2): traffic by hour (min / average / peak, weekday and weekend, as
              charts), replicas per "micro"-service each hour (today; simulated by peak and by average; chosen:
              editable) and the monthly cost meter by meter on AKS and on Container Apps, by peak, by average and
              chosen
  Latency     the UX -> BS -> Core -> mainframe chain: how much a request takes with each layer on AKS or Container
              Apps for the chosen application, a drawing of the chain and the measured latency per application;
              every parameter can be changed (to validate with the architecture team)
  Traffic     scope (cluster or application) x weekday/weekend x hour: TPS and simulated replicas
  Services    one row per "micro"-service: application, cluster, request and replicas today
  Prices      list prices with their links
  Variables   thresholds, saving margin, sizing rule, minimum replicas, discount, IGV, exchange rate
  Lists       the scopes the calculator offers
  Data, TrafficData, LatencyData, LatencyLayers, Lists, Services and the list in Apps are Power Query connections to
              the CSV files in results/ (Data > Refresh All reads them again; the folder is the first step of each
              query). Formulas live in their own tables (Criteria, Traffic, Apps, LatencyApps) and read the query
              tables row by row: after a refresh that adds rows, drag the corner of those tables down.
Reads only results/. Output: aks_aca_criteria_v2.xlsx in the newest reports/<date>-aks-aca folder (or --output).
"""

import argparse
import glob
import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import Legend
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

import aa_common as aa
import aa_criteria_workbook_v1 as v1wb
import aa_deck_exec as deck_entry
import aa_deck_exec_v1 as deck
import aa_k8s_snapshot as snap
import aa_powerquery as pq
import aa_v2_model as model

BOLD = Font(bold=True)
EDITABLE = PatternFill("solid", fgColor="FFF2CC")
USD, PEN, PCT = '"$"#,##0.00', '"S/ "#,##0', "0%"

CLUSTER_COLUMNS = ["cluster", "cod_app", "clusters_of_app", "bia", "microservices", "microservices_app",
                   "microservices_app_projected", "microservices_over_aca_size", "traffic_status", "traffic_layer",
                   "requests_month", "busiest_hour_tps", "peak_tps", "event_driven", "cost_month_usd",
                   "no_traffic_reason", "replicas_today_total", "month"]
CLUSTER_NUMBERS = {"clusters_of_app", "microservices", "microservices_app", "microservices_app_projected",
                   "microservices_over_aca_size", "requests_month", "busiest_hour_tps", "peak_tps", "cost_month_usd",
                   "replicas_today_total"}
TRAFFIC_COLUMNS = ["scope", "scope_type", "day_type", "hour", "office", "days", "tps_min", "tps_avg", "tps_max_hour",
                   "tps_peak", "replicas_today_avg", "replicas_today_max", "off_5min_min", "off_5min_avg", "tps_per_replica_today"] + \
    model.BUSIEST_COLUMNS
SERVICE_COLUMNS = ["cluster", "cod_app", "namespace", "deployment", "cpu_request_vcpu", "mem_request_gib",
                   "replicas_today_office", "cpu_used_vcpu", "mem_used_gib"]
GROUPS = [("AKS", "Stays on AKS (criteria)"), ("BELOW_MARGIN", "Meets criteria, saves less than the margin"),
          ("ACA_WITH_REVIEW", "ACA, with review"), ("ACA", "Candidate to ACA"),
          ("IN_USE", "In use, not measured"), ("NOT_IN_USE", "Not in use"), ("IN_DEVELOPMENT", "In development")]
LAYERS = ["UX", "BS", "CORE", "MAINFRAME"]
ALL_CLUSTER_COLUMNS = ["cluster", "environment", "category", "cod_app", "cost_month_usd"]
SCALE_TO_ZERO_NOTE = [
    "Scale to zero (a minimum of 0 replicas): Container Apps turns a \"micro\"-service off after 5 minutes "
    "without requests and on again with the next one (a few seconds of cold start).",
    "So an hour with traffic only part of the time is billed only for its 5-minute buckets with traffic or right "
    "after it. By peak takes that hour's day with the fewest buckets off; by average, the average day; by busiest "
    "hour, its own buckets. With a minimum of 1 the hour is billed whole."]
HIDDEN_SHEETS = ["Latency", "Traffic", "Data", "TrafficData", "LatencyData", "LatencyLayers", "AllClusters", "Lists",
                 "UsageData"]
USAGE_COLUMNS = ["cluster", "cod_app", "day_type", "hour", "cpu_used_avg", "cpu_system_avg", "cpu_allocatable",
                 "mem_used_avg_gib", "mem_system_avg_gib", "mem_allocatable_gib"]
USAGE_COLUMN = 50  # Calculator: column AX, the CPU and memory of each hour, right of their charts (9 columns)
TOP = 20  # rows of each Top costs table
GREEN, AMBER = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFEB9C")
BLOCK = 8  # Calculator: columns per day type (TPS min, average, peak, today, by peak, by busiest hour, by average,
#            chosen)
FIVE_MINUTES_COLUMN = 60  # Calculator: column BH, right of the CPU and memory table
HOUR_LABEL_STEP = 4  # hours 0, 4, 8 ... 20 on the charts: office hours are 8:00 to 19:59
LISTS = {"sizing_basis": '"PEAK,BUSIEST_HOUR,AVERAGE"', "capacity_from_today": '"NO,YES"',
         "size_from_usage": '"NO,YES"', "environment_fee": '"NO,YES"',
         "aks_from_bia": '"NONE,MUY ALTA,ALTA,MEDIA,BAJA"', "rule_event_driven": '"YES,NO"',
         "rule_shared": '"YES,NO"', "rule_high_load": '"YES,NO"', "rule_constant_load": '"YES,NO"'}
BIA_ORDER = '{"BAJA","MEDIA","ALTA","MUY ALTA"}'  # lowest first; SIN BIA and NONE match none  # variables picked from a list
BUCKETS = 12  # 5-minute buckets in an hour
# Official sources of the sizing rule, linked next to their variables.
SOURCES = {
    "concurrent_requests": "https://learn.microsoft.com/azure/container-apps/scale-app#http",
    "min_replicas_office": "https://learn.microsoft.com/azure/architecture/example-scenario/serverless/"
                           "microservices-with-container-apps-dapr#considerations",
    "min_replicas_offhours": "https://learn.microsoft.com/azure/architecture/example-scenario/serverless/"
                             "microservices-with-container-apps-dapr#considerations",
}


def number(value):
    return model.to_float(value) if value not in ("", None) else None


def write_variables(book, sheet, month, days, price_env_hour):
    blocks = [
        ("Criteria", [("max_tps", model.DEFAULTS["max_tps"], "Busiest hour up to this TPS (average of the hour)"),
                      ("max_services", model.DEFAULTS["max_services"], "Up to this many \"micro\"-services per cluster"),
                      ("aks_from_bia", model.DEFAULTS["aks_from_bia"],
                       "Clusters with this BIA or above stay on AKS (MUY ALTA > ALTA > MEDIA > BAJA); NONE: no cluster"),
                      ("rule_event_driven", model.DEFAULTS["rule_event_driven"],
                       "YES: event-driven clusters count as in use without measurement (their messages are not "
                       "measured as load). NO: evaluated as the rest, messages per second taken as TPS"),
                      ("rule_shared", model.DEFAULTS["rule_shared"],
                       "YES: shared clusters (shared_apps, Core layer) stay on AKS. NO: evaluated as the rest"),
                      ("rule_high_load", model.DEFAULTS["rule_high_load"],
                       "YES: high or constant load stays on AKS. NO: evaluated as the rest"),
                      ("max_avg_tps", model.DEFAULTS["max_avg_tps"],
                       "High load: more than this TPS on average over the day (the month's requests over its seconds)"),
                      ("min_quiet_share", model.DEFAULTS["min_quiet_share"],
                       "Constant load: its quietest weekday hour carries at least this share of the weekday average..."),
                      ("min_constant_tps", model.DEFAULTS["min_constant_tps"],
                       "...with at least this TPS on average over the day (one replica busy all day)"),
                      ("rule_constant_load", model.DEFAULTS["rule_constant_load"],
                       "YES: constant load (max_replicas_vs_aks) stays on AKS. NO (default): evaluated as the rest"),
                      ("max_replicas_vs_aks", model.DEFAULTS["max_replicas_vs_aks"],
                       "High, constant load: clusters where Container Apps would run, on average, at least this share "
                       "of the replicas they run today on AKS stay on AKS"),
                      ("shared_apps", model.DEFAULTS["shared_apps"],
                       "Shared applications (authentication and identity, comma-separated) stay on AKS, as does "
                       "every cluster whose traffic is on the Core layer")]),
        ("Saving filter (applied at the end, not a criterion)",
         [("min_saving", model.DEFAULTS["min_saving"], "A cluster that meets the criteria moves only if it saves at "
                                                        "least this share of its AKS cost")]),
        ("Replicas from traffic", [
            ("sizing_basis", model.DEFAULTS["sizing_basis"], "PEAK: busiest 5 minutes of each hour; BUSIEST_HOUR: "
                                                             "each hour's busiest occurrence in the month, 5 minutes "
                                                             "at a time; AVERAGE: the hour's average"),
            ("concurrent_requests", model.DEFAULTS["concurrent_requests"],
             "Requests per second per replica before Container Apps adds one (HTTP scale rule; Azure default 10)"),
            ("traffic_share", model.DEFAULTS["traffic_share"], "Share of the cluster's TPS each \"micro\"-service "
                                                               "receives (1 = all of it, conservative)"),
            ("min_replicas_office", model.DEFAULTS["min_replicas_office"], "Minimum replicas per \"micro\"-service, "
                                                                           "weekdays 8:00-20:00 (1 avoids the cold "
                                                                           "start)"),
            ("min_replicas_offhours", model.DEFAULTS["min_replicas_offhours"], "Minimum replicas per "
                                                                               "\"micro\"-service, the rest. 0 lets "
                                                                               "it scale to zero: see the note below"),
            ("capacity_from_today", model.DEFAULTS["capacity_from_today"],
             "YES: TPS per replica = what one replica handles today on AKS (peak hour) x capacity_safety, never below "
             "concurrent_requests (criteria v3)"),
            ("capacity_safety", model.DEFAULTS["capacity_safety"], "Share of today's TPS per replica taken as capacity")]),
        ("Container Apps", [("requests_buffer", model.DEFAULTS["requests_buffer"], "Extra on the requests of the month"),
                            ("environments_per_cluster", model.DEFAULTS["environments_per_cluster"],
                             "Environments (one load balancer each) per cluster moved"),
                            ("hours_month", model.DEFAULTS["hours_month"], "Hours billed for the load balancer"),
                            ("min_vcpu", model.DEFAULTS["min_vcpu"], "Smallest replica in vCPU"),
                            ("vcpu_step", model.DEFAULTS["vcpu_step"], "Sizes go up in steps of this many vCPU"),
                            ("gib_per_vcpu", model.DEFAULTS["gib_per_vcpu"], "GiB per vCPU"),
                            ("size_from_usage", model.DEFAULTS["size_from_usage"],
                             "YES: each \"micro\"-service's CPU and memory = its use per replica today x "
                             "usage_headroom, never above its request (criteria v3)"),
                            ("usage_headroom", model.DEFAULTS["usage_headroom"], "Times today's use kept as size"),
                            ("environment_fee", model.DEFAULTS["environment_fee"],
                             "YES: add the Dedicated plan management fee per environment (private endpoint or planned "
                             "maintenance; criteria v3)"),
                            ("price_env_hour", price_env_hour, "Dedicated plan management, US$ per environment hour "
                                                               "(Azure Retail Prices)")]),
        ("Money", [("bank_discount", model.DEFAULTS["bank_discount"], "Bank's discount on Azure list prices"),
                   ("igv", model.DEFAULTS["igv"], "IGV"), ("exchange_rate", model.DEFAULTS["exchange_rate"],
                                                         "Soles per US$")]),
        ("Time", [("data_month", month, "Month of the data"),
                  ("weekdays", days["weekday"], "Weekdays in the month"),
                  ("weekend_days", days["weekend"], "Weekend days in the month")]),
    ]
    sheet["A1"], sheet["A1"].font = "Variables (yellow: change them)", Font(bold=True, size=14)
    row = 3
    for title, items in blocks:
        sheet.cell(row=row, column=1, value=title).font = BOLD
        row += 1
        for label, value, comment in items:
            sheet.cell(row=row, column=1, value=label)
            cell = sheet.cell(row=row, column=2, value=value)
            cell.fill = EDITABLE
            if label in ("bank_discount", "igv", "min_saving", "traffic_share", "capacity_safety",
                         "max_replicas_vs_aks", "min_quiet_share"):
                cell.number_format = PCT
            sheet.cell(row=row, column=3, value=comment)
            if label in SOURCES:
                link = sheet.cell(row=row, column=4, value="Microsoft Learn")
                link.hyperlink, link.style = SOURCES[label], "Hyperlink"
            v1wb.define(book, label, f"Variables!$B${row}")
            if label in LISTS:
                validation = DataValidation(type="list", formula1=LISTS[label], allow_blank=False)
                sheet.add_data_validation(validation)
                validation.add(cell.coordinate)
            row += 1
        row += 1
    sheet.cell(row=row, column=1, value="Note").font = BOLD
    for offset, text in enumerate(SCALE_TO_ZERO_NOTE, start=1):
        sheet.cell(row=row + offset, column=1, value=text)
    for column, width in zip("ABCD", (26, 12, 90, 16)):
        sheet.column_dimensions[column].width = width




def r(table, column):
    return f"{table}[[#This Row],[{column}]]"


def pull(query, column, top=1):
    """A formula table's cell: the same row of a query table (Excel drops formula columns inside a query table, so
    formulas live in their own table and read the query table row by row)."""
    value = f"INDEX({query}[{column}],ROW()-{top})"
    return f'IF(ISBLANK({value}),"",{value})'


def formula_table(sheet, name, formulas, rows, top=1, left=1):
    """A table of formula columns only, at any position, with one row per row of the query table it reads."""
    for offset, (header, formula) in enumerate(formulas):
        sheet.cell(row=top, column=left + offset, value=header)
        for index in range(rows):
            sheet.cell(row=top + 1 + index, column=left + offset, value="=" + formula)
    ref = f"{get_column_letter(left)}{top}:{get_column_letter(left + len(formulas) - 1)}{top + max(rows, 1)}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name=v1wb.STYLE, showRowStripes=True)
    table.tableColumns = [TableColumn(id=index, name=header) for index, (header, _) in enumerate(formulas, start=1)]
    for column, (_, formula) in zip(table.tableColumns, formulas):
        column.calculatedColumnFormula = TableFormula(attr_text=formula)
    sheet.add_table(table)


# ---------------------------------------------------------------------------
# Rows each Power Query query returns (written as the table's first data; Data > Refresh All reads the CSVs again)
# ---------------------------------------------------------------------------


def cluster_rows(clusters):
    return [[number(row[c]) if c in CLUSTER_NUMBERS else (row[c] or None) for c in CLUSTER_COLUMNS]
            for row in sorted(clusters, key=lambda item: (item["cod_app"], item["cluster"]))]


def traffic_rows(profile):
    return [[row[c] if c in ("scope", "scope_type", "day_type") else number(row[c]) for c in TRAFFIC_COLUMNS]
            for row in profile]


def service_rows(services, clusters):
    code = {row["cluster_id"]: row["cod_app"] for row in clusters}
    rows = [[row["cluster"], code[row["cluster_id"]], row["namespace"], row["deployment"],
             number(row["cpu_request_vcpu"]), number(row["mem_request_gib"]), number(row["ci_replicas_avg_office"]),
             number(row["ci_cpu_used_per_replica"]), number(row["ci_mem_used_gib_per_replica"])]
            for row in services if row["environment"] == "PROD"]
    return sorted(rows, key=lambda values: (values[0], values[3]))


def all_cluster_rows(results):
    """Every AKS cluster with a cost in the cost month (any environment or category), the most expensive first."""
    read = lambda folder, name: aa.read_csv(os.path.join(results, folder, name))
    clusters = read("aa-k8s-snapshot", "cluster_snapshot.csv")
    _, cost = snap.app_cost(read("aa-cost", "cost_by_resource.csv"), {row["cluster_id"].lower() for row in clusters})
    rows = [[row["cluster"], row["environment"], row["category"], row["cod_app"] or None,
             round(cost[row["cluster_id"].lower()], 2)] for row in clusters if row["cluster_id"].lower() in cost]
    return sorted(rows, key=lambda values: (-values[4], values[0]))


def latency_inputs(results):
    """Per API of the month: (cod_app, layer, requests, backend ms, gateway ms), the join the queries also make."""
    requests = {(row["apim"], row["api_id"], row["month"]): number(row["requests"])
                for row in aa.read_csv(os.path.join(results, "aa-metrics", "apim_api_monthly.csv"))}
    out = []
    for row in aa.read_csv(os.path.join(results, "aa-signals", "apim_latency_monthly.csv")):
        weight = requests.get((row["apim"], row["api_id"], row["month"]))
        backend = number(row.get("backend_duration_avg_ms"))
        if row.get("cod_app") and weight and backend is not None:
            gateway = number(row.get("duration_avg_ms"))
            out.append((row["cod_app"], row["apim_layer"], weight, backend, backend if gateway is None else gateway))
    return out


def latency_app_rows(inputs):
    sums = defaultdict(lambda: [0.0, 0.0, 0.0])
    for code, layer, weight, backend, gateway in inputs:
        entry = sums[(code, layer)]
        entry[0] += weight
        entry[1] += weight * backend
        entry[2] += weight * gateway
    rows = []
    for code in sorted({code for code, _ in sums}):
        ux, bs = sums.get((code, "UX")), sums.get((code, "BS"))
        if ux and bs and ux[0] >= MIN_UX_REQUESTS and bs[0] > 0:
            rows.append([code, ux[0], ux[1] / ux[0], bs[0], bs[1] / bs[0], bs[2] / bs[0]])
    return rows


def latency_layer_rows(inputs):
    sums = defaultdict(lambda: [0.0, 0.0, 0.0])
    for _, layer, weight, backend, gateway in inputs:
        sums[layer][0] += weight
        sums[layer][1] += weight * backend
        sums[layer][2] += weight * gateway
    return [[layer, entry[0], entry[1] / entry[0], entry[2] / entry[0]] for layer, entry in sorted(sums.items())]


# Applications with fewer UX requests in the month are left out: a handful of calls gives meaningless ratios.
MIN_UX_REQUESTS = 1000
LATENCY_APP_COLUMNS = ["cod_app", "ux_requests", "ux_backend_ms", "bs_requests", "bs_backend_ms", "bs_gateway_ms"]
LATENCY_LAYER_COLUMNS = ["apim_layer", "requests", "backend_ms", "gateway_ms"]


# ---------------------------------------------------------------------------
# Power Query M code (one query per table; the folder is the first step, change it on another machine)
# ---------------------------------------------------------------------------


def m_list(columns):
    return ", ".join(f'"{c}"' for c in columns)


def m_types(columns):
    return ", ".join(f'{{"{c}", type number}}' for c in columns)


def m_typed(source, numbers):
    return (f'Table.TransformColumnTypes(Table.ReplaceValue({source}, "", null, Replacer.ReplaceValue, '
            f'{{{m_list(numbers)}}}), {{{m_types(numbers)}}}, "en-US")')


def query_clusters(folder):
    numbers = [c for c in CLUSTER_COLUMNS if c in CLUSTER_NUMBERS]
    return v1wb.query_head(folder) + f'''    Clusters = ReadCsv("aa-cluster-snapshot", "cluster_criteria_input.csv"),
    Columns = Table.SelectColumns(Clusters, {{{m_list(CLUSTER_COLUMNS)}}}),
    Typed = {m_typed("Columns", numbers)},
    Sorted = Table.Sort(Typed, {{{{"cod_app", Order.Ascending}}, {{"cluster", Order.Ascending}}}})
in
    Sorted'''


def query_usage(folder):
    numbers = [c for c in USAGE_COLUMNS if c not in ("cluster", "cod_app", "day_type")]
    return v1wb.query_head(folder) + f'''    Usage = ReadCsv("aa-metrics", "aks_usage_hourly_profile.csv"),
    Columns = Table.SelectColumns(Usage, {{{m_list(USAGE_COLUMNS)}}}),
    Typed = {m_typed("Columns", numbers)}
in
    Typed'''


def usage_rows(results):
    rows = aa.read_csv(os.path.join(results, "aa-metrics", "aks_usage_hourly_profile.csv"))
    return [[number(row[c]) if c not in ("cluster", "cod_app", "day_type") else row[c] for c in USAGE_COLUMNS]
            for row in rows]


def query_traffic(folder):
    numbers = [c for c in TRAFFIC_COLUMNS if c not in ("scope", "scope_type", "day_type")]
    return v1wb.query_head(folder) + f'''    Profile = ReadCsv("aa-cluster-snapshot", "traffic_hourly_profile.csv"),
    Columns = Table.SelectColumns(Profile, {{{m_list(TRAFFIC_COLUMNS)}}}),
    Typed = {m_typed("Columns", numbers)}
in
    Typed'''


def query_services(folder):
    numbers = ["cpu_request_vcpu", "mem_request_gib", "replicas_today_office", "cpu_used_vcpu", "mem_used_gib"]
    return v1wb.query_head(folder) + f'''    Deployments = ReadCsv("aa-k8s-snapshot", "deployment_snapshot.csv"),
    Clusters = Table.SelectColumns(ReadCsv("aa-cluster-snapshot", "cluster_criteria_input.csv"), {{"cluster_id", "cod_app"}}),
    // Each PROD "micro"-service of a cluster of the v2 criteria, with the cluster's application
    MicroServices = Table.RemoveColumns(Table.SelectRows(Deployments, each [microservice] = "YES" and [environment] = "PROD"), {{"cod_app"}}, MissingField.Ignore),
    Joined = Table.ExpandTableColumn(Table.NestedJoin(MicroServices, {{"cluster_id"}}, Clusters, {{"cluster_id"}}, "c", JoinKind.Inner), "c", {{"cod_app"}}),
    Renamed = Table.RenameColumns(Joined, {{{{"ci_replicas_avg_office", "replicas_today_office"}}, {{"ci_cpu_used_per_replica", "cpu_used_vcpu"}}, {{"ci_mem_used_gib_per_replica", "mem_used_gib"}}}}),
    Columns = Table.SelectColumns(Renamed, {{{m_list(SERVICE_COLUMNS)}}}),
    Typed = {m_typed("Columns", numbers)},
    Sorted = Table.Sort(Typed, {{{{"cluster", Order.Ascending}}, {{"deployment", Order.Ascending}}}})
in
    Sorted'''


def latency_head(folder):
    return v1wb.query_head(folder) + '''    Latency = ReadCsv("aa-signals", "apim_latency_monthly.csv"),
    Requests = Table.SelectColumns(ReadCsv("aa-metrics", "apim_api_monthly.csv"), {"apim", "api_id", "month", "requests"}),
    // Each API's latency weighted by its requests of the same month
    Joined = Table.ExpandTableColumn(Table.NestedJoin(Latency, {"apim", "api_id", "month"}, Requests, {"apim", "api_id", "month"}, "r", JoinKind.Inner), "r", {"requests"}),
    Typed = ''' + m_typed("Joined", ["requests", "backend_duration_avg_ms", "duration_avg_ms"]) + ''',
    Valid = Table.SelectRows(Typed, each [cod_app] <> null and [cod_app] <> "" and [requests] <> null and [requests] > 0 and [backend_duration_avg_ms] <> null),
    Weighted = Table.AddColumn(Table.AddColumn(Valid, "wb", each [requests] * [backend_duration_avg_ms], type number), "wg", each [requests] * (if [duration_avg_ms] = null then [backend_duration_avg_ms] else [duration_avg_ms]), type number),
'''


def query_latency_apps(folder):
    return latency_head(folder) + f'''    Layer = (t as table, layer as text, field as text) => List.Sum(Table.Column(Table.SelectRows(t, each [apim_layer] = layer), field)),
    Grouped = Table.Group(Weighted, {{"cod_app"}}, {{
        {{"ux_requests", each Layer(_, "UX", "requests"), type number}}, {{"ux_wb", each Layer(_, "UX", "wb"), type number}},
        {{"bs_requests", each Layer(_, "BS", "requests"), type number}}, {{"bs_wb", each Layer(_, "BS", "wb"), type number}},
        {{"bs_wg", each Layer(_, "BS", "wg"), type number}}}}),
    // Applications with traffic on both UX (at least {MIN_UX_REQUESTS} requests) and BS
    Both = Table.SelectRows(Grouped, each [ux_requests] <> null and [ux_requests] >= {MIN_UX_REQUESTS} and [bs_requests] <> null and [bs_requests] > 0),
    Ms = Table.AddColumn(Table.AddColumn(Table.AddColumn(Both, "ux_backend_ms", each [ux_wb] / [ux_requests], type number), "bs_backend_ms", each [bs_wb] / [bs_requests], type number), "bs_gateway_ms", each [bs_wg] / [bs_requests], type number),
    Columns = Table.SelectColumns(Ms, {{{m_list(LATENCY_APP_COLUMNS)}}}),
    Sorted = Table.Sort(Columns, {{{{"cod_app", Order.Ascending}}}})
in
    Sorted'''


def query_latency_layers(folder):
    return latency_head(folder) + f'''    Grouped = Table.Group(Weighted, {{"apim_layer"}}, {{{{"requests", each List.Sum([requests]), type number}}, {{"wb", each List.Sum([wb]), type number}}, {{"wg", each List.Sum([wg]), type number}}}}),
    Ms = Table.AddColumn(Table.AddColumn(Grouped, "backend_ms", each [wb] / [requests], type number), "gateway_ms", each [wg] / [requests], type number),
    Columns = Table.SelectColumns(Ms, {{{m_list(LATENCY_LAYER_COLUMNS)}}}),
    Sorted = Table.Sort(Columns, {{{{"apim_layer", Order.Ascending}}}})
in
    Sorted'''


def query_all_clusters(folder):
    return v1wb.query_head(folder) + f'''    Snapshot = ReadCsv("aa-k8s-snapshot", "cluster_snapshot.csv"),
    Cost = ReadCsv("aa-cost", "cost_by_resource.csv"),
    // AKS cost of the last month per cluster, every environment and category
    Aks = Table.SelectRows(Cost, each [owner_type] = "AKS"),
    LastMonth = List.Max(Aks[month]),
    Month = {m_typed('Table.SelectRows(Aks, each [month] = LastMonth)', ["cost_usd"])},
    PerCluster = Table.Group(Table.TransformColumns(Month, {{{{"owner_id", Text.Lower}}}}), {{"owner_id"}}, {{{{"cost_month_usd", each List.Sum([cost_usd]), type number}}}}),
    Clusters = Table.TransformColumns(Table.AddColumn(Snapshot, "key", each [cluster_id]), {{{{"key", Text.Lower}}}}),
    Joined = Table.ExpandTableColumn(Table.NestedJoin(Clusters, {{"key"}}, PerCluster, {{"owner_id"}}, "c", JoinKind.Inner), "c", {{"cost_month_usd"}}),
    Blank = Table.ReplaceValue(Joined, "", null, Replacer.ReplaceValue, {{"cod_app"}}),
    Columns = Table.SelectColumns(Blank, {{{m_list(ALL_CLUSTER_COLUMNS)}}}),
    Sorted = Table.Sort(Columns, {{{{"cost_month_usd", Order.Descending}}, {{"cluster", Order.Ascending}}}})
in
    Sorted'''


def query_scopes(folder):
    return v1wb.query_head(folder) + '''    Clusters = ReadCsv("aa-cluster-snapshot", "cluster_criteria_input.csv"),
    // The calculator offers every cluster, then every application
    Names = Table.Sort(Table.RenameColumns(Table.SelectColumns(Clusters, {"cluster"}), {{"cluster", "scope"}}), {{"scope", Order.Ascending}}),
    Codes = Table.Sort(Table.Distinct(Table.RenameColumns(Table.SelectColumns(Clusters, {"cod_app"}), {{"cod_app", "scope"}})), {{"scope", Order.Ascending}}),
    Scopes = Table.Combine({Names, Codes})
in
    Scopes'''


def query_apps(folder):
    return v1wb.query_head(folder) + '''    Clusters = ReadCsv("aa-cluster-snapshot", "cluster_criteria_input.csv"),
    Codes = Table.Sort(Table.Distinct(Table.SelectColumns(Clusters, {"cod_app"})), {{"cod_app", Order.Ascending}})
in
    Codes'''


# ---------------------------------------------------------------------------
# Sheets
# ---------------------------------------------------------------------------


def widths(sheet, headers, left=1):
    for index, header in enumerate(headers, start=left):
        sheet.column_dimensions[get_column_letter(index)].width = max(9, min(len(header) + 2, 20))


def write_query_table(sheet, name, columns, rows):
    v1wb.add_table(sheet, name, columns, rows)
    widths(sheet, columns)


def zero_factor(minimum, off_min, off_avg, basis='sizing_basis="PEAK"'):
    """Share of the hour billed: 1, or, when the hour's minimum is 0 (scale to zero allowed), its 5-minute buckets
    that cannot be at zero (the day with the fewest off buckets for the peak, their average for the average)."""
    return f'IF({minimum}=0,({BUCKETS}-IF({basis},{off_min},{off_avg}))/{BUCKETS},1)'


def capacity(today):
    """TPS per replica: concurrent_requests, or with capacity_from_today today's TPS per replica x capacity_safety,
    never below concurrent_requests."""
    return (f'IF(AND(capacity_from_today="YES",N({today})>0),MAX(concurrent_requests,{today}*capacity_safety),'
            'concurrent_requests)')


def busiest_hour_replicas(minimum, cell, per_replica):
    """Replicas of the hour at its busiest occurrence in the month: the average of its 12 buckets of 5 minutes, each with
    max(minimum, ceiling(TPS x traffic_share / concurrent_requests)), a bucket keeping the replicas of the one before
    (5 minutes before Container Apps scales in). cell(column) is a bucket's TPS."""
    columns = model.BUSIEST_COLUMNS
    terms = [f"MAX({minimum},CEILING(MAX({cell(now)},{cell(before)})*traffic_share/{per_replica},1))"
             for now, before in zip(columns, columns[:1] + columns[:-1])]
    return f"({'+'.join(terms)})/{BUCKETS}"


def write_traffic(sheet, count):
    """Traffic: the TrafficData query read row by row, and the replicas simulated each hour."""
    t = lambda column: r("Traffic", column)
    pulled = ["scope", "scope_type", "day_type", "hour", "office", "days", "tps_avg", "tps_peak", "off_5min_min",
              "off_5min_avg", "tps_per_replica_today"] + model.BUSIEST_COLUMNS
    minimum = f'IF({t("office")}=1,min_replicas_office,min_replicas_offhours)'
    formulas = [(column, pull("TrafficData", column)) for column in pulled] + [
        ("capacity", capacity(t("tps_per_replica_today"))),
        ("busiest_hour_replicas", busiest_hour_replicas(minimum, t, t("capacity"))),
        ("sim_replicas", f'IF(sizing_basis="BUSIEST_HOUR",{t("busiest_hour_replicas")},MAX({minimum},CEILING(IF('
                         f'sizing_basis="PEAK",{t("tps_peak")},{t("tps_avg")})*traffic_share/{t("capacity")},1))*'
                         f'{zero_factor(minimum, t("off_5min_min"), t("off_5min_avg"))})'),
        ("replica_seconds", f'{t("days")}*3600*{t("sim_replicas")}')]
    formula_table(sheet, "Traffic", formulas, count)
    widths(sheet, [header for header, _ in formulas])


def service_cost(match):
    """US$ per second of one replica of each "micro"-service where match (an array condition on Services) holds:
    Container Apps size = the largest of the CPU request, memory / gib_per_vcpu and min_vcpu, up to vcpu_step (MAX
    written as (a+b+|a-b|)/2 so SUMPRODUCT applies it row by row)."""
    # With size_from_usage the request comes down to the use x usage_headroom when lower (and the use is known):
    # request - (switch) x (use > 0) x max(0, request - use x headroom), all arithmetic for SUMPRODUCT.
    lower = lambda request, used: (f'({request}-(size_from_usage="YES")*({used}>0)*(({request}-{used}*usage_headroom)'
                                   f'+ABS({request}-{used}*usage_headroom))/2)')
    cpu = lower("Services[cpu_request_vcpu]", "Services[cpu_used_vcpu]")
    mem = lower("Services[mem_request_gib]", "Services[mem_used_gib]") + "/gib_per_vcpu"
    larger = f"(({cpu})+({mem})+ABS(({cpu})-({mem})))/2"
    size = f"CEILING(({larger}+min_vcpu+ABS({larger}-min_vcpu))/2,vcpu_step)"
    return f"SUMPRODUCT(({match})*{size}*(price_vcpu+gib_per_vcpu*price_gib))"


def criteria_formulas():
    c = lambda column: r("Criteria", column)
    return [
        # event-driven: its load is messages, not measured yet, so it counts as in use without measured traffic
        ("has_traffic", f'IF(AND(OR({c("event_driven")}<>"YES",rule_event_driven="NO"),{c("traffic_status")}="MEASURED",'
                        f'ISNUMBER({c("busiest_hour_tps")})),"YES","NO")'),
        ("size_ok", f'IF(N({c("microservices_over_aca_size")})=0,"YES","NO")'),
        ("tps_ok", f'IF({c("has_traffic")}="NO","-",IF({c("busiest_hour_tps")}<=max_tps,"YES","NO"))'),
        ("services_ok", f'IF(N({c("microservices")})<=max_services,"YES","NO")'),
        ("bia_ok", f'IF(IFERROR(MATCH({c("bia")},{BIA_ORDER},0),0)>=IFERROR(MATCH(aks_from_bia,{BIA_ORDER},0),99),'
                   '"NO","YES")'),
        ("critical", f'IF(OR({c("bia")}="MUY ALTA",{c("bia")}="ALTA"),"YES","NO")'),
        ("growing", f'IF(AND(ISNUMBER({c("microservices_app_projected")}),N({c("microservices_app_projected")})>'
                    f'max_services,N({c("microservices_app")})<=max_services),"YES","NO")'),
        # replicas of all its "micro"-services on Container Apps (average over the month) / today's on AKS
        ("aca_replicas", f'IFERROR({c("replica_seconds")}/(SUMIFS(Traffic[days],Traffic[scope],{c("cluster")},'
                         f'Traffic[scope_type],"CLUSTER")*3600)*N({c("microservices")}),"")'),
        ("replicas_vs_aks", f'IF(AND(ISNUMBER({c("aca_replicas")}),N({c("replicas_today_total")})>0),'
                            f'{c("aca_replicas")}/{c("replicas_today_total")},"")'),
        ("replicas_ok", f'IF(AND(rule_constant_load="YES",ISNUMBER({c("replicas_vs_aks")}),'
                        f'N({c("replicas_vs_aks")})>=max_replicas_vs_aks),"NO","YES")'),
        # requests in 24 hours: the month's requests over its days (DATE with day 0 is the month's last day)
        ("requests_day", f'IFERROR(N({c("requests_month")})/DAY(DATE(VALUE(LEFT({c("month")},4)),'
                         f'VALUE(MID({c("month")},6,2))+1,0)),0)'),
        ("avg_tps_day", f'{c("requests_day")}/86400'),
        # its quietest weekday hour over the weekday average (hourly averages): 1 = flat all day
        ("quiet_share", f'IFERROR(_xlfn.MINIFS(Traffic[tps_avg],Traffic[scope],{c("cluster")},Traffic[scope_type],"CLUSTER",'
                        f'Traffic[day_type],"weekday")/AVERAGEIFS(Traffic[tps_avg],Traffic[scope],{c("cluster")},'
                        'Traffic[scope_type],"CLUSTER",Traffic[day_type],"weekday"),"")'),
        ("load_ok", f'IF(AND(rule_high_load="YES",OR({c("avg_tps_day")}>max_avg_tps,AND(ISNUMBER({c("quiet_share")}),'
                    f'N({c("quiet_share")})>=min_quiet_share,{c("avg_tps_day")}>=min_constant_tps))),"NO","YES")'),
        ("shared_ok", f'IF(AND(rule_shared="YES",OR({c("traffic_layer")}="CORE",ISNUMBER(SEARCH(","&{c("cod_app")}&",",","&'
                      'SUBSTITUTE(shared_apps," ","")&",")))),"NO","YES")'),
        ("decision", f'IF({c("has_traffic")}="NO","NO_TRAFFIC",IF(OR({c("size_ok")}="NO",'
                     f'{c("tps_ok")}="NO",{c("services_ok")}="NO",{c("bia_ok")}="NO",{c("replicas_ok")}="NO",'
                     f'{c("shared_ok")}="NO",{c("load_ok")}="NO"),"AKS",IF(OR({c("critical")}="YES",'
                     f'{c("growing")}="YES"),"ACA_WITH_REVIEW","ACA")))'),
        ("replica_seconds", f'SUMIFS(Traffic[replica_seconds],Traffic[scope],{c("cluster")},Traffic[scope_type],'
                            '"CLUSTER")'),
        ("usd_per_second", service_cost(f'Services[cluster]={c("cluster")}')),
        ("aca_list_usd", f'{c("replica_seconds")}*{c("usd_per_second")}+IF({c("event_driven")}="YES",0,'
                         f'N({c("requests_month")}))*requests_buffer/1000000*price_requests'
                         '+environments_per_cluster*hours_month*(price_lb_hour+IF(environment_fee="YES",'
                         'price_env_hour,0))'),
        ("aca_month_pen", f'{c("aca_list_usd")}*(1-bank_discount)*(1+igv)*exchange_rate'),
        ("aks_month_pen", f'N({c("cost_month_usd")})*(1+igv)*exchange_rate'),
        ("saving_month_pen", f'{c("aks_month_pen")}-{c("aca_month_pen")}'),
        ("saving_share", f'IF({c("aks_month_pen")}>0,{c("saving_month_pen")}/{c("aks_month_pen")},0)'),
        ("group", f'IF({c("decision")}="NO_TRAFFIC",IF(AND({c("event_driven")}="YES",rule_event_driven="YES"),"IN_USE",'
                  f'IF({c("no_traffic_reason")}="IN_DEVELOPMENT","IN_DEVELOPMENT",'
                  f'IF({c("no_traffic_reason")}="IN_USE_NOT_MEASURED","IN_USE","NOT_IN_USE"))),IF(AND(OR('
                  f'{c("decision")}="ACA",{c("decision")}="ACA_WITH_REVIEW"),{c("saving_share")}<min_saving),'
                  f'"BELOW_MARGIN",{c("decision")}))'),
        ("goes_to_aca", f'IF(OR({c("group")}="ACA",{c("group")}="ACA_WITH_REVIEW"),"YES","NO")'),
        ("final_month_pen", f'IF({c("goes_to_aca")}="YES",{c("aca_month_pen")},{c("aks_month_pen")})'),
        # 1 = the most expensive on AKS; ties broken by cluster name
        ("cost_rank", f'RANK({c("aks_month_pen")},Criteria[aks_month_pen])+COUNTIFS(Criteria[aks_month_pen],'
                      f'{c("aks_month_pen")},Criteria[cluster],"<"&{c("cluster")})'),
    ]




def untitled(chart):
    """The chart without a title: Excel otherwise titles a one-series chart with the series name. (Wraps the
    instance's writer: a subclass would lose openpyxl's element list.)"""
    write = chart._write

    def write_untitled():
        tree = write()
        container = tree.find("chart")
        container.insert(0, container.makeelement("autoTitleDeleted", {"val": "1"}))
        return tree

    chart._write = write_untitled
    return chart


def value_labels():
    """Data labels with the value only (Excel adds the series and category names when they are left unset)."""
    labels = DataLabelList()
    labels.showVal, labels.showSerName, labels.showCatName, labels.showLegendKey = True, False, False, False
    return labels


def apart(chart):
    """Keep the title, axis titles and legend beside the plot. Excel draws them on it when the plot area is left to
    its automatic layout, so the plot gets a fixed inner box with room for each (fractions of the chart)."""
    left = 0.1 if chart.y_axis.title is not None else 0.06
    top = 0.13 if chart.title is not None else 0.05
    bottom = 0.27 if chart.x_axis.title is not None else 0.12
    right = 0.3 if chart.legend is not None and chart.legend.position == "r" else 0.04
    chart.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge", x=left, y=top,
                                                    w=1 - left - right, h=1 - top - bottom))
    for title in (chart.title, chart.x_axis.title, chart.y_axis.title):
        if title is not None:
            title.overlay = False
    if chart.legend is not None:
        chart.legend.overlay = False


def hour_axis(chart, step=HOUR_LABEL_STEP):
    """Show both axes (openpyxl 3.1 hides them by default), an hour label every HOUR_LABEL_STEP, whole numbers on
    the value axis and the legend below, none of them on the plot."""
    chart.x_axis.delete, chart.y_axis.delete = False, False
    chart.x_axis.tickLblSkip = chart.x_axis.tickMarkSkip = step
    chart.y_axis.number_format = "0"
    chart.legend = Legend()
    chart.legend.position = "b"
    apart(chart)


def write_results(sheet):
    sheet["A1"], sheet["A1"].font = "Result per destination (one row per cluster in Criteria)", Font(bold=True, size=14)
    headers = ["Destination", "Clusters", "AKS today (S/)", "After (S/)", "Saving (S/)"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=3, column=column, value=header)
    for row, (group, label) in enumerate(GROUPS, start=4):
        where = f'Criteria[group],"{group}"'
        values = [label, f"=COUNTIFS({where})", f"=SUMIFS(Criteria[aks_month_pen],{where})",
                  f"=SUMIFS(Criteria[final_month_pen],{where})", f"=C{row}-D{row}"]
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value).number_format = PEN if column >= 3 else "General"
    total = 4 + len(GROUPS)
    sheet.cell(row=total, column=1, value="Total")
    for column in (2, 3, 4, 5):
        letter = get_column_letter(column)
        sheet.cell(row=total, column=column, value=f"=SUM({letter}4:{letter}{total - 1})").number_format = \
            PEN if column >= 3 else "General"
    v1wb.plain_table(sheet, "Results", f"A3:E{total}")
    notes = ["Container Apps cost from traffic: replicas per hour = max(minimum, ceiling(TPS x traffic_share / "
             "concurrent_requests)) (Variables).",
             '="Saving a year (S/): "&TEXT(E' + str(total) + '*12,"#,##0")',
             "A cluster that meets the criteria but saves less than min_saving stays on AKS (BELOW_MARGIN)."]
    for offset, text in enumerate(notes, start=1):
        sheet.cell(row=total + offset, column=1, value=text)
    chart = untitled(BarChart())
    chart.type, chart.style, chart.varyColors = "col", 10, True
    chart.add_data(Reference(sheet, min_col=2, min_row=3, max_row=3 + len(GROUPS)), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=4, max_row=3 + len(GROUPS)))
    chart.legend = Legend()
    chart.legend.position = "r"
    apart(chart)
    chart.x_axis.delete = True  # the legend names each bar
    chart.y_axis.delete = False
    chart.dataLabels = value_labels()
    chart.height, chart.width = 9, 22
    sheet.add_chart(chart, f"A{total + 6}")
    for column, width in zip("ABCDE", (44, 10, 16, 16, 16)):
        sheet.column_dimensions[column].width = width




def write_top_costs(book, sheet, applications, everything):
    """Two tables side by side, the most expensive first: every PROD application cluster of the criteria with its
    Container Apps cost and saving, and every AKS cluster of any environment or category (AllClusters). The yellow
    top_n cell says how many rows show; the totals sit above, so the lists can grow down to every cluster."""
    sheet["A1"], sheet["A1"].font = "Top costs: where the AKS money is", Font(bold=True, size=14)
    sheet["A3"], sheet["A3"].font = "Show the top", BOLD
    sheet["C3"], sheet["C3"].fill = TOP, EDITABLE
    sheet["D3"] = f"clusters of each list (most expensive first; {max(applications, everything)} or more shows all)"
    v1wb.define(book, "top_n", "'Top costs'!$C$3")
    soles = "(1+igv)*exchange_rate"
    head = 12
    rank = f"ROW()-{head}"
    shown = lambda value: f'IF({rank}>top_n,"",IFERROR({value},""))'
    bill = f"(SUM(AllClusters[cost_month_usd])*{soles})"  # the whole AKS bill, every cluster

    # Left: PROD application clusters of the criteria, by AKS cost (Criteria[cost_rank]).
    sheet["A11"], sheet["A11"].font = "PROD application clusters by AKS cost (S/ a month, with IGV)", BOLD
    left = ["Rank", "Cluster", "App", "BIA", "AKS today (S/)", "Cost share", "Saving share", "Destination", "Stays on AKS because"]
    at = f"MATCH({rank},Criteria[cost_rank],0)"
    get = lambda column: f"INDEX(Criteria[{column}],{at})"
    flag = lambda column, value: f'{get(column)}="{value}"'
    # The saving share shows only for clusters that meet the criteria; the rest show "-" and why they stay.
    meets = f'OR({flag("group", "ACA")},{flag("group", "ACA_WITH_REVIEW")},{flag("group", "BELOW_MARGIN")})'
    why = (f'IF(AND({flag("event_driven", "YES")},rule_event_driven="YES"),"Event-driven: messages, not measured yet",'
           f'IF({flag("decision", "NO_TRAFFIC")},"No traffic of its own",IF({flag("size_ok", "NO")},"A ""micro""-service over 4 vCPU / 8 GiB",'
           f'IF({flag("tps_ok", "NO")},"Busiest hour over max_tps",IF({flag("services_ok", "NO")},'
           f'"Over max_services ""micro""-services",IF({flag("bia_ok", "NO")},"BIA "&aks_from_bia&" or above",'
           f'IF({flag("shared_ok", "NO")},"Shared: authentication or core",IF({flag("load_ok", "NO")},'
           f'"High or constant load",IF({flag("replicas_ok", "NO")},'
           f'"Constant load: nearly the replicas of AKS",'
           f'IF({flag("group", "BELOW_MARGIN")},"Saves less than min_saving",""))))))))))')
    left_formulas = [rank, get("cluster"), get("cod_app"), get("bia"), get("aks_month_pen"),
                     f'{get("aks_month_pen")}/SUM(Criteria[aks_month_pen])',                     f'IF({meets},{get("saving_share")},"-")', get("group"), why]
    formula_table(sheet, "TopApplications", [(header, shown(value)) for header, value in zip(left, left_formulas)],
                  applications, top=head)
    for column, fmt in ((5, PEN), (6, "0.0%"), (7, PCT)):
        for row in range(head + 1, head + 1 + applications):
            sheet.cell(row=row, column=column).number_format = fmt
    rows = f"$A${head + 1}:$I${head + applications}"
    sheet.conditional_formatting.add(rows, FormulaRule(formula=[f'OR($H{head + 1}="ACA",$H{head + 1}="ACA_WITH_REVIEW")'],
                                                       fill=GREEN))
    sheet.conditional_formatting.add(rows, FormulaRule(formula=[f'$H{head + 1}="BELOW_MARGIN"'], fill=AMBER))
    counted = 'SUMIFS(Criteria[saving_month_pen],Criteria[goes_to_aca],"YES"'
    totals = [('="Top "&MIN(top_n,COUNT(Criteria[cost_rank]))&" shown"', "SUM(TopApplications[AKS today (S/)])",
               f'{counted},Criteria[cost_rank],"<="&top_n)'),
              ("All PROD application clusters", "SUM(Criteria[aks_month_pen])", counted + ")")]
    sheet.cell(row=5, column=5, value="AKS today (S/)").font = BOLD
    sheet.cell(row=5, column=7, value="Saving counted (S/)").font = BOLD
    for row, (label, cost, saving) in enumerate(totals, start=6):
        sheet.cell(row=row, column=2, value=label).font = BOLD
        sheet.cell(row=row, column=5, value="=" + cost).number_format = PEN
        sheet.cell(row=row, column=7, value="=" + saving).number_format = PEN
    sheet.cell(row=8, column=2, value="Share of the shown in the PROD application cost")
    sheet["E8"], sheet["E8"].number_format = "=IF(E7>0,E6/E7,0)", PCT
    sheet.cell(row=9, column=2, value="Green: moves to Container Apps. Yellow: meets the criteria but saves less than "
                                      "min_saving. \"-\": a criterion keeps it on AKS.")

    # Right: every AKS cluster, any environment or category (AllClusters comes sorted by cost).
    start = 11  # column K
    letter = lambda offset: get_column_letter(start + offset)
    sheet.cell(row=11, column=start, value="AKS clusters of every environment and category "
                                           "(S/ a month, with IGV)").font = BOLD
    right = ["Rank", "Cluster", "Environment", "Category", "App", "AKS cost (S/)", "Cost share", "In the criteria", "Saving share",
             "Destination"]
    name = f"INDEX(AllClusters[cluster],{rank})"
    at = f"MATCH({name},Criteria[cluster],0)"
    group = f"INDEX(Criteria[group],{at})"
    inside = f"COUNTIF(Criteria[cluster],{name})>0"
    right_formulas = [rank, name, f"INDEX(AllClusters[environment],{rank})", f"INDEX(AllClusters[category],{rank})",
                      f'INDEX(AllClusters[cod_app],{rank})&""', f"INDEX(AllClusters[cost_month_usd],{rank})*{soles}",
                      f"INDEX(AllClusters[cost_month_usd],{rank})*{soles}/{bill}",
                      f'IF({inside},"YES","NO")',
                      f'IF({inside},IF(OR({group}="ACA",{group}="ACA_WITH_REVIEW",{group}="BELOW_MARGIN"),'
                      f'INDEX(Criteria[saving_share],{at}),"-"),"-")',
                      f'IF({inside},{group},"Not evaluated")']
    formula_table(sheet, "TopAllClusters", [(header, shown(value)) for header, value in zip(right, right_formulas)],
                  everything, top=head, left=start)
    for offset, fmt in ((5, PEN), (6, "0.0%"), (8, PCT)):
        for row in range(head + 1, head + 1 + everything):
            sheet.cell(row=row, column=start + offset).number_format = fmt
    rows = f"${letter(0)}${head + 1}:${letter(9)}${head + everything}"
    sheet.conditional_formatting.add(rows, FormulaRule(
        formula=[f'OR(${letter(9)}{head + 1}="ACA",${letter(9)}{head + 1}="ACA_WITH_REVIEW")'], fill=GREEN))
    sheet.conditional_formatting.add(rows, FormulaRule(formula=[f'${letter(9)}{head + 1}="BELOW_MARGIN"'], fill=AMBER))
    all_cost = f"SUM(AllClusters[cost_month_usd])*{soles}"
    app_cost = (f'SUMIFS(AllClusters[cost_month_usd],AllClusters[environment],"PROD",AllClusters[category],'
                f'"APPLICATION")*{soles}')
    saving = 'SUMIFS(Criteria[saving_month_pen],Criteria[goes_to_aca],"YES")'
    summary = [("All AKS clusters (S/ a month)", all_cost, PEN),
               ("PROD application clusters (the criteria)", app_cost, PEN),
               ("Their share of the AKS bill", f"IF({all_cost}>0,{app_cost}/({all_cost}),0)", PCT),
               ("Saving counted (S/ a month)", saving, PEN),
               ("Saving counted, share of the AKS bill", f"IF({all_cost}>0,{saving}/({all_cost}),0)", PCT)]
    for row, (label, formula, fmt) in enumerate(summary, start=5):
        sheet.cell(row=row, column=start + 1, value=label).font = BOLD
        sheet.cell(row=row, column=start + 5, value="=" + formula).number_format = fmt
    sheet.cell(row=10, column=start + 1,
               value="Not evaluated: other environments and infrastructure clusters are outside the criteria.")
    for column, width in zip("ABCDEFGHI", (6, 18, 7, 10, 14, 11, 18, 16, 34)):
        sheet.column_dimensions[column].width = width
    sheet.column_dimensions["J"].width = 3
    for offset, width in enumerate((6, 18, 12, 16, 7, 14, 11, 12, 11, 16)):
        sheet.column_dimensions[letter(offset)].width = width
    sheet.freeze_panes = f"A{head + 1}"


def write_calculator(sheet, default_scope):
    sheet["A1"], sheet["A1"].font = "Calculator: a cluster or an application, hour by hour", Font(bold=True, size=14)
    is_cluster = 'COUNTIF(Criteria[cluster],$B$2)>0'
    pick = lambda column, app_sum=True: (f'IF({is_cluster},INDEX(Criteria[{column}],MATCH($B$2,Criteria[cluster],0)),'
                                         + (f'SUMIFS(Criteria[{column}],Criteria[cod_app],$B$2))' if app_sum else '"-")'))
    lines = [("Cluster or application", default_scope),
             ("Type", f'=IF({is_cluster},"CLUSTER","APP")'),
             ("Decision (cluster)", f'=IFERROR(INDEX(Criteria[decision],MATCH($B$2,Criteria[cluster],0)),'
                                    '"see sheet Apps")'),
             ("\"Micro\"-services", f'=IF({is_cluster},COUNTIFS(Services[cluster],$B$2),'
                                    'COUNTIFS(Services[cod_app],$B$2))'),
             ("Clusters", f'=IF({is_cluster},1,COUNTIFS(Criteria[cod_app],$B$2))'),
             ("US$ per second, one replica of each", f"=IF({is_cluster},{service_cost('Services[cluster]=$B$2')},"
                                                    f"{service_cost('Services[cod_app]=$B$2')})"),
             ("Requests of the month", f"={pick('requests_month')}"),
             ("AKS today (S/, with IGV)", f"={pick('aks_month_pen')}"),
             ("Weekdays / weekend days", "=weekdays&\" / \"&weekend_days")]
    for row, (label, value) in enumerate(lines, start=2):
        sheet.cell(row=row, column=1, value=label).font = BOLD
        sheet.cell(row=row, column=2, value=value)
    sheet["B2"].fill = EDITABLE
    sheet["C2"] = "(choose it from the list)"
    validation = DataValidation(type="list", formula1='=INDIRECT("Scopes[scope]")', allow_blank=False)
    sheet.add_data_validation(validation)
    validation.add("B2")
    sheet["B9"].number_format = PEN

    top = 13
    block_headers = ["TPS min", "TPS average", "TPS peak", "Replicas today", "Replicas by peak",
                     "Replicas by busiest hour", "Replicas by average", "Chosen"]
    headers = ["Hour"] + ["Weekday " + block_headers[0]] + block_headers[1:] + \
        ["Weekend " + block_headers[0]] + block_headers[1:]
    sheet.cell(row=top - 1, column=1, value="Replicas per \"micro\"-service each hour (by peak, by busiest hour and by average; "
                                            "Chosen: yellow, change it; it starts as sizing_basis)").font = BOLD
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=top, column=column, value=header).font = BOLD
    fields = [("tps_min", "0.00"), ("tps_avg", "0.00"), ("tps_peak", "0.00"), ("replicas_today_avg", "0.0")]
    for hour in range(24):
        row = top + 1 + hour
        sheet.cell(row=row, column=1, value=hour)
        for block, day_type in ((0, "weekday"), (BLOCK, "weekend")):
            where = f'TrafficData[scope],$B$2,TrafficData[day_type],"{day_type}",TrafficData[hour],$A{row}'
            for offset, (field, fmt) in enumerate(fields):
                cell = sheet.cell(row=row, column=2 + block + offset, value=f"=SUMIFS(TrafficData[{field}],{where})")
                cell.number_format = fmt
            minimum = f"IF(SUMIFS(TrafficData[office],{where})>0,min_replicas_office,min_replicas_offhours)"
            average, peak = (get_column_letter(2 + block + offset) for offset in (1, 2))
            off = {name: f"SUMIFS(TrafficData[{name}],{where})" for name in ("off_5min_min", "off_5min_avg")}
            for offset, tps, basis in ((4, peak, "TRUE"), (6, average, "FALSE")):
                factor = zero_factor(minimum, off["off_5min_min"], off["off_5min_avg"], basis)
                sheet.cell(row=row, column=2 + block + offset,
                           value=f"=MAX({minimum},CEILING({tps}{row}*traffic_share/"
                                 f"{capacity(f'SUMIFS(TrafficData[tps_per_replica_today],{where})')},1))*{factor}"
                           ).number_format = "0.##"
            traffic = f'Traffic[scope],$B$2,Traffic[day_type],"{day_type}",Traffic[hour],$A{row}'
            sheet.cell(row=row, column=2 + block + 5,
                       value=f"=SUMIFS(Traffic[busiest_hour_replicas],{traffic})").number_format = "0.##"
            by_peak, by_busiest, by_average = (get_column_letter(2 + block + offset) for offset in (4, 5, 6))
            chosen = sheet.cell(row=row, column=9 + block,
                                value=f'=IF(sizing_basis="PEAK",{by_peak}{row},IF(sizing_basis="BUSIEST_HOUR",'
                                      f'{by_busiest}{row},{by_average}{row}))')
            chosen.fill, chosen.number_format = EDITABLE, "0.##"
    last = top + 24
    cost_top = last + 2
    sheet.cell(row=cost_top, column=1, value="Monthly cost (US$ list price unless said)").font = BOLD
    for column, header in ((2, "By peak"), (3, "By busiest hour"), (4, "By average"), (5, "Chosen")):
        sheet.cell(row=cost_top, column=column, value=header).font = BOLD
    seconds = lambda wd, we: f"(SUM({wd}{top + 1}:{wd}{last})*weekdays+SUM({we}{top + 1}:{we}{last})*weekend_days)*3600"
    grid = {2: ("F", "N"), 3: ("G", "O"), 4: ("H", "P"), 5: ("I", "Q")}  # cost column -> its replica columns
    each = lambda formula: {column: formula(get_column_letter(column), column) for column in grid}
    meters = [("Replica-seconds per \"micro\"-service", each(lambda x, column: seconds(*grid[column])), "#,##0"),
              ("Compute (vCPU + memory)", each(lambda x, _: f"{x}{cost_top + 1}*$B$7"), USD),
              ("Requests", each(lambda x, _: "N($B$8)*requests_buffer/1000000*price_requests"), USD),
              ("Environment (load balancer)",
               each(lambda x, _: "$B$6*environments_per_cluster*hours_month*(price_lb_hour+IF(environment_fee="
                                 '"YES",price_env_hour,0))'), USD),
              ("Total list price", each(lambda x, _: f"SUM({x}{cost_top + 2}:{x}{cost_top + 4})"), USD),
              ("Container Apps with discount and IGV (S/)",
               each(lambda x, _: f"{x}{cost_top + 5}*(1-bank_discount)*(1+igv)*exchange_rate"), PEN),
              ("AKS today (S/)", each(lambda x, _: "$B$9"), PEN),
              ("Saving (S/)", each(lambda x, _: f"{x}{cost_top + 7}-{x}{cost_top + 6}"), PEN),
              ("Saving share", each(lambda x, _: f"IF({x}{cost_top + 7}>0,{x}{cost_top + 8}/{x}{cost_top + 7},0)"),
               PCT)]
    for offset, (label, formulas, fmt) in enumerate(meters, start=1):
        sheet.cell(row=cost_top + offset, column=1, value=label)
        for column, formula in formulas.items():
            sheet.cell(row=cost_top + offset, column=column, value="=" + formula).number_format = fmt
    sheet.cell(row=cost_top + len(meters) + 2, column=1,
               value=f"By {model.DEFAULTS['sizing_basis'].lower().replace('_', ' ')} is the figure presented. By peak "
                     "is the most conservative; by busiest hour replays each hour's "
                     "busiest occurrence in the month 5 minutes at a time; by average is the optimistic one.")
    for offset, text in enumerate(SCALE_TO_ZERO_NOTE, start=3):
        sheet.cell(row=cost_top + len(meters) + offset, column=1, value=text)
    for block, title, anchor in ((0, "Weekday TPS by hour", "S2"), (BLOCK, "Weekend TPS by hour", "S20")):
        chart = LineChart()
        chart.title, chart.height, chart.width = title, 8, 16
        chart.y_axis.title, chart.x_axis.title = "TPS", "Hour"
        chart.add_data(Reference(sheet, min_col=2 + block, max_col=4 + block, min_row=top, max_row=last),
                       titles_from_data=True)
        for series, name in zip(chart.series, ("TPS min", "TPS average", "TPS peak")):
            series.tx = SeriesLabel(v=name)
        chart.set_categories(Reference(sheet, min_col=1, min_row=top + 1, max_row=last))
        hour_axis(chart)
        sheet.add_chart(chart, anchor)
    replicas_chart = BarChart()
    replicas_chart.type, replicas_chart.title, replicas_chart.height, replicas_chart.width = \
        "col", "Chosen replicas per \"micro\"-service", 8, 16
    replicas_chart.y_axis.title, replicas_chart.x_axis.title = "Replicas", "Hour"
    for column, name in ((9, "Weekday"), (9 + BLOCK, "Weekend")):
        replicas_chart.add_data(Reference(sheet, min_col=column, min_row=top + 1, max_row=last))
        replicas_chart.series[-1].tx = SeriesLabel(v=name)
    replicas_chart.set_categories(Reference(sheet, min_col=1, min_row=top + 1, max_row=last))
    hour_axis(replicas_chart)
    sheet.add_chart(replicas_chart, "S38")
    write_usage(sheet, top)
    write_five_minutes(sheet, top)
    sheet.column_dimensions["A"].width = 40
    for column in "BCDEFGHIJKLMNOPQ":
        sheet.column_dimensions[column].width = 13




def write_usage(sheet, top, left=USAGE_COLUMN):
    """CPU and memory the nodes use each hour, as a share of what they offer (weekday and weekend), and charts next to
    the TPS ones: CPU at its own scale (it moves in a few percent) and memory split into applications, system pool and
    unused (it does not follow the traffic)."""
    sheet.cell(row=top - 1, column=left, value="CPU and memory used each hour (share of what the nodes offer)").font = BOLD
    parts = ["CPU", "memory, applications", "memory, system", "memory, unused"]
    headers = ["Hour"] + [f"Weekday {part}" for part in parts] + [f"Weekend {part}" for part in parts]
    for offset, header in enumerate(headers):
        sheet.cell(row=top, column=left + offset, value=header).font = BOLD
    is_cluster = 'COUNTIF(Criteria[cluster],$B$2)>0'
    for hour in range(24):
        row = top + 1 + hour
        sheet.cell(row=row, column=left, value=hour)
        for block, day_type in ((0, "weekday"), (4, "weekend")):
            # SUMIFS takes no IF for its range: one per scope (a cluster, or every cluster of an application)
            total = lambda column, day=day_type: (
                f'IF({is_cluster},SUMIFS(UsageData[{column}],UsageData[cluster],$B$2,UsageData[day_type],"{day}",'
                f'UsageData[hour],{hour}),SUMIFS(UsageData[{column}],UsageData[cod_app],$B$2,UsageData[day_type],'
                f'"{day}",UsageData[hour],{hour}))')
            share = lambda used, offered: f'IF({offered}>0,{used}/{offered},"")'
            memory, system, offered = total("mem_used_avg_gib"), total("mem_system_avg_gib"), total("mem_allocatable_gib")
            values = [share(total("cpu_used_avg"), total("cpu_allocatable")),
                      share(f"MAX(0,{memory}-{system})", offered), share(system, offered),
                      share(f"MAX(0,{offered}-{memory})", offered)]
            for offset, value in enumerate(values, start=1):
                sheet.cell(row=row, column=left + block + offset, value="=" + value).number_format = PCT
    last = top + 24
    hours = Reference(sheet, min_col=left, min_row=top + 1, max_row=last)
    # Traffic and CPU: the hour's average TPS as columns (left axis) and the CPU used as a line (right axis, its own
    # scale), weekday and weekend, next to the TPS charts.
    for block, tps_block, day, anchor in ((0, 0, "Weekday", "AC2"), (4, BLOCK, "Weekend", "AC20")):
        traffic = BarChart()
        traffic.type, traffic.varyColors, traffic.gapWidth = "col", False, 30
        traffic.title, traffic.height, traffic.width = f"{day}: traffic and CPU", 8, 16
        traffic.y_axis.title, traffic.x_axis.title = "TPS (average of the hour)", "Hour"
        traffic.add_data(Reference(sheet, min_col=3 + tps_block, min_row=top, max_row=last), titles_from_data=True)
        traffic.series[0].tx = SeriesLabel(v="TPS average")
        traffic.series[0].graphicalProperties.solidFill = "B4C7E7"
        traffic.series[0].graphicalProperties.line.solidFill = "B4C7E7"
        traffic.set_categories(hours)
        hour_axis(traffic)
        cpu = LineChart()
        cpu.varyColors = False
        cpu.add_data(Reference(sheet, min_col=left + 1 + block, min_row=top, max_row=last), titles_from_data=True)
        cpu.series[0].tx = SeriesLabel(v="CPU used")
        cpu.series[0].graphicalProperties.line.solidFill = "C0504D"
        cpu.series[0].graphicalProperties.line.width = 28575
        cpu.y_axis.axId, cpu.y_axis.title, cpu.y_axis.crosses = 200, "CPU", "max"
        cpu.y_axis.majorGridlines, cpu.y_axis.delete = None, False
        cpu.y_axis.numFmt = NumFmt(formatCode="0%", sourceLinked=False)
        cpu.y_axis.scaling.min = 0
        traffic += cpu
        sheet.add_chart(traffic, anchor)
    # Capacity of the nodes in office hours (weekday): what the "micro"-services request and what is used, as a share
    # of what the nodes offer. Requests far above use: right-size the requests; requests far below 100%: too many nodes.
    box = last + 3
    sheet.cell(row=box - 1, column=left, value="Capacity in office hours, weekday (share of what the nodes offer)").font = BOLD
    for offset, header in enumerate(["", "Requested by the \"micro\"-services", "Used, office average",
                                      "Used, busiest hour"]):
        sheet.cell(row=box, column=left + offset, value=header).font = BOLD
    office = (top + 1 + 8, top + 1 + 19)  # 8:00 to 19:59
    span = lambda column: f"{get_column_letter(column)}{office[0]}:{get_column_letter(column)}{office[1]}"
    whole = lambda column: f"{get_column_letter(column)}{top + 1}:{get_column_letter(column)}{last}"
    offered = lambda column: (
        f'(IF({is_cluster},SUMIFS(UsageData[{column}],UsageData[cluster],$B$2,UsageData[day_type],"weekday",'
        f'UsageData[hour],">=8",UsageData[hour],"<20"),SUMIFS(UsageData[{column}],UsageData[cod_app],$B$2,'
        f'UsageData[day_type],"weekday",UsageData[hour],">=8",UsageData[hour],"<20"))/12)')
    requested = lambda column: (
        f"IF({is_cluster},SUMPRODUCT((Services[cluster]=$B$2)*Services[{column}]*Services[replicas_today_office]),"
        f"SUMPRODUCT((Services[cod_app]=$B$2)*Services[{column}]*Services[replicas_today_office]))")
    rows = [("CPU", requested("cpu_request_vcpu"), offered("cpu_allocatable"), f"AVERAGE({span(left + 1)})",
             f"MAX({whole(left + 1)})"),
            ("Memory", requested("mem_request_gib"), offered("mem_allocatable_gib"),
             f"1-AVERAGE({span(left + 4)})", f"1-MIN({whole(left + 4)})")]
    for offset, (label, asked, nodes, average, busiest) in enumerate(rows, start=1):
        sheet.cell(row=box + offset, column=left, value=label)
        for column, formula in ((1, f'IF({nodes}>0,{asked}/{nodes},"")'), (2, average), (3, busiest)):
            sheet.cell(row=box + offset, column=left + column, value=f"=IFERROR({formula},\"\")").number_format = PCT
    capacity_chart = BarChart()
    capacity_chart.type, capacity_chart.grouping, capacity_chart.varyColors = "bar", "clustered", False
    capacity_chart.title, capacity_chart.height, capacity_chart.width = "Capacity: requested and used", 8, 16
    capacity_chart.y_axis.title = "Share of what the nodes offer"
    capacity_chart.add_data(Reference(sheet, min_col=left + 1, max_col=left + 3, min_row=box, max_row=box + 2),
                            titles_from_data=True)
    for series, colour in zip(capacity_chart.series, ("A5A5A5", "4472C4", "C0504D")):
        series.graphicalProperties.solidFill = colour
        series.graphicalProperties.line.solidFill = colour
    capacity_chart.set_categories(Reference(sheet, min_col=left, min_row=box + 1, max_row=box + 2))
    capacity_chart.x_axis.delete = capacity_chart.y_axis.delete = False
    capacity_chart.y_axis.numFmt = NumFmt(formatCode="0%", sourceLinked=False)
    capacity_chart.y_axis.scaling.min, capacity_chart.y_axis.scaling.max, capacity_chart.y_axis.majorUnit = 0, 1, 0.25
    capacity_chart.legend.position = "b"
    sheet.add_chart(capacity_chart, "AM2")
    # Memory by hour, weekday only: it does not follow the traffic (weekend looks the same).
    memory = BarChart()
    memory.type, memory.grouping, memory.overlap, memory.gapWidth = "col", "percentStacked", 100, 40
    memory.title, memory.height, memory.width = "Weekday memory: apps, system, unused", 8, 16
    memory.y_axis.title, memory.x_axis.title = "Share of the nodes' memory", "Hour"
    memory.add_data(Reference(sheet, min_col=left + 2, max_col=left + 4, min_row=top, max_row=last),
                    titles_from_data=True)
    for series, name, colour in zip(memory.series, ("Applications", "System pool", "Unused"),
                                    ("4472C4", "C0504D", "D9D9D9")):
        series.tx = SeriesLabel(v=name)
        series.graphicalProperties.solidFill = colour
        series.graphicalProperties.line.solidFill = colour
    memory.set_categories(hours)
    hour_axis(memory)
    memory.y_axis.numFmt = NumFmt(formatCode="0%", sourceLinked=False)
    sheet.add_chart(memory, "AM20")
    for offset in range(len(headers)):
        sheet.column_dimensions[get_column_letter(left + offset)].width = 12


def write_five_minutes(sheet, top, left=FIVE_MINUTES_COLUMN):
    """Every 5 minutes of the day: the TPS of each hour's busiest occurrence in the month next to the hour's peak and average (the
    hour grid's), for weekday and weekend, and a chart of each."""
    sheet.cell(row=top - 1, column=left, value="Every 5 minutes (busiest hour of the month for each hour)").font = BOLD
    headers = ["Time", "Weekday busiest hour", "Weekday peak", "Weekday average", "Weekend busiest hour",
               "Weekend peak", "Weekend average"]
    for offset, header in enumerate(headers):
        sheet.cell(row=top, column=left + offset, value=header).font = BOLD
    for hour in range(24):
        hour_row = top + 1 + hour
        for bucket, name in enumerate(model.BUSIEST_COLUMNS):
            row = top + 1 + hour * BUCKETS + bucket
            sheet.cell(row=row, column=left, value=f"{hour:02d}:{bucket * 5:02d}")
            for offset, (block, day_type) in enumerate(((0, "weekday"), (BLOCK, "weekend"))):
                where = f'TrafficData[scope],$B$2,TrafficData[day_type],"{day_type}",TrafficData[hour],{hour}'
                values = [f"=SUMIFS(TrafficData[{name}],{where})",
                          f"=${get_column_letter(4 + block)}${hour_row}", f"=${get_column_letter(3 + block)}${hour_row}"]
                for index, value in enumerate(values):
                    sheet.cell(row=row, column=left + 1 + offset * 3 + index, value=value).number_format = "0.00"
    last = top + 24 * BUCKETS
    for offset, (title, anchor) in enumerate((("Weekday every 5 minutes: busiest hour, peak and average", "S56"),
                                              ("Weekend every 5 minutes: busiest hour, peak and average", "S74"))):
        chart = LineChart()
        chart.title, chart.height, chart.width = title, 8, 24
        chart.y_axis.title, chart.x_axis.title = "TPS", "Hour"
        first = left + 1 + offset * 3
        chart.add_data(Reference(sheet, min_col=first, max_col=first + 2, min_row=top + 1, max_row=last))
        for series, name in zip(chart.series, ("Busiest hour", "Peak of the hour", "Average of the hour")):
            series.tx = SeriesLabel(v=name)
            series.smooth = False
        chart.set_categories(Reference(sheet, min_col=left, min_row=top + 1, max_row=last))
        hour_axis(chart, step=HOUR_LABEL_STEP * BUCKETS)
        sheet.add_chart(chart, anchor)
    for offset in range(len(headers)):
        sheet.column_dimensions[get_column_letter(left + offset)].width = 12


def write_criteria(sheet, count):
    """Criteria: the ClusterData query read row by row, then the criteria, the decision and the costs."""
    formulas = [(column, pull("ClusterData", column)) for column in CLUSTER_COLUMNS] + criteria_formulas()
    formula_table(sheet, "Criteria", formulas, count)
    headers = [header for header, _ in formulas]
    widths(sheet, headers)
    for index, header in enumerate(headers, start=1):
        fmt = USD if header.endswith("_usd") else PEN if header.endswith("_pen") else PCT if header in ("saving_share", "replicas_vs_aks") \
            else "0.0" if header == "aca_replicas" else "#,##0" if header == "requests_day" else "0.0" if header == "avg_tps_day" \
            else PCT if header == "quiet_share" else None
        if fmt:
            for row in range(2, count + 2):
                sheet.cell(row=row, column=index).number_format = fmt
    sheet.freeze_panes = "C2"


def write_apps(sheet, codes):
    """Apps: the AppList query (column A) and, next to it, one formula row per application."""
    v1wb.add_table(sheet, "AppList", ["cod_app"], [[code] for code in codes])
    a = lambda column: r("Apps", column)
    formulas = [("application", pull("AppList", "cod_app")),
                ("clusters", f'COUNTIFS(Criteria[cod_app],{a("application")})'),
                ("clusters_to_aca", f'COUNTIFS(Criteria[cod_app],{a("application")},Criteria[goes_to_aca],"YES")'),
                ("status", f'IF({a("clusters_to_aca")}=0,"Stays on AKS",IF({a("clusters_to_aca")}={a("clusters")},'
                           '"All to ACA","Split"))'),
                ("aks_month_pen", f'SUMIFS(Criteria[aks_month_pen],Criteria[cod_app],{a("application")})'),
                ("after_month_pen", f'SUMIFS(Criteria[final_month_pen],Criteria[cod_app],{a("application")})'),
                ("saving_month_pen", f'{a("aks_month_pen")}-{a("after_month_pen")}')]
    formula_table(sheet, "Apps", formulas, len(codes), left=3)
    sheet.column_dimensions["A"].width = 12
    for index, (header, _) in enumerate(formulas, start=3):
        sheet.column_dimensions[get_column_letter(index)].width = 16
        if header.endswith("_pen"):
            for row in range(2, len(codes) + 2):
                sheet.cell(row=row, column=index).number_format = PEN


def box(sheet, cells, text, fill, color="FFFFFF"):
    sheet.merge_cells(cells)
    cell = sheet[cells.split(":")[0]]
    cell.value, cell.fill = text, PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_latency(book, sheet, app_count, default_app):
    """The chain UX -> BS -> Core -> mainframe, built from the bottom: latency(layer) = own time + calls to the next
    layer x (hop + latency(next)), plus, on Container Apps, its ingress, a cold start when no replica is on and the
    scale-out delay of a burst. Own times start from the measured APIM latency (LatencyData, LatencyLayers)."""
    sheet["A1"], sheet["A1"].font = "Latency of one request along the chain (ms) - parameters to validate", \
        Font(bold=True, size=14)
    sheet["A3"], sheet["A3"].font = "Application (UX and BS measured)", BOLD
    sheet["B3"], sheet["B3"].fill = default_app, EDITABLE
    validation = DataValidation(type="list", formula1='=INDIRECT("LatencyApps[application]")', allow_blank=False)
    sheet.add_data_validation(validation)
    validation.add("B3")
    v1wb.define(book, "latency_app", "Latency!$B$3")
    params = [("hop_aks_ms", 2, "Network hop between services inside AKS"),
              ("aca_ingress_ms", 5, "Extra of the Container Apps ingress (Envoy) on each call into a layer on ACA"),
              ("cross_platform_ms", 5, "Extra when a call goes from AKS to Container Apps or back"),
              ("cold_start_s", 30, "First request when the layer has no replica on (scale to zero)"),
              ("cold_probability", 0, "Share of requests that find no replica on (0 with 1 minimum replica)"),
              ("scale_delay_s", 15, "Container Apps checks its HTTP rule every 15 s"),
              ("burst_probability", 0.001, "Share of requests that arrive while a burst waits for a new replica")]
    for offset, (label, value, comment) in enumerate(params, start=5):
        sheet.cell(row=offset, column=1, value=label)
        cell = sheet.cell(row=offset, column=2, value=value)
        cell.fill = EDITABLE
        sheet.cell(row=offset, column=3, value=comment)
        v1wb.define(book, label, f"Latency!$B${offset}")
    app = lambda column: f"INDEX(LatencyApps[{column}],MATCH(latency_app,LatencyApps[application],0))"
    layer = lambda name, column: f'IFERROR(INDEX(LatencyLayers[{column}],MATCH("{name}",LatencyLayers[apim_layer],0)),0)'
    top = 14
    headers = ["Layer", "Own time (ms)", "Calls to next layer", "Measured (ms)", "All AKS", "UX on ACA",
               "UX+BS on ACA", "UX+BS+Core on ACA"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=top, column=column, value=header).font = BOLD
    platforms = {"All AKS": [], "UX on ACA": ["UX"], "UX+BS on ACA": ["UX", "BS"], "UX+BS+Core on ACA": ["UX", "BS", "CORE"]}
    # Own time: the selected application's UX minus what its BS answers through the gateway, and its BS minus what
    # Core answers (Core is measured for the whole layer: half Core, half mainframe, until measured per application).
    # From the bottom up the chain then adds to the application's measured UX latency.
    own = {"UX": f'=MAX(1,{app("ux_minus_bs")})',
           "BS": f'=MAX(1,{app("bs_gateway_ms")}-{layer("CORE", "gateway_ms")})',
           "CORE": f'=MAX(1,{layer("CORE", "backend_ms")}/2)', "MAINFRAME": f'={layer("CORE", "backend_ms")}/2'}
    measured = {"UX": f'={app("ux_backend_ms")}', "BS": f'={app("bs_gateway_ms")}',
                "CORE": f'={layer("CORE", "backend_ms")}', "MAINFRAME": '="-"'}
    calls = {"UX": 1, "BS": 1, "CORE": 1, "MAINFRAME": 0}
    for offset, name in enumerate(LAYERS, start=1):
        row = top + offset
        sheet.cell(row=row, column=1, value=name)
        for column, value in ((2, own[name]), (3, calls[name])):
            cell = sheet.cell(row=row, column=column, value=value)
            cell.fill, cell.number_format = EDITABLE, "0"
        sheet.cell(row=row, column=4, value=measured[name]).number_format = "0"
        for column, scenario in enumerate(platforms, start=5):
            if name == "MAINFRAME":
                sheet.cell(row=row, column=column, value="ON-PREM")
            else:
                sheet.cell(row=row, column=column, value="ACA" if name in platforms[scenario] else "AKS").fill = EDITABLE
    result = top + len(LAYERS) + 2
    sheet.cell(row=result, column=1, value="Latency from this layer down (ms)").font = BOLD
    for offset, name in enumerate(reversed(LAYERS)):
        row = result + 1 + offset
        sheet.cell(row=row, column=1, value=name)
        source = top + 1 + LAYERS.index(name)
        for column in range(5, 9):
            letter = get_column_letter(column)
            here = f"{letter}{source}"
            extra = (f'IF({here}="ACA",aca_ingress_ms+cold_probability*cold_start_s*1000+burst_probability*'
                     'scale_delay_s*1000,0)')
            if name == "MAINFRAME":
                formula = f"=$B{source}+{extra}"
            else:
                below = f"{letter}{source + 1}"
                hop = f'(hop_aks_ms+IF(AND({below}<>"ON-PREM",{here}<>{below}),cross_platform_ms,0))'
                formula = f"=$B{source}+{extra}+$C{source}*({hop}+{letter}{row - 1})"
            sheet.cell(row=row, column=column, value=formula).number_format = "0"
    total = result + len(LAYERS)
    sheet.cell(row=total + 1, column=1, value="Extra over all AKS (ms)").font = BOLD
    for column in range(5, 9):
        letter = get_column_letter(column)
        sheet.cell(row=total + 1, column=column, value=f"={letter}{total}-$E${total}").number_format = "0"
    sheet.cell(row=total + 2, column=1, value="BS requests per UX request (selected application)").font = BOLD
    sheet.cell(row=total + 2, column=2, value=f'={app("bs_calls_per_ux")}').number_format = "0.0"

    # The chain drawn with cells: one UX request can call BS several times.
    box(sheet, "K4:L6", "UX\n1 request", "1E3A78")
    box(sheet, "M5:M5", "→", "FFFFFF", "1E3A78")
    for index, cells in enumerate(("N3:O3", "N5:O5", "N7:O7"), start=1):
        box(sheet, cells, f"BS call {index}", "4A8FB8")
    box(sheet, "P5:P5", "→", "FFFFFF", "1E3A78")
    box(sheet, "Q4:R6", "Core", "F2A15A")
    box(sheet, "S5:S5", "→", "FFFFFF", "1E3A78")
    box(sheet, "T4:U6", "Mainframe", "5B6770")
    sheet["K9"] = ("A UX service is a wrapper: for one request it may call BS several times (\"Calls to next layer\"). "
                   "One after another the times add up (the model's case); in parallel only the slowest counts.")
    sheet["K9"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.merge_cells("K9:U11")

    # Measured latency per application, from the LatencyData query (one row per application with UX and BS).
    table_top = total + 5
    sheet.cell(row=table_top - 1, column=1, value="Measured per application (APIM, requests-weighted, month of the "
                                                   f"data; at least {MIN_UX_REQUESTS:,} UX requests)").font = BOLD
    l = lambda column: r("LatencyApps", column)
    formulas = [("application", pull("LatencyData", "cod_app", table_top)),
                ("ux_requests", pull("LatencyData", "ux_requests", table_top)),
                ("ux_backend_ms", pull("LatencyData", "ux_backend_ms", table_top)),
                ("bs_requests", pull("LatencyData", "bs_requests", table_top)),
                ("bs_gateway_ms", pull("LatencyData", "bs_gateway_ms", table_top)),
                ("ux_minus_bs", f'{l("ux_backend_ms")}-{l("bs_gateway_ms")}'),
                ("bs_calls_per_ux", f'IF({l("ux_requests")}>0,{l("bs_requests")}/{l("ux_requests")},0)')]
    formula_table(sheet, "LatencyApps", formulas, app_count, top=table_top)
    for row in range(table_top + 1, table_top + 1 + app_count):
        for column, fmt in ((2, "#,##0"), (3, "0"), (4, "#,##0"), (5, "0"), (6, "0"), (7, "0.0")):
            sheet.cell(row=row, column=column).number_format = fmt
    sheet.cell(row=table_top + app_count + 2, column=1, value=(
        "ux_minus_bs: what UX adds over what its BS answers. Negative: the UX APIs do not call these BS APIs (other "
        "application, cache or only the fast ones). bs_calls_per_ux: BS requests per UX request; above 1, UX calls BS "
        "several times (only meaningful when that BS serves this UX alone)."))
    for column, width in zip("ABCDEFGH", (34, 14, 18, 14, 12, 12, 14, 18)):
        sheet.column_dimensions[column].width = width


def build(results, output, criteria="v2"):
    data = model.load(results)
    clusters = data["clusters"]
    days = {row["day_type"]: int(row["days"]) for row in data["profile"]}
    month = clusters[0]["month"] if clusters else ""
    book = Workbook()
    names = ["Results", "Top costs", "Criteria", "Apps", "Calculator", "Latency", "Traffic", "Services", "Prices", "Variables",
             "Data", "TrafficData", "LatencyData", "LatencyLayers", "AllClusters", "Lists", "UsageData"]
    sheets = {}
    for index, name in enumerate(names):
        sheets[name] = book.active if index == 0 else book.create_sheet(name)
    sheets["Results"].title = "Results"
    write_variables(book, sheets["Variables"], month, days, data["prices"]["price_env_hour"])
    v1wb.write_prices(book, sheets["Prices"], v1wb.prices_rows(results))
    clusters_data = cluster_rows(clusters)
    traffic_data = traffic_rows(data["profile"])
    services_data = service_rows(data["services"], clusters)
    inputs = latency_inputs(results)
    latency_apps, latency_layers = latency_app_rows(inputs), latency_layer_rows(inputs)
    codes = sorted({row["cod_app"] for row in clusters})
    scopes = sorted(row["cluster"] for row in clusters) + codes
    write_query_table(sheets["Data"], "ClusterData", CLUSTER_COLUMNS, clusters_data)
    write_query_table(sheets["TrafficData"], "TrafficData", TRAFFIC_COLUMNS, traffic_data)
    write_query_table(sheets["Services"], "Services", SERVICE_COLUMNS, services_data)
    write_query_table(sheets["LatencyData"], "LatencyData", LATENCY_APP_COLUMNS, latency_apps)
    write_query_table(sheets["LatencyLayers"], "LatencyLayers", LATENCY_LAYER_COLUMNS, latency_layers)
    write_query_table(sheets["Lists"], "Scopes", ["scope"], [[scope] for scope in scopes])
    all_clusters = all_cluster_rows(results)
    write_query_table(sheets["AllClusters"], "AllClusters", ALL_CLUSTER_COLUMNS, all_clusters)
    write_query_table(sheets["UsageData"], "UsageData", USAGE_COLUMNS, usage_rows(results))
    write_traffic(sheets["Traffic"], len(traffic_data))
    write_criteria(sheets["Criteria"], len(clusters_data))
    write_apps(sheets["Apps"], codes)
    write_results(sheets["Results"])
    write_top_costs(book, sheets["Top costs"], len(clusters_data), len(all_clusters))
    example = next((row["cluster"] for row in clusters if row["cod_app"] == deck.EXAMPLE_APP), clusters[0]["cluster"])
    write_calculator(sheets["Calculator"], example)
    latency_default = deck.EXAMPLE_APP if any(row[0] == deck.EXAMPLE_APP for row in latency_apps) else \
        (latency_apps[0][0] if latency_apps else "")
    write_latency(book, sheets["Latency"], len(latency_apps), latency_default)
    for name in HIDDEN_SHEETS:  # raw query output: the visible sheets read it
        sheets[name].sheet_state = "hidden"
    book.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    aa.keep_version(output)
    book.save(output)
    folder = os.path.abspath(results)
    pq.load_queries(output, [
        ("Data", "ClusterData", "ClusterData", query_clusters(folder), CLUSTER_COLUMNS),
        ("TrafficData", "TrafficData", "TrafficData", query_traffic(folder), TRAFFIC_COLUMNS),
        ("Services", "Services", "Services", query_services(folder), SERVICE_COLUMNS),
        ("LatencyData", "LatencyData", "LatencyData", query_latency_apps(folder), LATENCY_APP_COLUMNS),
        ("LatencyLayers", "LatencyLayers", "LatencyLayers", query_latency_layers(folder), LATENCY_LAYER_COLUMNS),
        ("Lists", "Scopes", "Scopes", query_scopes(folder), ["scope"]),
        ("Apps", "AppList", "AppList", query_apps(folder), ["cod_app"]),
        ("AllClusters", "AllClusters", "AllClusters", query_all_clusters(folder), ALL_CLUSTER_COLUMNS),
        ("UsageData", "UsageData", "UsageData", query_usage(folder), USAGE_COLUMNS)])
    aa.azure.LOG.info("Wrote %s (criteria %s, %d clusters, 9 Power Query connections to %s)", output, criteria,
                      len(clusters), folder)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--criteria", choices=["v1"] + sorted(model.PRESETS), default="v2",
                        help="Criteria v1 (aa_criteria_workbook_v1.py), or the variables' defaults of v2 or v3 "
                             "(aa_v2_model.PRESETS)")
    parser.add_argument("--set", action="append", default=[], metavar="NAME=VALUE",
                        help="Change a variable's default, e.g. --set rule_shared=NO (repeatable; give the workbook "
                             "and the deck the same ones)")
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: aks_aca_criteria_<criteria>.xlsx in the newest reports/<date>-aks-aca")
    args, rest = parser.parse_known_args()
    if args.criteria == "v1":
        deck_entry.run_v1(args, rest, "aks_aca_criteria_v1.xlsx", v1wb.main)
        return
    if rest:
        parser.error("unrecognized arguments: " + " ".join(rest))
    aa.azure.setup_logging(False)
    model.use(args.criteria, args.set)
    output = args.output
    if not output:
        folders = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
        folder = folders[-1] if folders else os.path.join(aa.TOPIC_DIR, "reports", "aks-aca")
        output = os.path.join(folder, f"aks_aca_criteria_{args.criteria}.xlsx")
    build(args.results, output, args.criteria)


if __name__ == "__main__":
    main()
