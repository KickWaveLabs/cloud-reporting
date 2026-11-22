#!/usr/bin/env python3
"""Write the criteria workbook: which conditions each PROD application meets and what moving it to Container Apps
would save, from the CSV files in results/.

Three tables hold the CSV data as it is (Apps, Services, Prices); everything else is an Excel formula over those
tables and the yellow cells, so every figure can be traced and changed in Excel. Next to the workbook,
aks_aca_criteria.pq holds the Power Query (M) code that builds the same three tables from the CSV files, one query per
step, to read how they are linked or to load them in Excel (Data > Get Data > Blank Query > Advanced Editor):

  ResultsFolder (parameter)          folder with the aa-<source> result folders; change it to point elsewhere
  src_application_snapshot   ┐
  src_cluster_snapshot       │ connection only: each reads one CSV file as it is
  src_deployment_snapshot    │
  src_prices                 ┘
  Apps      -> sheet Criteria   one row per PROD application (application_snapshot.csv)
  Services  -> sheet Services   one row per "micro"-service: deployment_snapshot.csv joined with cluster_snapshot.csv
                                on cluster_id, keeping the PROD clusters with applications (that gives its cod_app)
  Prices    -> sheet Prices     the Container Apps and load balancer list prices (prices.csv, East US 2)

Formula columns (right of each loaded table, kept by Excel on refresh) apply the conditions of the corporate deck
(aa_deck_exec.criteria: peak hour traffic, "micro"-services, critical, growing) and its cost model (aca_breakdown):
measured replicas in office hours and, outside office hours, the replicas per "micro"-service of the
OffhoursReplicas sheet: empty takes default_offhours_replicas (1, as the deck), 0 lets that application scale to
zero. Results rebuilds the result slide (with an editable chart) and Summary the funnel and the saving; with the
defaults both give the deck's figures.

Reads only results/ (no Azure call). The AKS discount (amortized cost of PROD node VMs over their list price) is
measured here as in aa_deck_exec.py and written as a yellow cell; --aks-discount fixes it by hand.
Output: aks_aca_criteria.xlsx and aks_aca_criteria.pq in the newest reports/<date>-aks-aca folder (or --output). The
.pq file keeps the path of --results in ResultsFolder; on another machine, change that line.
"""

import argparse
import glob
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

import aa_common as aa
import aa_deck_exec as deck

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
FORMULA_FILL = PatternFill("solid", fgColor="2E6B4F")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
REGION = "eastus2"
PRICE_KEYS = ["aca_vcpu_active_s", "aca_vcpu_idle_s", "aca_mem_active_gib_s", "aca_mem_idle_gib_s",
              "aca_requests_million", "lb_standard_rules_hour", "lb_standard_data_gb", "aks_standard_tier_hour",
              f"vm_{deck.MIN_AKS_VM}_payg", f"vm_{deck.MIN_AKS_VM}_{deck.MIN_AKS_RESERVATION}",
              f"disk_{deck.MIN_AKS_DISK}_month"]

# Columns each loaded query returns, in order (the M code below selects exactly these).
APPS_COLUMNS = ["cod_app", "cluster_names", "bia", "microservices", "microservices_projected", "traffic_status",
                "traffic_layer", "busiest_hour_tps", "busiest_hour", "busiest_offhours_hour_tps",
                "offhours_active_share", "requests_month", "cost_month_usd", "possible_monolith", "no_traffic_reason",
                "adt_runtime"]
APPS_NUMBERS = ["microservices", "microservices_projected", "busiest_hour_tps", "busiest_offhours_hour_tps",
                "offhours_active_share", "requests_month", "cost_month_usd"]
SERVICES_COLUMNS = ["cod_app", "cluster", "namespace", "deployment", "cpu_request_vcpu", "mem_request_gib",
                    "ci_replicas_avg_office", "hpa_min"]
SERVICES_NUMBERS = ["cpu_request_vcpu", "mem_request_gib", "ci_replicas_avg_office", "hpa_min"]
PRICES_COLUMNS = ["key", "region", "meter", "unit", "retail_price"]


def m_list(values):
    return "{" + ", ".join(f'"{value}"' for value in values) + "}"


def m_types(numbers, columns):
    return "{" + ", ".join(f'{{"{column}", {"type number" if column in numbers else "type text"}}}'
                           for column in columns) + "}"


def section_m(results_folder):
    """The Power Query (M) code of the workbook, one shared query per step of the data flow."""
    folder = results_folder.replace('"', '""')
    csv = '[Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]'

    def source(name, folder_name, file_name, description):
        return f'''// {description}
shared {name} = let
    Source = Csv.Document(File.Contents(ResultsFile("{folder_name}", "{file_name}")), {csv}),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers;
'''
    return f'''section Section1;

// Folder with the aa-<source> result folders (aks-aca/results). Change it to read another copy of the results.
shared ResultsFolder = "{folder}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true];

// Full path of one result file, with the separator of the folder above (Windows or macOS).
shared ResultsFile = (folder as text, file as text) as text =>
    let
        Separator = if Text.Contains(ResultsFolder, "\\") then "\\" else "/",
        Base = if Text.EndsWith(ResultsFolder, Separator) then ResultsFolder else ResultsFolder & Separator
    in
        Base & folder & Separator & file;

{source("src_application_snapshot", "aa-k8s-snapshot", "application_snapshot.csv",
        "One row per PROD application, as aa_k8s_snapshot.py wrote it.")}
{source("src_cluster_snapshot", "aa-k8s-snapshot", "cluster_snapshot.csv",
        "One row per cluster with its category (APPLICATION, INFRASTRUCTURE...) and cod_app.")}
{source("src_deployment_snapshot", "aa-k8s-snapshot", "deployment_snapshot.csv",
        "One row per deployment of the platform team's export, with Container Insights replicas.")}
{source("src_prices", "aa-prices", "prices.csv", "Azure list prices (Retail Prices API), as aa_prices.py wrote them.")}
// Sheet Criteria: the application columns the conditions and the cost use. CSV numbers use a dot (en-US).
shared Apps = let
    Columns = Table.SelectColumns(src_application_snapshot, {m_list(APPS_COLUMNS)}),
    Blanks = Table.ReplaceValue(Columns, "", null, Replacer.ReplaceValue, {m_list(APPS_NUMBERS)}),
    Typed = Table.TransformColumnTypes(Blanks, {m_types(APPS_NUMBERS, APPS_COLUMNS)}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"cod_app", Order.Ascending}}}})
in
    Sorted;

// Sheet Services: each "micro"-service of a PROD cluster with applications. The join on cluster_id gives its
// cod_app and drops infrastructure clusters, clusters without applications and other environments.
shared Services = let
    ApplicationClusters = Table.SelectRows(src_cluster_snapshot, each [environment] = "PROD" and [category] = "APPLICATION"),
    ClusterApp = Table.SelectColumns(ApplicationClusters, {{"cluster_id", "cod_app"}}),
    Microservices = Table.SelectRows(src_deployment_snapshot, each [microservice] = "YES"),
    Deployments = Table.RemoveColumns(Microservices, {{"cod_app"}}, MissingField.Ignore),
    Linked = Table.NestedJoin(Deployments, {{"cluster_id"}}, ClusterApp, {{"cluster_id"}}, "application", JoinKind.Inner),
    WithApp = Table.ExpandTableColumn(Linked, "application", {{"cod_app"}}),
    Columns = Table.SelectColumns(WithApp, {m_list(SERVICES_COLUMNS)}),
    Blanks = Table.ReplaceValue(Columns, "", null, Replacer.ReplaceValue, {m_list(SERVICES_NUMBERS)}),
    Typed = Table.TransformColumnTypes(Blanks, {m_types(SERVICES_NUMBERS, SERVICES_COLUMNS)}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"cod_app", Order.Ascending}}, {{"cluster", Order.Ascending}}, {{"namespace", Order.Ascending}}, {{"deployment", Order.Ascending}}}})
in
    Sorted;

// Sheet Prices: the list prices of the cost model in {REGION}, one row per key.
shared Prices = let
    Wanted = Table.SelectRows(src_prices, each [region] = "{REGION}" and List.Contains({m_list(PRICE_KEYS)}, [key])),
    Columns = Table.SelectColumns(Wanted, {m_list(PRICES_COLUMNS)}),
    Distinct = Table.Distinct(Columns, {{"key"}}),
    Typed = Table.TransformColumnTypes(Distinct, {m_types(["retail_price"], PRICES_COLUMNS)}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"key", Order.Ascending}}}})
in
    Sorted;
'''


# ---------------------------------------------------------------------------
# The rows the queries return, computed here the same way, so the workbook is complete before the first refresh.
# ---------------------------------------------------------------------------


def number_or_none(value):
    return deck.to_float(value) if value not in ("", None) else None


def typed_rows(rows, columns, numbers):
    return [[number_or_none(row.get(column)) if column in numbers else (row.get(column) or None) for column in columns]
            for row in rows]


def query_rows(results):
    apps = sorted(deck.results_csv(results, "aa-k8s-snapshot", "application_snapshot.csv"), key=lambda r: r["cod_app"])
    clusters = {row["cluster_id"]: row["cod_app"]
                for row in deck.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["category"] == aa.APPLICATION}
    services = [dict(row, cod_app=clusters[row["cluster_id"]])
                for row in deck.results_csv(results, "aa-k8s-snapshot", "deployment_snapshot.csv")
                if row["microservice"] == "YES" and row["cluster_id"] in clusters]
    services.sort(key=lambda r: (r["cod_app"], r["cluster"], r["namespace"], r["deployment"]))
    prices, seen = [], set()
    for row in deck.results_csv(results, "aa-prices", "prices.csv"):
        if row["region"] == REGION and row["key"] in PRICE_KEYS and row["key"] not in seen:
            seen.add(row["key"])
            prices.append(row)
    prices.sort(key=lambda r: r["key"])
    return (typed_rows(apps, APPS_COLUMNS, APPS_NUMBERS), typed_rows(services, SERVICES_COLUMNS, SERVICES_NUMBERS),
            typed_rows(prices, PRICES_COLUMNS, ["retail_price"]), apps)


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------


def this_row(table, column):
    return f"{table}[[#This Row],[{column}]]"


def price_of(key):
    return f'INDEX(Prices[retail_price],MATCH("{key}",Prices[key],0))'


def assumptions(month, aks_ratio):
    """(name, value or formula, unit, note, editable) of the Assumptions sheet; each value gets a workbook name."""
    return [
        ("cost_month", month, "YYYY-MM", "Month of the cost and metrics in results/ (last full month).", True),
        ("max_tps", deck.MAX_TPS, "TPS", "Condition 1: peak hour traffic up to this many transactions per second.", True),
        ("max_microservices", deck.MAX_MICROSERVICES, "count",
         "Condition 2: up to this many \"micro\"-services in PROD. More with DESA marks the app as growing.", True),
        ("default_offhours_replicas", deck.OFFHOURS_REPLICAS, "replicas",
         "Replicas each \"micro\"-service keeps on outside office hours when OffhoursReplicas has no value for its "
         "application. 0 = scale to zero (the first request waits up to ~30 s).", True),
        ("offhours_rate", "ACTIVE", "ACTIVE / IDLE",
         "Rate of the off-hours replicas. ACTIVE is the safe side; IDLE applies only while a replica gets no "
         "request and uses under 0.01 vCPU.", True),
        ("requests_buffer", deck.REQUESTS_BUFFER, "factor", "Headroom on the requests of the month.", True),
        ("environments_per_app", deck.ENVIRONMENTS_PER_APP, "count",
         "Container Apps environments per application; each bills its internal load balancer.", True),
        ("bank_discount", deck.BANK_DISCOUNT, "share", "The bank's discount on Azure list prices (Container Apps).",
         True),
        ("aks_paid_ratio", aks_ratio, "share",
         "Share of the list price the bank pays for AKS node VMs (reservations included), measured in the month.",
         True),
        ("exchange_rate", deck.EXCHANGE_RATE, "PEN per USD", "Exchange rate of the decks.", True),
        ("igv", deck.IGV, "share", "IGV added to every amount in soles.", True),
        ("hours_month", deck.HOURS_MONTH, "hours", "Hours a month the load balancer is billed (as the calculator).",
         True),
        ("min_vcpu", 0.25, "vCPU", "Smallest Container Apps replica.", False),
        ("vcpu_step", 0.25, "vCPU", "Container Apps sizes go in steps of 0.25 vCPU.", False),
        ("gib_per_vcpu", 2, "GiB", "Container Apps Consumption: 2 GiB of memory per vCPU.", False),
        ("office_seconds",
         "=NETWORKDAYS(DATE(LEFT(cost_month,4),RIGHT(cost_month,2),1),EOMONTH(DATE(LEFT(cost_month,4),RIGHT(cost_month,2),1),0))"
         f"*{aa.OFFICE_END_HOUR - aa.OFFICE_START_HOUR}*3600", "seconds",
         f"Monday to Friday {aa.OFFICE_START_HOUR}:00-{aa.OFFICE_END_HOUR}:00 in the month (holidays not excluded).",
         False),
        ("offhours_seconds", "=DAY(EOMONTH(DATE(LEFT(cost_month,4),RIGHT(cost_month,2),1),0))*86400-office_seconds", "seconds",
         "The rest of the month.", False),
        ("price_vcpu_active", "=" + price_of("aca_vcpu_active_s"), "USD per vCPU-s", "Sheet Prices.", False),
        ("price_vcpu_idle", "=" + price_of("aca_vcpu_idle_s"), "USD per vCPU-s", "Sheet Prices.", False),
        ("price_gib_active", "=" + price_of("aca_mem_active_gib_s"), "USD per GiB-s", "Sheet Prices.", False),
        ("price_gib_idle", "=" + price_of("aca_mem_idle_gib_s"), "USD per GiB-s", "Sheet Prices.", False),
        ("price_requests_million", "=" + price_of("aca_requests_million"), "USD per million", "Sheet Prices.", False),
        ("price_lb_hour", "=" + price_of("lb_standard_rules_hour"), "USD per hour", "Sheet Prices.", False),
        ("price_lb_data_gb", "=" + price_of("lb_standard_data_gb"), "USD per GB", "Sheet Prices.", False),
        ("price_aks_tier_hour", "=" + price_of("aks_standard_tier_hour"), "USD per hour",
         "AKS Standard tier (Uptime SLA). Sheet Prices.", False),
        ("price_node_payg_hour", "=" + price_of(f"vm_{deck.MIN_AKS_VM}_payg"), "USD per hour",
         f"{deck.MIN_AKS_VM}, pay as you go. Sheet Prices.", False),
        ("price_node_reserved_term", "=" + price_of(f"vm_{deck.MIN_AKS_VM}_{deck.MIN_AKS_RESERVATION}"), "USD per term",
         f"{deck.MIN_AKS_VM}, 3-year reservation, whole term (26,280 h). Sheet Prices.", False),
        ("price_disk_month", "=" + price_of(f"disk_{deck.MIN_AKS_DISK}_month"), "USD per month",
         f"Managed OS disk {deck.MIN_AKS_DISK}. Sheet Prices.", False),
    ]


def apps_formulas():
    """(column, formula without '=') of the Criteria sheet, after the query columns."""
    r = lambda column: this_row("Apps", column)
    offhours = f"INDEX(OffhoursReplicas[offhours_replicas],MATCH({r('cod_app')},OffhoursReplicas[cod_app],0))"
    return [
        ("has_traffic", f'IF(AND({r("traffic_status")}="MEASURED",ISNUMBER({r("busiest_hour_tps")})),"YES","NO")'),
        ("c1_peak_hour_tps_ok", f'IF({r("has_traffic")}="NO","NO_DATA",IF({r("busiest_hour_tps")}<=max_tps,"YES","NO"))'),
        ("c2_microservices_ok", f'IF(N({r("microservices")})<=max_microservices,"YES","NO")'),
        ("critical", f'IF(OR({r("bia")}="MUY ALTA",{r("bia")}="ALTA"),"YES","NO")'),
        ("growing", f'IF(AND(N({r("microservices_projected")})>max_microservices,'
                    f'N({r("microservices")})<=max_microservices),"YES","NO")'),
        ("decision", f'IF({r("has_traffic")}="NO","NO_TRAFFIC",IF(OR({r("c1_peak_hour_tps_ok")}="NO",'
                     f'{r("c2_microservices_ok")}="NO"),"AKS",IF(OR({r("critical")}="YES",{r("growing")}="YES"),'
                     f'"ACA_WITH_REVIEW","ACA")))'),
        ("moves", f'IF(OR({r("decision")}="ACA",{r("decision")}="ACA_WITH_REVIEW"),1,0)'),
        ("offhours_replicas", f"IFERROR(IF(ISNUMBER({offhours}),{offhours},default_offhours_replicas),"
                              "default_offhours_replicas)"),
        ("aca_office_usd", f"SUMIFS(Services[office_cost_usd],Services[cod_app],{r('cod_app')})"),
        ("aca_offhours_usd", f"{r('offhours_replicas')}*SUMIFS(Services[offhours_cost_per_replica_usd],"
                             f"Services[cod_app],{r('cod_app')})"),
        ("aca_requests_usd", f"N({r('requests_month')})*requests_buffer/1000000*price_requests_million"),
        ("aca_environment_usd", "environments_per_app*hours_month*price_lb_hour"),
        ("aca_list_usd", f"{r('aca_office_usd')}+{r('aca_offhours_usd')}+{r('aca_requests_usd')}"
                         f"+{r('aca_environment_usd')}"),
        ("aca_discounted_usd", f"{r('aca_list_usd')}*(1-bank_discount)"),
        ("saving_month_usd", f"{r('moves')}*(N({r('cost_month_usd')})-{r('aca_discounted_usd')})"),
        ("saving_month_pen_igv", f"{r('saving_month_usd')}*exchange_rate*(1+igv)"),
        ("saving_year_pen_igv", f"{r('saving_month_pen_igv')}*12"),
        ("result_group", f'IF({r("decision")}<>"NO_TRAFFIC",{r("decision")},IF({r("no_traffic_reason")}="IN_DEVELOPMENT",'
                         f'"IN_DEVELOPMENT",IF({r("no_traffic_reason")}="IN_USE_NOT_MEASURED","IN_USE","NOT_IN_USE")))'),
        ("new_cost_usd", f"IF({r('moves')}=1,{r('aca_discounted_usd')},N({r('cost_month_usd')}))"),
    ]


def services_formulas():
    r = lambda column: this_row("Services", column)
    rate = lambda active, idle: f'IF(offhours_rate="IDLE",{idle},{active})'
    return [
        ("aca_vcpu", f"CEILING(MAX(N({r('cpu_request_vcpu')}),N({r('mem_request_gib')})/gib_per_vcpu,min_vcpu),vcpu_step)"),
        ("aca_gib", f"{r('aca_vcpu')}*gib_per_vcpu"),
        ("replicas_office", f"MAX(1,IF(N({r('ci_replicas_avg_office')})>0,{r('ci_replicas_avg_office')},"
                            f"IF(N({r('hpa_min')})>0,{r('hpa_min')},1)))"),
        ("office_cost_usd", f"{r('replicas_office')}*office_seconds*({r('aca_vcpu')}*price_vcpu_active"
                            f"+{r('aca_gib')}*price_gib_active)"),
        ("offhours_cost_per_replica_usd", f"offhours_seconds*({r('aca_vcpu')}*"
                                          f"{rate('price_vcpu_active', 'price_vcpu_idle')}+{r('aca_gib')}*"
                                          f"{rate('price_gib_active', 'price_gib_idle')})"),
    ]


def write_table(sheet, name, columns, rows, formulas=(), style="TableStyleMedium2", widths=None):
    """A table with the query columns (values) and formula columns (calculated columns of the table)."""
    headers = columns + [column for column, _ in formulas]
    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=index, value=header)
        cell.font = WHITE_BOLD
        cell.fill = FORMULA_FILL if index > len(columns) else HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row_index, values in enumerate(rows, start=2):
        for column_index, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)
        for offset, (_, formula) in enumerate(formulas):
            sheet.cell(row=row_index, column=len(columns) + 1 + offset, value="=" + formula)
    last = max(len(rows), 1) + 1
    table = Table(displayName=name, ref=f"A1:{get_column_letter(len(headers))}{last}")
    table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    table.autoFilter = AutoFilter(ref=table.ref)
    table.tableColumns = [TableColumn(id=index, name=header) for index, header in enumerate(headers, start=1)]
    for offset, (_, formula) in enumerate(formulas):
        table.tableColumns[len(columns) + offset].calculatedColumnFormula = TableFormula(attr_text=formula)
    sheet.add_table(table)
    sheet.freeze_panes = "B2"
    sheet.row_dimensions[1].height = 32
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = (widths or {}).get(header, max(12, len(header) + 2))


def build(results, aks_ratio, output):
    apps_rows, services_rows, prices_rows, apps = query_rows(results)
    estate_month = deck.read_estate(results)["month"]
    book = Workbook()
    readme = book.active
    readme.title = "README"
    sheets = {name: book.create_sheet(name) for name in ("Results", "Calculator", "Minimums", "Summary", "Criteria", "OffhoursReplicas", "Services",
                                                         "Prices", "Assumptions")}

    # Assumptions: one workbook name per value; yellow cells are inputs.
    sheet = sheets["Assumptions"]
    for column, header in enumerate(("Name", "Value", "Unit", "Note"), start=1):
        sheet.cell(row=1, column=column, value=header).font = WHITE_BOLD
        sheet.cell(row=1, column=column).fill = HEADER_FILL
    for row, (name, value, unit, note, editable) in enumerate(assumptions(estate_month, aks_ratio), start=2):
        sheet.cell(row=row, column=1, value=name)
        cell = sheet.cell(row=row, column=2, value=value)
        if editable:
            cell.fill = INPUT_FILL
        sheet.cell(row=row, column=3, value=unit)
        sheet.cell(row=row, column=4, value=note)
        book.defined_names[name] = DefinedName(name, attr_text=f"Assumptions!$B${row}")
    for column, width in zip("ABCD", (26, 16, 16, 100)):
        sheet.column_dimensions[column].width = width

    write_table(sheets["Prices"], "Prices", PRICES_COLUMNS, prices_rows,
                widths={"meter": 40, "retail_price": 14})
    write_table(sheets["Services"], "Services", SERVICES_COLUMNS, services_rows, services_formulas(),
                widths={"cluster": 20, "namespace": 22, "deployment": 36, "offhours_cost_per_replica_usd": 18})
    write_table(sheets["Criteria"], "Apps", APPS_COLUMNS, apps_rows, apps_formulas(),
                widths={"cluster_names": 24, "no_traffic_reason": 20, "busiest_hour": 16})

    # Off-hours replicas per application: the only per-application input. Empty = default_offhours_replicas.
    sheet = sheets["OffhoursReplicas"]
    lookup = lambda column: f"INDEX(Apps[{column}],MATCH({this_row('OffhoursReplicas', 'cod_app')},Apps[cod_app],0))"
    write_table(sheet, "OffhoursReplicas", ["cod_app", "offhours_replicas", "note"],
                [[row["cod_app"], None, None] for row in apps],
                [("decision", f'IFERROR({lookup("decision")},"NOT_IN_APPS")'),
                 ("offhours_active_share", f'IFERROR({lookup("offhours_active_share")},"")'),
                 ("busiest_offhours_hour_tps", f'IFERROR({lookup("busiest_offhours_hour_tps")},"")'),
                 ("replicas_used", f'IFERROR({lookup("offhours_replicas")},"")')],
                style="TableStyleLight9", widths={"note": 40, "offhours_active_share": 16})
    for row in range(2, len(apps) + 2):
        sheet.cell(row=row, column=2).fill = INPUT_FILL
        sheet.cell(row=row, column=5).number_format = "0%"

    write_summary(sheets["Summary"])
    write_results(sheets["Results"])
    candidates = [row["cod_app"] for row in apps if deck.decision(row) in deck.MOVABLE]
    write_calculator(book, sheets["Calculator"], candidates[0] if candidates else apps[0]["cod_app"])
    write_minimums(sheets["Minimums"])
    write_readme(readme, results)
    for name, sheet in sheets.items():
        if name in ("Criteria", "Services"):
            money_columns(sheet)
    book.calculation.fullCalcOnLoad = True
    book.save(output)
    with open(os.path.splitext(output)[0] + ".pq", "w", encoding="utf-8") as handle:
        handle.write(section_m(os.path.abspath(results)))
    return len(apps_rows), len(services_rows)


def money_columns(sheet):
    for column in range(1, sheet.max_column + 1):
        header = str(sheet.cell(row=1, column=column).value)
        if header.endswith("_usd") or header.endswith("_pen_igv"):
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=column).number_format = "#,##0.00"


def write_summary(sheet):
    """The funnel (no longer a slide), the decisions (slide 5) and the saving (slide 6), as formulas over Criteria."""
    moves = "Apps[moves],1"
    lines = [
        ("Funnel", None, None),
        ("PROD applications", "=ROWS(Apps[cod_app])", "count"),
        ("With measured traffic", '=COUNTIFS(Apps[has_traffic],"YES")', "count"),
        ("+ peak hour up to max_tps", '=COUNTIFS(Apps[has_traffic],"YES",Apps[c1_peak_hour_tps_ok],"YES")', "count"),
        ("+ up to max_microservices", '=COUNTIFS(Apps[has_traffic],"YES",Apps[c1_peak_hour_tps_ok],"YES",'
                                      'Apps[c2_microservices_ok],"YES")', "count"),
        ("   of which to validate (critical or growing)", '=COUNTIFS(Apps[decision],"ACA_WITH_REVIEW")', "count"),
        ("   of which keep 0 replicas off hours", f'=COUNTIFS({moves},Apps[offhours_replicas],0)', "count"),
        (None, None, None),
        ("Decision (slide 5)", None, None),
        ("AKS", '=COUNTIFS(Apps[decision],"AKS")', "count"),
        ("ACA_WITH_REVIEW", '=COUNTIFS(Apps[decision],"ACA_WITH_REVIEW")', "count"),
        ("ACA", '=COUNTIFS(Apps[decision],"ACA")', "count"),
        ("NO_TRAFFIC", '=COUNTIFS(Apps[decision],"NO_TRAFFIC")', "count"),
        (None, None, None),
        ("Cost of the applications that move (slide 6), per month", None, None),
        ("AKS paid today", f"=SUMIFS(Apps[cost_month_usd],{moves})", "USD"),
        ("Container Apps at list price", f"=SUMIFS(Apps[aca_list_usd],{moves})", "USD"),
        ("   of which off-hours replicas", f"=SUMIFS(Apps[aca_offhours_usd],{moves})", "USD"),
        ("Container Apps with the bank's discount", f"=SUMIFS(Apps[aca_discounted_usd],{moves})", "USD"),
        ("Saving at list price (AKS at list - ACA at list)", "=B17/aks_paid_ratio-B18", "USD"),
        ("Saving with discounts (AKS paid - ACA discounted)", "=B17-B20", "USD"),
        ("Saving with discounts per month", "=B22*exchange_rate*(1+igv)", "PEN with IGV"),
        ("Saving with discounts per year", "=B23*12", "PEN with IGV"),
        ("Saving at list price per year", "=B21*exchange_rate*(1+igv)*12", "PEN with IGV"),
    ]
    sheet.cell(row=1, column=1, value="Summary (formulas over sheet Criteria; change yellow cells in Assumptions "
                                       "and OffhoursReplicas)").font = Font(bold=True, size=12)
    for row, (label, formula, unit) in enumerate(lines, start=2):
        if label and formula is None:
            sheet.cell(row=row, column=1, value=label).font = Font(bold=True, color="1F3A5F")
            continue
        sheet.cell(row=row, column=1, value=label)
        if formula:
            cell = sheet.cell(row=row, column=2, value=formula)
            cell.number_format = "#,##0" if unit == "count" else "#,##0.00"
            sheet.cell(row=row, column=3, value=unit)
    sheet.column_dimensions["A"].width = 52
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 14


# ---------------------------------------------------------------------------
# Calculators in the user's format: Meter | Units | Cost / Unit | Seconds | Gross Total | Discount |
# Gross Total + Dsct | Net Total (with IGV) | Net Total (S/).
# ---------------------------------------------------------------------------

CALC_FILL = PatternFill("solid", fgColor="4472C4")
CALC_HEADERS = ["Meter", "Units", "Cost / Unit", "Seconds", "Gross Total", "Discount", "Gross Total + Dsct",
                "Net Total", "Net Total (S/)"]
USD, PEN, PCT = '"$"#,##0.00', '"S/ "#,##0.00', "0.00%"


def calculator_block(sheet, top, title, lines, time_header="Seconds"):
    """A title, the header row and one row per meter (meter, units, cost per unit, time, discount); the money
    columns and the total row are formulas. Returns the total row."""
    sheet.cell(row=top, column=3, value=title).font = Font(bold=True)
    for column, header in enumerate(CALC_HEADERS, start=1):
        cell = sheet.cell(row=top + 1, column=column, value=time_header if header == "Seconds" else header)
        cell.font, cell.fill = WHITE_BOLD, CALC_FILL
    row = top + 2
    for meter, units, cost, time, discount in lines:
        for column, value in enumerate((meter, units, cost, time), start=1):
            sheet.cell(row=row, column=column, value=value).number_format = "#,##0.00####" if column == 3 else "#,##0.####"
        sheet.cell(row=row, column=5, value=f"=B{row}*C{row}*D{row}").number_format = USD
        sheet.cell(row=row, column=6, value=discount).number_format = PCT
        sheet.cell(row=row, column=7, value=f"=E{row}*(1-F{row})").number_format = USD
        sheet.cell(row=row, column=8, value=f"=G{row}*(1+igv)").number_format = USD
        sheet.cell(row=row, column=9, value=f"=H{row}*exchange_rate").number_format = PEN
        row += 1
    sheet.cell(row=row, column=1, value="Total").font = Font(bold=True)
    for column, number_format in ((5, USD), (7, USD), (8, USD), (9, PEN)):
        letter = get_column_letter(column)
        cell = sheet.cell(row=row, column=column, value=f"=SUM({letter}{top + 2}:{letter}{row - 1})")
        cell.font, cell.number_format = Font(bold=True), number_format
    return row


def write_calculator(book, sheet, default_app):
    """The cost of one application, in the user's calculator format: Container Apps (requests, vCPU and memory in
    and out of office hours), its load balancer, and AKS as billed today. Pick the application in C3."""
    app = lambda column: f"INDEX(Apps[{column}],MATCH(calc_app,Apps[cod_app],0))"
    services = lambda column: f"SUMPRODUCT((Services[cod_app]=calc_app)*Services[replicas_office]*Services[{column}])"
    per_replica = lambda column: f"SUMIFS(Services[{column}],Services[cod_app],calc_app)"
    rate = lambda active, idle: f'=IF(offhours_rate="IDLE",{idle},{active})'
    sheet.cell(row=1, column=1, value="Cost calculator per application (formulas over Criteria, Services and "
                                      "Assumptions)").font = Font(bold=True, size=12)
    sheet.cell(row=3, column=1, value="Application (cod_app)").font = Font(bold=True)
    cell = sheet.cell(row=3, column=3, value=default_app)
    cell.fill, cell.font = INPUT_FILL, Font(bold=True)
    book.defined_names["calc_app"] = DefinedName("calc_app", attr_text="Calculator!$C$3")
    validation = DataValidation(type="list", formula1='=INDIRECT("Apps[cod_app]")', allow_blank=False)
    sheet.add_data_validation(validation)
    validation.add("C3")
    context = [("decision", app("decision")), ("c1 peak hour <= max_tps", app("c1_peak_hour_tps_ok")),
               ("c2 micro-services <= max", app("c2_microservices_ok")), ("critical", app("critical")),
               ("growing", app("growing")), ("micro-services", app("microservices")),
               ("off-hours replicas", app("offhours_replicas"))]
    for index, (label, formula) in enumerate(context):
        sheet.cell(row=4 + index, column=1, value=label)
        sheet.cell(row=4 + index, column=3, value=f"=IFERROR({formula},\"\")")
    # Inputs of the month, as the calculator's header row.
    for column, (label, formula) in enumerate((("weekdays (12 h)", "=office_seconds/3600/12"),
                                               ("office hours", "=office_seconds/3600"),
                                               ("office seconds", "=office_seconds"),
                                               ("off-hours seconds", "=offhours_seconds"),
                                               ("requests (M)", f"=IFERROR(N({app('requests_month')})*requests_buffer/1000000,0)"),
                                               ("ms per request", "=IFERROR(office_seconds*1000/(E13*1000000),\"\")")),
                                              start=1):
        header = sheet.cell(row=12, column=column, value=label)
        header.font, header.fill = WHITE_BOLD, CALC_FILL
        sheet.cell(row=13, column=column, value=formula).number_format = "#,##0.00"
    aca_total = calculator_block(sheet, 15, "ACA - COST - MONTHLY", [
        ("Requests (M)", "=E13", "=price_requests_million", 1, "=bank_discount"),
        ("vCPU office", f"={services('aca_vcpu')}", "=price_vcpu_active", "=office_seconds", "=bank_discount"),
        ("GiB office", f"={services('aca_gib')}", "=price_gib_active", "=office_seconds", "=bank_discount"),
        ("vCPU off hours", f"=IFERROR({app('offhours_replicas')},0)*{per_replica('aca_vcpu')}",
         rate("price_vcpu_active", "price_vcpu_idle"), "=offhours_seconds", "=bank_discount"),
        ("GiB off hours", f"=IFERROR({app('offhours_replicas')},0)*{per_replica('aca_gib')}",
         rate("price_gib_active", "price_gib_idle"), "=offhours_seconds", "=bank_discount"),
    ])
    lb_top = aca_total + 2
    lb_total = calculator_block(sheet, lb_top, "LOAD BALANCER - COST - MONTHLY", [
        ("Load Balancer rules", "=environments_per_app", "=price_lb_hour", "=hours_month", "=bank_discount"),
        ("Data processed (GB)", 0, "=price_lb_data_gb", 1, "=bank_discount"),
    ], time_header="Hours")
    sheet.cell(row=lb_top + 3, column=2).fill = INPUT_FILL
    aks_top = lb_total + 2
    aks_total = calculator_block(sheet, aks_top, "AKS - COST - MONTHLY (billed, amortized: discounts already in)", [
        ("AKS clusters of the application", 1, f"=IFERROR(N({app('cost_month_usd')}),0)", 1, 0),
    ], time_header="-")
    sheet.cell(row=aks_top + 2, column=3).number_format = USD
    top = aks_total + 2
    sheet.cell(row=top, column=3, value="COMPARISON - MONTHLY").font = Font(bold=True)
    rows = [("AKS today (S/)", f"=I{aks_total}"), ("ACA + load balancer (S/)", f"=I{aca_total}+I{lb_total}"),
            ("Saving per month (S/)", f"=B{top + 1}-B{top + 2}"), ("Saving per year (S/)", f"=B{top + 3}*12"),
            ("Check: ACA list price - Criteria (should be 0)",
             f"=ROUND(E{aca_total}+E{lb_total}-IFERROR({app('aca_list_usd')},0),2)")]
    for index, (label, formula) in enumerate(rows, start=1):
        sheet.cell(row=top + index, column=1, value=label).font = Font(bold=index in (3, 4))
        cell = sheet.cell(row=top + index, column=2, value=formula)
        cell.number_format, cell.font = (PEN if index <= 4 else "0.00"), Font(bold=index in (3, 4))
    sheet.cell(row=top + 7, column=1, value=(
        "Units of vCPU/GiB office = replicas measured in office hours x size of each \"micro\"-service (sheet "
        "Services); off hours = replicas of OffhoursReplicas x size. Discount and IGV as in Assumptions."))
    widths = {"A": 30, "B": 14, "C": 14, "D": 17, "E": 16, "F": 14, "G": 18, "H": 14, "I": 16}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def write_minimums(sheet):
    """Slide 3: the smallest AKS cluster and one Container Apps "micro"-service with one replica on 24 h, as the user's
    calculators (aks_estimation, aca_estimation)."""
    sheet.cell(row=1, column=1, value="Slide 3 minimums (yellow cells are inputs)").font = Font(bold=True, size=12)
    inputs = [("AKS nodes", deck.MIN_AKS_NODES), ("ACA vCPU per replica", deck.ACA_VCPU),
              ("ACA GiB per replica", deck.ACA_GIB), ("ACA requests (M)", deck.ACA_REQUESTS_MILLION),
              ("ACA hours on", deck.HOURS_MONTH)]
    for index, (label, value) in enumerate(inputs, start=3):
        sheet.cell(row=index, column=1, value=label)
        sheet.cell(row=index, column=2, value=value).fill = INPUT_FILL
    reserved = "1-price_node_reserved_term/26280/price_node_payg_hour"
    aks_total = calculator_block(sheet, 9, "AKS - SMALLEST CLUSTER - MONTHLY", [
        ("Cluster Management", 1, "=hours_month*price_aks_tier_hour", 1, "=bank_discount"),
        ("Nodes (3-year reservation)", "=B3", "=hours_month*price_node_payg_hour", 1, f"={reserved}"),
        ("OS Disks", "=B3", "=price_disk_month", 1, "=bank_discount"),
    ], time_header="-")
    aca_total = calculator_block(sheet, aks_total + 3, "ACA - ONE \"MICRO\"-SERVICE, 1 REPLICA ON - MONTHLY", [
        ("Requests (M)", "=B6", "=price_requests_million", 1, "=bank_discount"),
        ("vCPU", "=B4", "=price_vcpu_active", "=B7*3600", "=bank_discount"),
        ("GiB", "=B5", "=price_gib_active", "=B7*3600", "=bank_discount"),
    ])
    free = calculator_block(sheet, aca_total + 3, "ACA - FREE GRANT - MONTHLY (per subscription, not subtracted)", [
        ("Requests (M)", 2, "=price_requests_million", 1, "=bank_discount"),
        ("vCPU-s", 180000, "=price_vcpu_active", 1, "=bank_discount"),
        ("GiB-s", 360000, "=price_gib_active", 1, "=bank_discount"),
    ])
    sheet.cell(row=free + 2, column=1, value=(
        "Nodes: the reservation is entered as a discount on the pay-as-you-go price (1 - reserved term / 26,280 h / "
        "hourly price), as in the calculator. The bank's discount applies to the rest."))
    for column, width in {"A": 30, "B": 14, "C": 14, "D": 14, "E": 14, "F": 12, "G": 18, "H": 14, "I": 16}.items():
        sheet.column_dimensions[column].width = width


# Groups of the result slide: (result_group, label on the slide, runtime after the move).
RESULT_GROUPS = [("AKS", "AKS", "AKS"), ("ACA_WITH_REVIEW", "ACA con revisión", "ACA"), ("ACA", "ACA", "ACA"),
                 ("IN_USE", "Sin medición", "AKS"), ("NOT_IN_USE", "Sin uso", "AKS")]


def write_results(sheet):
    """The result slide: one row per group with the conditions it fails, its cost today and the new one, and a bar
    chart of the applications per group that can be edited in Excel (labels in Spanish, as on the slide)."""
    headers = ["result_group", "label", "apps", "fails_c1_peak_hour_tps", "fails_c2_microservices", "runtime_after",
               "today_usd", "new_usd", "difference_usd", "today_pen_igv", "new_pen_igv", "difference_pen_igv"]
    sheet.cell(row=1, column=1, value="Result per group (formulas over sheet Criteria; the chart reads columns B and C)"
               ).font = Font(bold=True, size=12)
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=column, value=header)
        cell.font, cell.fill = WHITE_BOLD, HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        sheet.column_dimensions[get_column_letter(column)].width = 16
    pen = "*exchange_rate*(1+igv)"
    # Applications in development stay out of the chart and the total, as on the slide (one row under the total).
    rows = RESULT_GROUPS + [None, ("IN_DEVELOPMENT", "En desarrollo (fuera del gráfico y del total)", "AKS")]
    for row, group_row in enumerate(rows, start=4):
        if group_row is None:
            continue
        group, label, runtime = group_row
        where = f'Apps[result_group],"{group}"'
        values = [group, label, f"=COUNTIFS({where})", f'=COUNTIFS({where},Apps[c1_peak_hour_tps_ok],"NO")',
                  f'=COUNTIFS({where},Apps[c2_microservices_ok],"NO")', runtime,
                  f"=SUMIFS(Apps[cost_month_usd],{where})", f"=SUMIFS(Apps[new_cost_usd],{where})", f"=G{row}-H{row}",
                  f"=G{row}{pen}", f"=H{row}{pen}", f"=I{row}{pen}"]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row, column=column, value=value)
            if column >= 7:
                cell.number_format = "#,##0"
    total = 4 + len(RESULT_GROUPS)
    sheet.cell(row=total, column=1, value="TOTAL").font = Font(bold=True)
    for column in (3, 7, 8, 9, 10, 11, 12):
        letter = get_column_letter(column)
        cell = sheet.cell(row=total, column=column, value=f"=SUM({letter}4:{letter}{total - 1})")
        cell.font = Font(bold=True)
        cell.number_format = "#,##0"
    sheet.cell(row=total + 3, column=1, value=(
        "new = Container Apps with the bank's discount for the applications that move (ACA, ACA_WITH_REVIEW), the "
        "same AKS cost for the rest. Change thresholds and prices in Assumptions, night replicas in OffhoursReplicas."))
    chart = BarChart()
    chart.type, chart.style, chart.legend = "col", 10, None
    chart.title = "Aplicaciones por destino"
    chart.add_data(Reference(sheet, min_col=3, min_row=3, max_row=3 + len(RESULT_GROUPS)), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=2, min_row=4, max_row=3 + len(RESULT_GROUPS)))
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = chart.dataLabels.showSerName = chart.dataLabels.showLegendKey = False
    chart.height, chart.width = 9, 16
    sheet.add_chart(chart, f"B{total + 5}")
    sheet.column_dimensions["B"].width = 30


def write_readme(sheet, results):
    lines = [
        ("AKS vs Container Apps: conditions and saving per application", True),
        ("", False),
        ("Where the data comes from (blue headers: values from the CSV files in " + os.path.abspath(results) + ")", True),
        ("Apps (sheet Criteria): the columns of aa-k8s-snapshot/application_snapshot.csv the conditions and the cost "
         "need.", False),
        ("Services: aa-k8s-snapshot/deployment_snapshot.csv joined with cluster_snapshot.csv on cluster_id, keeping "
         "the PROD clusters with applications; that join gives each \"micro\"-service its cod_app.", False),
        ("Prices: Container Apps and load balancer list prices (aa-prices/prices.csv, East US 2).", False),
        ("aks_aca_criteria.pq, next to this file, has the same three tables as Power Query queries (one step each); "
         "paste a query in Data > Get Data > Blank Query > Advanced Editor to see or rebuild it.", False),
        ("", False),
        ("How to use it", True),
        ("Green headers are formulas; blue headers come from the queries. Yellow cells are inputs.", False),
        ("Assumptions: thresholds, prices, discount and exchange rate. Each value has a name used in the formulas "
         "(Formulas > Name Manager).", False),
        ("OffhoursReplicas: replicas each \"micro\"-service of an application keeps on outside office hours. Empty "
         "= default_offhours_replicas (1). 0 = the application scales to zero at night (the first request waits up "
         "to ~30 s: check the caller's timeout).", False),
        ("Criteria: one row per application. c1 = peak hour traffic, c2 = \"micro\"-services in PROD, critical = BIA "
         "ALTA or MUY ALTA, growing = more than max_microservices with DESA. decision = AKS when c1 or c2 fail; "
         "ACA_WITH_REVIEW when it passes but is critical or growing; NO_TRAFFIC without measured PROD traffic.", False),
        ("Container Apps cost per application = office hours (measured replicas, sheet Services) + off-hours "
         "replicas + requests x requests_buffer + one internal load balancer per environment.", False),
        ("Results: the result slide (groups, conditions they fail, cost today against the new one) and its chart, "
         "which can be edited like any Excel chart.", False),
        ("Summary: the funnel, the decisions and the saving, as on the corporate deck.", False),
        ("", False),
        ("adt_runtime is the guideline's own result, kept for reference; the decision column is what the deck uses.",
         False),
    ]
    for row, (text, bold) in enumerate(lines, start=1):
        cell = sheet.cell(row=row, column=1, value=text)
        cell.font = Font(bold=bold, size=12 if row == 1 else 11)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    sheet.column_dimensions["A"].width = 140


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: aks_aca_criteria.xlsx in the newest reports/<date>-aks-aca folder")
    parser.add_argument("--aks-discount", type=float, default=None,
                        help="Discount on the list price of AKS node VMs, as a fraction (0.60 = 60%%). Default: measured")
    args = parser.parse_args()
    aa.azure.setup_logging(False)
    if not deck.results_csv(args.results, "aa-k8s-snapshot", "application_snapshot.csv"):
        raise SystemExit("No application_snapshot.csv in " + os.path.join(args.results, "aa-k8s-snapshot")
                         + ": run aa_k8s_snapshot.py first.")
    month = deck.read_estate(args.results)["month"]
    if args.aks_discount is not None:
        ratio = 1 - args.aks_discount
    else:
        ratio = deck.measured_discount(args.results, month, deck.read_prices(args.results))
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, "aks_aca_criteria.xlsx")
    apps, services = build(args.results, ratio, output)
    aa.azure.LOG.info("Wrote %s (%d applications, %d \"micro\"-services, AKS paid ratio %.2f)", output, apps, services,
                      ratio)


if __name__ == "__main__":
    main()
