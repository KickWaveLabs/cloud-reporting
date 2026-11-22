#!/usr/bin/env python3
"""Write the criteria workbook: which conditions each PROD application meets and what moving it to Container Apps
would save, from the CSV files in results/.

Seven sheets, plain Excel:
  Results     one row per destination (as the result slide) and a chart that can be edited
  Criteria    one row per application: its data (from the Data sheet, same row), the conditions, the decision, its
              off-hours replicas (1 by default, write 0 to let it scale to zero) and its monthly cost on AKS and ACA
  Calculator  one application's Container Apps and AKS cost, meter by meter (pick it in B2)
  Services    one row per "micro"-service (Services query), each one a container app in Container Apps: application,
              cluster, size, office replicas, HPA and measured use
  Prices      the list prices of the model, with the links where they come from
  Variables   thresholds, discount, IGV, exchange rate and the hours of the month
  Data        the Apps Power Query query: reads the CSV files and joins each "micro"-service to its application
Both Services and Data are Power Query connections to the CSV files.

Data > Refresh All reads the CSV files again (Folder, the first step of the query, is the --results folder; change it
on another machine). If a refresh brings more applications, drag the corner of the Criteria table down. The deck
(aa_deck_exec.py) uses the same conditions and cost model; with the defaults both give the same figures.
Reads only results/ (no Azure call). Output: aks_aca_criteria.xlsx in the newest reports/<date>-aks-aca folder (or
--output).
"""

import argparse
import glob
import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

import aa_common as aa
import aa_deck_exec as deck
import aa_powerquery as pq

REGION = "eastus2"
STYLE = "TableStyleMedium2"
BOLD = Font(bold=True)
USD, PEN, PCT = '"$"#,##0.00', '"S/ "#,##0.00', "0%"
PRICE_LINKS = [
    ("Azure Retail Prices API (where the prices are read from)",
     "https://prices.azure.com/api/retail/prices?$filter=serviceName eq 'Azure Container Apps' and "
     "armRegionName eq 'eastus2'"),
    ("Azure pricing calculator", "https://azure.microsoft.com/pricing/calculator/"),
    ("Container Apps pricing", "https://azure.microsoft.com/pricing/details/container-apps/"),
    ("Load Balancer pricing", "https://azure.microsoft.com/pricing/details/load-balancer/"),
]
# (name, key in prices.csv): one row of the Prices sheet each; the name is used in the formulas.
PRICES = [("price_vcpu", "aca_vcpu_active_s"), ("price_vcpu_idle", "aca_vcpu_idle_s"),
          ("price_gib", "aca_mem_active_gib_s"), ("price_gib_idle", "aca_mem_idle_gib_s"),
          ("price_requests", "aca_requests_million"), ("price_lb_hour", "lb_standard_rules_hour"),
          ("price_lb_gb", "lb_standard_data_gb")]

# Columns of the Apps query (sheet Criteria), in order; the last two come from the Services query.
APPS_COLUMNS = ["cod_app", "cluster_names", "clusters", "bia", "microservices", "microservices_projected", "traffic_status",
                "traffic_layer",
                "busiest_hour_tps", "busiest_offhours_hour_tps", "offhours_active_share", "requests_month",
                "cost_month_usd", "no_traffic_reason", "vcpu_office", "vcpu_one_replica_each"]
APPS_NUMBERS = ["clusters", "microservices", "microservices_projected", "busiest_hour_tps", "busiest_offhours_hour_tps",
                "offhours_active_share", "requests_month", "cost_month_usd", "vcpu_office", "vcpu_one_replica_each"]
PRICES_COLUMNS = ["name", "meter", "unit", "price"]
# Columns of the Services query (sheet Services): one row per "micro"-service, in order.
SERVICES_COLUMNS = ["cod_app", "cluster", "namespace", "deployment", "cpu_request_vcpu", "mem_request_gib",
                    "aca_vcpu", "aca_gib", "replicas_office", "hpa_min", "hpa_max", "replicas_max", "cpu_used_vcpu",
                    "mem_used_gib"]
# Services columns read as they are from deployment_snapshot.csv (CSV name -> sheet name), all numbers.
SERVICES_MEASURED = {"hpa_min": "hpa_min", "hpa_max": "hpa_max", "ci_replicas_max": "replicas_max",
                     "ci_cpu_used_per_replica": "cpu_used_vcpu", "ci_mem_used_gib_per_replica": "mem_used_gib"}


# ---------------------------------------------------------------------------
# Data: the rows the Power Query queries return
# ---------------------------------------------------------------------------


def number_or_none(value):
    return deck.to_float(value) if value not in ("", None) else None


def apps_rows(results):
    """application_snapshot.csv plus two sums per application from its "micro"-services (deployment_snapshot.csv,
    joined to the application through cluster_snapshot.csv): vCPU running in office hours (replicas x Container
    Apps size) and vCPU of one replica of each."""
    clusters = {row["cluster_id"]: row["cod_app"]
                for row in deck.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["category"] == aa.APPLICATION}
    sums = defaultdict(lambda: [0.0, 0.0])
    for row in deck.results_csv(results, "aa-k8s-snapshot", "deployment_snapshot.csv"):
        code = clusters.get(row["cluster_id"])
        if code and row["microservice"] == "YES":
            vcpu, _ = deck.aca_size(row)
            sums[code][0] += deck.office_replicas(row) * vcpu
            sums[code][1] += vcpu
    if not sums:
        aa.azure.LOG.warning("No micro-service sizes: deployment_snapshot.csv or cluster_snapshot.csv is missing or "
                             "empty in %s, so the Container Apps cost leaves out vCPU and memory",
                             os.path.join(results, "aa-k8s-snapshot"))
    out = []
    for row in sorted(deck.results_csv(results, "aa-k8s-snapshot", "application_snapshot.csv"),
                      key=lambda r: r["cod_app"]):
        row = dict(row, vcpu_office=sums[row["cod_app"]][0], vcpu_one_replica_each=sums[row["cod_app"]][1])
        out.append([number_or_none(row.get(c)) if c in APPS_NUMBERS else (row.get(c) or None) for c in APPS_COLUMNS])
    return out


def services_rows(results):
    """Each "micro"-service of a PROD cluster with applications: one container app in Container Apps, with its size and
    office replicas, the same figures its cost is made of."""
    clusters = {row["cluster_id"]: row["cod_app"]
                for row in deck.results_csv(results, "aa-k8s-snapshot", "cluster_snapshot.csv")
                if row["environment"] == "PROD" and row["category"] == aa.APPLICATION}
    out = []
    for row in deck.results_csv(results, "aa-k8s-snapshot", "deployment_snapshot.csv"):
        if row["microservice"] == "YES" and row["cluster_id"] in clusters:
            vcpu, gib = deck.aca_size(row)
            code = clusters[row["cluster_id"]]
            out.append([code, row["cluster"], row["namespace"], row["deployment"],
                        number_or_none(row["cpu_request_vcpu"]),
                        number_or_none(row["mem_request_gib"]), vcpu, gib, deck.office_replicas(row)]
                       + [number_or_none(row.get(source)) for source in SERVICES_MEASURED])
    return sorted(out, key=lambda values: (values[0], values[1], values[3]))


def prices_rows(results):
    found = {}
    for row in deck.results_csv(results, "aa-prices", "prices.csv"):
        if row["region"] == REGION and row["key"] not in found:
            found[row["key"]] = row
    return [[label, found[key]["meter"], found[key]["unit"], deck.to_float(found[key]["retail_price"])]
            for label, key in PRICES if key in found]


# ---------------------------------------------------------------------------
# Power Query: the Apps query, loaded into the Criteria sheet
# ---------------------------------------------------------------------------


def apps_query(results_folder):
    """M code of the Apps query: application_snapshot.csv, plus the vCPU of each application's "micro"-services
    (deployment_snapshot.csv joined to its application through cluster_snapshot.csv, sized as Container Apps)."""
    folder = results_folder.replace('"', '""')
    csv = '[Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]'

    def read(subfolder, file_name):
        return (f'Table.PromoteHeaders(Csv.Document(File.Contents(Folder & Separator & "{subfolder}" & Separator & '
                f'"{file_name}"), {csv}), [PromoteAllScalars = true])')
    columns = ", ".join(f'"{c}"' for c in APPS_COLUMNS[:-2])
    blanks = ", ".join(f'"{c}"' for c in APPS_NUMBERS[:-2])
    numbers = ", ".join(f'{{"{c}", type number}}' for c in APPS_NUMBERS)
    return f'''let
    // Folder with the CSV files (aks-aca/results): change it if they are somewhere else
    Folder = "{folder}",
    Separator = if Text.Contains(Folder, "\\") then "\\" else "/",
    ToNumber = (value) => if value = null or value = "" then 0 else Number.From(value, "en-US"),
    Applications = {read("aa-k8s-snapshot", "application_snapshot.csv")},
    Deployments = {read("aa-k8s-snapshot", "deployment_snapshot.csv")},
    Clusters = {read("aa-k8s-snapshot", "cluster_snapshot.csv")},
    // Each "micro"-service of a PROD cluster with applications gets its application through the cluster
    ApplicationClusters = Table.SelectColumns(Table.SelectRows(Clusters, each [environment] = "PROD" and [category] = "APPLICATION"), {{"cluster_id", "cod_app"}}),
    MicroServices = Table.RemoveColumns(Table.SelectRows(Deployments, each [microservice] = "YES"), {{"cod_app"}}, MissingField.Ignore),
    Services = Table.ExpandTableColumn(Table.NestedJoin(MicroServices, {{"cluster_id"}}, ApplicationClusters, {{"cluster_id"}}, "app", JoinKind.Inner), "app", {{"cod_app"}}),
    // Smallest Container Apps size that covers the request: steps of 0.25 vCPU, 2 GiB per vCPU
    AcaVcpu = Table.AddColumn(Services, "aca_vcpu", each Number.RoundUp(List.Max({{ToNumber([cpu_request_vcpu]), ToNumber([mem_request_gib]) / 2, 0.25}}) * 4) / 4, type number),
    // Replicas in office hours: Container Insights average, else the HPA minimum, at least 1
    Replicas = Table.AddColumn(AcaVcpu, "replicas_office", each List.Max({{1, if ToNumber([ci_replicas_avg_office]) > 0 then ToNumber([ci_replicas_avg_office]) else if ToNumber([hpa_min]) > 0 then ToNumber([hpa_min]) else 1}}), type number),
    PerApplication = Table.Group(Replicas, {{"cod_app"}}, {{{{"vcpu_office", each List.Sum(List.Transform(Table.ToRecords(_), (r) => r[replicas_office] * r[aca_vcpu])), type number}}, {{"vcpu_one_replica_each", each List.Sum([aca_vcpu]), type number}}}}),
    Columns = Table.SelectColumns(Applications, {{{columns}}}),
    Joined = Table.ExpandTableColumn(Table.NestedJoin(Columns, {{"cod_app"}}, PerApplication, {{"cod_app"}}, "sizes", JoinKind.LeftOuter), "sizes", {{"vcpu_office", "vcpu_one_replica_each"}}),
    Blanks = Table.ReplaceValue(Joined, "", null, Replacer.ReplaceValue, {{{blanks}}}),
    Typed = Table.TransformColumnTypes(Blanks, {{{numbers}}}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"cod_app", Order.Ascending}}}})
in
    Sorted'''


def services_query(results_folder):
    """M code of the Services query: the same steps as the Apps query up to each "micro"-service's size and replicas."""
    m = apps_query(results_folder)
    renames = ", ".join(f'{{"{source}", "{name}"}}' for source, name in SERVICES_MEASURED.items() if source != name)
    columns = ", ".join(f'"{c}"' for c in SERVICES_COLUMNS)
    numbers = ["cpu_request_vcpu", "mem_request_gib"] + list(SERVICES_MEASURED.values())
    blanks = ", ".join(f'"{c}"' for c in numbers)
    typed = ", ".join(f'{{"{c}", type number}}' for c in numbers)
    return m[:m.index("    PerApplication =")] + f'''    AcaGib = Table.AddColumn(Replicas, "aca_gib", each [aca_vcpu] * 2, type number),
    // Measured in Container Insights per replica (used) and the HPA configuration of the export
    Renamed = Table.RenameColumns(AcaGib, {{{renames}}}),
    Columns = Table.SelectColumns(Renamed, {{{columns}}}),
    Blanks = Table.ReplaceValue(Columns, "", null, Replacer.ReplaceValue, {{{blanks}}}),
    Typed = Table.TransformColumnTypes(Blanks, {{{typed}}}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"cod_app", Order.Ascending}}, {{"cluster", Order.Ascending}}, {{"deployment", Order.Ascending}}}})
in
    Sorted'''


# ---------------------------------------------------------------------------
# Sheets
# ---------------------------------------------------------------------------


def add_table(sheet, table_name, headers, rows, formulas=(), top=1):
    """An Excel table with value columns and formula columns (calculated columns, kept on refresh)."""
    all_headers = headers + [header for header, _ in formulas]
    for column, header in enumerate(all_headers, start=1):
        sheet.cell(row=top, column=column, value=header)
    for index, values in enumerate(rows, start=top + 1):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=index, column=column, value=value)
        for offset, (_, formula) in enumerate(formulas):
            sheet.cell(row=index, column=len(headers) + 1 + offset, value="=" + formula)
    last_row = top + max(len(rows), 1)
    table = Table(displayName=table_name, ref=f"A{top}:{get_column_letter(len(all_headers))}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name=STYLE, showRowStripes=True)
    table.tableColumns = [TableColumn(id=index, name=header) for index, header in enumerate(all_headers, start=1)]
    for offset, (_, formula) in enumerate(formulas):
        table.tableColumns[len(headers) + offset].calculatedColumnFormula = TableFormula(attr_text=formula)
    sheet.add_table(table)
    return last_row


def plain_table(sheet, table_name, ref):
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name=STYLE, showRowStripes=True)
    sheet.add_table(table)


def define(book, label, reference):
    book.defined_names[label] = DefinedName(label, attr_text=reference)


def write_variables(book, sheet, month):
    """Variables: one block per topic, a plain value and a short comment; each value has a name for the formulas."""
    blocks = [
        ("Conditions", [("max_tps", deck.MAX_TPS, "Peak hour traffic up to this TPS"),
                        ("max_services", deck.MAX_MICROSERVICES, "Up to this many \"micro\"-services in PROD")]),
        ("Container Apps", [("offhours_replicas", deck.OFFHOURS_REPLICAS,
                             "Replicas each \"micro\"-service keeps on outside office hours, weekends included (change it per app in Criteria)"),
                            ("offhours_rate", "ACTIVE", "ACTIVE or IDLE (idle: replica without requests)"),
                            ("requests_buffer", deck.REQUESTS_BUFFER, "Extra on the requests of the month"),
                            ("environments_per_app", deck.ENVIRONMENTS_PER_APP, "One load balancer each"),
                            ("gib_per_vcpu", 2, "Container Apps gives 2 GiB per vCPU")]),
        ("Money", [("bank_discount", deck.BANK_DISCOUNT, "Bank's discount on Azure list prices"),
                   ("igv", deck.IGV, "IGV"),
                   ("exchange_rate", deck.EXCHANGE_RATE, "Soles per US$")]),
        ("Time", [("data_month", month, "Month of the data"),
                  ("office_seconds", "=NETWORKDAYS(DATE(LEFT(data_month,4),RIGHT(data_month,2),1),EOMONTH(DATE(LEFT(data_month,4),"
                                     f"RIGHT(data_month,2),1),0))*{aa.OFFICE_END_HOUR - aa.OFFICE_START_HOUR}*3600",
                   f"Monday to Friday, {aa.OFFICE_START_HOUR}:00 to {aa.OFFICE_END_HOUR}:00"),
                  ("offhours_seconds", "=DAY(EOMONTH(DATE(LEFT(data_month,4),RIGHT(data_month,2),1),0))*86400-office_seconds",
                   "The rest of the month"),
                  ("hours_month", deck.HOURS_MONTH, "Hours billed for the load balancer")]),
    ]
    sheet["A1"] = "Variables"
    sheet["A1"].font = Font(bold=True, size=14)
    row = 3
    for title, items in blocks:
        sheet.cell(row=row, column=1, value=title).font = BOLD
        row += 1
        for label, value, comment in items:
            sheet.cell(row=row, column=1, value=label)
            cell = sheet.cell(row=row, column=2, value=value)
            if label in ("bank_discount", "igv"):
                cell.number_format = PCT
            elif label.endswith("seconds"):
                cell.number_format = "#,##0"
            sheet.cell(row=row, column=3, value=comment)
            define(book, label, f"Variables!$B${row}")
            row += 1
        row += 1
    for column, width in zip("ABC", (22, 12, 70)):
        sheet.column_dimensions[column].width = width


def write_prices(book, sheet, rows):
    last = add_table(sheet, "Prices", PRICES_COLUMNS, rows)
    for index, values in enumerate(rows, start=2):
        define(book, values[0], f"Prices!$D${index}")
        sheet.cell(row=index, column=4).number_format = "0.000000"
    row = last + 2
    sheet.cell(row=row, column=1, value="Source (list prices, East US 2):").font = BOLD
    for offset, (label, url) in enumerate(PRICE_LINKS, start=1):
        cell = sheet.cell(row=row + offset, column=1, value=label)
        cell.hyperlink, cell.style = url, "Hyperlink"
    for column, width in zip("ABCD", (18, 40, 16, 12)):
        sheet.column_dimensions[column].width = width


def criteria_formulas():
    r = lambda column: f"Criteria[[#This Row],[{column}]]"
    offhours_vcpu = 'IF(offhours_rate="IDLE",price_vcpu_idle,price_vcpu)'
    offhours_gib = 'IF(offhours_rate="IDLE",price_gib_idle,price_gib)'
    return [
        ("has_traffic", f'IF(AND({r("traffic_status")}="MEASURED",ISNUMBER({r("busiest_hour_tps")})),"YES","NO")'),
        ("tps_ok", f'IF({r("has_traffic")}="NO","-",IF({r("busiest_hour_tps")}<=max_tps,"YES","NO"))'),
        ("services_ok", f'IF(N({r("microservices")})<=max_services,"YES","NO")'),
        ("critical", f'IF(OR({r("bia")}="MUY ALTA",{r("bia")}="ALTA"),"YES","NO")'),
        ("growing", f'IF(AND(N({r("microservices_projected")})>max_services,N({r("microservices")})<=max_services),'
                    '"YES","NO")'),
        ("decision", f'IF({r("has_traffic")}="NO","NO_TRAFFIC",IF(OR({r("tps_ok")}="NO",{r("services_ok")}="NO"),'
                     f'"AKS",IF(OR({r("critical")}="YES",{r("growing")}="YES"),"ACA_WITH_REVIEW","ACA")))'),
        ("group", f'IF({r("decision")}<>"NO_TRAFFIC",{r("decision")},IF({r("no_traffic_reason")}="IN_DEVELOPMENT",'
                  f'"IN_DEVELOPMENT",IF({r("no_traffic_reason")}="IN_USE_NOT_MEASURED","IN_USE","NOT_IN_USE")))'),
        ("offhours_replicas", "offhours_replicas"),
        ("aca_list_usd", f"N({r('vcpu_office')})*office_seconds*(price_vcpu+gib_per_vcpu*price_gib)"
                         f"+{r('offhours_replicas')}*N({r('vcpu_one_replica_each')})*offhours_seconds*({offhours_vcpu}"
                         f"+gib_per_vcpu*{offhours_gib})+N({r('requests_month')})*requests_buffer/1000000*price_requests"
                         "+environments_per_app*hours_month*price_lb_hour"),
        ("aca_month_pen", f"{r('aca_list_usd')}*(1-bank_discount)*(1+igv)*exchange_rate"),
        ("aks_month_pen", f"N({r('cost_month_usd')})*(1+igv)*exchange_rate"),
        ("saving_month_pen", f'IF(OR({r("decision")}="ACA",{r("decision")}="ACA_WITH_REVIEW"),'
                             f"{r('aks_month_pen')}-{r('aca_month_pen')},0)"),
    ]


def write_data(sheet, rows):
    """The Apps query's table: the CSV data, refreshed from the files (Data > Refresh All)."""
    add_table(sheet, "Apps", APPS_COLUMNS, rows)
    for index, header in enumerate(APPS_COLUMNS, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(10, min(len(header) + 3, 22))


def write_services(sheet, rows):
    """The Services query's table: filter cod_app or cluster to see what each application or cluster runs."""
    add_table(sheet, "Services", SERVICES_COLUMNS, rows)
    for column, width in zip("ABCDEFGHIJKLMN", (10, 18, 22, 48, 16, 16, 10, 10, 16, 9, 9, 13, 14, 13)):
        sheet.column_dimensions[column].width = width


def write_criteria(sheet, rows):
    """One row per application: each data column comes from the same row of the Data sheet (the Apps query), then
    the conditions, the decision, the off-hours replicas (a value that can be changed) and the costs."""
    def pull(column):
        value = f"INDEX(Apps[{column}],ROW()-1)"
        return f'IF(ISBLANK({value}),"",{value})'
    formulas = [(column, pull(column)) for column in APPS_COLUMNS] + criteria_formulas()
    add_table(sheet, "Criteria", [], [[] for _ in rows], formulas)
    for index, (header, _) in enumerate(formulas, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(10, min(len(header) + 3, 22))
        number_format = (USD if header.endswith("_usd") else PEN if header.endswith("_pen")
                         else PCT if header == "offhours_active_share" else None)
        if number_format:
            for row in range(2, len(rows) + 2):
                sheet.cell(row=row, column=index).number_format = number_format
    # Off-hours replicas are a value, not a formula: 1 for every application, 0 where it may scale to zero outside
    # office hours.
    column = [header for header, _ in formulas].index("offhours_replicas") + 1
    for row in range(2, len(rows) + 2):
        sheet.cell(row=row, column=column, value=deck.OFFHOURS_REPLICAS)
    table = sheet.tables["Criteria"]
    table.tableColumns[column - 1].calculatedColumnFormula = None


# (group in Criteria, label, where it goes)
GROUPS = [("AKS", "Stays on AKS", "AKS"), ("ACA_WITH_REVIEW", "ACA, with review", "ACA"),
          ("ACA", "Candidate to ACA", "ACA"), ("IN_USE", "In use, not measured", "AKS"),
          ("NOT_IN_USE", "Not in use", "AKS")]


GROUP_COLORS = {"AKS": "1E3A78", "ACA_WITH_REVIEW": "F2A15A", "ACA": "4CB07A", "IN_USE": "4A8FB8",
                "NOT_IN_USE": "5B6770"}


def write_results(sheet):
    sheet["A1"] = "Result per destination (clusters; the decision is taken per application)"
    sheet["A1"].font = Font(bold=True, size=14)
    headers = ["Destination", "Clusters", "Apps", "Fail TPS", "Fail services", "Goes to", "AKS today (S/)",
               "With ACA (S/)", "Saving (S/)"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=3, column=column, value=header)
    for row, (group, label, runtime) in enumerate(GROUPS, start=4):
        where = f'Criteria[group],"{group}"'
        values = [label, f"=SUMIFS(Criteria[clusters],{where})", f"=COUNTIFS({where})",
                  f'=COUNTIFS({where},Criteria[tps_ok],"NO")', f'=COUNTIFS({where},Criteria[services_ok],"NO")', runtime,
                  f"=SUMIFS(Criteria[aks_month_pen],{where})", f"=G{row}-SUMIFS(Criteria[saving_month_pen],{where})",
                  f"=G{row}-H{row}"]
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value).number_format = PEN if column >= 7 else "General"
    total = 4 + len(GROUPS)
    sheet.cell(row=total, column=1, value="Total")
    for column in (2, 3, 7, 8, 9):
        letter = get_column_letter(column)
        cell = sheet.cell(row=total, column=column, value=f"=SUM({letter}4:{letter}{total - 1})")
        cell.number_format = PEN if column >= 7 else "General"
    plain_table(sheet, "Results", f"A3:I{total}")
    sheet.cell(row=total + 1, column=1,
               value='=COUNTIFS(Criteria[group],"IN_DEVELOPMENT")&" apps in development are not counted. Fail columns count '
                     'applications."')
    chart = BarChart()
    chart.type, chart.style, chart.title = "col", 10, "Clusters per destination"
    chart.legend = Legend()
    chart.legend.position = "b"
    chart.varyColors = True
    chart.add_data(Reference(sheet, min_col=2, min_row=3, max_row=3 + len(GROUPS)), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=4, max_row=3 + len(GROUPS)))
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = chart.dataLabels.showSerName = chart.dataLabels.showLegendKey = False
    # One colour per destination, as the deck, so the legend at the bottom names each bar.
    for index, (group, _, _) in enumerate(GROUPS):
        point = DataPoint(idx=index)
        point.graphicalProperties.solidFill = GROUP_COLORS[group]
        chart.series[0].dPt.append(point)
    chart.height, chart.width = 9, 16
    sheet.add_chart(chart, f"A{total + 3}")
    for column, width in zip("ABCDEFGHI", (22, 9, 7, 10, 13, 10, 16, 16, 16)):
        sheet.column_dimensions[column].width = width


def write_calculator(sheet, default_app):
    """One application, meter by meter, as the team's calculator: units x cost per unit x time = gross; less the
    discount; plus IGV; in soles."""
    app = lambda column: f"INDEX(Criteria[{column}],MATCH($B$2,Criteria[cod_app],0))"
    sheet["A2"], sheet["B2"], sheet["C2"] = "Application", default_app, "(choose it from the list)"
    sheet["A2"].font = sheet["B2"].font = BOLD
    validation = DataValidation(type="list", formula1='=INDIRECT("Criteria[cod_app]")', allow_blank=False)
    sheet.add_data_validation(validation)
    validation.add("B2")
    sheet["A3"], sheet["B3"] = "Decision", f"={app('decision')}"
    sheet["A4"], sheet["B4"] = "Off-hours replicas", f"={app('offhours_replicas')}"
    sheet["A5"], sheet["B5"] = "\"Micro\"-services", "=COUNTIFS(Services[cod_app],$B$2)"
    sheet["C5"] = "(the list: sheet Services, filter cod_app)"
    headers = ["Meter", "Units", "Cost / Unit", "Seconds", "Gross Total", "Discount", "Gross Total + Dsct",
               "Net Total", "Net Total (S/)"]

    def block(top, title, table_name, lines, time_header="Seconds"):
        sheet.cell(row=top, column=4, value=title).font = BOLD
        for column, header in enumerate(headers, start=1):
            sheet.cell(row=top + 1, column=column, value=time_header if header == "Seconds" else header)
        row = top + 2
        for meter, units, cost, time, discount in lines:
            values = [meter, units, cost, time, f"=B{row}*C{row}*D{row}", discount, f"=E{row}*(1-F{row})",
                      f"=G{row}*(1+igv)", f"=H{row}*exchange_rate"]
            for column, value in enumerate(values, start=1):
                cell = sheet.cell(row=row, column=column, value=value)
                cell.number_format = {3: "0.000000", 5: USD, 6: PCT, 7: USD, 8: USD, 9: PEN}.get(column, "General")
            row += 1
        sheet.cell(row=row, column=1, value="Total").font = BOLD
        for column, number_format in ((5, USD), (7, USD), (8, USD), (9, PEN)):
            letter = get_column_letter(column)
            cell = sheet.cell(row=row, column=column, value=f"=SUM({letter}{top + 2}:{letter}{row - 1})")
            cell.font, cell.number_format = BOLD, number_format
        plain_table(sheet, table_name, f"A{top + 1}:I{row}")
        return row

    sheet.cell(row=6, column=4, value="MONTH").font = BOLD
    month = [("days (12h)", "=office_seconds/3600/12"), ("hours", "=office_seconds/3600"),
             ("seconds", "=office_seconds"), ("off-hours seconds", "=offhours_seconds"),
             ("Requests(M)", f"=N({app('requests_month')})*requests_buffer/1000000")]
    for column, (header, formula) in enumerate(month, start=1):
        sheet.cell(row=7, column=column, value=header)
        sheet.cell(row=8, column=column, value=formula).number_format = "#,##0.00"
    plain_table(sheet, "Month", "A7:E8")
    offhours = lambda active, idle: f'=IF(offhours_rate="IDLE",{idle},{active})'
    aca = block(10, "ACA - COST - MONTHLY", "AcaCost", [
        ("Requests (M)", "=E8", "=price_requests", 1, "=bank_discount"),
        ("vCPU office", f"={app('vcpu_office')}", "=price_vcpu", "=C8", "=bank_discount"),
        ("GiB office", f"={app('vcpu_office')}*gib_per_vcpu", "=price_gib", "=C8", "=bank_discount"),
        ("vCPU off-hours", f"=B4*{app('vcpu_one_replica_each')}", offhours("price_vcpu", "price_vcpu_idle"), "=D8",
         "=bank_discount"),
        ("GiB off-hours", f"=B4*{app('vcpu_one_replica_each')}*gib_per_vcpu", offhours("price_gib", "price_gib_idle"),
         "=D8", "=bank_discount"),
    ])
    lb = block(aca + 3, "LOAD BALANCER - COST - MONTHLY", "LbCost", [
        ("Load Balancer rules", "=environments_per_app", "=price_lb_hour", "=hours_month", "=bank_discount"),
        ("Data processed", 0, "=price_lb_gb", 1, "=bank_discount"),
    ], time_header="Hours")
    aks = block(lb + 3, "AKS - COST - MONTHLY (as billed)", "AksCost", [
        ("AKS clusters", 1, f"=N({app('cost_month_usd')})", 1, 0),
    ], time_header="-")
    top = aks + 3
    for offset, (label, formula) in enumerate((("AKS today", f"=I{aks}"), ("ACA", f"=I{aca}+I{lb}"),
                                               ("Saving per month", f"=B{top}-B{top + 1}"),
                                               ("Saving per year", f"=B{top + 2}*12"))):
        sheet.cell(row=top + offset, column=1, value=label).font = BOLD
        sheet.cell(row=top + offset, column=2, value=formula).number_format = PEN
    for column, width in zip("ABCDEFGHI", (20, 14, 12, 14, 13, 10, 18, 12, 15)):
        sheet.column_dimensions[column].width = width


def build(results, output):
    month = deck.read_estate(results)["month"]
    apps = deck.read_applications(results)
    book = Workbook()
    results_sheet = book.active
    results_sheet.title = "Results"
    criteria, calculator, services, prices, variables, data = (
        book.create_sheet(title) for title in ("Criteria", "Calculator", "Services", "Prices", "Variables", "Data"))
    write_variables(book, variables, month)
    write_prices(book, prices, prices_rows(results))
    rows = apps_rows(results)
    write_data(data, rows)
    write_services(services, services_rows(results))
    write_criteria(criteria, rows)
    candidates = sorted(row["cod_app"] for row in apps if deck.decision(row) in deck.MOVABLE)
    write_calculator(calculator, candidates[0] if candidates else rows[0][0])
    write_results(results_sheet)
    book.calculation.fullCalcOnLoad = True
    book.save(output)
    folder = os.path.abspath(results)
    pq.load_queries(output, [("Data", "Apps", "Apps", apps_query(folder), APPS_COLUMNS),
                             ("Services", "Services", "Services", services_query(folder), SERVICES_COLUMNS)])
    return len(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: aks_aca_criteria.xlsx in the newest reports/<date>-aks-aca folder")
    args = parser.parse_args()
    aa.azure.setup_logging(False)
    if not deck.results_csv(args.results, "aa-k8s-snapshot", "application_snapshot.csv"):
        raise SystemExit("No application_snapshot.csv in " + os.path.join(args.results, "aa-k8s-snapshot")
                         + ": run aa_k8s_snapshot.py first.")
    reports = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
    output = args.output or os.path.join(reports[-1] if reports else aa.TOPIC_DIR, "aks_aca_criteria.xlsx")
    count = build(args.results, output)
    aa.azure.LOG.info("Wrote %s (%d applications; the Data sheet is the Apps query)", output, count)


if __name__ == "__main__":
    main()
