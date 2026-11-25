#!/usr/bin/env python3
"""Write the criteria workbook v2: one row per PROD application cluster, a summary per application, the traffic of
each hour and a calculator where the replicas of every hour can be chosen.

Same rules and cost model as aa_v2_model.py, written as Excel formulas over the same inputs (the module's docstring
explains them); with the default variables both give the same figures. Sheets:
  Results     clusters, AKS today, with Container Apps and saving per destination, and a chart
  Criteria    one row per cluster: data, criteria, decision, Container Apps cost from traffic, saving, group
  Apps        one row per application: its clusters, how many move, cost today and after
  Calculator  pick a cluster or an application (B2): traffic by hour (min / average / peak, weekday and weekend, as
              charts), replicas per "micro"-service each hour (today, simulated from traffic, chosen: editable) and
              the monthly cost meter by meter on AKS and on Container Apps
  Latency     the UX -> BS -> Core -> mainframe chain: how much a request takes with each layer on AKS or Container
              Apps; every parameter can be changed (to validate with the architecture team)
  Traffic     scope (cluster or application) x weekday/weekend x hour: TPS and simulated replicas
  Services    one row per "micro"-service: request, Container Apps size and its cost per second
  Prices      list prices with their links
  Variables   thresholds, saving margin, sizing rule, minimum replicas, discount, IGV, exchange rate
  Lists       the scopes the calculator offers
Plain values and formulas (no Power Query): run the script again to refresh. Reads only results/. Output:
aks_aca_criteria_v2.xlsx in the newest reports/<date>-aks-aca folder (or --output).
"""

import argparse
import glob
import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import aa_common as aa
import aa_criteria_workbook as v1wb
import aa_deck_exec as deck
import aa_v2_model as model

BOLD = Font(bold=True)
EDITABLE = PatternFill("solid", fgColor="FFF2CC")
USD, PEN, PCT = '"$"#,##0.00', '"S/ "#,##0', "0%"

CLUSTER_COLUMNS = ["cluster", "cod_app", "clusters_of_app", "bia", "microservices", "microservices_app",
                   "microservices_app_projected", "microservices_over_aca_size", "traffic_status", "traffic_layer",
                   "requests_month", "busiest_hour_tps", "peak_tps", "event_driven", "cost_month_usd",
                   "no_traffic_reason"]
CLUSTER_NUMBERS = {"clusters_of_app", "microservices", "microservices_app", "microservices_app_projected",
                   "microservices_over_aca_size", "requests_month", "busiest_hour_tps", "peak_tps", "cost_month_usd"}
TRAFFIC_COLUMNS = ["scope", "scope_type", "day_type", "hour", "office", "days", "tps_min", "tps_avg", "tps_max_hour",
                   "tps_peak", "replicas_today_avg", "replicas_today_max"]
SERVICE_COLUMNS = ["cluster", "cod_app", "namespace", "deployment", "cpu_request_vcpu", "mem_request_gib",
                   "replicas_today_office"]
GROUPS = [("AKS", "Stays on AKS (criteria)"), ("BELOW_MARGIN", "Meets criteria, saves less than the margin"),
          ("ACA_WITH_REVIEW", "ACA, with review"), ("ACA", "Candidate to ACA"),
          ("IN_USE", "In use, not measured"), ("NOT_IN_USE", "Not in use"), ("IN_DEVELOPMENT", "In development")]
LAYERS = ["UX", "BS", "CORE", "MAINFRAME"]
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


def write_variables(book, sheet, month, days):
    blocks = [
        ("Criteria", [("max_tps", model.DEFAULTS["max_tps"], "Busiest hour up to this TPS (average of the hour)"),
                      ("max_services", model.DEFAULTS["max_services"], "Up to this many \"micro\"-services per cluster")]),
        ("Saving filter (applied at the end, not a criterion)",
         [("min_saving", model.DEFAULTS["min_saving"], "A cluster that meets the criteria moves only if it saves at "
                                                        "least this share of its AKS cost")]),
        ("Replicas from traffic", [
            ("sizing_basis", model.DEFAULTS["sizing_basis"], "PEAK: busiest 5 minutes of each hour; AVERAGE: the "
                                                             "hour's average"),
            ("concurrent_requests", model.DEFAULTS["concurrent_requests"],
             "Requests per second per replica before Container Apps adds one (HTTP scale rule; Azure default 10)"),
            ("traffic_share", model.DEFAULTS["traffic_share"], "Share of the cluster's TPS each \"micro\"-service "
                                                               "receives (1 = all of it, conservative)"),
            ("min_replicas_office", model.DEFAULTS["min_replicas_office"], "Minimum replicas per \"micro\"-service, "
                                                                           "weekdays 8:00-20:00 (1 avoids the cold "
                                                                           "start)"),
            ("min_replicas_offhours", model.DEFAULTS["min_replicas_offhours"], "Minimum replicas per "
                                                                               "\"micro\"-service, the rest")]),
        ("Container Apps", [("requests_buffer", model.DEFAULTS["requests_buffer"], "Extra on the requests of the month"),
                            ("environments_per_cluster", model.DEFAULTS["environments_per_cluster"],
                             "Environments (one load balancer each) per cluster moved"),
                            ("hours_month", model.DEFAULTS["hours_month"], "Hours billed for the load balancer"),
                            ("min_vcpu", model.DEFAULTS["min_vcpu"], "Smallest replica in vCPU"),
                            ("vcpu_step", model.DEFAULTS["vcpu_step"], "Sizes go up in steps of this many vCPU"),
                            ("gib_per_vcpu", model.DEFAULTS["gib_per_vcpu"], "GiB per vCPU")]),
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
            if label in ("bank_discount", "igv", "min_saving", "traffic_share"):
                cell.number_format = PCT
            sheet.cell(row=row, column=3, value=comment)
            if label in SOURCES:
                link = sheet.cell(row=row, column=4, value="Microsoft Learn")
                link.hyperlink, link.style = SOURCES[label], "Hyperlink"
            v1wb.define(book, label, f"Variables!$B${row}")
            if label == "sizing_basis":
                validation = DataValidation(type="list", formula1='"PEAK,AVERAGE"', allow_blank=False)
                sheet.add_data_validation(validation)
                validation.add(cell.coordinate)
            row += 1
        row += 1
    for column, width in zip("ABCD", (26, 12, 90, 16)):
        sheet.column_dimensions[column].width = width


def r(table, column):
    return f"{table}[[#This Row],[{column}]]"


def write_traffic(sheet, profile):
    rows = [[row[c] if c in ("scope", "scope_type", "day_type") else number(row[c]) for c in TRAFFIC_COLUMNS]
            for row in profile]
    t = lambda column: r("Traffic", column)
    formulas = [("sim_replicas", f'MAX(IF({t("office")}=1,min_replicas_office,min_replicas_offhours),'
                                 f'CEILING(IF(sizing_basis="PEAK",{t("tps_peak")},{t("tps_avg")})*traffic_share/'
                                 'concurrent_requests,1))'),
                ("replica_seconds", f'{t("days")}*3600*{t("sim_replicas")}')]
    v1wb.add_table(sheet, "Traffic", TRAFFIC_COLUMNS, rows, formulas)
    for index, header in enumerate(TRAFFIC_COLUMNS + [f for f, _ in formulas], start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(9, min(len(header) + 2, 18))


def write_services(sheet, services, names):
    rows = []
    for row in sorted(services, key=lambda item: (names[item["cluster_id"]], item["deployment"])):
        rows.append([names[row["cluster_id"]], row["cod_app"], row["namespace"], row["deployment"],
                     number(row["cpu_request_vcpu"]), number(row["mem_request_gib"]),
                     number(row["ci_replicas_avg_office"])])
    s = lambda column: r("Services", column)
    larger = f'MAX({s("cpu_request_vcpu")},{s("mem_request_gib")}/gib_per_vcpu,min_vcpu)'
    formulas = [("size_vcpu", f"CEILING({larger},vcpu_step)"),
                ("usd_per_second", f'{s("size_vcpu")}*(price_vcpu+gib_per_vcpu*price_gib)')]
    v1wb.add_table(sheet, "Services", SERVICE_COLUMNS, rows, formulas)
    for column, width in zip("ABCDEFGHI", (18, 9, 14, 48, 15, 15, 18, 10, 15)):
        sheet.column_dimensions[column].width = width


def criteria_formulas():
    c = lambda column: r("Criteria", column)
    return [
        ("has_traffic", f'IF(AND({c("traffic_status")}="MEASURED",ISNUMBER({c("busiest_hour_tps")})),"YES","NO")'),
        ("size_ok", f'IF(N({c("microservices_over_aca_size")})=0,"YES","NO")'),
        ("tps_ok", f'IF({c("has_traffic")}="NO","-",IF({c("busiest_hour_tps")}<=max_tps,"YES","NO"))'),
        ("services_ok", f'IF(N({c("microservices")})<=max_services,"YES","NO")'),
        ("critical", f'IF(OR({c("bia")}="MUY ALTA",{c("bia")}="ALTA"),"YES","NO")'),
        ("growing", f'IF(AND(ISNUMBER({c("microservices_app_projected")}),N({c("microservices_app_projected")})>'
                    f'max_services,N({c("microservices_app")})<=max_services),"YES","NO")'),
        ("app_over_15", f'IF(N({c("microservices_app")})>max_services,"YES","NO")'),
        ("decision", f'IF({c("has_traffic")}="NO","NO_TRAFFIC",IF(OR({c("event_driven")}="YES",{c("size_ok")}="NO",'
                     f'{c("tps_ok")}="NO",{c("services_ok")}="NO"),"AKS",IF(OR({c("critical")}="YES",'
                     f'{c("growing")}="YES",{c("app_over_15")}="YES"),"ACA_WITH_REVIEW","ACA")))'),
        ("replica_seconds", f'SUMIFS(Traffic[replica_seconds],Traffic[scope],{c("cluster")},Traffic[scope_type],'
                            '"CLUSTER")'),
        ("usd_per_second", f'SUMIFS(Services[usd_per_second],Services[cluster],{c("cluster")})'),
        ("aca_list_usd", f'{c("replica_seconds")}*{c("usd_per_second")}+IF({c("event_driven")}="YES",0,'
                         f'N({c("requests_month")}))*requests_buffer/1000000*price_requests'
                         '+environments_per_cluster*hours_month*price_lb_hour'),
        ("aca_month_pen", f'{c("aca_list_usd")}*(1-bank_discount)*(1+igv)*exchange_rate'),
        ("aks_month_pen", f'N({c("cost_month_usd")})*(1+igv)*exchange_rate'),
        ("saving_month_pen", f'{c("aks_month_pen")}-{c("aca_month_pen")}'),
        ("saving_share", f'IF({c("aks_month_pen")}>0,{c("saving_month_pen")}/{c("aks_month_pen")},0)'),
        ("group", f'IF({c("decision")}="NO_TRAFFIC",IF({c("no_traffic_reason")}="IN_DEVELOPMENT","IN_DEVELOPMENT",'
                  f'IF({c("no_traffic_reason")}="IN_USE_NOT_MEASURED","IN_USE","NOT_IN_USE")),IF(AND(OR('
                  f'{c("decision")}="ACA",{c("decision")}="ACA_WITH_REVIEW"),{c("saving_share")}<min_saving),'
                  f'"BELOW_MARGIN",{c("decision")}))'),
        ("goes_to_aca", f'IF(OR({c("group")}="ACA",{c("group")}="ACA_WITH_REVIEW"),"YES","NO")'),
        ("final_month_pen", f'IF({c("goes_to_aca")}="YES",{c("aca_month_pen")},{c("aks_month_pen")})'),
    ]


def write_criteria(sheet, clusters):
    rows = [[number(row[c]) if c in CLUSTER_NUMBERS else (row[c] or None) for c in CLUSTER_COLUMNS]
            for row in sorted(clusters, key=lambda item: (item["cod_app"], item["cluster"]))]
    formulas = criteria_formulas()
    v1wb.add_table(sheet, "Criteria", CLUSTER_COLUMNS, rows, formulas)
    headers = CLUSTER_COLUMNS + [header for header, _ in formulas]
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(9, min(len(header) + 2, 20))
        fmt = USD if header.endswith("_usd") else PEN if header.endswith("_pen") else PCT if header == "saving_share" \
            else None
        if fmt:
            for row in range(2, len(rows) + 2):
                sheet.cell(row=row, column=index).number_format = fmt
    sheet.freeze_panes = "C2"


def write_apps(sheet, codes):
    a = lambda column: r("Apps", column)
    formulas = [("clusters", f'COUNTIFS(Criteria[cod_app],{a("cod_app")})'),
                ("clusters_to_aca", f'COUNTIFS(Criteria[cod_app],{a("cod_app")},Criteria[goes_to_aca],"YES")'),
                ("status", f'IF({a("clusters_to_aca")}=0,"Stays on AKS",IF({a("clusters_to_aca")}={a("clusters")},'
                           '"All to ACA","Split"))'),
                ("aks_month_pen", f'SUMIFS(Criteria[aks_month_pen],Criteria[cod_app],{a("cod_app")})'),
                ("after_month_pen", f'SUMIFS(Criteria[final_month_pen],Criteria[cod_app],{a("cod_app")})'),
                ("saving_month_pen", f'{a("aks_month_pen")}-{a("after_month_pen")}')]
    v1wb.add_table(sheet, "Apps", ["cod_app"], [[code] for code in codes], formulas)
    for index, header in enumerate(["cod_app"] + [f for f, _ in formulas], start=1):
        sheet.column_dimensions[get_column_letter(index)].width = 16
        if header.endswith("_pen"):
            for row in range(2, len(codes) + 2):
                sheet.cell(row=row, column=index).number_format = PEN


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
    chart = BarChart()
    chart.type, chart.style, chart.title = "col", 10, "Clusters per destination"
    chart.legend = None
    chart.add_data(Reference(sheet, min_col=2, min_row=3, max_row=3 + len(GROUPS)), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=4, max_row=3 + len(GROUPS)))
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.height, chart.width = 8, 18
    sheet.add_chart(chart, f"A{total + 6}")
    for column, width in zip("ABCDE", (44, 10, 16, 16, 16)):
        sheet.column_dimensions[column].width = width


def write_lists(sheet, clusters):
    scopes = sorted(row["cluster"] for row in clusters) + sorted({row["cod_app"] for row in clusters})
    v1wb.add_table(sheet, "Scopes", ["scope"], [[scope] for scope in scopes])
    sheet.column_dimensions["A"].width = 20


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
             ("US$ per second, one replica of each", f'=IF({is_cluster},SUMIFS(Services[usd_per_second],'
                                                    'Services[cluster],$B$2),SUMIFS(Services[usd_per_second],'
                                                    'Services[cod_app],$B$2))'),
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
    headers = ["Hour", "Weekday TPS min", "TPS average", "TPS peak", "Replicas today", "Simulated", "Chosen",
               "Weekend TPS min", "TPS average", "TPS peak", "Replicas today", "Simulated", "Chosen"]
    sheet.cell(row=top - 1, column=1, value="Replicas per \"micro\"-service each hour (Chosen: yellow, change it; "
                                            "it starts as Simulated)").font = BOLD
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=top, column=column, value=header).font = BOLD
    fields = [("tps_min", "0.00"), ("tps_avg", "0.00"), ("tps_peak", "0.00"), ("replicas_today_avg", "0.0"),
              ("sim_replicas", "0")]
    for hour in range(24):
        row = top + 1 + hour
        sheet.cell(row=row, column=1, value=hour)
        for block, day_type in ((0, "weekday"), (6, "weekend")):
            for offset, (field, fmt) in enumerate(fields):
                cell = sheet.cell(row=row, column=2 + block + offset,
                                  value=f'=SUMIFS(Traffic[{field}],Traffic[scope],$B$2,Traffic[day_type],'
                                        f'"{day_type}",Traffic[hour],$A{row})')
                cell.number_format = fmt
            simulated = get_column_letter(6 + block)
            chosen = sheet.cell(row=row, column=7 + block, value=f"={simulated}{row}")
            chosen.fill, chosen.number_format = EDITABLE, "0"
    last = top + 24
    cost_top = last + 2
    sheet.cell(row=cost_top, column=1, value="Monthly cost (US$ list price unless said)").font = BOLD
    sheet.cell(row=cost_top, column=2, value="Simulated").font = BOLD
    sheet.cell(row=cost_top, column=3, value="Chosen").font = BOLD
    seconds = lambda wd, we: f"(SUM({wd}{top + 1}:{wd}{last})*weekdays+SUM({we}{top + 1}:{we}{last})*weekend_days)*3600"
    meters = [("Replica-seconds per \"micro\"-service", seconds("F", "L"), seconds("G", "M"), "#,##0"),
              ("Compute (vCPU + memory)", f"B{cost_top + 1}*$B$7", f"C{cost_top + 1}*$B$7", USD),
              ("Requests", "N($B$8)*requests_buffer/1000000*price_requests",
               "N($B$8)*requests_buffer/1000000*price_requests", USD),
              ("Environment (load balancer)", "$B$6*environments_per_cluster*hours_month*price_lb_hour",
               "$B$6*environments_per_cluster*hours_month*price_lb_hour", USD),
              ("Total list price", f"SUM(B{cost_top + 2}:B{cost_top + 4})", f"SUM(C{cost_top + 2}:C{cost_top + 4})",
               USD),
              ("Container Apps with discount and IGV (S/)", f"B{cost_top + 5}*(1-bank_discount)*(1+igv)*exchange_rate",
               f"C{cost_top + 5}*(1-bank_discount)*(1+igv)*exchange_rate", PEN),
              ("AKS today (S/)", "$B$9", "$B$9", PEN),
              ("Saving (S/)", f"B{cost_top + 7}-B{cost_top + 6}", f"C{cost_top + 7}-C{cost_top + 6}", PEN),
              ("Saving share", f"IF(B{cost_top + 7}>0,B{cost_top + 8}/B{cost_top + 7},0)",
               f"IF(C{cost_top + 7}>0,C{cost_top + 8}/C{cost_top + 7},0)", PCT)]
    for offset, (label, simulated, chosen, fmt) in enumerate(meters, start=1):
        sheet.cell(row=cost_top + offset, column=1, value=label)
        for column, formula in ((2, simulated), (3, chosen)):
            sheet.cell(row=cost_top + offset, column=column, value="=" + formula).number_format = fmt
    for block, title, anchor in ((0, "Weekday TPS by hour", "O2"), (6, "Weekend TPS by hour", "O20")):
        chart = LineChart()
        chart.title, chart.height, chart.width = title, 8, 16
        chart.y_axis.title, chart.x_axis.title = "TPS", "Hour"
        chart.add_data(Reference(sheet, min_col=2 + block, max_col=4 + block, min_row=top, max_row=last),
                       titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=top + 1, max_row=last))
        chart.legend = Legend()
        chart.legend.position = "b"
        sheet.add_chart(chart, anchor)
    replicas_chart = BarChart()
    replicas_chart.type, replicas_chart.title, replicas_chart.height, replicas_chart.width = \
        "col", "Chosen replicas per \"micro\"-service", 8, 16
    for column in (7, 13):
        replicas_chart.add_data(Reference(sheet, min_col=column, min_row=top, max_row=last), titles_from_data=True)
    replicas_chart.set_categories(Reference(sheet, min_col=1, min_row=top + 1, max_row=last))
    sheet.add_chart(replicas_chart, "O38")
    sheet.column_dimensions["A"].width = 40
    for column in "BCDEFGHIJKLM":
        sheet.column_dimensions[column].width = 13


def layer_latency(results):
    """Request-weighted average backend latency per APIM layer (ms), from aa_signals.py, for reference."""
    latency = aa.read_csv(os.path.join(results, "aa-signals", "apim_latency_monthly.csv"))
    requests = {(row["apim"], row["api_id"], row["month"]): model.to_float(row["requests"]) or 0
                for row in aa.read_csv(os.path.join(results, "aa-metrics", "apim_api_monthly.csv"))}
    sums = defaultdict(lambda: [0.0, 0.0])
    for row in latency:
        weight = requests.get((row["apim"], row["api_id"], row["month"]), 0)
        value = model.to_float(row.get("backend_duration_avg_ms"))
        if weight and value is not None:
            sums[row["apim_layer"]][0] += value * weight
            sums[row["apim_layer"]][1] += weight
    return {layer: round(total / weight, 1) for layer, (total, weight) in sums.items() if weight}


def write_latency(book, sheet, measured):
    """The chain UX -> BS -> Core -> mainframe, built from the bottom: latency(layer) = own time + calls to the next
    layer x (hop + latency(next)), plus, on Container Apps, its ingress, a cold start when no replica is on and the
    scale-out delay of a burst."""
    sheet["A1"], sheet["A1"].font = "Latency of one request along the chain (ms) - parameters to validate", \
        Font(bold=True, size=14)
    params = [("hop_aks_ms", 2, "Network hop between services inside AKS"),
              ("aca_ingress_ms", 5, "Extra of the Container Apps ingress (Envoy) on each call into a layer on ACA"),
              ("cross_platform_ms", 5, "Extra when a call goes from AKS to Container Apps or back"),
              ("cold_start_s", 30, "First request when the layer has no replica on (scale to zero)"),
              ("cold_probability", 0, "Share of requests that find no replica on (0 with 1 minimum replica)"),
              ("scale_delay_s", 15, "Container Apps checks its HTTP rule every 15 s"),
              ("burst_probability", 0.001, "Share of requests that arrive while a burst waits for a new replica")]
    sheet.cell(row=3, column=1, value="Parameter").font = BOLD
    sheet.cell(row=3, column=2, value="Value").font = BOLD
    for offset, (label, value, comment) in enumerate(params, start=4):
        sheet.cell(row=offset, column=1, value=label)
        cell = sheet.cell(row=offset, column=2, value=value)
        cell.fill = EDITABLE
        sheet.cell(row=offset, column=3, value=comment)
        v1wb.define(book, label, f"Latency!$B${offset}")
    top = 13
    headers = ["Layer", "Own time (ms)", "Calls to next layer", "Measured APIM backend (ms)",
               "All AKS", "UX on ACA", "UX+BS on ACA", "UX+BS+Core on ACA"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(row=top, column=column, value=header).font = BOLD
    platforms = {"All AKS": [], "UX on ACA": ["UX"], "UX+BS on ACA": ["UX", "BS"], "UX+BS+Core on ACA": ["UX", "BS", "CORE"]}
    # Own time from the measured APIM backend latency: each layer's minus the one below (a layer's latency includes
    # the calls it makes); the mainframe takes half of Core's until it is measured. Fallbacks when a layer is missing.
    ux, bs, core = measured.get("UX", 170), measured.get("BS", 150), measured.get("CORE", 40)
    mainframe = round(core / 2, 1)
    defaults = {"UX": (max(round(ux - bs, 1), 1), 1), "BS": (max(round(bs - core, 1), 1), 1),
                "CORE": (max(round(core - mainframe, 1), 1), 1), "MAINFRAME": (mainframe, 0)}
    for offset, layer in enumerate(LAYERS, start=1):
        row = top + offset
        sheet.cell(row=row, column=1, value=layer)
        for column, value in ((2, defaults[layer][0]), (3, defaults[layer][1])):
            cell = sheet.cell(row=row, column=column, value=value)
            cell.fill = EDITABLE
        sheet.cell(row=row, column=4, value=measured.get(layer, "-"))
        for column, scenario in enumerate(platforms, start=5):
            if layer == "MAINFRAME":
                sheet.cell(row=row, column=column, value="ON-PREM")
            else:
                sheet.cell(row=row, column=column, value="ACA" if layer in platforms[scenario] else "AKS").fill = EDITABLE
    result = top + len(LAYERS) + 2
    sheet.cell(row=result, column=1, value="Latency from this layer down (ms)").font = BOLD
    for offset, layer in enumerate(reversed(LAYERS)):
        row = result + 1 + offset
        sheet.cell(row=row, column=1, value=layer)
        source = top + 1 + LAYERS.index(layer)
        for column in range(5, 9):
            letter = get_column_letter(column)
            here = f"{letter}{source}"
            extra = (f'IF({here}="ACA",aca_ingress_ms+cold_probability*cold_start_s*1000+burst_probability*'
                     'scale_delay_s*1000,0)')
            if layer == "MAINFRAME":
                formula = f"=$B{source}+{extra}"
            else:
                below_platform = f"{letter}{source + 1}"
                below_latency = f"{letter}{row - 1}"
                hop = f'(hop_aks_ms+IF(AND({below_platform}<>"ON-PREM",{here}<>{below_platform}),cross_platform_ms,0))'
                formula = f"=$B{source}+{extra}+$C{source}*({hop}+{below_latency})"
            sheet.cell(row=row, column=column, value=formula).number_format = "0"
    total = result + len(LAYERS)
    sheet.cell(row=total + 1, column=1, value="Extra over all AKS (ms)").font = BOLD
    for column in range(5, 9):
        letter = get_column_letter(column)
        sheet.cell(row=total + 1, column=column, value=f"={letter}{total}-$E${total}").number_format = "0"
    sheet.cell(row=total + 3, column=1, value=(
        "Own time: what the layer itself takes (measured APIM backend latency of the layer minus that of the layer "
        "below, approximately). Calls to next layer: sequential calls per request. Yellow cells can be changed."))
    for column, width in zip("ABCDEFGH", (34, 14, 18, 26, 12, 12, 14, 18)):
        sheet.column_dimensions[column].width = width


def build(results, output):
    data = model.load(results)
    clusters = data["clusters"]
    summary = aa.read_csv(os.path.join(results, "aa-cluster-snapshot", "traffic_hourly_profile.csv"))
    days = {"weekday": 0, "weekend": 0}
    for row in summary:
        days[row["day_type"]] = int(row["days"])
    month = clusters[0]["month"] if clusters else ""
    book = Workbook()
    sheets = {name: (book.active if index == 0 else book.create_sheet(name))
              for index, name in enumerate(["Results", "Criteria", "Apps", "Calculator", "Latency", "Traffic",
                                            "Services", "Prices", "Variables", "Lists"])}
    sheets["Results"].title = "Results"
    write_variables(book, sheets["Variables"], month, days)
    v1wb.write_prices(book, sheets["Prices"], v1wb.prices_rows(results))
    write_traffic(sheets["Traffic"], data["profile"])
    write_services(sheets["Services"], data["services"], data["names"])
    write_criteria(sheets["Criteria"], clusters)
    write_apps(sheets["Apps"], sorted({row["cod_app"] for row in clusters}))
    write_results(sheets["Results"])
    write_lists(sheets["Lists"], clusters)
    example = next((row["cluster"] for row in clusters if row["cod_app"] == deck.EXAMPLE_APP), clusters[0]["cluster"])
    write_calculator(sheets["Calculator"], example)
    write_latency(book, sheets["Latency"], layer_latency(results))
    book.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    book.save(output)
    aa.azure.LOG.info("Wrote %s (%d clusters)", output, len(clusters))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: aks_aca_criteria_v2.xlsx in the newest reports/<date>-aks-aca")
    args = parser.parse_args()
    aa.azure.setup_logging(False)
    output = args.output
    if not output:
        folders = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
        folder = folders[-1] if folders else os.path.join(aa.TOPIC_DIR, "reports", "aks-aca")
        output = os.path.join(folder, "aks_aca_criteria_v2.xlsx")
    build(args.results, output)


if __name__ == "__main__":
    main()
