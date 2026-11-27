#!/usr/bin/env python3
"""Build the backup cost calculator: monthly cost of each application's backup by plan, in Azure and at the vendors.

The user enters, per row, the service (VM, SQL Database or Cosmos DB), its data size in TB and the plan code the
decision tree gave it (P0, PN, P1 to P4). The workbook returns the TB stored for the short and long term and the
monthly cost in Azure, at D3 and at Opción 1. Region, discount, redundancy (LRS, ZRS or GRS) and the Archive tier are
chosen once at the top; the assumptions and prices sit in their own sheets and can be edited.

The model is the one in bu_guidelines_deck.py (plan_cost), which also feeds the deck. Prices come from the public
Azure Retail Prices API; no Azure sign-in and no estate data are needed.

Sheets:
  Calculator    inputs and monthly cost per row, with totals
  Plans         long-term copies kept by each plan in steady state and the share that can sit in Archive
  Assumptions   change rates and copy weights (editable)
  Prices        list price per TB-month for Central US and East US 2, and the vendors' quoted tariffs
  Check         the model in Python for 1 TB per service and plan, to compare with the formulas

Open the file in Excel and save it once so the formulas calculate.

Usage:
  python backup/src/bu_guidelines_calculator.py
Output: backup/reports/<date>-lineamientos/calculadora_backup.xlsx (or --output).
"""

import argparse
import os
from datetime import date

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import bu_guidelines_deck as deck

NAVY, WHITE, INPUT = "1E3A78", "FFFFFF", "FFF2CC"
HEADER = PatternFill("solid", fgColor=NAVY)
EDITABLE = PatternFill("solid", fgColor=INPUT)
REGIONS = {"Central US": "centralus", "East US 2": "eastus2"}
PLAN_CODES = ["P0", "PN", "P1", "P2", "P3", "P4"]
ROWS = 30  # input rows on the Calculator sheet
FIRST = 12  # first input row


def header(sheet, row, values, widths=None):
    for column, value in enumerate(values, start=1):
        cell = sheet.cell(row, column, value)
        cell.fill, cell.font = HEADER, Font(bold=True, color=WHITE)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for column, width in enumerate(widths or [], start=1):
        sheet.column_dimensions[chr(64 + column)].width = width


def price_rows(list_prices):
    """(key, description) of every price the model uses."""
    rows = [("vault_lrs", "Backup vault, LRS"), ("vault_zrs", "Backup vault, ZRS"), ("vault_grs", "Backup vault, GRS"),
            ("archive_lrs", "Archive tier (VM), LRS"), ("archive_grs", "Archive tier (VM), GRS"),
            ("sqldb_pitr_lrs", "SQL Database PITR, LRS"), ("sqldb_pitr_zrs", "SQL Database PITR, ZRS"),
            ("sqldb_pitr_geo", "SQL Database PITR, RA-GRS"), ("sqldb_ltr_lrs", "SQL Database LTR, LRS"),
            ("sqldb_ltr_zrs", "SQL Database LTR, ZRS"), ("sqldb_ltr_geo", "SQL Database LTR, RA-GRS"),
            ("vm_instance", "Azure Backup protected instance, VM (per TB protected)"),
            ("cosmos_instance", "Azure Backup protected instance, Cosmos DB (per TB protected)")]
    return [(key, text, list_prices["centralus"][key], list_prices["eastus2"][key]) for key, text in rows]


def build(output, list_prices):
    book = Workbook()
    calc = book.active
    calc.title = "Calculator"

    # Prices: list US$ per TB-month by region; column E is the net price for the region chosen in Calculator.
    prices = book.create_sheet("Prices")
    header(prices, 1, ["Key", "Price", "Central US (list, US$/TB-month)", "East US 2 (list, US$/TB-month)",
                       "Net, selected region"], [18, 52, 16, 16, 16])
    for row, (key, text, central, east) in enumerate(price_rows(list_prices), start=2):
        prices.append([key, text, round(central, 4), round(east, 4),
                       f'=IF(Calculator!$B$3="East US 2",D{row},C{row})*(1-Calculator!$B$4)'])
    last_price = prices.max_row
    vendor_row = last_price + 3
    prices.cell(vendor_row - 1, 1, "Vendors (quoted, US$/TB-month, net)").font = Font(bold=True)
    header(prices, vendor_row, ["Vendor", "Active (short term)", "Archive (long term)"])
    for offset, (name, active, archive) in enumerate(deck.VENDORS, start=1):
        prices.append([name, active, archive])
        for column in (2, 3):
            prices.cell(vendor_row + offset, column).fill = EDITABLE
    vendors = {name: vendor_row + offset for offset, (name, _, _) in enumerate(deck.VENDORS, start=1)}

    def price(key):
        return f'INDEX(Prices!$E$2:$E${last_price},MATCH("{key}",Prices!$A$2:$A${last_price},0))'

    def price_by(kind, redundancy_cell):
        keys = deck.REDUNDANCY_KEYS[kind]
        options = ",".join(f'IF({redundancy_cell}="{name}",{price(key)}' for name, key in keys.items())
        return options + ",NA()" + ")" * len(keys)

    # Plans: steady-state copies and Archive share.
    plans = book.create_sheet("Plans")
    header(plans, 1, ["Plan", "Weekly kept", "Monthly kept", "Yearly kept", "Long-term copies (steady state)",
                      "Share of copies in Archive (VM)"], [8, 12, 12, 12, 18, 18])
    for code in PLAN_CODES:
        policy = deck.PLAN_POLICIES.get(code, (0, 0, 0))
        copies, share = deck.steady_copies(policy) if code in deck.PLAN_POLICIES else (0, 0)
        plans.append([code, *policy, round(copies, 3), round(share, 3)])

    # Assumptions.
    notes = {"vm_daily_change": "Share of a VM's data that changes per day (Microsoft: usually 2% to 7%)",
             "sql_pitr_factor": "Billable PITR TB per TB of database (PITR is free up to the database size)",
             "sql_ltr_weight": "Each SQL LTR copy as a share of the database (measured on the Oct 2026 bill)",
             "vm_ltr_weight": "Each VM long-term point in the vault as a share of the VM data (assumption)",
             "cosmos_ltr_weight": "Each Cosmos DB long-term copy after the first full one (assumption)"}
    assumptions = book.create_sheet("Assumptions")
    header(assumptions, 1, ["Key", "Value", "Meaning"], [20, 10, 80])
    cells = {}
    for row, (key, value) in enumerate(deck.ASSUMPTIONS.items(), start=2):
        assumptions.append([key, value, notes[key]])
        assumptions.cell(row, 2).fill = EDITABLE
        cells[key] = f"Assumptions!$B${row}"

    # Calculator: global inputs.
    calc["A1"] = "Backup cost calculator (US$ per month, net of discount)"
    calc["A1"].font = Font(bold=True, size=14, color=NAVY)
    inputs = [("Region", "Central US"), ("Discount on Azure list prices", deck.DISCOUNT),
              ("Short-term redundancy", deck.SHORT_REDUNDANCY.upper()),
              ("Long-term redundancy", deck.LONG_REDUNDANCY.upper()),
              ("VM long-term copies to Archive (LRS or GRS only)", "Yes")]
    for row, (label, value) in enumerate(inputs, start=3):
        calc.cell(row, 1, label)
        calc.cell(row, 2, value).fill = EDITABLE
    calc["B4"].number_format = "0%"
    for formula, target in (('"Central US,East US 2"', "B3"), ('"LRS,ZRS,GRS"', "B5"), ('"LRS,ZRS,GRS"', "B6"),
                            ('"Yes,No"', "B7")):
        rule = DataValidation(type="list", formula1=formula, allow_blank=False)
        calc.add_data_validation(rule)
        rule.add(target)
    calc["A9"] = ("One row per application and service. Plan: the code the decision tree gives (P0 no backup, PN "
                  "non-production, P1 base, P2 base + 10 years, P3 operations register, P4 LBTR).")
    calc["A9"].alignment = Alignment(wrap_text=True)
    calc.merge_cells("A9:J9")
    calc.row_dimensions[9].height = 30

    columns = ["Application", "Service", "Data size (TB)", "Plan", "Short-term TB stored", "Long-term copies",
               "Long-term TB stored", "Azure (US$/month)", *[f"{name} (US$/month)" for name, _, _ in deck.VENDORS]]
    header(calc, FIRST - 1, columns, [46, 16, 14, 8, 14, 12, 14, 14, 14, 14])
    examples = [("Example VM", "VM", 1, "P1"), ("Example SQL Database", "SQL Database", 1, "P1"),
                ("Example Cosmos DB", "Cosmos DB", 1, "P1")]
    service_rule = DataValidation(type="list", formula1='"' + ",".join(deck.SERVICES) + '"', allow_blank=True)
    plan_rule = DataValidation(type="list", formula1='"' + ",".join(PLAN_CODES) + '"', allow_blank=True)
    calc.add_data_validation(service_rule)
    calc.add_data_validation(plan_rule)
    stored = 'OR(D{r}="P1",D{r}="P2",D{r}="P3",D{r}="P4")'
    for index in range(ROWS):
        r = FIRST + index
        if index < len(examples):
            for column, value in enumerate(examples[index], start=1):
                calc.cell(r, column, value)
        for column in range(1, 5):
            calc.cell(r, column).fill = EDITABLE
        service_rule.add(f"B{r}")
        plan_rule.add(f"D{r}")
        on = stored.format(r=r)
        copies = f"IF({on},VLOOKUP(D{r},Plans!$A$2:$F$7,5,FALSE),0)"
        share = f"VLOOKUP(D{r},Plans!$A$2:$F$7,6,FALSE)"
        calc[f"E{r}"] = (f'=IF(C{r}="","",IF(NOT({on}),0,IF(B{r}="VM",C{r}*(1+6*{cells["vm_daily_change"]}),'
                         f'IF(B{r}="SQL Database",C{r}*{cells["sql_pitr_factor"]},0))))')
        calc[f"F{r}"] = f'=IF(C{r}="","",{copies})'
        calc[f"G{r}"] = (f'=IF(C{r}="","",IF(F{r}=0,0,IF(B{r}="VM",F{r}*C{r}*{cells["vm_ltr_weight"]},'
                         f'IF(B{r}="SQL Database",F{r}*C{r}*{cells["sql_ltr_weight"]},'
                         f'C{r}*(1+(F{r}-1)*{cells["cosmos_ltr_weight"]})))))')
        vm_long = (f'IF(AND($B$7="Yes",$B$6<>"ZRS"),{share}*{price_by("archive", "$B$6")}'
                   f'+(1-{share})*{price_by("vault", "$B$6")},{price_by("vault", "$B$6")})')
        vm = (f'E{r}*{price_by("vault", "$B$5")}+IF({on},C{r}*{price("vm_instance")},0)+G{r}*{vm_long}')
        sql = f'E{r}*{price_by("sql_pitr", "$B$5")}+G{r}*{price_by("sql_ltr", "$B$6")}'
        cosmos = f'G{r}*{price_by("vault", "$B$6")}+IF(F{r}>0,C{r}*{price("cosmos_instance")},0)'
        calc[f"H{r}"] = f'=IF(C{r}="","",IF(B{r}="VM",{vm},IF(B{r}="SQL Database",{sql},{cosmos})))'
        for offset, (name, _, _) in enumerate(deck.VENDORS):
            column = chr(ord("I") + offset)
            calc[f"{column}{r}"] = (f'=IF(C{r}="","",E{r}*Prices!$B${vendors[name]}+G{r}*Prices!$C${vendors[name]})')
        for column in "EFGHIJ":
            calc[f"{column}{r}"].number_format = "#,##0.00"
    total = FIRST + ROWS
    calc.cell(total, 1, "Total").font = Font(bold=True)
    for column in "CEGHIJ":
        calc[f"{column}{total}"] = f"=SUM({column}{FIRST}:{column}{total - 1})"
        calc[f"{column}{total}"].font = Font(bold=True)
        calc[f"{column}{total}"].number_format = "#,##0.00"

    # Check: the Python model for 1 TB, Central US, deck redundancy, to compare with the formulas.
    check = book.create_sheet("Check")
    net = {key: value * (1 - deck.DISCOUNT) for key, value in list_prices["centralus"].items()}
    header(check, 1, ["Service", "Plan", "Azure (US$/month)", *[name for name, _, _ in deck.VENDORS]], [16, 8, 16, 14, 14])
    for service in deck.SERVICES:
        for plan in PLAN_CODES:
            result = deck.plan_cost(service, 1, plan, net, deck.SHORT_REDUNDANCY.upper(),
                                    deck.LONG_REDUNDANCY.upper())
            check.append([service, plan, round(result["azure"], 2),
                          *[round(result[name], 2) for name, _, _ in deck.VENDORS]])
    check.append([])
    check.append(["Python model for 1 TB, Central US, 30% discount, short-term ZRS, long-term LRS, Archive on. "
                  "Set the same inputs in Calculator to compare."])

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    book.save(output)
    return output


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default=None,
                        help="Output .xlsx. Default: backup/reports/<date>-lineamientos/calculadora_backup.xlsx")
    args = parser.parse_args()
    output = args.output or os.path.join(deck.TOPIC_DIR, "reports", f"{date.today():%Y-%m-%d}-lineamientos",
                                         "calculadora_backup.xlsx")
    try:
        list_prices, _ = deck.load_prices()
    except requests.RequestException as error:
        raise SystemExit(f"Could not read the Azure Retail Prices API: {error}")
    print(build(output, list_prices))


if __name__ == "__main__":
    main()
