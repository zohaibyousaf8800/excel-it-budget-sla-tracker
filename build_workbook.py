#!/usr/bin/env python3
"""
IT Budget & SLA Tracker — workbook generator.

Builds a formula-driven Excel workbook for two monthly IT management jobs:

  1. Budget control — annual IT budget by category, quarterly actual spend,
     remaining budget, variance and over-budget flags.
  2. Helpdesk SLA reporting — tickets logged per month, how many were
     resolved inside the SLA target, and the resulting SLA percentage.

A vendor contract sheet with renewal alerts and a one-page Dashboard of
live KPIs and charts sit on top of those two registers. No totals are
hard-coded: every headline figure is an Excel formula, so the workbook
recalculates as soon as a register is updated.

Author: Zohaib Yousaf (info@zohaibyousaf.com)
Requires: openpyxl  (pip install openpyxl)
Run:      python build_workbook.py
Output:   IT_Budget_SLA_Tracker.xlsx

The data shipped in the workbook is realistic SAMPLE data for
demonstration. Replace it with your own figures before use.
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).with_name("IT_Budget_SLA_Tracker.xlsx")

# ------------------------------------------------------------------ theme
NAVY = "0F2A44"
TEAL = "0E7C7B"
LIGHT = "E8EEF3"
WHITE = "FFFFFF"
GREY_TEXT = "475569"
GOLD = "F2B134"

header_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
header_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name="Calibri", size=18, bold=True, color=NAVY)
sub_font = Font(name="Calibri", size=11, color=GREY_TEXT)
kpi_label_font = Font(name="Calibri", size=11, bold=True, color=GREY_TEXT)
kpi_value_font = Font(name="Calibri", size=22, bold=True, color=NAVY)
total_font = Font(name="Calibri", size=11, bold=True, color=NAVY)
thin = Side(style="thin", color="CBD5E1")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
wrap = Alignment(vertical="center", wrap_text=True)
MONEY = '"$"#,##0'
PCT = "0.0%"


def style_header(ws, row, ncols):
    """Apply the shared navy header style to one row."""
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[row].height = 30


def add_table(ws, ref, name):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(tab)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def d(text):
    return date.fromisoformat(text)


# ------------------------------------------------------------- sample data
# Budget lines: (Category, Item, Annual Budget, Q1, Q2, Q3, Q4 actual)
# Q4 is 0 in the sample because the reporting year is only at October.
BUDGET = [
    ("Hardware", "End-user devices refresh (laptops & desktops)", 45000, 12800, 11500, 9400, 0),
    ("Hardware", "Servers & storage", 30000, 8200, 7600, 5100, 0),
    ("Software & Licences", "Microsoft 365 & software licences", 24000, 6000, 6000, 6000, 0),
    ("Software & Licences", "Design & productivity tools", 4000, 2350, 1900, 0, 0),
    ("Cloud (Azure/M365)", "Azure infrastructure & cloud services", 36000, 8400, 9100, 9800, 0),
    ("Cloud (Azure/M365)", "SaaS subscriptions", 15000, 3700, 3800, 3900, 0),
    ("Network", "Internet, WAN links & network hardware", 32000, 7900, 8100, 8050, 0),
    ("Security", "Security tools, firewall licences & backup", 28000, 6900, 7050, 7000, 0),
    ("Support Contracts", "Vendor support & maintenance contracts", 20000, 5000, 5000, 5000, 0),
    ("Training", "IT staff training & certifications", 8000, 1800, 2400, 1200, 0),
]

# SLA rows: (Month, Logged, Resolved, Resolved Within SLA, Avg Resolution hrs, P1)
SLA = [
    ("Jan 2026", 148, 141, 140, 7.4, 1),
    ("Feb 2026", 132, 128, 126, 6.8, 0),
    ("Mar 2026", 171, 158, 149, 9.6, 2),
    ("Apr 2026", 156, 150, 148, 7.1, 1),
    ("May 2026", 164, 159, 156, 7.9, 1),
    ("Jun 2026", 149, 144, 142, 6.5, 0),
    ("Jul 2026", 138, 133, 131, 6.9, 1),
    ("Aug 2026", 127, 123, 121, 6.2, 0),
    ("Sep 2026", 186, 172, 166, 10.8, 3),
    ("Oct 2026", 173, 160, 157, 8.7, 1),
]

# Vendors: (Vendor, Service, Contract Start, Renewal Date, Annual Cost, Notice days)
VENDORS = [
    ("Microsoft", "Microsoft 365 & Azure agreement", "2026-03-01", "2027-03-01", 60000, 90),
    ("Adobe", "Creative & document tools", "2025-12-15", "2026-12-15", 5400, 30),
    ("Primary ISP", "Fibre internet & WAN links", "2025-11-15", "2026-11-15", 21600, 30),
    ("Endpoint Security Vendor", "Antivirus & endpoint protection", "2026-01-31", "2027-01-31", 8400, 60),
    ("Backup Vendor", "Backup & recovery platform", "2026-05-30", "2027-05-30", 9600, 60),
    ("Hardware Maintenance Co.", "Server & network support contract", "2026-07-01", "2027-07-01", 14400, 90),
    ("Domain & Hosting Provider", "Domains, DNS & web hosting", "2026-02-01", "2027-02-01", 1800, 30),
]

# ---------------------------------------------------------------- budgets
wb = Workbook()

ws = wb.active
ws.title = "IT_Budget"
ws.sheet_properties.tabColor = NAVY
headers = ["Category", "Item", "Annual Budget", "Q1 Actual", "Q2 Actual",
           "Q3 Actual", "Q4 Actual", "Total Actual", "Remaining",
           "Variance %", "Status"]
ws.append(headers)
for row in BUDGET:
    r = ws.max_row + 1
    ws.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6],
               f"=SUM(D{r}:G{r})",            # Total Actual
               f"=C{r}-H{r}",                  # Remaining
               f"=IF(C{r}=0,0,(H{r}-C{r})/C{r})",  # Variance % (over = positive)
               f'=IF(H{r}>C{r},"Over Budget","On Track")'])
style_header(ws, 1, len(headers))
last = ws.max_row
# Totals row — every cell a formula.
r = last + 1
ws.cell(row=r, column=1, value="TOTAL").font = total_font
ws.cell(row=r, column=2, value="All categories").font = total_font
for col, letter in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")]:
    c = ws.cell(row=r, column=col, value=f"=SUM({letter}2:{letter}{last})")
    c.font = total_font
c = ws.cell(row=r, column=10, value=f"=IF(C{r}=0,0,(H{r}-C{r})/C{r})")
c.font = total_font
c = ws.cell(row=r, column=11,
            value=f'=IF(COUNTIF(K2:K{last},"Over Budget")>0,"Review — see lines above","On Track")')
c.font = total_font
for row_i in range(2, r + 1):
    for col in range(3, 10):
        ws.cell(row=row_i, column=col).number_format = MONEY
    ws.cell(row=row_i, column=10).number_format = PCT
    ws.cell(row=row_i, column=11).alignment = center
    ws.cell(row=row_i, column=1).alignment = wrap
    ws.cell(row=row_i, column=2).alignment = wrap
add_table(ws, f"A1:K{last}", "BudgetTable")
set_widths(ws, [21, 44, 14, 12, 12, 12, 12, 13, 13, 11, 30])
ws.freeze_panes = "A2"
# Variance: positive (over budget) is the problem, so flag it red.
ws.conditional_formatting.add(f"J2:J{last}",
    CellIsRule(operator="greaterThan", formula=["0"],
               fill=PatternFill("solid", fgColor="FECACA")))
ws.conditional_formatting.add(f"K2:K{last}",
    FormulaRule(formula=['$K2="Over Budget"'],
                fill=PatternFill("solid", fgColor="FECACA")))
BUDGET_TOTAL_ROW = r

# ------------------------------------------------------------- SLA monthly
ws = wb.create_sheet("SLA_Monthly")
ws.sheet_properties.tabColor = TEAL
headers = ["Month", "Tickets Logged", "Resolved", "Resolved Within SLA",
           "SLA %", "Avg Resolution (hrs)", "P1 Incidents"]
ws.append(headers)
for row in SLA:
    r = ws.max_row + 1
    ws.append([row[0], row[1], row[2], row[3],
               f"=IF(B{r}=0,0,D{r}/B{r})",   # SLA % = within SLA / logged
               row[4], row[5]])
style_header(ws, 1, len(headers))
last = ws.max_row
# Totals / average row — SLA % is worked out from the totals, not averaged.
r = last + 1
ws.cell(row=r, column=1, value="TOTAL / AVERAGE (Jan–Oct)").font = total_font
for col, letter in [(2, "B"), (3, "C"), (4, "D"), (7, "G")]:
    c = ws.cell(row=r, column=col, value=f"=SUM({letter}2:{letter}{last})")
    c.font = total_font
c = ws.cell(row=r, column=5, value=f"=IF(B{r}=0,0,D{r}/B{r})")
c.font = total_font
c = ws.cell(row=r, column=6, value=f"=ROUND(AVERAGE(F2:F{last}),1)")
c.font = total_font
for row_i in range(2, r + 1):
    ws.cell(row=row_i, column=5).number_format = PCT
    for col in (2, 3, 4, 6, 7):
        ws.cell(row=row_i, column=col).alignment = center
add_table(ws, f"A1:G{last}", "SLATable")
set_widths(ws, [24, 15, 11, 19, 10, 19, 13])
ws.freeze_panes = "A2"
# SLA target is 95% — anything under it needs attention.
ws.conditional_formatting.add(f"E2:E{last}",
    CellIsRule(operator="lessThan", formula=["0.95"],
               fill=PatternFill("solid", fgColor="FDE68A")))
SLA_TOTAL_ROW = r

# ---------------------------------------------------------------- vendors
ws = wb.create_sheet("Vendors")
ws.sheet_properties.tabColor = GOLD
headers = ["Vendor", "Service", "Contract Start", "Renewal Date",
           "Annual Cost", "Notice Period (days)", "Days to Renewal",
           "Renewal Alert"]
ws.append(headers)
for v in VENDORS:
    r = ws.max_row + 1
    ws.append([v[0], v[1], d(v[2]), d(v[3]), v[4], v[5],
               f"=D{r}-TODAY()",  # days left before the contract renews
               f'=IF(G{r}<=F{r},"RENEW SOON","OK")'])
style_header(ws, 1, len(headers))
last = ws.max_row
for row_i in range(2, last + 1):
    for col in (3, 4):
        ws.cell(row=row_i, column=col).number_format = "yyyy-mm-dd"
    ws.cell(row=row_i, column=5).number_format = MONEY
    for col in (6, 7, 8):
        ws.cell(row=row_i, column=col).alignment = center
add_table(ws, f"A1:H{last}", "VendorsTable")
set_widths(ws, [26, 38, 14, 13, 13, 17, 15, 14])
ws.freeze_panes = "A2"
ws.conditional_formatting.add(f"H2:H{last}",
    FormulaRule(formula=['$H2="RENEW SOON"'],
                fill=PatternFill("solid", fgColor="FDE68A")))

# ---------------------------------------------------------------- summary
# Small formula-only table that feeds the Dashboard charts. Nothing here is
# typed in by hand — change SLA_Monthly or IT_Budget and this follows.
ws = wb.create_sheet("Summary")
ws.sheet_properties.tabColor = "64748B"
ws.append(["Month", "Tickets Logged", "SLA %"])
for i in range(len(SLA)):
    src = i + 2  # matching row on SLA_Monthly
    ws.append([f"=SLA_Monthly!A{src}", f"=SLA_Monthly!B{src}",
               f"=SLA_Monthly!E{src}"])
style_header(ws, 1, 3)
for row_i in range(2, ws.max_row + 1):
    ws.cell(row=row_i, column=3).number_format = PCT
set_widths(ws, [14, 15, 10])
SUMMARY_LAST = ws.max_row

# --------------------------------------------------------------- dashboard
ws = wb.create_sheet("Dashboard", 0)
ws.sheet_properties.tabColor = TEAL
ws["B2"] = "IT Budget & SLA Tracker — Dashboard"
ws["B2"].font = title_font
ws["B3"] = ("Built by Zohaib Yousaf — IT Manager  |  info@zohaibyousaf.com  |  "
            "Every figure below is a live formula from IT_Budget and SLA_Monthly. "
            "Sample data — replace with actuals.")
ws["B3"].font = sub_font
ws.merge_cells("B3:I3")

kpis = [
    ("Total Annual IT Budget", f"=IT_Budget!C{BUDGET_TOTAL_ROW}"),
    ("Actual Spend YTD", f"=IT_Budget!H{BUDGET_TOTAL_ROW}"),
    ("Budget Variance %", f"=IT_Budget!J{BUDGET_TOTAL_ROW}"),
    ("SLA Met % (YTD)", f"=SLA_Monthly!E{SLA_TOTAL_ROW}"),
    ("Tickets This Month (Oct 2026)", f"=SLA_Monthly!B{SLA_TOTAL_ROW - 1}"),
    ("Contracts Flagged RENEW SOON", '=COUNTIF(Vendors!H:H,"RENEW SOON")'),
]
for i, (label, formula) in enumerate(kpis):
    col = 2 + (i % 3) * 3
    r = 5 + (i // 3) * 4
    lbl = ws.cell(row=r, column=col, value=label)
    lbl.font = kpi_label_font
    val = ws.cell(row=r + 1, column=col, value=formula)
    val.font = kpi_value_font
    if "Variance" in label or "SLA Met" in label:
        val.number_format = PCT
    elif "Tickets" in label or "RENEW" in label:
        val.number_format = "#,##0"
    else:
        val.number_format = MONEY
    for rr in (r, r + 1):
        for cc in range(col, col + 2):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
            ws.cell(row=rr, column=cc).border = border

ws["B14"] = ("How to use:  1) Update Q1–Q4 actuals on IT_Budget as invoices land — "
             "variance and status update themselves.  2) Add one row per month to "
             "SLA_Monthly; months under the 95% SLA target are highlighted.  "
             "3) Keep Vendors renewal dates current — contracts inside their notice "
             "period flag RENEW SOON on this Dashboard.")
ws["B14"].font = sub_font
ws["B14"].alignment = wrap
ws.merge_cells("B14:I16")
set_widths(ws, [2, 18, 18, 4, 18, 18, 4, 18, 18])

# Chart 1 — Annual budget vs actual spend, by budget line (from IT_Budget).
bar = BarChart()
bar.type = "col"
bar.style = 10
bar.title = "Annual Budget vs Actual Spend by Item"
bar.y_axis.title = "USD"
bar.x_axis.title = "Budget item"
data = Reference(wb["IT_Budget"], min_col=3, min_row=1, max_row=len(BUDGET) + 1)
data2 = Reference(wb["IT_Budget"], min_col=8, min_row=1, max_row=len(BUDGET) + 1)
cats = Reference(wb["IT_Budget"], min_col=2, min_row=2, max_row=len(BUDGET) + 1)
bar.add_data(data, titles_from_data=True)
bar.add_data(data2, titles_from_data=True)
bar.set_categories(cats)
bar.height, bar.width = 8, 15
ws.add_chart(bar, "B18")

# Chart 2 — SLA % by month (from Summary, which itself reads SLA_Monthly).
line = LineChart()
line.style = 10
line.title = "SLA Met % by Month (target 95%)"
line.y_axis.title = "SLA %"
line.y_axis.number_format = "0%"
line.x_axis.title = "Month"
data = Reference(wb["Summary"], min_col=3, min_row=1, max_row=SUMMARY_LAST)
cats = Reference(wb["Summary"], min_col=1, min_row=2, max_row=SUMMARY_LAST)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
line.height, line.width = 8, 15
ws.add_chart(line, "B34")

for name in wb.sheetnames:
    wb[name].sheet_properties.pageSetUpPr.fitToPage = True
    wb[name].page_setup.fitToWidth = 1
    wb[name].page_setup.fitToHeight = 0
    wb[name].page_setup.orientation = "landscape"

wb.properties.creator = "Zohaib Yousaf"
wb.properties.title = "IT Budget & SLA Tracker"
wb.properties.description = (
    "Formula-driven IT budget control and helpdesk SLA monthly reporting "
    "workbook by Zohaib Yousaf, IT Manager."
)
wb.save(OUT)
print(f"Saved {OUT} ({OUT.stat().st_size} bytes)")
