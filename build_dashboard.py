import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

# ---- load raw data ----
df = pd.read_csv("sales_data.csv")

wb = Workbook()

# =========================================================
# SHEET 1: Data (raw data + a TotalPrice formula column)
# =========================================================
ws_data = wb.active
ws_data.title = "Data"

headers = list(df.columns) + ["TotalPrice"]
for col_num, header in enumerate(headers, start=1):
    cell = ws_data.cell(row=1, column=col_num, value=header)
    cell.font = Font(bold=True, name="Arial")

for row_num, row in enumerate(df.itertuples(index=False), start=2):
    for col_num, value in enumerate(row, start=1):
        ws_data.cell(row=row_num, column=col_num, value=value)
    # TotalPrice formula = Quantity * UnitPrice (columns F and G)
    ws_data.cell(row=row_num, column=8, value=f"=F{row_num}*G{row_num}")

for col in range(1, 9):
    ws_data.column_dimensions[get_column_letter(col)].width = 16

last_row = len(df) + 1  # last data row number

# =========================================================
# SHEET 2: Dashboard
# =========================================================
ws = wb.create_sheet("Dashboard")
ws.sheet_view.showGridLines = False

arial_bold = Font(name="Arial", bold=True)
title_font = Font(name="Arial", bold=True, size=16, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
kpi_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
thin = Side(style="thin", color="B7B7B7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---- Title bar ----
ws.merge_cells("A1:H1")
ws["A1"] = "Sales Performance Dashboard"
ws["A1"].font = title_font
ws["A1"].fill = header_fill
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 30

# ---- KPI cards (row 3-5) ----
kpi_labels = ["Total Sales", "Average Order Value", "Total Orders"]
kpi_formulas = [
    f"=SUM(Data!H2:H{last_row})",
    f"=AVERAGE(Data!H2:H{last_row})",
    f"=COUNTA(Data!A2:A{last_row})",
]
kpi_formats = ["$#,##0.00", "$#,##0.00", "0"]

col_start = 1
for i, (label, formula, fmt) in enumerate(zip(kpi_labels, kpi_formulas, kpi_formats)):
    c = col_start + i * 3
    ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + 1)
    cell = ws.cell(row=3, column=c, value=label)
    cell.font = arial_bold
    cell.alignment = Alignment(horizontal="center")
    cell.fill = kpi_fill

    ws.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
    val_cell = ws.cell(row=4, column=c, value=formula)
    val_cell.font = Font(name="Arial", bold=True, size=14, color="1F4E78")
    val_cell.alignment = Alignment(horizontal="center")
    val_cell.number_format = fmt
    val_cell.fill = kpi_fill

# ---- Table: Sales by Category (SUMIFS formulas) ----
ws["A7"] = "Sales by Category"
ws["A7"].font = arial_bold

categories = sorted(df["Category"].unique())
ws["A8"] = "Category"
ws["B8"] = "Total Sales"
ws["A8"].font = ws["B8"].font = arial_bold
for i, cat in enumerate(categories, start=9):
    ws.cell(row=i, column=1, value=cat)
    ws.cell(row=i, column=2, value=f'=SUMIFS(Data!$H$2:$H${last_row},Data!$D$2:$D${last_row},A{i})')
    ws.cell(row=i, column=2).number_format = "$#,##0.00"
cat_end_row = 8 + len(categories)

# ---- Table: Sales by Region (SUMIFS formulas) ----
ws["D7"] = "Sales by Region"
ws["D7"].font = arial_bold

regions = sorted(df["Region"].unique())
ws["D8"] = "Region"
ws["E8"] = "Total Sales"
ws["D8"].font = ws["E8"].font = arial_bold
for i, reg in enumerate(regions, start=9):
    ws.cell(row=i, column=4, value=reg)
    ws.cell(row=i, column=5, value=f'=SUMIFS(Data!$H$2:$H${last_row},Data!$E$2:$E${last_row},D{i})')
    ws.cell(row=i, column=5).number_format = "$#,##0.00"
reg_end_row = 8 + len(regions)

# ---- Table: Top Products (SUMIFS formulas) ----
ws["G7"] = "Sales by Product"
ws["G7"].font = arial_bold

products = sorted(df["Product"].unique())
ws["G8"] = "Product"
ws["H8"] = "Total Sales"
ws["G8"].font = ws["H8"].font = arial_bold
for i, prod in enumerate(products, start=9):
    ws.cell(row=i, column=7, value=prod)
    ws.cell(row=i, column=8, value=f'=SUMIFS(Data!$H$2:$H${last_row},Data!$C$2:$C${last_row},G{i})')
    ws.cell(row=i, column=8).number_format = "$#,##0.00"
prod_end_row = 8 + len(products)

# apply borders to all three tables
for start_row, end_row, cols in [(8, cat_end_row, (1, 2)), (8, reg_end_row, (4, 5)), (8, prod_end_row, (7, 8))]:
    for r in range(start_row, end_row + 1):
        for c in cols:
            ws.cell(row=r, column=c).border = border

# =========================================================
# Charts
# =========================================================
# Bar chart: Sales by Category
bar1 = BarChart()
bar1.title = "Sales by Category"
bar1.y_axis.title = "Total Sales ($)"
data_ref = Reference(ws, min_col=2, min_row=8, max_row=cat_end_row)
cats_ref = Reference(ws, min_col=1, min_row=9, max_row=cat_end_row)
bar1.add_data(data_ref, titles_from_data=True)
bar1.set_categories(cats_ref)
bar1.width, bar1.height = 10, 7
ws.add_chart(bar1, "A15")

# Bar chart: Sales by Region
bar2 = BarChart()
bar2.title = "Sales by Region"
bar2.y_axis.title = "Total Sales ($)"
data_ref2 = Reference(ws, min_col=5, min_row=8, max_row=reg_end_row)
cats_ref2 = Reference(ws, min_col=4, min_row=9, max_row=reg_end_row)
bar2.add_data(data_ref2, titles_from_data=True)
bar2.set_categories(cats_ref2)
bar2.width, bar2.height = 10, 7
ws.add_chart(bar2, "D15")

# Pie chart: Sales by Product
pie1 = PieChart()
pie1.title = "Sales Share by Product"
data_ref3 = Reference(ws, min_col=8, min_row=8, max_row=prod_end_row)
cats_ref3 = Reference(ws, min_col=7, min_row=9, max_row=prod_end_row)
pie1.add_data(data_ref3, titles_from_data=True)
pie1.set_categories(cats_ref3)
pie1.width, pie1.height = 10, 7
ws.add_chart(pie1, "G15")

# column widths for dashboard sheet
for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
    ws.column_dimensions[col].width = 14

# page setup so it prints/exports as one readable page
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save("Sales_Dashboard.xlsx")
print("Saved Sales_Dashboard.xlsx")
