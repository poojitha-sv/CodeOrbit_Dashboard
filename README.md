# CodeOrbit Dashboard Creation Task

This is my Task 3 submission for the CodeOrbit Tech Data Analyst internship.

## Project Title
Dashboard Creation Using Excel - Sales Performance Dashboard

## Technologies Used
- Excel (built with Python's openpyxl library, which writes real .xlsx files)

## Setup Instructions
1. Download Sales_Dashboard.xlsx
2. Open it in Excel (use the Excel app, not a generic file viewer, so charts render correctly)
3. Click the "Dashboard" tab at the bottom to see the KPIs, tables, and charts (raw data is on the "Data" tab)

## GitHub Repository Link
https://github.com/poojitha-sv/CodeOrbit_Dashboard

## What This Project Does
I built a one-page sales dashboard in Excel using the same sample sales data from Task 2. It shows 3 KPI cards, 3 summary tables, and 3 charts, all using live formulas (SUM, AVERAGE, COUNTA, SUMIFS) so the numbers update automatically if the data changes.

## Files Here
- Sales_Dashboard.xlsx - the dashboard (open the "Dashboard" tab; raw data is in the "Data" tab)
- build_dashboard.py - the Python script (openpyxl) I used to build it

## Dashboard Contents
- KPI cards: Total Sales, Average Order Value, Total Orders
- Table + bar chart: Sales by Category
- Table + bar chart: Sales by Region
- Table + pie chart: Sales Share by Product
