Product Purchase & Sales Trend Tracker Tool
===========================================

This tool imports a product list and a sales report Excel file,
calculates the quantity ordered for each product over the last 3 months
and 12 months, merges the data, and exports the result as an Excel file.

Usage:
------
1. Install dependencies:
   pip install -r requirements.txt

2. Run the script:
   python purchase_trend_tracker.py product_list.xlsx sales_report.xlsx output.xlsx

Inputs:
- product_list.xlsx must include columns: Internal Reference, Code, Min, Max, Quantity on Hand, Forecasted quantity
- sales_report.xlsx must include columns with at least: Code, Order Date, Quantity

Output:
- An Excel file with all original product data plus two new columns:
  'Quantity Ordered (3 months)' and 'Quantity Ordered (12 months)'