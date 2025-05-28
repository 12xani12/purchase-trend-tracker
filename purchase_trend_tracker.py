import pandas as pd
from datetime import datetime, timedelta

def load_product_list(file_path):
    # Load product list Excel
    df = pd.read_excel(file_path)
    return df

def load_sales_report(file_path):
    # Load sales report Excel
    df = pd.read_excel(file_path)
    return df

def calculate_sales_trends(product_df, sales_df, date_col='Order Date', code_col='Code', quantity_col='Quantity'):
    # Convert order date to datetime if needed
    if not pd.api.types.is_datetime64_any_dtype(sales_df[date_col]):
        sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors='coerce')

    # Define date thresholds
    today = datetime.today()
    three_months_ago = today - timedelta(days=90)
    twelve_months_ago = today - timedelta(days=365)

    # Filter sales in last 3 months and 12 months
    sales_3m = sales_df[sales_df[date_col] >= three_months_ago]
    sales_12m = sales_df[sales_df[date_col] >= twelve_months_ago]

    # Aggregate quantities by product code
    sales_3m_agg = sales_3m.groupby(code_col)[quantity_col].sum().reset_index()
    sales_3m_agg.rename(columns={quantity_col: 'Quantity Ordered (3 months)'}, inplace=True)

    sales_12m_agg = sales_12m.groupby(code_col)[quantity_col].sum().reset_index()
    sales_12m_agg.rename(columns={quantity_col: 'Quantity Ordered (12 months)'}, inplace=True)

    # Merge aggregated sales with product list
    merged_df = product_df.merge(sales_3m_agg, how='left', left_on='Code', right_on=code_col)
    merged_df = merged_df.merge(sales_12m_agg, how='left', left_on='Code', right_on=code_col)

    # Drop duplicate code columns from merges
    merged_df.drop(columns=[code_col+'_x', code_col+'_y'], inplace=True, errors='ignore')

    # Fill NaN with zeros for products with no sales
    merged_df['Quantity Ordered (3 months)'] = merged_df['Quantity Ordered (3 months)'].fillna(0).astype(int)
    merged_df['Quantity Ordered (12 months)'] = merged_df['Quantity Ordered (12 months)'].fillna(0).astype(int)

    return merged_df

def export_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"Exported combined data to {output_path}")

def main(product_list_path, sales_report_path, output_path):
    print("Loading product list...")
    product_df = load_product_list(product_list_path)
    print("Loading sales report...")
    sales_df = load_sales_report(sales_report_path)

    print("Calculating sales trends...")
    combined_df = calculate_sales_trends(product_df, sales_df)

    print("Exporting results...")
    export_to_excel(combined_df, output_path)
    print("Done.")

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Product Purchase & Sales Trend Tracker Tool')
    parser.add_argument('product_list', help='Path to the product list Excel file')
    parser.add_argument('sales_report', help='Path to the sales report Excel file')
    parser.add_argument('output_file', help='Path to save the combined output Excel file')

    args = parser.parse_args()

    main(args.product_list, args.sales_report, args.output_file)