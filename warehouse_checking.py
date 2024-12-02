import pandas as pd
import os

# File paths
base_dir = r"C:\Users\asservices012\KBTC\Data Science Study Group - Hyundai Work Data - Hyundai Work Data\warehouse_data"
take_in_file = os.path.join(base_dir, "Parts_Take_In.xlsx")
take_out_file = os.path.join(base_dir, "Parts_Take_Out.xlsx")
checking_file = os.path.join(base_dir, "Warehouse_Checking_Data.xlsx")

# Load data
take_in_df = pd.read_excel(take_in_file)
take_out_df = pd.read_excel(take_out_file)
checking_df = pd.read_excel(checking_file)

# Aggregate Take In and Take Out data by Part No
take_in_summary = take_in_df.groupby('Parts No')['Qty'].sum().reset_index().rename(columns={'Qty': 'Total Take In Qty'})
take_out_summary = take_out_df.groupby('Parts No')['Qty'].sum().reset_index().rename(columns={'Qty': 'Total Take Out Qty'})

# Merge Take In and Take Out data with the Warehouse Checking Data
merged_df = checking_df.merge(take_in_summary, on='Parts No', how='left').merge(take_out_summary, on='Parts No', how='left')

# Fill missing values with 0 (for parts without any take-in or take-out records)
merged_df['Total Take In Qty'] = merged_df['Total Take In Qty'].fillna(0)
merged_df['Total Take Out Qty'] = merged_df['Total Take Out Qty'].fillna(0)

# Calculate Expected Stock
merged_df['Expected Stock'] = merged_df['Base Stock Qty'] + merged_df['Total Take In Qty'] - merged_df['Total Take Out Qty']

# Calculate Actual Stock
merged_df['Actual Stock'] = (
    merged_df['Good Part Qty'] +
    merged_df['Damage Qty'] +
    merged_df['Wrong Part Qty'] +
    merged_df['Refound Qty']
)

# Identify Discrepancies
merged_df['Discrepancy'] = merged_df['Expected Stock'] - merged_df['Actual Stock']
merged_df['Status'] = merged_df['Discrepancy'].apply(
    lambda x: 'Missing' if x > 0 else ('Extra' if x < 0 else 'Match')
)

# Save the results to a new Excel file
output_file = os.path.join(base_dir, "Stock_Checking_Results.xlsx")
merged_df.to_excel(output_file, index=False)

print(f"Stock checking results saved to: {output_file}")
