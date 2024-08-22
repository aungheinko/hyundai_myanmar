import pandas as pd
import os
from datetime import datetime, timedelta

# Load the Excel file
file_path = rf"D:\Database\AS_Daily_Service_Status_2024.xlsx"
df = pd.read_excel(file_path, engine="openpyxl", sheet_name="IKLM_data", skiprows=1)

# Convert the "Take In Date" column to datetime format
df["Take In Date"] = pd.to_datetime(df["Take In Date"])

# Sort the data by "Take In Date" in descending order
df = df.sort_values(by="Take In Date", ascending=False)

# Get the last 30 unique working days
last_30_days = df["Take In Date"].dt.date.unique()[:30]

# Filter the DataFrame to include only rows from the last 30 unique working days
filtered_df = df[df['Take In Date'].dt.date.isin(last_30_days)]

# Further filter out rows where "Service Type" is "Auto Parts"
filtered_autopart = filtered_df[filtered_df['Service Type'] != 'Auto Parts']

# Further filter out rows where "Service Type" is one of the excluded values
exclude_services = ['Inspection', 'Pending', 'Quotation']
filtered_invoice = filtered_df[~filtered_df['Service Type'].isin(exclude_services)]

# Calculate the total sum of the "Invoice Amount" column
total_invoice_amount = filtered_invoice['Invoice Amount'].sum()
total_cartakein = filtered_autopart.shape[0]


sheet_name = "IKLM_Monthly_Report"
df_report = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# Extract values from specific cells
remaining_credit = df_report.iloc[42, 5]  # F49 corresponds to row 48, column 5 (0-indexed)
total_warranty_sale = df_report.iloc[47, 5]  # F44 corresponds to row 43, column 5 (0-indexed)

# Print Answers
os.system('cls')
print(f"Total Invoice Amount: {round(total_invoice_amount,2)}")
print(f'Totak Car Take In: {total_cartakein}')

print(f'30 days Rolling Average Car Take In : {round(total_cartakein/30,1)}')
print(f'30 days Rolling Average Revenue Per Day : {round(total_invoice_amount/30,2)}')
print(f'30 days Rolling Average Revenue Per Day on one Tech : {round((total_invoice_amount/30)/8,2)}')

print(f"Total Remaining Credit: {round(remaining_credit,0)}")
print(f"Total Warranty Sales: {round(total_warranty_sale,0)}")