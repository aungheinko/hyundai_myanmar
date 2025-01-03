import pandas as pd
import os
from datetime import datetime, timedelta

# Load the Excel file
file_path = rf"C:\Users\asservices012\OneDrive - KBTC\Database\AS_Daily_Service_Status_2025.xlsx"
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
# Define the service types to include
service_types = ["Campaign", "General", "Warranty"]
# Apply the filters
filtered_df = filtered_df[
    (filtered_df["Lead Time"] == 1) &
    (filtered_df["Service Type"].isin(service_types))
]
# Convert 'Take in Time' and 'Take Out Time' from 24-hour format to datetime.time objects
filtered_df['Take in Time'] = pd.to_datetime(filtered_df['Take in Time'], format='%H:%M:%S').dt.time
filtered_df['Take Out Time'] = pd.to_datetime(filtered_df['Take Out Time'], format='%H:%M:%S').dt.time

# Define lunch break times (24-hour format)
lunch_start = datetime.strptime('12:00:00', '%H:%M:%S').time()
lunch_end = datetime.strptime('13:00:00', '%H:%M:%S').time()

def calculate_working_hours(start_time, end_time):
    # Convert time objects to datetime objects for calculation
    today = datetime.today().date()
    start_datetime = datetime.combine(today, start_time)
    end_datetime = datetime.combine(today, end_time)
    
    # Handle cases where end_time is on the next day
    if end_datetime < start_datetime:
        end_datetime += timedelta(days=1)
    
    # Calculate working hours excluding lunch break
    total_seconds = 0
    
    # Time before lunch
    if start_time < lunch_start:
        if end_time <= lunch_start:
            delta = end_datetime - start_datetime
        else:
            delta = datetime.combine(today, lunch_start) - start_datetime
        total_seconds += delta.total_seconds()
    
    # Time after lunch
    if end_time > lunch_end:
        if start_time >= lunch_end:
            delta = end_datetime - start_datetime
        else:
            delta = end_datetime - datetime.combine(today, lunch_end)
        total_seconds += delta.total_seconds()

    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    
    return hours * 60 + minutes  # Return total minutes for easier average calculation

# Print Answers
os.system('cls')
print(f"Total Invoice Amount: {round(total_invoice_amount,2)}")
print(f'Totak Car Take In: {total_cartakein}')

print(f'30 days Rolling Average Car Take In : {round(total_cartakein/30,1)}')
print(f'30 days Rolling Average Revenue Per Day : {round(total_invoice_amount/30,2)}')
print(f'30 days Rolling Average Revenue Per Day on one Tech : {round((total_invoice_amount/30)/8,2)}')

print(f"Total Remaining Credit: {round(remaining_credit,0)}")
print(f"Total Warranty Sales: {round(total_warranty_sale,0)}")


filtered_df['Working Hours (Minutes)'] = filtered_df.apply(lambda row: calculate_working_hours(row['Take in Time'], row['Take Out Time']), axis=1)

# Calculate the average working hours in minutes
average_working_minutes = filtered_df['Working Hours (Minutes)'].mean()

# Convert the average minutes back to hours and minutes
average_hours = int(average_working_minutes // 60)
average_minutes = int(average_working_minutes % 60)

print(f"Average Working Hours: {average_hours} Hour(s) {average_minutes} Minute(s)")