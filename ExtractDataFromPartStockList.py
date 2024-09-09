import pandas as pd
import msoffcrypto
import io
import datetime

# Define the file path and password
file = r"C:\Users\asservices012\KBTC\Data Science Study Group - Hyundai Work Data - Hyundai Work Data\Part Data\2024 Parts stock list.xlsx"
password = "iklmpart24"

# Open the password-protected file using msoffcrypto
with open(file, "rb") as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=password)  # Load password
    
    # Decrypt to an in-memory file (BytesIO)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)
    
# Read the decrypted file with pandas
df = pd.read_excel(decrypted, engine="openpyxl", sheet_name="Sep",skiprows=27)
df = df.drop(index=0)
df = df.reset_index(drop=True)
df = df.drop(df.columns[0], axis=1)
column_names = list(df.columns)
columns_to_drop = ['LP',
 'Retail',
 'FOB',
 'LC',
 'LP.1',
 'Retail.1','Last update',
 'Remark','Aloc Qty','Model','MODEL','CLASS','S','Type']
df = df.drop(columns=columns_to_drop)
today = datetime.date.today()

# Get current month name
current_month_name = today.strftime("%B")

# Get previous month by subtracting one month
previous_month = today.replace(day=1) - datetime.timedelta(days=1)
previous_month_name = previous_month.strftime("%B")
df = df.rename(columns={'QTY':'PE QTY','FOB.2':'PE FOB Total','Qty':'In Qty','FOB/TT':'In FOB Total',
                       'Qty.1':'Out Qty','FOB/TT.1':'Out FOB Total','QTY.1':'End Qty','FOB.3':'End FOB Total'})

# Filter out rows where 'Updated Model' is '-'
model_filter = df[df['Updated Model'] != '-'].copy()  # Making a copy to operate on

# Convert 'End FOB Total' to numeric, forcing errors to NaN and filling NaNs with 0
model_filter.loc[:, 'End FOB Total'] = pd.to_numeric(model_filter['End FOB Total'], errors='coerce').fillna(0)

# Group by 'Updated Model' and sum 'End FOB Total'
group_by_model = model_filter.groupby('Updated Model')["End FOB Total"].sum().reset_index()

# Calculate the grand total of 'End FOB Total'
end_fob_total = group_by_model['End FOB Total'].sum()

# Create a row for the grand total
group_grand_total_row = pd.DataFrame({'Updated Model': ['End FOB Total'], 'End FOB Total': [end_fob_total]})

# Concatenate the grouped DataFrame with the grand total row
group_by_model_final = pd.concat([group_by_model, group_grand_total_row], ignore_index=True)

print(group_by_model_final)

class_filter = df[df['Updated Class'] != '-'].copy()

# Convert 'End FOB Total' to numeric (if not done earlier)
class_filter.loc[:, 'End FOB Total'] = pd.to_numeric(class_filter['End FOB Total'], errors='coerce').fillna(0)

# Group by 'Updated Class' and sum 'End FOB Total'
group_by_class = class_filter.groupby('Updated Class')["End FOB Total"].sum().reset_index()

# Calculate the grand total for 'End FOB Total' in the grouped data
end_fob_total_class = group_by_class['End FOB Total'].sum()

# Create a row for the grand total
group_grand_total_row_class = pd.DataFrame({'Updated Class': ['End FOB Total'], 'End FOB Total': [end_fob_total_class]})

# Concatenate the grouped DataFrame with the grand total row
group_by_class_final = pd.concat([group_by_class, group_grand_total_row_class], ignore_index=True)

print(group_by_class_final)

columns_to_calculate = ['In', 'Out', 'In.1', 'Out.1', 'In.2',
                        'Out.2', 'In.3', 'Out.3', 'In.4', 'Out.4', 'In.5', 'Out.5', 'In.6',
                        'Out.6', 'In.7', 'Out.7', 'In.8', 'Out.8', 'In.9', 'Out.9', 'In.10',
                        'Out.10', 'In.11', 'Out.11', 'In.12', 'Out.12', 'In.13', 'Out.13',
                        'In.14', 'Out.14', 'In.15', 'Out.15', 'In.16', 'Out.16', 'In.17',
                        'Out.17', 'In.18', 'Out.18', 'In.19', 'Out.19', 'In.20', 'Out.20',
                        'In.21', 'Out.21', 'In.22', 'Out.22', 'In.23', 'Out.23', 'In.24',
                        'Out.24', 'In.25', 'Out.25', 'In.26', 'Out.26', 'In.27', 'Out.27',
                        'In.28', 'Out.28', 'In.29', 'Out.29', 'In.30', 'Out.30']

# Ensure that 'FOB.1' is numeric
df['FOB.1'] = pd.to_numeric(df['FOB.1'], errors='coerce').fillna(0)

for column in columns_to_calculate:
    # Ensure that each column is numeric
    df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)

    # Multiply 'FOB.1' by the respective column and sum
    total_fob = (df['FOB.1'] * df[column]).sum()
    
    # Output the total FOB value for each column
    print(f"{column}\t: {round(total_fob, 2)}")
