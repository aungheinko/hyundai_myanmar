import pandas as pd
import msoffcrypto
import io
import datetime

# Define the file path and password
file = r"C:\Users\asservices012\KBTC\Data Science Study Group - Hyundai Work Data - Hyundai Work Data\Part Data\2024 Parts stock list.xlsx"

def load_excel_file(file_path, password):
    """Decrypt and load the password-protected Excel file."""
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        
        # Decrypt to an in-memory file (BytesIO)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
    return decrypted

def process_excel_data(decrypted):
    """Read and process the Excel file."""
    df = pd.read_excel(decrypted, engine="openpyxl", sheet_name="Oct", skiprows=27)
    
    # Drop unnecessary columns and reset index
    df = df.drop(index=0).reset_index(drop=True)
    df = df.drop(df.columns[0], axis=1)
    
    # Drop unwanted columns
    columns_to_drop = ['LP', 'Retail', 'FOB', 'LC', 'LP.1', 'Retail.1', 'Last update', 'Remark',
                       'Aloc Qty', 'Model', 'MODEL', 'CLASS', 'S', 'Type']
    df = df.drop(columns=columns_to_drop)
    
    # Rename columns
    df = df.rename(columns={
        'QTY': 'PE QTY', 'FOB.2': 'PE FOB Total', 'Qty': 'In Qty', 
        'FOB/TT': 'In FOB Total', 'Qty.1': 'Out Qty', 'FOB/TT.1': 'Out FOB Total',
        'QTY.1': 'End Qty', 'FOB.3': 'End FOB Total'
    })
    
    return df

def group_by_model(df):
    """Group by 'Updated Model' and calculate End FOB Total."""
    model_filter = df[df['Updated Model'] != '-'].copy()
    model_filter['End FOB Total'] = pd.to_numeric(model_filter['End FOB Total'], errors='coerce').fillna(0)
    
    group_by_model = model_filter.groupby('Updated Model')["End FOB Total"].sum().reset_index()
    end_fob_total = group_by_model['End FOB Total'].sum()
    
    # Add grand total row
    grand_total_row = pd.DataFrame({'Updated Model': ['End FOB Total'], 'End FOB Total': [end_fob_total]})
    group_by_model_final = pd.concat([group_by_model, grand_total_row], ignore_index=True)
    
    return group_by_model_final

def group_by_class(df):
    """Group by 'Updated Class' and calculate End FOB Total."""
    class_filter = df[df['Updated Class'] != '-'].copy()
    class_filter['End FOB Total'] = pd.to_numeric(class_filter['End FOB Total'], errors='coerce').fillna(0)
    
    group_by_class = class_filter.groupby('Updated Class')["End FOB Total"].sum().reset_index()
    end_fob_total_class = group_by_class['End FOB Total'].sum()
    
    # Add grand total row
    grand_total_row_class = pd.DataFrame({'Updated Class': ['End FOB Total'], 'End FOB Total': [end_fob_total_class]})
    group_by_class_final = pd.concat([group_by_class, grand_total_row_class], ignore_index=True)
    
    return group_by_class_final

def calculate_fob(df, columns_to_calculate):
    """Calculate FOB totals for the given columns."""
    df['FOB.1'] = pd.to_numeric(df['FOB.1'], errors='coerce').fillna(0)
    
    for column in columns_to_calculate:
        df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)
        total_fob = (df['FOB.1'] * df[column]).sum()
        print(f"{column}\t: {round(total_fob, 2)}")

def main():
    # Get password from user input
    password = input("Enter Password for PartStock List: ")
    
    try:
        # Load and process the file
        decrypted = load_excel_file(file, password)
        df = process_excel_data(decrypted)
        
        # Define the columns to calculate FOB
        columns_to_calculate = ['In', 'Out', 'In.1', 'Out.1', 'In.2',
                        'Out.2', 'In.3', 'Out.3', 'In.4', 'Out.4', 'In.5', 'Out.5', 'In.6',
                        'Out.6', 'In.7', 'Out.7', 'In.8', 'Out.8', 'In.9', 'Out.9', 'In.10',
                        'Out.10', 'In.11', 'Out.11', 'In.12', 'Out.12', 'In.13', 'Out.13',
                        'In.14', 'Out.14', 'In.15', 'Out.15', 'In.16', 'Out.16', 'In.17',
                        'Out.17', 'In.18', 'Out.18', 'In.19', 'Out.19', 'In.20', 'Out.20',
                        'In.21', 'Out.21', 'In.22', 'Out.22', 'In.23', 'Out.23', 'In.24',
                        'Out.24', 'In.25', 'Out.25', 'In.26', 'Out.26', 'In.27', 'Out.27',
                        'In.28', 'Out.28', 'In.29', 'Out.29', 'In.30', 'Out.30']
        
        print("Good Morning, Aung! Have a Great Day!")
        print("What would you like to know today?")
        print("1. End of FOB Total By Model\n2. End of FOB Total By Class\n3. Daily In & Out FOB\n0. Quit Program")
        
        while True:
            user_input = input("Enter your request number: ")
            
            if user_input == '1':
                print(group_by_model(df))
            elif user_input == '2':
                print(group_by_class(df))
            elif user_input == '3':
                calculate_fob(df, columns_to_calculate)
            elif user_input == '0':
                print("Goodbye!")
                break
            else:
                print("Invalid input. Please try again.")
                
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
