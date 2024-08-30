import os
import pandas as pd

def get_excel_files_in_folder(folder_path):
    try:
        # List all files in the folder
        files = os.listdir(folder_path)
        
        # Filter files to include only those with the '.xlsx' extension
        excel_files = [file.split('.')[0] for file in files if file.endswith('.xlsx')]
        
        # Return the list of Excel file names without the extension
        return excel_files
    except FileNotFoundError:
        print("Folder not found.")
        return []

def is_empty_row(row):
    return all(value is None for value in row)

# Load the Excel file into a pandas DataFrame, skipping the first column and the first row
excel_file_path = rf"C:\Users\asservices012\KBTC\Data Science Study Group - Hyundai Work Data - Hyundai Work Data\Database\AS_Daily_Service_Status_2024.xlsx"
sheet_name = "IKLM_data"
df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=1)

# Remove the first column
df = df.iloc[:, 1:]

# Filter the DataFrame to only include rows where 'Pending Cars' column has value 'Pending'
pending_df = df[df['Pending Cars'] == 'Pending']

# Convert the 'RO No' column to a list
ronumber_list = pending_df["RO No"].tolist()

# Prompt the user to input the folder path
# folder_path = input("Enter the folder path: ")
folder_path = rf"Z:\IKLM_ASD\Reception\Mr.LEE\Pending"

# Validate the folder path
if not os.path.exists(folder_path):
    print("Invalid folder path.")
    exit()

# Get the list of Excel files with the '.xlsx' extension
excel_files = get_excel_files_in_folder(folder_path)
sequence = 0

# Open a text file for writing the output
output_file_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'output.txt')
with open(output_file_path, 'w') as f:
    for ronumber in ronumber_list:
        ro_found = False
        for file in excel_files:
            if ronumber in file:
                sequence += 1
                if 'finished' in file.lower():
                    output_text = f"{sequence} {ronumber} Finished\n"
                    print(output_text.strip())
                    f.write(output_text)
                else:
                    output_text = f"{sequence} {ronumber} Pending\n"
                    print(output_text.strip())
                    f.write(output_text)
                ro_found = True
                break
        if not ro_found:
            output_text = f"RO number {ronumber} not found in folder.\n"
            print(output_text.strip())
            f.write(output_text)

print("Output written to", output_file_path)
