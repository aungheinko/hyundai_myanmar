import os
import pandas as pd

def list_excel_files(directory_path):
    try:
        # List all files in the directory
        files = os.listdir(directory_path)
        
        # Filter files to include only those with the '.xlsx' extension
        excel_files = [file.split('.')[0] for file in files if file.endswith('.xlsx')]
        
        # Return the list of Excel file names without the extension
        return excel_files
    except FileNotFoundError:
        print("Directory not found.")
        return []

def row_is_empty(row):
    return all(value is None for value in row)

# Load the Excel file into a pandas DataFrame, skipping the first column and the first row
excel_file_path = r"D:\Working Enviroment\Database\Data_2024_v2.2.xlsx"
sheet_name = "IKLM_data"
df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=1)

# Remove the first column
df = df.iloc[:, 1:]

# Filter the DataFrame to only include rows where 'Pending Cars' column has value 'Pending'
pending_df = df[df['Pending Cars'] == 'Pending']

# Convert the 'RO No' column to a list
ro_number_list = pending_df["RO No"].tolist()

# Prompt the user to input the folder path
# directory_path = input("Enter the folder path: ")
directory_path = rf"Z:\IKLM_ASD\Reception\Mr.LEE\Pending"

# Validate the folder path
if not os.path.exists(directory_path):
    print("Invalid directory path.")
    exit()

# Get the list of Excel files with the '.xlsx' extension
excel_files = list_excel_files(directory_path)
sequence = 0

# Open a text file for writing the output in the same directory path
output_file_path = os.path.join(folder_path , 'output.txt')
with open(output_file_path, 'w') as f:
    for ro_number in ro_number_list:
        ro_found = False
        for file in excel_files:
            if ro_number in file:
                sequence += 1
                if 'finished' in file.lower():
                    output_text = f"{sequence} {ro_number} Finished\n"
                    print(output_text.strip())
                    f.write(output_text)
                else:
                    output_text = f"{sequence} {ro_number} Pending\n"
                    print(output_text.strip())
                    f.write(output_text)
                ro_found = True
                break
        if not ro_found:
            output_text = f"RO number {ro_number} not found in folder.\n"
            print(output_text.strip())
            f.write(output_text)

print("Output written to", output_file_path)
