import os
from openpyxl import load_workbook

def get_files_in_folder(folder_path):
    try:
        files = os.listdir(folder_path)
        return files
    except FileNotFoundError:
        print("Folder not found.")
        return []

def get_last_value_below_total_amount(filename):
    try:
        wb = load_workbook(filename, data_only=True)  # Use data_only=True to get calculated values instead of formulas
        sheet = wb["Invoice"]
        total_amount_row = None

        # Find the row containing "Total Amount"
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "Total Amount":
                    total_amount_row = cell.row
                    break
            if total_amount_row:
                break

        if total_amount_row:
            total_amount_cell = sheet.cell(row=total_amount_row, column=2)
            next_row = sheet[total_amount_row + 1]
            last_value = None
            for cell in reversed(next_row):
                if cell.value is not None:
                    if cell.data_type == "f":  # Check if the cell contains a formula
                        last_value = cell.value  # If it's a formula, get the formula value
                    else:
                        last_value = cell.value  # Otherwise, get the displayed value
                    break
            return last_value
        else:
            print(f"Total Amount not found in the sheet of file: {filename}")
            return None
    except FileNotFoundError:
        print(f"File not found: {filename}")
        return None
    except Exception as e:
        print(f"An error occurred while processing file {filename}: {e}")
        return None

# Usage example
folder_path = input("Enter the folder path: ")

if not os.path.exists(folder_path):
    print("Invalid folder path.")
    exit()

file_list = get_files_in_folder(folder_path)

for file_name in file_list:
    excel_file = os.path.join(folder_path, file_name)
    last_underscore_index = excel_file.rfind("_")
    dot_index = excel_file.rfind(".")
    last_value = get_last_value_below_total_amount(excel_file)
    if last_value is not None:
        desired_text = excel_file[last_underscore_index + 1:dot_index]
        if isinstance(last_value,str):
            print(desired_text,0)
        else:print(desired_text,last_value)
        # print(f"Last non-null value below 'Total Amount' in {file_name}: {int(last_value)/2100}")
