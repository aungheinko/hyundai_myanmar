import openpyxl

def extract_data_from_excel(file_path):
    # Load the workbook and select the specific sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Repair_Order_01']
    
    # Extract data from merged cells L26:M26
    merged_cell_LM26 = sheet['L26'].value
    
    # Extract data from cell H34
    cell_H34 = sheet['H34'].value
    
    # Extract data from merged cells I34:J34
    merged_cell_IJ34 = sheet['I34'].value
    
    # Print the extracted data
    print("Data from merged cells L26:M26:", merged_cell_LM26)
    print("Data from cell H34:", cell_H34)
    print("Data from merged cells I34:J34:", merged_cell_IJ34)

# Ask the user for the Excel file path
file_path = input("Enter the Excel file path: ")
extract_data_from_excel(file_path)
