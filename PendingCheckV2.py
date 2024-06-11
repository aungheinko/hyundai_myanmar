import os
import pandas as pd

def get_excel_files_in_folder(folder_path):
    """Get a list of Excel files (without extension) in the specified folder."""
    try:
        files = os.listdir(folder_path)
        excel_files = [file.split('.')[0] for file in files if file.endswith('.xlsx')]
        return excel_files
    except FileNotFoundError:
        print("Folder not found.")
        return []

def is_empty_row(row):
    """Check if a row is empty."""
    return all(value is None for value in row)

def process_excel_file(excel_file_path, sheet_name):
    """Load the Excel file and filter the DataFrame to include only pending cars."""
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=1)
    df = df.iloc[:, 1:]  # Remove the first column
    pending_df = df[df['Pending Cars'] == 'Pending']
    return pending_df["RONo"].tolist()

def write_output_to_file(output_file_path, ronumber_list, excel_files):
    """Write the output to a text file."""
    sequence = 0
    with open(output_file_path, 'w') as f:
        for ronumber in ronumber_list:
            ro_found = False
            for file in excel_files:
                if ronumber in file:
                    sequence += 1
                    status = "Finished" if 'finished' in file.lower() else "Pending"
                    output_text = f"{sequence} {ronumber} {status}\n"
                    print(output_text.strip())
                    f.write(output_text)
                    ro_found = True
                    break
            if not ro_found:
                output_text = f"RO number {ronumber} not found in folder.\n"
                print(output_text.strip())
                f.write(output_text)

def main():
    # File paths
    excel_file_path = r"Z:\IKLM_ASD\AUNG HEIN KO\Database\Data_2024_v2.2.xlsx"
    sheet_name = "IKLM_data"
    folder_path = rf"Z:\IKLM_ASD\Reception\Mr.LEE\Pending"
    output_file_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'output.txt')

    # Validate folder path
    if not os.path.exists(folder_path):
        print("Invalid folder path.")
        return

    # Process Excel file
    ronumber_list = process_excel_file(excel_file_path, sheet_name)

    # Get Excel files in the folder
    excel_files = get_excel_files_in_folder(folder_path)

    # Write output to file
    write_output_to_file(output_file_path, ronumber_list, excel_files)
    print("Output written to", output_file_path)

if __name__ == "__main__":
    main()
