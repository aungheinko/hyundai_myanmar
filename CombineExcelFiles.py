import os
import shutil

def combine_excel_files(source_folder, destination_folder):
    # Create the destination folder if it doesn't exist
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    # Iterate through all subfolders and files in the source folder
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            # Check if the file is an Excel file
            if file.endswith(".xlsx"):
                # Get the full path of the source file
                source_file_path = os.path.join(root, file)
                # Construct the destination file path
                destination_file_path = os.path.join(destination_folder, file)
                # Copy the Excel file to the destination folder
                shutil.copy(source_file_path, destination_file_path)

# Source folder containing subfolders with Excel files
source_folder = r"D:\RO_2024"
# Destination folder where all Excel files will be combined
destination_folder = r"D:\Combine Excel"

# Combine Excel files
combine_excel_files(source_folder, destination_folder)
print("Excel files combined successfully.")
