import os

def find_files_with_string(root_folder, search_string):
    matching_files = []
    
    # Walk through each directory and subdirectory
    for root, dirs, files in os.walk(root_folder):
        # Iterate through each file in the current directory
        for file in files:
            # Check if the search string is in the filename
            if search_string in file:
                matching_files.append(os.path.join(root, file))
    
    return matching_files

# Root folder path where you want to search for files
root_folder_path = 'Z:\IKLM_ASD\Reception\RO\RO_2023'

# Search string
search_string = '4R-4616'

# Find files containing the search string in their filenames
matching_files = find_files_with_string(root_folder_path, search_string)

# Print the matching file paths
if matching_files:
    print("Matching files:")
    for file_path in matching_files:
        print(file_path)
else:
    print("No matching files found.")
