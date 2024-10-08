import os

def delete_thumb_pictures(folder_path):
    # Loop through all files in the given folder
    for filename in os.listdir(folder_path):
        # Check if the file is a picture and contains "_thumb" in the name
        if "_thumb" in filename and (filename.endswith(".jpg") or filename.endswith(".png") or filename.endswith(".jpeg")):
            file_path = os.path.join(folder_path, filename)
            try:
                os.remove(file_path)  # Delete the file
                print(f"Deleted: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

# Example usage
folder_path = rf"C:\Users\asservices012\Downloads\Telegram Desktop\ChatExport_2024-10-07\photos"  # Replace with the folder path
delete_thumb_pictures(folder_path)
