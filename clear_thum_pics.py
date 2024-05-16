import os
import shutil

def clear_temporary_files():
    # Clear temporary files
    temp_folders = [
        os.environ.get("TEMP", os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "Temp")),
        os.environ.get("TMP", os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "Temp"))
    ]

    for folder in temp_folders:
        if os.path.exists(folder):
            print(f"Clearing temporary files in {folder}...")
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Failed to delete {file_path}: {e}")

def clear_recent_files():
    # Clear recent files
    recent_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Recent")
    if os.path.exists(recent_folder):
        print("Clearing recent files...")
        for filename in os.listdir(recent_folder):
            file_path = os.path.join(recent_folder, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}: {e}")

def main():
    clear_temporary_files()
    clear_recent_files()
    print("Cleanup complete.")

if __name__ == "__main__":
    main()
