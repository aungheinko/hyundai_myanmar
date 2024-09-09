import pandas as pd
import msoffcrypto
import io

# Define the file path and password
file = r"C:\Users\asservices012\KBTC\Data Science Study Group - Hyundai Work Data - Hyundai Work Data\Part Data\2024 Parts stock list.xlsx"
password = input("Enter Your Password For PartStockList : ")

# Open the password-protected file using msoffcrypto
with open(file, "rb") as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=password)  # Load password
    
    # Decrypt to an in-memory file (BytesIO)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)

# Read the decrypted file with pandas
df = pd.read_excel(decrypted, engine="openpyxl", sheet_name="Sep")

# Display the first 5 rows of the DataFrame
print(df.head(5))
