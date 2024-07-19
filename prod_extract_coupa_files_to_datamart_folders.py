import zipfile
import os
from datetime import datetime, timedelta

# Create zipfile archive name, including 2 days back for LI archives extracted over the weekend

today = datetime.now()

year = today.year
month = today.month
day = today.day
yesterday_year = (today - timedelta(days=1)).year
yesterday_month = (today - timedelta(days=1)).month
yesterday_day = (today - timedelta(days=1)).day
twodaysago_year = (today - timedelta(days=2)).year
twodaysago_month = (today - timedelta(days=2)).month
twodaysago_day = (today - timedelta(days=2)).day
threedaysago_year = (today - timedelta(days=3)).year
threedaysago_month = (today - timedelta(days=3)).month
threedaysago_day = (today - timedelta(days=3)).day
fourdaysago_year = (today - timedelta(days=4)).year
fourdaysago_month = (today - timedelta(days=4)).month
fourdaysago_day = (today - timedelta(days=4)).day

archive_file_prefixes = [ f"drive-download-{year:04d}{month:02d}{day:02d}",
f"Invoice LI Datamart - Daily_{year:04d}-{month:02d}-{day:02d}",
f"PO LI Datamart - Daily_{year:04d}-{month:02d}-{day:02d}",
f"Invoice LI Datamart - Daily_{yesterday_year:04d}-{yesterday_month:02d}-{yesterday_day:02d}",
f"PO LI Datamart - Daily_{yesterday_year:04d}-{yesterday_month:02d}-{yesterday_day:02d}",
f"Invoice LI Datamart - Daily_{twodaysago_year:04d}-{twodaysago_month:02d}-{twodaysago_day:02d}",
f"PO LI Datamart - Daily_{twodaysago_year:04d}-{twodaysago_month:02d}-{twodaysago_day:02d}"
]

print(archive_file_prefixes)

# Define the path to the Downloads directory and the ZIP file

downloads_path = os.path.expanduser("~/Downloads")
# destination_archive = os.path.expanduser("~/Desktop/extracttestfile")
destination_archive = os.path.expanduser("C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart")

# Define a function to determine the destination folder based on the Excel file name
def get_destination_folder(file_name):
    if "invoice datamart - abbrev" in file_name.lower():
        return "Invoice Datamart files"
    elif "invoice li datamart - daily" in file_name.lower():
        return "Invoice LI Datamart files"
    elif "invoice pmt details - daily" in file_name.lower():
        return "Invoice Pmt Detail Datamart files"
    elif "new supplier req - daily" in file_name.lower():
        return "New Supplier Req Datamart files"
    elif "payments - daily" in file_name.lower():
        return "Payments Datamart files"
    elif "pmt acct datamart - daily" in file_name.lower():
        return "Pmt Acct Datamart files"
    elif "po datamart - daily" in file_name.lower():
        return "PO Datamart files"
    elif "po li datamart - daily" in file_name.lower():
        return "PO LI Datamart files"
    elif "supplier info datamart - daily" in file_name.lower():
        return "Supplier Info Datamart files"
    elif "suppliers datamart - daily" in file_name.lower():
        return "Supplier Datamart files"
    # Add more conditions as needed
    else:
        return "Others"

# Find the ZIP files that start with any of the prefixes
found_files = []
for file_name in os.listdir(downloads_path):
    if any(file_name.startswith(prefix) for prefix in archive_file_prefixes) and file_name.endswith(".zip"):
        found_files.append(file_name)

if not found_files:
    raise FileNotFoundError(f"No ZIP files found starting with any of the prefixes: {archive_file_prefixes}")

# Open and extract each found ZIP file
for zip_file in found_files:
    zip_file_path = os.path.join(downloads_path, zip_file)
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith(".xlsx") or file.endswith(".xls"):
                # Determine the destination folder
                destination_folder = get_destination_folder(file)
                destination_path = os.path.join(destination_archive, destination_folder)
                
                # Create the destination folder if it doesn't exist
                os.makedirs(destination_path, exist_ok=True)
                
                # Extract the file to the destination folder
                zip_ref.extract(file, destination_path)
                print(f"Extracted {file} to {destination_path}")
