import os
import pandas as pd
from datetime import datetime
import glob

pd.options.mode.chained_assignment = None  # Suppress pandas warnings

def process_begun():
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM Source compilation began on", str_date_time)

def process_ended():
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM Source compilation completed on", str_date_time, "\n")

def compile_excel_files(existing_file, folder_path, output_file):
    process_begun()

    # Read existing file into DataFrame
    DM_df = pd.read_csv(existing_file, dtype=Inv_dtype_dict)

    # Save DataFrame to temporary Excel file
    temp_excel_file = os.path.join(folder_path, 'invoice_DM_temp.xlsx')
    DM_df.to_excel(temp_excel_file, index=False)

    # Get list of last 4 Excel files in folder
    excel_files = sorted(glob.glob(os.path.join(folder_path, "*.xlsx")), key=os.path.getmtime, reverse=True)[:4]

    # Check if any Excel files found
    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Read Excel files into DataFrame and concatenate
    dfs = [pd.read_excel(file, dtype=Inv_dtype_dict) for file in excel_files]
    compiled_data = pd.concat(dfs, ignore_index=True)

    # Convert date columns to datetime dtype
    date_columns = ['Last Updated Date', 'Payment Date', 'Date Received']
    compiled_data[date_columns] = compiled_data[date_columns].apply(pd.to_datetime, format='%m/%d/%y')

    # Find latest date among columns
    compiled_data['Latest Record Date'] = compiled_data[date_columns].max(axis=1)

    # Sort DataFrame by latest record date
    compiled_data = compiled_data.sort_values(by='Latest Record Date', ascending=False)

    # Write compiled data to CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

def remove_dupes_from_csv_source(output_file, output_file2):
    # Read output from compilation
    ddf = pd.read_csv(output_file)

    # Remove duplicates from compilation
    ddf.drop_duplicates(subset=['Invoice ID', 'Latest Record Date'], inplace=True)

    # Write results to new CSV
    ddf.to_csv(output_file2, index=False)

# Define file paths
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice Datamart files/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source dedup temp.csv'
output_file2 = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source dedup TEST.csv'
existing_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source.csv'

# Dictionary specifying column data types
Inv_dtype_dict = {
    'Invoice #': 'str', 'Supplier': 'str', 'Net Due Date': 'str', 'Total': 'float', 'Status': 'str',
    'Delivery Method': 'str', 'Dispute Reason': 'str', 'Invoice Date': 'str', 'Invoice ID': 'str',
    'Last Updated Date': 'str', 'Last Updated By': 'str', 'Payment Date': 'str', 'Payment Term': 'str',
    'PO Number': 'str', 'PO ID': 'str', 'Requester': 'str', 'Supplier #': 'str', 'Chart Of Accounts': 'str',
    'Created By': 'str', 'Created Date': 'str', 'Credit Applied': 'str', 'Current Approver': 'str',
    'Date Received': 'str', 'Exported?': 'str', 'Last Exported At': 'str', 'Line Tags': 'str',
    'Line Tolerance Failures': 'str', 'Original Invoice Number': 'str', 'Paid': 'str', 'Payment Channel': 'str',
    'Payment Notes': 'str', 'Payment Num''s': 'str', 'Remit To Code': 'str',
    'Resolved Line Tolerance Failures': 'str', 'Resolved Tolerance Failures': 'str', 'Source': 'str',
    'Supplier Default Commodity': 'str', 'Supplier Total': 'float', 'Tags': 'str',
    'Time with Current Approver': 'str', 'Tolerance Failures': 'str', 'Unanswered Comments': 'str',
    'Latest Record Date': 'str'
}

# Perform operations
compile_excel_files(existing_file, folder_path, output_file)
remove_dupes_from_csv_source(output_file2, output_file2)
process_ended()
