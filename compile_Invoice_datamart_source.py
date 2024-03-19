import os
import pandas as pd
from datetime import datetime  

pd.options.mode.copy_on_write = True

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM Source compilation began on", str_date_time)

def process_ended():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM Source compilation completed on", str_date_time)
    print(" ")

def compile_excel_files(folder_path, output_file):

    process_begun()

    # Get a list of all Excel files in the folder
    excel_files = [file for file in os.listdir(folder_path) if file.endswith(('.xls', '.xlsx'))]

    # Check if there are any Excel files in the folder
    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Initialize an empty list to store DataFrames
    dfs = []

    # Loop through each Excel file and append its data to the list
    for excel_file in excel_files:
        file_path = os.path.join(folder_path, excel_file)
        df = pd.read_excel(file_path) #, dtype=Inv_dtype_dict)
        dfs.append(df)

    # Concatenate the list of DataFrames into one
    compiled_data = pd.concat(dfs, ignore_index=True)

    # Convert date columns to datetime dtype

    date_columns = ['Last Updated Date', 'Payment Date', 'Date Received']

    compiled_data[date_columns] = compiled_data[date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))

    # Function to find the most recent date among three columns
    def find_latest_date(row):
        #return max(row['Created Date'], row['Last Updated Date'], row['Payment Date'], row['Date Received'])
        return max(row['Last Updated Date'], row['Payment Date'], row['Date Received'])

    # Apply the function row-wise to find the latest date
    compiled_data['Latest Record Date'] = compiled_data.apply(find_latest_date, axis=1)
    #clean_data['Latest Record Date'] = clean_data.loc[:, date_columns].apply(find_latest_date, axis=1)

    #order by latest record date this will ensure latest records show first when manually viewing final source
    compiled_data = compiled_data.sort_values(by= 'Latest Record Date', ascending=False)

    # Drop duplicates based on all columns
    compiled_data = compiled_data.drop_duplicates(subset=['Invoice ID', 'Latest Record Date'])

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice Datamart files/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source.csv'

# Inv_dtype_dict = {'Invoice #': 'str',
# 'Supplier': 'str',
# 'Net Due Date': 'str',
# 'Total': 'float',
# 'Status': 'str',
# 'Delivery Method': 'str',
# 'Dispute Reason': 'str',
# 'Invoice Date': 'str',
# 'Invoice ID': 'str',
# 'Last Updated Date': 'str',
# 'Last Updated By': 'str',
# 'Payment Date': 'str',
# 'Payment Term': 'str',
# 'PO Number': 'str',
# 'PO ID': 'str',
# 'Requester': 'str',
# 'Supplier #': 'str',
# 'Chart Of Accounts': 'str',
# 'Created By': 'str',
# 'Created Date': 'str',
# 'Credit Applied': 'str',
# 'Current Approver': 'str',
# 'Date Received': 'str',
# 'Exported?': 'str',
# 'Last Exported At': 'str',
# 'Line Tags': 'str',
# 'Line Tolerance Failures': 'str',
# 'Original Invoice Number': 'str',
# 'Paid': 'str',
# 'Payment Channel': 'str',
# 'Payment Notes': 'str',
# 'Payment Num''s': 'str',
# 'Remit To Code': 'str',
# 'Resolved Line Tolerance Failures': 'str',
# 'Resolved Tolerance Failures': 'str',
# 'Source': 'str',
# 'Supplier Default Commodity': 'str',
# 'Supplier Total': 'float',
# 'Tags': 'str',
# 'Time with Current Approver': 'str',
# 'Tolerance Failures': 'str',
# 'Unanswered Comments': 'str',
# 'Latest Record Date': 'str'}

compile_excel_files(folder_path, output_file)

process_ended()