import os
import pandas as pd
from datetime import datetime  

pd.options.mode.copy_on_write = True

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM2 Source compilation began on", str_date_time)

def process_ended():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice DM2 Source compilation completed on", str_date_time)
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
        df = pd.read_excel(file_path)
        dfs.append(df)

    # Concatenate the list of DataFrames into one
    compiled_data = pd.concat(dfs, ignore_index=True)

    # Drop duplicates based on all columns
    compiled_data = compiled_data.drop_duplicates()

    # Convert date columns to datetime dtype
    date_columns = ['Created Date', 'Last Updated Date', 'Payment Date', 'Date Received']
    #clean_data[date_columns] = clean_data[date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))
    compiled_data[date_columns] = compiled_data[date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))

    #clean_data = clean_data.loc[:, date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))

    # Function to find the most recent date among three columns
    def find_latest_date(row):
        return max(row['Created Date'], row['Last Updated Date'], row['Payment Date'], row['Date Received'])

    # Apply the function row-wise to find the latest date
    compiled_data['Latest Record Date'] = compiled_data.apply(find_latest_date, axis=1)
    #clean_data['Latest Record Date'] = clean_data.loc[:, date_columns].apply(find_latest_date, axis=1)

    # Drop duplicates based on all columns
    compiled_data = compiled_data.drop_duplicates()

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice Datamart2 files/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM2 Source.csv'

compile_excel_files(folder_path, output_file)

process_ended()