import os
import pandas as pd
import glob
from datetime import datetime  

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice LI DM Source compilation began on", str_date_time)

def process_ended():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Invoice LI DM Source compilation completed on", str_date_time)
    print(" ")

def load_latest_excel_files(folder_path, num_files=3):
    # Get the list of all excel files in the folder
    excel_files = glob.glob(f"{folder_path}/*.xlsx")
    
    # Sort the list based on creation time to get the most recent files
    latest_files = sorted(excel_files, key=os.path.getmtime, reverse=True)[:num_files]
    
    # Read the excel files into DataFrames and store them in a list
    dfs = [pd.read_excel(file) for file in latest_files]
    
    return dfs

def compile_excel_files(folder_path, output_file):

    process_begun()

    dfs3 = load_latest_excel_files(folder_path)

    csv_df = pd.read_csv(output_file)

    # Concatenate the list of DataFrames into one
    compiled_data = pd.concat([csv_df] + dfs3, ignore_index=True)

    # Convert date columns to datetime dtype
    date_columns = ['Last Updated Date', 'PO Order Date', 'Invoice Created Date', 'Local Payment Date']
    compiled_data.loc[:, date_columns] = compiled_data.loc[:, date_columns].apply(lambda x: pd.to_datetime(x, format='mixed'))

    compiled_data = compiled_data.drop_duplicates()

    # Function to find the most recent date among three columns
    def find_latest_date(row):
        return max(row['Last Updated Date'], row['PO Order Date'], row['Invoice Created Date'], row['Local Payment Date'])

    # Apply the function row-wise to find the latest date
    compiled_data['Latest Record Date'] = compiled_data.apply(find_latest_date, axis=1)

    #order by latest record date this will ensure latest records show first when manually viewing final source
    compiled_data = compiled_data.sort_values(by='Latest Record Date',ascending=False)

    compiled_data = compiled_data.drop_duplicates(subset=['Invoice ID', 'Line #', 'Latest Record Date'])

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice LI Datamart files/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice LI DM Source.csv'

compile_excel_files(folder_path, output_file)

process_ended()