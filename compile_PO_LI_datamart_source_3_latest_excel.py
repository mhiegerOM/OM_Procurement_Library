import os
import pandas as pd
import glob
from datetime import datetime  

def process_begun():

    # Create timestamp for start of process
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("PO LI DM Source compilation began on", str_date_time)

def process_ended():

    # Create timestamp for end of process
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("PO LI DM Source compilation completed on", str_date_time)
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

    #order by last updated date this will ensure latest records show first when manually viewing final source
    compiled_data = compiled_data.sort_values(by='Last Updated Date', ascending=False)

    #remove equation markings from Active Invoiced Column in datasource
    compiled_data['Active Invoiced'] = compiled_data['Active Invoiced'].str.replace('=', '').str.replace('"', '').astype(float)

    # Drop duplicates based on all columns
    compiled_data = compiled_data.drop_duplicates(subset=['PO ID', 'Line', 'Last Updated Date'])

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/PO LI Datamart files'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/PO LI DM Source.csv'

compile_excel_files(folder_path, output_file)

process_ended()