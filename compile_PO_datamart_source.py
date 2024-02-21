import os
import pandas as pd
from datetime import datetime  

def process_begun():

    # Create timestamp for start of process
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("PO DM Source compilation began on", str_date_time)

def process_ended():

    # Create timestamp for end of process
    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("PO DM Source compilation completed on", str_date_time)
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

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/PO Datamart files'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/PO DM Source.csv'

compile_excel_files(folder_path, output_file)

process_ended()