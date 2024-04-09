import os
import pandas as pd
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

def eval_timestamp():
    timestamp = datetime.now()
    step_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("step completed", step_time)

def compile_excel_files(folder_path, output_file):

    process_begun()

        # Get a list of all Excel files in the directory
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    # Sort the Excel files by size (largest first)
    excel_files.sort(key=lambda x: os.path.getsize(os.path.join(folder_path, x)), reverse=True)

    print("excel files sorted largest to smallest")
    eval_timestamp()

    # Initialize the base dataframe
    compiled_data = pd.DataFrame()
    print("compiled_data df created")
    eval_timestamp()

    # Read the first Excel file into the base dataframe
    first_file = True

    # Iterate over each Excel file in the directory
    for filename in excel_files:
        filepath = os.path.join(folder_path, filename)
        
        # Read the Excel file into a temporary dataframe
        temp_df = pd.read_excel(filepath)
        print("1st df completed")
        eval_timestamp()
        
        # Remove rows from temp_df that are duplicates with compiled_data
        if not first_file:
            temp_df = temp_df[~temp_df.isin(compiled_data)].dropna()
            print("temp df evaluated")
            eval_timestamp()
        
        # Append the temporary dataframe to the base dataframe
        compiled_data = pd.concat([compiled_data, temp_df], ignore_index=True)
        print("df appended")
        eval_timestamp()

        if first_file:
            first_file = False

    # Convert date columns to datetime dtype
    date_columns = ['Last Updated Date', 'PO Order Date', 'Invoice Created Date', 'Local Payment Date']
    compiled_data.loc[:, date_columns] = compiled_data.loc[:, date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))

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
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/Final Result/Invoice DM LI TEST Source.csv'

compile_excel_files(folder_path, output_file)

process_ended()