import os
import pandas as pd

def compile_excel_files(folder_path, output_file):
    # Get a list of all Excel files in the folder
    excel_files = [file for file in os.listdir(folder_path) if file.endswith(('.xls', '.xlsx'))]

    # Check if there are any Excel files in the folder
    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Initialize an empty DataFrame to store the compiled data
    compiled_data = pd.DataFrame()

    # Loop through each Excel file and append its data to the compiled_data DataFrame
    for excel_file in excel_files:
        file_path = os.path.join(folder_path, excel_file)
        df = pd.read_excel(file_path)
       ## compiled_data = compiled_data.append(df, ignore_index=True)
        compiled_data = compiled_data.set_index(keys={'Invoice','Supplier','Last Updated Date'},append=True)

    # Write the compiled data to a new CSV file
    compiled_data.to_csv(output_file, index=False)
    print(f"Compiled data saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/Invoice DM Source.csv'
compile_excel_files(folder_path, output_file)
