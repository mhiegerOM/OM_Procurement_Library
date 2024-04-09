import pandas as pd
import os

# Directory containing the Excel files
directory = 'path_to_directory'

# Get a list of all Excel files in the directory
excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

# Sort the Excel files by size (largest first)
excel_files.sort(key=lambda x: os.path.getsize(os.path.join(directory, x)), reverse=True)

# Initialize the base dataframe
base_df = pd.DataFrame()

# Read the first Excel file into the base dataframe
first_file = True

# Iterate over each Excel file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):  # assuming all files are .xlsx format
        filepath = os.path.join(directory, filename)
        
        # Read the Excel file into a temporary dataframe
        temp_df = pd.read_excel(filepath)
        
        # Remove rows from temp_df that are duplicates with base_df
        if not first_file:
            temp_df = temp_df[~temp_df.isin(base_df)].dropna()
        
        # Append the temporary dataframe to the base dataframe
        base_df = pd.concat([base_df, temp_df], ignore_index=True)
        
        if first_file:
            first_file = False

# Drop duplicate rows from the base dataframe
base_df.drop_duplicates(inplace=True)

# Write the base dataframe to a CSV file
base_df.to_csv('output.csv', index=False)

print("Processing completed!")
