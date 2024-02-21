import os
import pandas as pd

def combine_csv_files(folder_path, output_file):
    # Get a list of all CSV files in the folder
    csv_files = [file for file in os.listdir(folder_path) if file.endswith('.csv')]

    # Check if there are any CSV files in the folder
    if not csv_files:
        print("No CSV files found in the specified folder.")
        return

    # Initialize an empty DataFrame to store the combined data
    combined_data = pd.DataFrame()

    # Loop through each CSV file and append its data to the combined_data DataFrame
    for csv_file in csv_files:
        file_path = os.path.join(folder_path, csv_file)
        df = pd.read_csv(file_path)
        combined_data = combined_data.append(df, ignore_index=True)

    # Write the combined data to a new CSV file
    combined_data.to_csv(output_file, index=False)
    print(f"Combined data saved to {output_file}")

# Example usage
folder_path = '/path/to/your/csv/files'
output_file = '/path/to/combined/output.csv'
combine_csv_files(folder_path, output_file)
