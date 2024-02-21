import os
import pandas as pd
from datetime import datetime  
#from numpy import row_stack

def add_latest_record_date_to_datasource(folder_path, output_file):

    # Initialize an empty list to store DataFrames
    df = []

    file_path = os.path.join(folder_path, 'C:/Users/MatthewHieger/Desktop/testing invoice dates/invoice dates baseline.xlsx')
    df = pd.read_excel(file_path)

    # Drop duplicates based on all columns
    clean_df = df.drop_duplicates()

    # Convert date columns to datetime dtype
    date_columns = ['Last Updated Date', 'Payment Date', 'Date Received']
    clean_df[date_columns] = df[date_columns].apply(pd.to_datetime)

    # Function to find the most recent date among three columns
    def find_latest_date(row):
        return max(row['Last Updated Date'], row['Payment Date'], row['Date Received'])

    # Apply the function row-wise to find the latest date
    clean_df['Latest Record Date'] = clean_df.apply(find_latest_date, axis=1)

    # Print the updated DataFrame
    #print(df)

    # Write the appended data to a new CSV file
    clean_df.to_csv(output_file, index=False)
    print(f"basline saved to {output_file}")

# Example usage
folder_path = 'C:/Users/MatthewHieger/Desktop/testing invoice dates'
output_file = 'C:/Users/MatthewHieger/Desktop/testing invoice dates/invoice dates baseline.csv'

add_latest_record_date_to_datasource(folder_path, output_file)