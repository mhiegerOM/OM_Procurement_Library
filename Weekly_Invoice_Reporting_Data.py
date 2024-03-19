import os
import pandas as pd
from datetime import datetime  

pd.options.mode.copy_on_write = True

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Weekly Invoice Dataset Set Compilation began on", str_date_time)

def process_ended():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Weekly Invoice Dataset Set Compilation completed on", str_date_time)
    print(" ")

def dedupe_csv(folder_path, output_file):

    process_begun()

    #file_path = os.path.join(folder_path, 'Invoice DM TEST.csv')
    file_path = os.path.join(folder_path, 'Invoice DM Source.csv')
    df = pd.read_csv(file_path, dtype=Inv_dtype_dict)

    df = df.sort_values(by= 'Latest Record Date', ascending=False)

    #df = df.drop_duplicates(subset=['Invoice #','Supplier #'])
    df = df.drop_duplicates(subset=['Invoice ID'])


    # Write the compiled data to a new CSV file
    df.to_csv(output_file, columns=['Invoice #', 
        'PO Number', 
        'Requester', 
        'Supplier',
        'Created Date', 
        'Net Due Date', 
        'Total', 
        'Status', 
        'Delivery Method', 
        'Created By', 
        'Current Approver', 
        'Time with Current Approver', 
        'Paid', 
        'Payment Date'], 
        index=False)
    print(f"Weekly Invoice dataframe saved to {output_file}")

# Example usage
    
Inv_dtype_dict = {'Invoice #': 'str',
'Supplier': 'str',
'Net Due Date': 'str',
'Total': 'float',
'Status': 'str',
'Delivery Method': 'str',
'Dispute Reason': 'str',
'Invoice Date': 'str',
'Invoice ID': 'str',
'Last Updated Date': 'str',
'Last Updated By': 'str',
'Payment Date': 'str',
'Payment Term': 'str',
'PO Number': 'str',
'PO ID': 'str',
'Requester': 'str',
'Supplier #': 'str',
'Chart Of Accounts': 'str',
'Created By': 'str',
'Created Date': 'str',
'Credit Applied': 'str',
'Current Approver': 'str',
'Date Received': 'str',
'Exported?': 'str',
'Last Exported At': 'str',
'Line Tags': 'str',
'Line Tolerance Failures': 'str',
'Original Invoice Number': 'str',
'Paid': 'str',
'Payment Channel': 'str',
'Payment Notes': 'str',
'Payment Num''s': 'str',
'Remit To Code': 'str',
'Resolved Line Tolerance Failures': 'str',
'Resolved Tolerance Failures': 'str',
'Source': 'str',
'Supplier Default Commodity': 'str',
'Supplier Total': 'float',
'Tags': 'str',
'Time with Current Approver': 'str',
'Tolerance Failures': 'str',
'Unanswered Comments': 'str',
'Latest Record Date': 'str'}

#folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/'
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Weekly Invoice Report Data.csv'

dedupe_csv(folder_path, output_file)

process_ended()