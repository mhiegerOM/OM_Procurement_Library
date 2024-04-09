import os
import pandas as pd
from datetime import datetime

pd.options.mode.copy_on_write = True

def eval_timestamp():
    timestamp = datetime.now()
    step_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("step completed", step_time)

def copy_source_to_excel_file(input_file, output_file):

   # file_path = os.path(input_file)
    df = pd.read_csv(input_file, dtype=Inv_dtype_dict)
    print("dataframe read from csv")
    eval_timestamp()

    date_columns = ['Net Due Date', 'Invoice Date', 'Created Date', 'Last Updated Date', 'Payment Date', 'Date Received', 'Last Exported At', 'Latest Record Date']

    df[date_columns] = df[date_columns].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))
    print("dates converted")
    eval_timestamp()    

    df = df.to_excel(output_file, index=False)
    print(f"Compiled data saved to {output_file}")
    eval_timestamp()


# Example usage
#folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice Datamart files/'
#output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source.csv'

# Example usage in TEST folders
##folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/'
input_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source.csv'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/Invoice DM Source copy3.xlsx'

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

copy_source_to_excel_file(input_file, output_file)