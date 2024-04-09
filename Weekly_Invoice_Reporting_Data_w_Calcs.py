import os
import pandas as pd
from datetime import datetime, timedelta

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

    #format Payment Date to dtype date
    df['Payment Date'] = df['Payment Date'].apply(lambda x: pd.to_datetime(x, format='%Y-%m-%d'))

    #create column with payment month
    df['Payment MMYY'] = df['Payment Date'].apply(lambda x: pd.to_datetime(x, format='%m/%y'))

    #format Order Date to dtype date
    df['Net Due Date'] = df['Net Due Date'].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))
    
    monday_of_week = datetime.now() - timedelta(days=datetime.now().weekday())

    #define calculation for Invoice age
    def find_PO_Age(row):
        return (monday_of_week - row['Net Due Date']).days

    #create column with Invoice age in days
    df['Aging Based on Net Due Date'] = df.apply(find_PO_Age, axis=1)

    #define aging buckets
    def assign_aging_bucket(row):
        if row['Aging Based on Net Due Date'] <= 0:
            return '(1) Current'
        elif row['Aging Based on Net Due Date'] <= -10:
            return '(2) 10 Day'
        elif row['Aging Based on Net Due Date'] <= 30:
            return '(3) 1 to 30'
        elif row['Aging Based on Net Due Date'] <= 90:
            return '(4) 31 to 90'
        elif row['Aging Based on Net Due Date'] <= 180:
            return '(5) 91 to 180'
        elif row['Aging Based on Net Due Date'] <= 270:
            return '(6) 181 to 270'
        elif row['Aging Based on Net Due Date'] <= 365:
            return '(7) 271 to 365'
        else:
            return '(8) Over 365'

    #create column with PO aging buckets
    df['Aging Bucket Based on Net Due Date'] = df.apply(assign_aging_bucket, axis=1)

    def assign_approval_bucket(row):
        if row['Total'] <= 0:
            return 'Credit'
        elif row['Total'] <= 5000:
            return '$0 to $5,000'
        elif row['Total'] <= 25000:
            return '$5,001 to $25,000'
        elif row['Total'] <= 50000:
            return '$25,001 to $50,000'
        elif row['Total'] <= 500000:
            return '$50,001 to $500,000'
        elif row['Total'] <= 5000000:
            return '$500,001 to $5,000,000'
        else:
            return '> $5,000,000'
    
    #create column with Invoice Total Approval buckets
    df['Approval Bucket'] = df.apply(assign_approval_bucket, axis=1)

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
        'Payment Date',
        'Payment MMYY',
        'Aging Based on Net Due Date',
        'Aging Bucket Based on Net Due Date',
        'Approval Bucket'
        ], 
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