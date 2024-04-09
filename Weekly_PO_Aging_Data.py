import os
import pandas as pd
from datetime import datetime, timedelta

pd.options.mode.copy_on_write = True

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Weekly PO LI Dataset Set Compilation began on", str_date_time)

def process_ended():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("Weekly PO LI Dataset Set Compilation completed on", str_date_time)
    print(" ")

def dedupe_csv(folder_path, output_file):

    process_begun()

    #file_path = os.path.join(folder_path, 'Invoice DM TEST.csv')
    file_path = os.path.join(folder_path, 'PO LI DM Source.csv')
    df = pd.read_csv(file_path, dtype=POLI_dtype_dict)

    df = df.sort_values(by= 'Last Updated Date', ascending=False)

    #df = df.drop_duplicates(subset=['Invoice #','Supplier #'])
    df = df.drop_duplicates(subset=['PO ID', 'Line'])

    #clean up Active Invoiced equation strings from coupa extract if not already done
    df['Active Invoiced'] = df['Active Invoiced'].str.replace('=', '').str.replace('"', '').astype(float)
    
    #translate invoiced ammounts into bolean statuses
    df['Received'] = df['Received'].apply(lambda x: 'Received' if x > 0 else 'Not received')
    df['Invoiced'] = df['Active Invoiced'].apply(lambda x: 'Invoiced' if x > 0 else 'Not Invoiced')
    df['Approved'] = df['Approved Invoiced'].apply(lambda x: 'Approved' if x > 0 else 'Not Approved')
    
    #format Order Date to dtype date
    df['Order Date (Header)'] = df['Order Date (Header)'].apply(lambda x: pd.to_datetime(x, format='%m/%d/%y'))

    monday_of_week = datetime.now() - timedelta(days=datetime.now().weekday())

    #define calculation for PO age
    def find_PO_Age(row):
        return (monday_of_week - row['Order Date (Header)']).days

    #create column with PO age in days
    df['PO Days Aging'] = df.apply(find_PO_Age, axis=1)

    #define aging buckets
    def assign_aging_bucket(row):
        if row['PO Days Aging'] <= 30:
            return '(1) 0 to 30'
        elif row['PO Days Aging'] <= 60:
            return '(2) 31 to 60'
        elif row['PO Days Aging'] <= 90:
            return '(3) 61 to 90'
        elif row['PO Days Aging'] <= 120:
            return '(4) 91 to 120'
        else:
            return '(5) Over 120+'

    #create column with PO aging buckets
    df['Aging Bucket'] = df.apply(assign_aging_bucket, axis=1)

    # Write the compiled data to a new CSV file
    df.to_csv(output_file, columns=['PO Number (Header)',
        'Supplier',
        'Line',
        'Chart Of Accounts',
        'Expected Delivery Date',
        'Order Date (Header)',
        'Services Start Date',
        'Order Status (Header)',
        'Services End Date',
        'Order Transmission Status (Header)',
        'Item #',
        'UOM',
        'Item',
        'Line Total',
        # 'Received',
        # 'Active Invoiced',
        # 'Approved Invoiced',
        'Commodity',
        'Requested By (Header)',
        'Type',
        'Received',
        'Invoiced',
        'Approved',
        'PO Days Aging',
        'Aging Bucket'
        # 'Expected Delivery Month',
        # 'Sevices End Month',
        # 'Order Month']
        ], index=False)
    print(f"Weekly PO LI dataframe saved to {output_file}")

# Example usage
    
POLI_dtype_dict = {'Acknowledged At': 'str',
'Active Invoiced': 'str',
'Approved Invoiced': 'float',
'Change Type': 'str',
'Chart Of Accounts': 'str',
'Commodity': 'str',
'Contract': 'str',
'Contract Number': 'str',
'Created By': 'str',
'Department': 'str',
'Expected Delivery Date': 'str',
'Inventory Item': 'str',
'Item': 'str',
'Item #': 'str',
'Item Name': 'str',
'Last Updated By': 'str',
'Last Updated Date': 'str',
'Line': 'str',
'Line Owner': 'str',
'Line Status': 'str',
'Line Total': 'float',
'Need By': 'str',
'Order Date (Header)': 'str',
'Order Status (Header)': 'str',
'Order Total (Header)': 'float',
'Order Transmission Status (Header)': 'str',
'P-Card Expiration (Header)': 'str',
'P-Card Number (Header)': 'str',
'Payment Term': 'str',
'Pending PO Change (Header)': 'str',
'PO ID': 'str',
'PO Number (Header)': 'str',
'Price': 'float',
'Qty': 'str',
'Receipt Approval Required': 'str',
'Received': 'float',
'Receiving Warehouse': 'str',
'Reporting Total': 'float',
'Req. Need By Date': 'str',
'Req. Submitted Date': 'str',
'Requested By (Header)': 'str',
'Revision Date': 'str',
'Savings (%)': 'str',
'Services End Date': 'str',
'Services Start Date': 'str',
'Site Account Number (Header)': 'str',
'Supplier': 'str',
'Supplier #': 'str',
'Supplier id': 'str',
'Supplier Part Number': 'str',
'Type': 'str',
'UOM': 'str',
'Version': 'str'
}

#folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/'
folder_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/'
output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Weekly PO LI Data.csv'

dedupe_csv(folder_path, output_file)

process_ended()