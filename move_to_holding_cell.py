import os
from datetime import datetime

def eval_timestamp():
    timestamp = datetime.now()
    step_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("step completed", step_time)

def move_files_to_archive(folder_path, archive_path):
    
    used_files = [file for file in os.listdir(folder_path) if file.startswith(('Invoice Datamart'))]
    print("coupa file list retrieved")
    print(used_files)
    eval_timestamp()

    for file in used_files:
        origination = os.path.join(origin_path, file)
        destination = os.path.join(archive_path, file)
        os.rename(origination, destination)
    
    print("files moved")
    eval_timestamp()

# Example usage in TEST folders
origin_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST'
archive_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/holding cell'
#input_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/Invoice DM Source.csv'
#output_file = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/Invoice DM Source copy3.xlsx'


move_files_to_archive(origin_path, archive_path)