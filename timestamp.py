from datetime import datetime  

def process_begun():

    timestamp = datetime.now()
    str_date_time = timestamp.strftime("%Y-%m-%d @ %H:%M:%S")
    print("[Process] process began on", str_date_time)

process_begun()