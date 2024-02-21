import imaplib
import email
import os
import zipfile
import pandas as pd
from email.header import decode_header
from datetime import datetime, timedelta
import yaml

def load_credentials(yaml_file):
    with open(yaml_file, "r") as stream:
        try:
            credentials = yaml.safe_load(stream)
            return credentials["gmail_credentials"]["username"], credentials["gmail_credentials"]["password"]
        except yaml.YAMLError as exc:
            print(exc)

def download_and_process_attachment(username, password, subject_keyword, extract_folder):
    # Connect to Gmail IMAP server
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)

    # Select the mailbox you want to work with (e.g., "inbox")
    mail.select("inbox")

    # Get today's date
    today_date = datetime.now().strftime("%d-%b-%Y")

    # Search for emails with the specified subject and received today
    search_criteria = f'SINCE "{today_date}" SUBJECT "{subject_keyword}"'
    status, messages = mail.search(None, search_criteria)
    message_ids = messages[0].split()

    for message_id in message_ids:
        # Fetch the email by ID
        status, msg_data = mail.fetch(message_id, "(RFC822)")
        raw_email = msg_data[0][1]

        # Parse the raw email content
        msg = email.message_from_bytes(raw_email)
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")

        print(f"Processing email with subject: {subject}")

        # Iterate through email parts
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            # Check if the part is an attachment and is a .zip file
            filename = part.get_filename()
            if filename is not None and filename.endswith(".zip"):
                # Save the attachment to a temporary path
                temp_zip_path = "temp_attachment.zip"
                with open(temp_zip_path, "wb") as f:
                    f.write(part.get_payload(decode=True))
                
                # Extract the contents of the .zip file to the specified folder
                with zipfile.ZipFile(temp_zip_path, "r") as zip_ref:
                    zip_ref.extractall(extract_folder)

                # Remove the temporary .zip file
                os.remove(temp_zip_path)

                # Process the extracted Excel file
                excel_files = [f for f in os.listdir(extract_folder) if f.startswith("Invoice Datamart - Daily2_") and f.endswith(".xlsx")]

                if excel_files:
                    for excel_file in excel_files:
                        excel_file_path = os.path.join(extract_folder, excel_file)

                        # Load the Excel file into a DataFrame
                        df = pd.read_excel(excel_file_path)

                # Add a column "As_of_Date" and populate it with today's date
                df["As_of_Date"] = today_date

                # Save the updated Excel file
                df.to_excel(excel_file_path, index=False)
                
                print(f"Excel file updated and saved to: {excel_file_path}")

    # Logout and close the connection
    mail.logout()

# Replace this with the path to your YAML file
yaml_file = "C:/Users/MatthewHieger/Desktop/Procurement_Local_Library/creds.yaml"

# Load credentials from the YAML file
username, password = load_credentials(yaml_file)

# Replace these with other parameters
subject_keyword = "Coupa Report: Invoice Datamart - Daily2"
extract_folder = "C:/Users/MatthewHieger/Desktop/gmail extraction test folder"

# Call the function to download and process the .zip attachment
download_and_process_attachment(username, password, subject_keyword, extract_folder)
