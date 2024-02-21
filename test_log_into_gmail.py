import imaplib
import email
import yaml
from datetime import datetime

def read_yaml(file_path):
    with open(file_path, 'r') as file:
        return yaml.safe_load(file)

def get_email_body(email_message):
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode(part.get_content_charset())
    else:
        return email_message.get_payload(decode=True).decode(email_message.get_content_charset())

def main():
    # Load credentials from YAML file
    credentials = read_yaml('creds.yaml')
    gmailusername = credentials['username']
    gmailpassword = credentials['password']

    # Connect to Gmail
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(gmailusername, gmailpassword)
    mail.select('inbox')

    # Search for emails with specific criteria
    result, data = mail.search(None, '(UNSEEN SUBJECT "Invoice Datamart - Daily" SINCE "'
                                 + datetime.today().strftime('%d-%b-%Y') + '")')

    # Get the email IDs
    email_ids = data[0].split()

    if email_ids:
        # Fetch the latest email
        latest_email_id = email_ids[-1]
        result, email_data = mail.fetch(latest_email_id, '(RFC822)')
        raw_email = email_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        # Get and print the body of the email
        email_body = get_email_body(email_message)
        print("Email Body:")
        print(email_body)

    mail.close()
    mail.logout()

if __name__ == "__main__":
    main()
