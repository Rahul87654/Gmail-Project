import os
import base64
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd

# Define the scopes required for accessing Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate():
    """
    Authenticate and authorize access to Gmail API.
    """
    creds = None
    
    # Check if token.json file exists
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no valid credentials or if they've expired, obtain new credentials
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Specify the path to your credentials.json file
            credentials_path = os.path.join(os.path.dirname(__file__), 'credentials.json')
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(f"Credentials file '{credentials_path}' not found.")
            
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def read_latest_email():
    try:
        # Authenticate and build the Gmail service
        creds = authenticate()
        service = build('gmail', 'v1', credentials=creds)
        # Call the Gmail API to retrieve the latest email in the Inbox
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=1).execute()
        messages = results.get('messages', [])
        if messages:
            latest_message_id = messages[0]['id']
            msg = service.users().messages().get(userId='me', id=latest_message_id).execute()
            # Extract email details
            headers = msg['payload']['headers']
            sender = next((header['value'] for header in headers if header['name'] == 'From'), '')
            cc = next((header['value'] for header in headers if header['name'] == 'Cc'), '')
            bcc = next((header['value'] for header in headers if header['name'] == 'Bcc'), '')
            # Extract timestamp (date and time)
            timestamp = next((header['value'] for header in headers if header['name'] == 'Date'), '')
            # Extract subject and body
            subject = next((header['value'] for header in headers if header['name'] == 'Subject'), '')
            body = ''
            parts = msg['payload'].get('parts', [])
            for part in parts:
                if part['mimeType'] == 'text/plain' or part['mimeType'] == 'text/html':
                    body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                    break
            # Print email details
            print(f"Sender: {sender}")
            print(f"CC: {cc}")
            print(f"BCC: {bcc}")
            print(f"Timestamp: {timestamp}")
            print(f"Subject: {subject}")
           
            # Extract email body
            email_body = ''
            if 'parts' in msg['payload']:
                for part in msg['payload']['parts']:
                    if part['mimeType'] == 'text/plain':
                        email_body += base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
            
            # Extract all file paths from the email body using regex
            file_paths = re.findall(r'PALSFTPHOME/[\w-]+/[\w.-]+', email_body)
            
            # Load Excel data from the Excel file
            excel_file_path = os.path.join(os.path.dirname(__file__), 'File.xlsx')
            try:
                excel_data = pd.read_excel(excel_file_path)
                excel_loaded = True
            except FileNotFoundError:
                excel_data = None
                excel_loaded = False
            
            print("Excel File Loaded:", "Yes" if excel_loaded else "No")
            
            if excel_data is not None and not excel_data.empty:
                # Check each file path against the 'File_path' column in Excel data
                for file_path in file_paths:
                    matches = excel_data[excel_data['File_path'].apply(lambda x: file_path.startswith(x))]
                    if not matches.empty:
                        client_name = matches.iloc[0]['Client_name']
                        print(f"File path '{file_path}' matches. Client: {client_name}")
                    else:
                        print(f"No matching client found for file path '{file_path}'")
            else:
                print("Excel data is empty or could not be loaded.")
        else:
            print("No emails found in the Inbox.")
    
    except HttpError as error:
        print(f"An error occurred: {error}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    # Read the latest email and process file paths against Excel data
    read_latest_email()
