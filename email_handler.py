import os.path
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from config import CLIENT_SECRET_FILE, TOKEN_FILE, SCOPES

def get_gmail_service():
    """Authenticates the user and returns the Gmail API service object."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def get_unread_emails_by_subject(service, subject):
    """Searches for unread emails with a specific subject and returns them."""
    query = f'is:unread subject:"{subject}"'
    results = service.users().messages().list(userId='me', q=query).execute()
    return results.get('messages', [])

def download_excel_attachments(service, msg_id, output_dir):
    """Fetches an email, finds .xlsx attachments, and saves them locally."""
    message = service.users().messages().get(userId='me', id=msg_id).execute()
    downloaded_files = []
    
    parts = message['payload'].get('parts', [])
    for part in parts:
        if part.get('filename') and part.get('filename').endswith('.xlsx'):
            attachment_id = part['body']['attachmentId']
            attachment = service.users().messages().attachments().get(
                userId='me', messageId=msg_id, id=attachment_id).execute()
            
            data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            filepath = os.path.join(output_dir, part['filename'])
            with open(filepath, 'wb') as f:
                f.write(data)
            downloaded_files.append(filepath)
            
    return downloaded_files