import imaplib
import email
import os
from config import EMAIL_ACCOUNT, APP_PASSWORD, IMAP_SERVER

def connect_to_email():
    """Establishes and returns a connection to the IMAP server."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ACCOUNT, APP_PASSWORD)
        return mail
    except Exception as e:
        print(f"Connection failed: {e}")
        return None

def get_unread_emails_by_subject(mail, subject):
    """Searches for unread emails with a specific subject and returns their IDs."""
    mail.select('inbox')
    search_criteria = f'(UNSEEN SUBJECT "{subject}")'
    status, messages = mail.search(None, search_criteria)
    
    if status == 'OK' and messages[0]:
        return messages[0].split()
    return []

def download_excel_attachments(mail, email_id, output_dir):
    """Fetches an email, finds .xlsx attachments, and saves them locally."""
    res, msg_data = mail.fetch(email_id, '(RFC822)')
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    
    downloaded_files = []
    
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
            continue
            
        filename = part.get_filename()
        if filename and filename.endswith('.xlsx'):
            filepath = os.path.join(output_dir, filename)
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            downloaded_files.append(filepath)
            
    return downloaded_files

def mark_as_read(mail, email_id):
    """Marks a specific email as read so it isn't processed again."""
    mail.store(email_id, '+FLAGS', '\\Seen')