import time
from config import TARGET_SUBJECT, INPUT_DIR, CHECK_INTERVAL, setup_directories
from email_handler import get_gmail_service, get_unread_emails_by_subject, download_excel_attachments
from excel_processor import extract_data_from_excel
from email_sender import send_html_email

def process_new_emails(service):
    """Orchestrates finding, downloading, extracting, and forwarding."""
    messages = get_unread_emails_by_subject(service, TARGET_SUBJECT)
    
    if messages:
        print(f"\n[SUCCESS] Found {len(messages)} new email(s)! Processing...")
        
        for msg_info in messages:
            msg_id = msg_info['id']
            
            # 1. Download the Excel file from the trigger email
            saved_files = download_excel_attachments(service, msg_id, INPUT_DIR)
            
            for filepath in saved_files:
                print(f"[DOWNLOADED] Saved attachment to: {filepath}")
                
                # 2. Extract the data using Pandas
                records, headers = extract_data_from_excel(filepath)
                
                if records and headers:
                    # 3. Generate the HTML table and send the email
                    print("Initiating email forward sequence...")
                    send_html_email(service, records, headers)
            
            # 4. Remove the UNREAD label so we don't process it again
            service.users().messages().modify(
                userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}
            ).execute()
            print("[DONE] Pipeline complete. Email marked as read. Waiting for the next one...\n")

def start_pipeline():
    """The main loop that keeps the script running and listening."""
    setup_directories()
    
    print("Authenticating with Google Workspace...")
    service = get_gmail_service()
    
    print(f"Waiting for unread emails with subject: '{TARGET_SUBJECT}'...")
    
    while True:
        try:
            process_new_emails(service)
        except Exception as e:
            print(f"[ERROR] Pipeline encountered an issue: {e}")
            print("Attempting to refresh connection...")
            service = get_gmail_service()
            
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    start_pipeline()