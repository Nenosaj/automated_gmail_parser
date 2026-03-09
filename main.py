import time
from config import TARGET_SUBJECT, INPUT_DIR, CHECK_INTERVAL, setup_directories
from email_handler import connect_to_email, get_unread_emails_by_subject, download_excel_attachments, mark_as_read
from excel_processor import extract_data_from_excel
from email_sender import send_html_email

def process_new_emails(mail):
    """Orchestrates finding, downloading, extracting, and forwarding."""
    email_ids = get_unread_emails_by_subject(mail, TARGET_SUBJECT)
    
    if email_ids:
        print(f"\n[SUCCESS] Found {len(email_ids)} new email(s)! Processing...")
        
        for e_id in email_ids:
            # 1. Download the Excel file from the trigger email
            saved_files = download_excel_attachments(mail, e_id, INPUT_DIR)
            
            for filepath in saved_files:
                print(f"[DOWNLOADED] Saved attachment to: {filepath}")
                
                # 2. Extract the data using Pandas
                records, headers = extract_data_from_excel(filepath)
                
                if records and headers:
                    # 3. Generate the HTML table and send the email
                    print("Initiating email forward sequence...")
                    send_html_email(records, headers)
            
            # 4. Mark the trigger email as read so we don't process it again
            mark_as_read(mail, e_id)
            print("[DONE] Pipeline complete. Email marked as read. Waiting for the next one...\n")

def start_pipeline():
    """The main loop that keeps the script running and listening."""
    setup_directories()
    print(f"Waiting for unread emails with subject: '{TARGET_SUBJECT}'...")
    
    while True:
        mail = connect_to_email()
        if mail:
            process_new_emails(mail)
            mail.logout() # Cleanly disconnect after checking
        
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    start_pipeline()