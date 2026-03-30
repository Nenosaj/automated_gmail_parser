import time
import sys
import os
import argparse

# CRITICAL: Forces the script to run in its own folder (Required for Task Scheduler)
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from config import EMAIL_RULES, INPUT_DIR, CHECK_INTERVAL, setup_directories
from email_handler import get_gmail_service, get_unread_emails_by_subject, download_excel_attachments
from excel_processor import extract_data_from_excel
from email_sender import send_html_email

def process_single_rule(service, subject, recipients):
    """Processes only one specific rule, then stops and cleans up."""
    print(f"Checking for emails matching: '{subject}'...")
    messages = get_unread_emails_by_subject(service, subject)
    
    if messages:
        print(f"[SUCCESS] Found {len(messages)} new email(s) for '{subject}'! Processing...")
        for msg_info in messages:
            msg_id = msg_info['id']
            saved_files = download_excel_attachments(service, msg_id, INPUT_DIR)
            
            for filepath in saved_files:
                print(f"[DOWNLOADED] Saved attachment to: {filepath}")
                records, headers = extract_data_from_excel(filepath)
                
                if records and headers:
                    print("Initiating email forward sequence...")
                    send_html_email(service, records, headers, subject, recipients)
                    
                    # --- AUTO-CLEANUP STEP ---
                    try:
                        print(f"[CLEANUP] Deleting processed file: {filepath}")
                        os.remove(filepath)
                    except Exception as e:
                        print(f"[WARNING] Could not delete file {filepath}: {e}")
                
            # Mark as read
            service.users().messages().modify(
                userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}
            ).execute()
            print(f"[DONE] Pipeline complete for '{subject}'.\n")
    else:
        print(f"[INFO] No unread emails found for '{subject}'.")

def start_pipeline():
    """The original infinite loop for the GUI's 'Start Pipeline' button."""
    setup_directories()
    service = get_gmail_service()
    print("Listening for unread emails continuously...")
    while True:
        try:
            for rule in EMAIL_RULES:
                process_single_rule(service, rule["subject"], rule["recipients"])
        except Exception as e:
            print(f"[ERROR] Pipeline encountered an issue: {e}")
            service = get_gmail_service()
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    # Setup the argument parser for Task Scheduler
    parser = argparse.ArgumentParser()
    parser.add_argument("--subject", type=str, help="Run a specific rule by subject")
    args = parser.parse_args()

    setup_directories()
    service = get_gmail_service()

    if args.subject:
        # Run in Scheduled Job mode (Process one and exit)
        rule = next((r for r in EMAIL_RULES if r["subject"] == args.subject), None)
        if rule:
            process_single_rule(service, rule["subject"], rule["recipients"])
        else:
            print(f"[ERROR] No rule found in settings for subject: {args.subject}")
    else:
        # Run in Continuous Listener mode
        start_pipeline()