import time
import sys
import os
import argparse

import config # Imports absolute paths safely
from config import INPUT_DIR, CHECK_INTERVAL, setup_directories
from email_handler import get_gmail_service, get_unread_emails_by_subject, download_excel_attachments
from excel_processor import extract_data_from_excel
from email_sender import send_html_email

def process_single_rule(service, subject, recipients, custom_body=""):
    """Processes only one specific rule, then stops and cleans up."""
    print(f"Checking for emails matching: '{subject}'...")
    sys.stdout.flush() 
    
    messages = get_unread_emails_by_subject(service, subject)
    
    if messages:
        print(f"[SUCCESS] Found {len(messages)} new email(s) for '{subject}'! Processing...")
        sys.stdout.flush()
        for msg_info in messages:
            msg_id = msg_info['id']
            saved_files = download_excel_attachments(service, msg_id, INPUT_DIR)
            
            for filepath in saved_files:
                print(f"[DOWNLOADED] Saved attachment to: {filepath}")
                sys.stdout.flush()
                records, headers = extract_data_from_excel(filepath)
                
                if records and headers:
                    print("Initiating email forward sequence...")
                    sys.stdout.flush()
                    send_html_email(service, records, headers, subject, recipients, custom_body)
                    
                    # --- AUTO-CLEANUP STEP ---
                    try:
                        print(f"[CLEANUP] Deleting processed file: {filepath}")
                        os.remove(filepath)
                    except Exception as e:
                        print(f"[WARNING] Could not delete file {filepath}: {e}")
                    sys.stdout.flush()
                
            # Mark as read
            service.users().messages().modify(
                userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}
            ).execute()
            print(f"[DONE] Pipeline complete for '{subject}'.\n")
            sys.stdout.flush()
    else:
        print(f"[INFO] No unread emails found for '{subject}'.")
        sys.stdout.flush()

# --- GUI ADAPTER FUNCTIONS ---
def load_settings():
    return config.load_settings()

def process_rule(rule, sender_email):
    setup_directories()
    service = get_gmail_service()
    process_single_rule(service, rule["subject"], rule["recipients"], rule.get("body", ""))

def continuous_loop():
    """The original infinite loop for the GUI's 'Start Pipeline' button."""
    setup_directories()
    service = get_gmail_service()
    print("[SYSTEM] Listening for unread emails continuously...")
    sys.stdout.flush()
    
    while True:
        try:
            current_settings = config.load_settings()
            rules = current_settings.get("email_rules", [])
            
            for rule in rules:
                process_single_rule(service, rule["subject"], rule["recipients"], rule.get("body", ""))
        except Exception as e:
            print(f"[ERROR] Pipeline encountered an issue: {e}")
            sys.stdout.flush()
            service = get_gmail_service()
            
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--subject", type=str, help="Run a specific rule by subject")
    args, unknown = parser.parse_known_args()

    setup_directories()
    service = get_gmail_service()

    if args.subject:
        current_settings = config.load_settings()
        rules = current_settings.get("email_rules", [])
        rule = next((r for r in rules if r["subject"] == args.subject), None)
        if rule:
            process_single_rule(service, rule["subject"], rule["recipients"], rule.get("body", ""))
        else:
            print(f"[ERROR] No rule found in settings for subject: {args.subject}")
    else:
        continuous_loop()