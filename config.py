import os

# --- IMAP SETTINGS (For Reading the Trigger Email) ---
EMAIL_ACCOUNT = "daohog.jason1@gmail.com"        # Replace with your email
APP_PASSWORD = "hzgd qcfy aoqk audi"  # Replace with your App Password
IMAP_SERVER = "imap.gmail.com"
TARGET_SUBJECT = "Fwd: CNFM OPEN FT and WO"

# --- SMTP SETTINGS (For Sending the HTML Table Email) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = EMAIL_ACCOUNT                  # Reusing the same email to send
SENDER_PASSWORD = APP_PASSWORD                # Reusing the same App Password

# --- RECIPIENT SETTINGS ---
# Add the emails of everyone who needs this report. 
# For testing, just put your own email here first!
RECIPIENT_EMAILS = [
    "cacao.thesis08@gmail.com", 

]

# --- DIRECTORY SETTINGS ---
INPUT_DIR = "inputs"

# --- PIPELINE SETTINGS ---
CHECK_INTERVAL = 60  # How often to check the inbox (in seconds)

def setup_directories():
    """Ensures the necessary directories exist before the pipeline runs."""
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)