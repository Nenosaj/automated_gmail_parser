import os

# --- GOOGLE API SETTINGS ---
CLIENT_SECRET_FILE = "credentials.json" # Downloaded from Google Cloud Console
TOKEN_FILE = "token.json"               # Created automatically after login
SCOPES = ['https://mail.google.com/']   # Full access to Gmail

# --- EMAIL CONTENT SETTINGS ---
TARGET_SUBJECT = "Fwd: CNFM OPEN FT and WO"
SENDER_EMAIL = "daohog.jason1@gmail.com" 

# --- RECIPIENT SETTINGS ---
RECIPIENT_EMAILS = [
    "cacao.thesis08@gmail.com",

]

# --- DIRECTORY & PIPELINE SETTINGS ---
INPUT_DIR = "inputs"
CHECK_INTERVAL = 60 

def setup_directories():
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)