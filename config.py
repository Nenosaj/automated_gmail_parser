import os
import json

# --- GOOGLE API SETTINGS ---
CLIENT_SECRET_FILE = "credentials.json" 
TOKEN_FILE = "token.json"               
SCOPES = ['https://mail.google.com/']   

SETTINGS_FILE = "settings.json"

# Load settings from the JSON file
with open(SETTINGS_FILE, 'r') as f:
    settings = json.load(f)

# --- DYNAMIC SETTINGS ---
EMAIL_RULES = settings.get("email_rules", [])
SENDER_EMAIL = settings.get("sender_email", "")
INPUT_DIR = settings.get("input_dir", "inputs")
CHECK_INTERVAL = settings.get("check_interval", 60)

def setup_directories():
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)