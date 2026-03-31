import os
import sys
import json

# --- GET TRUE APPLICATION PATH ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- GOOGLE API SETTINGS ---
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")               
SCOPES = ['https://mail.google.com/']   

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

# --- SAFELY LOAD SETTINGS ---
# Check if the file exists. If it doesn't, create a default one to prevent crashes!
if not os.path.exists(SETTINGS_FILE):
    settings = {
        "email_rules": [],
        "sender_email": "",
        "input_dir": os.path.join(BASE_DIR, "inputs"),
        "check_interval": 60
    }
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception:
        pass  # Failsafe just in case folder permissions are locked
else:
    try:
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
    except Exception:
        settings = {}  # Fallback if the JSON file is corrupted or empty

# --- DYNAMIC SETTINGS ---
EMAIL_RULES = settings.get("email_rules", [])
SENDER_EMAIL = settings.get("sender_email", "")
INPUT_DIR = settings.get("input_dir", os.path.join(BASE_DIR, "inputs"))
CHECK_INTERVAL = settings.get("check_interval", 60)

def setup_directories():
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            pass
    return {"sender_email": "", "email_rules": [], "input_dir": os.path.join(BASE_DIR, "inputs")}