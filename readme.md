# Automated Excel Report Forwarder (Gmail API)

This Python automation pipeline monitors a Gmail inbox for specific trigger emails, downloads attached Excel (`.xlsx`) reports, extracts the data, and forwards a formatted HTML summary table to a designated mailing list. 

It is designed for secure, corporate environments by utilizing **OAuth 2.0** and the **Gmail API** rather than outdated App Passwords or standard SMTP/IMAP protocols.

## Features
* **Secure Authentication:** Uses OAuth 2.0 via the Google Cloud Gmail API.
* **Automated Polling:** Continuously checks for unread emails matching a specific subject.
* **Data Extraction:** Utilizes `pandas` to clean and extract tabular data from Excel attachments.
* **HTML Generation:** Automatically builds an inline-styled HTML table for maximum compatibility with corporate email clients (Outlook, Gmail).
* **State Management:** Removes the 'UNREAD' label after processing to prevent duplicate forwards.

## Prerequisites
* Python 3.8 or higher
* A Google Cloud Project with the **Gmail API** enabled.

## Installation

1. **Clone or download the repository** to your local machine.
2. **Install the required Python packages:**
   Open your command prompt or terminal and run:
   ```bash
   pip install pandas openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib