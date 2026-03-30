# ⚙️ Automated Email Parser

A standalone Windows desktop application designed to automatically parse incoming emails, extract attached Excel data, convert it into clean HTML reports, and forward it to designated recipients based on a daily schedule.

![App Icon](app_icon.png)

## 🚀 Features
* **Automated Workflow:** Reads unread emails, downloads Excel attachments, and forwards the data as a clean HTML table.
* **Custom Rules Engine:** Route different Excel reports to different teams based on the incoming email subject line.
* **Windows Task Scheduler Integration:** Set it and forget it. Syncs directly with Windows to run in the background every day, even if you are AFK.
* **Auto-Cleanup:** Deletes local Excel files immediately after processing to save disk space and protect data privacy.

---

## 🛠️ Installation

**Note:** You do not need Python installed to use this application!

1. Go to the [Releases](../../releases) page of this repository.
2. Download the latest `Automated_Email_Parser_Setup_vX.X.exe`.
3. Double-click to install. 
*(**Antivirus Notice:** Because this is an open-source tool without a corporate digital certificate, Windows Defender SmartScreen may flag it. Click **"More Info"** -> **"Run Anyway"** to proceed).*
4. Open the application from your Start Menu or Desktop shortcut.

---

## 🔑 Step 1: Google Cloud API Setup (One-Time Only)

To allow the application to read and send emails on your behalf, you must provide it with Google Cloud API credentials.

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Click the dropdown at the top and create a **New Project** (e.g., "Email Parser App").
3. In the left sidebar, navigate to **APIs & Services > Library**.
4. Search for **Gmail API** and click **Enable**.
5. Go to **APIs & Services > OAuth consent screen**.
    * Choose **External** (or Internal if using a Google Workspace account).
    * Fill in the required app name and developer email.
    * Click **Save and Continue** until you finish the setup.
6. Go to **APIs & Services > Credentials**.
    * Click **Create Credentials > OAuth client ID**.
    * Select **Desktop app** as the application type and name it.
    * Click **Create**.
7. Click the **Download JSON** button next to your new credentials. 
8. Rename this downloaded file to `credentials.json`.

---

## 💻 Step 2: Using the Application

1. **Upload Credentials:** Open the Automated Email Parser app. Under Step 1, click **Upload JSON** and select the `credentials.json` file you just downloaded.
2. **First Login:** The first time the script runs, a browser window will open asking you to log into your Google account and grant the app permission.
3. **Sender Setup:** Enter your email address in the **Step 2: Sender Setup** box.
4. **Create Rules:** * Under Step 3, define the exact **Subject** line of the incoming emails you want to target.
    * Enter the **Recipients** (comma-separated) who should receive the parsed HTML report.
    * Set a **Time** (24-hour format) for when the app should check for this specific report.
    * Click **+ Add New Rule**.
5. **Sync with Windows:** Once your rules are added to the table, click **📅 Sync Tasks to Windows Scheduler**. 

The app will now automatically wake up in the background at your scheduled times, process any unread emails matching your rules, forward the reports, and clean up after itself! 

---

## 👨‍💻 Built By
Created by **Jason T. Daohog**