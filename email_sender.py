import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import config 

def send_html_email(service, records, headers, subject, recipients, custom_body=""):
    """Generates an HTML table and emails it via the Gmail API."""
    if not records:
        print("[ERROR] No data provided to send email.")
        return

    print(f"Generating HTML email body for {subject}...")
    import sys
    sys.stdout.flush()
    
    html_content = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p></p>
        <p style="white-space: pre-wrap; margin-bottom: 15px;">{custom_body}</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 11px;">
            <tr>
                <td colspan="{len(headers)}" style="background-color: #559ED6; text-align: center; font-weight: bold; font-size: 16px; color: white; border: 1px solid black; padding: 8px;">
                    {subject}
                </td>
            </tr>
            <tr>
    """
    
    # Add column headers
    for header in headers:
        html_content += f'<th style="background-color: #1A4280; color: white; font-weight: bold; border: 1px solid black; padding: 5px;">{header.upper()}</th>'
    html_content += "</tr>"
    
    # Add data rows
    for row in records:
        html_content += "<tr>"
        for header in headers:
             html_content += f'<td style="border: 1px solid black; padding: 5px; white-space: nowrap;">{row.get(header, "")}</td>'
        html_content += "</tr>"
        
    html_content += """
        </table>
    </body>
    </html>
    """

    # Setup the Email Message
    msg = MIMEMultipart()
    
    # GRAB THE FRESH SENDER EMAIL FROM SETTINGS
    current_settings = config.load_settings()
    msg['From'] = current_settings.get("sender_email", "")
    
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    
    msg.attach(MIMEText(html_content, 'html'))

    # Connect to API and Send
    try:
        print("Connecting to API to send emails...")
        sys.stdout.flush()
        raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        print(f"[SUCCESS] Formatted email successfully sent to {len(recipients)} recipient(s)!")
        sys.stdout.flush()
        
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")
        sys.stdout.flush()