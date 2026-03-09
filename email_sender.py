import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import SENDER_EMAIL, RECIPIENT_EMAILS

def send_html_email(service, records, headers):
    """Generates an HTML table and emails it via the Gmail API."""
    if not records:
        print("[ERROR] No data provided to send email.")
        return

    print("Generating HTML email body...")
    
    html_content = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p>Please find the daily broadband activities below:</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 11px;">
            <tr>
                <td colspan="{len(headers)}" style="background-color: yellow; text-align: center; font-weight: bold; font-size: 16px; color: blue; border: 1px solid black; padding: 8px;">
                    DAYTIME/SICE MOBILE BROADBAND ACTIVITIES
                </td>
            </tr>
            <tr>
    """
    
    # Add column headers
    for header in headers:
        html_content += f'<th style="background-color: blue; color: yellow; font-weight: bold; border: 1px solid black; padding: 5px;">{header.upper()}</th>'
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
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENT_EMAILS)
    msg['Subject'] = "Fwd: CNFM OPEN FT and WO"
    
    msg.attach(MIMEText(html_content, 'html'))

    # Connect to API and Send
    try:
        print("Connecting to API to send emails...")
        raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        print(f"[SUCCESS] Formatted email successfully sent to {len(RECIPIENT_EMAILS)} recipient(s)!")
        
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")