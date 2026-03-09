import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT, RECIPIENT_EMAILS

def send_html_email(records, headers):
    """Generates an HTML table and emails it to the recipient list."""
    if not records:
        print("[ERROR] No data provided to send email.")
        return

    print("Generating HTML email body...")
    
    # 1. Build the HTML email with INLINE CSS so it renders correctly in Outlook/Gmail
    html_content = """
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p>Please find the daily broadband activities below:</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 11px;">
            <tr>
                <td colspan="{}" style="background-color: yellow; text-align: center; font-weight: bold; font-size: 16px; color: blue; border: 1px solid black; padding: 8px;">
                    DAYTIME/SICE MOBILE BROADBAND ACTIVITIES
                </td>
            </tr>
            <tr>
    """.format(len(headers))
    
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

    # 2. Setup the Email Message
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENT_EMAILS)
    msg['Subject'] = "Fwd: CNFM OPEN FT and WO"
    
    msg.attach(MIMEText(html_content, 'html'))

    # 3. Connect to SMTP and Send
    try:
        print("Connecting to SMTP server to send emails...")
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls() # Secure the connection
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        
        server.send_message(msg)
        server.quit()
        print(f"[SUCCESS] Formatted email successfully sent to {len(RECIPIENT_EMAILS)} recipient(s)!")
        
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")