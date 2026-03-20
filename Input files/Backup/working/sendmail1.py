import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# --- CONFIGURATION ---
SENDER_EMAIL = "harsh.rawat@paipai.mobi"
SENDER_PASSWORD = "ryklyfoiqfkqyevv"  # Your 16-digit App Password
RECEIVER_EMAIL = "harsh.rawat@paipai.mobi"
SUBJECT = "Update: Automated Payout Report"

def send_gmail():
    try:
        # 1. Create the email container
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = SUBJECT

        # 2. Create the Sample Mail Body
        body = """
        <html>
        <body>
            <h3>Daily Payout Summary</h3>
            <p>Dear Team,</p>
            <p>The processing for today's revenue data has been completed successfully. 
            Attached is the automated checklist for your review.</p>
            
            <table border="1" style="border-collapse: collapse; width: 50%;">
                <tr style="background-color: #f2f2f2;">
                    <th>Metric</th>
                    <th>Status</th>
                </tr>
                <tr>
                    <td>Data Validation</td>
                    <td style="color: green;">Passed</td>
                </tr>
                <tr>
                    <td>Nodal Status Check</td>
                    <td style="color: green;">Completed</td>
                </tr>
            </table>
            
            <p>Please let us know if any further action is required.</p>
            <p>Best Regards,<br><b>Automated System</b></p>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))

        # 3. Connect to Gmail's SMTP Server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Secure the connection
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        
        # 4. Send the email
        server.send_message(msg)
        server.quit()
        
        print("✅ Email sent successfully!")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    send_gmail()