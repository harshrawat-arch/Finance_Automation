import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --- CONFIGURATION ---
SENDER_EMAIL = "harsh.rawat@paipai.mobi"
SENDER_PASSWORD = "ryklyfoiqfkqyevv" 

# ✅ Multiple recipients (Fixed)
RECEIVER_EMAIL = ["harsh.rawat@paipai.mobi", "gaurav43.kumar@paipai.mobi"]

CHECKLIST_FILE = r'output_files\checklist.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'

def send_gmail():
    try:
        if not os.path.exists(CHECKLIST_FILE):
            print(f"❌ Error: {CHECKLIST_FILE} not found.")
            return

        # 1. Load and format the date
        date_display = ""
        if os.path.exists(DATA_FILE):
            date_df = pd.read_csv(DATA_FILE, usecols=['settled_at'], encoding='latin1')
            if not date_df.empty:
                raw_date = pd.to_datetime(date_df['settled_at'].iloc[0], errors='coerce')
                if pd.notnull(raw_date):
                    date_display = raw_date.strftime('%d-%B-%Y')
        
        subject_line = f"MKPL REGULAR PAYOUT -Testing - {date_display}"

        # 2. Load checklist data
        df = pd.read_excel(CHECKLIST_FILE)

        # 3. Metric sequence
        metric_sequence = [
            'GMV', 'pg_payable2', 'cod_payable2', 'Seller Payable', 
            'Commission', 'Tax', 'cust_shipping_reversal', 
            'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 
            'additional_delivery_charges2', 'cart_conv_fee2', 
            'mp_shipping', 'shipping_amount2'
        ]

        # 4. Transpose and Enforce Order
        df_transposed = df.set_index('direction_name').T
        df_transposed = df_transposed.apply(pd.to_numeric, errors='coerce').fillna(0)
        
        cols_in_df = df_transposed.columns.tolist()
        ordered_cols = []
        if 'payout' in cols_in_df: ordered_cols.append('payout')
        if 'Return' in cols_in_df: ordered_cols.append('Return')
        df_transposed = df_transposed[ordered_cols]

        # 5. Calculate Total and Reorder Rows
        df_transposed['Total'] = df_transposed.sum(axis=1)
        available_metrics = [m for m in metric_sequence if m in df_transposed.index]
        df_transposed = df_transposed.reindex(available_metrics)

        # 6. Create Styled HTML Table
        html_table = df_transposed.to_html(classes='executive-table', border=0, float_format='{:,.2f}'.format)
        
        body = f"""
        <html>
        <head>
        <style>
            .executive-table {{ 
                border-collapse: collapse; 
                width: auto; 
                min-width: 500px;
                font-family: 'Tahoma', sans-serif; 
                font-size: 13px; 
                border: 5px double #000000 !important; 
                margin: 10px 0;
            }}
            .executive-table th {{ 
                background-color: #92D050 !important; 
                color: #000000 !important; 
                padding: 10px 15px; 
                text-align: center;
                border: 4px double #000000 !important; 
                font-weight: bold !important;
            }}
            .executive-table td {{ 
                padding: 8px 12px; 
                text-align: right; 
                border: 4px double #000000 !important; 
                color: #000000 !important; 
                font-weight: normal !important; 
            }}
            .executive-table tr td:first-child {{ 
                text-align: left; 
                font-weight: bold !important; 
                color: #000000 !important; 
                background-color: #D9D9D9 !important; 
                white-space: nowrap; 
            }}
            .executive-table tr td:last-child {{
                background-color: #F2F2F2 !important;
                font-weight: normal !important;
            }}
            .body-text {{
                font-family: 'Tahoma', sans-serif;
                font-size: 14px;
                color: #000000;
            }}
        </style>
        </head>
        <body>
            <div class="body-text">
                <p>Dear Team,</p>
                <p>Please find the Payout Summary for {date_display}:</p>
            </div>
            {html_table}
            <div class="body-text">
                <p><br>Regards,<br>Finance Automation System</p>
            </div>
        </body>
        </html>
        """

        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = ", ".join(RECEIVER_EMAIL)  # ✅ Display properly in email
        msg['Subject'] = subject_line
        msg.attach(MIMEText(body, 'html'))

        # Attach files (unchanged)
        files_to_attach = [
            r'output_files\checklist.xlsx',
            r'output_files\output.xlsx',
            r'output_files\payment sheet.xlsx',
            r'output_files\payout_summary.xlsx'
        ]

        for file_path in files_to_attach:
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={os.path.basename(file_path)}"
                    )
                    msg.attach(part)
            else:
                print(f"⚠ File not found (skipped): {file_path}")

        # Send Email
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)

        # ✅ Send to multiple recipients (Fixed)
        server.send_message(msg, from_addr=SENDER_EMAIL, to_addrs=RECEIVER_EMAIL)

        server.quit()
    
        print(f"✅ SUCCESS: Email sent to multiple recipients with attachments.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    send_gmail()