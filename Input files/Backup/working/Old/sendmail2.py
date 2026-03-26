import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# --- CONFIGURATION ---
SENDER_EMAIL = "harsh.rawat@paipai.mobi"
SENDER_PASSWORD = "ryklyfoiqfkqyevv" 
RECEIVER_EMAIL = "harsh.rawat@paipai.mobi"
CHECKLIST_FILE = r'output_files\checklist.xlsx'

def send_gmail():
    try:
        if not os.path.exists(CHECKLIST_FILE):
            print(f"❌ Error: {CHECKLIST_FILE} not found.")
            return

        # 1. Load data
        df = pd.read_excel(CHECKLIST_FILE)

        # 2. Metric sequence as per your requirement
        metric_sequence = [
            'GMV', 'pg_payable2', 'cod_payable2', 'Seller Payable', 
            'Commission', 'Tax', 'cust_shipping_reversal', 
            'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 
            'additional_delivery_charges2', 'cart_conv_fee2', 
            'mp_shipping', 'shipping_amount2'
        ]

        # 3. Transpose data
        df_transposed = df.set_index('direction_name').T
        df_transposed = df_transposed.apply(pd.to_numeric, errors='coerce').fillna(0)
        
        # 4. Force Column Order: payout first, then Return
        # We check if columns exist to prevent errors
        cols_in_df = df_transposed.columns.tolist()
        ordered_cols = []
        if 'payout' in cols_in_df: ordered_cols.append('payout')
        if 'Return' in cols_in_df: ordered_cols.append('Return')
        
        # Reorder columns based on Payout -> Return
        df_transposed = df_transposed[ordered_cols]

        # 5. Calculate Total
        df_transposed['Total'] = df_transposed.sum(axis=1)

        # 6. Reorder Rows based on metric_sequence
        available_metrics = [m for m in metric_sequence if m in df_transposed.index]
        df_transposed = df_transposed.reindex(available_metrics)

        # 7. Create Styled HTML Table
        html_table = df_transposed.to_html(classes='summary-table', border=1, float_format='{:.2f}'.format)
        
        body = f"""
        <html>
        <head>
        <style>
            .summary-table {{ 
                border-collapse: collapse; 
                width: 100%; 
                font-family: Calibri, sans-serif; 
                font-size: 14px;
                border: 1px solid #000;
            }}
            /* Header Style: Yellow Background */
            .summary-table th {{ 
                background-color: #FFFF00; 
                color: black; 
                padding: 10px; 
                text-align: center;
                border: 1px solid #000;
            }}
            /* Metric Column Style: Light Grey */
            .summary-table td {{ 
                padding: 8px; 
                text-align: right; 
                border: 1px solid #000;
            }}
            .summary-table tr td:first-child {{ 
                text-align: left; 
                font-weight: bold; 
                background-color: #D9D9D9; 
                border: 1px solid #000;
            }}
        </style>
        </head>
        <body>
            <p>Dear Team,</p>
            <p>Please find the <b>Check1: Payout Calculations Summary</b> below:</p>
            {html_table}
            <p><br>Regards,<br><b>Finance Automation System</b></p>
        </body>
        </html>
        """

        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = ":MKPL REGULAR PAYOUT -TEST :-)"
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        print("✅ SUCCESS: Payout now appears before Return. Email sent.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    send_gmail()