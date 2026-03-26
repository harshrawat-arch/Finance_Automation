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

RECEIVER_EMAIL = ["harsh.rawat@paipai.mobi", "gaurav43.kumar@paipai.mobi"]

CHECKLIST_FILE = r'output_files\checklist.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'
PAYMENT_FILE = r'output_files\payment sheet.xlsx'


# ---------------- EXISTING FUNCTION (UNCHANGED) ---------------- #
def prepare_email(subject_line, date_display, attach_files=True):
    df = pd.read_excel(CHECKLIST_FILE)

    metric_sequence = [
        'GMV', 'pg_payable2', 'cod_payable2', 'Seller Payable', 
        'Commission', 'Tax', 'cust_shipping_reversal', 
        'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 
        'additional_delivery_charges2', 'cart_conv_fee2', 
        'mp_shipping', 'shipping_amount2'
    ]

    df_transposed = df.set_index('direction_name').T
    df_transposed = df_transposed.apply(pd.to_numeric, errors='coerce').fillna(0)
    
    cols_in_df = df_transposed.columns.tolist()
    ordered_cols = []
    if 'payout' in cols_in_df: ordered_cols.append('payout')
    if 'Return' in cols_in_df: ordered_cols.append('Return')
    df_transposed = df_transposed[ordered_cols]

    df_transposed['Total'] = df_transposed.sum(axis=1)
    available_metrics = [m for m in metric_sequence if m in df_transposed.index]
    df_transposed = df_transposed.reindex(available_metrics)

    html_table = df_transposed.to_html(classes='executive-table', border=0, float_format='{:,.2f}'.format)

    body = f"""
    <html><body>
        <p>Dear Team,</p>
        <p>Please find the Payout Summary for {date_display}:</p>
        {html_table}
        <p><br>Regards,<br>Finance Automation System</p>
    </body></html>
    """

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECEIVER_EMAIL)
    msg['Subject'] = subject_line
    msg.attach(MIMEText(body, 'html'))

    if attach_files:
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

    return msg


# ---------------- UPDATED FUNCTION (FORMAT FIX ONLY) ---------------- #
def prepare_filtered_email(subject_line, date_display, mid_wid):
    if not os.path.exists(PAYMENT_FILE):
        print(f"❌ Payment file not found: {PAYMENT_FILE}")
        return None

    df = pd.read_excel(PAYMENT_FILE, sheet_name='Sheet1')

    df.columns = df.columns.str.strip()
    df['MID_WID'] = df['MID_WID'].astype(str).str.strip()

    df_filtered = df[df['MID_WID'] == str(mid_wid)]

    required_cols = ['MID_WID', 'SAP', 'merchant_id_name2', 'GMV']
    available_cols = [col for col in required_cols if col in df_filtered.columns]

    if df_filtered.empty:
        html_table = f"<p style='color:red;'>No data found for MID_WID {mid_wid}</p>"
    else:
        # ✅ FORMAT NUMERIC VALUES (comma + 2 decimal)
        for col in df_filtered.columns:
            if pd.api.types.is_numeric_dtype(df_filtered[col]):
                df_filtered[col] = df_filtered[col].apply(lambda x: "{:,.2f}".format(x))

        html_table = df_filtered[available_cols].to_html(index=False, border=1)

    body = f"""
    <html><body>
        <p>Dear Team,</p>
        <p>Please find the payout details for {date_display}:</p>
        {html_table}
        <p><br>Regards,<br>Finance Automation System</p>
    </body></html>
    """

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECEIVER_EMAIL)
    msg['Subject'] = subject_line
    msg.attach(MIMEText(body, 'html'))

    return msg


# ---------------- MAIN FUNCTION ---------------- #
def send_gmail():
    try:
        if not os.path.exists(CHECKLIST_FILE):
            print(f"❌ Error: {CHECKLIST_FILE} not found.")
            return

        # Get date
        date_display = ""
        if os.path.exists(DATA_FILE):
            date_df = pd.read_csv(DATA_FILE, usecols=['settled_at'], encoding='latin1')
            if not date_df.empty:
                raw_date = pd.to_datetime(date_df['settled_at'].iloc[0], errors='coerce')
                if pd.notnull(raw_date):
                    date_display = raw_date.strftime('%d-%B-%Y')

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)

        # 1️⃣ EXISTING MAIL
        subject_line = f"MKPL REGULAR PAYOUT -Testing - {date_display}"
        msg1 = prepare_email(subject_line, date_display, attach_files=True)
        server.send_message(msg1, from_addr=SENDER_EMAIL, to_addrs=RECEIVER_EMAIL)

        # 2️⃣ GOOGLE PLAY
        subject_line = f"GOOGLE PLAY PAYOUT -Testing - {date_display}"
        msg2 = prepare_filtered_email(subject_line, date_display, 1139089)
        if msg2:
            server.send_message(msg2, from_addr=SENDER_EMAIL, to_addrs=RECEIVER_EMAIL)

        # 3️⃣ APPLE BRAND
        subject_line = f"Apple Brand PAYOUT -Testing - {date_display}"
        msg3 = prepare_filtered_email(subject_line, date_display, 1182161)
        if msg3:
            server.send_message(msg3, from_addr=SENDER_EMAIL, to_addrs=RECEIVER_EMAIL)

        server.quit()

        print("✅ SUCCESS: All emails sent with formatted amounts.")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    send_gmail()