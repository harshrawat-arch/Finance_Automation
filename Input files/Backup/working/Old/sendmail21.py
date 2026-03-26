import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import warnings

# Suppress fragmentation warnings
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

# --- CONFIGURATION ---
SENDER_EMAIL = "harsh.rawat@paipai.mobi"
SENDER_PASSWORD = "ryklyfoiqfkqyevv"
RECEIVER_EMAIL = ["harsh.rawat@paipai.mobi"]

CHECKLIST_FILE = r'output_files\checklist.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'
PAYMENT_FILE = r'output_files\payment sheet.xlsx'
HOLD_EXCEL_FILE = r'output_files\HOLD MID_WID.xlsx'
OPEN_ORDERS_FILE = r'output_files\Payout _Open orders Direction 1 & 2.xlsx'

# =========== PAYMENT SUMMARY =========== #
def generate_payment_summary(is_main=False):
    if not os.path.exists(PAYMENT_FILE):
        return "<p style='color:red;'>Payment Sheet Missing</p>"

    df = pd.read_excel(PAYMENT_FILE, sheet_name="Sheet1")
    df.columns = df.columns.str.strip()
    df["MID_WID"] = df["MID_WID"].astype(str).str.strip()

    # Convert numeric columns safely
    df["pg_payable2"] = pd.to_numeric(df["pg_payable2"], errors="coerce").fillna(0)
    df["cod_payable2"] = pd.to_numeric(df["cod_payable2"], errors="coerce").fillna(0)

    # Handle Hold/Nodal status
    df["Hold"] = df.get("Hold", df.get("HOLD", "")).astype(str).str.strip().str.lower()
    df["Nodal_is_numeric"] = pd.to_numeric(df.get("Nodal_Status", ""), errors="coerce").notna()

    # Core Calculations
    as_escrow = df["pg_payable2"].sum()
    as_cod = df["cod_payable2"].sum()

    euronet_mids = ["1139089", "1182161"]
    df_euronet = df[df["MID_WID"].isin(euronet_mids)]
    he_escrow, he_cod = df_euronet["pg_payable2"].sum(), df_euronet["cod_payable2"].sum()

    df_other = df[(df["Hold"] == "hold") | ((~df["Nodal_is_numeric"]) & (df["pg_payable2"] >= 0) & (df["cod_payable2"] >= 0))]
    df_other = df_other[~df_other["MID_WID"].isin(euronet_mids)]
    ho_escrow, ho_cod = df_other["pg_payable2"].sum(), df_other["cod_payable2"].sum()

    ad_escrow = df[df["pg_payable2"] < 0]["pg_payable2"].sum() * -1
    ad_cod = df[df["cod_payable2"] < 0]["cod_payable2"].sum() * -1

    # --- FIX: Mapping Adjust_Recovery to the correct columns from the Payment Sheet ---
    def get_sum(col_name):
        if col_name in df.columns:
            return pd.to_numeric(df[col_name], errors="coerce").fillna(0).sum()
        return 0.0

    # These column names now match the ones generated in your main processing script
    adj_recov_escrow = get_sum("Recovered_pg") 
    adj_recov_cod = get_sum("Recovered_cod")

    # Final Totals
    # Note: Recovery is already negative in the sheet, so we ADD it here to reduce the payable
    payable_escrow = as_escrow - he_escrow - ho_escrow + ad_escrow + adj_recov_escrow
    payable_cod = as_cod - he_cod - ho_cod + ad_cod + adj_recov_cod
    total_val = payable_escrow + payable_cod

    # Table sequence exactly as per image
    summary_df = pd.DataFrame({
        "Mode": ["As Per Panel", "Hold_Euronet", "Hold_Other", "Added-Recovery", "Adjust_Recovery", "Payable amount"],
        "ESCROW": [as_escrow, he_escrow, ho_escrow, ad_escrow, adj_recov_escrow, payable_escrow],
        "COD": [as_cod, he_cod, ho_cod, ad_cod, adj_recov_cod, payable_cod]
    })
    summary_df["Total"] = summary_df["ESCROW"] + summary_df["COD"]

    for col in ["ESCROW", "COD", "Total"]:
        summary_df[col] = summary_df[col].apply(lambda x: f"{x:,.2f}")

    html = summary_df.to_html(index=False, border=1, classes='styled-table')

    if is_main:
        html = html.replace("<td>Payable amount</td>", '<td style="background:#ff6600;color:white;"><b>Payable amount</b></td>')
        kpi_html = f"""
        <table cellpadding="6" style="margin-bottom:10px; font-family:Arial;">
        <tr>
            <td style="border:2px solid #ff6600;border-radius:8px;text-align:center;width:120px;"><b>ESCROW</b><br>{payable_escrow:,.2f}</td>
            <td style="border:2px solid #ff6600;border-radius:8px;text-align:center;width:120px;"><b>COD</b><br>{payable_cod:,.2f}</td>
            <td style="background:#ff6600;color:white;border-radius:8px;text-align:center;width:140px;"><b>TOTAL</b><br>{total_val:,.2f}</td>
        </tr>
        </table>
        """
        return kpi_html + html
    return html

# ================= EMAIL UTILS ================= #
def prepare_email(subject_line, date_display, is_main=False):
    df = pd.read_excel(CHECKLIST_FILE)
    metric_sequence = ['GMV', 'pg_payable2', 'cod_payable2', 'Seller Payable', 'Commission', 'Tax', 'pf', 'TCS', 'TDS', 'mp_shipping', 'shipping_amount2']
    df_transposed = df.set_index('direction_name').T.apply(pd.to_numeric, errors='coerce').fillna(0)
    
    cols_in_df = df_transposed.columns.tolist()
    ordered_cols = [c for c in ['payout', 'Return'] if c in cols_in_df]
    df_transposed = df_transposed[ordered_cols]
    df_transposed['Total'] = df_transposed.sum(axis=1)
    
    available_metrics = [m for m in metric_sequence if m in df_transposed.index]
    df_transposed = df_transposed.reindex(available_metrics)
    
    html_table = df_transposed.to_html(classes='styled-table', border=1, float_format='{:,.2f}'.format)
    payment_summary_html = generate_payment_summary(is_main)

    css_style = "<style>.styled-table { border-collapse: collapse; font-size: 13px; width: 100%; font-family: Arial; } .styled-table th { background-color: #1f4e79; color: white; padding: 6px; border: 1px solid black; } .styled-table td { padding: 6px; border: 1px solid black; text-align: right; }</style>"
    body = f"<html>{css_style}<body><p>Dear Team,</p><h3>Payment Summary dated ({date_display})</h3>{payment_summary_html}<br><h3>Payout Summary</h3>{html_table}<p><br>Regards,<br>Finance Automation System</p></body></html>"
    msg = MIMEMultipart(); msg['From'] = SENDER_EMAIL; msg['To'] = ", ".join(RECEIVER_EMAIL); msg['Subject'] = subject_line; msg.attach(MIMEText(body, 'html'))
    return msg

def prepare_hold_email(date_display):
    if not os.path.exists(PAYMENT_FILE): return None
    df = pd.read_excel(PAYMENT_FILE, sheet_name='Sheet1')
    df.columns = df.columns.str.strip()
    df['Hold'] = df.get('Hold', df.get('HOLD', "")).astype(str).str.lower()
    df["Nodal_is_numeric"] = pd.to_numeric(df.get("Nodal_Status", ""), errors="coerce").notna()
    
    hold_df = df[(df["Hold"] == "hold") | (~df["Nodal_is_numeric"])].copy()
    if hold_df.empty: return None

    cols = ['MID_WID', 'SAP', 'merchant_id_name2', 'pg_payable2', 'cod_payable2', 'Hold']
    hold_df[cols].to_excel(HOLD_EXCEL_FILE, index=False)
    html_table = hold_df[cols].to_html(index=False, border=1, classes='styled-table')
    
    body = f"<html><body><p>Dear Team,</p><p>Please find the <b>HOLD Details</b> attached for {date_display}.</p>{html_table}<p>Regards,<br>Finance Automation System</p></body></html>"
    msg = MIMEMultipart(); msg['From'] = SENDER_EMAIL; msg['To'] = ", ".join(RECEIVER_EMAIL); msg['Subject'] = f"HOLD Details for the payout dated {date_display}"
    msg.attach(MIMEText(body, 'html'))
    with open(HOLD_EXCEL_FILE, "rb") as f:
        part = MIMEBase("application", "octet-stream"); part.set_payload(f.read()); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(HOLD_EXCEL_FILE)}"); msg.attach(part)
    return msg

def prepare_filtered_email(subject_line, date_display, mid_wid):
    if not os.path.exists(PAYMENT_FILE): return None
    df = pd.read_excel(PAYMENT_FILE, sheet_name='Sheet1')
    df.columns = df.columns.str.strip()
    df_filtered = df[df['MID_WID'].astype(str) == str(mid_wid)]
    html_table = df_filtered[['MID_WID', 'SAP', 'merchant_id_name2', 'GMV']].to_html(index=False, border=1, classes='styled-table')
    body = f"<html><body><p>Dear Team,</p><p>Payout details for {mid_wid} ({date_display}):</p>{html_table}</body></html>"
    msg = MIMEMultipart(); msg['From'] = SENDER_EMAIL; msg['To'] = ", ".join(RECEIVER_EMAIL); msg['Subject'] = subject_line; msg.attach(MIMEText(body, 'html'))
    return msg

def prepare_attachment_email(date_display):
    files = [CHECKLIST_FILE, PAYMENT_FILE, r"output_files\output.xlsx"]
    msg = MIMEMultipart(); msg['From'] = SENDER_EMAIL; msg['To'] = ", ".join(RECEIVER_EMAIL); msg['Subject'] = f"payout file for checking dated {date_display}"
    msg.attach(MIMEText("<p>Please find the final checking files attached.</p>", 'html'))
    for f_path in files:
        if os.path.exists(f_path):
            with open(f_path, "rb") as f:
                part = MIMEBase("application", "octet-stream"); part.set_payload(f.read()); encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(f_path)}"); msg.attach(part)
    return msg

# ================= SEQUENTIAL SEND ================= #
def send_gmail():
    try:
        date_display = "TBD"
        if os.path.exists(DATA_FILE):
            date_df = pd.read_csv(DATA_FILE, usecols=['settled_at'], encoding='latin1', nrows=1)
            if not date_df.empty:
                date_display = pd.to_datetime(date_df['settled_at'].iloc[0]).strftime('%d-%B-%Y')
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)

        # 1) MKPL REGULAR PAYOUT
        print("Sending 1/5: MKPL REGULAR PAYOUT...")
        server.send_message(prepare_email(f"MKPL REGULAR PAYOUT -Testing - {date_display}", date_display, is_main=True))

        # 2) GOOGLE PLAY PAYOUT
        print("Sending 2/5: GOOGLE PLAY PAYOUT...")
        server.send_message(prepare_filtered_email(f"GOOGLE PLAY PAYOUT -Testing - {date_display}", date_display, 1139089))

        # 3) Apple BRAND PAYOUT
        print("Sending 3/5: Apple BRAND PAYOUT...")
        server.send_message(prepare_filtered_email(f"Apple BRAND PAYOUT -Testing - {date_display}", date_display, 1182161))

        # 4) HOLD Details
        print("Sending 4/5: HOLD Details...")
        hold_msg = prepare_hold_email(date_display)
        if hold_msg: server.send_message(hold_msg)

        # 5) Payout file for checking
        print("Sending 5/5: Payout checking files...")
        server.send_message(prepare_attachment_email(date_display))

        server.quit()
        print(f"✅ SUCCESS: All 5 emails sent in sequence for {date_display}.")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    send_gmail()