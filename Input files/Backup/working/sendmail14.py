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


# =========== PAYMENT SUMMARY =========== #
def generate_payment_summary(is_main=False):

    if not os.path.exists(PAYMENT_FILE):
        return "<p style='color:red;'>Payment Sheet Missing</p>"

    df = pd.read_excel(PAYMENT_FILE, sheet_name="Sheet1")

    df.columns = df.columns.str.strip()
    df["MID_WID"] = df["MID_WID"].astype(str).str.strip()

    df["pg_payable2"] = pd.to_numeric(df["pg_payable2"], errors="coerce").fillna(0)
    df["cod_payable2"] = pd.to_numeric(df["cod_payable2"], errors="coerce").fillna(0)

    # Normalize HOLD column
    if "Hold" in df.columns:
        df["Hold"] = df["Hold"].astype(str).str.strip().str.lower()
    elif "HOLD" in df.columns:
        df.rename(columns={"HOLD": "Hold"}, inplace=True)
        df["Hold"] = df["Hold"].astype(str).str.strip().str.lower()
    else:
        df["Hold"] = ""

    # Numeric check for Nodal_Status
    if "Nodal_Status" in df.columns:
        df["Nodal_is_numeric"] = pd.to_numeric(df["Nodal_Status"], errors="coerce").notna()
    else:
        df["Nodal_is_numeric"] = True

    # ===== CALCULATION (UNCHANGED) =====
    as_escrow = df["pg_payable2"].sum()
    as_cod = df["cod_payable2"].sum()

    euronet_mids = ["1139089", "1182161"]
    df_euronet = df[df["MID_WID"].isin(euronet_mids)]

    he_escrow = df_euronet["pg_payable2"].sum()
    he_cod = df_euronet["cod_payable2"].sum()

    df_other = df[
        (df["Hold"] == "hold") |
        ((~df["Nodal_is_numeric"]) &
         (df["pg_payable2"] >= 0) &
         (df["cod_payable2"] >= 0))
    ]

    df_other = df_other[~df_other["MID_WID"].isin(euronet_mids)]

    ho_escrow = df_other["pg_payable2"].sum()
    ho_cod = df_other["cod_payable2"].sum()

    ad_escrow = df[df["pg_payable2"] < 0]["pg_payable2"].sum() * -1
    ad_cod = df[df["cod_payable2"] < 0]["cod_payable2"].sum() * -1

    payable_escrow = as_escrow - he_escrow - ho_escrow + ad_escrow
    payable_cod = as_cod - he_cod - ho_cod + ad_cod
    total_val = payable_escrow + payable_cod

    summary_df = pd.DataFrame({
        "Mode": [
            "As Per Panel",
            "Hold_Euronet",
            "Hold_Other",
            "Added-Recovery",
            "Payable amount"
        ],
        "ESCROW": [as_escrow, he_escrow, ho_escrow, ad_escrow, payable_escrow],
        "COD": [as_cod, he_cod, ho_cod, ad_cod, payable_cod]
    })

    summary_df["Total"] = summary_df["ESCROW"] + summary_df["COD"]

    for col in ["ESCROW", "COD", "Total"]:
        summary_df[col] = summary_df[col].apply(lambda x: f"{x:,.2f}")

    html = summary_df.to_html(index=False, border=1, classes='styled-table')

    # ===== UI ONLY FOR MAIN MAIL =====
    if is_main:

        html = html.replace(
            "<td>Payable amount</td>",
            '<td style="background:#ff6600;color:white;"><b>Payable amount</b></td>'
        )

        html = html.replace(
            f"<td>{payable_escrow:,.2f}</td>",
            f'<td style="background:#ff6600;color:white;"><b>{payable_escrow:,.2f}</b></td>'
        )

        html = html.replace(
            f"<td>{payable_cod:,.2f}</td>",
            f'<td style="background:#ff6600;color:white;"><b>{payable_cod:,.2f}</b></td>'
        )

        html = html.replace(
            f"<td>{total_val:,.2f}</td>",
            f'<td style="background:#ff6600;color:white;"><b>{total_val:,.2f}</b></td>'
        )

        # KPI CARDS (COMPACT)
        kpi_html = f"""
        <table cellpadding="6" style="margin-bottom:10px;">
        <tr>
            <td style="border:2px solid #ff6600;border-radius:8px;text-align:center;width:120px;">
                <b>ESCROW</b><br>{payable_escrow:,.2f}
            </td>
            <td style="border:2px solid #ff6600;border-radius:8px;text-align:center;width:120px;">
                <b>COD</b><br>{payable_cod:,.2f}
            </td>
            <td style="background:#ff6600;color:white;border-radius:8px;text-align:center;width:140px;">
                <b>TOTAL</b><br>{total_val:,.2f}
            </td>
        </tr>
        </table>
        """

        return kpi_html + html

    return html


# ================= EMAIL BODY ================= #
def prepare_email(subject_line, date_display, is_main=False):

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

    html_table = df_transposed.to_html(
        classes='styled-table',
        border=1,
        float_format='{:,.2f}'.format
    )

    payment_summary_html = generate_payment_summary(is_main)

    css_style = """
    <style>
        .styled-table {
            border-collapse: collapse;
            font-size: 13px;
            width: 100%;
        }
        .styled-table th {
            background-color: #1f4e79;
            color: white;
            text-align: center;
            padding: 6px;
            border: 1px solid black;
        }
        .styled-table td {
            padding: 6px;
            border: 1px solid black;
            text-align: right;
        }
    </style>
    """

    body = f"""
    <html>
    {css_style}
    <body>
        <p>Dear Team,</p>

        <h3>Please find the Payment Summary dated ({date_display})</h3>
        {payment_summary_html}
        <br>

        <h3>Payout Summary</h3>
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


# ================= FILTERED EMAIL ================= #
def prepare_filtered_email(subject_line, date_display, mid_wid):

    if not os.path.exists(PAYMENT_FILE):
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
        for col in df_filtered.columns:
            if pd.api.types.is_numeric_dtype(df_filtered[col]):
                df_filtered[col] = df_filtered[col].apply(lambda x: "{:,.2f}".format(x))

        html_table = df_filtered[available_cols].to_html(
            index=False,
            border=1,
            classes='styled-table'
        )

    body = f"""
    <html>
    <body>
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


# ================= ATTACHMENT MAIL ================= #
def prepare_attachment_email(date_display):

    files_to_attach = [
        r"output_files\checklist.xlsx",
        r"output_files\output.xlsx",
        r"output_files\payment sheet.xlsx",
        r"output_files\payout_summary.xlsx"
    ]

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECEIVER_EMAIL)
    msg['Subject'] = f"payout file for checking dated {date_display}"

    msg.attach(MIMEText("<p>Please find attached files</p>", 'html'))

    for filepath in files_to_attach:
        if os.path.exists(filepath):
            with open(filepath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filepath)}")
                msg.attach(part)

    return msg


# ================= SEND ================= #
def send_gmail():
    try:
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

        # MAIN MAIL (ONLY HERE UI)
        msg1 = prepare_email(
            f"MKPL REGULAR PAYOUT -Testing_Final11 - {date_display}",
            date_display,
            is_main=True
        )
        server.send_message(msg1)

        # OTHER MAILS (NO CHANGE)
        server.send_message(prepare_filtered_email(
            f"GOOGLE PLAY PAYOUT -Testing_Final11 - {date_display}",
            date_display, 1139089
        ))

        server.send_message(prepare_filtered_email(
            f"Apple BRAND PAYOUT -Testing_Final11 - {date_display}",
            date_display, 1182161
        ))

        server.send_message(prepare_attachment_email(date_display))

        server.quit()
        print("✅ SUCCESS")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    send_gmail()