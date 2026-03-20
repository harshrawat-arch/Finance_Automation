import pandas as pd
import numpy as np
import os
from checkList import create_checklist_pivot

# --- CONFIGURATION (Input/Output Paths) ---
HEADER_FILE = r'input_files\Header.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'
HOLD_FILE = r'input_files\Hold_list.csv'
MM_FILE = r'input_files\MM.csv'
RECOVERY_FILE = r'input_files\Recovery.csv'

OUTPUT_FILE = r'output_files\output.xlsx'
SUMMARY_FILE = r'output_files\payout_summary.xlsx'
PAYMENT_SHEET_FILE = r'output_files\payment sheet.xlsx'
CHECKLIST_FILE = r'output_files\checklist.xlsx'

def process_revenue_report():
    try:
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(OUTPUT_FILE)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        # 1. Load Header Template
        header_df = pd.read_excel(HEADER_FILE)
        template_headers = header_df.columns.tolist()

        # 2. Load CSV Data
        df = pd.read_csv(DATA_FILE, encoding='latin1')

        # Helper to clean numeric data
        def to_num(cols, target_df=df):
            for col in cols:
                if col not in target_df.columns: 
                    target_df[col] = 0.0
                target_df[col] = pd.to_numeric(target_df[col], errors='coerce').fillna(0)

        # --- 3. BASIC LOGIC (MID_WID & Status) ---
        m_id = df['merchant_id'].astype(str)
        w_id = df['warehouse_id'].fillna('').astype(str)
        f_type = df['finance_key_type'].fillna('').astype(str)
        df['MID_WID'] = np.where((f_type == 'default') | (f_type == ''), m_id, m_id + "_" + w_id)

        dir_str = df['direction'].astype(str)
        df['Status'] = np.select([(dir_str == '1'), (dir_str == '2')], ['Payout', 'Returned'], default='Unknown')

        # --- 4. FINANCIAL CALCULATIONS ---
        to_num(['price', 'qty_ordered'])
        base_gmv = (df['price'] * df['qty_ordered']).round(2)
        df['GMV'] = np.where(dir_str == '1', base_gmv, np.where(dir_str == '2', -base_gmv, 0))

        comm_cols = [
            'pg_commission', 'mp_commission', 'logistics_penalty', 'reverse_logistics_penalty',
            'merchant_cancel_mp_penalty', 'merchant_cancel_pg_penalty', 'logistics',
            'shipping_recovery_mp_fee', 'shipping_recovery_pg_fee', 'mp_commission_reversal',
            'sla_breach_mp_penalty', 'marketing_fee', 'closing_fee', 'deal_setup_fees',
            'partial_shipping_rev_pg_fee', 'partial_shipping_rev_mp_fee', 'pf_taxes_comm',
            'pf_pac_comm', 'pf_scf_comm', 'pf_rev_tax_comm', 'pf_rev_pac_comm', 'pf_rev_scf_comm'
        ]
        to_num(comm_cols)
        total_comm = df[comm_cols].sum(axis=1)
        
        df['Commission'] = np.where(dir_str == '1', total_comm, 
                                   pd.to_numeric(df.get('commission', total_comm), errors='coerce').fillna(0)).round(2)
        df['Tax'] = (total_comm * 0.18).round(2)

        to_num(['cust_shipping_reversal', 'partial_shipping_rev', 'mp_shipping', 'pf_tax', 'pf_packing', 
                'pf_seller_convenience', 'product_igst', 'product_cgst', 'product_sgst',
                'tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 
                'shipping_tcs_sgst', 'tds_ecom', 'merchant_payable'])
        
        df['pf'] = ((df['pf_tax'] + df['pf_packing'] + df['pf_seller_convenience']) * -1).round(2)
        df['product_gst'] = (df['product_igst'] + df['product_cgst'] + df['product_sgst']).round(2)
        df['TCS'] = df[['tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 'shipping_tcs_sgst']].sum(axis=1).round(2)
        df['TDS'] = df['tds_ecom'].round(2)

        df['Seller Payable'] = (df['GMV'] - df['Commission'] - df['Tax'] - df['cust_shipping_reversal'] - 
                                  df['partial_shipping_rev'] - df['pf'] - df['product_gst'] - df['TCS'] - df['TDS']).round(2)
        df['Diff'] = (df['merchant_payable'] - df['Seller Payable']).round(2)

        # --- 5. ALIAS COLUMNS (Suffix 2) ---
        if 'merchant_id_name' not in df.columns: df['merchant_id_name'] = ''
        df['merchant_id_name2'] = df['merchant_id_name'].fillna('').astype(str)

        add_delivery_src = 'additional_delivery_charges'
        for col in df.columns:
            if col.startswith('additional_delivery'):
                add_delivery_src = col
                break

        to_num(['pg_payable', 'cod_payable', 'shipping_amount', add_delivery_src, 'cart_conv_fee'])
        
        df['pg_payable2'] = df['pg_payable']
        df['cod_payable2'] = df['cod_payable']
        df['shipping_amount2'] = df['shipping_amount']
        df['additional_delivery_charges2'] = df[add_delivery_src]
        df['cart_conv_fee2'] = df['cart_conv_fee']
        df['merchant_payable2'] = df['merchant_payable']

        # --- 6. CREATE SUMMARY PIVOT TABLE (payout_summary.xlsx) ---
        summary_sum_cols = [
            'GMV', 'Commission', 'Tax', 'pf', 'product_gst', 'TCS', 'TDS', 
            'Seller Payable', 'Diff', 'pg_payable2', 'cod_payable2', 
            'shipping_amount2', 'additional_delivery_charges2', 'cart_conv_fee2',
            'cust_shipping_reversal', 'partial_shipping_rev', 'mp_shipping', 'merchant_payable2'
        ]
        
        summary_df = df.groupby('MID_WID').agg({
            'merchant_id_name2': 'first',
            'merchant_id': 'first',
            **{col: 'sum' for col in summary_sum_cols}
        }).reset_index()
        summary_df.to_excel(SUMMARY_FILE, index=False)

        # --- 7. CREATE PAYMENT SHEET (Based on payout_summary data) ---
        payment_df = summary_df.copy()

        # Mapping SAP and Nodal Status from MM.csv
        if os.path.exists(MM_FILE):
            mm_df = pd.read_csv(MM_FILE, encoding='latin1')
            # Clean headers to handle potential trailing spaces
            mm_df.columns = mm_df.columns.str.strip()
            mm_df['MID_WID'] = mm_df['MID_WID'].astype(str)
            
            payment_df = pd.merge(payment_df, mm_df[['MID_WID', 'SAP_New_Code', 'Nodal']], on='MID_WID', how='left')
            payment_df.rename(columns={'SAP_New_Code': 'SAP', 'Nodal': 'Nodal_Status'}, inplace=True)
        else:
            payment_df['SAP'] = ""
            payment_df['Nodal_Status'] = ""

        # Mapping Hold
        if os.path.exists(HOLD_FILE):
            hold_df = pd.read_csv(HOLD_FILE, encoding='latin1')
            hold_df['MID_WID'] = hold_df['MID_WID'].astype(str)
            hold_df['Hold_Flag'] = "Hold"
            payment_df = pd.merge(payment_df, hold_df[['MID_WID', 'Hold_Flag']], on='MID_WID', how='left')
            payment_df['Hold'] = payment_df['Hold_Flag'].fillna("")
        else:
            payment_df['Hold'] = ""

        # Mapping Recovery
        if os.path.exists(RECOVERY_FILE):
            rec_df = pd.read_csv(RECOVERY_FILE, encoding='latin1')
            rec_df['Merchant_ID'] = rec_df['Merchant_ID'].astype(str)
            to_num(['TOTAL'], target_df=rec_df)
            rec_df['TOTAL'] = rec_df['TOTAL'].abs()
            rec_grouped = rec_df.groupby('Merchant_ID')['TOTAL'].sum().reset_index()
            payment_df['m_id_str'] = payment_df['merchant_id'].astype(str)
            payment_df = pd.merge(payment_df, rec_grouped, left_on='m_id_str', right_on='Merchant_ID', how='left')
            payment_df.rename(columns={'TOTAL': 'recovery'}, inplace=True)
        else:
            payment_df['recovery'] = 0

        # --- CALCULATE NEW COLUMNS FOR PAYMENT SHEET ---
        payment_df['recovery'] = pd.to_numeric(payment_df['recovery'], errors='coerce').fillna(0)
        
        # 1. net pg_payable calculation
        payment_df['net pg_payable'] = np.where(
            (payment_df['Hold'] == "Hold" ) | (payment_df['Nodal_Status'].astype(str).str.isalpha()), 
            0, 
            (payment_df['pg_payable2'] - payment_df['recovery']).clip(lower=0)
        ).round(2)
        
        # Calculate pending recovery after adjusting from pg_payable2
        payment_df['pending_recovery'] = (payment_df['recovery'] - payment_df['pg_payable2']).clip(lower=0)

        # 2. net cod_payable calculation
        payment_df['net_cod_payable'] = np.where(
            (payment_df['Hold'] == "Hold") | (payment_df['Nodal_Status'].astype(str).str.isalpha()), 
            0, 
            (payment_df['cod_payable2'] - payment_df['pending_recovery']).clip(lower=0)
        ).round(2)

        # Select exact headers for Payment Sheet in required order (Nodal_Status after net_cod_payable)
        pay_headers = ['MID_WID', 'SAP', 'merchant_id_name2', 'GMV', 'merchant_payable2', 
                       'pg_payable2', 'cod_payable2', 'Hold', 'recovery', 
                       'net pg_payable', 'net_cod_payable', 'Nodal_Status']
        
        for h in pay_headers:
            if h not in payment_df.columns: payment_df[h] = ""
            
        payment_df[pay_headers].fillna("").to_excel(PAYMENT_SHEET_FILE, index=False)

        # --- 8. FINAL EXPORT (Main Data - logic unchanged) ---
        final_cols = [c for c in template_headers if c in df.columns]
        force_add_cols = [
            'MID_WID', 'Status', 'GMV', 'Commission', 'Tax', 'cust_shipping_reversal', 
            'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 'Seller Payable', 'Diff',
            'merchant_id_name2', 'pg_payable2', 'cod_payable2', 'shipping_amount2', 
            'additional_delivery_charges2', 'cart_conv_fee2', 'mp_shipping'
        ]
        for col in force_add_cols:
            if col not in final_cols: final_cols.append(col)

        df[final_cols].to_excel(OUTPUT_FILE, index=False)
        
        print("-" * 40)
        print(f"✅ SUCCESS! Nodal_Status added to Payment Sheet.")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error: {e}")




if __name__ == "__main__":
    process_revenue_report()

    create_checklist_pivot(OUTPUT_FILE, CHECKLIST_FILE)