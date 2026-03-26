import pandas as pd
import numpy as np
import os

# --- CONFIGURATION (Input/Output Paths) ---
HEADER_FILE = r'input_files\Header.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'

# Mapping Input Files (New)
HOLD_FILE = r'input_files\Hold_list.csv'
MM_FILE = r'input_files\MM.csv'
RECOVERY_FILE = r'input_files\Recovery.csv'

# Output Files
OUTPUT_FILE = r'output_files\output.xlsx'
SUMMARY_FILE = r'output_files\payout_summary.xlsx'
PAYMENT_SHEET = r'output_files\payment_sheet.xlsx'

def process_revenue_report():
    try:
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(OUTPUT_FILE)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        # 1. Load Data
        header_df = pd.read_excel(HEADER_FILE)
        template_headers = header_df.columns.tolist()

        # Load with latin1 to handle special characters and ensure low_memory is False for large datasets
        df = pd.read_csv(DATA_FILE, encoding='latin1', low_memory=False)
        hold_df = pd.read_csv(HOLD_FILE, encoding='latin1', low_memory=False)
        mm_df = pd.read_csv(MM_FILE, encoding='latin1', low_memory=False)
        recovery_df = pd.read_csv(RECOVERY_FILE, encoding='latin1', low_memory=False)

        # Helper to clean numeric data and handle missing columns safely
        def to_num(cols):
            for col in cols:
                if col not in df.columns: 
                    df[col] = 0.0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # --- 2. BASIC LOGIC (MID_WID & Status) ---
        m_id = df['merchant_id'].astype(str)
        w_id = df['warehouse_id'].fillna('').astype(str)
        f_type = df['finance_key_type'].fillna('').astype(str)
        df['MID_WID'] = np.where((f_type == 'default') | (f_type == ''), m_id, m_id + "_" + w_id)

        dir_str = df['direction'].astype(str)
        df['Status'] = np.select([(dir_str == '1'), (dir_str == '2')], ['Payout', 'Returned'], default='Unknown')

        # --- 3. FINANCIAL CALCULATIONS (From process.py) ---
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

        to_num(['merchant_payable', 'pg_payable', 'cod_payable', 'shipping_amount', 'cust_shipping_reversal', 'partial_shipping_rev'])
        
        df['merchant_payable2'] = df['merchant_payable']
        df['pg_payable2'] = df['pg_payable']
        df['cod_payable2'] = df['cod_payable']
        
        if 'merchant_id_name' not in df.columns: df['merchant_id_name'] = ''
        df['merchant_id_name2'] = df['merchant_id_name'].fillna('').astype(str)

        # --- 4. OUTPUT 1: PAYOUT SUMMARY (Remains clean of new file imports) ---
        summary_col_order = [
            'MID_WID', 'merchant_id_name2', 'GMV', 'merchant_payable2', 
            'pg_payable2', 'cod_payable2', 'Commission', 'Tax'
        ]
        summary_df = df.groupby(['MID_WID', 'merchant_id_name2'])[summary_col_order[2:]].sum().reset_index()
        summary_df.to_excel(SUMMARY_FILE, index=False)

        # --- 5. OUTPUT 2: PAYMENT SHEET (Impacted by New Imports) ---
        # A. Start with base aggregate data
        payment_df = df.groupby(['MID_WID', 'merchant_id_name2'])[['GMV', 'merchant_payable2', 'pg_payable2', 'cod_payable2']].sum().reset_index()

        # B. Map SAP from MM.csv (MID_WID to MID_WID)
        mm_map = mm_df[['MID_WID', 'SAP_New_Code']].rename(columns={'SAP_New_Code': 'SAP'})
        mm_map['MID_WID'] = mm_map['MID_WID'].astype(str)
        payment_df = payment_df.merge(mm_map, on='MID_WID', how='left')

        # C. Map Hold from Hold_list.csv (MID_WID to MID_WID)
        hold_map = hold_df[['MID_WID']].copy()
        hold_map['MID_WID'] = hold_map['MID_WID'].astype(str)
        hold_map['Hold'] = 'Hold'
        payment_df = payment_df.merge(hold_map, on='MID_WID', how='left')
        payment_df['Hold'] = payment_df['Hold'].fillna('')

        # D. Map Recovery from Recovery.csv (Merchant_ID to merchant_id)
        mid_to_merchant = df[['MID_WID', 'merchant_id']].astype(str).drop_duplicates('MID_WID')
        rec_map = recovery_df[['Merchant_ID', 'TOTAL']].rename(columns={'TOTAL': 'recovery'})
        rec_map['Merchant_ID'] = rec_map['Merchant_ID'].astype(str)
        
        payment_df = payment_df.merge(mid_to_merchant, on='MID_WID', how='left')
        payment_df = payment_df.merge(rec_map, left_on='merchant_id', right_on='Merchant_ID', how='left')
        payment_df['recovery'] = payment_df['recovery'].fillna(0)

        # Select and re-order headers for Payment Sheet exactly as requested
        final_headers_payment = [
            'MID_WID', 'SAP', 'merchant_id_name2', 'GMV', 'merchant_payable2', 
            'pg_payable2', 'cod_payable2', 'Hold', 'recovery'
        ]
        payment_df[final_headers_payment].to_excel(PAYMENT_SHEET, index=False)

        # --- 6. OUTPUT 3: MAIN DATA EXPORT ---
        final_cols = [c for c in template_headers if c in df.columns]
        extra_cols = ['MID_WID', 'Status', 'GMV', 'Commission', 'Tax', 'merchant_id_name2']
        for col in extra_cols:
            if col not in final_cols:
                final_cols.append(col)
        
        df[final_cols].to_excel(OUTPUT_FILE, index=False)
        
        print("-" * 40)
        print("✅ SUCCESS!")
        print(f"1. Payment Sheet (with mappings): {PAYMENT_SHEET}")
        print(f"2. Payout Summary (clean): {SUMMARY_FILE}")
        print("-" * 40)

    except Exception as e:
        print(f"❌ CRITICAL ERROR: {e}")

if __name__ == "__main__":
    process_revenue_report()