import pandas as pd
import numpy as np
import os

# --- CONFIGURATION (Input/Output Paths) ---
HEADER_FILE = r'input_files\Header.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'
OUTPUT_FILE = r'output_files\output.xlsx'
SUMMARY_FILE = r'output_files\payout_summary.xlsx'

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
        df = pd.read_csv(DATA_FILE)

        # Helper to clean numeric data
        def to_num(cols):
            for col in cols:
                if col not in df.columns: 
                    df[col] = 0.0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

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

        # Commission and Tax logic
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

        # pf, product_gst, TCS, TDS
        to_num(['cust_shipping_reversal', 'partial_shipping_rev'])
        to_num(['pf_tax', 'pf_packing', 'pf_seller_convenience', 'product_igst', 'product_cgst', 'product_sgst',
                'tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 'shipping_tcs_sgst', 'tds_ecom'])
        
        df['pf'] = ((df['pf_tax'] + df['pf_packing'] + df['pf_seller_convenience']) * -1).round(2)
        df['product_gst'] = (df['product_igst'] + df['product_cgst'] + df['product_sgst']).round(2)
        df['TCS'] = df[['tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 'shipping_tcs_sgst']].sum(axis=1).round(2)
        df['TDS'] = df['tds_ecom'].round(2)

        # Merchant Payable & Diff
        df['Seller Payable'] = (df['GMV'] - df['Commission'] - df['Tax'] - df['cust_shipping_reversal'] - 
                                  df['partial_shipping_rev'] - df['pf'] - df['product_gst'] - df['TCS'] - df['TDS']).round(2)
        
        to_num(['merchant_payable'])
        df['Diff'] = (df['merchant_payable'] - df['Seller Payable']).round(2)

        # --- 5. NEW ALIAS COLUMNS (Suffix 2) ---
        # Map original source columns and create duplicates with "2"
        if 'merchant_id_name' not in df.columns: df['merchant_id_name'] = ''
        df['merchant_id_name2'] = df['merchant_id_name'].fillna('').astype(str)

        # Check for variation in additional_delivery column name
        add_delivery_src = 'additional_delivery_charges'
        if add_delivery_src not in df.columns:
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

        # --- 6. CREATE SUMMARY PIVOT TABLE ---
        summary_columns = ['GMV', 'Commission', 'Tax', 'pf', 'product_gst', 'TCS', 'TDS', 'Seller Payable', 'Diff']
        summary_df = df.groupby('MID_WID')[summary_columns].sum().reset_index()
        summary_df.to_excel(SUMMARY_FILE, index=False)

        # --- 7. FINAL EXPORT (Main Data) ---
        # Maintain template columns first
        final_cols = [c for c in template_headers if c in df.columns]
        
        # Force add calculated and aliased columns in order
        force_add_cols = [
            'MID_WID', 'Status', 'GMV', 'Commission', 'Tax', 'cust_shipping_reversal', 
            'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 'Seller Payable', 'Diff',
            'merchant_id_name2', 'pg_payable2', 'cod_payable2', 'shipping_amount2', 
            'additional_delivery_charges2', 'cart_conv_fee2'
        ]
        
        for col in force_add_cols:
            if col not in final_cols:
                final_cols.append(col)

        df[final_cols].to_excel(OUTPUT_FILE, index=False)
        
        print("-" * 40)
        print(f"✅ SUCCESS!")
        print(f"Main Report: {OUTPUT_FILE}")
        print(f"Summary Pivot: {SUMMARY_FILE}")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    process_revenue_report()