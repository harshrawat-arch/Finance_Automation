import pandas as pd
import numpy as np
import os

# --- CONFIGURATION ---
HEADER_FILE = r'input_files\Header.xlsx'
DATA_FILE = r'input_files\compact_revenue_report.csv'
OUTPUT_FILE = r'output_files\output.xlsx'
SUMMARY_FILE = r'output_files\payout_summary.xlsx' # New output file

def process_revenue_report():
    try:
        if not os.path.exists(os.path.dirname(OUTPUT_FILE)):
            os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

        # 1. Load Mappings/Headers
        header_df = pd.read_excel(HEADER_FILE)
        template_headers = header_df.columns.tolist()

        # 2. Load CSV Data
        df = pd.read_csv(DATA_FILE)

        # Helper to clean numeric data
        def to_num(cols):
            for col in cols:
                if col not in df.columns: df[col] = 0.0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # --- 3. LOGIC CALCULATIONS ---
        
        # MID_WID & Status
        m_id = df['merchant_id'].astype(str)
        w_id = df['warehouse_id'].fillna('').astype(str)
        f_type = df['finance_key_type'].fillna('').astype(str)
        df['MID_WID'] = np.where((f_type == 'default') | (f_type == ''), m_id, m_id + "_" + w_id)

        dir_str = df['direction'].astype(str)
        df['Status'] = np.select([(dir_str == '1'), (dir_str == '2')], ['Payout', 'Returned'], default='Unknown')

        # GMV
        to_num(['price', 'qty_ordered'])
        base_gmv = (df['price'] * df['qty_ordered']).round(2)
        df['GMV'] = np.where(dir_str == '1', base_gmv, np.where(dir_str == '2', -base_gmv, 0))

        # Commission Sum & Tax
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

        # Shipping Reversals
        to_num(['cust_shipping_reversal', 'partial_shipping_rev'])

        # pf, product_gst, TCS, TDS
        to_num(['pf_tax', 'pf_packing', 'pf_seller_convenience', 'product_igst', 'product_cgst', 'product_sgst',
                'tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 'shipping_tcs_sgst', 'tds_ecom'])
        
        df['pf'] = ((df['pf_tax'] + df['pf_packing'] + df['pf_seller_convenience']) * -1).round(2)
        df['product_gst'] = (df['product_igst'] + df['product_cgst'] + df['product_sgst']).round(2)
        df['TCS'] = df[['tcs_cgst', 'tcs_igst', 'tcs_sgst', 'shipping_tcs_cgst', 'shipping_tcs_igst', 'shipping_tcs_sgst']].sum(axis=1).round(2)
        df['TDS'] = df['tds_ecom'].round(2)

        # Merchant Payable (Calculated) & Diff
        df['Merchant Payable'] = (df['GMV'] - df['Commission'] - df['Tax'] - df['cust_shipping_reversal'] - 
                                  df['partial_shipping_rev'] - df['pf'] - df['product_gst'] - df['TCS'] - df['TDS']).round(2)
        
        to_num(['merchant_payable'])
        df['Diff'] = (df['merchant_payable'] - df['Merchant Payable']).round(2)


        # --- 4. CREATE SUMMARY PIVOT TABLE ---
        print("Generating Payout Summary Pivot Table...")
        summary_columns = ['GMV', 'Commission', 'Tax', 'pf', 'product_gst', 'TCS', 'TDS', 'Merchant Payable', 'Diff']
        
        # Group by MID_WID and sum the required columns
        summary_df = df.groupby('MID_WID')[summary_columns].sum().reset_index()
        
        # Export Summary Table
        summary_df.to_excel(SUMMARY_FILE, index=False)
        print(f"✅ SUMMARY SAVED: {SUMMARY_FILE}")


        # --- 5. FINAL EXPORT (Main Data) ---
        final_cols = [c for c in template_headers if c in df.columns]
        
        calc_cols = ['MID_WID', 'Status', 'GMV', 'Commission', 'Tax', 'cust_shipping_reversal', 
                     'partial_shipping_rev', 'pf', 'product_gst', 'TCS', 'TDS', 'Merchant Payable', 'Diff']
        
        for col in calc_cols:
            if col not in final_cols:
                final_cols.append(col)

        # Export Main Report
        df[final_cols].to_excel(OUTPUT_FILE, index=False)
        print(f"✅ MAIN REPORT SAVED: Output with {len(final_cols)} columns saved to {OUTPUT_FILE}")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    process_revenue_report()