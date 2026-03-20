import pandas as pd
import numpy as np
import os

# --- CONFIGURATION ---
OUTPUT_FILE = r'output_files\output.xlsx'
CHECKLIST_FILE = r'output_files\checklist_file.xlsx'

def create_checklist_pivot():
    try:
        # 1. Check if the output file exists
        if not os.path.exists(OUTPUT_FILE):
            print(f"❌ Error: {OUTPUT_FILE} not found. Please run the main processing script first.")
            return

        # 2. Load the processed output data
        df = pd.read_excel(OUTPUT_FILE)

        # 3. Update Direction Names: 1 -> payout, 2 -> Return
        df['direction_name'] = df['direction'].map({1: 'payout', 2: 'Return', '1': 'payout', '2': 'Return'})
        df['direction_name'] = df['direction_name'].fillna('Unknown')

        # 4. List of columns to sum in the exact sequence requested
        sum_cols = [
            'GMV', 
            'merchant_payable',
            'pg_payable2', 
            'cod_payable2', 
            'Commission', 
            'Tax', 
            'cust_shipping_reversal', 
            'partial_shipping_rev', 
            'pf', 
            'product_gst', 
            'TCS', 
            'TDS', 
            'Seller Payable',
            'Diff',
            'shipping_amount2', 
            'mp_shipping', 
            'additional_delivery_charges2', 
            'cart_conv_fee2'
        ]

        # Filter to ensure we only sum columns that actually exist in output.xlsx
        available_cols = [col for col in sum_cols if col in df.columns]

        # 5. Create the Pivot Table
        # We use the list sequence to define the column order in the final pivot
        pivot_df = pd.pivot_table(
            df, 
            index='direction_name', 
            values=available_cols, 
            aggfunc='sum'
        )
        
        # Reorder columns to match your requested sequence exactly
        pivot_df = pivot_df.reindex(columns=available_cols).reset_index()

        # 6. Save to the checklist file
        pivot_df.to_excel(CHECKLIST_FILE, index=False)
        
        print("-" * 40)
        print(f"✅ SUCCESS! Checklist pivot created at: {CHECKLIST_FILE}")
        print(f"Sequence: {', '.join(available_cols)}")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error creating checklist: {e}")

if __name__ == "__main__":
    create_checklist_pivot()