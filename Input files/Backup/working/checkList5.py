import pandas as pd
import numpy as np
import os

# --- CONFIGURATION ---
# OUTPUT_FILE = r'output_files\output.xlsx'
# CHECKLIST_FILE = r'output_files\checklist_file.xlsx'

def create_checklist_pivot(OUTPUT_FILE, CHECKLIST_FILE):
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

        # Filter to ensure we only sum columns that actually exist
        available_cols = [col for col in sum_cols if col in df.columns]

        # 5. Create Pivot Table for Sheet1
        pivot_df = pd.pivot_table(
            df,
            index='direction_name',
            values=available_cols,
            aggfunc='sum'
        )

        pivot_df = pivot_df.reindex(columns=available_cols).reset_index()

        # -------------------------------------------------
        # NEW PIVOT FOR SHEET2
        # -------------------------------------------------

        pivot_sheet2 = pd.pivot_table(
            df,
            index='direction_name',
            values=['merchant_payable','pg_payable2','cod_payable2'],
            aggfunc='sum'
        ).reset_index()

        pivot_sheet2['Difference'] = (
            pivot_sheet2['merchant_payable']
            - pivot_sheet2['pg_payable2']
            - pivot_sheet2['cod_payable2']
        )

        pivot_sheet2 = pivot_sheet2[
            ['direction_name','merchant_payable','pg_payable2','cod_payable2','Difference']
        ]

        # 6. Save to checklist Excel
        with pd.ExcelWriter(CHECKLIST_FILE, engine='openpyxl') as writer:

            # Sheet1 (existing pivot)
            pivot_df.to_excel(writer, sheet_name='Sheet1', index=False)

            # Sheet2 (new reconciliation pivot)
            pivot_sheet2.to_excel(writer, sheet_name='Sheet2', index=False)

            # Blank sheets
            pd.DataFrame().to_excel(writer, sheet_name='Sheet3', index=False)
            pd.DataFrame().to_excel(writer, sheet_name='Sheet4', index=False)

        print("-" * 40)
        print(f"✅ SUCCESS! Checklist pivot created at: {CHECKLIST_FILE}")
        print(f"Sequence: {', '.join(available_cols)}")
        print("Sheets created: Sheet1, Sheet2, Sheet3, Sheet4")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error creating checklist: {e}")

# if __name__ == "__main__":
#     create_checklist_pivot()