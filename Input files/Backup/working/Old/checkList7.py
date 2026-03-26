import pandas as pd
import numpy as np
import os

def create_checklist_pivot(OUTPUT_FILE, CHECKLIST_FILE):
    try:

        if not os.path.exists(OUTPUT_FILE):
            print(f"❌ Error: {OUTPUT_FILE} not found.")
            return

        df = pd.read_excel(OUTPUT_FILE)

        # --- Direction Mapping ---
        df['direction_name'] = df['direction'].map(
            {1: 'payout', 2: 'Return', '1': 'payout', '2': 'Return'}
        )
        df['direction_name'] = df['direction_name'].fillna('Unknown')

        # ---------------------------------------------------
        # SHEET1 PIVOT (existing logic)
        # ---------------------------------------------------

        sum_cols = [
            'GMV','merchant_payable','pg_payable2','cod_payable2','Commission','Tax',
            'cust_shipping_reversal','partial_shipping_rev','pf','product_gst','TCS',
            'TDS','Seller Payable','Diff','shipping_amount2','mp_shipping',
            'additional_delivery_charges2','cart_conv_fee2'
        ]

        available_cols = [c for c in sum_cols if c in df.columns]

        pivot_df = pd.pivot_table(
            df,
            index='direction_name',
            values=available_cols,
            aggfunc='sum'
        )

        pivot_df = pivot_df.reindex(columns=available_cols).reset_index()

        # ---------------------------------------------------
        # SHEET2 FIRST PIVOT
        # ---------------------------------------------------

        pivot_sheet2_top = pd.pivot_table(
            df,
            index='direction_name',
            values=['merchant_payable','pg_payable2','cod_payable2'],
            aggfunc='sum'
        ).reset_index()

        pivot_sheet2_top['Difference'] = (
            pivot_sheet2_top['merchant_payable']
            - pivot_sheet2_top['pg_payable2']
            - pivot_sheet2_top['cod_payable2']
        )

        pivot_sheet2_top = pivot_sheet2_top[
            ['direction_name','merchant_payable','pg_payable2','cod_payable2','Difference']
        ]

        # ---------------------------------------------------
        # SHEET2 SECOND PIVOT
        # ---------------------------------------------------

        cols_needed = [
            'GMV','Commission','Tax','cust_shipping_reversal','partial_shipping_rev',
            'pf','product_gst','TCS','TDS','Seller Payable','merchant_payable'
        ]

        for c in cols_needed:
            if c not in df.columns:
                df[c] = 0

        pivot_sheet2_bottom = pd.pivot_table(
            df,
            index='direction_name',
            values=cols_needed,
            aggfunc='sum'
        ).reset_index()

        pivot_sheet2_bottom.rename(
            columns={'partial_shipping_rev':'partial_shipping_reversal'},
            inplace=True
        )

        # --- Diff calculation (unchanged logic) ---
        pivot_sheet2_bottom['Diff'] = (
            pivot_sheet2_bottom['Seller Payable']
            - pivot_sheet2_bottom['merchant_payable']
        )

        pivot_sheet2_bottom = pivot_sheet2_bottom[
            [
                'direction_name','GMV','Commission','Tax',
                'cust_shipping_reversal','partial_shipping_reversal',
                'pf','product_gst','TCS','TDS',
                'Seller Payable','merchant_payable','Diff'
            ]
        ]

        # ---------------------------------------------------
        # WRITE EXCEL (FORMAT LIKE IMAGE)
        # ---------------------------------------------------

        with pd.ExcelWriter(CHECKLIST_FILE, engine='openpyxl') as writer:

            sheet_name = "Sheet2"

            # Title 1
            title1 = pd.DataFrame({"A": ["1  Check list _merchant_payable"]})
            title1.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            # First Pivot
            pivot_sheet2_top.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=2,
                index=False
            )

            # Title 2
            title2_row = len(pivot_sheet2_top) + 5
            title2 = pd.DataFrame({"A": ["2  Check list _Payout Calculation"]})
            title2.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=title2_row,
                index=False,
                header=False
            )

            # Second Pivot
            pivot_sheet2_bottom.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=title2_row + 2,
                index=False
            )

            # Other sheets unchanged
            pivot_df.to_excel(writer, sheet_name='Sheet1', index=False)
            pd.DataFrame().to_excel(writer, sheet_name='Sheet3', index=False)
            pd.DataFrame().to_excel(writer, sheet_name='Sheet4', index=False)

        print("✅ Checklist.xlsx created successfully")

    except Exception as e:
        print(f"❌ Error creating checklist: {e}")