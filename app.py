import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(layout="wide", page_title="QM Cards Reconciliation Report")

st.title("QM Cards Recon VA - Reconciliation Report Generator")

st.markdown("""
**Instructions:**  
Upload the following files (as originally required in Colab):  
- KCB Excel file (e.g. `QUICK MART 11.6.2025.xlsx`)
- Equity Excel file (e.g. `QUICKMART 11062025.xlsx`)
- Aspire CSV file (e.g. `ZEDS_CARDS_TILLWISE_2025-06-11.csv`)
- Card Key Excel file (e.g. `card_key.xlsx`)
""")

kcb_file = st.file_uploader("Upload KCB Excel", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel", type=["xlsx"])
aspire_file = st.file_uploader("Upload Aspire CSV", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel", type=["xlsx"])

if st.button("Generate Reconciliation Report"):

    if not (kcb_file and equity_file and aspire_file and key_file):
        st.error("Please upload all four required files before proceeding.")
        st.stop()

    # --- Load all files ---
    kcb = pd.read_excel(kcb_file)
    equity = pd.read_excel(equity_file)
    aspire = pd.read_csv(aspire_file)
    key = pd.read_excel(key_file)

    # --- Data Cleaning and Alignment ---
    kcb.columns = kcb.columns.str.strip()
    equity.columns = equity.columns.str.strip()

    kcb_renamed = kcb.rename(columns={
        'Card No': 'Card_Number',
        'Trans Date': 'TRANS_DATE',
        'RRN': 'R_R_N',
        'Amount': 'Purchase',
        'Comm': 'Commission',
        'NetPaid': 'Settlement_Amount',
        'Merchant': 'store'
    })
    kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -1 * x if x < 0 else 0)
    kcb_renamed['Source'] = 'KCB'
    equity = equity.rename(columns={'Outlet_Name': 'store'})
    equity['Source'] = 'Equity'
    columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
               'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
    kcb_final = kcb_renamed[columns]
    equity_final = equity[columns]
    merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)

    # Drop rows where Card_Number is NaN or empty string
    merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
    merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']

    # --- Key/branch mapping ---
    card_key = pd.read_excel(key_file)
    card_key.columns = card_key.columns.str.strip()
    card_key['Col_1'] = card_key['Col_1'].astype(str).str.strip()
    card_key['Col_2'] = card_key['Col_2'].astype(str).str.strip()
    merged_cards['store'] = merged_cards['store'].astype(str).str.strip()
    lookup_dict = dict(zip(card_key['Col_1'], card_key['Col_2']))
    merged_cards['branch'] = merged_cards['store'].map(lookup_dict)
    cols = list(merged_cards.columns)
    source_index = cols.index('Source')
    cols.insert(source_index + 1, cols.pop(cols.index('branch')))
    merged_cards = merged_cards[cols]

    # --- Card checks ---
    merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
    merged_cards['card_check'] = merged_cards['Card_Number'].apply(
        lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
    )
    cols = merged_cards.columns.tolist()
    if 'branch' in cols and 'card_check' in cols:
        cols.remove('card_check')
        branch_index = cols.index('branch')
        cols.insert(branch_index + 1, 'card_check')
        merged_cards = merged_cards[cols]

    # --- Aspire checks ---
    aspire['CARD_NUMBER'] = aspire['CARD_NUMBER'].astype(str).str.strip()
    aspire['card_check'] = aspire['CARD_NUMBER'].apply(
        lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
    )
    cols = aspire.columns.tolist()
    if 'CARD_NUMBER' in cols and 'card_check' in cols:
        cols.remove('card_check')
        insert_index = cols.index('CARD_NUMBER') + 1
        cols.insert(insert_index, 'card_check')
        aspire = aspire[cols]
    aspire = aspire[[
        'STORE_CODE', 'STORE_NAME', 'ZED_DATE', 'TILL', 'SESSION', 'RCT', 'CUSTOMER_NAME',
        'CARD_TYPE', 'CARD_NUMBER', 'card_check', 'AMOUNT', 'REF_NO', 'RCT_TRN_DATE'
    ]]
    aspire = aspire.rename(columns={'REF_NO': 'R_R_N'})
    aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
    merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()

    # --- Merge on R_R_N ---
    rrntable = pd.merge(
        aspire,
        merged_cards,
        on='R_R_N',
        how='inner',
        suffixes=('_aspire', '_merged')
    )

    # --- Card summary (Aspire, KCB, Equity) ---
    aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
    card_summary = (
        aspire['STORE_NAME'].dropna().drop_duplicates().sort_values().reset_index(drop=True).to_frame(name='STORE_NAME')
    )
    card_summary.index = card_summary.index + 1
    card_summary.reset_index(inplace=True)
    card_summary.rename(columns={'index': 'No'}, inplace=True)

    aspire_sums = aspire.groupby('STORE_NAME')['AMOUNT'].sum().reset_index().rename(columns={'AMOUNT': 'Aspire_Zed'})
    card_summary = card_summary.merge(aspire_sums, on='STORE_NAME', how='left')
    card_summary['Aspire_Zed'] = card_summary['Aspire_Zed'].fillna(0)

    merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')

    kcb_grouped = (
        merged_cards[merged_cards['Source'] == 'KCB']
        .groupby('branch')['Purchase'].sum().reset_index()
        .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'kcb_paid'})
    )
    card_summary = card_summary.merge(kcb_grouped, on='STORE_NAME', how='left')
    card_summary['kcb_paid'] = card_summary['kcb_paid'].fillna(0)

    equity_grouped = (
        merged_cards[merged_cards['Source'] == 'Equity']
        .groupby('branch')['Purchase'].sum().reset_index()
        .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'equity_paid'})
    )
    card_summary = card_summary.merge(equity_grouped, on='STORE_NAME', how='left')
    card_summary['equity_paid'] = card_summary['equity_paid'].fillna(0)
    cols = list(card_summary.columns)
    if 'kcb_paid' in cols and 'equity_paid' in cols:
        kcb_index = cols.index('kcb_paid')
        cols.insert(kcb_index + 1, cols.pop(cols.index('equity_paid')))
        card_summary = card_summary[cols]

    # --- Totals row and formatting ---
    card_summary['kcb_paid'] = card_summary['kcb_paid'].replace({',': ''}, regex=True).astype(float)
    card_summary['equity_paid'] = card_summary['equity_paid'].replace({',': ''}, regex=True).astype(float)
    card_summary['Gross_Banking'] = card_summary['kcb_paid'] + card_summary['equity_paid']
    numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking']
    for col in numeric_cols:
        card_summary[col] = card_summary[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float, np.float64)) else x)
    totals = {
        'No': '',
        'STORE_NAME': 'TOTAL',
        'Aspire_Zed': f"{card_summary[:-1]['Aspire_Zed'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
        'kcb_paid': f"{card_summary[:-1]['kcb_paid'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
        'equity_paid': f"{card_summary[:-1]['equity_paid'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
        'Gross_Banking': f"{card_summary[:-1]['Gross_Banking'].replace({',': ''}, regex=True).astype(float).sum():,.2f}"
    }
    card_summary.iloc[-1] = totals

    # --- Prepare for export ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        card_summary.to_excel(writer, sheet_name='card_summary', index=False)
        rrntable.to_excel(writer, sheet_name='RNN_Merge', index=False)
        merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
        aspire.to_excel(writer, sheet_name='aspire', index=False)
        # Add any other DataFrames as new sheets as per your Colab logic

    st.success("✅ All reports generated successfully! Download below:")

    st.download_button(
        label="Download Reconciliation_Report.xlsx",
        data=output.getvalue(),
        file_name="Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write("Preview of Card Summary:")
    st.dataframe(card_summary.head(10))

st.markdown("---")
st.markdown("© 2025 QM Cards Recon VA")
