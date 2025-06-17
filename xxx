import pandas as pd
from io import BytesIO
import zipfile

def main(uploaded_zip):
    # Extract and load files from uploaded ZIP
    with zipfile.ZipFile(uploaded_zip) as z:
        kcb = pd.read_excel(z.open([f for f in z.namelist() if "QUICK MART" in f][0]))
        equity = pd.read_excel(z.open([f for f in z.namelist() if "QUICKMART" in f][0]))
        aspire = pd.read_csv(z.open([f for f in z.namelist() if "ZEDS_CARDS" in f][0]))
        key = pd.read_excel(z.open([f for f in z.namelist() if "card_key" in f.lower()][0]))

    # Sample logic (replace with your real processing steps)
    kcb['Net_Variance'] = kcb.get('Variance', 0) - kcb.get('KCB_Recs', 0) - kcb.get('Equity_Recs', 0) - kcb.get('Asp_Recs', 0)

    # Prepare downloadable Excel output
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        kcb.to_excel(writer, index=False, sheet_name='KCB')
        equity.to_excel(writer, index=False, sheet_name='Equity')
        aspire.to_excel(writer, index=False, sheet_name='Aspire')
        key.to_excel(writer, index=False, sheet_name='Card_Key')
    output.seek(0)

    return kcb, output
