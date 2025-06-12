import streamlit as st
import pandas as pd
from io import BytesIO
from app_logic import main  # assuming your notebook logic is in a 'main()' function

st.set_page_config(page_title="Card Reconciliation App", layout="wide")

st.title("🧾 Card Reconciliation Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    st.success("File uploaded. Processing...")

    try:
        # Process the uploaded Excel file
        result_df, excel_bytes = main(uploaded_file)

        st.subheader("✅ Sample Output")
        st.dataframe(result_df.head(10))

        st.download_button(
            label="📥 Download Full Reconciliation Report",
            data=excel_bytes,
            file_name="Card_Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ An error occurred while processing the file: {e}")
else:
    st.info("Please upload your latest card Excel dump to generate the report.")