import streamlit as st
import pandas as pd
from io import BytesIO
import model  # your local model.py file

# --- Page setup ---
st.set_page_config(page_title="Chấm Điểm Vay Margin", layout="centered")
st.title("📊 Chấm Điểm Tự Động")
st.write("Upload your Excel file, choose a sheet, and automatically process it.")

# --- File upload ---
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Load all sheet names
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        # Let user select a sheet
        sheet_choice = st.selectbox("Select a sheet to process:", sheet_names)

        # Ask user for save name

        if st.button("🚀 Run Processing"):
            # Read chosen sheet
            df = pd.read_excel(uploaded_file, sheet_name=sheet_choice)

            # Run your model’s processing
            if hasattr(model, "add_score"):
                df_result = model.add_score(df)
            else:
                df_result = df

            # Prepare result for download
            buffer = BytesIO()
            df_result.to_excel(buffer, index=False)
            buffer.seek(0)

            st.success("✅ File processed successfully!")
            st.download_button(
                label="📥 Download Processed File",
                data=buffer,
                file_name=f"xep_hang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Something went wrong:\n\n{e}")
