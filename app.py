import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Stock Report Cleaner", page_icon="📊")

st.title("📊 Stock Report Cleaner")

uploaded_files = st.file_uploader(
    "Upload yearly reports",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for file in uploaded_files:

        # ---------- Read year from first row ----------
        header = pd.read_excel(file, header=None, nrows=1)

        header_text = " ".join(header.iloc[0].fillna("").astype(str))

        year = ""

        match = re.search(r"Date From\s+\d+/\d+/(\d{4})", header_text)

        if match:
            year = match.group(1)

        # ---------- Read actual table ----------
        df = pd.read_excel(file, header=1)

        # Remove completely empty rows
        df = df.dropna(how="all")

        # Add Year column
        df.insert(0, "Year", year)

        all_data.append(df)

    # Merge all reports
    final_df = pd.concat(all_data, ignore_index=True)

    st.success("Processing Complete!")

    st.dataframe(final_df)

    # Export to Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False)

    st.download_button(
        label="📥 Download Final Report",
        data=output.getvalue(),
        file_name="Final_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
