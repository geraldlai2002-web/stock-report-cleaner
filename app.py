import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="Stock Report Cleaner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Stock Report Cleaner")

uploaded_files = st.file_uploader(
    "Upload yearly reports",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for file in uploaded_files:

        # Read first row to get year
        header = pd.read_excel(file, header=None, nrows=1)
        header_text = " ".join(header.fillna("").astype(str).values.flatten())

        year = ""

        match = re.search(r"Date From\s+\d+/\d+/(\d{4})", header_text)
        if match:
            year = match.group(1)

        # Read data table
        df = pd.read_excel(file, header=1)

        # Remove empty rows
        df = df.dropna(how="all")

        # Remove rows containing "Page" anywhere
        df = df[
            ~df.astype(str)
              .apply(lambda row: row.str.contains("Page", case=False, na=False).any(), axis=1)
        ]

        if "Stk Code" in df.columns:

            df["Stk Code"] = df["Stk Code"].astype(str).str.strip()

            # Remove category rows
            remove_codes = [
                "PMAT",
                "RMAT",
                "FG",
                "WIP",
                "PACKING",
                "OTHERS"
            ]

            df = df[~df["Stk Code"].isin(remove_codes)]

            # Remove Grand Total
            df = df[
                ~df["Stk Code"].str.contains(
                    "Grand",
                    case=False,
                    na=False
                )
            ]

        # Add Year column
        df.insert(0, "Year", year)

        all_data.append(df)

    # Merge all files
    final_df = pd.concat(all_data, ignore_index=True)

    st.success("Processing Complete!")

    st.dataframe(final_df, use_container_width=True)

    # Download Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False)

    st.download_button(
        "📥 Download Final_Report.xlsx",
        output.getvalue(),
        file_name="Final_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
