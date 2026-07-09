import streamlit as st
import pandas as pd
import re
from io import BytesIO
import time

st.set_page_config(
    page_title="Stock Report Cleaner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Stock Report Cleaner")
st.markdown("### Upload yearly inventory reports for Power BI")

uploaded_files = st.file_uploader(
    "📂 Drag and drop Excel files here or click Browse Files",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Upload all yearly reports (2015 - 2026)"
)

if uploaded_files:

    st.success(f"✅ {len(uploaded_files)} file(s) selected")

    st.markdown("### Selected Files")

    file_list = pd.DataFrame({
        "No.": range(1, len(uploaded_files)+1),
        "File Name": [f.name for f in uploaded_files]
    })

    st.dataframe(file_list, use_container_width=True)

    if st.button("🚀 Process Reports"):

        progress = st.progress(0)

        all_data = []
        years = []

        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):

            # Read first row
            header = pd.read_excel(file, header=None, nrows=1)

            header_text = " ".join(
                header.fillna("").astype(str).values.flatten()
            )

            year = ""

            match = re.search(
                r"Date From\s+\d+/\d+/(\d{4})",
                header_text
            )

            if match:
                year = match.group(1)
                years.append(int(year))

            # Read report
            df = pd.read_excel(file, header=1)

            # Remove blank rows
            df = df.dropna(how="all")

            # Remove rows containing "Page"
            df = df[
                ~df.astype(str)
                .apply(
                    lambda row: row.str.contains(
                        "Page",
                        case=False,
                        na=False
                    ).any(),
                    axis=1
                )
            ]

            if "Stk Code" in df.columns:

                df["Stk Code"] = (
                    df["Stk Code"]
                    .astype(str)
                    .str.strip()
                )

                remove_codes = [
                    "PMAT",
                    "RMAT",
                    "FG",
                    "WIP",
                    "PACKING",
                    "OTHERS"
                ]

                df = df[
                    ~df["Stk Code"].isin(remove_codes)
                ]

                df = df[
                    ~df["Stk Code"].str.contains(
                        "Grand",
                        case=False,
                        na=False
                    )
                ]

            df.insert(0, "Year", year)

            all_data.append(df)

            progress.progress((i + 1) / total_files)

            time.sleep(0.1)

        final_df = pd.concat(all_data, ignore_index=True)

        if years:
            final_df = final_df.sort_values("Year")

        st.success("✅ Processing Complete!")

        st.markdown("## 📊 Processing Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric("Files Processed", len(uploaded_files))
        col2.metric("Rows Processed", len(final_df))
        col3.metric(
            "Years",
            f"{min(years)} - {max(years)}"
        )

        st.markdown("### Preview")

        st.dataframe(
            final_df,
            use_container_width=True
        )

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            final_df.to_excel(
                writer,
                index=False
            )

        st.download_button(
            "📥 Download Final Report",
            output.getvalue(),
            file_name=f"Stock_Report_{min(years)}_{max(years)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
