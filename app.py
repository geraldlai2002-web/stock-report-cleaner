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
st.write("Prepare yearly inventory reports for Power BI")

uploaded_files = st.file_uploader(
    "📂 Drag & Drop Excel Files Here",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.success(f"✅ {len(uploaded_files)} file(s) selected")

    file_df = pd.DataFrame({
        "No": range(1, len(uploaded_files) + 1),
        "File Name": [f.name for f in uploaded_files]
    })

    st.subheader("Uploaded Files")
    st.dataframe(file_df, use_container_width=True)

   if st.button("🚀 Process Reports"):
    st.session_state["processed"] = True

  if st.session_state.get("processed", False):

    progress = st.progress(0)
    status = st.empty()

        all_data = []
        years = []

        total_rows_before = 0
        total_rows_after = 0

        total_files = len(uploaded_files)

        for index, file in enumerate(uploaded_files):

            status.write(f"Processing {file.name}...")

            # Read first row to detect year
            header = pd.read_excel(
                file,
                header=None,
                nrows=1
            )

            header_text = " ".join(
                header.fillna("").astype(str).values.flatten()
            )

            match = re.search(
                r"Date From\s+\d+/\d+/(\d{4})",
                header_text
            )

            if not match:
                st.error(f"Cannot detect year from {file.name}")
                st.stop()

            year = int(match.group(1))

            if year in years:
                st.error(f"❌ Duplicate Year Detected: {year}")
                st.stop()

            years.append(year)

            # Read report
            df = pd.read_excel(
                file,
                header=1
            )

            total_rows_before += len(df)

            # Remove blank rows
            df = df.dropna(how="all")

            # Remove page footer
            df = df[
                ~df.astype(str).apply(
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

            df.insert(
                0,
                "Year",
                year
            )

            total_rows_after += len(df)

            all_data.append(df)

            progress.progress(
                (index + 1) / total_files
            )

            time.sleep(0.1)

        status.empty()

        # Missing year check
        if len(years) > 1:

            expected = list(
                range(
                    min(years),
                    max(years) + 1
                )
            )

            missing = sorted(
                set(expected) - set(years)
            )

            if missing:

                st.warning(
                    "⚠ Missing Year(s): "
                    + ", ".join(map(str, missing))
                )

        final_df = pd.concat(
            all_data,
            ignore_index=True
        )

        final_df = final_df.sort_values(
            "Year"
        )

        st.session_state["final_df"] = final_df

        st.success("✅ Processing Complete")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Files",
            len(uploaded_files)
        )

        col2.metric(
            "Rows",
            len(final_df)
        )

        col3.metric(
            "Years",
            f"{min(years)} - {max(years)}"
        )

        # -----------------------------
        # Search Stock Code
        # -----------------------------
        st.markdown("---")

        if "final_df" not in st.session_state:
            st.stop()

        final_df = st.session_state["final_df"]

        st.subheader("🔍 Search Stock Code")

        search_code = st.text_input(
            "Enter Stock Code (Example: D10001)"
        )

        if search_code:

            if "Stk Code" not in final_df.columns:
                st.error("Column 'Stk Code' not found.")

            else:

                result = final_df[
                    final_df["Stk Code"]
                    .astype(str)
                    .str.upper()
                    .str.strip()
                    .str.contains(
                        search_code.upper().strip(),
                        na=False
                    )
                ]

                if result.empty:

                    st.warning("No record found.")

                else:

                    st.success(
                        f"Found {len(result)} record(s)."
                    )

                    st.dataframe(
                        result,
                        use_container_width=True,
                        height=350
                    )

                    output_search = BytesIO()

                    with pd.ExcelWriter(
                        output_search,
                        engine="openpyxl"
                    ) as writer:

                        result.to_excel(
                            writer,
                            index=False
                        )

                    st.download_button(
                        "📤 Export Search Result",
                        output_search.getvalue(),
                        file_name=f"{search_code.upper()}_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # -----------------------------
        # Complete Report
        # -----------------------------
        st.markdown("---")
        st.subheader("📊 Complete Report")

        st.dataframe(
            final_df,
            use_container_width=True,
            height=450
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

        # -----------------------------
        # Processing Log
        # -----------------------------
        st.markdown("---")
        st.subheader("📋 Processing Log")

        removed_rows = total_rows_before - total_rows_after

        log_df = pd.DataFrame({
            "Item": [
                "Files Uploaded",
                "Rows Before Cleaning",
                "Rows Removed",
                "Rows After Cleaning",
                "Years",
                "Status"
            ],
            "Value": [
                len(uploaded_files),
                total_rows_before,
                removed_rows,
                total_rows_after,
                f"{min(years)} - {max(years)}",
                "SUCCESS"
            ]
        })

        st.dataframe(
            log_df,
            use_container_width=True,
            hide_index=True
        )

st.markdown("---")
st.caption("📊 Stock Report Cleaner v3.0")
st.caption("Developed by JUN THANG LAI")
st.caption("Internship Project")
