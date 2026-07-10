import streamlit as st
import pandas as pd
import re
from io import BytesIO
import zipfile

st.set_page_config(
    page_title="Stock Report Cleaner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Stock Report Cleaner")
st.write("Prepare yearly inventory reports for Power BI")

DEFAULT_REMOVE_CODES = [
    "PMAT",
    "RMAT",
    "FG",
    "WIP",
    "PACKING",
    "OTHERS"
]

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

uploaded_files = st.file_uploader(
    "📂 Drag & Drop Excel Files Here",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)


@st.cache_data(show_spinner=False)
def process_files(files, remove_codes):
    """
    Reads, cleans, and combines all uploaded Excel reports.
    Cached on (file content, remove_codes), so re-running the app (e.g.
    typing in the search box, toggling a checkbox) will NOT re-read/
    re-clean the files — it only recomputes if the uploaded files OR
    the exclusion list changes.

    remove_codes: tuple of stock codes to exclude (tuple, not list,
    so it's hashable and works as a cache key).
    """

    all_data = []
    years = []

    total_rows_before = 0
    total_rows_after = 0

    for file in files:

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
            raise ValueError(f"Cannot detect year from {file.name}")

        year = int(match.group(1))

        if year in years:
            raise ValueError(f"Duplicate Year Detected: {year}")

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

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    final_df = final_df.sort_values("Year")

    return final_df, years, total_rows_before, total_rows_after


if uploaded_files:

    col1, col2 = st.columns([4, 1])

    with col1:
        st.success(f"✅ {len(uploaded_files)} file(s) selected")

    with col2:
        if st.button("🔄 Reset"):

            st.session_state["uploader_key"] += 1
            st.session_state.pop("processed", None)
            process_files.clear()

            st.rerun()

    file_df = pd.DataFrame({
        "No": range(1, len(uploaded_files) + 1),
        "File Name": [f.name for f in uploaded_files]
    })

    st.subheader("Uploaded Files")
    st.dataframe(file_df, use_container_width=True)

    # -----------------------------
    # Configurable Exclusion List
    # -----------------------------
    st.subheader("⚙️ Stock Code Exclusion Settings")

    st.caption(
        "Rows whose 'Stk Code' matches any of the selected codes below "
        "will be removed from the cleaned report."
    )

    selected_codes = st.multiselect(
        "Codes to exclude",
        options=DEFAULT_REMOVE_CODES,
        default=DEFAULT_REMOVE_CODES
    )

    custom_codes_input = st.text_input(
        "Add custom code(s) to exclude (comma-separated, optional)",
        placeholder="e.g. SCRAP, SAMPLE"
    )

    custom_codes = [
        code.strip().upper()
        for code in custom_codes_input.split(",")
        if code.strip()
    ]

    remove_codes = tuple(
        sorted(set(selected_codes) | set(custom_codes))
    )

    if remove_codes:
        st.caption(f"Currently excluding: {', '.join(remove_codes)}")
    else:
        st.caption("No codes selected — no rows will be excluded by code.")

    if st.button("🚀 Process Reports"):
        st.session_state["processed"] = True

    if st.session_state.get("processed", False):

        # process_files is cached, so this only actually does work the
        # first time (or when the uploaded files / remove_codes change) —
        # every later rerun of the script (search box, checkboxes, etc.)
        # just reuses the cached result instantly.
        try:
            with st.spinner("Processing reports..."):
                final_df, years, total_rows_before, total_rows_after = process_files(
                    uploaded_files, remove_codes
                )
        except ValueError as e:
            st.error(f"❌ {e}")
            st.stop()

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
                        file_name=f"{search_code.upper()}_{min(years)}-{max(years)}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # -----------------------------
        # Complete Report
        # -----------------------------
        st.subheader("📥 Download Reports")
        col1, col2 = st.columns(2)
        with col1:
            download_merged = st.checkbox(
                "📄 Merged Excel (.xlsx)",
                value=True
            )
        with col2:
            download_separate = st.checkbox(
                "📦 Separate Excel Files (.zip)"
            )

        st.markdown("---")
        st.subheader("📊 Preview Report")

        st.dataframe(
            final_df,
            use_container_width=True,
            height=450
        )

        # Export
        if download_merged:

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
                "📥 Download Merged Report",
                output.getvalue(),
                file_name=f"Stock_Report_{min(years)}-{max(years)}_Merged.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if download_separate:

            zip_buffer = BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for year in sorted(final_df["Year"].unique()):

                    yearly_df = final_df[
                        final_df["Year"] == year
                    ]

                    excel_buffer = BytesIO()

                    with pd.ExcelWriter(
                        excel_buffer,
                        engine="openpyxl"
                    ) as writer:

                        yearly_df.to_excel(
                            writer,
                            index=False
                        )

                    zip_file.writestr(
                        f"Stock_Report_{year}.xlsx",
                        excel_buffer.getvalue()
                    )

            st.download_button(
                "📦 Download Separate Reports (.zip)",
                zip_buffer.getvalue(),
                file_name=f"Stock_Report_{min(years)}-{max(years)}_Separate.zip",
                mime="application/zip"
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
st.caption("Developed by Gerald")
st.caption("Internship Project")

