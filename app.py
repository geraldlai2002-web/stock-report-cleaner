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

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

uploaded_files = st.file_uploader(
    "📂 Drag & Drop Excel Files Here",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)


@st.cache_data(show_spinner=False)
def process_files(files):
    """
    Reads, cleans, and combines all uploaded Excel reports.

    Each report has section markers embedded in the 'Stk Code' column
    (e.g. PMAT, CCTN, HH01...) — every real item between one marker and
    the next belongs to that warehouse/category. Each marker also
    reappears once more as a subtotal row for its section.

    Real item rows always have a value in 'Uom'; marker/subtotal/grand
    total/page-footer rows never do. That's a structural fact, not a
    name we have to know in advance, so it works regardless of which
    warehouse or how many different section codes a file uses.

    We use that fact to:
      1. Tag every real item row with its 'Warehouse' (carried forward
         from the marker row above it).
      2. Drop every non-item row (markers, subtotals, grand total,
         page footer) from the final data.
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

        # Remove fully blank rows
        df = df.dropna(how="all")

        if "Stk Code" not in df.columns or "Uom" not in df.columns:
            raise ValueError(
                f"{file.name} is missing an expected column "
                f"('Stk Code' or 'Uom') — check the report format."
            )

        df["Stk Code"] = (
            df["Stk Code"]
            .astype(str)
            .str.strip()
        )

        # A row with no Uom is never a real stock item — it's a section
        # marker, a section subtotal, the grand total, or the page footer.
        marker_mask = df["Uom"].isna()

        # The grand total / page footer rows are also marker rows, but
        # they aren't warehouse names, so they shouldn't be carried
        # forward as one.
        footer_mask = (
            df["Stk Code"].str.contains("Grand", case=False, na=False)
            | df["Stk Code"].str.contains("Page", case=False, na=False)
        )

        warehouse_labels = df["Stk Code"].where(marker_mask & ~footer_mask)

        # Carry the last-seen warehouse name down onto every item row
        # below it, until the next marker changes it.
        df["Warehouse"] = warehouse_labels.ffill()

        # Now drop every non-item row — markers, subtotals, grand total,
        # page footer — all identified by having no Uom.
        df = df[~marker_mask]

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

    if st.button("🚀 Process Reports"):
        st.session_state["processed"] = True

    if st.session_state.get("processed", False):

        # process_files is cached, so this only actually does work the
        # first time (or when the uploaded files change) — every later
        # rerun of the script (search box, checkboxes, etc.) just reuses
        # the cached result instantly.
        try:
            with st.spinner("Processing reports..."):
                final_df, years, total_rows_before, total_rows_after = process_files(
                    uploaded_files
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
        # Warehouse Filter
        # -----------------------------
        st.markdown("---")
        st.subheader("🏬 Warehouse Filter")

        all_warehouses = sorted(
            final_df["Warehouse"].dropna().unique().tolist()
        )

        st.caption(
            f"Detected {len(all_warehouses)} warehouse(s)/categor(ies) "
            f"across the uploaded files: {', '.join(all_warehouses)}"
        )

        selected_warehouses = st.multiselect(
            "Include warehouses",
            options=all_warehouses,
            default=all_warehouses
        )

        final_df = final_df[
            final_df["Warehouse"].isin(selected_warehouses)
        ]

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
