import streamlit as st

st.set_page_config(
    page_title="Stock Report Cleaner",
    page_icon="📊"
)

st.title("📊 Stock Report Cleaner")

st.write("Upload your yearly reports (2015-2026)")

uploaded_files = st.file_uploader(
    "Choose Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write("### Uploaded Files")

    for file in uploaded_files:
        st.write(file.name)

if st.button("Process"):
    st.success("Processing...")
