import streamlit as st
import pandas as pd
from openpyxl import Workbook
import zipfile
import struct
import io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="⚡ Electrochemistry Extractor",
    layout="wide",
    page_icon="🧪"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
.block-container {padding-top: 2rem;}
.stButton>button {
    border-radius: 10px;
    height: 3em;
    font-weight: 600;
}
.stDownloadButton>button {
    border-radius: 10px;
    height: 3em;
    background-color: #4CAF50;
    color: white;
}
.file-box {
    padding: 10px;
    border-radius: 10px;
    background-color: #f5f7fa;
    margin-bottom: 5px;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.title("🧪 Electrochemistry Data Extractor")
st.caption("Extract CV & GCD data from `.xlsb` files in seconds")

# ---------------- SESSION STATE ----------------
if "files" not in st.session_state:
    st.session_state.files = []

if "processed_file" not in st.session_state:
    st.session_state.processed_file = None

# ---------------- MODE SELECTION ----------------
col1, col2 = st.columns([1, 2])

with col1:
    option = st.radio("⚙️ Select Mode", ("CV", "GCD"))

with col2:
    uploaded_files = st.file_uploader(
        "📂 Upload .xlsb files",
        type=["xlsb"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.session_state.files = uploaded_files

# ---------------- FILE PREVIEW ----------------
if st.session_state.files:
    st.subheader(f"📁 Uploaded Files ({len(st.session_state.files)})")

    for f in st.session_state.files:
        st.markdown(f"<div class='file-box'>📄 {f.name}</div>", unsafe_allow_html=True)

    colA, colB = st.columns([1, 1])

    with colA:
        if st.button("🗑 Remove All Files"):
            st.session_state.files = []
            st.rerun()

    with colB:
        process_btn = st.button("⚡ Process Files")

# ---------------- PROCESSING ----------------
if st.session_state.files and 'process_btn' in locals() and process_btn:

    results = []
    progress_bar = st.progress(0)
    status = st.empty()

    for i, file in enumerate(st.session_state.files):
        try:
            status.info(f"Processing: {file.name}")

            if option == "CV":
                df = process_cv_file(file)
            else:
                df = process_gcd_file(file)

            if df is not None and not df.empty:
                results.append((file.name, df))
            else:
                st.warning(f"⚠️ No valid data in {file.name}")

        except Exception as e:
            st.error(f"❌ Error in {file.name}: {e}")

        progress_bar.progress((i + 1) / len(st.session_state.files))

    # ---------------- OUTPUT ----------------
    if results:
        status.success("✅ Processing complete! Preparing download...")

        output = io.BytesIO()

        if option == "CV":
            final_df = pd.concat([res[1] for res in results], axis=1)
            wb = Workbook()
            ws = wb.active
            ws.title = "Combined Data"

            filenames = [res[0] for res in results]

            row1 = []
            for fn in filenames:
                row1.extend([fn, ""])
            ws.append(row1)

            row2 = []
            for _ in filenames:
                row2.extend(["Voltage_V", "_Current_A"])
            ws.append(row2)

            final_df.fillna("", inplace=True)
            for r in final_df.itertuples(index=False, name=None):
                ws.append(r)

            wb.save(output)

        else:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                wb = writer.book
                ws = wb.create_sheet("Extracted Data")

                col_start = 1
                for file_name, df_data in results:
                    ws.cell(row=1, column=col_start, value=file_name)
                    ws.cell(row=2, column=col_start, value="StepTime_s")
                    ws.cell(row=2, column=col_start + 1, value="Voltage_V")

                    for row_idx, (st_val, vv) in enumerate(
                        zip(df_data["StepTime_s"], df_data["Voltage_V"]), start=3
                    ):
                        ws.cell(row=row_idx, column=col_start, value=round(float(st_val), 4))
                        ws.cell(row=row_idx, column=col_start + 1, value=round(float(vv), 6))

                    col_start += 3

                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]

        output.seek(0)
        st.session_state.processed_file = output

# ---------------- AUTO DOWNLOAD ----------------
if st.session_state.processed_file:
    st.success("📥 Your file is ready!")

    st.download_button(
        label="⬇️ Download Now",
        data=st.session_state.processed_file,
        file_name=f"Combined_{option}_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download"
    )

    # AUTO DOWNLOAD TRIGGER (JS)
    st.markdown("""
        <script>
        const button = window.parent.document.querySelector('[data-testid="stDownloadButton"] button');
        if (button) { button.click(); }
        </script>
    """, unsafe_allow_html=True)
