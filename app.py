import streamlit as st
import pandas as pd
from openpyxl import Workbook
import zipfile
import struct
import io
import qrcode
from io import BytesIO
import base64

# ---------------------- PAGE CONFIG ----------------------
st.set_page_config(page_title="Electrochemistry Extractor", layout="centered")

# ---------------------- MODERN CSS ----------------------
st.markdown("""
<style>
.main {
    background: linear-gradient(135deg, #0f172a, #1e293b);
    color: white;
}
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}
.card {
    background: #111827;
    padding: 2rem;
    border-radius: 20px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.5);
}
h1, h2, h3 {
    text-align: center;
}
.stButton>button {
    width: 100%;
    border-radius: 12px;
    height: 3em;
    font-size: 16px;
    font-weight: bold;
    background: linear-gradient(90deg, #06b6d4, #3b82f6);
    color: white;
    border: none;
}
.stDownloadButton>button {
    width: 100%;
    border-radius: 12px;
    height: 3em;
    font-size: 16px;
    font-weight: bold;
    background: linear-gradient(90deg, #22c55e, #16a34a);
    color: white;
}
.upload-box {
    border: 2px dashed #3b82f6;
    padding: 1.5rem;
    border-radius: 15px;
    text-align: center;
}
.footer {
    text-align: center;
    font-size: 14px;
    opacity: 0.7;
}
</style>
""", unsafe_allow_html=True)

# ---------------------- HEADER ----------------------
st.markdown('<div class="card">', unsafe_allow_html=True)

st.title("⚡ Electrochemistry Data Extractor")
st.caption("Upload `.xlsb` files → Extract CV / GCD → Download instantly")

# ---------------------- INPUT SECTION ----------------------
option = st.radio("Select Mode", ["CV", "GCD"], horizontal=True)

uploaded_files = st.file_uploader(
    "📂 Upload Files",
    type=["xlsb"],
    accept_multiple_files=True
)

colA, colB = st.columns(2)

process = colA.button("🚀 Process Files")
clear = colB.button("🗑 Remove All Files")

if clear:
    st.session_state.clear()
    st.rerun()

# ---------------------- CORE FUNCTIONS ----------------------
COL_CYCLENO, COL_STEPNO, COL_STEPTIME, COL_VOLTAGE = 4, 6, 7, 9

def _iter_records(raw: bytes):
    i = 0
    while i < len(raw):
        b0 = raw[i]; i += 1
        if b0 & 0x80:
            b1 = raw[i]; i += 1
            rec_type = (b0 & 0x7F) | (b1 << 7)
        else:
            rec_type = b0

        size, shift = 0, 0
        for _ in range(4):
            b = raw[i]; i += 1
            size |= (b & 0x7F) << shift
            shift += 7
            if not (b & 0x80): break

        rec_data = raw[i:i+size]
        i += size
        yield rec_type, rec_data

def read_sheet(file_obj):
    with zipfile.ZipFile(file_obj) as z:
        raw = z.read("xl/worksheets/sheet2.bin")

    cycle, step, step_time, voltage = {}, {}, {}, {}
    cur_row = -1

    for rec_type, rec_data in _iter_records(raw):
        if rec_type == 0 and len(rec_data) >= 4:
            cur_row = struct.unpack_from('<I', rec_data, 0)[0]
            continue
        if cur_row <= 0:
            continue

        if rec_type == 62:
            col = struct.unpack_from('<I', rec_data, 0)[0]
            if col == COL_STEPTIME:
                try:
                    val = rec_data[13:].decode(errors="ignore")
                    step_time[cur_row] = float(val.split(":")[-1])
                except:
                    pass

        elif rec_type == 5:
            col = struct.unpack_from('<I', rec_data, 0)[0]
            val = struct.unpack_from('<d', rec_data, 8)[0]
            if col == COL_CYCLENO: cycle[cur_row] = val
            elif col == COL_STEPNO: step[cur_row] = val
            elif col == COL_VOLTAGE: voltage[cur_row] = val

    return cycle, step, step_time, voltage

def process_gcd_file(file):
    cycle, step, step_time, voltage = read_sheet(file)
    rows = sorted(set(cycle) & set(step) & set(step_time) & set(voltage))

    data = []
    for r in rows:
        data.append([step_time[r], voltage[r]])

    return pd.DataFrame(data, columns=["StepTime_s", "Voltage_V"])

def process_cv_file(file):
    df = pd.read_excel(file, engine="pyxlsb", header=None)
    df = df.dropna()

    return df.iloc[:, [9, 8]].rename(columns={9: "Voltage_V", 8: "_Current_A"})

# ---------------------- PROCESSING ----------------------
if process:
    if not uploaded_files:
        st.warning("⚠️ Upload files first")
    else:
        progress = st.progress(0)
        results = []

        for i, file in enumerate(uploaded_files):
            try:
                df = process_cv_file(file) if option == "CV" else process_gcd_file(file)
                if df is not None:
                    results.append((file.name, df))
            except Exception as e:
                st.error(f"{file.name}: {e}")
            progress.progress((i+1)/len(uploaded_files))

        if results:
            output = io.BytesIO()

            if option == "CV":
                final = pd.concat([r[1] for r in results], axis=1)
                final.to_excel(output, index=False)
            else:
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for name, df in results:
                        df.to_excel(writer, sheet_name=name[:30], index=False)

            output.seek(0)

            st.success("✅ Done! File ready")

            st.download_button(
                "⬇ Download Excel",
                data=output,
                file_name=f"{option}_Data.xlsx"
            )

            # 🔥 AUTO DOWNLOAD
            st.markdown("""
                <script>
                const btn = window.parent.document.querySelector('button[kind="primary"]');
                if (btn) btn.click();
                </script>
            """, unsafe_allow_html=True)

# ---------------------- SUPPORT SECTION ----------------------
with st.expander("❤️ Support Research"):
    col1, col2 = st.columns(2)

    with col1:
        upi = "upi://pay?pa=satyajeet1396@oksbi&pn=Satyajeet"
        qr = qrcode.make(upi)

        buffer = BytesIO()
        qr.save(buffer, format="PNG")
        img = base64.b64encode(buffer.getvalue()).decode()

        st.markdown(f"""
        <div style="text-align:center">
        <img src="data:image/png;base64,{img}" width="180"><br>
        <b>satyajeet1396@oksbi</b>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="text-align:center">
        <a href="https://www.buymeacoffee.com/researcher13">
        ☕ Buy Me a Coffee
        </a>
        </div>
        """, unsafe_allow_html=True)

# ---------------------- FOOTER ----------------------
st.markdown("""
<div class="footer">
Created by Dr. Satyajeet Patil • Research Tools for Scientists 🚀
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
