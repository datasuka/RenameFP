# Rename PDF Faktur Pajak Berdasarkan Metadata (Dark Mode & Custom UI)

import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
import zipfile

st.set_page_config(page_title="Rename Faktur Pajak", layout="centered")

# CSS untuk dark theme DJP style
st.markdown("""
<style>
    .stApp { background-color: #0d1117; color: white; }
    h1, h2, h3, h4, h5, h6, p, label, .markdown-text-container, .stText, .stMarkdown {
        color: white !important;
    }
    .stButton>button {
        background-color: #1f6feb;
        color: white !important;
        border-radius: 6px;
        padding: 0.4em 0.9em;
    }
    .stDownloadButton>button {
        background-color: #2ea44f;
        color: white !important;
        border-radius: 6px;
        padding: 0.4em 0.9em;
    }
    .stFileUploader {
        background-color: #161b22;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("## 🧾 Rename PDF Faktur Pajak Berdasarkan Format yang ditentukan.")
st.markdown("*By: Reza Fahlevi Lubis BKP @zavibis*")

st.markdown("### 📌 Petunjuk Penggunaan")
st.markdown("""
1. **Upload** satu atau beberapa file PDF Faktur Pajak.
2. Aplikasi akan otomatis membaca isi metadata dari setiap faktur.
3. Pilih dan urutkan **kolom-kolom metadata** untuk dijadikan format nama file PDF.
4. Klik **Rename PDF & Download** untuk mengunduh file hasil rename dalam 1 file ZIP.
""")

bulan_map = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

def extract(pattern, text, flags=re.DOTALL, default="-", postproc=lambda x: x.strip()):
    match = re.search(pattern, text, flags)
    return postproc(match.group(1)) if match else default

def extract_tanggal(text):
    match = re.search(r",\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    return f"{match.group(1).zfill(2)}/{match.group(2)}/{match.group(3)}" if match else "-"

def extract_nitku_pembeli(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "NPWP" in line and i > 0:
            prev_line = lines[i-1]
            match = re.search(r"#(\d{22})", prev_line)
            if match:
                return match.group(1)
    return "-"

def extract_data_from_text(text):
    raw_data = {
        "NomorSeriFaktur": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", text),
        "NamaPKP": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", text),
        "AlamatPKP": extract(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.*?)\s*NPWP", text),
        "NPWPPKP": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9\.]+)", text),
        "NamaPembeli": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", text),
        "NPWPPembeli": extract(r"NPWP\s*:\s*([0-9\.]+)\s*NIK", text),
        "Referensi": extract(r"Referensi:\s*(.*?)\n", text, postproc=lambda x: x.strip().rstrip(")")),
        "TanggalFaktur": extract_tanggal(text),
        "NITKU": extract_nitku_pembeli(text),
    }

    try:
        d, m_str, y = raw_data["TanggalFaktur"].split("/")
        month_num = bulan_map.get(m_str, "00")
        raw_data["TanggalFaktur"] = f"{d} {month_num} {y}"
        raw_data["Masa"] = month_num
        raw_data["Tahun"] = y
    except:
        raw_data["Masa"] = "-"
        raw_data["Tahun"] = "-"

    return raw_data

def sanitize_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "_", str(text))

def generate_filename(row, selected_cols):
    parts = [sanitize_filename(str(row[col])) for col in selected_cols]
    return "Faktur Pajak " + "_".join(parts) + ".pdf"

uploaded_files = st.file_uploader("📎 Upload PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data_rows = []
    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.read()
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            text = "".join(page.get_text() for page in doc)
        raw_data = extract_data_from_text(text)
        raw_data["OriginalName"] = uploaded_file.name
        raw_data["FileBytes"] = file_bytes
        data_rows.append(raw_data)

    df = pd.DataFrame(data_rows).drop(columns=["FileBytes", "OriginalName"])
    column_options = df.columns.tolist()

    st.markdown("### ✏️ Pilih Kolom untuk Format Nama File")
    selected_columns = st.multiselect("Urutan Nama File", column_options, default=[], key="selector", help="Klik lalu geser untuk atur urutan")

    if st.button("🔁 Rename PDF & Download"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for i, row in df.iterrows():
                filename = generate_filename(row, selected_columns)
                zipf.writestr(filename, data_rows[i]["FileBytes"])
        zip_buffer.seek(0)
        st.success("✅ Berhasil! Klik tombol di bawah ini untuk mengunduh file ZIP.")
        st.download_button("📦 Download ZIP PDF Hasil Rename", zip_buffer, file_name="faktur_renamed.zip", mime="application/zip")
