
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import os

st.title("Rename Faktur Pajak PDF Berdasarkan Metadata")

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
    return {
        "KodeFaktur": extract(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", text),
        "NamaPKP": extract(r"Pengusaha Kena Pajak:\s*Nama\s*:\s*(.*?)\s*Alamat", text),
        "AlamatPKP": extract(r"Pengusaha Kena Pajak:.*?Alamat\s*:\s*(.*?)\s*NPWP", text),
        "NPWPPKP": extract(r"Pengusaha Kena Pajak:.*?NPWP\s*:\s*([0-9\.]+)", text),
        "NamaPembeli": extract(r"Pembeli Barang Kena Pajak.*?Nama\s*:\s*(.*?)\s*Alamat", text),
        "AlamatPembeli": extract(r"Pembeli Barang Kena Pajak.*?Alamat\s*:\s*(.*?)\s*#", text),
        "NPWPPembeli": extract(r"NPWP\s*:\s*([0-9\.]+)\s*NIK", text),
        "Referensi": extract(r"Referensi:\s*(.*?)\n", text),
        "TanggalFaktur": extract_tanggal(text),
        "NITKU": extract_nitku_pembeli(text),
    }

def sanitize_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "_", str(text))

def generate_filename(row, selected_cols):
    parts = [sanitize_filename(str(row[col])) for col in selected_cols]
    return "_".join(parts) + ".pdf"

uploaded_files = st.file_uploader("Upload PDF Faktur Pajak", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    st.markdown("### 1. Pilih Kolom untuk Format Nama File")
    data_rows = []
    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.read()
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            text = "".join(page.get_text() for page in doc)
        data = extract_data_from_text(text)
        data["FileBytes"] = file_bytes
        data["OriginalName"] = uploaded_file.name
        data_rows.append(data)

    df = pd.DataFrame(data_rows).drop(columns=["FileBytes", "OriginalName"])
    column_options = df.columns.tolist()
    selected_columns = st.multiselect("Urutkan dan pilih kolom", column_options, default=column_options)

    if st.button("Rename dan Download Zip"):
        from io import BytesIO
        import zipfile

        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for i, row in df.iterrows():
                filename = generate_filename(row, selected_columns)
                zipf.writestr(filename, data_rows[i]["FileBytes"])

        buffer.seek(0)
        st.download_button("Download PDF Rename ZIP", buffer, file_name="faktur_rename.zip", mime="application/zip")
