import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import re
import io

# Konfigurasi Tesseract (ganti path jika beda)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Fungsi OCR
def proses_gambar(image_file):
    image = Image.open(image_file)
    ocr_text = pytesseract.image_to_string(image, lang='eng+ind')

    lines = ocr_text.strip().split("\n")
    lines = [line.strip() for line in lines if line.strip()]

    grouped_data = []
    temp_group = []

    for line in lines:
        temp_group.append(line)
        if "Rp" in line:
            grouped_data.append(temp_group)
            temp_group = []

    final_data = []
    for group in grouped_data:
        name = group[0]
        price_match = re.search(r"Rp\s*([\d.]+)", group[-1])
        price = price_match.group(1) if price_match else "N/A"
        final_data.append([name, price])

    return final_data


# =================== UI Streamlit ===================
st.set_page_config(page_title="OCR Gambar ke Excel", page_icon="📄", layout="centered")

st.markdown("<h1 style='text-align: center;'>📸 OCR Gambar ke Excel</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>Ekstrak nama & harga dari gambar dan simpan langsung ke Excel</p>", unsafe_allow_html=True)

st.markdown("---")

# Layout kolom input
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_files = st.file_uploader("📂 Unggah gambar (bisa lebih dari satu)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
with col2:
    file_title = st.text_input("📝 Nama file Excel", value="output")

if uploaded_files:
    all_data = []

    with st.spinner("⏳ Memproses gambar..."):
        for uploaded_file in uploaded_files:
            result = proses_gambar(uploaded_file)
            all_data.extend(result)

    df = pd.DataFrame(all_data, columns=["Nama Paket", "Harga"])
    
    st.success("✅ Semua gambar berhasil diproses!")
    st.markdown("### 🧾 Hasil Ekstraksi")
    st.dataframe(df, use_container_width=True)

    # Simpan ke Excel
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    file_name = file_title.strip() + ".xlsx"

    st.markdown("---")
    st.download_button(
        label="⬇️ Download Excel",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Klik untuk mengunduh hasil dalam format Excel (.xlsx)"
    )
else:
    st.info("📌 Silakan unggah gambar untuk mulai memproses data.")
