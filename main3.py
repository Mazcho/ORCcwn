import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import re
import io

# Hilangkan path lokal, asumsikan tesseract terinstall otomatis di server (paket: tesseract-ocr)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

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
        price = f"Rp {price_match.group(1)}" if price_match else "Rp N/A"
        final_data.append([name, price])

    return final_data

# =================== STREAMLIT UI ===================
st.set_page_config(page_title="OCR Gambar Editable ke Excel", page_icon="📝", layout="centered")

st.title("📄 OCR Gambar ke Excel (Editable)")
st.caption("Unggah gambar yang berisi data nama & harga, edit langsung hasilnya, dan simpan sebagai Excel.")

st.markdown("---")

col1, col2 = st.columns([2, 1])
with col1:
    uploaded_files = st.file_uploader("📂 Upload Gambar (JPG, PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
with col2:
    file_title = st.text_input("📝 Nama File Excel", value="output")

if uploaded_files:
    all_data = []

    with st.spinner("🔍 Memproses OCR dari gambar..."):
        for uploaded_file in uploaded_files:
            result = proses_gambar(uploaded_file)
            all_data.extend(result)

    df = pd.DataFrame(all_data, columns=["Nama Paket", "Harga"])
    
    st.success("✅ OCR selesai! Anda bisa mengedit hasilnya di bawah ini.")

    edited_df = st.data_editor(df, use_container_width=True, num_rows="fixed")

    st.markdown("### 📦 Hasil Akhir (Siap Diunduh)")
    st.dataframe(edited_df, use_container_width=True)

    buffer = io.BytesIO()
    edited_df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    filename = file_title.strip() + ".xlsx"

    st.download_button(
        label="⬇️ Download Excel dengan Perubahan",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📌 Unggah gambar terlebih dahulu untuk memulai proses OCR.")


