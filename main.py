import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import re
import io
import zipfile
from pathlib import Path

# Hilangkan path lokal, asumsikan tesseract terinstall otomatis di server
# Jika kamu di lokal Windows, kamu bisa aktifkan baris ini
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Fungsi OCR untuk gambar
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
        final_data.append((name, price))

    return final_data

# Konfigurasi halaman
st.set_page_config(page_title="🧾 OCR Gambar & ZIP ➜ Excel", layout="wide")
st.markdown("<h1 style='text-align: center;'>📤 Konversi Gambar Nota ke Excel</h1>", unsafe_allow_html=True)

# Input nama file output
file_title = st.text_input("📄 Nama File Output Excel", value="hasil_ocr")

# Tabs: Gambar Satuan & ZIP
tab1, tab2 = st.tabs(["🖼️ Upload Gambar", "🗂️ Upload ZIP Folder Daerah"])

with tab1:
    st.subheader("🖼️ Upload Gambar Satuan")
    uploaded_images = st.file_uploader("Pilih satu atau beberapa gambar", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

    if uploaded_images:
        all_data = []
        for img in uploaded_images:
            try:
                hasil = proses_gambar(img)
                for nama, harga in hasil:
                    all_data.append({
                        "Nama Paket": nama,
                        "Harga": harga,
                        "Free PSB": "None"
                    })
            except Exception as e:
                st.warning(f"Gagal proses gambar {img.name} ({e})")

        if all_data:
            df = pd.DataFrame(all_data)
            st.success("✅ Data berhasil diproses! Anda bisa mengedit tabel sebelum download.")
            edited_df = st.data_editor(df, use_container_width=True)

            # Download
            output = io.BytesIO()
            edited_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="⬇️ Download Excel",
                data=output,
                file_name=f"{file_title.strip()}_gambar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tab2:
    st.subheader("🗂️ Upload ZIP Folder Daerah")
    uploaded_zip = st.file_uploader("Upload file ZIP berisi folder daerah dan gambar", type=["zip"])

    if uploaded_zip:
        all_data = []

        with zipfile.ZipFile(uploaded_zip) as zip_file:
            file_list = zip_file.namelist()
            image_files = [f for f in file_list if f.lower().endswith((".png", ".jpg", ".jpeg"))]

            with st.spinner("🔍 Memproses semua gambar..."):
                for image_path in image_files:
                    path_obj = Path(image_path)
                    daerah = path_obj.parent.name if path_obj.parent.name else "Tidak Diketahui"

                    with zip_file.open(image_path) as img_file:
                        try:
                            hasil_ocr = proses_gambar(img_file)
                            for nama, harga in hasil_ocr:
                                all_data.append({
                                    "Daerah": daerah,
                                    "Nama Paket": nama,
                                    "Harga": harga,
                                    "Free PSB": "None"
                                })
                        except Exception as e:
                            st.warning(f"Gagal proses gambar: {image_path} ({e})")

        if all_data:
            df = pd.DataFrame(all_data)
            st.success("✅ Data dari ZIP berhasil diproses! Anda bisa mengedit sebelum download.")
            edited_df = st.data_editor(df, use_container_width=True)

            # Download
            output = io.BytesIO()
            edited_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="⬇️ Download Excel",
                data=output,
                file_name=f"{file_title.strip()}_zip.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
