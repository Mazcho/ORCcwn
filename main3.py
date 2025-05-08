import streamlit as st
import pytesseract
from PIL import Image, ImageOps
import pandas as pd
import re
import io
import zipfile
from pathlib import Path

# ------------------- Konstanta untuk deteksi FEE -------------------
BLOCK_HEIGHT = 180  # tinggi area 1 paket (perkiraan)

def extract_fee_from_crop(image, block_index):
    width, height = image.size
    top = block_index * BLOCK_HEIGHT
    bottom = top + 50
    left = int(width * 0.75)
    right = width

    cropped = image.crop((left, top, right, bottom))
    gray = cropped.convert('L')
    inverted = ImageOps.invert(gray)
    bw = inverted.point(lambda x: 0 if x < 128 else 255, '1')
    
    ocr_text = pytesseract.image_to_string(bw, lang='eng+ind', config='--psm 6')
    match = re.search(r"([\d]{2,3}[.,]?\d{0,3})", ocr_text)
    return match.group(1).replace(",", ".") if match else "N/A"

def proses_gambar(image_file):
    image = Image.open(image_file)
    ocr_text = pytesseract.image_to_string(image, lang='eng+ind', config='--psm 6')

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
    for i, group in enumerate(grouped_data):
        name = group[0]
        price_match = re.search(r"Rp\s*([\d.]+)", " ".join(group))
        price = f"Rp {price_match.group(1)}" if price_match else "Rp N/A"
        fee = extract_fee_from_crop(image, i)
        final_data.append([name, price, fee])

    return final_data

# ------------------- STREAMLIT APP -------------------
st.set_page_config(page_title="OCR Gambar ke Excel", page_icon="📄", layout="wide")

# ===== Sidebar Navigation =====
st.sidebar.title("🔧 Menu Tools")
tool_option = st.sidebar.radio("Pilih Tools:", ["Tools Dataps Nasional", "Tools Preprocessing DataJATENG"])

# ===== Tools 1: OCR Tools =====
if tool_option == "Tools Dataps Nasional":
    st.title("📄 OCR Gambar (Nama Paket, Harga, FEE) ke Excel")

    tab1, tab2, tab3 = st.tabs(["🖼️ Gambar Manual", "🗂️ ZIP Folder Daerah", "❓ Help"])

    with tab1:
        st.subheader("🖼️ Upload Gambar Tunggal atau Multiple")

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

            df = pd.DataFrame(all_data, columns=["Nama Paket", "Harga", "FEE"])

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
                                for nama, harga, fee in hasil_ocr:
                                    all_data.append({
                                        "Daerah": daerah,
                                        "Nama Paket": nama,
                                        "Harga": harga,
                                        "FEE": fee
                                    })
                            except Exception as e:
                                st.warning(f"Gagal proses gambar: {image_path} ({e})")

            if all_data:
                df = pd.DataFrame(all_data)
                st.success("✅ Data dari ZIP berhasil diproses! Anda bisa mengedit sebelum download.")
                edited_df = st.data_editor(df, use_container_width=True)

                st.markdown("### 📦 Hasil Akhir (Siap Diunduh)")
                st.dataframe(edited_df, use_container_width=True)

                output = io.BytesIO()
                edited_df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)

                st.download_button(
                    label="⬇️ Download Excel",
                    data=output,
                    file_name=f"{file_title.strip()}_zip.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("📌 Tidak ada gambar yang bisa diproses dalam ZIP.")

            # --- TAB 3: HELP ---
    with tab3:
        st.subheader("📖 Panduan Penggunaan Aplikasi")

        st.markdown("### 📌 Mode Manual (Gambar Tunggal / Multiple)")
        st.markdown("""
        1. Klik tab **🖼️ Gambar Manual**.
        2. Upload 1 atau lebih file gambar (format: JPG, PNG).
        3. Masukkan nama file Excel yang akan disimpan.
        4. Hasil akan muncul secara otomatis setelah diproses.
        5. Anda dapat mengedit hasil tabel jika diperlukan.
        6. Klik **⬇️ Download** untuk mengunduh file Excel.

        **Catatan:**
        - Pastikan gambar jelas dan memiliki teks nama paket, harga (Rp), dan FEE di sebelah kanan bawah.
        - Sistem akan otomatis memotong area untuk membaca FEE.
        """)

        st.markdown("### 📦 Mode ZIP Folder (Gambar Kelompok Daerah)")
        st.markdown("""
        1. Klik tab **🗂️ ZIP Folder Daerah**.
        2. Upload file **ZIP** berisi gambar-gambar dari beberapa daerah.
        3. Setiap gambar dalam folder akan diproses otomatis.
        4. Nama folder akan dianggap sebagai **nama daerah**.
        5. Hasil akan digabung dalam satu tabel dengan kolom "Daerah".
        6. Anda dapat mengedit tabel sebelum mengunduh.
        7. Klik **⬇️ Download** untuk menyimpan sebagai Excel.

        **Catatan:**
        - Struktur folder ZIP harus seperti berikut:
        
            ```
            DataPsNasional
            └── Daerah_A/
                ├── gambar1.jpg
                ├── gambar2.png
            └── Daerah_B/
                ├── gambar3.jpg
            ```
        - Sistem akan mengabaikan file selain gambar.
        - Jika folder tidak ada (gambar langsung di root), kolom daerah akan diisi “Tidak Diketahui”.
        """)

        st.info("Jika ada kesalahan OCR, Anda bisa edit hasilnya sebelum diunduh.")

# ===== Tools 2: Placeholder Preprocessing JATENG =====
elif tool_option == "Tools Preprocessing DataJATENG":
    st.title("🧪 Tools Preprocessing Data JATENG")
    st.info("📌 Fitur preprocessing DataJATENG akan segera hadir. Silakan tunggu pembaruan selanjutnya.")

# Footer
st.markdown(
    """
    <hr style="margin-top: 50px;">
    <div style="text-align: center; font-size: 14px; color: grey;">
        Created by <strong>MazCho</strong> - <em>CWN Python Developer</em>
    </div>
    """,
    unsafe_allow_html=True
)
