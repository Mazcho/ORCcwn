import streamlit as st
import pytesseract
from PIL import Image, ImageOps
import pandas as pd
import re
import io
import zipfile
from pathlib import Path
import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

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
    st.title("📊 Aplikasi Preprocessing Data Internet")
    # Fungsi pembaca file universal
    def load_file(uploaded_file):
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.txt'):
            return pd.read_csv(uploaded_file, delimiter='\t')
        else:
            st.warning(f"Format file {uploaded_file.name} tidak dikenali.")
            return None

    # Inisialisasi session_state
    if 'final_df' not in st.session_state:
        st.session_state.final_df = None

    # Tab navigasi
    tab1, tab2, tab3 = st.tabs(["🔧 Preprocessing", "📂 Split & Download","HELP"])

    with tab1:
        st.subheader("📁 Upload File")
        data_file = st.file_uploader("Upload file data mentah", type=["csv", "xlsx", "xls", "txt"], key="data")
        sf_file = st.file_uploader("Upload file full_sales_force", type=["csv", "xlsx", "xls", "txt"], key="sf")
        tl_file = st.file_uploader("Upload file full_team_leader", type=["csv", "xlsx", "xls", "txt"], key="tl")

        if data_file and sf_file and tl_file:
            df = load_file(data_file)
            sf_df = load_file(sf_file)
            tl_df = load_file(tl_file)

            if df is not None and sf_df is not None and tl_df is not None:
                # Preprocessing nama: hilangkan spasi depan-belakang dan kapitalisasi
                sf_df['Nama SF'] = sf_df['user_name'].str.strip().str.title().str.replace("'", "", regex=False)
                tl_df['Nama TL'] = tl_df['user_name'].str.strip().str.title().str.replace("'", "", regex=False)
                df['Nama SF'] = df['Nama SF'].astype(str).str.strip().str.title().str.replace("'", "", regex=False)
                df['Nama TL'] = df['Nama TL'].astype(str).str.strip().str.title().str.replace("'", "", regex=False)

                # Ambil sf_id berdasarkan kodesf
                sf_df = sf_df.rename(columns={'sf_id': 'sf_id_master'})
                df = df.merge(sf_df[['sf_id_master']], left_on='Kode SF', right_on='sf_id_master', how='left')
                df = df.rename(columns={'sf_id_master': 'sf_id'})

                # Ambil tl_id berdasarkan Nama TL
                df = df.merge(tl_df[['Nama TL', 'tl_id']], on='Nama TL', how='left')

                # Ubah tanggal
                df['Tanggal PS'] = pd.to_datetime(df['Tanggal PS'], errors='coerce')
                df['waktu_ps'] = df['Tanggal PS'].dt.strftime('%Y-%m-%d %H:%M:%S')
                df['tanggal_angka'] = df['Tanggal PS'].dt.strftime('%d')

                # Isi NaN ID
                df['sf_id'] = df['sf_id'].fillna(0)
                df['tl_id'] = df['tl_id'].fillna(0)

                # Final dataframe
                final_df = df[['sf_id', 'tl_id', 'Order ID', 'Nomor Internet', 'waktu_ps', 'tanggal_angka', 'Paket', 'ARPU']]
                final_df.columns = ['sf_id', 'tl_id', 'track_id', 'no_internet', 'waktu_ps', 'tanggal_angka', 'paket_name', 'arpu']

                # Simpan ke session
                st.session_state.final_df = final_df

                # Tampilkan
                st.subheader("✅ Data Hasil Preprocessing")
                st.dataframe(final_df)

                # Download CSV
                csv = final_df.to_csv(index=False).encode('utf-8')
                st.download_button("⬇️ Download CSV", csv, "data_preprocessed.csv", "text/csv")

    with tab2:
        st.subheader("📂 Split dan Download Excel per 450 Baris")

        if st.session_state.final_df is not None:
            final_df = st.session_state.final_df

            if st.button("🔀 Split & Zip Excel Files"):
                buffer = BytesIO()
                with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    total_rows = final_df.shape[0]
                    num_chunks = (total_rows // 450) + int(total_rows % 450 != 0)

                    for i in range(num_chunks):
                        start = i * 450
                        end = start + 450
                        chunk = final_df.iloc[start:end]

                        excel_buffer = BytesIO()
                        chunk.to_excel(excel_buffer, index=False, sheet_name=f'Data_{i+1}')
                        excel_filename = f'output_part_{i+1}.xlsx'
                        zipf.writestr(excel_filename, excel_buffer.getvalue())

                st.success(f"Berhasil membuat {num_chunks} file dan mengompres ke ZIP.")

                # Tombol download ZIP
                st.download_button(
                    label="⬇️ Download ZIP",
                    data=buffer.getvalue(),
                    file_name="split_excel_output.zip",
                    mime="application/zip"
                )
        else:
            st.info("⚠️ Silakan lakukan preprocessing terlebih dahulu di Tab 1.")
    
    with tab3:
        st.subheader("🆘 Panduan Penggunaan Aplikasi HELP")
        st.markdown("""
        ### 📌 Langkah-langkah Penggunaan:

        1. **Upload File** (di Tab 🔧 Preprocessing)
            - File **data mentah**: berisi data transaksi atau pemasangan.
            - File **full_sales_force**: wajib ada kolom `user_name` dan `sf_id`.
            - File **full_team_leader**: wajib ada kolom `user_name` dan `tl_id`.

        2. **Proses Preprocessing**
            - Nama SF dan TL akan dibersihkan dan disesuaikan.
            - Dicocokkan dengan ID dari master SF & TL.
            - Tanggal akan diformat otomatis.

        3. **Split dan Download**
            - Buka tab 📂 *Split & Download*.
            - Klik tombol **🔀 Split & Zip Excel Files**.
            - Data akan dibagi tiap 450 baris dan dijadikan file Excel.
            - File dikompresi ke format ZIP dan bisa langsung diunduh.

        ---

        ### 📄 Unduh Panduan Lengkap (PDF)

        👉 [Klik di sini untuk download data Team Leader (wajib login akun CWN)](https://creativewidyanusantara.co.id/admincwn/manage_tl/)
                    

        👉 [Klik di sini untuk download data Sales FOrce (wajib login akun CWN)](https://creativewidyanusantara.co.id/admincwn/admincwn/updatesf/)

        ---

        ### ❗ Tips & Catatan:
        - Gunakan format file yang didukung: `.csv`, `.xlsx`, `.xls`, `.txt`
        - Pastikan kolom pada data sesuai dengan ketentuan aplikasi.
        - Proses split bisa dilakukan *setelah* preprocessing berhasil.

        ---

        📬 **Kontak Bantuan Teknis:**
        Jika mengalami kendala, silakan hubungi tim support melalui:
        - 📧 Email: support@creativewidyanusantara.co.id
        - 🌐 Website: [creativewidyanusantara.co.id](https://creativewidyanusantara.co.id)
        """)



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
