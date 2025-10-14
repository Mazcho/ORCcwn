import streamlit as st
import pandas as pd
import io
import math

# ===== Streamlit Config =====
st.set_page_config(page_title="WA Blast", page_icon="📲", layout="wide")

# ===== Header =====
st.markdown("<h1 style='text-align:center;'>📲 WhatsApp Blast IndiHome</h1>", unsafe_allow_html=True)
st.markdown("---")

# ===== Upload File =====
st.subheader("📁 Upload File Pelanggan")
uploaded_file = st.file_uploader(
    "Upload file CSV atau Excel berisi kolom: No_hp, no_internet, nominal",
    type=["csv", "xlsx"]
)

if uploaded_file:
    # Baca file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)

    # Bersihkan nomor HP
    df["No_hp"] = df["No_hp"].astype(str).str.replace(",", "").str.strip()

    st.markdown("### 📋 Preview Data")
    st.dataframe(df.astype(str), use_container_width=True)

    # ===== Input Delay =====
    delay = st.number_input(
        "⏳ Delay antar pesan (detik, hanya referensi untuk manual sending)", 
        min_value=1, max_value=30, value=5
    )

    # ===== Template Pesan =====
    st.subheader("✏️ Template Pesan")
    pesan_template = st.text_area(
        "Gunakan {no_internet} dan {nominal} untuk personalisasi pesan",
        value="Halo, jangan lupa membayar tagihan IndiHome Anda dengan nomor internet {no_internet}, nominal {nominal}."
    )

    # ===== Generate Preview Pesan =====
    if st.button("🔎 Generate Preview Pesan"):
        df["Pesan"] = df.apply(
            lambda x: pesan_template.format(
                no_internet=x.get("no_internet", ""), 
                nominal=x.get("nominal", "")
            ), axis=1
        )

        st.markdown("### 📨 Preview Pesan Personal")
        st.dataframe(df[["No_hp", "Pesan"]], use_container_width=True)

        # ===== Estimasi Waktu =====
        total_rows = len(df)
        total_seconds = total_rows * delay
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        st.markdown(f"⏱️ **Estimasi waktu selesai:** ~ {minutes} menit {seconds} detik untuk {total_rows} pesan")

        # ===== Download Excel =====
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, sheet_name="WA_Blast", engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            label="📥 Download File Pesan Siap Kirim (Excel)",
            data=buffer,
            file_name="WA_Blast.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ===== Footer =====
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:grey;'>Created by MazCho - CWN Python Developer</p>", 
    unsafe_allow_html=True
)
