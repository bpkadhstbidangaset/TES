import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="Aplikasi Rekon Persediaan & Modal", layout="wide")
st.title("📊 Aplikasi Rekonsiliasi Realisasi Belanja")
st.caption("Pemerintah Kabupaten Hulu Sungai Tengah")

# --- FUNGSI LOAD MASTER PATOKAN (CACHED AGAR CEPAT) ---
@st.cache_data
def load_master_rak():
    file_persediaan = "RAK PERSEDIAAN.xlsx"
    file_modal = "RAK BELANJA MODAL.xlsx"
    
    # Cek apakah file master ada di direktori
    if not os.path.exists(file_persediaan) or not os.path.exists(file_modal):
        return None, None, f"File patokan '{file_persediaan}' atau '{file_modal}' tidak ditemukan di folder aplikasi."

    # 1. Baca Master Persediaan
    df_p_raw = pd.read_excel(file_persediaan, header=None).iloc[2:]
    rek_persediaan = set(df_p_raw[1].dropna().astype(str).str.strip()) - {'-', 'KODE REKENING'}

    # 2. Baca Master Belanja Modal
    df_m_raw = pd.read_excel(file_modal, header=None).iloc[1:]
    df_m_raw.columns = ['Kode Kategori', 'Uraian Kategori', 'Uraian Rekening', 'Kode Rekening']
    df_m_clean = df_m_raw.dropna(subset=['Kode Rekening']).copy()
    df_m_clean['Kode Rekening'] = df_m_clean['Kode Rekening'].astype(str).str.strip()
    
    # Ambil mapping unik kode rekening ke kategori
    df_m_map = df_m_clean[['Kode Rekening', 'Kode Kategori', 'Uraian Kategori']].drop_duplicates('Kode Rekening')

    return rek_persediaan, df_m_map, None

# Load master RAK otomatis
rek_persediaan, df_modal_map, error_msg = load_master_rak()

if error_msg:
    st.error(error_msg)
    st.info("Pastikan file 'RAK PERSEDIAAN.xlsx' dan 'RAK BELANJA MODAL.xlsx' berada di folder yang sama dengan script ini.")
    st.stop()

# --- 1. UPLOAD HANYA FILE LRA ---
with st.sidebar:
    st.header("📂 Upload LRA")
    f_lra = st.file_uploader("Upload LRA Realisasi (.xlsx)", type=["xlsx"])
    st.success("✅ Master RAK Persediaan & Modal aktif (Otomatis)")

if f_lra:
    with st.spinner("Memproses data realisasi..."):
        # Baca LRA
        df_lra = pd.read_excel(f_lra, sheet_name='Data Realisasi Dokumen', header=4)
        df_lra['Kode Rekening'] = df_lra['Kode Rekening'].astype(str).str.strip()

        # Filter Persediaan
        df_rekon_pers = df_lra[df_lra['Kode Rekening'].isin(rek_persediaan)].copy()

        # Filter Modal
        df_rekon_modal = df_lra[df_lra['Kode Rekening'].isin(df_modal_map['Kode Rekening'].unique())].copy()
        df_rekon_modal = df_rekon_modal.merge(df_modal_map, on='Kode Rekening', how='left')

    # --- 2. PILIH SKPD ---
    daftar_skpd = ["Semua SKPD"] + sorted(list(df_lra['Nama SKPD'].dropna().unique()))
    pilihan_skpd = st.selectbox("📌 Pilih SKPD:", daftar_skpd)

    if pilihan_skpd != "Semua SKPD":
        df_rekon_pers = df_rekon_pers[df_rekon_pers['Nama SKPD'] == pilihan_skpd]
        df_rekon_modal = df_rekon_modal[df_rekon_modal['Nama SKPD'] == pilihan_skpd]

    # --- 3. TAMPILKAN HASIL ---
    tab1, tab2 = st.tabs(["📦 Rekon Rekening Persediaan", "🏢 Rekon Belanja Modal"])

    # TAB PERSEDIAAN
    with tab1:
        total_pers = df_rekon_pers['Nilai Realisasi'].sum()
        st.metric("Total Realisasi Persediaan", f"Rp {total_pers:,.2f}")
        
        st.subheader("Rekap per Rekening")
        rekap_pers = df_rekon_pers.groupby(['Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
        st.dataframe(rekap_pers, use_container_width=True)

        st.subheader("Detail Transaksi Dokumen")
        kolom_p = ['Tanggal SP2D', 'Nomor SP2D', 'Kode Rekening', 'Nama Rekening', 'Keterangan Dokumen', 'Nilai Realisasi']
        st.dataframe(df_rekon_pers[[c for c in kolom_p if c in df_rekon_pers.columns]], use_container_width=True)

    # TAB MODAL
    with tab2:
        total_modal = df_rekon_modal['Nilai Realisasi'].sum()
        st.metric("Total Realisasi Belanja Modal", f"Rp {total_modal:,.2f}")

        st.subheader("Rekap per Kategori Aset & Rekening")
        rekap_modal = df_rekon_modal.groupby(['Kode Kategori', 'Uraian Kategori', 'Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
        st.dataframe(rekap_modal, use_container_width=True)

        st.subheader("Detail Transaksi Dokumen")
        kolom_m = ['Tanggal SP2D', 'Nomor SP2D', 'Uraian Kategori', 'Kode Rekening', 'Nama Rekening', 'Keterangan Dokumen', 'Nilai Realisasi']
        st.dataframe(df_rekon_modal[[c for c in kolom_m if c in df_rekon_modal.columns]], use_container_width=True)

else:
    st.info("👋 Silakan upload file **LRA Realisasi (.xlsx)** pada menu di sebelah kiri.")