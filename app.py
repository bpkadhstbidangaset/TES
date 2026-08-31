import streamlit as st
import pandas as pd
import os

# Konfigurasi Halaman
st.set_page_config(
    page_title="Sistem Rekon Persediaan & Belanja Modal",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header { font-size: 26px; font-weight: 700; color: #1E3A8A; margin-bottom: 2px; }
    .sub-header { font-size: 14px; color: #6B7280; margin-bottom: 20px; }
    .metric-box {
        color: white; padding: 16px 20px; border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 15px;
    }
    .metric-title { font-size: 13px; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; opacity: 0.9; }
    .metric-value { font-size: 22px; font-weight: 700; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)

def format_rupiah(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# 8 Kolom Pilihan Sesuai Permintaan
KOLOM_DETAIL_PILIHAN = [
    'Kode Sub Kegiatan',
    'Nama Sub Kegiatan',
    'Kode Rekening',
    'Nama Rekening',
    'Nomor Dokumen',
    'Tanggal Dokumen',
    'Keterangan Dokumen',
    'Nilai Realisasi'
]

# --- LOAD MASTER PATOKAN RAK ---
@st.cache_data
def load_master_rak():
    file_persediaan = "RAK PERSEDIAAN.xlsx"
    file_modal = "RAK BELANJA MODAL.xlsx"
    
    if not os.path.exists(file_persediaan) or not os.path.exists(file_modal):
        return None, None, f"File '{file_persediaan}' atau '{file_modal}' tidak ditemukan."

    # 1. Master Persediaan
    df_p_raw = pd.read_excel(file_persediaan, header=None).iloc[2:]
    rek_persediaan = set(df_p_raw[1].dropna().astype(str).str.strip()) - {'-', 'KODE REKENING'}

    # 2. Master Belanja Modal (HANYA KODE REKENING AWALAN 5.2)
    df_m_raw = pd.read_excel(file_modal, header=None).iloc[1:]
    df_m_raw.columns = ['Kode Kategori', 'Uraian Kategori', 'Uraian Rekening', 'Kode Rekening']
    df_m_clean = df_m_raw.dropna(subset=['Kode Rekening']).copy()
    df_m_clean['Kode Rekening'] = df_m_clean['Kode Rekening'].astype(str).str.strip()
    
    # FILTER: HANYA BELANJA MODAL (5.2...)
    df_m_clean = df_m_clean[df_m_clean['Kode Rekening'].str.startswith('5.2')]
    
    df_m_map = df_m_clean[['Kode Rekening', 'Kode Kategori', 'Uraian Kategori']].drop_duplicates('Kode Rekening')

    return rek_persediaan, df_m_map, None

rek_persediaan, df_modal_map, error_msg = load_master_rak()

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135679.png", width=70)
    st.markdown("### **Panel Kontrol**")
    st.info("📌 **Master RAK Modal:** Terpasang di sistem")
    
    f_lra = st.file_uploader("📥 Upload LRA Realisasi (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.caption("Pemerintah Kabupaten Hulu Sungai Tengah © 2026")

# --- HEADER UTAMA ---
st.markdown('<div class="main-header">🏛️ Rekonsiliasi Rekening Persediaan & Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Realisasi LRA terhadap Klasifikasi Persediaan dan Belanja Modal (Akun 5.2)</div>', unsafe_allow_html=True)

if error_msg:
    st.error(f"⚠️ {error_msg}")
    st.stop()

if f_lra:
    with st.spinner("Sedang memproses data realisasi..."):
        # Baca LRA
        df_lra = pd.read_excel(f_lra, sheet_name='Data Realisasi Dokumen', header=4)
        df_lra['Kode Rekening'] = df_lra['Kode Rekening'].astype(str).str.strip()
        df_lra['Nilai Realisasi'] = pd.to_numeric(df_lra['Nilai Realisasi'], errors='coerce').fillna(0)

        # 1. Filter Persediaan
        df_rekon_pers = df_lra[df_lra['Kode Rekening'].isin(rek_persediaan)].copy()

        # 2. Filter Belanja Modal (Hanya Akun 5.2 yang cocok dengan RAK Modal)
        df_rekon_modal = df_lra[
            (df_lra['Kode Rekening'].str.startswith('5.2')) & 
            (df_lra['Kode Rekening'].isin(df_modal_map['Kode Rekening'].unique()))
        ].copy()
        df_rekon_modal = df_rekon_modal.merge(df_modal_map, on='Kode Rekening', how='left')

    # --- FILTER SKPD ---
    daftar_skpd = ["-- SEMUA SKPD --"] + sorted(list(df_lra['Nama SKPD'].dropna().unique()))
    col_filter, _ = st.columns([2, 1])
    with col_filter:
        pilihan_skpd = st.selectbox("🎯 **Pilih Perangkat Daerah (SKPD):**", daftar_skpd)

    if pilihan_skpd != "-- SEMUA SKPD --":
        df_pers_filtered = df_rekon_pers[df_rekon_pers['Nama SKPD'] == pilihan_skpd].copy()
        df_modal_filtered = df_rekon_modal[df_rekon_modal['Nama SKPD'] == pilihan_skpd].copy()
    else:
        df_pers_filtered = df_rekon_pers.copy()
        df_modal_filtered = df_rekon_modal.copy()

    # --- RINGKASAN METRICS ---
    total_p = df_pers_filtered['Nilai Realisasi'].sum()
    total_m = df_modal_filtered['Nilai Realisasi'].sum()
    grand_total = total_p + total_m

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"""
        <div class="metric-box" style="background: linear-gradient(135deg, #0D9488, #14B8A6);">
            <div class="metric-title">📦 Total Belanja Persediaan</div>
            <div class="metric-value">{format_rupiah(total_p)}</div>
        </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-box" style="background: linear-gradient(135deg, #4F46E5, #6366F1);">
            <div class="metric-title">🏢 Total Belanja Modal (Akun 5.2)</div>
            <div class="metric-value">{format_rupiah(total_m)}</div>
        </div>
        """, unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="metric-box" style="background: linear-gradient(135deg, #1E293B, #334155);">
            <div class="metric-title">📊 Total Gabungan Realisasi</div>
            <div class="metric-value">{format_rupiah(grand_total)}</div>
        </div>
        """, unsafe_allow_html=True)

    # --- TABS KONTEN ---
    tab1, tab2 = st.tabs(["📦 REKON PERSEDIAAN", "🏢 REKON BELANJA MODAL"])

    # ================= TAB 1: PERSEDIAAN =================
    with tab1:
        st.subheader("1. Rekapitulasi per Kode Rekening Persediaan")
        if not df_pers_filtered.empty:
            rekap_p = df_pers_filtered.groupby(['Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
            baris_total_p = pd.DataFrame([{
                'Kode Rekening': 'TOTAL',
                'Nama Rekening': 'JUMLAH KESELURUHAN',
                'Nilai Realisasi': rekap_p['Nilai Realisasi'].sum()
            }])
            rekap_p_tampil = pd.concat([rekap_p, baris_total_p], ignore_index=True)
            rekap_p_tampil['Realisasi (Rp)'] = rekap_p_tampil['Nilai Realisasi'].apply(format_rupiah)
            
            st.dataframe(
                rekap_p_tampil[['Kode Rekening', 'Nama Rekening', 'Realisasi (Rp)']],
                use_container_width=True,
                hide_index=True
            )

            st.markdown("---")
            st.subheader("2. Rincian Dokumen Realisasi")
            df_detail_p = df_pers_filtered[[c for c in KOLOM_DETAIL_PILIHAN if c in df_pers_filtered.columns]].copy()
            df_detail_p['Nilai Realisasi'] = df_detail_p['Nilai Realisasi'].apply(format_rupiah)
            st.dataframe(df_detail_p, use_container_width=True, hide_index=True)
        else:
            st.warning("Tidak ditemukan data realisasi persediaan untuk SKPD ini.")

    # ================= TAB 2: BELANJA MODAL (AKUN 5.2) =================
    with tab2:
        st.subheader("1. Rekapitulasi per Kategori & Rekening Belanja Modal (5.2...)")
        if not df_modal_filtered.empty:
            rekap_m = df_modal_filtered.groupby(['Kode Kategori', 'Uraian Kategori', 'Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
            
            baris_total_m = pd.DataFrame([{
                'Kode Kategori': 'TOTAL',
                'Uraian Kategori': 'JUMLAH KESELURUHAN',
                'Kode Rekening': '-',
                'Nama Rekening': '-',
                'Nilai Realisasi': rekap_m['Nilai Realisasi'].sum()
            }])
            rekap_m_tampil = pd.concat([rekap_m, baris_total_m], ignore_index=True)
            rekap_m_tampil['Realisasi (Rp)'] = rekap_m_tampil['Nilai Realisasi'].apply(format_rupiah)

            st.dataframe(
                rekap_m_tampil[['Kode Kategori', 'Uraian Kategori', 'Kode Rekening', 'Nama Rekening', 'Realisasi (Rp)']],
                use_container_width=True,
                hide_index=True
            )

            st.markdown("---")
            st.subheader("2. Rincian Dokumen Realisasi")
            df_detail_m = df_modal_filtered[[c for c in KOLOM_DETAIL_PILIHAN if c in df_modal_filtered.columns]].copy()
            df_detail_m['Nilai Realisasi'] = df_detail_m['Nilai Realisasi'].apply(format_rupiah)
            st.dataframe(df_detail_m, use_container_width=True, hide_index=True)
        else:
            st.warning("Tidak ditemukan data realisasi belanja modal (Akun 5.2) untuk SKPD ini.")

else:
    st.info("👈 Silakan upload file **`LRA REALISASI JAN-JUNI30.xlsx`** pada menu di sebelah kiri.")
