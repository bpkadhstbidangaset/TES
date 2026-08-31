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
    .status-card-match {
        background-color: #ECFDF5; border-left: 5px solid #10B981;
        padding: 14px 18px; border-radius: 8px; margin-top: 10px; margin-bottom: 20px;
    }
    .status-card-diff {
        background-color: #FEF2F2; border-left: 5px solid #EF4444;
        padding: 14px 18px; border-radius: 8px; margin-top: 10px; margin-bottom: 20px;
    }
    table.custom-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        margin-bottom: 20px;
        font-size: 14px;
    }
    table.custom-table th {
        background-color: #F1F5F9;
        color: #1E293B;
        font-weight: 700;
        text-align: left;
        padding: 10px 12px;
        border: 1px solid #CBD5E1;
    }
    table.custom-table td {
        padding: 9px 12px;
        border: 1px solid #E2E8F0;
        color: #334155;
    }
    table.custom-table tr:nth-child(even) {
        background-color: #F8FAFC;
    }
    table.custom-table tr.total-row td {
        background-color: #E2E8F0 !important;
        font-weight: 800 !important;
        color: #0F172A !important;
        border-top: 2px solid #94A3B8 !important;
        border-bottom: 2px solid #94A3B8 !important;
    }
    .text-right { text-align: right; }
    .text-center { text-align: center; }
</style>
""", unsafe_allow_html=True)

def format_rupiah(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# 8 Kolom Pilihan
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

    # 2. Master Belanja Modal (Akun 5.2)
    df_m_raw = pd.read_excel(file_modal, header=None).iloc[1:]
    df_m_raw.columns = ['Kode Kategori', 'Uraian Kategori', 'Uraian Rekening', 'Kode Rekening']
    df_m_clean = df_m_raw.dropna(subset=['Kode Rekening']).copy()
    df_m_clean['Kode Rekening'] = df_m_clean['Kode Rekening'].astype(str).str.strip()
    df_m_clean = df_m_clean[df_m_clean['Kode Rekening'].str.startswith('5.2')]
    df_m_map = df_m_clean[['Kode Rekening', 'Kode Kategori', 'Uraian Kategori']].drop_duplicates('Kode Rekening')

    return rek_persediaan, df_m_map, None

rek_persediaan, df_modal_map, error_msg = load_master_rak()

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135679.png", width=70)
    st.markdown("### **Panel Kontrol**")
    st.info("📌 **Master RAK Modal:** Terkunci otomatis (Akun `5.2...`)")
    
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
    with st.spinner("Sedang memproses data realisasi LRA..."):
        df_lra = pd.read_excel(f_lra, sheet_name='Data Realisasi Dokumen', header=4)
        df_lra['Kode Rekening'] = df_lra['Kode Rekening'].astype(str).str.strip()
        df_lra['Nilai Realisasi'] = pd.to_numeric(df_lra['Nilai Realisasi'], errors='coerce').fillna(0)

        df_rekon_pers = df_lra[df_lra['Kode Rekening'].isin(rek_persediaan)].copy()
        df_rekon_modal = df_lra[
            (df_lra['Kode Rekening'].str.startswith('5.2')) & 
            (df_lra['Kode Rekening'].isin(df_modal_map['Kode Rekening'].unique()))
        ].copy()
        df_rekon_modal = df_rekon_modal.merge(df_modal_map, on='Kode Rekening', how='left')

    # Filter SKPD
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

    # Metric Cards
    total_p = df_pers_filtered['Nilai Realisasi'].sum()
    total_m = df_modal_filtered['Nilai Realisasi'].sum()
    grand_total = total_p + total_m

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"""<div class="metric-box" style="background: linear-gradient(135deg, #0D9488, #14B8A6);"><div class="metric-title">📦 Total Belanja Persediaan (LRA)</div><div class="metric-value">{format_rupiah(total_p)}</div></div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="metric-box" style="background: linear-gradient(135deg, #4F46E5, #6366F1);"><div class="metric-title">🏢 Total Belanja Modal (LRA)</div><div class="metric-value">{format_rupiah(total_m)}</div></div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""<div class="metric-box" style="background: linear-gradient(135deg, #1E293B, #334155);"><div class="metric-title">📊 Total Gabungan Realisasi</div><div class="metric-value">{format_rupiah(grand_total)}</div></div>""", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📦 REKON PERSEDIAAN", "🏢 REKON BELANJA MODAL"])

    # ================= TAB 1: PERSEDIAAN =================
    with tab1:
        st.markdown("##### 📥 Pembanding Data SIPPER (Opsional)")
        f_sipper = st.file_uploader("Upload Data Aplikasi SIPPER (.xlsx)", type=["xlsx"], key="up_sipper")
        
        df_sipper_rekap = None
        if f_sipper:
            try:
                raw_sipper = pd.read_excel(f_sipper, header=None)
                rows_s = raw_sipper.iloc[11:].copy()
                df_s = rows_s[[5, 7]].dropna(subset=[5]).copy()
                df_s.columns = ['Kode Rekening', 'Nilai SIPPER']
                df_s['Kode Rekening'] = df_s['Kode Rekening'].astype(str).str.strip()
                df_s['Nilai SIPPER'] = pd.to_numeric(df_s['Nilai SIPPER'], errors='coerce').fillna(0)
                df_sipper_rekap = df_s.groupby('Kode Rekening')['Nilai SIPPER'].sum().reset_index()
                st.success("✅ Data Aplikasi SIPPER berhasil dimuat.")
            except Exception as e:
                st.error(f"Format file SIPPER tidak sesuai: {e}")

        st.markdown("---")
        st.subheader("1. Rekapitulasi per Kode Rekening Persediaan")
        if not df_pers_filtered.empty:
            rekap_p = df_pers_filtered.groupby(['Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
            
            if df_sipper_rekap is not None:
                rekap_p = pd.merge(rekap_p, df_sipper_rekap, on='Kode Rekening', how='outer').fillna(0)
                rekap_p['Nama Rekening'] = rekap_p['Nama Rekening'].replace(0, 'Dari Data SIPPER')
                rekap_p['Selisih'] = rekap_p['Nilai Realisasi'] - rekap_p['Nilai SIPPER']
                rekap_p['Status'] = rekap_p['Selisih'].apply(lambda x: '✅ COCOK' if round(x, 2) == 0 else '❌ SELISIH')
                
                tot_lra_p = rekap_p['Nilai Realisasi'].sum()
                tot_sipper_p = rekap_p['Nilai SIPPER'].sum()
                tot_selisih_p = rekap_p['Selisih'].sum()

                rows_html = "".join([
                    f"<tr><td>{r['Kode Rekening']}</td><td>{r['Nama Rekening']}</td><td class='text-right'>{format_rupiah(r['Nilai Realisasi'])}</td><td class='text-right'>{format_rupiah(r['Nilai SIPPER'])}</td><td class='text-right'>{format_rupiah(r['Selisih'])}</td><td class='text-center'>{r['Status']}</td></tr>"
                    for _, r in rekap_p.iterrows()
                ])
                tot_status_p = '✅ COCOK' if round(tot_selisih_p, 2) == 0 else '❌ SELISIH'
                
                table_html = f"<table class='custom-table'><thead><tr><th>Kode Rekening</th><th>Nama Rekening</th><th class='text-right'>Realisasi LRA (Rp)</th><th class='text-right'>Nilai SIPPER (Rp)</th><th class='text-right'>Selisih (Rp)</th><th class='text-center'>Status</th></tr></thead><tbody>{rows_html}<tr class='total-row'><td>TOTAL</td><td>JUMLAH KESELURUHAN</td><td class='text-right'>{format_rupiah(tot_lra_p)}</td><td class='text-right'>{format_rupiah(tot_sipper_p)}</td><td class='text-right'>{format_rupiah(tot_selisih_p)}</td><td class='text-center'>{tot_status_p}</td></tr></tbody></table>"
                st.markdown(table_html, unsafe_allow_html=True)

                if round(tot_selisih_p, 2) == 0:
                    st.markdown(f"<div class='status-card-match'><h4 style='color: #065F46; margin:0;'>✅ STATUS: COCOK DENGAN LRA</h4><p style='color: #047857; margin: 4px 0 0 0; font-size:14px;'>Total Realisasi LRA <b>({format_rupiah(tot_lra_p)})</b> sama persis dengan SIPPER <b>({format_rupiah(tot_sipper_p)})</b>. Tidak ada selisih.</p></div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<div class='status-card-diff'><h4 style='color: #991B1B; margin:0;'>⚠️ STATUS: TERDAPAT SELISIH REKONSILIASI</h4><p style='color: #B91C1C; margin: 4px 0 0 0; font-size:14px;'>Ditemukan selisih sebesar <b>{format_rupiah(tot_selisih_p)}</b> antara LRA ({format_rupiah(tot_lra_p)}) dan SIPPER ({format_rupiah(tot_sipper_p)}).</p></div>", unsafe_allow_html=True)
            else:
                tot_lra_p = rekap_p['Nilai Realisasi'].sum()
                rows_html = "".join([
                    f"<tr><td>{r['Kode Rekening']}</td><td>{r['Nama Rekening']}</td><td class='text-right'>{format_rupiah(r['Nilai Realisasi'])}</td></tr>"
                    for _, r in rekap_p.iterrows()
                ])
                table_html = f"<table class='custom-table'><thead><tr><th>Kode Rekening</th><th>Nama Rekening</th><th class='text-right'>Realisasi (Rp)</th></tr></thead><tbody>{rows_html}<tr class='total-row'><td>TOTAL</td><td>JUMLAH KESELURUHAN</td><td class='text-right'>{format_rupiah(tot_lra_p)}</td></tr></tbody></table>"
                st.markdown(table_html, unsafe_allow_html=True)
                st.info("💡 *Upload file SIPPER di atas untuk menampilkan perbandingan dan status selisih secara otomatis.*")

            st.markdown("---")
            st.subheader("2. Rincian Dokumen Realisasi LRA")
            df_detail_p = df_pers_filtered[[c for c in KOLOM_DETAIL_PILIHAN if c in df_pers_filtered.columns]].copy()
            df_detail_p['Nilai Realisasi'] = df_detail_p['Nilai Realisasi'].apply(format_rupiah)
            st.dataframe(df_detail_p, use_container_width=True, hide_index=True)
        else:
            st.warning("Tidak ditemukan data persediaan untuk SKPD ini.")

    # ================= TAB 2: BELANJA MODAL =================
    with tab2:
        st.markdown("##### 📥 Pembanding Data Aplikasi Belanja Modal / Aset (Opsional)")
        f_app_modal = st.file_uploader("Upload Data Rincian Pengadaan Aset (.xlsx)", type=["xlsx"], key="up_modal")
        
        df_modal_app_rekap = None
        if f_app_modal:
            try:
                raw_m = pd.read_excel(f_app_modal, header=None)
                rows_m = raw_m[raw_m[0].astype(str).str.startswith('5.2')].copy()
                df_m = rows_m[[0, 3]].dropna(subset=[0]).copy()
                df_m.columns = ['Kode Rekening', 'Nilai Aset']
                df_m['Kode Rekening'] = df_m['Kode Rekening'].astype(str).str.strip()
                df_m['Nilai Aset'] = pd.to_numeric(df_m['Nilai Aset'], errors='coerce').fillna(0)
                df_modal_app_rekap = df_m.groupby('Kode Rekening')['Nilai Aset'].sum().reset_index()
                st.success("✅ Data Aplikasi Belanja Modal / Aset berhasil dimuat.")
            except Exception as e:
                st.error(f"Format file Belanja Modal tidak sesuai: {e}")

        st.markdown("---")
        st.subheader("1. Rekapitulasi per Kategori & Rekening Belanja Modal")
        if not df_modal_filtered.empty:
            rekap_m = df_modal_filtered.groupby(['Kode Kategori', 'Uraian Kategori', 'Kode Rekening', 'Nama Rekening'])['Nilai Realisasi'].sum().reset_index()
            
            if df_modal_app_rekap is not None:
                rekap_m = pd.merge(rekap_m, df_modal_app_rekap, on='Kode Rekening', how='outer').fillna(0)
                rekap_m['Kode Kategori'] = rekap_m['Kode Kategori'].replace(0, '-')
                rekap_m['Uraian Kategori'] = rekap_m['Uraian Kategori'].replace(0, '-')
                rekap_m['Nama Rekening'] = rekap_m['Nama Rekening'].replace(0, 'Dari Data Aplikasi Aset')
                rekap_m['Selisih'] = rekap_m['Nilai Realisasi'] - rekap_m['Nilai Aset']
                rekap_m['Status'] = rekap_m['Selisih'].apply(lambda x: '✅ COCOK' if round(x, 2) == 0 else '❌ SELISIH')

                tot_lra_m = rekap_m['Nilai Realisasi'].sum()
                tot_aset_m = rekap_m['Nilai Aset'].sum()
                tot_selisih_m = rekap_m['Selisih'].sum()

                rows_html = "".join([
                    f"<tr><td>{r['Kode Kategori']}</td><td>{r['Uraian Kategori']}</td><td>{r['Kode Rekening']}</td><td>{r['Nama Rekening']}</td><td class='text-right'>{format_rupiah(r['Nilai Realisasi'])}</td><td class='text-right'>{format_rupiah(r['Nilai Aset'])}</td><td class='text-right'>{format_rupiah(r['Selisih'])}</td><td class='text-center'>{r['Status']}</td></tr>"
                    for _, r in rekap_m.iterrows()
                ])
                tot_status_m = '✅ COCOK' if round(tot_selisih_m, 2) == 0 else '❌ SELISIH'
                
                table_html = f"<table class='custom-table'><thead><tr><th>Kode Kategori</th><th>Uraian Kategori</th><th>Kode Rekening</th><th>Nama Rekening</th><th class='text-right'>Realisasi LRA (Rp)</th><th class='text-right'>Nilai Aplikasi Aset (Rp)</th><th class='text-right'>Selisih (Rp)</th><th class='text-center'>Status</th></tr></thead><tbody>{rows_html}<tr class='total-row'><td>TOTAL</td><td>JUMLAH KESELURUHAN</td><td>-</td><td>-</td><td class='text-right'>{format_rupiah(tot_lra_m)}</td><td class='text-right'>{format_rupiah(tot_aset_m)}</td><td class='text-right'>{format_rupiah(tot_selisih_m)}</td><td class='text-center'>{tot_status_m}</td></tr></tbody></table>"
                st.markdown(table_html, unsafe_allow_html=True)

                if round(tot_selisih_m, 2) == 0:
                    st.markdown(f"<div class='status-card-match'><h4 style='color: #065F46; margin:0;'>✅ STATUS: COCOK DENGAN LRA</h4><p style='color: #047857; margin: 4px 0 0 0; font-size:14px;'>Total Belanja Modal LRA <b>({format_rupiah(tot_lra_m)})</b> sama persis dengan Aplikasi Aset <b>({format_rupiah(tot_aset_m)})</b>. Tidak ada selisih.</p></div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<div class='status-card-diff'><h4 style='color: #991B1B; margin:0;'>⚠️ STATUS: TERDAPAT SELISIH REKONSILIASI</h4><p style='color: #B91C1C; margin: 4px 0 0 0; font-size:14px;'>Ditemukan selisih sebesar <b>{format_rupiah(tot_selisih_m)}</b> antara Belanja Modal LRA ({format_rupiah(tot_lra_m)}) dan Aplikasi Aset ({format_rupiah(tot_aset_m)}).</p></div>", unsafe_allow_html=True)
            else:
                tot_lra_m = rekap_m['Nilai Realisasi'].sum()
                rows_html = "".join([
                    f"<tr><td>{r['Kode Kategori']}</td><td>{r['Uraian Kategori']}</td><td>{r['Kode Rekening']}</td><td>{r['Nama Rekening']}</td><td class='text-right'>{format_rupiah(r['Nilai Realisasi'])}</td></tr>"
                    for _, r in rekap_m.iterrows()
                ])
                table_html = f"<table class='custom-table'><thead><tr><th>Kode Kategori</th><th>Uraian Kategori</th><th>Kode Rekening</th><th>Nama Rekening</th><th class='text-right'>Realisasi (Rp)</th></tr></thead><tbody>{rows_html}<tr class='total-row'><td>TOTAL</td><td>JUMLAH KESELURUHAN</td><td>-</td><td>-</td><td class='text-right'>{format_rupiah(tot_lra_m)}</td></tr></tbody></table>"
                st.markdown(table_html, unsafe_allow_html=True)
                st.info("💡 *Upload file Pengadaan Aset di atas untuk menampilkan perbandingan dan status selisih secara otomatis.*")

            st.markdown("---")
            st.subheader("2. Rincian Dokumen Realisasi LRA")
            df_detail_m = df_modal_filtered[[c for c in KOLOM_DETAIL_PILIHAN if c in df_modal_filtered.columns]].copy()
            df_detail_m['Nilai Realisasi'] = df_detail_m['Nilai Realisasi'].apply(format_rupiah)
            st.dataframe(df_detail_m, use_container_width=True, hide_index=True)
        else:
            st.warning("Tidak ditemukan data belanja modal untuk SKPD ini.")

else:
    st.info("👈 Silakan upload file **`LRA REALISASI JAN-JUNI30.xlsx`** pada menu di sebelah kiri.")
