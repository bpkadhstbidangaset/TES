import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistem Rekonsiliasi Aset & LRA",
    page_icon="📑",
    layout="wide"
)

st.title("📑 Sistem Rekonsiliasi Belanja Modal & Persediaan vs LRA")
st.markdown("Alat pencocokan otomatis realisasi belanja anggaran (LRA) dengan mutasi penatausahaan Aset/Persediaan.")

# Upload File Template Excel
uploaded_file = st.file_uploader(
    "Unggah File Workbook Excel Hasil Pengisian SKPD (.xlsx)", 
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        # Membaca sheet LRA
        df_lra = pd.read_excel(uploaded_file, sheet_name="Input_LRA")
        # Membaca sheet Mutasi Aset
        df_aset = pd.read_excel(uploaded_file, sheet_name="Input_Mutasi_Aset")
        # Membaca sheet Persediaan
        df_persediaan = pd.read_excel(uploaded_file, sheet_name="Input_Persediaan")

        st.success("File berhasil dimuat!")

        # 1. Pilihan Mode Rekonsiliasi
        tab1, tab2 = st.tabs(["🏛️ Rekonsiliasi Belanja Modal (Aset Tetap)", "📦 Rekonsiliasi Belanja Persediaan"])

        # TAB 1: BELANJA MODAL
        with tab1:
            st.subheader("Pencocokan Belanja Modal (Akun 5.2.x)")
            
            # Filter LRA belanja modal saja
            df_lra_modal = df_lra[df_lra['Kode_Rekening'].astype(str).str.startswith('5.2')].copy()
            rekap_lra_modal = df_lra_modal.groupby(
                ['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening', 'Uraian_Akun_Belanja'], 
                as_index=False
            )['Realisasi_LRA'].sum()

            rekap_aset = df_aset.groupby(
                ['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening_Belanja'], 
                as_index=False
            )['Total_Catat_Aset'].sum().rename(columns={'Kode_Rekening_Belanja': 'Kode_Rekening'})

            # Merge
            rekon_modal = pd.merge(
                rekap_lra_modal, 
                rekap_aset, 
                on=['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening'], 
                how='outer'
            ).fillna(0)

            rekon_modal['Selisih'] = rekon_modal['Realisasi_LRA'] - rekon_modal['Total_Catat_Aset']
            
            def get_status_modal(row):
                if row['Selisih'] == 0:
                    return "Cocok / Seimbang"
                elif row['Selisih'] > 0:
                    return "Kurang Catat di Aset (BMD Belum Terbit BAST/Input)"
                else:
                    return "Aset Lebih Besar (Non-Kas / Kapitalisasi / Hibah)"

            rekon_modal['Status'] = rekon_modal.apply(get_status_modal, axis=1)

            # Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("Total LRA Modal", f"Rp {rekon_modal['Realisasi_LRA'].sum():,.0f}")
            m2.metric("Total Input Aset Tetap", f"Rp {rekon_modal['Total_Catat_Aset'].sum():,.0f}")
            m3.metric("Total Selisih", f"Rp {rekon_modal['Selisih'].sum():,.0f}")

            st.dataframe(rekon_modal.style.format({
                'Realisasi_LRA': 'Rp {:,.0f}',
                'Total_Catat_Aset': 'Rp {:,.0f}',
                'Selisih': 'Rp {:,.0f}'
            }), use_container_width=True)

        # TAB 2: PERSEDIAAN
        with tab2:
            st.subheader("Pencocokan Belanja Persediaan (Akun 5.1.02.01)")
            
            df_lra_persediaan = df_lra[df_lra['Kode_Rekening'].astype(str).str.startswith('5.1.02.01')].copy()
            rekap_lra_persediaan = df_lra_persediaan.groupby(
                ['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening', 'Uraian_Akun_Belanja'], 
                as_index=False
            )['Realisasi_LRA'].sum()

            rekap_persediaan = df_persediaan.groupby(
                ['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening_Belanja'], 
                as_index=False
            )['Total_Nilai_Pengadaan'].sum().rename(columns={'Kode_Rekening_Belanja': 'Kode_Rekening'})

            # Merge
            rekon_persediaan = pd.merge(
                rekap_lra_persediaan, 
                rekap_persediaan, 
                on=['Kode_SKPD', 'Nama_SKPD', 'Kode_Rekening'], 
                how='outer'
            ).fillna(0)

            rekon_persediaan['Selisih'] = rekon_persediaan['Realisasi_LRA'] - rekon_persediaan['Total_Nilai_Pengadaan']

            def get_status_persediaan(row):
                if row['Selisih'] == 0:
                    return "Cocok / Seimbang"
                elif row['Selisih'] > 0:
                    return "Kurang Catat di Mutasi Persediaan"
                else:
                    return "Persediaan Lebih Besar (Koreksi / Saldo Awal)"

            rekon_persediaan['Status'] = rekon_persediaan.apply(get_status_persediaan, axis=1)

            # Metrics
            p1, p2, p3 = st.columns(3)
            p1.metric("Total LRA Persediaan", f"Rp {rekon_persediaan['Realisasi_LRA'].sum():,.0f}")
            p2.metric("Total Pengadaan Masuk", f"Rp {rekon_persediaan['Total_Nilai_Pengadaan'].sum():,.0f}")
            p3.metric("Total Selisih", f"Rp {rekon_persediaan['Selisih'].sum():,.0f}")

            st.dataframe(rekon_persediaan.style.format({
                'Realisasi_LRA': 'Rp {:,.0f}',
                'Total_Nilai_Pengadaan': 'Rp {:,.0f}',
                'Selisih': 'Rp {:,.0f}'
            }), use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
