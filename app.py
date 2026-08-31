import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistem Rekonsiliasi BMD vs LRA SIPD", layout="wide"
)

st.title("🏛️ Aplikasi Rekonsiliasi Belanja Modal & Persediaan BMD vs LRA SIPD")
st.caption(
    "Dirancang khusus untuk rekonsiliasi data LRA SIPD Hulu Sungai Tengah & Register BMD"
)


# --- 1. FUNGSI PARSER DATA ---
def load_lra_sipd(file):
    """Membaca file LRA SIPD yang headernya berada pada baris ke-5"""
    df = pd.read_excel(file, skiprows=4)
    df.columns = [str(c).strip() for c in df.columns]

    # Standardisasi tipe data dan pembersihan string
    df["Kode Rekening"] = df["Kode Rekening"].astype(str).str.strip()
    df["Nama Rekening"] = df["Nama Rekening"].astype(str).str.strip()
    df["Nama SKPD"] = df["Nama SKPD"].astype(str).str.strip()
    df["Nomor SP2D"] = df["Nomor SP2D"].astype(str).str.strip()
    df["Nilai Realisasi"] = pd.to_numeric(
        df["Nilai Realisasi"], errors="coerce"
    ).fillna(0)
    return df


def load_bmd_entry(file):
    """Membaca file entrian BMD bertingkat seperti DISDIK Data Aplikasi"""
    raw = pd.read_excel(file, header=None)

    # Identifikasi baris header tabel (mencari baris bertuliskan 'KODE' atau 'PENGADAAN')
    header_idx = None
    for idx, row in raw.iterrows():
        if "KODE" in str(row.values) and "PENGADAAN" in str(row.values):
            header_idx = idx
            break

    if header_idx is None:
        header_idx = 7  # default fallback

    df = pd.read_excel(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Bersihkan nama kolom umum
    col_mapping = {}
    for col in df.columns:
        c_upper = col.upper()
        if "KODE" in c_upper:
            col_mapping[col] = "Kode"
        elif "URAIAN" in c_upper or "DESKRIPSI" in c_upper:
            col_mapping[col] = "Uraian"
        elif "PENGADAAN" in c_upper:
            col_mapping[col] = "Nilai_Pengadaan"
        elif "ASET" in c_upper:
            col_mapping[col] = "Nilai_Aset"
        elif "SELISIH" in c_upper:
            col_mapping[col] = "Selisih_BMD"

    df = df.rename(columns=col_mapping)
    for col in ["Nilai_Pengadaan", "Nilai_Aset", "Selisih_BMD"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def load_rak_reference(file_rak):
    """Membaca file kamus/referensi akun RAK"""
    df = pd.read_excel(file_rak)
    # Cari baris yang memuat 'KODE REKENING'
    header_row = 0
    for idx, row in df.iterrows():
        if "KODE REKENING" in str(row.values).upper():
            header_row = idx + 1
            break
    df_clean = pd.read_excel(file_rak, skiprows=header_row)
    df_clean.columns = [str(c).strip().upper() for c in df_clean.columns]
    return df_clean


# --- 2. KOMPONEN UPLOAD FILE ---
st.sidebar.header("📁 Unggah Dokumen")
file_lra = st.sidebar.file_uploader(
    "1. File LRA SIPD (.xlsx)", type=["xlsx", "xls"], key="lra"
)
file_bmd = st.sidebar.file_uploader(
    "2. File Register BMD SKPD (.xlsx)", type=["xlsx", "xls"], key="bmd"
)
file_rak_modal = st.sidebar.file_uploader(
    "3. RAK Belanja Modal (Opsional)", type=["xlsx", "xls"], key="rak_m"
)
file_rak_persediaan = st.sidebar.file_uploader(
    "4. RAK Persediaan (Opsional)", type=["xlsx", "xls"], key="rak_p"
)


# --- 3. PEMROSESAN & TAMPILAN TAB ---
if file_lra and file_bmd:
    df_lra = load_lra_sipd(file_lra)
    df_bmd = load_bmd_entry(file_bmd)

    # Filter Berdasarkan SKPD
    daftar_skpd = sorted(df_lra["Nama SKPD"].dropna().unique().tolist())
    selected_skpd = st.selectbox("🏢 Pilih SKPD untuk Direkonsiliasi:", daftar_skpd)
    df_lra_skpd = df_lra[df_lra["Nama SKPD"] == selected_skpd]

    tab_modal, tab_persediaan, tab_download = st.tabs(
        [
            "🏢 Rekon Belanja Modal (Akun 5.2.x & Kapitalisasi)",
            "📦 Rekon Belanja Persediaan (Akun 5.1.02.01.x)",
            "📥 Download Kertas Kerja Lengkap",
        ]
    )

    # -------------------------------------------------------------
    # TAB 1: BELANJA MODAL
    # -------------------------------------------------------------
    with tab_modal:
        st.subheader("1. Belanja Modal Murni (Akun 5.2.x)")
        lra_modal_52 = df_lra_skpd[
            df_lra_skpd["Kode Rekening"].str.startswith("5.2")
        ]

        # Cek juga belanja jasa konsultansi / pengawasan yang masuk kapitalisasi (5.1.02.02.008)
        st.caption(
            "💡 Termasuk Belanja Modal Akun 5.2 dan Belanja Jasa Konsultansi/Perencanaan Proyek (Kapitalisasi)"
        )
        lra_kapitalisasi = df_lra_skpd[
            df_lra_skpd["Kode Rekening"].str.startswith("5.1.02.02.008")
        ]
        lra_modal_gabung = pd.concat(
            [lra_modal_52, lra_kapitalisasi], ignore_index=True
        )

        # Rekap per Rekening Belanja
        rekap_rekening_modal = (
            lra_modal_gabung.groupby(
                ["Kode Rekening", "Nama Rekening"], as_index=False
            )["Nilai Realisasi"]
            .sum()
            .rename(columns={"Nilai Realisasi": "Realisasi LRA"})
        )

        col_m1, col_m2, col_m3 = st.columns(3)
        total_lra_modal = lra_modal_gabung["Nilai Realisasi"].sum()
        total_bmd_modal = (
            df_bmd["Nilai_Aset"].sum()
            if "Nilai_Aset" in df_bmd.columns
            else df_bmd["Nilai_Pengadaan"].sum()
        )
        selisih_modal = total_lra_modal - total_bmd_modal

        col_m1.metric("Total Belanja Modal LRA", f"Rp {total_lra_modal:,.0f}")
        col_m2.metric("Total Aset Masuk BMD", f"Rp {total_bmd_modal:,.0f}")
        col_m3.metric(
            "Selisih Belanja Modal",
            f"Rp {selisih_modal:,.0f}",
            delta_color="inverse",
        )

        st.markdown("#### Ringkasan Realisasi per Kode Rekening Belanja")
        st.dataframe(
            rekap_rekening_modal.style.format({"Realisasi LRA": "Rp {:,.2f}"}),
            use_container_width=True,
        )

        st.markdown("#### Rincian Realisasi Dokumen SP2D Modal di LRA")
        st.dataframe(
            lra_modal_gabung[
                [
                    "Nomor SP2D",
                    "Kode Rekening",
                    "Nama Rekening",
                    "Keterangan Dokumen",
                    "Nilai Realisasi",
                ]
            ].style.format({"Nilai Realisasi": "Rp {:,.2f}"}),
            use_container_width=True,
        )

    # -------------------------------------------------------------
    # TAB 2: BELANJA PERSEDIAAN
    # -------------------------------------------------------------
    with tab_persediaan:
        st.subheader("2. Belanja Persediaan / Pakai Habis (Akun 5.1.02.01.x)")
        lra_persediaan = df_lra_skpd[
            df_lra_skpd["Kode Rekening"].str.startswith("5.1.02.01")
        ]

        rekap_persediaan = (
            lra_persediaan.groupby(
                ["Kode Rekening", "Nama Rekening"], as_index=False
            )["Nilai Realisasi"]
            .sum()
            .rename(columns={"Nilai Realisasi": "Realisasi LRA Persediaan"})
        )

        tot_lra_persediaan = lra_persediaan["Nilai Realisasi"].sum()
        st.metric(
            "Total Realisasi Belanja Persediaan di LRA",
            f"Rp {tot_lra_persediaan:,.0f}",
        )

        st.markdown("#### Daftar Rekening Belanja Persediaan")
        st.dataframe(
            rekap_persediaan.style.format(
                {"Realisasi LRA Persediaan": "Rp {:,.2f}"}
            ),
            use_container_width=True,
        )

        st.markdown("#### Rincian Transaksi Belanja Persediaan per SP2D")
        st.dataframe(
            lra_persediaan[
                [
                    "Nomor SP2D",
                    "Kode Rekening",
                    "Nama Rekening",
                    "Keterangan Dokumen",
                    "Nilai Realisasi",
                ]
            ].style.format({"Nilai Realisasi": "Rp {:,.2f}"}),
            use_container_width=True,
        )

    # -------------------------------------------------------------
    # TAB 3: UNDUH KERTAS KERJA EXCEL
    # -------------------------------------------------------------
    with tab_download:
        st.subheader("📥 Unduh Kertas Kerja Rekonsiliasi Format Excel")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            rekap_rekening_modal.to_excel(
                writer, index=False, sheet_name="Rekap_Modal"
            )
            lra_modal_gabung[
                [
                    "Nomor SP2D",
                    "Kode Rekening",
                    "Nama Rekening",
                    "Keterangan Dokumen",
                    "Nilai Realisasi",
                ]
            ].to_excel(writer, index=False, sheet_name="Rincian_SP2D_Modal")
            rekap_persediaan.to_excel(
                writer, index=False, sheet_name="Rekap_Persediaan"
            )
            lra_persediaan[
                [
                    "Nomor SP2D",
                    "Kode Rekening",
                    "Nama Rekening",
                    "Keterangan Dokumen",
                    "Nilai Realisasi",
                ]
            ].to_excel(
                writer, index=False, sheet_name="Rincian_SP2D_Persediaan"
            )

        st.download_button(
            label="📊 Download Excel Kertas Kerja Rekon",
            data=output.getvalue(),
            file_name=f"Kertas_Kerja_Rekon_{selected_skpd}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info(
        "👋 Silakan unggah minimal file **LRA SIPD** dan **File Register BMD SKPD** pada panel samping kiri."
    )
