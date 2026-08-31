import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Rekonsiliasi SIPD vs BMD per Dokumen", layout="wide"
)

st.title("📑 Rekonsiliasi Belanja Modal & Persediaan (SIPD vs BMD)")
st.caption(
    "Pencocokan Data Realisasi LRA SIPD per Dokumen/SP2D dengan Register Entrian SKPD"
)

# Upload File
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Data LRA SIPD (per SP2D / Dokumen)")
    file_sipd = st.file_uploader(
        "Unggah Ekspor SIPD (Excel/CSV)", type=["xlsx", "xls", "csv"], key="sipd"
    )

with col2:
    st.subheader("2. Data Entrian BMD / Pengurus Barang")
    file_bmd = st.file_uploader(
        "Unggah Ekspor BMD (Excel/CSV)", type=["xlsx", "xls", "csv"], key="bmd"
    )


def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, dtype=str)
    return pd.read_excel(file, dtype=str)


def clean_number(series):
    return (
        series.astype(str)
        .str.replace(r"[^\d.-]", "", regex=True)
        .replace("", "0")
        .astype(float)
    )


def clean_doc_no(series):
    # Membersihkan spasi dan huruf kapital agar matching tidak gagal karena typo spasi
    return series.astype(str).str.strip().str.upper()


if file_sipd and file_bmd:
    df_sipd_raw = load_file(file_sipd)
    df_bmd_raw = load_file(file_bmd)

    st.write("---")
    st.subheader("⚙️ Pemetaan Kolom Data")

    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    with col_s1:
        skpd_sipd = st.selectbox("SKPD / Sub-Unit (SIPD)", df_sipd_raw.columns)
        doc_sipd = st.selectbox("No. Dokumen / SP2D (SIPD)", df_sipd_raw.columns)
    with col_s2:
        rek_sipd = st.selectbox("Kode/Nama Rekening (SIPD)", df_sipd_raw.columns)
        val_sipd = st.selectbox("Nilai Realisasi (SIPD)", df_sipd_raw.columns)

    with col_s3:
        skpd_bmd = st.selectbox("SKPD / Unit (BMD)", df_bmd_raw.columns)
        doc_bmd = st.selectbox("No. Dokumen / SP2D (BMD)", df_bmd_raw.columns)
    with col_s4:
        rek_bmd = st.selectbox("Kode/Akun Aset (BMD)", df_bmd_raw.columns)
        val_bmd = st.selectbox("Nilai Entrian (BMD)", df_bmd_raw.columns)

    # Pra-pemrosesan Data
    df_sipd = pd.DataFrame(
        {
            "SKPD": df_sipd_raw[skpd_sipd].astype(str).str.strip(),
            "No_Dokumen": clean_doc_no(df_sipd_raw[doc_sipd]),
            "Rekening_SIPD": df_sipd_raw[rek_sipd].astype(str).str.strip(),
            "Nilai_SIPD": clean_number(df_sipd_raw[val_sipd]),
        }
    )

    df_bmd = pd.DataFrame(
        {
            "SKPD": df_bmd_raw[skpd_bmd].astype(str).str.strip(),
            "No_Dokumen": clean_doc_no(df_bmd_raw[doc_bmd]),
            "Rekening_BMD": df_bmd_raw[rek_bmd].astype(str).str.strip(),
            "Nilai_BMD": clean_number(df_bmd_raw[val_bmd]),
        }
    )

    # Agregasi per Dokumen jika dalam 1 SP2D terdapat rincian barang ganda
    sipd_agg = (
        df_sipd.groupby(["SKPD", "No_Dokumen", "Rekening_SIPD"], as_index=False)[
            "Nilai_SIPD"
        ].sum()
    )
    bmd_agg = (
        df_bmd.groupby(["SKPD", "No_Dokumen", "Rekening_BMD"], as_index=False)[
            "Nilai_BMD"
        ].sum()
    )

    # Matching menggunakan Outer Join pada SKPD dan Nomor Dokumen
    df_merged = pd.merge(
        sipd_agg,
        bmd_agg,
        on=["SKPD", "No_Dokumen"],
        how="outer",
    ).fillna(
        {
            "Nilai_SIPD": 0,
            "Nilai_BMD": 0,
            "Rekening_SIPD": "-",
            "Rekening_BMD": "-",
        }
    )

    df_merged["Selisih"] = df_merged["Nilai_SIPD"] - df_merged["Nilai_BMD"]

    # Klasifikasi Status Hasil Rekonsiliasi
    conditions = [
        (df_merged["Nilai_SIPD"] > 0)
        & (df_merged["Nilai_BMD"] > 0)
        & (abs(df_merged["Selisih"]) < 1),
        (df_merged["Nilai_SIPD"] > 0)
        & (df_merged["Nilai_BMD"] > 0)
        & (abs(df_merged["Selisih"]) >= 1),
        (df_merged["Nilai_SIPD"] > 0) & (df_merged["Nilai_BMD"] == 0),
        (df_merged["Nilai_SIPD"] == 0) & (df_merged["Nilai_BMD"] > 0),
    ]
    choices = [
        "✅ Klop / Sesuai",
        "⚠️ Selisih Nilai SP2D",
        "❌ Ada di SIPD, Belum di BMD",
        "❓ Ada di BMD, Tidak ada di SIPD",
    ]
    df_merged["Status_Rekon"] = np.select(
        conditions, choices, default="Lainnya"
    )

    # Ringkasan Eksekutif
    st.write("---")
    st.subheader("📊 Ringkasan Hasil Rekon")

    # Filter SKPD jika multi-SKPD
    list_skpd = ["SEMUA SKPD"] + sorted(df_merged["SKPD"].unique().tolist())
    selected_skpd = st.selectbox("Pilih SKPD untuk Dilihat:", list_skpd)

    if selected_skpd != "SEMUA SKPD":
        df_view = df_merged[df_merged["SKPD"] == selected_skpd]
    else:
        df_view = df_merged

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total SP2D SIPD", f"Rp {df_view['Nilai_SIPD'].sum():,.0f}")
    kpi2.metric("Total Entrian BMD", f"Rp {df_view['Nilai_BMD'].sum():,.0f}")
    kpi3.metric("Total Selisih", f"Rp {df_view['Selisih'].sum():,.0f}")
    kpi4.metric(
        "Dokumen Belum Sesuai",
        f"{(df_view['Status_Rekon'] != '✅ Klop / Sesuai').sum()} Dokumen",
    )

    # Filter Berdasarkan Status
    st.write("---")
    st.subheader("🔍 Rincian Transaksi per Dokumen")

    filter_status = st.multiselect(
        "Filter Kategori Status:",
        options=choices,
        default=[
            "⚠️ Selisih Nilai SP2D",
            "❌ Ada di SIPD, Belum di BMD",
            "❓ Ada di BMD, Tidak ada di SIPD",
        ],
    )

    df_filtered = df_view[df_view["Status_Rekon"].isin(filter_status)]

    st.dataframe(
        df_filtered[
            [
                "SKPD",
                "No_Dokumen",
                "Rekening_SIPD",
                "Nilai_SIPD",
                "Rekening_BMD",
                "Nilai_BMD",
                "Selisih",
                "Status_Rekon",
            ]
        ].style.format(
            {
                "Nilai_SIPD": "Rp {:,.2f}",
                "Nilai_BMD": "Rp {:,.2f}",
                "Selisih": "Rp {:,.2f}",
            }
        ),
        use_container_width=True,
    )

    # Export Rekap & Detail ke Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: Detail Per Dokumen
        df_view.to_excel(
            writer, index=False, sheet_name="Rincian_per_Dokumen"
        )

        # Sheet 2: Rekapitulasi per SKPD
        rekap_skpd = (
            df_merged.groupby("SKPD")
            .agg(
                Total_SIPD=("Nilai_SIPD", "sum"),
                Total_BMD=("Nilai_BMD", "sum"),
                Total_Selisih=("Selisih", "sum"),
                Jumlah_Dok_Belum_Klop=(
                    "Status_Rekon",
                    lambda x: (x != "✅ Klop / Sesuai").sum(),
                ),
            )
            .reset_index()
        )
        rekap_skpd.to_excel(writer, index=False, sheet_name="Rekap_per_SKPD")

    st.download_button(
        label="📥 Unduh Kertas Kerja Rekonsiliasi Lengkap (Excel)",
        data=output.getvalue(),
        file_name="Kertas_Kerja_Rekon_SIPD_vs_BMD.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
