import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistem Rekon BMD vs SIPD", layout="wide")

st.title("📑 Sistem Rekonsiliasi Otomatis: SIPD vs BMD")
st.caption(
    "Pemisahan Otomatis Rekonsiliasi Belanja Modal (Akun 5.2) & Belanja Persediaan (Akun 5.1.02.01)"
)

# --- FUNGSI BANTUAN ---
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


def clean_str(series):
    return series.astype(str).str.strip().str.upper()


def process_reconciliation(df_sipd_sub, df_bmd_sub, jenis_rekon):
    """Fungsi umum untuk matching data per SKPD dan No Dokumen"""
    if df_sipd_sub.empty and df_bmd_sub.empty:
        st.warning(f"Tidak ada data ditemukan untuk {jenis_rekon}.")
        return None

    # Agregasi per Dokumen
    sipd_agg = (
        df_sipd_sub.groupby(["SKPD", "No_Dokumen", "Rekening"], as_index=False)["Nilai_SIPD"]
        .sum()
    )
    bmd_agg = (
        df_bmd_sub.groupby(["SKPD", "No_Dokumen"], as_index=False)["Nilai_BMD"]
        .sum()
    )

    # Outer Join
    merged = pd.merge(sipd_agg, bmd_agg, on=["SKPD", "No_Dokumen"], how="outer").fillna(
        {"Nilai_SIPD": 0, "Nilai_BMD": 0, "Rekening": "-"}
    )

    merged["Selisih"] = merged["Nilai_SIPD"] - merged["Nilai_BMD"]

    # Klasifikasi Status
    conditions = [
        (merged["Nilai_SIPD"] > 0) & (merged["Nilai_BMD"] > 0) & (abs(merged["Selisih"]) < 1),
        (merged["Nilai_SIPD"] > 0) & (merged["Nilai_BMD"] > 0) & (abs(merged["Selisih"]) >= 1),
        (merged["Nilai_SIPD"] > 0) & (merged["Nilai_BMD"] == 0),
        (merged["Nilai_SIPD"] == 0) & (merged["Nilai_BMD"] > 0),
    ]
    choices = [
        "✅ Klop / Sesuai",
        "⚠️ Selisih Nilai",
        "❌ Ada di SIPD, Belum di BMD",
        "❓ Ada di BMD, Belum/Tidak ada di SIPD",
    ]
    merged["Status_Rekon"] = np.select(conditions, choices, default="Lainnya")
    return merged


def render_tab_content(df_rekon, label_rekon):
    """Fungsi untuk menampilkan visual, ringkasan metrik, dan tabel interaktif"""
    if df_rekon is None or df_rekon.empty:
        st.info(f"Belum ada data untuk {label_rekon}.")
        return

    # Filter SKPD
    list_skpd = ["SEMUA SKPD"] + sorted(df_rekon["SKPD"].unique().tolist())
    selected_skpd = st.selectbox(f"Pilih SKPD ({label_rekon}):", list_skpd, key=f"skpd_{label_rekon}")

    if selected_skpd != "SEMUA SKPD":
        df_view = df_rekon[df_rekon["SKPD"] == selected_skpd]
    else:
        df_view = df_rekon

    # Metric Cards
    m1, m2, m3, m4 = st.columns(4)
    tot_sipd = df_view["Nilai_SIPD"].sum()
    tot_bmd = df_view["Nilai_BMD"].sum()
    tot_selisih = df_view["Selisih"].sum()
    unmatched_count = (df_view["Status_Rekon"] != "✅ Klop / Sesuai").sum()

    m1.metric("Total LRA SIPD", f"Rp {tot_sipd:,.0f}")
    m2.metric("Total Mutasi BMD", f"Rp {tot_bmd:,.0f}")
    m3.metric("Total Selisih", f"Rp {tot_selisih:,.0f}")
    m4.metric("Item Belum Klop", f"{unmatched_count} Dokumen")

    # Filter Tabel
    st.write("---")
    st.subheader(f"Daftar Transaksi: {label_rekon}")

    filter_status = st.multiselect(
        f"Filter Status Transaksi ({label_rekon}):",
        options=[
            "✅ Klop / Sesuai",
            "⚠️ Selisih Nilai",
            "❌ Ada di SIPD, Belum di BMD",
            "❓ Ada di BMD, Belum/Tidak ada di SIPD",
        ],
        default=[
            "⚠️ Selisih Nilai",
            "❌ Ada di SIPD, Belum di BMD",
            "❓ Ada di BMD, Belum/Tidak ada di SIPD",
        ],
        key=f"status_{label_rekon}",
    )

    df_filtered = df_view[df_view["Status_Rekon"].isin(filter_status)]

    st.dataframe(
        df_filtered.style.format(
            {
                "Nilai_SIPD": "Rp {:,.2f}",
                "Nilai_BMD": "Rp {:,.2f}",
                "Selisih": "Rp {:,.2f}",
            }
        ),
        use_container_width=True,
    )


# --- SECTION 1: UPLOAD DATA ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    st.subheader("1. Data Realisasi LRA SIPD (Seluruh Belanja)")
    file_sipd = st.file_uploader("Unggah File SIPD (Excel/CSV)", type=["xlsx", "xls", "csv"], key="sipd")

with col_up2:
    st.subheader("2. Data Entrian BMD (Aset Tetap & Persediaan)")
    file_bmd = st.file_uploader("Unggah File BMD/SIMBADA (Excel/CSV)", type=["xlsx", "xls", "csv"], key="bmd")


# --- SECTION 2: PROSES & PEMISAHAN TAB ---
if file_sipd and file_bmd:
    df_sipd_raw = load_file(file_sipd)
    df_bmd_raw = load_file(file_bmd)

    with st.expander("⚙️ Konfigurasi Pemetaan Kolom", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            col_skpd_sipd = st.selectbox("SKPD (SIPD)", df_sipd_raw.columns)
            col_doc_sipd = st.selectbox("No. SP2D / Dokumen (SIPD)", df_sipd_raw.columns)
        with c2:
            col_rek_sipd = st.selectbox("Kode Rekening Belanja (SIPD)", df_sipd_raw.columns)
            col_val_sipd = st.selectbox("Nilai Realisasi (SIPD)", df_sipd_raw.columns)
        with c3:
            col_skpd_bmd = st.selectbox("SKPD (BMD)", df_bmd_raw.columns)
            col_doc_bmd = st.selectbox("No. SP2D / Dokumen (BMD)", df_bmd_raw.columns)
        with c4:
            col_rek_bmd = st.selectbox("Kode/Jenis Aset (BMD)", df_bmd_raw.columns)
            col_val_bmd = st.selectbox("Nilai Entrian (BMD)", df_bmd_raw.columns)

    # Standardisasi Dataframe Dasar
    df_sipd_clean = pd.DataFrame(
        {
            "SKPD": clean_str(df_sipd_raw[col_skpd_sipd]),
            "No_Dokumen": clean_str(df_sipd_raw[col_doc_sipd]),
            "Rekening": clean_str(df_sipd_raw[col_rek_sipd]),
            "Nilai_SIPD": clean_number(df_sipd_raw[col_val_sipd]),
        }
    )

    df_bmd_clean = pd.DataFrame(
        {
            "SKPD": clean_str(df_bmd_raw[col_skpd_bmd]),
            "No_Dokumen": clean_str(df_bmd_raw[col_doc_bmd]),
            "Rekening": clean_str(df_bmd_raw[col_rek_bmd]),
            "Nilai_BMD": clean_number(df_bmd_raw[col_val_bmd]),
        }
    )

    # Filter Otomatis Berdasarkan Prefix Kode Rekening LRA SIPD
    # Belanja Modal: diawali '5.2'
    sipd_modal = df_sipd_clean[df_sipd_clean["Rekening"].str.startswith("5.2")]
    # Belanja Persediaan: diawali '5.1.02.01' atau '5.1.2.01'
    sipd_persediaan = df_sipd_clean[
        df_sipd_clean["Rekening"].str.startswith("5.1.02.01")
        | df_sipd_clean["Rekening"].str.startswith("5.1.2.01")
    ]

    # Matching untuk masing-masing kategori
    # Catatan: Data BMD dicocokkan berdasarkan kesamaan nomor dokumen yang ada di masing-masing subset SIPD
    bmd_modal = df_bmd_clean[
        df_bmd_clean["No_Dokumen"].isin(sipd_modal["No_Dokumen"])
        | df_bmd_clean["Rekening"].str.startswith("1.3")  # Kode Akun Neraca Aset Tetap
    ]
    bmd_persediaan = df_bmd_clean[
        df_bmd_clean["No_Dokumen"].isin(sipd_persediaan["No_Dokumen"])
        | df_bmd_clean["Rekening"].str.startswith("1.1.7")  # Kode Akun Neraca Persediaan
    ]

    df_rekon_modal = process_reconciliation(sipd_modal, bmd_modal, "Belanja Modal")
    df_rekon_persediaan = process_reconciliation(
        sipd_persediaan, bmd_persediaan, "Belanja Persediaan"
    )

    # --- TAB NAVIGATION ---
    st.write("---")
    tab1, tab2, tab3 = st.tabs(
        [
            "🏢 Tab 1: Belanja Modal (Akun 5.2.x)",
            "📦 Tab 2: Belanja Persediaan (Akun 5.1.02.01.x)",
            "📥 Tab 3: Ekspor Laporan Rekon",
        ]
    )

    with tab1:
        render_tab_content(df_rekon_modal, "Belanja Modal")

    with tab2:
        render_tab_content(df_rekon_persediaan, "Belanja Persediaan")

    with tab3:
        st.subheader("Unduh Kertas Kerja Rekonsiliasi Gabungan")
        st.write(
            "File Excel ini berisi 2 *sheet terpisah* (Belanja Modal dan Belanja Persediaan) yang siap dijadikan lampiran Berita Acara Rekonsiliasi (BAR)."
        )

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            if df_rekon_modal is not None:
                df_rekon_modal.to_excel(
                    writer, index=False, sheet_name="Rekon_Belanja_Modal"
                )
            if df_rekon_persediaan is not None:
                df_rekon_persediaan.to_excel(
                    writer, index=False, sheet_name="Rekon_Persediaan"
                )

        st.download_button(
            label="📥 Unduh Kertas Kerja Rekonsiliasi (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="Kertas_Kerja_Rekon_Modal_dan_Persediaan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("💡 Silakan unggah kedua file di atas untuk memulai rekonsiliasi otomatis.")
