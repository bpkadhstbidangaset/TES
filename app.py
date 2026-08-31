import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Rekonsiliasi Belanja Modal & Persediaan", layout="wide")

st.title("📊 Rekonsiliasi Rekening Belanja Modal & Persediaan")
st.write("Aplikasi untuk mencocokkan data transaksi Belanja Modal/Persediaan dengan Buku Persediaan/Aset.")

# 1. Upload File
col1, col2 = st.columns(2)

with col1:
    st.subheader("Data Realisasi Belanja (SPM / SP2D / LRA)")
    file_belanja = st.file_uploader("Upload File Belanja (Excel/CSV)", type=["xlsx", "xls", "csv"], key="belanja")

with col2:
    st.subheader("Data Buku Persediaan / Barang (BAST / SIMAK)")
    file_persediaan = st.file_uploader("Upload File Persediaan (Excel/CSV)", type=["xlsx", "xls", "csv"], key="persediaan")

def load_data(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)

if file_belanja and file_persediaan:
    df_belanja = load_data(file_belanja)
    df_persediaan = load_data(file_persediaan)

    st.divider()
    st.subheader("⚙️ Konfigurasi Pencocokan Data")

    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        key_belanja = st.selectbox("Pilih Kolom Kunci Data Belanja (misal: No SP2D / Kode Barang):", df_belanja.columns)
        val_belanja = st.selectbox("Pilih Kolom Nominal Belanja:", df_belanja.columns)

    with col_cfg2:
        key_persediaan = st.selectbox("Pilih Kolom Kunci Data Persediaan:", df_persediaan.columns)
        val_persediaan = st.selectbox("Pilih Kolom Nominal Persediaan:", df_persediaan.columns)

    if st.button("🚀 Jalankan Rekonsiliasi", type="primary"):
        # Standardisasi format kunci & nilai
        df_belanja[key_belanja] = df_belanja[key_belanja].astype(str).str.strip()
        df_persediaan[key_persediaan] = df_persediaan[key_persediaan].astype(str).str.strip()

        df_belanja[val_belanja] = pd.to_numeric(df_belanja[val_belanja], errors="coerce").fillna(0)
        df_persediaan[val_persediaan] = pd.to_numeric(df_persediaan[val_persediaan], errors="coerce").fillna(0)

        # Agregasi data per kunci jika ada transaksi berulang
        agg_belanja = df_belanja.groupby(key_belanja, as_index=False)[val_belanja].sum()
        agg_persediaan = df_persediaan.groupby(key_persediaan, as_index=False)[val_persediaan].sum()

        # Merge Outer untuk melihat Cocok, Kurang Catat di Belanja, atau Kurang Catat di Persediaan
        merged = pd.merge(
            agg_belanja,
            agg_persediaan,
            left_on=key_belanja,
            right_on=key_persediaan,
            how="outer"
        )

        # Isi ID yang hilang karena outer join
        merged["ID_Kunci"] = merged[key_belanja].combine_first(merged[key_persediaan])
        merged[val_belanja] = merged[val_belanja].fillna(0)
        merged[val_persediaan] = merged[val_persediaan].fillna(0)

        # Hitung Selisih
        merged["Selisih"] = merged[val_belanja] - merged[val_persediaan]
        merged["Status"] = merged["Selisih"].apply(
            lambda x: "Cocok" if x == 0 else ("Selisih Lebih Belanja" if x > 0 else "Selisih Lebih Persediaan")
        )

        # Susun kolom akhir
        result_df = merged[["ID_Kunci", val_belanja, val_persediaan, "Selisih", "Status"]].rename(
            columns={val_belanja: "Nominal_Belanja", val_persediaan: "Nominal_Persediaan"}
        )

        # Ringkasan KPI
        st.divider()
        st.subheader("📈 Hasil Rekonsiliasi")
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Belanja", f"Rp {result_df['Nominal_Belanja'].sum():,.2f}")
        kpi2.metric("Total Persediaan", f"Rp {result_df['Nominal_Persediaan'].sum():,.2f}")
        kpi3.metric("Total Selisih", f"Rp {result_df['Selisih'].sum():,.2f}")
        kpi4.metric("Item Tidak Cocok", len(result_df[result_df["Status"] != "Cocok"]))

        # Filter tampilan
        status_filter = st.multiselect(
            "Filter Status Transaksi:",
            options=result_df["Status"].unique(),
            default=result_df["Status"].unique()
        )
        filtered_df = result_df[result_df["Status"].isin(status_filter)]

        st.dataframe(filtered_df, use_container_width=True)

        # Export Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Hasil_Rekonsiliasi")
        
        st.download_button(
            label="📥 Download Hasil Rekonsiliasi (.xlsx)",
            data=output.getvalue(),
            file_name="hasil_rekonsiliasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
