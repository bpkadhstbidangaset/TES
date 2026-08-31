import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

st.set_page_config(page_title="Rekonsiliasi Belanja Modal & Persediaan", layout="wide")

st.title("📊 Rekonsiliasi Rekening Belanja Modal & Persediaan")
st.caption("Aplikasi pencocokan data transaksi belanja modal/persediaan dengan buku persediaan & aset.")

# 1. Bagian Upload Data
col1, col2 = st.columns(2)

with col1:
    st.subheader("Data Belanja (SPM / SP2D / LRA)")
    file_belanja = st.file_uploader("Upload File Belanja (Excel/CSV)", type=["xlsx", "xls", "csv"], key="belanja")

with col2:
    st.subheader("Data Persediaan / Barang (BAST / SIMAK)")
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
    st.subheader("⚙️ Konfigurasi Kolom Pencocokan")

    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        key_belanja = st.selectbox("Kolom Kunci Belanja (Kode Akun / No SPM / Kode Barang):", df_belanja.columns)
        val_belanja = st.selectbox("Kolom Nominal Belanja:", df_belanja.columns)

    with col_cfg2:
        key_persediaan = st.selectbox("Kolom Kunci Persediaan:", df_persediaan.columns)
        val_persediaan = st.selectbox("Kolom Nominal Persediaan:", df_persediaan.columns)

    if st.button("🚀 Jalankan Rekonsiliasi & Analisis", type="primary"):
        # Standardisasi data string & numeric
        df_belanja[key_belanja] = df_belanja[key_belanja].astype(str).str.strip()
        df_persediaan[key_persediaan] = df_persediaan[key_persediaan].astype(str).str.strip()

        df_belanja[val_belanja] = pd.to_numeric(df_belanja[val_belanja], errors="coerce").fillna(0)
        df_persediaan[val_persediaan] = pd.to_numeric(df_persediaan[val_persediaan], errors="coerce").fillna(0)

        # Agregasi data jika terdapat baris berulang
        agg_belanja = df_belanja.groupby(key_belanja, as_index=False)[val_belanja].sum()
        agg_persediaan = df_persediaan.groupby(key_persediaan, as_index=False)[val_persediaan].sum()

        # Full Outer Join
        merged = pd.merge(
            agg_belanja,
            agg_persediaan,
            left_on=key_belanja,
            right_on=key_persediaan,
            how="outer"
        )

        merged["ID_Kunci"] = merged[key_belanja].combine_first(merged[key_persediaan])
        merged[val_belanja] = merged[val_belanja].fillna(0)
        merged[val_persediaan] = merged[val_persediaan].fillna(0)

        # Hitung Nilai Selisih dan Kategori
        merged["Selisih"] = merged[val_belanja] - merged[val_persediaan]
        merged["Selisih_Absolut"] = merged["Selisih"].abs()
        
        merged["Status"] = merged["Selisih"].apply(
            lambda x: "Cocok" if abs(x) < 0.01 else ("Lebih Catat Belanja" if x > 0 else "Lebih Catat Persediaan")
        )

        result_df = merged[["ID_Kunci", val_belanja, val_persediaan, "Selisih", "Selisih_Absolut", "Status"]].rename(
            columns={val_belanja: "Nominal_Belanja", val_persediaan: "Nominal_Persediaan"}
        )

        # Perhitungan Metrik & Persentase
        total_items = len(result_df)
        total_matched = len(result_df[result_df["Status"] == "Cocok"])
        total_unmatched = total_items - total_matched
        
        match_rate_item = (total_matched / total_items * 100) if total_items > 0 else 0
        
        total_val_belanja = result_df["Nominal_Belanja"].sum()
        total_val_persediaan = result_df["Nominal_Persediaan"].sum()
        total_selisih = result_df["Selisih"].sum()
        total_selisih_abs = result_df["Selisih_Absolut"].sum()

        st.divider()
        st.subheader("📈 Ringkasan & Tingkat Kesesuaian (Match Rate)")

        # Baris Metrik Utama
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Belanja", f"Rp {total_val_belanja:,.2f}")
        kpi2.metric("Total Persediaan", f"Rp {total_val_persediaan:,.2f}")
        kpi3.metric("Total Selisih Netto", f"Rp {total_selisih:,.2f}")
        kpi4.metric("Kesesuaian Data (Match Rate)", f"{match_rate_item:.1f}%")

        # Indikator Progress Bar
        st.write(f"**Tingkat Ketercapaian Rekonsiliasi Item ({total_matched} dari {total_items} item cocok):**")
        st.progress(min(max(match_rate_item / 100, 0.0), 1.0))

        st.divider()
        st.subheader("📊 Visualisasi Breakdown Data")

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            # 1. Donut Chart Komposisi Status Rekonsiliasi
            status_summary = result_df["Status"].value_counts().reset_index()
            status_summary.columns = ["Status", "Jumlah"]

            color_map = {
                "Cocok": "#2ca02c",
                "Lebih Catat Belanja": "#ff7f0e",
                "Lebih Catat Persediaan": "#d62728"
            }

            fig_donut = px.pie(
                status_summary,
                names="Status",
                values="Jumlah",
                hole=0.45,
                title="Proporsi Status Rekonsiliasi (Berdasarkan Item)",
                color="Status",
                color_discrete_map=color_map
            )
            fig_donut.update_traces(textinfo="percent+label")
            st.plotly_chart(fig_donut, use_container_width=True)

        with chart_col2:
            # 2. Bar Chart Top 10 Selisih Terbesar
            unmatched_df = result_df[result_df["Status"] != "Cocok"].sort_values(
                by="Selisih_Absolut", ascending=False
            ).head(10)

            if not unmatched_df.empty:
                fig_bar = px.bar(
                    unmatched_df,
                    x="ID_Kunci",
                    y="Selisih",
                    color="Status",
                    title="Top 10 Item/Akun dengan Selisih Terbesar (Netto)",
                    color_discrete_map=color_map,
                    labels={"ID_Kunci": "Kunci / Akun", "Selisih": "Nominal Selisih (Rp)"}
                )
                fig_bar.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.success("🎉 Tidak ada selisih ditemukan. Seluruh data cocok 100%!")

        # Tabel Rincian Data
        st.divider()
        st.subheader("📋 Rincian Data Rekonsiliasi")

        status_filter = st.multiselect(
            "Filter Tampilan Berdasarkan Status:",
            options=result_df["Status"].unique(),
            default=result_df["Status"].unique()
        )
        
        filtered_df = result_df[result_df["Status"].isin(status_filter)].drop(columns=["Selisih_Absolut"])
        st.dataframe(filtered_df, use_container_width=True)

        # Fitur Download Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Rekonsiliasi")
            
            # Buat sheet ringkasan
            summary_export = pd.DataFrame({
                "Indikator": ["Total Belanja", "Total Persediaan", "Selisih Netto", "Total Selisih Absolut", "Total Item", "Item Cocok", "Persentase Cocok (%)"],
                "Nilai": [total_val_belanja, total_val_persediaan, total_selisih, total_selisih_abs, total_items, total_matched, match_rate_item]
            })
            summary_export.to_excel(writer, index=False, sheet_name="Ringkasan")

        st.download_button(
            label="📥 Download Laporan Lengkap (.xlsx)",
            data=output.getvalue(),
            file_name="laporan_rekonsiliasi_lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
