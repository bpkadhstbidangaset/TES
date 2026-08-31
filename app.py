import io
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Sistem Rekon BMD per SKPD - Bidang Aset",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🏛️ Rekonsiliasi Belanja Modal & Persediaan BMD per SKPD")
st.caption(
    "Bidang Pengelolaan Aset Daerah / BPKAD — Alat Bantu Verifikasi Realisasi"
    " SP2D vs Register KIB/Persediaan"
)


# ==========================================
# 1. FUNGSI PEMBERSIHAN & LOGIKA BMD
# ==========================================
def clean_string_column(series):
  """Membersihkan format nomor dokumen / nomor SKPD."""
  return (
      series.astype(str)
      .str.strip()
      .str.replace(r"\s+", " ", regex=True)
      .str.upper()
  )


def clean_currency(series):
  """Membersihkan format rupiah menjadi float numerik."""
  if pd.api.types.is_numeric_dtype(series):
    return series.fillna(0.0)
  return (
      series.astype(str)
      .str.replace("Rp", "", case=False, regex=False)
      .str.replace(".", "", regex=False)
      .str.replace(",", ".", regex=False)
      .str.strip()
      .apply(pd.to_numeric, errors="coerce")
      .fillna(0.0)
  )


def evaluate_bmd_row(row, toleransi_rp=100.0):
  kas = row["TOTAL_KAS"]
  aset = row["TOTAL_ASET"]
  selisih = kas - aset
  abs_selisih = abs(selisih)

  if abs_selisih <= toleransi_rp:
    return pd.Series(["Sesuai (Cocok)", "Nilai belanja modal/persediaan sinkron"])
  elif kas > 0 and aset == 0:
    return pd.Series([
        "Belum Dicatat di KIB/Persediaan",
        "SP2D terbit, belum diinput pengurus barang",
    ])
  elif kas == 0 and aset > 0:
    return pd.Series([
        "Aset Tercatat Tanpa SP2D",
        "Barang masuk KIB tanpa nomor SP2D/Realisasi Kasda",
    ])
  else:
    # Cek indikasi PPN 11%
    bruto = max(kas, aset)
    dpp = bruto / 1.11
    if abs(abs_selisih - (bruto - dpp)) <= (toleransi_rp * 5):
      return pd.Series([
          "Indikasi Selisih PPN 11%",
          "Kasda catat bruto, KIB catat DPP (atau sebaliknya)",
      ])
    # Cek indikasi PPh 22 1.5%
    if abs(abs_selisih - (dpp * 0.015)) <= (toleransi_rp * 5):
      return pd.Series([
          "Indikasi Selisih PPh 22 (1.5%)",
          "Selisih potongan PPh belanja barang/modal",
      ])
    return pd.Series(
        ["Tercatat tapi Selisih Nilai", "Perlu cek BAST fisik / rincian belanja"]
    )


# ==========================================
# 2. SIDEBAR UPLOAD & PEMETAAN KOLOM
# ==========================================
st.sidebar.header("📁 1. Unggah Data Berkas")
file_kas = st.sidebar.file_uploader(
    "Data SP2D Kasda (.xlsx / .csv)", type=["xlsx", "csv"]
)
file_aset = st.sidebar.file_uploader(
    "Data Register KIB/Persediaan (.xlsx / .csv)", type=["xlsx", "csv"]
)

if file_kas and file_aset:
  df_k_raw = (
      pd.read_csv(file_kas)
      if file_kas.name.endswith(".csv")
      else pd.read_excel(file_kas)
  )
  df_a_raw = (
      pd.read_csv(file_aset)
      if file_aset.name.endswith(".csv")
      else pd.read_excel(file_aset)
  )

  st.sidebar.header("🔗 2. Pemetaan Kolom SKPD & Nilai")

  col_skpd_k = st.sidebar.selectbox("Kolom Nama SKPD (File SP2D):", df_k_raw.columns)
  col_skpd_a = st.sidebar.selectbox("Kolom Nama SKPD (File Aset):", df_a_raw.columns)
  col_key_k = st.sidebar.selectbox(
      "Kolom No. SP2D / BAST (File SP2D):", df_k_raw.columns
  )
  col_key_a = st.sidebar.selectbox(
      "Kolom No. SP2D / BAST (File Aset):", df_a_raw.columns
  )
  col_val_k = st.sidebar.selectbox("Kolom Nominal SP2D:", df_k_raw.columns)
  col_val_a = st.sidebar.selectbox(
      "Kolom Nilai Perolehan Aset:", df_a_raw.columns
  )

  if st.sidebar.button("🚀 Jalankan Rekonsiliasi", type="primary"):
    df_k = df_k_raw.copy()
    df_a = df_a_raw.copy()

    df_k["SKPD_CLEAN"] = clean_string_column(df_k[col_skpd_k])
    df_a["SKPD_CLEAN"] = clean_string_column(df_a[col_skpd_a])
    df_k["KEY_CLEAN"] = clean_string_column(df_k[col_key_k])
    df_a["KEY_CLEAN"] = clean_string_column(df_a[col_key_a])
    df_k["VAL_CLEAN"] = clean_currency(df_k[col_val_k])
    df_a["VAL_CLEAN"] = clean_currency(df_a[col_val_a])

    # Agregasi per SKPD dan No Dokumen
    rekap_k = (
        df_k.groupby(["SKPD_CLEAN", "KEY_CLEAN"])["VAL_CLEAN"]
        .sum()
        .reset_index(name="TOTAL_KAS")
    )
    rekap_a = (
        df_a.groupby(["SKPD_CLEAN", "KEY_CLEAN"])["VAL_CLEAN"]
        .sum()
        .reset_index(name="TOTAL_ASET")
    )

    merged = pd.merge(
        rekap_k, rekap_a, on=["SKPD_CLEAN", "KEY_CLEAN"], how="outer"
    ).fillna(0.0)
    merged["SELISIH"] = merged["TOTAL_KAS"] - merged["TOTAL_ASET"]

    # Status awal dari sistem
    merged[["STATUS_SISTEM", "ANALISIS_SISTEM"]] = merged.apply(
        evaluate_bmd_row, axis=1
    )

    # Kolom Verifikasi Khas Bidang Aset BMD
    merged["SUDAH_VERIFIKASI"] = merged["STATUS_SISTEM"] == "Sesuai (Cocok)"
    merged["KLASIFIKASI_BMD"] = merged["STATUS_SISTEM"].apply(
        lambda s: "Sesuai / Valid" if s == "Sesuai (Cocok)" else "Pending / Belum Dicek"
    )
    merged["CATATAN_BIDANG_ASET"] = ""

    merged.rename(
        columns={"SKPD_CLEAN": "NAMA_SKPD", "KEY_CLEAN": "NO_DOKUMEN"},
        inplace=True,
    )
    st.session_state["df_master_rekon"] = merged
    st.session_state["raw_kas"] = df_k_raw
    st.session_state["raw_aset"] = df_a_raw
    st.success("Rekonsiliasi seluruh SKPD berhasil diproses!")

# Fallback Data Simulasi jika belum ada file yang diunggah
elif "df_master_rekon" not in st.session_state:
  dummy_data = {
      "NAMA_SKPD": [
          "DINAS KESEHATAN",
          "DINAS KESEHATAN",
          "DINAS PENDIDIKAN",
          "DINAS PENDIDIKAN",
          "DINAS PEKERJAAN UMUM",
      ],
      "NO_DOKUMEN": [
          "SP2D-001/DINKES",
          "SP2D-002/DINKES",
          "SP2D-010/DISDIK",
          "SP2D-011/DISDIK",
          "SP2D-099/DPUPR",
      ],
      "TOTAL_KAS": [
          50000000.0,
          111000000.0,
          25000000.0,
          8000000.0,
          750000000.0,
      ],
      "TOTAL_ASET": [
          50000000.0,
          100000000.0,
          25000000.0,
          0.0,
          0.0,
      ],
      "SELISIH": [0.0, 11000000.0, 0.0, 8000000.0, 750000000.0],
      "STATUS_SISTEM": [
          "Sesuai (Cocok)",
          "Indikasi Selisih PPN 11%",
          "Sesuai (Cocok)",
          "Belum Dicatat di KIB/Persediaan",
          "Belum Dicatat di KIB/Persediaan",
      ],
      "ANALISIS_SISTEM": [
          "Nilai belanja modal/persediaan sinkron",
          "Kasda catat bruto, KIB catat DPP",
          "Nilai belanja modal/persediaan sinkron",
          "SP2D terbit, belum diinput pengurus barang",
          "SP2D terbit, belum diinput pengurus barang",
      ],
      "SUDAH_VERIFIKASI": [True, False, True, False, False],
      "KLASIFIKASI_BMD": [
          "Sesuai / Valid",
          "Pending / Belum Dicek",
          "Sesuai / Valid",
          "Aset Ekstrakompatibel (Bawah Batas)",
          "Perlu Pencatatan KDP (KIB F)",
      ],
      "CATATAN_BIDANG_ASET": [
          "Sesuai",
          "",
          "Sesuai",
          "Pengadaan printer Rp 8 jt (Bawah batas kapitalisasi)",
          "Termin 1 Pembangunan Jembatan (Wajib KDP)",
      ],
  }
  st.session_state["df_master_rekon"] = pd.DataFrame(dummy_data)


# ==========================================
# 3. FILTER PER SKPD & DASHBOARD TAMPILAN
# ==========================================
if "df_master_rekon" in st.session_state:
  df_master = st.session_state["df_master_rekon"]

  # Daftar SKPD unik
  list_skpd = ["SEMUA SKPD"] + sorted(df_master["NAMA_SKPD"].unique().tolist())

  col_filter1, col_filter2 = st.columns([2, 1])
  with col_filter1:
    selected_skpd = st.selectbox(
        "🏢 Pilih SKPD / OPD yang akan Direkonsiliasi:", list_skpd
    )
  with col_filter2:
    mode_tampilan = st.radio(
        "Mode Tampilan:",
        ["Kertas Kerja Detail", "Rekapitulasi Global Per SKPD"],
        horizontal=True,
    )

  st.markdown("---")

  # --- MODE 1: REKAPITULASI GLOBAL PER SKPD ---
  if mode_tampilan == "Rekapitulasi Global Per SKPD":
    st.subheader("📊 Tabel Rekapitulasi Realisasi vs Mutasi Aset Seluruh SKPD")

    rekap_opd = (
        df_master.groupby("NAMA_SKPD")
        .agg(
            Total_SP2D=("TOTAL_KAS", "sum"),
            Total_Input_Aset=("TOTAL_ASET", "sum"),
            Total_Selisih=("SELISIH", "sum"),
            Jumlah_Item=("NO_DOKUMEN", "count"),
            Item_Selesai=("SUDAH_VERIFIKASI", "sum"),
        )
        .reset_index()
    )

    rekap_opd["Status_Rekon"] = rekap_opd.apply(
        lambda r: (
            "✅ Lengkap / Sinkron"
            if r["Total_Selisih"] == 0 and r["Item_Selesai"] == r["Jumlah_Item"]
            else "⚠️ Ada Selisih / Pending"
        ),
        axis=1,
    )

    # Format Tampilan Rupiah
    rekap_opd_display = rekap_opd.copy()
    for col in ["Total_SP2D", "Total_Input_Aset", "Total_Selisih"]:
      rekap_opd_display[col] = rekap_opd_display[col].apply(
          lambda x: f"Rp {x:,.0f}"
      )

    st.dataframe(rekap_opd_display, use_container_width=True, hide_index=True)

    # Visualisasi Bar Chart Selisih per SKPD
    fig = px.bar(
        rekap_opd,
        x="NAMA_SKPD",
        y=["Total_SP2D", "Total_Input_Aset"],
        barmode="group",
        title="Perbandingan Nilai SP2D Kasda vs Nilai Register Aset per OPD",
        labels={"value": "Rupiah", "variable": "Kategori"},
    )
    st.plotly_chart(fig, use_container_width=True)

  # --- MODE 2: KERTAS KERJA DETAIL PER SKPD ---
  else:
    if selected_skpd != "SEMUA SKPD":
      df_view = df_master[df_master["NAMA_SKPD"] == selected_skpd].copy()
    else:
      df_view = df_master.copy()

    # Ringkasan Metrik SKPD Terpilih
    c1, c2, c3, c4 = st.columns(4)
    tot_sp2d = df_view["TOTAL_KAS"].sum()
    tot_aset = df_view["TOTAL_ASET"].sum()
    tot_selisih = df_view["SELISIH"].sum()
    progres = (
        (df_view["SUDAH_VERIFIKASI"].sum() / len(df_view)) * 100
        if len(df_view) > 0
        else 0
    )

    c1.metric("Total Belanja Kasda (SP2D)", f"Rp {tot_sp2d:,.0f}")
    c2.metric("Total Dicatat di KIB/Persediaan", f"Rp {tot_aset:,.0f}")
    c3.metric(
        "Total Selisih",
        f"Rp {tot_selisih:,.0f}",
        delta=-tot_selisih if tot_selisih != 0 else 0,
    )
    c4.metric("Progres Verifikasi", f"{progres:.1f}%")

    st.markdown(f"### Kertas Kerja Rekonsiliasi: **{selected_skpd}**")

    # Konfigurasi Data Editor Khusus Bidang Aset
    col_config = {
        "NAMA_SKPD": st.column_config.TextColumn("Nama SKPD", disabled=True),
        "NO_DOKUMEN": st.column_config.TextColumn("No. SP2D", disabled=True),
        "TOTAL_KAS": st.column_config.NumberColumn(
            "SP2D Kasda (Rp)", format="Rp %d", disabled=True
        ),
        "TOTAL_ASET": st.column_config.NumberColumn(
            "Nilai KIB (Rp)", format="Rp %d", disabled=True
        ),
        "SELISIH": st.column_config.NumberColumn(
            "Selisih (Rp)", format="Rp %d", disabled=True
        ),
        "STATUS_SISTEM": st.column_config.TextColumn(
            "Deteksi Sistem", disabled=True
        ),
        "ANALISIS_SISTEM": st.column_config.TextColumn(
            "Analisis Indikasi", disabled=True
        ),
        "SUDAH_VERIFIKASI": st.column_config.CheckboxColumn(
            "Valid?",
            help="Centang jika status belanja ini sudah jelas/selesai",
            default=False,
        ),
        "KLASIFIKASI_BMD": st.column_config.SelectboxColumn(
            "Klasifikasi Bidang Aset",
            width="medium",
            options=[
                "Sesuai / Valid",
                "Pending / Belum Dicek",
                "Belum Input KIB (Intrakompatibel)",
                "Perlu Pencatatan KDP (KIB F)",
                "Aset Ekstrakompatibel (Bawah Batas)",
                "Persediaan Langsung Habis (Tanpa Gudang)",
                "Perlu Konfirmasi Potongan Pajak",
                "Koreksi Akun Belanja",
            ],
            required=True,
        ),
        "CATATAN_BIDANG_ASET": st.column_config.TextColumn(
            "Catatan / Rekomendasi Bidang Aset",
            width="large",
            max_chars=255,
            help="Tuliskan arahan ke pengurus barang OPD (misal: input ke KIB F atau koreksi BAST)",
        ),
    }

    # Data Editor Interaktif
    edited_view = st.data_editor(
        df_view,
        column_config=col_config,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=f"editor_{selected_skpd}",
    )

    # Sinkronisasi kembali hasil edit ke session_state utama
    if selected_skpd != "SEMUA SKPD":
      df_master.update(edited_view)
    else:
      df_master = edited_view.copy()
    st.session_state["df_master_rekon"] = df_master

    # ==========================================
    # 4. EKSPOR BERITA ACARA REKON PER SKPD
    # ==========================================
    st.markdown("---")
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
      edited_view.to_excel(
          writer,
          sheet_name=f"Rekon_{selected_skpd[:25]}",
          index=False,
      )
      df_master.to_excel(writer, sheet_name="Rekap Seluruh SKPD", index=False)

    nama_file_excel = f"BAR_Rekon_Aset_{selected_skpd.replace(' ', '_')}.xlsx"
    st.download_button(
        label=f"📥 Unduh Kertas Kerja BAR ({selected_skpd})",
        data=output_buffer.getvalue(),
        file_name=nama_file_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
