import io
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Sistem Rekonsiliasi Belanja Modal & Aset BMD",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🏛️ Aplikasi Rekonsiliasi Belanja Modal & Persediaan")
st.caption(
    "Bidang Pengelolaan Aset Daerah / BPKAD — Alat Bantu Pencocokan Realisasi"
    " SP2D Kasda vs Register Aset/Persediaan (SIPD-RI)"
)


# ==========================================
# 1. FUNGSI STANDARISASI & ANALISIS
# ==========================================
def clean_string_column(series):
  """Membersihkan format nomor dokumen / kontrak."""
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


def check_tax_discrepancy(nilai_kas, nilai_aset, tax_config):
  """Mendeteksi apakah selisih nilai dipicu oleh komponen pajak belanja."""
  selisih = nilai_kas - nilai_aset
  abs_selisih = abs(selisih)
  toleransi_rp = tax_config.get("toleransi", 100.0)

  if abs_selisih <= toleransi_rp:
    return "Cocok (Sesuai)", 0.0, "Toleransi pembulatan nilai"

  bruto = max(nilai_kas, nilai_aset)
  ppn = tax_config.get("PPN")
  dpp = bruto / (1.0 + ppn) if ppn and ppn > 0 else bruto

  tax_rules = []
  if ppn and ppn > 0:
    tax_rules.append(
        (f"Selisih PPN {ppn*100:.1f}% (DPP vs Bruto)", bruto - dpp)
    )

  for pph_name in ["PPh 22", "PPh 23", "PPh Final"]:
    rate = tax_config.get(pph_name)
    if rate and rate > 0:
      tax_rules.append((f"Selisih {pph_name} ({rate*100:.2f}%)", dpp * rate))

  if ppn and ppn > 0:
    for pph_name in ["PPh 22", "PPh 23"]:
      rate = tax_config.get(pph_name)
      if rate and rate > 0:
        tax_rules.append((
            f"Selisih PPN + {pph_name}",
            (bruto - dpp) + (dpp * rate),
        ))

  for label_aturan, nominal_estimasi in tax_rules:
    if abs(abs_selisih - nominal_estimasi) <= (toleransi_rp * 5):
      keterangan = (
          "Kas catat Bruto / Aset catat DPP/Neto"
          if selisih > 0
          else "Aset catat Bruto / Kas catat Neto"
      )
      return f"Indikasi {label_aturan}", nominal_estimasi, keterangan

  return "Tercatat tapi Selisih Nominal", 0.0, "Perlu Cek BAST / Bukti Fisik"


def evaluate_row(row, tax_config):
  kas = row["TOTAL_KAS"]
  aset = row["TOTAL_ASET"]

  if kas > 0 and aset == 0:
    return pd.Series([
        "Belum Dicatat di Aset/Persediaan",
        0.0,
        "SP2D cair, BAST/KIB belum diinput",
    ])
  elif kas == 0 and aset > 0:
    return pd.Series([
        "Aset Tercatat Tanpa Realisasi Kas",
        0.0,
        "Barang masuk tapi SP2D belum terbit",
    ])
  else:
    status, pot, ket = check_tax_discrepancy(kas, aset, tax_config)
    return pd.Series([status, pot, ket])


# ==========================================
# 2. SIDEBAR PARAMETER
# ==========================================
st.sidebar.header("📁 1. Unggah Berkas")
file_kas = st.sidebar.file_uploader(
    "Data SP2D / BKU Belanja (.xlsx / .csv)", type=["xlsx", "csv"]
)
file_aset = st.sidebar.file_uploader(
    "Data Register Aset / KIB / Persediaan (.xlsx / .csv)", type=["xlsx", "csv"]
)

st.sidebar.header("⚙️ 2. Konfigurasi Analisis Pajak")
with st.sidebar.expander("Pengaturan Tarif Pajak", expanded=False):
  toleransi_val = st.number_input(
      "Toleransi Pembulatan (Rp):", min_value=0.0, value=100.0, step=50.0
  )
  aktif_ppn = st.checkbox("Evaluasi PPN (11%)", value=True)
  tarif_ppn = (
      st.number_input("Tarif PPN (%):", value=11.0, step=0.5) / 100.0
      if aktif_ppn
      else None
  )

  aktif_pph22 = st.checkbox("Evaluasi PPh 22 Barang (1.5%)", value=True)
  tarif_pph22 = (
      st.number_input("Tarif PPh 22 (%):", value=1.5, step=0.1) / 100.0
      if aktif_pph22
      else None
  )

  aktif_pph23 = st.checkbox("Evaluasi PPh 23 Jasa (2%)", value=True)
  tarif_pph23 = (
      st.number_input("Tarif PPh 23 (%):", value=2.0, step=0.1) / 100.0
      if aktif_pph23
      else None
  )

  aktif_pph_final = st.checkbox("Evaluasi PPh Final Konstruksi", value=False)
  tarif_pph_final = (
      st.number_input("Tarif PPh Final (%):", value=1.75, step=0.05) / 100.0
      if aktif_pph_final
      else None
  )

  tax_config = {
      "PPN": tarif_ppn,
      "PPh 22": tarif_pph22,
      "PPh 23": tarif_pph23,
      "PPh Final": tarif_pph_final,
      "toleransi": toleransi_val,
  }

# ==========================================
# 3. PROSES REKONSILIASI
# ==========================================
if file_kas and file_aset:
  df_kas_raw = (
      pd.read_csv(file_kas)
      if file_kas.name.endswith(".csv")
      else pd.read_excel(file_kas)
  )
  df_aset_raw = (
      pd.read_csv(file_aset)
      if file_aset.name.endswith(".csv")
      else pd.read_excel(file_aset)
  )

  st.sidebar.header("🔗 3. Pemetaan Kolom")
  col_key_kas = st.sidebar.selectbox(
      "Kolom Kunci SP2D (No. SP2D / BAST):", df_kas_raw.columns
  )
  col_key_aset = st.sidebar.selectbox(
      "Kolom Kunci Aset (No. SP2D / BAST):", df_aset_raw.columns
  )
  col_val_kas = st.sidebar.selectbox(
      "Kolom Nominal Realisasi SP2D:", df_kas_raw.columns
  )
  col_val_aset = st.sidebar.selectbox(
      "Kolom Nilai Perolehan Aset:", df_aset_raw.columns
  )

  if st.sidebar.button("🚀 Jalankan Rekonsiliasi", type="primary"):
    df_k = df_kas_raw.copy()
    df_a = df_aset_raw.copy()

    df_k["_KEY"] = clean_string_column(df_k[col_key_kas])
    df_a["_KEY"] = clean_string_column(df_a[col_key_aset])
    df_k["_VAL"] = clean_currency(df_k[col_val_kas])
    df_a["_VAL"] = clean_currency(df_a[col_val_aset])

    # Agregasi
    rekap_k = (
        df_k.groupby("_KEY")["_VAL"].sum().reset_index(name="TOTAL_KAS")
    )
    rekap_a = (
        df_a.groupby("_KEY")["_VAL"].sum().reset_index(name="TOTAL_ASET")
    )

    merged = pd.merge(rekap_k, rekap_a, on="_KEY", how="outer").fillna(0.0)
    merged["SELISIH"] = merged["TOTAL_KAS"] - merged["TOTAL_ASET"]

    # Evaluasi Status & Pajak
    merged[["STATUS_SISTEM", "POTONGAN_PAJAK", "KETERANGAN_ANALISIS"]] = (
        merged.apply(lambda r: evaluate_row(r, tax_config), axis=1)
    )

    # Inisialisasi kolom verifikasi
    merged["SUDAH_DITINDAKLANJUTI"] = merged["STATUS_SISTEM"] == "Cocok (Sesuai)"
    merged["STATUS_VERIFIKASI"] = merged["STATUS_SISTEM"].apply(
        lambda s: "Selesai / Clear" if s == "Cocok (Sesuai)" else "Pending"
    )
    merged["CATATAN_VERIFIKATOR"] = ""
    merged["PIC_TINDAK_LANJUT"] = "-"

    merged.rename(columns={"_KEY": "NO_DOKUMEN"}, inplace=True)
    st.session_state["df_rekon_audit"] = merged
    st.session_state["raw_kas"] = df_kas_raw
    st.session_state["raw_aset"] = df_aset_raw
    st.success("Rekonsiliasi berhasil diproses!")

# Fallback ke Dummy Data jika belum upload file
elif "df_rekon_audit" not in st.session_state:
  st.info(
      "💡 Menampilkan data simulasi awal. Silakan unggah file SP2D dan Aset"
      " pada panel samping untuk memproses data riil."
  )
  dummy_data = {
      "NO_DOKUMEN": [
          "SP2D-0012/MODAL",
          "SP2D-0045/PERSED",
          "SP2D-0089/MODAL",
          "SP2D-0099/MODAL",
      ],
      "TOTAL_KAS": [111000000.0, 22200000.0, 50000000.0, 15000000.0],
      "TOTAL_ASET": [100000000.0, 21900000.0, 47500000.0, 15000000.0],
      "SELISIH": [11000000.0, 300000.0, 2500000.0, 0.0],
      "STATUS_SISTEM": [
          "Indikasi Selisih PPN 11% (DPP vs Bruto)",
          "Indikasi Selisih PPh 22 (1.5%)",
          "Tercatat tapi Selisih Nominal",
          "Cocok (Sesuai)",
      ],
      "SUDAH_DITINDAKLANJUTI": [False, False, False, True],
      "STATUS_VERIFIKASI": [
          "Perlu Konfirmasi Pajak",
          "Perlu Konfirmasi Pajak",
          "Pending",
          "Selesai / Clear",
      ],
      "CATATAN_VERIFIKATOR": [
          "",
          "",
          "Kuitansi fisik belum diterima dari PPK",
          "Sesuai BAST",
      ],
      "PIC_TINDAK_LANJUT": ["Bendahara Pengeluaran", "PPK", "Pengurus Barang", "-"],
  }
  st.session_state["df_rekon_audit"] = pd.DataFrame(dummy_data)

# ==========================================
# 4. DASHBOARD & DATA EDITOR INTERAKTIF
# ==========================================
if "df_rekon_audit" in st.session_state:
  df_current = st.session_state["df_rekon_audit"]

  # Metric Ringkasan
  col1, col2, col3, col4 = st.columns(4)
  total_item = len(df_current)
  selesai = df_current["SUDAH_DITINDAKLANJUTI"].sum()
  total_selisih_nom = (
      (df_current["STATUS_SISTEM"] == "Tercatat tapi Selisih Nominal")
      .astype(int)
      .sum()
  )
  selisih_pajak = df_current["STATUS_SISTEM"].str.contains("Indikasi").sum()

  col1.metric("Total Transaksi", f"{total_item} Item")
  col2.metric("Selesai Ditindaklanjuti", f"{selesai} Item")
  col3.metric("Indikasi Pajak", f"{selisih_pajak} Item")
  col4.metric("Perlu Koreksi Fisik", f"{total_selisih_nom} Item")

  # Visualisasi Status
  fig = px.pie(
      df_current,
      names="STATUS_SISTEM",
      title="Proporsi Hasil Rekonsiliasi",
      hole=0.4,
      color_discrete_sequence=px.colors.qualitative.Safe,
  )
  fig.update_layout(height=280, margin=dict(l=20, r=20, t=40, b=20))
  st.plotly_chart(fig, use_container_width=True)

  st.markdown("### Form Kertas Kerja & Tindak Lanjut Verifikator")

  column_configuration = {
      "NO_DOKUMEN": st.column_config.TextColumn("No. Dokumen", disabled=True),
      "TOTAL_KAS": st.column_config.NumberColumn(
          "Realisasi SP2D (Rp)", format="Rp %d", disabled=True
      ),
      "TOTAL_ASET": st.column_config.NumberColumn(
          "Nilai Aset (Rp)", format="Rp %d", disabled=True
      ),
      "SELISIH": st.column_config.NumberColumn(
          "Selisih (Rp)", format="Rp %d", disabled=True
      ),
      "STATUS_SISTEM": st.column_config.TextColumn(
          "Hasil Deteksi Sistem", disabled=True
      ),
      "SUDAH_DITINDAKLANJUTI": st.column_config.CheckboxColumn(
          "Selesai?", default=False
      ),
      "STATUS_VERIFIKASI": st.column_config.SelectboxColumn(
          "Status Verifikasi",
          width="medium",
          options=[
              "Pending",
              "Perlu Konfirmasi Pajak",
              "Perlu Koreksi Jurnal",
              "Menunggu BAST Fisik",
              "Selesai / Clear",
              "Diabaikan (Wajar)",
          ],
          required=True,
      ),
      "CATATAN_VERIFIKATOR": st.column_config.TextColumn(
          "Catatan Tindak Lanjut / Temuan", width="large", max_chars=255
      ),
      "PIC_TINDAK_LANJUT": st.column_config.SelectboxColumn(
          "PIC Terkait",
          options=[
              "-",
              "PPK",
              "Bendahara Pengeluaran",
              "Pengurus Barang / Pengurus BMD",
              "Bidang Akuntansi",
              "Penyedia",
          ],
          required=True,
      ),
  }

  edited_df = st.data_editor(
      df_current,
      column_config=column_configuration,
      use_container_width=True,
      hide_index=True,
      num_rows="fixed",
      key="data_editor_rekon_table",
  )

  # Simpan perubahan
  st.session_state["df_rekon_audit"] = edited_df

  # ==========================================
  # 5. EKSPOR HASIL VERIFIKASI KE EXCEL
  # ==========================================
  output_buffer = io.BytesIO()
  with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
    edited_df.to_excel(
        writer, sheet_name="Berita Acara Rekon", index=False
    )
    if "raw_kas" in st.session_state:
      st.session_state["raw_kas"].to_excel(
          writer, sheet_name="Data SP2D Kasda", index=False
      )
    if "raw_aset" in st.session_state:
      st.session_state["raw_aset"].to_excel(
          writer, sheet_name="Data Register Aset", index=False
      )

  st.download_button(
      label="📥 Unduh Kertas Kerja Berita Acara Rekon (.xlsx)",
      data=output_buffer.getvalue(),
      file_name="Berita_Acara_Rekonsiliasi_BMD.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
