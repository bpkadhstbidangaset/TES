import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Kertas Kerja Verifikasi Rekon", layout="wide")

st.title("📝 Kertas Kerja Verifikasi & Tindak Lanjut Rekonsiliasi")
st.caption(
    "Verifikator dapat menandai status penyelesaian, memilih rekomendasi, dan"
    " menuliskan catatan tindak lanjut langsung pada tabel."
)

# ---------------------------------------------------------
# 1. INISIALISASI SESSION STATE (Menyimpan Data Edit)
# ---------------------------------------------------------
# Contoh data hasil rekon (jika belum ada di session_state)
if "df_rekon_audit" not in st.session_state:
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
          "Indikasi Selisih PPN 11%",
          "Indikasi Selisih PPh 22 (1.5%)",
          "Tercatat tapi Selisih Nominal",
          "Cocok (Sesuai)",
      ],
      # Kolom tambahan khusus untuk verifikator:
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

# ---------------------------------------------------------
# 2. KONFIGURASI KOLOM INTERAKTIF (st.column_config)
# ---------------------------------------------------------
column_configuration = {
    # Kolom Read-Only (Hanya Baca / Hasil Engine)
    "NO_DOKUMEN": st.column_config.TextColumn("No. Dokumen", disabled=True),
    "TOTAL_KAS": st.column_config.NumberColumn(
        "Realisasi Kas (Rp)", format="Rp %d", disabled=True
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
    # Kolom Interaktif yang Dapat Diedit Verifikator
    "SUDAH_DITINDAKLANJUTI": st.column_config.CheckboxColumn(
        "Selesai?",
        help="Centang jika selisih sudah selesai ditindaklanjuti/diklarifikasi",
        default=False,
    ),
    "STATUS_VERIFIKASI": st.column_config.SelectboxColumn(
        "Status Verifikator",
        help="Pilih status validasi manual",
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
        "Catatan Tindak Lanjut / Temuan",
        help="Ketik rincian tindak lanjut atau penjelasan selisih",
        width="large",
        max_chars=255,
    ),
    "PIC_TINDAK_LANJUT": st.column_config.SelectboxColumn(
        "PIC / Unit Terkait",
        options=[
            "-",
            "PPK",
            "Bendahara Pengeluaran",
            "Pengurus Barang / Logistik",
            "Akuntansi / GLP",
            "Penyedia / Pihak Ketiga",
        ],
        required=True,
    ),
}

# ---------------------------------------------------------
# 3. MENAMPILKAN DATA EDITOR
# ---------------------------------------------------------
st.markdown("### Daftar Transaksi & Form Verifikasi")

# Data editor interaktif
edited_df = st.data_editor(
    st.session_state["df_rekon_audit"],
    column_config=column_configuration,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",  # Mengunci agar user tidak menghapus/menambah baris data mentah
    key="tabel_editor_rekon",
)

# Simpan perubahan ke session_state saat ada interaksi edit
st.session_state["df_rekon_audit"] = edited_df

# ---------------------------------------------------------
# 4. STATISTIK PROGRESS VERIFIKASI & EKSPOR
# ---------------------------------------------------------
st.markdown("---")
col1, col2, col3 = st.columns(3)

total_item = len(edited_df)
total_selesai = edited_df["SUDAH_DITINDAKLANJUTI"].sum()
total_pending = total_item - total_selesai

col1.metric("Total Transaksi", f"{total_item} Item")
col2.metric("Selesai Ditindaklanjuti", f"{total_selesai} Item")
col3.metric("Belum Selesai / Pending", f"{total_pending} Item")

# Ekspor Kertas Kerja Hasil Verifikasi
output_buffer = io.BytesIO()
with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
  edited_df.to_excel(
      writer, sheet_name="Kertas Kerja Verifikasi", index=False
  )

st.download_button(
    label="📥 Unduh Kertas Kerja Berita Acara Rekon (.xlsx)",
    data=output_buffer.getvalue(),
    file_name="Berita_Acara_Hasil_Verifikasi_Rekon.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
