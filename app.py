import html
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# KONFIGURASI
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Rekon Persediaan & Belanja Modal",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).parent
FILE_RAK_PERSEDIAAN = BASE_DIR / "RAK PERSEDIAAN.xlsx"
FILE_RAK_MODAL = BASE_DIR / "RAK BELANJA MODAL.xlsx"
SHEET_LRA = "Data Realisasi Dokumen"

# Hanya kolom ini yang dibaca dari file LRA (file asli punya 46 kolom).
KOLOM_LRA = [
    'BULAN',
    'Nama SKPD',
    'Kode Sub Kegiatan', 'Nama Sub Kegiatan',
    'Kode Rekening', 'Nama Rekening',
    'Nomor Dokumen', 'Tanggal Dokumen', 'Keterangan Dokumen',
    'Nilai Realisasi',
]
KOLOM_WAJIB = ['Nama SKPD', 'Kode Rekening', 'Nilai Realisasi']
KOLOM_TEKS = [c for c in KOLOM_LRA if c != 'Nilai Realisasi']

NAMA_BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
              'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
# Cadangan jika kode bulan tidak diawali nomor: dicocokkan dari 3 huruf pertama
NAMA_BULAN_PREFIX = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'MEI': 4, 'JUN': 5, 'JUL': 6, 'AGU': 7, 'AUG': 7,
                     'SEP': 8, 'OKT': 9, 'OCT': 9, 'NOV': 10, 'DES': 11, 'DEC': 11}

MAX_BARIS_RINCI = 1000  # batas baris tabel rincian yang digambar sebagai HTML

# Engine Excel: calamine jauh lebih cepat dari openpyxl untuk file besar.
# Kalau paketnya belum terpasang, otomatis pakai engine bawaan pandas.
try:
    import python_calamine  # noqa: F401
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = None

# Custom CSS Kunci Ukuran 70% (Compact & Pas 1 Layar)
CSS = """<style>
    .main-header { font-size: 24px; font-weight: 700; color: #1E3A8A; margin-bottom: 2px; }
    .sub-header { font-size: 13px; color: #6B7280; margin-bottom: 18px; }
    
    .metric-box {
        color: white; padding: 14px 18px; border-radius: 10px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 12px;
    }
    .metric-title { font-size: 12px; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; opacity: 0.9; }
    .metric-value { font-size: 20px; font-weight: 700; margin-top: 3px; }
    
    .status-card-match {
        background-color: #ECFDF5; border-left: 4px solid #10B981;
        padding: 12px 16px; border-radius: 8px; margin-top: 12px; margin-bottom: 18px;
    }
    .status-card-diff {
        background-color: #FEF2F2; border-left: 4px solid #EF4444;
        padding: 12px 16px; border-radius: 8px; margin-top: 12px; margin-bottom: 18px;
    }
    
    /* Container Box / Bubble Table Mandiri (Skala 70%) */
    .table-container {
        background-color: #FFFFFF;
        border-radius: 10px;
        border: 1px solid #CBD5E1;
        box-shadow: 0 3px 5px -1px rgba(0, 0, 0, 0.05);
        margin-top: 8px;
        margin-bottom: 20px;
        max-height: 460px;
        overflow-y: auto;
        overflow-x: auto;
        position: relative;
    }
    
    /* Table Styling 70% Compact */
    table.custom-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 11px; /* Ukuran teks 70% */
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    table.custom-table th {
        background-color: #F1F5F9;
        color: #1E293B;
        font-weight: 700;
        text-align: left;
        padding: 7px 8px;
        border-bottom: 2px solid #94A3B8;
        white-space: nowrap;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    table.custom-table td {
        padding: 6px 8px;
        border-bottom: 1px solid #F1F5F9;
        color: #334155;
        vertical-align: middle;
    }
    table.custom-table tr:hover td {
        background-color: #F8FAFC;
    }
    table.custom-table tr.total-row td {
        background-color: #E2E8F0 !important;
        font-weight: 800 !important;
        color: #0F172A !important;
        border-top: 2px solid #94A3B8 !important;
        border-bottom: 2px solid #94A3B8 !important;
        position: sticky;
        bottom: 0;
        z-index: 9;
    }
    
    .text-right { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
    .text-center { text-align: center; white-space: nowrap; }
    .text-code { white-space: nowrap; font-family: monospace; font-size: 10.5px; color: #475569; }
    
    .diff-match { color: #059669 !important; font-weight: 600; }
    .diff-alert { color: #DC2626 !important; font-weight: 700; background-color: #FEF2F2; border-radius: 4px; padding: 1px 4px; }
    
    .badge-pill-match {
        background-color: #D1FAE5; color: #065F46; padding: 2px 7px; border-radius: 9999px;
        font-weight: 600; font-size: 10px; display: inline-block; border: 1px solid #A7F3D0;
    }
    .badge-pill-diff {
        background-color: #FEE2E2; color: #991B1B; padding: 2px 7px; border-radius: 9999px;
        font-weight: 700; font-size: 10px; display: inline-block; border: 1px solid #FECACA;
    }
</style>"""
st.markdown(CSS, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# HELPER FORMAT & RENDER HTML
# ----------------------------------------------------------------------------
def format_rupiah(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _esc(v):
    return html.escape(str(v))


def _badge(selisih):
    if abs(selisih) < 0.005:
        return '<span class="badge-pill-match">✅ COCOK</span>'
    return '<span class="badge-pill-diff">❌ SELISIH</span>'


def _cell(kind, v):
    """Satu sel <td> sesuai jenis kolomnya."""
    if kind == 'money':
        return f"<td class='text-right'>{format_rupiah(v)}</td>"
    if kind == 'diff':
        cls = "diff-match" if abs(v) < 0.005 else "diff-alert"
        return f"<td class='text-right'><span class='{cls}'>{format_rupiah(v)}</span></td>"
    if kind == 'status':
        return f"<td class='text-center'>{_badge(v)}</td>"
    if kind == 'code':
        return f"<td class='text-code'>{_esc(v)}</td>"
    if kind == 'center':
        return f"<td class='text-center'>{_esc(v)}</td>"
    return f"<td>{_esc(v)}</td>"


def _cell_total(kind, v):
    if v is None:
        return "<td class='text-center'>-</td>" if kind == 'center' else "<td>-</td>"
    if kind in ('money', 'diff', 'status'):
        return _cell(kind, v)
    return f"<td>{_esc(v)}</td>"


_TH_CLASS = {'money': 'text-right', 'diff': 'text-right', 'center': 'text-center', 'status': 'text-center'}


def render_table(df, cols, total, max_rows=None):
    """Bangun tabel HTML. `cols` = [(header, nama_kolom_df, jenis)], `total` = nilai per kolom."""
    keys = [k for _, k, _ in cols]
    kinds = [kd for _, _, kd in cols]
    data = df[keys].head(max_rows) if max_rows else df[keys]

    head = "".join(f"<th class='{_TH_CLASS.get(kd, '')}'>{_esc(h)}</th>" for h, _, kd in cols)
    body = "".join(
        "<tr>" + "".join(_cell(kd, v) for kd, v in zip(kinds, row)) + "</tr>"
        for row in data.itertuples(index=False, name=None)
    )
    foot = "<tr class='total-row'>" + "".join(_cell_total(kd, v) for kd, v in zip(kinds, total)) + "</tr>"

    st.markdown(
        f"<div class='table-container'><table class='custom-table'>"
        f"<thead><tr>{head}</tr></thead><tbody>{body}{foot}</tbody></table></div>",
        unsafe_allow_html=True,
    )


def metric_box(judul, nilai, gradient):
    st.markdown(
        f"""<div class="metric-box" style="background: linear-gradient(135deg, {gradient});">
        <div class="metric-title">{judul}</div><div class="metric-value">{format_rupiah(nilai)}</div></div>""",
        unsafe_allow_html=True,
    )


def kartu_status(tot_lra, tot_pembanding, tot_selisih, teks_cocok, teks_selisih):
    if abs(tot_selisih) < 0.005:
        st.markdown(
            f"<div class='status-card-match'><h4 style='color: #065F46; margin:0;'>✅ STATUS: COCOK DENGAN LRA</h4>"
            f"<p style='color: #047857; margin: 4px 0 0 0; font-size:12px;'>{teks_cocok[0]} <b>({format_rupiah(tot_lra)})</b> "
            f"sama persis dengan {teks_cocok[1]} <b>({format_rupiah(tot_pembanding)})</b>. Tidak ada selisih.</p></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div class='status-card-diff'><h4 style='color: #991B1B; margin:0;'>⚠️ STATUS: TERDAPAT SELISIH REKONSILIASI</h4>"
            f"<p style='color: #B91C1C; margin: 4px 0 0 0; font-size:12px;'>Ditemukan selisih sebesar "
            f"<b>{format_rupiah(tot_selisih)}</b> antara {teks_selisih[0]} ({format_rupiah(tot_lra)}) "
            f"dan {teks_selisih[1]} ({format_rupiah(tot_pembanding)}).</p></div>",
            unsafe_allow_html=True,
        )


# ----------------------------------------------------------------------------
# LOAD DATA (semua di-cache supaya tidak dibaca ulang tiap interaksi)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_master_rak():
    if not FILE_RAK_PERSEDIAAN.exists() or not FILE_RAK_MODAL.exists():
        return None, None, f"File '{FILE_RAK_PERSEDIAAN.name}' atau '{FILE_RAK_MODAL.name}' tidak ditemukan."

    kode_p = pd.read_excel(FILE_RAK_PERSEDIAAN, header=None, usecols=[1], skiprows=2).iloc[:, 0]
    rek_persediaan = set(kode_p.dropna().astype(str).str.strip()) - {'-', 'KODE REKENING'}

    df_m = pd.read_excel(FILE_RAK_MODAL, header=None, usecols=[0, 1, 2, 3], skiprows=1)
    df_m.columns = ['Kode Kategori', 'Uraian Kategori', 'Uraian Rekening', 'Kode Rekening']
    df_m = df_m.dropna(subset=['Kode Rekening']).copy()
    df_m['Kode Rekening'] = df_m['Kode Rekening'].astype(str).str.strip()
    df_m = df_m[df_m['Kode Rekening'].str.startswith('5.2')]
    df_modal_map = df_m[['Kode Rekening', 'Kode Kategori', 'Uraian Kategori']].drop_duplicates('Kode Rekening')

    return rek_persediaan, df_modal_map, None


def label_bulan(raw):
    """'04_APRIL' / '08_AGUST' / '09_SEPT' -> nama bulan lengkap. Tak dikenali -> apa adanya."""
    raw = str(raw).strip()
    nomor = raw.split('_')[0]
    if nomor.isdigit() and 1 <= int(nomor) <= 12:
        return NAMA_BULAN[int(nomor) - 1]
    idx = NAMA_BULAN_PREFIX.get(raw.split('_')[-1][:3].upper())
    return NAMA_BULAN[idx] if idx is not None else (raw or '-')


def _format_tanggal(v):
    if hasattr(v, 'strftime'):
        return v.strftime('%d/%m/%Y')
    return '' if pd.isna(v) else str(v)


@st.cache_data(show_spinner="Sedang memproses data realisasi LRA...")
def proses_lra(data: bytes):
    """Baca LRA sekali saja, simpan hanya baris persediaan & belanja modal."""
    rek_persediaan, df_modal_map, _ = load_master_rak()

    xls = pd.ExcelFile(BytesIO(data), engine=EXCEL_ENGINE)
    sheet = SHEET_LRA if SHEET_LRA in xls.sheet_names else xls.sheet_names[0]

    # Cari baris header otomatis (tahan terhadap file dengan/tanpa kop di atas tabel)
    probe = xls.parse(sheet, header=None, nrows=30)
    cocok = probe.apply(lambda r: r.astype(str).str.strip().eq('Kode Rekening').any(), axis=1)
    if not cocok.any():
        raise ValueError(
            f"Kolom 'Kode Rekening' tidak ditemukan di sheet '{sheet}'. "
            f"Pastikan file yang diupload adalah LRA Realisasi Dokumen."
        )
    baris_header = int(cocok.idxmax())

    df = xls.parse(sheet, header=baris_header, usecols=lambda c: str(c).strip() in KOLOM_LRA)
    df.columns = df.columns.astype(str).str.strip()

    kurang = [c for c in KOLOM_WAJIB if c not in df.columns]
    if kurang:
        raise ValueError(f"Kolom wajib tidak ada di file LRA: {', '.join(kurang)}")

    df = df.reindex(columns=KOLOM_LRA)
    df[KOLOM_TEKS] = df[KOLOM_TEKS].fillna('')
    df['Kode Rekening'] = df['Kode Rekening'].astype(str).str.strip()

    daftar_skpd = sorted(s for s in df['Nama SKPD'].astype(str).unique() if s)

    df_pers = df[df['Kode Rekening'].isin(rek_persediaan)].copy()
    df_modal = df[df['Kode Rekening'].isin(df_modal_map['Kode Rekening'])].merge(
        df_modal_map, on='Kode Rekening', how='left'
    )

    for d in (df_pers, df_modal):
        d['Nilai Realisasi'] = pd.to_numeric(d['Nilai Realisasi'], errors='coerce').fillna(0)
        d['Tanggal Dokumen'] = d['Tanggal Dokumen'].map(_format_tanggal)
        d['Bulan'] = d['BULAN'].map(label_bulan)

    # Rekening 5.2 di LRA yang tidak ada di master RAK (tidak ikut dihitung sebagai Belanja Modal)
    tak = df[df['Kode Rekening'].str.startswith('5.2') & ~df['Kode Rekening'].isin(df_modal_map['Kode Rekening'])].copy()
    tak['Nilai Realisasi'] = pd.to_numeric(tak['Nilai Realisasi'], errors='coerce').fillna(0)
    tak['Bulan'] = tak['BULAN'].map(label_bulan)
    df_tak = tak[['BULAN', 'Bulan', 'Nama SKPD', 'Kode Rekening', 'Nama Rekening', 'Nilai Realisasi']].reset_index(drop=True)

    # Urutan bulan mengikuti kode di file ('01_JAN' < '02_FEB' < ...)
    daftar_bulan = sorted(set(df['BULAN'].astype(str)) - {''})

    return df_pers.reset_index(drop=True), df_modal.reset_index(drop=True), daftar_skpd, daftar_bulan, df_tak


@st.cache_data(show_spinner=False)
def baca_sipper(data: bytes):
    df = pd.read_excel(BytesIO(data), header=None, usecols=[5, 7], skiprows=11, engine=EXCEL_ENGINE)
    df.columns = ['Kode Rekening', 'Nilai SIPPER']
    df = df.dropna(subset=['Kode Rekening']).copy()
    df['Kode Rekening'] = df['Kode Rekening'].astype(str).str.strip()
    df['Nilai SIPPER'] = pd.to_numeric(df['Nilai SIPPER'], errors='coerce').fillna(0)
    return df.groupby('Kode Rekening', as_index=False)['Nilai SIPPER'].sum(), {}


@st.cache_data(show_spinner=False)
def baca_aset(data: bytes):
    """Laporan 'Rincian Pengadaan Aset'. Mengembalikan (nilai per rekening 5.2, info kop laporan)."""
    raw = pd.read_excel(BytesIO(data), header=None, usecols=[0, 1, 2, 3], engine=EXCEL_ENGINE)
    raw.columns = ['Kode', 'Kolom B', 'Kolom C', 'Nilai']
    kode = raw['Kode'].astype(str).str.strip()
    nilai = pd.to_numeric(raw['Nilai'], errors='coerce').fillna(0)

    # Kop laporan: unit kerja, tahun, dan tanggal cetak ("Barabai, 6 Oktober 2026")
    meta = {}
    for label, isi in zip(kode.head(15).str.upper(), raw['Kolom C'].head(15)):
        if label == 'UNIT KERJA' and pd.notna(isi):
            meta['unit_kerja'] = str(isi).strip()
        elif label == 'TAHUN' and pd.notna(isi):
            meta['tahun'] = str(isi).strip().removesuffix('.0')
    cetak = raw['Nilai'].astype(str).str.extract(r',\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s*$')[0].dropna()
    if not cetak.empty:
        meta['dicetak'] = cetak.iloc[-1]

    # Baris rekening non-5.2 (mis. 5.1.02 = belanja barang/jasa yang dikapitalisasi jadi aset):
    # tidak dibandingkan dengan Belanja Modal, tapi nilainya dilaporkan agar total file bisa ditelusuri.
    is_52 = kode.str.startswith('5.2')
    non_52 = kode.str.match(r'^5\.\d\.\d{2}\.\d{2}\.\d{3}\.\d{5}$') & ~is_52
    meta['non_52_nilai'] = float(nilai[non_52].sum())
    meta['non_52_rekening'] = int(kode[non_52].nunique())

    df = pd.DataFrame({'Kode Rekening': kode[is_52], 'Nilai Aset': nilai[is_52]})
    return df.groupby('Kode Rekening', as_index=False)['Nilai Aset'].sum(), meta


@st.cache_data(show_spinner=False)
def ke_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    return buf.getvalue()


# ----------------------------------------------------------------------------
# KONFIGURASI DUA TAB (logika tampilan sama, hanya label & kolom yang beda)
# ----------------------------------------------------------------------------
KOLOM_RINCI = [
    ('Kode Sub Kegiatan', 'Kode Sub Kegiatan', 'code'),
    ('Nama Sub Kegiatan', 'Nama Sub Kegiatan', 'text'),
    ('Kode Rekening', 'Kode Rekening', 'code'),
    ('Nama Rekening', 'Nama Rekening', 'text'),
    ('Nomor Dokumen', 'Nomor Dokumen', 'code'),
    ('Tanggal Dokumen', 'Tanggal Dokumen', 'center'),
    ('Bulan', 'Bulan', 'center'),
    ('Keterangan Dokumen', 'Keterangan Dokumen', 'text'),
    ('Nilai Realisasi (Rp)', 'Nilai Realisasi', 'money'),
]

CFG_PERSEDIAAN = {
    'key': 'sipper',
    'judul_uploader': "##### 📥 Pembanding Data SIPPER (Opsional)",
    'label_uploader': "Upload Data Aplikasi SIPPER (.xlsx)",
    'parser': baca_sipper,
    'pesan_sukses': "✅ Data Aplikasi SIPPER berhasil dimuat.",
    'pesan_gagal': "Format file SIPPER tidak sesuai",
    'judul_rekap': "1. Rekapitulasi per Kode Rekening Persediaan",
    'kolom_id': [('Kode Rekening', 'Kode Rekening', 'code'), ('Nama Rekening', 'Nama Rekening', 'text')],
    'kolom_pembanding': 'Nilai SIPPER',
    'header_pembanding': 'Nilai SIPPER (Rp)',
    'isi_kosong': {'Nama Rekening': 'Dari Data SIPPER'},
    'teks_cocok': ("Total Realisasi LRA", "SIPPER"),
    'teks_selisih': ("LRA", "SIPPER"),
    'info_upload': "💡 *Upload file SIPPER di atas untuk menampilkan perbandingan dan status selisih secara otomatis.*",
    'kosong': "Tidak ditemukan data persediaan untuk SKPD ini.",
}

CFG_MODAL = {
    'key': 'modal',
    'judul_uploader': "##### 📥 Pembanding Data Aplikasi Belanja Modal / Aset (Opsional)",
    'label_uploader': "Upload Data Rincian Pengadaan Aset (.xlsx)",
    'parser': baca_aset,
    'pesan_sukses': "✅ Data Aplikasi Belanja Modal / Aset berhasil dimuat.",
    'pesan_gagal': "Format file Belanja Modal tidak sesuai",
    'judul_rekap': "1. Rekapitulasi per Kategori & Rekening Belanja Modal",
    'kolom_id': [
        ('Kode Kategori', 'Kode Kategori', 'code'),
        ('Uraian Kategori', 'Uraian Kategori', 'text'),
        ('Kode Rekening', 'Kode Rekening', 'code'),
        ('Nama Rekening', 'Nama Rekening', 'text'),
    ],
    'kolom_pembanding': 'Nilai Aset',
    'header_pembanding': 'Nilai Aplikasi Aset (Rp)',
    'isi_kosong': {'Kode Kategori': '-', 'Uraian Kategori': '-', 'Nama Rekening': 'Dari Data Aplikasi Aset'},
    'teks_cocok': ("Total Belanja Modal LRA", "Aplikasi Aset"),
    'teks_selisih': ("Belanja Modal LRA", "Aplikasi Aset"),
    'info_upload': "💡 *Upload file Pengadaan Aset di atas untuk menampilkan perbandingan dan status selisih secara otomatis.*",
    'kosong': "Tidak ditemukan data belanja modal untuk SKPD ini.",
}


def _norm(teks):
    return " ".join(str(teks).upper().split())


def info_file_pembanding(meta, skpd_terpilih):
    """Tampilkan identitas file pembanding & peringatan jika tidak sesuai dengan SKPD terpilih."""
    unit = meta.get('unit_kerja')
    if unit:
        bagian = [f"Unit kerja di file: **{unit}**"]
        if meta.get('tahun'):
            bagian.append(f"Tahun {meta['tahun']}")
        if meta.get('dicetak'):
            bagian.append(f"dicetak {meta['dicetak']}")
        st.caption(" · ".join(bagian))
        if skpd_terpilih is None:
            st.warning(
                f"⚠️ File ini hanya memuat satu unit kerja (**{unit}**), sedangkan tampilan saat ini "
                f"**semua SKPD**. Selisih akan sangat besar dan menyesatkan. Pilih SKPD yang sesuai di bagian atas."
            )
        elif _norm(unit) != _norm(skpd_terpilih):
            st.warning(
                f"⚠️ File ini untuk unit kerja **{unit}**, sedangkan SKPD yang dipilih **{skpd_terpilih}**. "
                f"Hasil perbandingan di bawah tidak valid."
            )
    if meta.get('non_52_nilai'):
        st.info(
            f"ℹ️ File ini juga memuat {meta['non_52_rekening']:,} rekening di luar 5.2 (belanja yang dikapitalisasi "
            f"menjadi aset) senilai {format_rupiah(meta['non_52_nilai'])}. Bagian ini tidak ikut dibandingkan, "
            f"sehingga angka TOTAL di file lebih besar daripada angka pembanding di bawah."
        )


SEMUA_PILIHAN = "__SEMUA__"


def filter_rincian(df, kunci):
    """Filter rincian: per rekening, per sub kegiatan, dan pencarian teks pada nomor/keterangan dokumen."""
    nama_rek = df.drop_duplicates('Kode Rekening').set_index('Kode Rekening')['Nama Rekening']
    nama_sub = df.drop_duplicates('Kode Sub Kegiatan').set_index('Kode Sub Kegiatan')['Nama Sub Kegiatan']

    c1, c2, c3 = st.columns(3)
    with c1:
        rek = st.selectbox(
            "Filter rekening", [SEMUA_PILIHAN] + sorted(nama_rek.index), key=f"f_rek_{kunci}",
            format_func=lambda k: "-- Semua rekening --" if k == SEMUA_PILIHAN else f"{k} — {nama_rek[k]}",
        )
    with c2:
        sub = st.selectbox(
            "Filter sub kegiatan", [SEMUA_PILIHAN] + sorted(nama_sub.index), key=f"f_sub_{kunci}",
            format_func=lambda k: "-- Semua sub kegiatan --" if k == SEMUA_PILIHAN else f"{k or '(tanpa kode)'} — {nama_sub[k]}",
        )
    with c3:
        cari = st.text_input("Cari nomor / keterangan dokumen", key=f"f_cari_{kunci}").strip()

    if rek != SEMUA_PILIHAN:
        df = df[df['Kode Rekening'] == rek]
    if sub != SEMUA_PILIHAN:
        df = df[df['Kode Sub Kegiatan'] == sub]
    if cari:
        cocok = (df['Nomor Dokumen'].astype(str).str.contains(cari, case=False, regex=False)
                 | df['Keterangan Dokumen'].astype(str).str.contains(cari, case=False, regex=False))
        df = df[cocok]
    return df


def tampilkan_rekon(df, cfg, df_semua_bulan, urut_bulan, skpd_terpilih):
    """Render satu tab rekonsiliasi (uploader pembanding + rekap + rincian + per bulan).

    df = data setelah filter SKPD & bulan; df_semua_bulan = data setelah filter SKPD saja."""
    st.markdown(cfg['judul_uploader'])
    f = st.file_uploader(cfg['label_uploader'], type=["xlsx"], key=f"up_{cfg['key']}")

    pembanding = None
    if f:
        try:
            pembanding, meta_file = cfg['parser'](f.getvalue())
            st.success(cfg['pesan_sukses'])
            info_file_pembanding(meta_file, skpd_terpilih)
        except Exception as e:
            st.error(f"{cfg['pesan_gagal']}: {e}")

    st.markdown("---")
    st.subheader(cfg['judul_rekap'])
    if df.empty:
        st.warning(cfg['kosong'])
        tampilkan_bulanan(df_semua_bulan, cfg, urut_bulan)
        return

    kolom_id = cfg['kolom_id']
    key_id = [k for _, k, _ in kolom_id]
    awal_total = ['TOTAL', 'JUMLAH KESELURUHAN'] + [None] * (len(kolom_id) - 2)

    rekap = df.groupby(key_id, as_index=False)['Nilai Realisasi'].sum()

    if pembanding is not None:
        nilai_p = cfg['kolom_pembanding']
        rekap = rekap.merge(pembanding, on='Kode Rekening', how='outer')
        rekap['Nilai Realisasi'] = rekap['Nilai Realisasi'].fillna(0)
        rekap[nilai_p] = rekap[nilai_p].fillna(0)
        for kolom, isi in cfg['isi_kosong'].items():
            rekap[kolom] = rekap[kolom].fillna(isi)
        rekap['Selisih'] = rekap['Nilai Realisasi'] - rekap[nilai_p]

        tot_lra = rekap['Nilai Realisasi'].sum()
        tot_p = rekap[nilai_p].sum()
        tot_selisih = rekap['Selisih'].sum()

        cols = kolom_id + [
            ('Realisasi LRA (Rp)', 'Nilai Realisasi', 'money'),
            (cfg['header_pembanding'], nilai_p, 'money'),
            ('Selisih (Rp)', 'Selisih', 'diff'),
            ('Status', 'Selisih', 'status'),
        ]
        hanya_selisih = st.checkbox("🔍 Tampilkan hanya rekening yang selisih", key=f"selisih_{cfg['key']}")
        tampil = rekap[rekap['Selisih'].abs() >= 0.005] if hanya_selisih else rekap
        if tampil.empty:
            st.success("✅ Tidak ada rekening yang selisih.")
        else:
            render_table(tampil, cols, awal_total + [tot_lra, tot_p, tot_selisih, tot_selisih])
        if hanya_selisih:
            st.caption(f"Menampilkan {len(tampil):,} dari {len(rekap):,} rekening. Baris TOTAL tetap dihitung dari seluruh rekening.")
        kartu_status(tot_lra, tot_p, tot_selisih, cfg['teks_cocok'], cfg['teks_selisih'])

        ekspor = rekap[key_id + ['Nilai Realisasi', nilai_p, 'Selisih']].rename(columns={
            'Nilai Realisasi': 'Realisasi LRA (Rp)', nilai_p: cfg['header_pembanding'], 'Selisih': 'Selisih (Rp)'})
        ekspor['Status'] = rekap['Selisih'].abs().lt(0.005).map({True: 'COCOK', False: 'SELISIH'})
    else:
        cols = kolom_id + [('Realisasi (Rp)', 'Nilai Realisasi', 'money')]
        render_table(rekap, cols, awal_total + [rekap['Nilai Realisasi'].sum()])
        st.info(cfg['info_upload'])
        ekspor = rekap[key_id + ['Nilai Realisasi']].rename(columns={'Nilai Realisasi': 'Realisasi (Rp)'})

    st.download_button(
        "⬇️ Unduh hasil rekon (Excel)",
        data=ke_excel(ekspor),
        file_name=f"rekon_{cfg['key']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_rekap_{cfg['key']}",
    )

    # --- 2. RINCIAN ---
    st.markdown("---")
    st.subheader("2. Rincian Dokumen Realisasi LRA")
    d = filter_rincian(df, cfg['key'])
    if d.empty:
        st.info("Tidak ada dokumen yang cocok dengan filter.")
    else:
        tot_rinci = d['Nilai Realisasi'].sum()
        total_rinci = ['TOTAL', 'JUMLAH REALISASI DOKUMEN'] + [None] * (len(KOLOM_RINCI) - 3) + [tot_rinci]
        render_table(d, KOLOM_RINCI, total_rinci, max_rows=MAX_BARIS_RINCI)

        if len(d) > MAX_BARIS_RINCI:
            st.caption(
                f"Menampilkan {MAX_BARIS_RINCI:,} dari {len(d):,} baris agar halaman tetap ringan. "
                f"Total di atas dihitung dari seluruh baris. Persempit dengan filter di atas, atau unduh Excel untuk data lengkap."
            )
        st.download_button(
            "⬇️ Unduh rincian (Excel)",
            data=ke_excel(d[[k for _, k, _ in KOLOM_RINCI]]),
            file_name=f"rincian_{cfg['key']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{cfg['key']}",
        )

    tampilkan_bulanan(df_semua_bulan, cfg, urut_bulan)


def tampilkan_bulanan(df, cfg, urut_bulan):
    """Bagian 'Realisasi per Bulan' di dalam tab rekon. `df` = data SKPD terpilih, SEMUA bulan."""
    st.markdown("---")
    st.subheader("3. Realisasi per Bulan")
    st.caption("Bagian ini selalu menampilkan semua bulan (tidak terpengaruh filter bulan).")
    if not urut_bulan:
        st.warning("Kolom BULAN tidak ditemukan atau kosong di file LRA.")
        return
    if df.empty:
        st.warning(cfg['kosong'])
        return

    # --- ringkasan per bulan + kumulatif ---
    ring = df.groupby('BULAN')['Nilai Realisasi'].sum().reindex(urut_bulan, fill_value=0).rename('Realisasi').to_frame()
    ring['Kumulatif'] = ring['Realisasi'].cumsum()
    ring = ring.reset_index()
    ring['Bulan'] = ring['BULAN'].map(label_bulan)

    st.markdown("##### Ringkasan per Bulan")
    cols = [
        ('Bulan', 'Bulan', 'text'),
        ('Realisasi Bulan Ini (Rp)', 'Realisasi', 'money'),
        ('Kumulatif s.d. Bulan Ini (Rp)', 'Kumulatif', 'money'),
    ]
    render_table(ring, cols, ['TOTAL', ring['Realisasi'].sum(), None])

    # Label sumbu '01 Jan' agar urutan grafik mengikuti kalender
    grafik = ring.assign(Label=ring['BULAN'].str.replace('_', ' ').str.title()).set_index('Label')
    st.bar_chart(grafik['Realisasi'])

    # --- rekening x bulan ---
    st.markdown("##### Per Rekening per Bulan")
    kolom_id = cfg['kolom_id']
    key_id = [k for _, k, _ in kolom_id]
    pv = df.pivot_table(index=key_id, columns='BULAN', values='Nilai Realisasi', aggfunc='sum', fill_value=0)
    pv = pv.reindex(columns=urut_bulan, fill_value=0)
    pv['Total'] = pv.sum(axis=1)
    pv.columns.name = None
    pv = pv.reset_index()

    cols = kolom_id + [(label_bulan(b), b, 'money') for b in urut_bulan] + [('Total (Rp)', 'Total', 'money')]
    total = ['TOTAL', 'JUMLAH KESELURUHAN'] + [None] * (len(kolom_id) - 2)
    total += [pv[b].sum() for b in urut_bulan] + [pv['Total'].sum()]
    render_table(pv, cols, total)

    st.download_button(
        "⬇️ Unduh tabel per bulan (Excel)",
        data=ke_excel(pv.rename(columns={b: label_bulan(b) for b in urut_bulan})),
        file_name=f"realisasi_bulanan_{cfg['key']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_bln_{cfg['key']}",
    )


# ----------------------------------------------------------------------------
# SIDEBAR & HEADER
# ----------------------------------------------------------------------------
rek_persediaan, df_modal_map, error_msg = load_master_rak()

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135679.png", width=70)
    st.markdown("### **Panel Kontrol**")
    st.info("📌 **Master RAK Belanja sudah di SISTEM**")

    f_lra = st.file_uploader("📥 Upload LRA Realisasi (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.caption("Pemerintah Kabupaten Hulu Sungai Tengah © 2026")

st.markdown('<div class="main-header">🏛️ Rekonsiliasi Rekening Persediaan & Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Realisasi LRA terhadap Klasifikasi Persediaan dan Belanja Modal (Akun 5.2)</div>', unsafe_allow_html=True)

if error_msg:
    st.error(f"⚠️ {error_msg}")
    st.stop()

if not f_lra:
    st.info("👈 Silakan upload file **LRA REALISASI** pada menu di sebelah kiri.")
    st.stop()

# ----------------------------------------------------------------------------
# PROSES UTAMA
# ----------------------------------------------------------------------------
try:
    df_rekon_pers, df_rekon_modal, daftar_skpd, daftar_bulan, df_tak = proses_lra(f_lra.getvalue())
except Exception as e:
    st.error(f"⚠️ Gagal membaca file LRA: {e}")
    st.stop()

SEMUA = "-- SEMUA SKPD --"
col_skpd, col_bulan = st.columns([2, 2])
with col_skpd:
    pilihan_skpd = st.selectbox("🎯 **Pilih Perangkat Daerah (SKPD):**", [SEMUA] + daftar_skpd)
with col_bulan:
    bulan_pilih = st.multiselect(
        "📅 **Pilih Bulan** (kosong = semua bulan):",
        options=daftar_bulan, format_func=label_bulan, placeholder="Semua bulan",
    )

# Filter SKPD dulu (dipakai tab Per Bulan), lalu filter bulan (dipakai metrik & tab rekon)
if pilihan_skpd == SEMUA:
    df_pers_skpd, df_modal_skpd = df_rekon_pers, df_rekon_modal
else:
    df_pers_skpd = df_rekon_pers[df_rekon_pers['Nama SKPD'] == pilihan_skpd]
    df_modal_skpd = df_rekon_modal[df_rekon_modal['Nama SKPD'] == pilihan_skpd]

if bulan_pilih:
    df_pers_filtered = df_pers_skpd[df_pers_skpd['BULAN'].isin(bulan_pilih)]
    df_modal_filtered = df_modal_skpd[df_modal_skpd['BULAN'].isin(bulan_pilih)]
    st.caption("Periode ditampilkan: " + ", ".join(label_bulan(b) for b in bulan_pilih))
else:
    df_pers_filtered, df_modal_filtered = df_pers_skpd, df_modal_skpd

df_tak_f = df_tak if pilihan_skpd == SEMUA else df_tak[df_tak['Nama SKPD'] == pilihan_skpd]
if bulan_pilih:
    df_tak_f = df_tak_f[df_tak_f['BULAN'].isin(bulan_pilih)]

total_p = df_pers_filtered['Nilai Realisasi'].sum()
total_m = df_modal_filtered['Nilai Realisasi'].sum()

m1, m2, m3 = st.columns(3)
with m1:
    metric_box("📦 Total Belanja Persediaan (LRA)", total_p, "#0D9488, #14B8A6")
with m2:
    metric_box("🏢 Total Belanja Modal (LRA)", total_m, "#4F46E5, #6366F1")
with m3:
    metric_box("📊 Total Gabungan Realisasi", total_p + total_m, "#1E293B, #334155")

if df_tak_f.empty:
    st.caption("✅ Semua rekening Belanja Modal (5.2) di LRA terpetakan di master RAK.")
else:
    st.warning(
        f"⚠️ Ada {df_tak_f['Kode Rekening'].nunique():,} rekening Belanja Modal (5.2) di LRA senilai "
        f"{format_rupiah(df_tak_f['Nilai Realisasi'].sum())} yang **tidak ada di master RAK**, "
        f"sehingga tidak ikut dihitung di angka di atas. Periksa apakah master RAK perlu diperbarui."
    )
    with st.expander("Lihat rekening yang tidak terpetakan"):
        ringkas = df_tak_f.groupby(['Nama SKPD', 'Kode Rekening', 'Nama Rekening'], as_index=False)['Nilai Realisasi'].sum()
        render_table(
            ringkas,
            [('Nama SKPD', 'Nama SKPD', 'text'), ('Kode Rekening', 'Kode Rekening', 'code'),
             ('Nama Rekening', 'Nama Rekening', 'text'), ('Realisasi (Rp)', 'Nilai Realisasi', 'money')],
            ['TOTAL', None, None, ringkas['Nilai Realisasi'].sum()],
        )

skpd_terpilih = None if pilihan_skpd == SEMUA else pilihan_skpd

tab1, tab2 = st.tabs(["📦 REKON PERSEDIAAN", "🏢 REKON BELANJA MODAL"])
with tab1:
    tampilkan_rekon(df_pers_filtered, CFG_PERSEDIAAN, df_pers_skpd, daftar_bulan, skpd_terpilih)
with tab2:
    tampilkan_rekon(df_modal_filtered, CFG_MODAL, df_modal_skpd, daftar_bulan, skpd_terpilih)
