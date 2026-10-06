"""
SISTEM MONITORING KEUANGAN & KINERJA NIAGA
PT PLN (Persero) UP3 Medan Utara

Jalankan :  streamlit run app.py
Kebutuhan:  streamlit, pandas, numpy, openpyxl  (+ xlrd bila ada file .xls)

Alur eksekusi:
  1. Konstanta & fungsi bantu
  2. Parser file (di-cache, in-memory, tanpa database)
  3. PRE-COMPUTATION: semua file dibaca dari st.session_state dan semua
     metrik dihitung SEBELUM tab dirender, sehingga Tab 2 tidak pernah
     bergantung pada urutan render Tab 3/4/5.
  4. Sidebar (pemilih bulan evaluasi + status upload)
  5. Render 5 tab
"""

import datetime as dt
import html
import io
import re

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# =========================================================================
# 1. KONFIGURASI, KONSTANTA, STYLING
# =========================================================================
st.set_page_config(
    page_title="Dashboard Niaga - UP3 Medan Utara",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

TAHUN = 2026
ID_UP3 = "12128"
NAMA_UP3 = "UP3 Medan Utara"

TARGET_ULP = [
    {"id": "12801", "nama": "MEDAN TIMUR", "key": "medan_timur", "t_pal": 515_000_000, "t_ts": 10_000_000, "t_total": 525_000_000},
    {"id": "12802", "nama": "BELAWAN", "key": "belawan", "t_pal": 295_000_000, "t_ts": 5_000_000, "t_total": 300_000_000},
    {"id": "12803", "nama": "HELVETIA", "key": "helvetia", "t_pal": 43_000_000, "t_ts": 2_000_000, "t_total": 45_000_000},
    {"id": "12804", "nama": "LABUHAN", "key": "labuhan", "t_pal": 247_000_000, "t_ts": 3_000_000, "t_total": 250_000_000},
    {"id": "12805", "nama": "DENAI", "key": "denai", "t_pal": 197_000_000, "t_ts": 3_000_000, "t_total": 200_000_000},
]

# Baris Excel target pelunasan per ULP pada sheet 'PELUNASAN PRR'
BARIS_TARGET_PELUNASAN = {"12801": 10, "12802": 14, "12803": 18, "12804": 22, "12805": 26}

# Riwayat cadangan TOTAL (PAL + TS) Jan s.d. Jul 2026. Hanya dipakai bila
# File Master Kinerja belum diunggah / sheet riwayatnya tidak terbaca.
HISTORIS_CADANGAN = {
    "12801": [538413056, 788770092, 524042368, 463692774, 578753035, 407488699, 491208963],
    "12802": [406965719, 450904955, 409064637, 266140474, 354805477, 245161324, 310085436],
    "12803": [33255070, 54611771, 39395628, 18567809, 47109229, 25783095, 31593876],
    "12804": [187565928, 287709388, 178054633, 237142328, 205740200, 189616330, 230045125],
    "12805": [112969944, 256403336, 112244960, 116244782, 111320801, 119660549, 123729277],
}

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
KODE_BULAN = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGU", "SEP", "OKT", "NOV", "DES"]

# Semua variasi penulisan nama bulan -> nomor bulan (1..12)
ALIAS_BULAN = {
    "JANUARI": 1, "JAN": 1, "JANUARY": 1,
    "FEBRUARI": 2, "PEBRUARI": 2, "FEB": 2, "FEBRUARY": 2,
    "MARET": 3, "MAR": 3, "MARCH": 3,
    "APRIL": 4, "APR": 4,
    "MEI": 5, "MAY": 5,
    "JUNI": 6, "JUN": 6, "JUNE": 6,
    "JULI": 7, "JUL": 7, "JULY": 7,
    "AGUSTUS": 8, "AGU": 8, "AGT": 8, "AGS": 8, "AUG": 8, "AUGUST": 8,
    "SEPTEMBER": 9, "SEP": 9, "SEPT": 9,
    "OKTOBER": 10, "OKT": 10, "OCT": 10, "OCTOBER": 10,
    "NOVEMBER": 11, "NOV": 11, "NOP": 11, "NOPEMBER": 11,
    "DESEMBER": 12, "DES": 12, "DEC": 12, "DECEMBER": 12,
}

BOBOT = 2.0  # bobot tiap indikator Percepatan Cash In

# Kunci session_state semua slot upload
K_CLOSING = "up_closing"
K_KINERJA = "up_kinerja"
K_TS_PRABAYAR = "up_ts_prabayar"
JENIS_PER_ULP = ["pal", "ts", "prr_tunai", "prr_cicilan", "eks_tunai", "eks_cicilan"]


def k(jenis, ulp):
    return f"up_{jenis}_{ulp['key']}"


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
.hero-banner { background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); padding: 22px 28px;
  border-radius: 14px; color: white; margin-bottom: 22px; box-shadow: 0 4px 15px rgba(30,64,175,.12); }
.hero-title { font-size: 22px; font-weight: 800; margin-bottom: 3px; letter-spacing: -0.3px; }
.hero-subtitle { font-size: 13px; opacity: .92; }
.kpi-container { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
.kpi-card { flex: 1; min-width: 210px; background: white; padding: 15px 18px; border-radius: 12px;
  border: 1px solid #e2e8f0; box-shadow: 0 2px 6px rgba(0,0,0,.02); }
.kpi-label { font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 6px; }
.kpi-value { font-size: 20px; font-weight: 800; color: #0f172a; }
.upload-box-empty { background: #fff5f5; border: 2px dashed #ef4444; border-radius: 12px;
  padding: 10px 14px; margin-bottom: 6px; }
.upload-box-filled { background: #f0fdf4; border: 2px solid #22c55e; border-radius: 12px;
  padding: 10px 14px; margin-bottom: 6px; box-shadow: 0 2px 8px rgba(34,197,94,.12); }
.upload-title { font-size: 12px; font-weight: 700; color: #334155; margin-bottom: 4px; }
.badge-red, .badge-green { display: inline-block; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 6px; }
.badge-red { background: #fee2e2; color: #b91c1c; }
.badge-green { background: #dcfce7; color: #15803d; }
.rekap-wrap { width: 100%; overflow-x: auto; margin-bottom: 18px; border: 1px solid #93c5fd; border-radius: 8px; }
.rekap-table { width: 100%; border-collapse: collapse; background: #fff; }
.rekap-table th { background: #bfdbfe !important; color: #000 !important; font-weight: 700; font-size: 15px;
  text-align: center; padding: 10px 12px; border: 1px solid #93c5fd; white-space: nowrap; }
.rekap-table td { font-size: 15px; padding: 9px 12px; border: 1px solid #e2e8f0; color: #1e293b;
  background: #fff; white-space: nowrap; }
.rekap-table td.l { text-align: left; } .rekap-table td.c { text-align: center; } .rekap-table td.r { text-align: right; }
.rekap-table tr.rekap-footer td { background: #bfdbfe !important; color: #000 !important; font-weight: 700;
  border: 1px solid #93c5fd; }
.rp { display: flex; justify-content: space-between; gap: 12px; }
.upload-box-empty .badge-red, .upload-box-empty .badge-green,
.upload-box-filled .badge-green { margin: 2px 4px 2px 0; }
.cashin-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; background: #fff; border: 1px solid #cbd5e1; }
.cashin-table th { background: #334155; color: #fff; padding: 8px 10px; font-size: 12px; text-align: center;
  border: 1px solid #475569; font-weight: 700; }
.cashin-table td { padding: 7px 10px; font-size: 12px; border: 1px solid #cbd5e1; color: #1e293b; }
.cashin-table td.r { text-align: right; } .cashin-table td.c { text-align: center; }
.cashin-table td.cashin-unit { font-weight: 700; vertical-align: middle; text-align: center; background: #f8fafc; }
.cashin-table td.cashin-rank { font-weight: 800 !important; font-size: 48px !important; line-height: 1;
  vertical-align: middle; text-align: center; background: #f8fafc; color: #0284c7 !important; }
.cashin-total-row td { background: #f8fafc; font-weight: 700; }
.cashin-pct-row td { background: #f1f5f9; font-weight: 800; }
</style>
""", unsafe_allow_html=True)


# =========================================================================
# 2. FUNGSI BANTU: FORMAT & KOMPONEN
# =========================================================================
def fmt_angka(x, rupiah=True):
    """1234567 -> 'Rp 1.234.567'. Nol/kosong -> '-'. Nilai negatif tetap tampil."""
    if x is None or pd.isna(x) or float(x) == 0:
        return "-"
    s = f"{abs(float(x)):,.0f}".replace(",", ".")
    s = f"Rp {s}" if rupiah else s
    return f"-{s}" if x < 0 else s


def fmt_persen(x):
    """85.5 -> '85,50%' (dua angka di belakang koma)."""
    if x is None or pd.isna(x):
        return "-"
    return f"{x:.2f}".replace(".", ",") + "%"


def fmt_desimal(x, n=2):
    return f"{x:.{n}f}".replace(".", ",")


def format_tabel(df, kolom_uang=(), kolom_persen=(), rupiah=True):
    out = df.copy()
    for c in kolom_uang:
        if c in out.columns:
            out[c] = out[c].apply(lambda v: fmt_angka(v, rupiah))
    for c in kolom_persen:
        if c in out.columns:
            out[c] = out[c].apply(fmt_persen)
    return out


KOLOM_KIRI = {"Nama ULP", "Nama Unit"}
KOLOM_TENGAH = {"No", "ID ULP", "Kode Unit", "Kode", "Proporsional", "Baris Excel"}


def tampil_df(df):
    """
    Tabel HTML selebar kontainer: header biru terang berfont hitam, dan baris penutup
    (total UP3) bergaya sama sebagai footer. st.dataframe tidak bisa diwarnai header-nya.
    """
    def kelas(c):
        return "l" if c in KOLOM_KIRI else ("c" if c in KOLOM_TENGAH else "r")

    kepala = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    n = len(df)
    ada_footer = n > 0 and ID_UP3 in [str(v) for v in df.iloc[-1].tolist()[:3]]
    badan = ""
    for i in range(n):
        sel = "".join(
            f'<td class="{kelas(c)}">{"" if pd.isna(v) else sel_rp(v)}</td>'
            for c, v in zip(df.columns, df.iloc[i].tolist()))
        badan += f'<tr class="rekap-footer">{sel}</tr>' if (ada_footer and i == n - 1) else f"<tr>{sel}</tr>"
    st.markdown(
        f'<div class="rekap-wrap"><table class="rekap-table"><thead><tr>{kepala}</tr></thead>'
        f"<tbody>{badan}</tbody></table></div>",
        unsafe_allow_html=True,
    )


VERSI_APP = "v7 · 6 Okt 2026"
perlu_rerun = []         # diisi bila ada slot upload yang berubah pada run ini


class Berkas:
    """Salinan file terunggah (nama + isi) yang disimpan sendiri di session_state."""

    def __init__(self, f):
        self.name = f.name
        self._isi = f.getvalue()
        self.tanda = (f.name, len(self._isi), getattr(f, "file_id", None))

    def getvalue(self):
        return self._isi


PETA = {}                # {kunci slot per-ULP: Berkas} hasil pemetaan upload grup
PERLU_MANUAL = {}        # {jenis: [Berkas yang ULP-nya tidak terdeteksi otomatis]}
PILIH_KOSONG = "— pilih ULP —"


def file_di(key):
    """Berkas pada slot `key` (None bila kosong). Aman dipanggil sebelum widget dirender."""
    if key in PETA:
        return PETA[key]
    return st.session_state.get("_berkas_" + key)


def bytes_di(key):
    f = file_di(key)
    return f.getvalue() if f is not None else None


def custom_file_uploader(label, key, allowed_types=("xls", "xlsx")):
    """Slot upload dengan indikator: merah putus-putus (kosong) / hijau solid (terisi)."""
    f = file_di(key)
    if f is not None:
        kelas = "upload-box-filled"
        badge = f'<span class="badge-green">✅ Terupload: {html.escape(f.name)}</span>'
    else:
        kelas = "upload-box-empty"
        badge = '<span class="badge-red">⚠️ Belum Diupload</span>'
    st.markdown(
        f'<div class="{kelas}"><div class="upload-title">{html.escape(label)}</div>{badge}</div>',
        unsafe_allow_html=True,
    )
    baru = st.file_uploader(label, type=list(allowed_types), key=key, label_visibility="collapsed")
    # Sinkronkan ke salinan milik aplikasi; bila berubah, jalankan ulang agar seluruh
    # perhitungan (yang dilakukan sebelum tab dirender) langsung memakai file ini.
    tanda_baru = (baru.name, baru.size, getattr(baru, "file_id", None)) if baru is not None else None
    if tanda_baru != (f.tanda if f is not None else None):
        st.session_state["_berkas_" + key] = Berkas(baru) if baru is not None else None
        perlu_rerun.append(key)
    return baru


def uploader_grup(label, jenis, allowed_types=("xls", "xlsx")):
    """
    Satu slot untuk kelima file ULP sekaligus (pilih/seret 5 file). Tiap file dipetakan
    otomatis ke ULP-nya; status per ULP tampil sebagai badge merah/hijau.
    """
    daftar = st.session_state.get("_berkas_grp_" + jenis) or []
    badge, terisi = "", 0
    for u in TARGET_ULP:
        bk = PETA.get(k(jenis, u))
        if bk is not None:
            terisi += 1
            badge += f'<span class="badge-green">✅ {u["nama"]}: {html.escape(bk.name)}</span> '
        else:
            badge += f'<span class="badge-red">⚠️ {u["nama"]}: Belum Diupload</span> '
    kelas = "upload-box-filled" if terisi == len(TARGET_ULP) else "upload-box-empty"
    st.markdown(
        f'<div class="{kelas}"><div class="upload-title">{html.escape(label)} — {terisi}/{len(TARGET_ULP)} ULP</div>'
        f"{badge}</div>",
        unsafe_allow_html=True,
    )
    baru = st.file_uploader(label, type=list(allowed_types), key="grp_" + jenis,
                            accept_multiple_files=True, label_visibility="collapsed") or []
    tanda_baru = [(f.name, f.size, getattr(f, "file_id", None)) for f in baru]
    if tanda_baru != [b.tanda for b in daftar]:
        st.session_state["_berkas_grp_" + jenis] = [Berkas(f) for f in baru]
        perlu_rerun.append(jenis)
    # File yang ULP-nya tidak dikenali dari nama/isi: pengguna memilih sendiri
    for bk in PERLU_MANUAL.get(jenis, []):
        st.selectbox(f"ULP untuk file “{bk.name}”", [PILIH_KOSONG] + [u["nama"] for u in TARGET_ULP],
                     key=f"pilih_{jenis}_{bk.name}")
    return baru


def sel_rp(teks):
    """'Rp 1.234' -> 'Rp' rata kiri dan angka rata kanan dalam satu sel (gaya akuntansi)."""
    m = re.match(r"^(-?)Rp\s+(.*)$", str(teks))
    if not m:
        return html.escape(str(teks))
    return f'<div class="rp"><span>Rp</span><span>{m.group(1)}{html.escape(m.group(2))}</span></div>'


def angka(v):
    """Konversi sel Excel ke float; kosong/teks -> 0."""
    n = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(n) else float(n)


def bulan_dari_label(v):
    """Nomor bulan dari isi sel header ('AGUSTUS', 'Agt 2026', tanggal) atau None."""
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
        return v.month
    if not isinstance(v, str):
        return None
    token = re.findall(r"[A-Z]+", v.upper())
    if len(token) == 0 or len(token) > 2:
        return None
    return ALIAS_BULAN.get(token[0])


def bulan_dari_nama_file(nama):
    """Deteksi bulan dari nama file: nama bulan, MMYYYY, atau YYYYMM."""
    s = nama.upper()
    for alias in sorted(ALIAS_BULAN, key=len, reverse=True):
        if re.search(rf"(?<![A-Z]){alias}(?![A-Z])", s):
            return ALIAS_BULAN[alias]
    m = re.search(r"(?<!\d)(0[1-9]|1[0-2])[-_. ]?20\d{2}(?!\d)", s)
    if m:
        return int(m.group(1))
    m = re.search(r"(?<!\d)20\d{2}[-_. ]?(0[1-9]|1[0-2])(?!\d)", s)
    if m:
        return int(m.group(1))
    return None


# =========================================================================
# 3. PARSER FILE (semua menerima bytes agar bisa di-cache)
# =========================================================================
def _baca(b, **kw):
    return pd.read_excel(io.BytesIO(b), header=None, **kw)


def _tabel_berheader(b, sheet, wajib):
    """Baca sheet, cari baris header yang memuat semua label `wajib`."""
    raw = _baca(b, sheet_name=sheet)
    for i in range(min(40, len(raw))):
        isi = {str(x).strip().upper() for x in raw.iloc[i] if pd.notna(x)}
        if all(w in isi for w in wajib):
            kolom = [str(x).strip() if pd.notna(x) else f"col_{j}" for j, x in enumerate(raw.iloc[i])]
            df = raw.iloc[i + 1:].copy()
            df.columns = kolom
            return df
    raise ValueError(f"Header {wajib} tidak ditemukan di sheet '{sheet}'.")


def _kolom_berisi(df, *kata):
    for c in df.columns:
        if any(w in str(c).upper() for w in kata):
            return c
    raise ValueError(f"Kolom {kata} tidak ditemukan.")


def _baris_closing(r, df, c_real, c_target, dasar):
    baris = dict(dasar)
    for c in df.columns:
        m = ALIAS_BULAN.get(str(c).strip().upper())
        if m:
            v = pd.to_numeric(r[c], errors="coerce")
            baris[KODE_BULAN[m - 1]] = float(v) if pd.notna(v) else np.nan
    real, target = angka(r[c_real]), angka(r[c_target])
    baris["Realisasi s.d Bulan"] = real
    baris["Target s.d Bulan"] = target
    baris["GAP"] = real - target
    baris["% Pencapaian"] = (real / target * 100) if target > 0 else np.nan
    return baris


@st.cache_data(show_spinner=False)
def parse_closing(b, sheet_ulp, sheet_up3):
    """Return (DataFrame rekap, nomor bulan terakhir yang berisi data atau None)."""
    df_u = _tabel_berheader(b, sheet_ulp, ["NAMA ULP"])
    df_p = _tabel_berheader(b, sheet_up3, ["KODE UNIT", "UP3"])
    kol_nama = [c for c in df_u.columns if c.strip().upper() == "NAMA ULP"][0]
    kol_up3 = [c for c in df_p.columns if c.strip().upper() == "UP3"][0]

    rows = []
    for ulp in TARGET_ULP:
        cocok = df_u[df_u[kol_nama].astype(str).str.strip().str.upper() == ulp["nama"]]
        if not cocok.empty:
            rows.append(_baris_closing(
                cocok.iloc[0], df_u, _kolom_berisi(df_u, "REALISASI"), _kolom_berisi(df_u, "TARGET"),
                {"No": len(rows) + 1, "ID ULP": ulp["id"], "Nama ULP": ulp["nama"].title()}))
    cocok = df_p[df_p[kol_up3].astype(str).str.strip().str.upper() == "MEDAN UTARA"]
    if not cocok.empty:
        rows.append(_baris_closing(
            cocok.iloc[0], df_p, _kolom_berisi(df_p, "REALISASI"), _kolom_berisi(df_p, "TARGET"),
            {"No": "", "ID ULP": ID_UP3, "Nama ULP": NAMA_UP3}))
    if not rows:
        raise ValueError(f"Tidak ada baris ULP/UP3 Medan Utara di sheet '{sheet_ulp}' / '{sheet_up3}'.")

    df = pd.DataFrame(rows)
    kol_bulan = [m for m in KODE_BULAN if m in df.columns]
    terakhir = None
    for i, m in enumerate(KODE_BULAN, start=1):
        if m in df.columns and (df[m].fillna(0) > 0).any():
            terakhir = i
    urut = ["No", "ID ULP", "Nama ULP"] + kol_bulan + ["Realisasi s.d Bulan", "Target s.d Bulan", "GAP", "% Pencapaian"]
    return df[urut], terakhir


def _nilai_blok(df, baris_akhir, col, tinggi=4):
    """
    Nilai target satu ULP. Sel target di-merge 4 baris (mis. baris 7-10 untuk 12801) dan
    Excel hanya menyimpan nilainya di baris teratas, jadi blok dibaca dari atas ke bawah.
    """
    for r in range(baris_akhir - tinggi, baris_akhir):
        if 0 <= r < len(df):
            v = pd.to_numeric(df.iat[r, col], errors="coerce")
            if pd.notna(v):
                return float(v)
    return 0.0


@st.cache_data(show_spinner=False)
def parse_kinerja(b):
    """
    File Master Kinerja ->
      historis   : {id_ulp: [12 nilai total PAL+TS, None bila kosong]}
      target_pel : {bulan(1..12): {id_ulp: target pelunasan}}
      kolom_pel  : {bulan: huruf kolom Excel yang dipakai}
      catatan    : daftar peringatan
    """
    hasil = {"historis": {}, "target_pel": {}, "kolom_pel": {}, "catatan": []}
    xl = pd.ExcelFile(io.BytesIO(b))
    nama_sheet = {s.strip().upper(): s for s in xl.sheet_names}

    # --- Riwayat saldo: sheet 'PIUTANG RATA-RATA TUNGG', baris 150-154, kolom I..T (Jan..Des)
    s_hist = next((v for n, v in nama_sheet.items() if n.startswith("PIUTANG RATA-RATA TUNGG")), None)
    if s_hist is None:
        hasil["catatan"].append("Sheet 'PIUTANG RATA-RATA TUNGG' tidak ditemukan; riwayat cadangan dipakai.")
    else:
        df = xl.parse(s_hist, header=None)
        for i, ulp in enumerate(TARGET_ULP):
            r = 149 + i
            nilai = []
            for c in range(8, 20):
                v = angka(df.iat[r, c]) if r < len(df) and c < df.shape[1] else 0.0
                nilai.append(v if v > 0 else None)
            hasil["historis"][ulp["id"]] = nilai

    # --- Target pelunasan: sheet 'PELUNASAN PRR', kolom berlabel 'T' per bulan (range B:BB)
    s_pel = nama_sheet.get("PELUNASAN PRR")
    if s_pel is None:
        hasil["catatan"].append("Sheet 'PELUNASAN PRR' tidak ditemukan; target pelunasan tidak tersedia.")
        return hasil
    df = xl.parse(s_pel, header=None)
    kol_akhir = min(54, df.shape[1])          # s.d. kolom BB
    n_header = min(9, len(df))                # header berada di atas baris 10
    kolom = range(1, kol_akhir)               # kolom B s.d. BB

    def label_t(v):
        return isinstance(v, str) and v.strip().upper() in ("T", "TARGET")

    # Baris sub-header = baris dengan label 'T' terbanyak
    jml_t = [sum(label_t(df.iat[r, c]) for c in kolom) for r in range(n_header)]
    brs_t = int(np.argmax(jml_t)) if jml_t and max(jml_t) > 0 else None
    kol_t = [c for c in kolom if label_t(df.iat[brs_t, c])] if brs_t is not None else []

    # Baris header bulan = baris (di atas/sejajar sub-header) dengan nama bulan terbanyak
    kol_bulan = {}
    if brs_t is not None:
        kandidat = []
        for r in range(brs_t + 1):
            peta = {}
            for c in kolom:
                m = bulan_dari_label(df.iat[r, c])
                if m and m not in peta:
                    peta[m] = c
            kandidat.append(peta)
        terbaik = max(kandidat, key=len)
        if len(terbaik) >= 2:
            kol_bulan = terbaik

    # Tiap kolom 'T' milik bulan yang header-nya terdekat di sebelah kiri (sel merge)
    t_per_bulan = {}
    for c in kol_t:
        kiri = [(cb, m) for m, cb in kol_bulan.items() if cb <= c]
        if kiri:
            t_per_bulan.setdefault(max(kiri)[1], c)

    for m in range(1, 13):
        col = t_per_bulan.get(m)
        if col is None and not kol_bulan and len(kol_t) >= m:   # cadangan: 'T' urutan ke-m
            col = kol_t[m - 1]
        if col is None:
            continue
        hasil["kolom_pel"][m] = get_column_letter(col + 1)
        hasil["target_pel"][m] = {
            uid: _nilai_blok(df, brs, col) for uid, brs in BARIS_TARGET_PELUNASAN.items()
        }
    if not hasil["target_pel"]:
        hasil["catatan"].append("Kolom berlabel 'T' tidak ditemukan pada sheet 'PELUNASAN PRR' (B:BB).")
    return hasil


@st.cache_data(show_spinner=False)
def parse_saldo_akhir(b):
    """File PAL/TS -> (lembar, rupiah) = baris 119 (berjalan) + baris 120 (tunggakan)."""
    df = _baca(b)

    def v(r, c):
        return angka(df.iat[r, c]) if r < len(df) and c < df.shape[1] else 0.0

    # Lembar: Umum I+J, Pemda X, BUMN AB  |  Rupiah: Umum K+L, Pemda Z, BUMN AD
    lembar = sum(v(r, c) for r in (118, 119) for c in (8, 9, 23, 27))
    rupiah = sum(v(r, c) for r in (118, 119) for c in (10, 11, 25, 29))
    return lembar, rupiah


@st.cache_data(show_spinner=False)
def parse_baris_21(b, kata_baris, kata_kolom, kolom_default):
    """Ambil nilai pada baris '2.1 <kata_baris>' di kolom yang header-nya memuat semua `kata_kolom`."""
    df = _baca(b)
    baris = None
    for i in range(len(df)):
        teks = " ".join(str(x).upper() for x in df.iloc[i] if pd.notna(x))
        if "2.1" in teks and kata_baris in teks:
            baris = i
            break
    if baris is None:
        baris = 3                              # posisi baku: baris 4
    if baris >= len(df):
        return 0.0
    kolom = None
    for c in range(df.shape[1]):
        teks = " ".join(str(x).upper() for x in df.iloc[:max(5, baris), c] if pd.notna(x))
        if all(w in teks for w in kata_kolom):
            kolom = c
            break
    if kolom is None:
        kolom = kolom_default
    return angka(df.iat[baris, kolom]) if kolom < df.shape[1] else 0.0


def parse_prr_tunai(b):
    return parse_baris_21(b, "PELUNASAN", ("GAB",), 5)             # F4


def parse_cicilan(b):
    return parse_baris_21(b, "CICILAN", ("RUPIAH", "TOTAL"), 4)    # E4


@st.cache_data(show_spinner=False)
def parse_eks_prr_tunai(b):
    """Jumlahkan kolom 'RP PRR' (file boleh kosong)."""
    df = _baca(b)
    for i in range(min(15, len(df))):
        for c in range(df.shape[1]):
            label = str(df.iat[i, c]).upper().replace("_", " ")
            if pd.notna(df.iat[i, c]) and "RP" in label and "PRR" in label:
                return float(pd.to_numeric(df.iloc[i + 1:, c], errors="coerce").sum())
    return 0.0


@st.cache_data(show_spinner=False)
def parse_ts_prabayar(b):
    """File nontaglis UP3: cari header UNIT UP & RP TAG (baris ~10), jumlahkan RP TAG per ULP."""
    df = _baca(b)
    for i in range(min(25, len(df))):
        label = [str(x).strip().upper().replace("_", " ") if pd.notna(x) else "" for x in df.iloc[i]]
        c_unit = next((j for j, s in enumerate(label) if "UNIT UP" in s), None)
        c_tag = next((j for j, s in enumerate(label) if "RP TAG" in s), None)
        if c_unit is not None and c_tag is not None:
            unit = df.iloc[i + 1:, c_unit].astype(str).str.strip().str.upper()
            tag = pd.to_numeric(df.iloc[i + 1:, c_tag], errors="coerce").fillna(0)
            return {
                u["id"]: float(tag[unit.str.contains(u["id"], regex=False)
                                   | unit.str.contains(u["nama"], regex=False)].sum())
                for u in TARGET_ULP
            }
    raise ValueError("Header 'UNIT UP' dan 'RP TAG' tidak ditemukan.")


ALIAS_ULP = {
    "12801": ["MEDANTIMUR", "MDNTIMUR", "TIMUR"],
    "12802": ["BELAWAN"],
    "12803": ["HELVETIA", "HELVET"],
    "12804": ["LABUHAN"],
    "12805": ["DENAI"],
}


@st.cache_data(show_spinner=False)
def deteksi_ulp(nama, b):
    """ID ULP pemilik file: dari nama file (nama/kode unit), bila gagal dari isi 25 baris pertama."""
    rapat = re.sub(r"[^A-Z0-9]", "", nama.upper())
    cocok = [uid for uid, alias in ALIAS_ULP.items() if uid in rapat or any(a in rapat for a in alias)]
    if len(cocok) == 1:
        return cocok[0]
    try:
        df = pd.read_excel(io.BytesIO(b), header=None, nrows=25)
    except Exception:  # noqa: BLE001
        return None
    teks = " ".join(str(x).upper() for x in df.to_numpy().ravel() if pd.notna(x))
    cocok = [u["id"] for u in TARGET_ULP
             if u["nama"] in teks or re.search(rf"(?<!\d){u['id']}", teks)]
    return cocok[0] if len(cocok) == 1 else None


# =========================================================================
# 4. PRE-COMPUTATION (dijalankan sebelum sidebar & tab dirender)
# =========================================================================
peringatan = []          # (tab, pesan) -> ditampilkan di tab terkait


def aman(tab, key, fungsi, default):
    """Jalankan parser untuk slot `key`; file rusak tidak menjatuhkan aplikasi."""
    b = bytes_di(key)
    if b is None:
        return default
    try:
        return fungsi(b)
    except Exception as e:  # noqa: BLE001
        peringatan.append((tab, f"**{file_di(key).name}** gagal dibaca: {e}"))
        return default


# --- 4.0 Petakan upload grup (5 file sekaligus) ke slot per-ULP
NAMA_KE_ID = {u["nama"]: u["id"] for u in TARGET_ULP}
for jenis in JENIS_PER_ULP:
    tab_asal = 4 if jenis in ("pal", "ts") else 5
    hasil_peta, PERLU_MANUAL[jenis] = {}, []
    for bk in st.session_state.get("_berkas_grp_" + jenis) or []:
        uid = deteksi_ulp(bk.name, bk.getvalue())
        if uid is None:
            PERLU_MANUAL[jenis].append(bk)
            uid = NAMA_KE_ID.get(st.session_state.get(f"pilih_{jenis}_{bk.name}"))
            if uid is None:
                peringatan.append((tab_asal, f"ULP untuk file **{bk.name}** tidak dikenali dari nama maupun "
                                             "isinya. Pilih ULP-nya pada kotak pilihan di bawah slot upload."))
                continue
        if uid in hasil_peta:
            peringatan.append((tab_asal, f"**{bk.name}** dan **{hasil_peta[uid].name}** sama-sama terbaca sebagai "
                                         f"ULP {uid}; yang dipakai **{bk.name}**."))
        hasil_peta[uid] = bk
    for u in TARGET_ULP:
        PETA[k(jenis, u)] = hasil_peta.get(u["id"])

# --- 4a. Tab 1: Data Closing
closing = {"rupiah": None, "kali": None, "bulan": None}
if bytes_di(K_CLOSING) is not None:
    df_r, bln_r = aman(1, K_CLOSING, lambda b: parse_closing(b, "Rupiah-ULP", "Rupiah-UP3"), (None, None))
    df_k, bln_k = aman(1, K_CLOSING, lambda b: parse_closing(b, "Kali-ULP", "Kali-UP3"), (None, None))
    closing = {"rupiah": df_r, "kali": df_k, "bulan": bln_r or bln_k}

# --- 4b. Tab 3: Master Kinerja
kinerja = aman(3, K_KINERJA, parse_kinerja, None)
if kinerja:
    peringatan += [(3, c) for c in kinerja["catatan"]]

# --- 4c. Tab 4: PAL & TS
saldo = {}
for u in TARGET_ULP:
    lbr_p, rp_p = aman(4, k("pal", u), parse_saldo_akhir, (0.0, 0.0))
    lbr_t, rp_t = aman(4, k("ts", u), parse_saldo_akhir, (0.0, 0.0))
    saldo[u["id"]] = {
        "lbr_p": lbr_p, "rp_p": rp_p, "lbr_t": lbr_t, "rp_t": rp_t,
        "tot_lbr": lbr_p + lbr_t, "tot_rp": rp_p + rp_t,
        "ada_file": file_di(k("pal", u)) is not None or file_di(k("ts", u)) is not None,
    }

# --- 4d. Tab 5: Pelunasan (21 slot)
ts_prabayar = aman(5, K_TS_PRABAYAR, parse_ts_prabayar, {})
pelunasan = {}
for u in TARGET_ULP:
    p = {
        "prr_tunai": aman(5, k("prr_tunai", u), parse_prr_tunai, 0.0),
        "prr_cicilan": aman(5, k("prr_cicilan", u), parse_cicilan, 0.0),
        "eks_tunai": aman(5, k("eks_tunai", u), parse_eks_prr_tunai, 0.0),
        "eks_cicilan": aman(5, k("eks_cicilan", u), parse_cicilan, 0.0),
        "ts": ts_prabayar.get(u["id"], 0.0),
    }
    p["prr"] = p["prr_tunai"] + p["prr_cicilan"]
    p["eks"] = p["eks_tunai"] + p["eks_cicilan"]
    p["total"] = p["prr"] + p["eks"] + p["ts"]
    pelunasan[u["id"]] = p

# --- 4e. Deteksi bulan otomatis dari file yang diunggah
semua_key = [K_CLOSING] + [k(j, u) for j in JENIS_PER_ULP for u in TARGET_ULP] + [K_TS_PRABAYAR, K_KINERJA]
bulan_terdeteksi, sumber_deteksi = closing["bulan"], "isi file Data Closing"
if bulan_terdeteksi is None:
    for key in semua_key:
        f = file_di(key)
        if f is not None:
            m = bulan_dari_nama_file(f.name)
            if m:
                bulan_terdeteksi, sumber_deteksi = m, f"nama file {f.name}"
                break

# Hasil deteksi baru menimpa pilihan; setelah itu pengguna bebas mengubah manual.
if "bulan_eval" not in st.session_state:
    st.session_state["bulan_eval"] = BULAN[(bulan_terdeteksi or dt.date.today().month) - 1]
tanda = (bulan_terdeteksi, sumber_deteksi) if bulan_terdeteksi else None
if tanda and st.session_state.get("_tanda_deteksi") != tanda:
    st.session_state["bulan_eval"] = BULAN[bulan_terdeteksi - 1]
st.session_state["_tanda_deteksi"] = tanda

# --- 4f. Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Pengaturan Laporan")
    nama_bulan = st.selectbox("Pilih Bulan Evaluasi Laporan", BULAN, key="bulan_eval")
    if bulan_terdeteksi:
        st.caption(f"🔎 Terdeteksi otomatis: **{BULAN[bulan_terdeteksi - 1]}** (dari {sumber_deteksi}).")
    else:
        st.caption("🔎 Bulan belum terdeteksi dari file; pilih manual.")

    st.caption(f"Versi aplikasi: {VERSI_APP}")
    st.markdown("### 📦 Status Upload")
    n_tab4 = sum(file_di(k(j, u)) is not None for j in ("pal", "ts") for u in TARGET_ULP)
    n_tab5 = sum(file_di(k(j, u)) is not None for j in JENIS_PER_ULP[2:] for u in TARGET_ULP) \
        + (file_di(K_TS_PRABAYAR) is not None)
    for label, n, total in [
        ("Tab 1 · Data Closing", int(file_di(K_CLOSING) is not None), 1),
        ("Tab 3 · Master Kinerja", int(file_di(K_KINERJA) is not None), 1),
        ("Tab 4 · PAL & TS", n_tab4, 10),
        ("Tab 5 · Pelunasan", n_tab5, 21),
    ]:
        ikon = "🟢" if n == total else ("🟡" if n else "🔴")
        st.markdown(f"{ikon} {label}: **{n}/{total}**")

BLN = BULAN.index(nama_bulan) + 1           # nomor bulan aktif (1..12)
PERIODE = f"{nama_bulan} {TAHUN}"

# --- 4g. Kalkulasi yang bergantung pada bulan aktif
target_pel = (kinerja or {}).get("target_pel", {}).get(BLN, {})
kolom_pel = (kinerja or {}).get("kolom_pel", {}).get(BLN)


def rata_rata_saldo(uid):
    """Rata-rata saldo bulan 1 s.d. bulan aktif. 0 selama PAL/TS ULP tsb belum diunggah."""
    d = saldo[uid]
    if not d["ada_file"]:
        return 0.0
    riwayat = (kinerja or {}).get("historis", {}).get(uid)
    if not riwayat or not any(riwayat):
        riwayat = HISTORIS_CADANGAN.get(uid, [])
    riwayat = list(riwayat) + [None] * (12 - len(riwayat))
    nilai = [x for x in riwayat[:BLN - 1] if x]               # bulan 1 .. bulan-1
    berjalan = d["tot_rp"] if d["tot_rp"] > 0 else riwayat[BLN - 1]
    if berjalan:
        nilai.append(berjalan)                                # bulan berjalan
    return float(np.mean(nilai)) if nilai else 0.0


def pct_realisasi(pct_capaian):
    """IF(%Pencapaian >= 110%, 110% x Bobot, Bobot x %Pencapaian)"""
    return 1.10 * BOBOT if pct_capaian >= 1.10 else BOBOT * pct_capaian


def evaluasi(nama, t_saldo, r_saldo, t_pel, r_pel):
    cap_prr = 1.0                                                           # Target 100, Realisasi 100
    cap_saldo = 2.0 - (r_saldo / t_saldo) if (r_saldo > 0 and t_saldo > 0) else 1.0
    cap_pel = (r_pel / t_pel) if t_pel > 0 else 1.0
    real = [pct_realisasi(c) for c in (cap_prr, cap_saldo, cap_pel)]
    return {
        "nama": nama, "t_saldo": t_saldo, "r_saldo": r_saldo, "t_pel": t_pel, "r_pel": r_pel,
        "cap_prr": cap_prr, "cap_saldo": cap_saldo, "cap_pel": cap_pel,
        "real_prr": real[0], "real_saldo": real[1], "real_pel": real[2],
        "tot_real": sum(real), "pencapaian": sum(real) / (3 * BOBOT) * 100.0, "rank": "-",
    }


rata = {u["id"]: rata_rata_saldo(u["id"]) for u in TARGET_ULP}

cash_in = [
    evaluasi(u["nama"], u["t_total"], rata[u["id"]], target_pel.get(u["id"], 0.0), pelunasan[u["id"]]["total"])
    for u in TARGET_ULP
]
for peringkat, e in enumerate(sorted(cash_in, key=lambda e: e["pencapaian"], reverse=True), start=1):
    e["rank"] = peringkat
cash_in.append(evaluasi(
    "MEDAN UTARA",
    sum(u["t_total"] for u in TARGET_ULP), sum(rata.values()),
    sum(target_pel.values()), sum(p["total"] for p in pelunasan.values()),
))


def tampil_peringatan(tab):
    for t, pesan in peringatan:
        if t == tab:
            st.warning(pesan)


# =========================================================================
# 5. RENDER
# =========================================================================
st.markdown(f"""
<div class="hero-banner">
  <div class="hero-title">⚡ SISTEM MONITORING KEUANGAN & KINERJA NIAGA</div>
  <div class="hero-subtitle">PT PLN (Persero) UP3 Medan Utara — Transaksi PLN Mobile & Percepatan Cash In · Periode evaluasi: <b>{PERIODE}</b></div>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📱 MONITORING TRANSAKSI PLN MOBILE",
    "⚡ PERCEPATAN CASH IN",
    "📋 PENGHAPUSAN PRR",
    "📊 SALDO RATA-RATA (PAL & TS)",
    "💳 PELUNASAN",
])

# ------------------------------------------------------------------ TAB 1
with tab1:
    st.markdown("### 📂 Upload File Data Closing Bulanan")
    st.caption("Unggah file Excel closing bulanan (*Data Closing Transaksi s.d...*) berisi sheet "
               "`Rupiah-ULP`, `Rupiah-UP3`, `Kali-ULP`, dan `Kali-UP3`.")
    custom_file_uploader("File Data Closing Transaksi", K_CLOSING)
    tampil_peringatan(1)

    df_rupiah, df_kali = closing["rupiah"], closing["kali"]
    if df_rupiah is None and df_kali is None:
        st.info("💡 Unggah file Data Closing untuk memunculkan KPI dan tabel rekapitulasi.")
    else:
        sd = KODE_BULAN[(closing["bulan"] or BLN) - 1]

        def baris_up3(df):
            if df is None:
                return None
            r = df[df["ID ULP"] == ID_UP3]
            return r.iloc[0] if not r.empty else None

        up_r, up_k = baris_up3(df_rupiah), baris_up3(df_kali)
        kartu = [
            (f"Total Realisasi Rupiah (s.d {sd})", fmt_angka(up_r["Realisasi s.d Bulan"]) if up_r is not None else "-"),
            (f"Pencapaian Rupiah s.d {sd}", fmt_persen(up_r["% Pencapaian"]) if up_r is not None else "-"),
            (f"Total Frekuensi Transaksi (s.d {sd})",
             fmt_angka(up_k["Realisasi s.d Bulan"], rupiah=False) + " Kali" if up_k is not None else "-"),
            (f"Pencapaian Frekuensi s.d {sd}", fmt_persen(up_k["% Pencapaian"]) if up_k is not None else "-"),
        ]
        st.markdown('<div class="kpi-container">' + "".join(
            f'<div class="kpi-card"><div class="kpi-label">{a}</div><div class="kpi-value">{b}</div></div>'
            for a, b in kartu) + "</div>", unsafe_allow_html=True)

        kol_uang = KODE_BULAN + ["Realisasi s.d Bulan", "Target s.d Bulan", "GAP"]
        if df_rupiah is not None:
            st.markdown(f"#### 1. Realisasi Rupiah Transaksi Keuangan (s.d. {sd} {TAHUN})")
            tampil_df(format_tabel(df_rupiah, kol_uang, ["% Pencapaian"], rupiah=True))
        if df_kali is not None:
            st.markdown(f"#### 2. Realisasi Frekuensi Transaksi (s.d. {sd} {TAHUN})")
            tampil_df(format_tabel(df_kali, kol_uang, ["% Pencapaian"], rupiah=False))

# ------------------------------------------------------------------ TAB 2
with tab2:
    st.markdown(f"### ⚡ TABEL PERCEPATAN CASH IN — {PERIODE.upper()}")

    def td(isi, kelas="r"):
        return f'<td class="{kelas}">{isi}</td>'

    baris_html = ""
    for e in cash_in:
        baris_html += (
            f'<tr><td rowspan="5" class="cashin-unit">{e["nama"]}</td>'
            + td("Penghapusan PRR", "") + td("2", "c") + td("100") + td("100")
            + td(fmt_persen(e["cap_prr"] * 100)) + td(fmt_desimal(e["real_prr"]))
            + f'<td rowspan="5" class="cashin-rank">{e["rank"]}</td></tr>'
            + "<tr>" + td("Saldo Rata-rata Pascabayar Non Kogol 1", "") + td("2", "c")
            + td(sel_rp(fmt_angka(e["t_saldo"]))) + td(sel_rp(fmt_angka(e["r_saldo"])))
            + td(fmt_persen(e["cap_saldo"] * 100)) + td(fmt_desimal(e["real_saldo"])) + "</tr>"
            + "<tr>" + td("Pelunasan PRR, Eks PRR, dan TS Prabayar", "") + td("2", "c")
            + td(sel_rp(fmt_angka(e["t_pel"]))) + td(sel_rp(fmt_angka(e["r_pel"])))
            + td(fmt_persen(e["cap_pel"] * 100)) + td(fmt_desimal(e["real_pel"])) + "</tr>"
            + '<tr class="cashin-total-row">' + td("TOTAL", "c") + td("6", "c") + '<td colspan="3"></td>'
            + f'<td class="r" style="color:#0284c7;">{fmt_desimal(e["tot_real"], 3)}</td></tr>'
            + '<tr class="cashin-pct-row">' + td("PENCAPAIAN %", "c") + td("1", "c")
            + f'<td colspan="4" class="c" style="color:#0369a1;font-size:13px;">{fmt_persen(e["pencapaian"])}</td></tr>'
        )

    st.markdown(
        '<table class="cashin-table"><thead>'
        '<tr><th rowspan="2" style="width:140px;">UNIT</th><th rowspan="2">URAIAN</th>'
        '<th rowspan="2" style="width:70px;">BOBOT</th>'
        f'<th colspan="4">BULAN EVALUASI: {nama_bulan.upper()}</th>'
        '<th rowspan="2" style="width:90px;">PERINGKAT</th></tr>'
        '<tr><th>TARGET</th><th>REALISASI</th><th>% PENCAPAIAN</th><th>% REALISASI</th></tr>'
        f'</thead><tbody>{baris_html}</tbody></table>',
        unsafe_allow_html=True,
    )

    if kinerja is not None and target_pel and sum(target_pel.values()) > 0:
        st.caption(f"🎯 Target pelunasan {nama_bulan}: sheet `PELUNASAN PRR` kolom **{kolom_pel}** "
                   f"dari {file_di(K_KINERJA).name}.")
    elif kinerja is not None and target_pel:
        st.caption(f"🎯 File master terbaca, tetapi target pelunasan {nama_bulan} di kolom **{kolom_pel}** bernilai 0.")
    belum = []
    if not target_pel:
        belum.append("target pelunasan (File Master Kinerja, Tab 3)")
    if not any(d["ada_file"] for d in saldo.values()):
        belum.append("saldo rata-rata (file PAL/TS, Tab 4)")
    if belum:
        st.caption("ℹ️ Belum tersedia: " + " dan ".join(belum)
                   + ". Indikator terkait sementara dihitung 100% sampai datanya diunggah.")

# ------------------------------------------------------------------ TAB 3
with tab3:
    st.markdown("### 📂 Unggah File Master Target / Kinerja")
    st.caption("Unggah **082026.KINERJA_ADMNIAGA_CUEX_TARGET_2026 revisi.xlsx** sebagai acuan target "
               "pelunasan dan riwayat saldo bulanan.")
    custom_file_uploader("File Master Kinerja Tahunan", K_KINERJA)
    tampil_peringatan(3)
    if kinerja is None:
        st.info("💡 Unggah File Master Kinerja untuk menyinkronkan target pelunasan dan riwayat saldo tunggakan.")
    else:
        st.success("✅ File Master Kinerja terhubung.")

    st.markdown("---")
    st.markdown(f"### 🎯 Target Pelunasan PRR — {PERIODE}")
    if target_pel:
        st.caption(f"Sumber: sheet `PELUNASAN PRR`, kolom **{kolom_pel}** (label `T` bulan {nama_bulan}), "
                   "blok baris 7–10 / 11–14 / 15–18 / 19–22 / 23–26 (sel merge).")
        df_tp = pd.DataFrame(
            [{"Kode Unit": u["id"], "Nama Unit": u["nama"].title(),
              "Baris Excel": f"{BARIS_TARGET_PELUNASAN[u['id']] - 3}–{BARIS_TARGET_PELUNASAN[u['id']]}", "Target Pelunasan": target_pel.get(u["id"], 0.0)}
             for u in TARGET_ULP]
            + [{"Kode Unit": ID_UP3, "Nama Unit": NAMA_UP3, "Baris Excel": "", "Target Pelunasan": sum(target_pel.values())}]
        )
        df_tp["Baris Excel"] = df_tp["Baris Excel"].astype(str)
        tampil_df(format_tabel(df_tp, ["Target Pelunasan"]))
    elif kinerja is not None:
        st.warning(f"Target bulan {nama_bulan} tidak ditemukan pada sheet `PELUNASAN PRR`.")
    else:
        st.caption("Belum ada data — menunggu File Master Kinerja.")

    st.markdown("---")
    st.markdown("### 📋 TABEL TARGET USULAN PENGHAPUSAN PRR")
    st.caption("Format mengikuti sheet `USULAN PENGHAPUSAN PRR`. Realisasi masih placeholder 100%.")
    usulan = [("12801", "Medan Timur", "30%", 972_000_000), ("12802", "Belawan", "16%", 485_000_000),
              ("12803", "Helvetia", "10%", 387_000_000), ("12804", "Labuhan", "14%", 415_000_000),
              ("12805", "Denai", "30%", 905_000_000)]
    df_usulan = pd.DataFrame(
        [{"No": str(i), "Kode Unit": kd, "Nama Unit": nm, "Proporsional": pr, f"Target {TAHUN}": tg,
          "Realisasi Plgn": 100, "Realisasi Rupiah": tg, "GAP": 0, "% Pencapaian": 100.0}
         for i, (kd, nm, pr, tg) in enumerate(usulan, start=1)]
        + [{"No": "", "Kode Unit": ID_UP3, "Nama Unit": NAMA_UP3, "Proporsional": "100%",
            f"Target {TAHUN}": sum(x[3] for x in usulan), "Realisasi Plgn": 500,
            "Realisasi Rupiah": sum(x[3] for x in usulan), "GAP": 0, "% Pencapaian": 100.0}]
    )
    tampil_df(format_tabel(df_usulan, [f"Target {TAHUN}", "Realisasi Rupiah", "GAP"], ["% Pencapaian"]))

# ------------------------------------------------------------------ TAB 4
with tab4:
    st.markdown(f"### 📂 Upload File PAL & TS — {PERIODE} (10 File)")
    st.caption("Pilih atau seret **kelima file ULP sekaligus** ke tiap slot; sistem mengenali ULP dari nama "
               "file (nama/kode unit) atau isinya. Sistem mengambil Baris 119 (*Rekening Berjalan*) dan Baris 120 "
               "(*Rekening Tunggakan*) untuk kolom Umum, Pemda, dan BUMN.")
    kiri, kanan = st.columns(2)
    with kiri:
        st.markdown("##### 📁 File PAL (5 ULP)")
        uploader_grup("File PAL", "pal")
    with kanan:
        st.markdown("##### 📁 File TS / Tagihan Susulan (5 ULP)")
        uploader_grup("File TS", "ts")
    tampil_peringatan(4)

    with st.expander(f"📌 TABEL TARGET RESMI (PAL, TS & RATA-RATA TUNGGAKAN {TAHUN})"):
        for kol, judul, f in zip(st.columns(3),
                                 ["1. Target PAL Bulanan", "2. Target TS Bulanan", "3. Target Rata-Rata Tunggakan"],
                                 ["t_pal", "t_ts", "t_total"]):
            with kol:
                st.markdown(f"##### {judul}")
                df_t = pd.DataFrame(
                    [{"Kode": u["id"], "Nama Unit": u["nama"].title(), "Target": u[f]} for u in TARGET_ULP]
                    + [{"Kode": ID_UP3, "Nama Unit": NAMA_UP3, "Target": sum(u[f] for u in TARGET_ULP)}])
                tampil_df(format_tabel(df_t, ["Target"]))

    def tabel_saldo(satuan, f_pal, f_ts, f_tot, f_rata, target):
        rows = [{"No": str(i), "ID ULP": u["id"], "Nama Unit": u["nama"].title(), "Target": target(u),
                 f"Saldo PAL ({satuan})": saldo[u["id"]][f_pal], f"Saldo TS ({satuan})": saldo[u["id"]][f_ts],
                 f"Total Saldo ({satuan})": saldo[u["id"]][f_tot], f"Rata-Rata ({satuan})": f_rata(u["id"])}
                for i, u in enumerate(TARGET_ULP, start=1)]
        df = pd.DataFrame(rows)
        total = {"No": "", "ID ULP": ID_UP3, "Nama Unit": NAMA_UP3}
        total.update({c: df[c].sum() for c in df.columns[3:]})
        return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    st.markdown("#### 1. Tabel Rekapitulasi Saldo Rupiah (PAL + TS)")
    df_rp = tabel_saldo("Rp", "rp_p", "rp_t", "tot_rp", lambda uid: rata[uid], lambda u: u["t_total"])
    tampil_df(format_tabel(df_rp, df_rp.columns[3:], rupiah=True))
    sumber = "File Master Kinerja" if (kinerja and kinerja["historis"]) else "riwayat cadangan (Jan–Jul 2026)"
    st.caption(f"Rata-rata = rata-rata saldo Januari s.d. {nama_bulan}; bulan sebelumnya dari {sumber}, "
               "bulan berjalan dari file yang diunggah. Bernilai “-” sampai file PAL/TS ULP tersebut diunggah.")

    st.markdown("#### 2. Tabel Rekapitulasi Lembar Rekening (PAL + TS)")
    df_lbr = tabel_saldo("Lbr", "lbr_p", "lbr_t", "tot_lbr",
                         lambda uid: saldo[uid]["tot_lbr"] if saldo[uid]["ada_file"] else 0.0, lambda u: 0)
    tampil_df(format_tabel(df_lbr, df_lbr.columns[3:], rupiah=False))

# ------------------------------------------------------------------ TAB 5
with tab5:
    st.markdown(f"### 💳 PELUNASAN TUNGGAKAN — {PERIODE}")
    st.caption("21 file: 10 file PRR, 10 file Eks PRR, 1 file TS Prabayar. Tiap slot ULP menerima 5 file sekaligus.")
    tampil_peringatan(5)
    sub1, sub2, sub3 = st.tabs(["1. PELUNASAN PRR", "2. PELUNASAN EKS PRR", "3. PELUNASAN TS Prabayar"])

    def rekap(kolom):
        """kolom: {judul kolom: field pada dict pelunasan}; otomatis menambah baris UP3."""
        df = pd.DataFrame([{"ID ULP": u["id"], "Nama Unit": u["nama"].title(),
                            **{j: pelunasan[u["id"]][f] for j, f in kolom.items()}} for u in TARGET_ULP])
        total = {"ID ULP": ID_UP3, "Nama Unit": NAMA_UP3, **{j: df[j].sum() for j in kolom}}
        df = pd.concat([df, pd.DataFrame([total])], ignore_index=True)
        tampil_df(format_tabel(df, list(kolom)))

    def dua_kolom_upload(label, jenis_tunai, jenis_cicilan):
        a, b = st.columns(2)
        with a:
            st.markdown(f"##### 📁 File TUNAI ({label}, 5 ULP)")
            uploader_grup(f"{label} Tunai", jenis_tunai)
        with b:
            st.markdown(f"##### 📁 File CICILAN ({label}, 5 ULP)")
            uploader_grup(f"{label} Cicilan", jenis_cicilan)

    with sub1:
        st.markdown("#### 📂 Upload Data PELUNASAN PRR (10 File)")
        st.caption("**TUNAI**: baris `2.1 PELUNASAN`, kolom `GAB` (F4). "
                   "**CICILAN**: baris `2.1 Pelunasan Cicilan PRR`, kolom `RUPIAH Total` (E4).")
        dua_kolom_upload("PRR", "prr_tunai", "prr_cicilan")
        st.markdown("##### Rekap Pelunasan PRR")
        rekap({"Pelunasan Tunai (Rp)": "prr_tunai", "Pelunasan Cicilan (Rp)": "prr_cicilan", "Total PRR (Rp)": "prr"})

    with sub2:
        st.markdown("#### 📂 Upload Data PELUNASAN EKS PRR (10 File)")
        st.caption("**TUNAI**: jumlah kolom `RP PRR` (bila berisi data). **CICILAN**: posisi sama dengan cicilan PRR.")
        dua_kolom_upload("Eks PRR", "eks_tunai", "eks_cicilan")
        st.markdown("##### Rekap Pelunasan Eks PRR")
        rekap({"Eks PRR Tunai (Rp)": "eks_tunai", "Eks PRR Cicilan (Rp)": "eks_cicilan", "Total Eks PRR (Rp)": "eks"})

    with sub3:
        st.markdown("#### 📂 Upload Data PELUNASAN TS Prabayar (1 File)")
        st.caption("File nontaglis UP3 (*DETAIL Monitoring Pelunasan Piutang - NO REGISTER...*). "
                   "Sistem mencari header `UNIT UP` dan `RP TAG`, lalu menjumlahkan `RP TAG` per ULP.")
        custom_file_uploader("File TS Prabayar (UP3)", K_TS_PRABAYAR)
        st.markdown("##### Rekap Pelunasan TS Prabayar")
        rekap({"Total TS Prabayar (Rp)": "ts"})

    st.markdown("---")
    st.markdown("#### 📊 Ringkasan Total Pelunasan")
    rekap({"Pelunasan PRR (Rp)": "prr", "Pelunasan Eks PRR (Rp)": "eks",
           "Pelunasan TS Prabayar (Rp)": "ts", "TOTAL PELUNASAN (Rp)": "total"})

# Ada slot upload yang berubah pada run ini -> jalankan ulang supaya pre-computation memakainya.
if perlu_rerun:
    st.rerun()
