import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import warnings
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
import io
import time
import requests
import json
from openai import OpenAI

warnings.filterwarnings('ignore')

# 1. KONFIGURASI FILE & PATH
FILE_MASTER_REF         = "cost & bbm 2022 sd 2025 HP & Type.xlsx"


# ⚙️  KONFIGURASI TIM IT
# 1) OPENAI_API_KEY
# Isi API key OpenAI di bawah ini
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "xxxxxx")

# 2) PATH PENYIMPANAN FILE "Konsumsi BBM Standar Pabrik" DI SERVER
# File hasil crawling akan disimpan & dibaca dari path ini.
PATH_STANDAR_PABRIK_SERVER = "C:/Users/asus/Downloads/SPIL/bbm/Analisa BBM/Konsumsi_BBM_Standar_Pabrik.xlsx"

if "fcst_hm_processed" not in st.session_state:
    st.session_state.fcst_hm_processed      = False
    st.session_state.fcst_hm_df_komparasi   = None
    st.session_state.fcst_hm_df_metrik      = None
    st.session_state.fcst_hm_df_excl        = None
    st.session_state.fcst_hm_df_monthly_raw = None
    st.session_state.fcst_hm_out_file       = None

# 2. SETUP HALAMAN
st.set_page_config(page_title="Dashboard Efisiensi BBM", layout="wide")
st.title("Dashboard BBM Alat Berat")

# 6. SIDEBAR & MENU NAVIGASI
category_filter = "Forecast Data"

st.sidebar.markdown("---")

# BAGIAN A: MENU FORECASTING (ARIMA + GRADIENT BOOSTING + ENSEMBLE)
if category_filter == "Forecast Data":
    st.header("Forecast Hour Meter & Kebutuhan BBM")

    # PANDUAN FORMAT & TEMPLATE
    with st.expander("📋 Panduan Format & Template File", expanded=False):
        st.markdown("#### 🔗 Template File (klik untuk membuka di Google Sheets)")
        tl1, tl2 = st.columns(2)
        with tl1:
            st.link_button(
                "📄 Template Detail Alat Berat",
                url="https://docs.google.com/spreadsheets/d/1BQIn_Ju51Y7Cmr-rhV1_jS0vhEgqn2BWLrBwrXy-JMs/edit?usp=sharing",
                use_container_width=True
            )
        with tl2:
            st.link_button(
                "📄 Template Data Train / Test",
                url="https://docs.google.com/spreadsheets/d/1AUIM5Fs36cB04ELIHoSGkobqWiKEn-3R1pEAaXlnAGw/edit?usp=sharing",
                use_container_width=True
            )

        st.markdown("---")
        st.markdown("#### 📌 Ketentuan dan Aturan Upload File")

        st.markdown("**Data Train & Data Test**")
        st.markdown(
            "- Format file wajib **`.xlsx`** (bukan `.xls` atau `.csv`).\n"
            "- Boleh upload **lebih dari satu file** untuk masing-masing Data Train dan Data Test, "
            "semua file akan digabungkan secara otomatis.\n"
            "- **Setiap file wajib mewakili tepat 1 tahun penuh data (Januari s.d. Desember)**, "
            "tidak boleh lebih dan tidak boleh kurang dari itu agar data dapat terbaca dengan benar.\n"
            "- **Setiap file wajib memiliki tepat 12 sheet** dengan nama persis: "
            "**`JAN, FEB, MAR, APR, MEI, JUN, JUL, AGT, SEP, OKT, NOV, DES`**. "
            "Sheet yang namanya tidak sesuai atau jumlahnya kurang/lebih dari 12 tidak akan terbaca.\n"
            "- Struktur header di setiap sheet wajib terdiri dari **3 baris**: "
            "baris pertama berisi **nama unit** (sebagai merged cell horizontal), "
            "baris kedua berisi group/kategori (tidak digunakan, boleh diisi bebas), "
            "dan baris ketiga berisi label kolom: **`TANGGAL`**, **`HM`**, **`LITER`**, dan kolom lain "
            "seperti `KEGIATAN`, `SUMBER FUEL`, `PHOTOS` yang akan diabaikan.\n"
            "- Kolom **`TANGGAL`** harus berada di **kolom paling kiri** setiap sheet.\n"
            "- Nilai di kolom `TANGGAL` harus berformat **tanggal Excel yang proper** "
            "(bukan teks). Jika tanggal diketik sebagai teks, data bulan tersebut tidak akan terbaca.\n"
            "- **Nama unit di file Data Train/Test harus sama persis** (termasuk huruf kapital dan spasi) "
            "dengan nama yang ada di file Detail Alat Berat. Ketidakcocokan nama akan menyebabkan "
            "unit masuk ke daftar yang di-exclude."
        )

        st.markdown("**Detail Alat Berat**")
        st.markdown(
            "- Format file wajib **`.xlsx`**.\n"
            "- Hanya upload **1 file**.\n"
            "- File wajib memiliki sheet bernama tepat **`Sheet2`**, penamaan lain seperti "
            "`Sheet 2` atau `Data` tidak akan terbaca.\n"
            "- Header kolom harus berada di **baris kedua** sheet tersebut (baris pertama "
            "adalah baris grup/judul yang bisa diisi bebas).\n"
            "- Kolom yang wajib ada: **`NAMA ALAT BERAT`**, **`ALAT BERAT`**, "
            "**`TYPE/MERK`**, **`CAP`**, **`HP`**. Kolom lain tidak berpengaruh.\n"
            "- Nilai di kolom `CAP` dan `HP` harus berupa **angka**, bukan teks."
        )

        st.markdown("**Konsumsi BBM Standar Pabrik**")
        st.markdown(
            "- File ini **tidak perlu diupload secara manual**. Data standar konsumsi BBM pabrik "
            "akan didapatkan secara otomatis melalui sistem, tergantung pilihan proses yang dipilih:\n"
            "  - **Crawl Standar Pabrik + Forecast**: mengambil ulang data standar terbaru "
            "berdasarkan file Detail Alat Berat yang diupload, lalu otomatis melanjutkan ke proses forecast.\n"
            "  - **Forecast Langsung**: menggunakan hasil crawling standar pabrik yang tersimpan "
            "paling akhir di server, tanpa mengambil ulang data baru.\n"
            "- Jika belum pernah ada hasil crawling tersimpan di server, kolom standar pabrik pada "
            "hasil forecast akan dikosongkan dan perbandingan vs standar pabrik tidak akan tersedia."
        )

    # UPLOAD FILE DATA
    st.markdown("### 📂 Upload File Data")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        f_train_files = st.file_uploader(
            "1. Data Train (boleh lebih dari 1 file)",
            type=["xlsx"], accept_multiple_files=True, key="fcst_train"
        )
        f_master = st.file_uploader(
            "3. Detail Alat Berat (1 file)",
            type=["xlsx"], accept_multiple_files=False, key="fcst_master"
        )
    with col_u2:
        f_test_files = st.file_uploader(
            "2. Data Test (boleh lebih dari 1 file)",
            type=["xlsx"], accept_multiple_files=True, key="fcst_test"
        )

    col_batas1, col_batas2 = st.columns(2)
    with col_batas1:
        batas_train_akhir = st.text_input(
            "Akhir periode Data Train (format: YYYY-MM)",
            value="2024-12",
            help="Semua data sampai bulan ini akan dijadikan data train."
        )
    with col_batas2:
        batas_test_awal = st.text_input(
            "Awal periode Data Test (format: YYYY-MM)",
            value="2025-01",
            help="Semua data mulai bulan ini akan dijadikan data test."
        )

    st.markdown("### ⚙️ Pilihan Proses")
    mode_proses = st.radio(
        "Pilih jenis proses yang ingin dijalankan:",
        ["🌐 Crawl Standar Pabrik + Forecast", "🚀 Forecast Langsung (pakai standar pabrik terakhir)"],
        help=(
            "**Crawl Standar Pabrik + Forecast**: mengambil ulang data standar konsumsi BBM pabrik "
            "berdasarkan file Detail Alat Berat yang diupload, lalu otomatis melanjutkan ke proses forecast.\n\n"
            "**Forecast Langsung**: melewati proses crawl dan langsung menjalankan forecast menggunakan "
            "hasil crawling standar pabrik yang tersimpan paling akhir di server."
        )
    )
    is_mode_crawl = mode_proses.startswith("🌐")

    if st.button("🚀 Jalankan Proses Forecast HM"):
        if f_train_files and f_test_files and f_master:
            with st.spinner("Memuat library forecasting dan melatih model (Proses ini bisa memakan 3-5 menit)"):
                try:
                    import pmdarima as pm
                    from sklearn.ensemble import HistGradientBoostingRegressor
                    from sklearn.metrics import mean_squared_error

                    # FUNGSI LOAD DATA
                    def load_and_melt_excel_fcst(file_obj, target_sheets=None):
                        all_data = []
                        xls = pd.ExcelFile(file_obj)
                        for sheet_name in xls.sheet_names:
                            if target_sheets is not None and sheet_name not in target_sheets:
                                continue
                            try:
                                df = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1, 2])
                                df = df.set_index(df.columns[0])
                                df.index.name = 'TANGGAL'
                                df.columns    = df.columns.droplevel(1)
                                df_stacked    = df.stack(level=0).reset_index()
                                df_stacked.rename(columns={'level_1': 'EQUIP NAME'}, inplace=True)
                                all_data.append(df_stacked)
                            except Exception:
                                continue
                        if not all_data:
                            return pd.DataFrame()
                        df_final = pd.concat(all_data, ignore_index=True)
                        df_final['TANGGAL'] = pd.to_datetime(df_final['TANGGAL'], dayfirst=True, errors='coerce')
                        df_final = df_final.dropna(subset=['TANGGAL'])
                        return df_final

                    # Baca semua file train dan test
                    df_train_list = []
                    for f in f_train_files:
                        f.seek(0)
                        df_train_list.append(load_and_melt_excel_fcst(f, target_sheets=None))

                    df_test_list = []
                    for f in f_test_files:
                        f.seek(0)
                        df_test_list.append(load_and_melt_excel_fcst(f, target_sheets=None))

                    df_all = pd.concat(df_train_list + df_test_list, ignore_index=True)
                    df_all = df_all.sort_values(['EQUIP NAME', 'TANGGAL'])

                    df_all['HM_Clean']   = pd.to_numeric(df_all['HM'], errors='coerce').replace(0, np.nan)
                    df_all['HM_Clean']   = df_all.groupby('EQUIP NAME')['HM_Clean'].ffill().fillna(0)
                    df_all['Delta_HM']   = df_all.groupby('EQUIP NAME')['HM_Clean'].diff().fillna(0)
                    df_all.loc[df_all['Delta_HM'] < 0,   'Delta_HM'] = 0
                    df_all.loc[df_all['Delta_HM'] > 744, 'Delta_HM'] = 0  # max 31 hari x 24 jam
                    df_all['LITER_Clean'] = pd.to_numeric(df_all['LITER'], errors='coerce').fillna(0)

                    df_all['TAHUN_BULAN'] = df_all['TANGGAL'].dt.to_period('M')
                    agg_data = df_all.groupby(['EQUIP NAME', 'TAHUN_BULAN']).agg(
                        {'Delta_HM': 'sum', 'LITER_Clean': 'sum'}
                    ).reset_index()
                    agg_data.rename(columns={'Delta_HM': 'HM', 'LITER_Clean': 'LITER'}, inplace=True)

                    agg_data_str = agg_data.copy()
                    agg_data_str['TAHUN_BULAN'] = agg_data_str['TAHUN_BULAN'].astype(str)
                    st.session_state.fcst_hm_df_monthly_raw = agg_data_str

                    train_agg = agg_data[agg_data['TAHUN_BULAN'] <= batas_train_akhir]
                    test_agg = agg_data[
                        (agg_data['TAHUN_BULAN'] >= batas_test_awal) &
                        (agg_data['TAHUN_BULAN'] <= batas_test_awal[:4] + '-12')
                    ]

                    # MAPPING MASTER + EKSTRAK unit_to_key & key_to_specs
                    f_master.seek(0)
                    df_map_master = pd.read_excel(f_master, sheet_name='Sheet2', header=1)
                    df_map_master.columns = df_map_master.columns.str.strip()

                    col_name_m  = next((c for c in df_map_master.columns if 'NAMA' in str(c).upper()), None)
                    col_jenis_m = next((c for c in df_map_master.columns
                                        if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper()
                                        and c != col_name_m), None)
                    col_type_m  = next((c for c in df_map_master.columns
                                        if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
                    col_cap_m   = next((c for c in df_map_master.columns
                                        if str(c).strip().upper() == 'CAP'
                                        or 'CAPAC' in str(c).upper()), None)
                    col_hp_m    = next((c for c in df_map_master.columns
                                        if str(c).strip().upper() == 'HP'
                                        or 'HORSE' in str(c).upper()), None)

                    master_names_set = set()
                    unit_to_key      = {}
                    key_to_specs     = {}

                    if col_name_m:
                        for _, row in df_map_master.iterrows():
                            if pd.isna(row[col_name_m]): continue
                            nama = str(row[col_name_m]).strip()
                            if nama in ('nan', '-', ''): continue
                            master_names_set.add(nama)

                            jenis  = str(row[col_jenis_m]).strip() if col_jenis_m and pd.notna(row[col_jenis_m]) else 'Tidak Diketahui'
                            type_m = str(row[col_type_m]).strip()  if col_type_m  and pd.notna(row[col_type_m])  else 'Tidak Diketahui'
                            cap    = str(row[col_cap_m]).strip()    if col_cap_m   and pd.notna(row[col_cap_m])   else 'Tidak Diketahui'
                            hp     = str(row[col_hp_m]).strip()     if col_hp_m    and pd.notna(row[col_hp_m])    else 'Tidak Diketahui'

                            for val, attr in [(cap, 'cap'), (hp, 'hp')]:
                                try:
                                    fval = float(val)
                                    val  = str(int(fval)) if fval == int(fval) else str(fval)
                                    if attr == 'cap': cap = val
                                    else:             hp  = val
                                except (ValueError, TypeError):
                                    pass

                            if cap in ('0', '0.0', '-', 'nan', ''): cap = 'Tidak Diketahui'
                            if hp  in ('0', '0.0', '-', 'nan', ''): hp  = 'Tidak Diketahui'

                            composite_key = f"{jenis}|{type_m}|{cap}|{hp}"
                            unit_to_key[nama] = composite_key
                            if composite_key not in key_to_specs:
                                key_to_specs[composite_key] = {
                                    'jenis':     jenis,
                                    'type_merk': type_m,
                                    'cap':       cap,
                                    'hp':        hp,
                                }

                    # ----------------------------------------------------------
                    # STANDAR PABRIK: CRAWL (jika mode crawl) ATAU BACA DARI SERVER
                    # ----------------------------------------------------------
                    def crawl_standar_pabrik_openai(key_to_specs_dict, api_key):
                        """Crawl estimasi standar konsumsi BBM per kombinasi jenis/type/cap/hp via OpenAI."""
                        client     = OpenAI(api_key=api_key)
                        hasil      = {}
                        total_keys = len(key_to_specs_dict)

                        CAPS_PER_JENIS = {
                            'FORKLIFT':      (0.5, 15.0),
                            'REACH STACKER': (5.0, 25.0),
                            'CRANE':         (4.0, 40.0),
                            'LOADER':        (3.0, 20.0),
                            'TRONTON':       (2.0, 12.0),
                            'TRAILER':       (2.0, 12.0),
                            'HEAD':          (2.0, 12.0),
                            'TRUCK':         (2.0, 12.0),
                            'DEFAULT':       (0.5, 50.0),
                        }

                        def get_cap(jenis_str):
                            j = jenis_str.upper()
                            for k, v in CAPS_PER_JENIS.items():
                                if k in j:
                                    return v
                            return CAPS_PER_JENIS['DEFAULT']

                        def clamp(val, lo, hi):
                            try:
                                return max(lo, min(hi, float(val)))
                            except (TypeError, ValueError):
                                return None

                        crawl_prog = st.progress(0, text="Memulai crawling standar pabrik...")

                        for idx, (key, specs) in enumerate(key_to_specs_dict.items()):
                            crawl_prog.progress((idx + 1) / total_keys,
                                                text=f"Crawling: {key} ({idx+1}/{total_keys})")
                            jenis  = specs['jenis']
                            type_m = specs['type_merk']
                            cap    = specs['cap']
                            hp     = specs['hp']

                            jenis_upper = jenis.upper()
                            is_wheeled  = any(k in jenis_upper for k in ['TRONTON', 'TRAILER', 'HEAD', 'TRUCK'])

                            if is_wheeled:
                                cap_display = (f"{cap} feet (ukuran peti kemas/container)"
                                              if cap != 'Tidak Diketahui' else cap)
                                panduan_khusus = (
                                    "\nPERHATIAN KHUSUS — Kendaraan Angkut Beroda (Tronton/Trailer/Head/Truck):\n"
                                    f"- Kapasitas '{cap}' di atas adalah ukuran peti kemas dalam FEET (bukan ton beban).\n"
                                    "  Contoh: 20 feet = kontainer 20 kaki, 40 feet = kontainer 40 kaki.\n"
                                    "  Jangan menafsirkan angka kapasitas ini sebagai berat muatan dalam ton.\n"
                                    "- Satuan konsumsi yang diminta adalah Liter per Jam operasional mesin menyala (bukan L/km).\n"
                                    "- Konsumsi wajar kendaraan kelas ini saat operasional normal: 2-12 L/Jam.\n"
                                    "- HP tinggi (200-500 HP) pada kendaraan beroda TIDAK berarti konsumsi 15+ L/Jam;\n"
                                    "  engine truck beroperasi jauh di bawah kapasitas maksimum saat berjalan normal.\n"
                                    "- Nilai di atas 15 L/Jam untuk jenis kendaraan ini hampir pasti TIDAK WAJAR.\n"
                                    "- Gunakan referensi konsumsi BBM riil truck/trailer logistik, bukan mesin stasioner.\n"
                                )
                            else:
                                cap_display    = f"{cap} ton"
                                panduan_khusus = ""

                            prompt = f"""Anda adalah database spesifikasi teknis alat berat industri pelabuhan dan logistik.
Berikan estimasi standar konsumsi bahan bakar (fuel consumption) dalam satuan Liter per Jam (L/Jam)
untuk alat berat dengan spesifikasi berikut:

- Jenis Alat   : {jenis}
- Type / Merk  : {type_m}
- Kapasitas    : {cap_display}
- Horse Power  : {hp} HP ('Tidak Diketahui' jika tidak tersedia)
{panduan_khusus}
Panduan estimasi:
1. Gunakan data spesifikasi resmi pabrik jika type/merk dikenali secara spesifik.
2. Jika tidak dikenali, gunakan rentang umum berdasarkan jenis alat dan HP-nya.
3. Jika HP tidak diketahui, estimasi berdasarkan jenis dan kapasitas saja.
4. Asumsikan kondisi operasional normal: beban penuh, medan datar, suhu 25-35 derajat Celsius.
5. Jawab HANYA dalam format JSON berikut tanpa penjelasan atau teks tambahan apapun:
{{
  "konsumsi_min_L_per_jam": <angka float>,
  "konsumsi_max_L_per_jam": <angka float>,
  "konsumsi_tengah_L_per_jam": <angka float>,
  "sumber_info": "<nama model spesifik jika diketahui, atau rentang kelas jenis alat>",
  "catatan": "<asumsi kondisi operasional, maks 120 karakter>"
}}"""

                            MAX_RETRIES = 3
                            berhasil    = False

                            for attempt in range(MAX_RETRIES):
                                try:
                                    response = client.chat.completions.create(
                                        model="gpt-4o-mini",
                                        messages=[{"role": "user", "content": prompt}],
                                        temperature=0.1,
                                        max_tokens=300
                                    )
                                    raw_text = response.choices[0].message.content.strip()
                                    raw_text = re.sub(r'```json|```', '', raw_text).strip()
                                    parsed   = json.loads(raw_text)

                                    cap_lo, cap_hi = get_cap(jenis)
                                    val_min = clamp(parsed.get('konsumsi_min_L_per_jam',    0), cap_lo, cap_hi)
                                    val_max = clamp(parsed.get('konsumsi_max_L_per_jam',    0), cap_lo, cap_hi)
                                    val_tgh = clamp(parsed.get('konsumsi_tengah_L_per_jam', 0), cap_lo, cap_hi)

                                    if val_min is not None and val_max is not None and val_tgh is not None:
                                        val_min = min(val_min, val_max)
                                        val_tgh = max(val_min, min(val_tgh, val_max))

                                    hasil[key] = {
                                        'Standar Pabrik Konsumsi BBM Per Jam':     round(val_tgh, 2) if val_tgh is not None else None,
                                        'Standar Pabrik Konsumsi BBM Min (L/Jam)': round(val_min, 2) if val_min is not None else None,
                                        'Standar Pabrik Konsumsi BBM Max (L/Jam)': round(val_max, 2) if val_max is not None else None,
                                        'Sumber Data Standar Pabrik':              str(parsed.get('sumber_info', '-')),
                                        'Catatan Standar Pabrik':                  str(parsed.get('catatan', '-')),
                                    }
                                    berhasil = True
                                    time.sleep(1)
                                    break

                                except json.JSONDecodeError as e:
                                    hasil[key] = {
                                        'Standar Pabrik Konsumsi BBM Per Jam':     None,
                                        'Standar Pabrik Konsumsi BBM Min (L/Jam)': None,
                                        'Standar Pabrik Konsumsi BBM Max (L/Jam)': None,
                                        'Sumber Data Standar Pabrik':              'Gagal diproses (JSON error)',
                                        'Catatan Standar Pabrik':                  str(e)[:100],
                                    }
                                    break

                                except Exception as e:
                                    if '429' in str(e) or 'rate' in str(e).lower():
                                        time.sleep(30 * (attempt + 1))
                                    else:
                                        hasil[key] = {
                                            'Standar Pabrik Konsumsi BBM Per Jam':     None,
                                            'Standar Pabrik Konsumsi BBM Min (L/Jam)': None,
                                            'Standar Pabrik Konsumsi BBM Max (L/Jam)': None,
                                            'Sumber Data Standar Pabrik':              'Error API',
                                            'Catatan Standar Pabrik':                  str(e)[:100],
                                        }
                                        break

                            if not berhasil and key not in hasil:
                                hasil[key] = {
                                    'Standar Pabrik Konsumsi BBM Per Jam':     None,
                                    'Standar Pabrik Konsumsi BBM Min (L/Jam)': None,
                                    'Standar Pabrik Konsumsi BBM Max (L/Jam)': None,
                                    'Sumber Data Standar Pabrik':              'Gagal setelah max retry',
                                    'Catatan Standar Pabrik':                  '-',
                                }

                        crawl_prog.empty()
                        return hasil

                    def load_standar_pabrik_dari_server(path_server):
                        """Baca file Konsumsi BBM Standar Pabrik hasil crawl terakhir dari path server."""
                        hasil = {}
                        if not os.path.exists(path_server):
                            return hasil
                        try:
                            df_sp = pd.read_excel(path_server)
                            rename_sp = {
                                'Standar_Konsumsi_L_per_Jam':     'Standar Pabrik Konsumsi BBM Per Jam',
                                'Standar_Konsumsi_Min_L_per_Jam': 'Standar Pabrik Konsumsi BBM Min (L/Jam)',
                                'Standar_Konsumsi_Max_L_per_Jam': 'Standar Pabrik Konsumsi BBM Max (L/Jam)',
                                'Sumber_Data_Standar':            'Sumber Data Standar Pabrik',
                                'Catatan_Standar':                'Catatan Standar Pabrik',
                            }
                            df_sp.rename(columns=rename_sp, inplace=True)
                            for _, row in df_sp.iterrows():
                                key = str(row['Composite_Key'])
                                hasil[key] = {
                                    'Standar Pabrik Konsumsi BBM Per Jam':     row.get('Standar Pabrik Konsumsi BBM Per Jam'),
                                    'Standar Pabrik Konsumsi BBM Min (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Min (L/Jam)'),
                                    'Standar Pabrik Konsumsi BBM Max (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Max (L/Jam)'),
                                    'Sumber Data Standar Pabrik':              row.get('Sumber Data Standar Pabrik', '-'),
                                    'Catatan Standar Pabrik':                  row.get('Catatan Standar Pabrik', '-'),
                                }
                        except Exception as e_sp:
                            st.warning(f"Gagal membaca file standar pabrik dari server: {e_sp}")
                        return hasil

                    standar_per_key = {}

                    if is_mode_crawl:
                        if not OPENAI_API_KEY.strip():
                            st.error("API Key belum dikonfigurasi oleh Tim IT. Hubungi Tim IT untuk mengisi API Key di kode aplikasi.")
                            st.stop()

                        st.info("🌐 Menjalankan crawling standar pabrik terlebih dahulu...")
                        standar_per_key = crawl_standar_pabrik_openai(key_to_specs, OPENAI_API_KEY)

                        # Simpan hasil crawl ke path server agar bisa dipakai ulang untuk mode "Forecast Langsung"
                        try:
                            rows_standar = []
                            for key, val in standar_per_key.items():
                                parts = key.split('|')
                                rows_standar.append({
                                    'Jenis_Alat':                              parts[0] if len(parts) > 0 else '-',
                                    'Type_Merk':                               parts[1] if len(parts) > 1 else '-',
                                    'Capacity':                                parts[2] if len(parts) > 2 else '-',
                                    'Horse_Power':                             parts[3] if len(parts) > 3 else '-',
                                    'Composite_Key':                           key,
                                    'Standar Pabrik Konsumsi BBM Per Jam':     val['Standar Pabrik Konsumsi BBM Per Jam'],
                                    'Standar Pabrik Konsumsi BBM Min (L/Jam)': val['Standar Pabrik Konsumsi BBM Min (L/Jam)'],
                                    'Standar Pabrik Konsumsi BBM Max (L/Jam)': val['Standar Pabrik Konsumsi BBM Max (L/Jam)'],
                                    'Sumber Data Standar Pabrik':              val['Sumber Data Standar Pabrik'],
                                    'Catatan Standar Pabrik':                  val['Catatan Standar Pabrik'],
                                })
                            os.makedirs(os.path.dirname(PATH_STANDAR_PABRIK_SERVER), exist_ok=True)
                            pd.DataFrame(rows_standar).to_excel(PATH_STANDAR_PABRIK_SERVER, index=False)
                            st.success("✅ Hasil crawling disimpan ke server. Melanjutkan ke proses forecast...")
                        except Exception as e_save:
                            st.warning(f"Hasil crawling berhasil didapat tetapi gagal disimpan ke server: {e_save}. "
                                       f"Proses forecast tetap dilanjutkan menggunakan hasil crawling saat ini.")
                    else:
                        standar_per_key = load_standar_pabrik_dari_server(PATH_STANDAR_PABRIK_SERVER)
                        if not standar_per_key:
                            st.warning("⚠️ File standar pabrik belum tersedia di server. Forecast tetap berjalan "
                                       "tapi kolom standar pabrik tidak akan terisi. Jalankan mode 'Crawl Standar "
                                       "Pabrik + Forecast' minimal sekali sebelumnya.")

                    def get_mapped_unit_name_fcst(unit_name):
                        hardcoded = {
                            "FL RENTAL 01":               "FL RENTAL 01 TIMIKA",
                            "TOBATI (EX.FL KALMAR 32T)":  "TOP LOADER KALMAR 35T/TOBATI",
                            "L 8477 UUC (EX.L 9902 UR)":  "L 9902 UR / S75",
                            "WIND RIVER (EX.TL BOSS 42T)": "TOP LOADER BOSS"
                        }
                        unit_name = str(unit_name).strip()
                        if unit_name in hardcoded and hardcoded[unit_name] in master_names_set:
                            return hardcoded[unit_name]
                        if unit_name in master_names_set:
                            return unit_name
                        if " (" in unit_name:
                            before = unit_name.split(" (")[0].strip()
                            if before in master_names_set:
                                return before
                        if "EX." in unit_name.upper():
                            m = re.search(r'EX\.([^\)]+)', unit_name.upper())
                            if m:
                                after = m.group(1).strip()
                                if after in master_names_set:
                                    return after
                        return None

                    # FUNGSI PREPROCESSING & MODEL
                    def preprocess_ts(series):
                        df  = pd.DataFrame(series, columns=['HM'])
                        p05 = df['HM'].quantile(0.05)
                        p95 = df['HM'].quantile(0.95)
                        df['HM_Capped']   = df['HM'].clip(lower=p05, upper=p95) if p95 > 0 else df['HM']
                        df['HM_Smoothed'] = df['HM_Capped'].ewm(span=3, min_periods=1).mean()
                        return df['HM_Smoothed']

                    def prepare_gb_features(series, n_lags=3):
                        df = pd.DataFrame(series.values, columns=['y'])
                        for i in range(1, n_lags + 1):
                            df[f'lag_{i}'] = df['y'].shift(i)
                        df['rolling_mean_3'] = df['y'].shift(1).rolling(window=3, min_periods=1).mean()
                        df['rolling_std_3']  = df['y'].shift(1).rolling(window=3, min_periods=1).std().fillna(0)
                        df['trend']          = np.arange(len(df))
                        df = df.dropna()
                        fcols = [f'lag_{i}' for i in range(1, n_lags + 1)] + ['rolling_mean_3', 'rolling_std_3', 'trend']
                        return df[fcols], df['y'], fcols

                    def predict_gb(train_series, steps_ahead):
                        n = len(train_series)
                        if n >= 7:   n_lags = 3
                        elif n >= 5: n_lags = 2
                        elif n >= 4: n_lags = 1
                        else:        return np.full(steps_ahead, max(0.0, float(train_series.mean())))
                        X_train, y_train, fcols = prepare_gb_features(train_series, n_lags)
                        if len(X_train) < 3:
                            return np.full(steps_ahead, max(0.0, float(train_series.mean())))
                        model = HistGradientBoostingRegressor(
                            max_iter=200, learning_rate=0.05, max_depth=4, random_state=42
                        ).fit(X_train, y_train)
                        preds, history = [], list(train_series.values)
                        trend_offset   = len(history)
                        for step in range(steps_ahead):
                            lag_vals  = [history[-(i)] for i in range(1, n_lags + 1)]
                            window    = history[-3:] if len(history) >= 3 else history
                            roll_mean = float(np.mean(window))
                            roll_std  = float(np.std(window)) if len(window) > 1 else 0.0
                            row_feat  = lag_vals + [roll_mean, roll_std, trend_offset + step]
                            pred      = max(0.0, model.predict(pd.DataFrame([row_feat], columns=fcols))[0])
                            preds.append(pred)
                            history.append(pred)
                        return np.array(preds)

                    def mape_aman(actual, pred):
                        actual, pred = np.array(actual), np.array(pred)
                        mask = actual != 0
                        if not np.any(mask): return 0.0
                        return float(np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100)

                    def ensemble_bobot(p_arima, p_gb, aktual_h):
                        rmse_a = float(np.sqrt(mean_squared_error(aktual_h, p_arima))) + 1e-6
                        rmse_g = float(np.sqrt(mean_squared_error(aktual_h, p_gb)))    + 1e-6
                        w_a    = (1 / rmse_a) / (1 / rmse_a + 1 / rmse_g)
                        w_g    = (1 / rmse_g) / (1 / rmse_a + 1 / rmse_g)
                        return w_a * p_arima + w_g * p_gb, round(w_a, 4), round(w_g, 4)

                    def kategorikan_deviasi(pct):
                        if pct is None or pd.isna(pct): return 'Data Tidak Tersedia'
                        if pct <= -10: return 'Sangat Hemat (≤ -10%)'
                        if pct <    0: return 'Hemat (-10% s.d. 0%)'
                        if pct <=  10: return 'Normal (0% s.d. +10%)'
                        if pct <=  25: return 'Sedikit Boros (+10% s.d. +25%)'
                        return 'Boros (> +25%)'

                    # PIPELINE UTAMA
                    list_unit_raw    = train_agg['EQUIP NAME'].unique()
                    results_combined = []
                    metrics_list     = []
                    excl_list        = []
                    all_actual, all_arima, all_gb, all_ens = [], [], [], []
                    total_pop = 0

                    prog_bar = st.progress(0, text="Memulai training model...")
                    n_units  = len(list_unit_raw)

                    for idx_u, unit in enumerate(list_unit_raw):
                        prog_bar.progress((idx_u + 1) / n_units,
                                          text=f"Training: {unit} ({idx_u+1}/{n_units})")
                        mapped_name = None
                        try:
                            mapped_name = get_mapped_unit_name_fcst(unit)
                            if not mapped_name:
                                continue
                            total_pop += 1

                            df_u_train = train_agg[train_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
                            df_u_test  = test_agg[test_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()

                            comp_key    = unit_to_key.get(mapped_name, '')
                            sp          = standar_per_key.get(comp_key, {})
                            std_per_jam = sp.get('Standar Pabrik Konsumsi BBM Per Jam',     None)
                            std_min     = sp.get('Standar Pabrik Konsumsi BBM Min (L/Jam)', None)
                            std_max     = sp.get('Standar Pabrik Konsumsi BBM Max (L/Jam)', None)
                            std_sumber  = sp.get('Sumber Data Standar Pabrik',              'Tidak tersedia')
                            std_catatan = sp.get('Catatan Standar Pabrik',                  '-')

                            specs_unit = key_to_specs.get(comp_key, {})
                            spec_jenis = specs_unit.get('jenis',     'Tidak Diketahui')
                            spec_type  = specs_unit.get('type_merk', 'Tidak Diketahui')
                            spec_cap   = specs_unit.get('cap',       'Tidak Diketahui')
                            spec_hp    = specs_unit.get('hp',        'Tidak Diketahui')

                            if df_u_test.empty:
                                excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                  'Alasan': 'Tidak ada data aktual di periode uji.'}); continue
                            if len(df_u_train) < 12:
                                excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                  'Alasan': 'Data latih kurang dari 12 bulan.'}); continue
                            if len(df_u_test) < 12:
                                excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                  'Alasan': 'Data uji kurang dari 12 bulan.'}); continue

                            # Cek per tahun dalam data train secara dinamis (tidak hardcode 2023/2024)
                            train_years = df_u_train.index.map(lambda p: p.year).unique()
                            skip_unit = False
                            for yr in train_years:
                                df_yr = df_u_train[df_u_train.index.map(lambda p: p.year) == yr]
                                if len(df_yr) == 12 and (df_yr['HM'].sum() == 0 or df_yr['LITER'].sum() == 0):
                                    excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                      'Alasan': f'HM/LITER = 0 selama 1 tahun penuh pada data train tahun {yr}'})
                                    skip_unit = True
                                    break
                            if skip_unit: continue
                            if df_u_test['HM'].sum() == 0 or df_u_test['LITER'].sum() == 0:
                                excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                  'Alasan': 'HM/LITER = 0 selama 1 tahun penuh pada data test'}); continue

                            true_ratio = float(df_u_train['LITER'].sum() / df_u_train['HM'].sum()) \
                                         if df_u_train['HM'].sum() > 0 else 0.0
                            aktual_l   = df_u_test['LITER'].values.astype(float)
                            aktual_h   = df_u_test['HM'].values.astype(float)
                            steps      = len(df_u_test)

                            total_hm_test    = float(aktual_h.sum())
                            total_liter_test = float(aktual_l.sum())
                            aktual_l_per_jam = (total_liter_test / total_hm_test
                                                if total_hm_test > 0 else None)
                            deviasi_pct      = (
                                (aktual_l_per_jam - std_min) / std_min * 100
                                if std_min and std_min > 0 and aktual_l_per_jam
                                else None
                            )

                            t_utuh = df_u_train['HM'].copy()
                            try:
                                first_idx = df_u_train[df_u_train['HM'] > 0].index[0]
                                t_potong  = df_u_train.loc[first_idx:]['HM'].copy()
                            except IndexError:
                                t_potong = t_utuh

                            best_arima, best_gb = np.zeros(steps), np.zeros(steps)
                            min_rmse, model_ok  = float('inf'), False

                            for _, ds in [("U", t_utuh), ("P", t_potong)]:
                                if len(ds) < 6: continue
                                model_ok = True
                                ds_s = preprocess_ts(ds)
                                try:
                                    arima_m  = pm.auto_arima(ds_s, seasonal=False, max_d=1,
                                                             suppress_warnings=True, error_action="ignore")
                                    p_ar_raw = arima_m.predict(n_periods=steps)
                                    baseline = float(ds_s.tail(6).mean())
                                    p_ar     = np.clip(p_ar_raw, baseline * 0.1, baseline * 2.0)
                                    p_ar     = np.maximum(0.0, p_ar).astype(float)
                                except Exception:
                                    p_ar = np.full(steps, max(0.0, float(ds_s.mean())))
                                try:
                                    p_gb = predict_gb(ds_s, steps).astype(float)
                                except Exception:
                                    p_gb = np.full(steps, max(0.0, float(ds_s.mean())))

                                p_ens_t, _, _ = ensemble_bobot(p_ar, p_gb, aktual_h)
                                rmse_t = float(np.sqrt(mean_squared_error(aktual_h, p_ens_t)))
                                if rmse_t < min_rmse:
                                    min_rmse, best_arima, best_gb = rmse_t, p_ar.copy(), p_gb.copy()

                            if not model_ok:
                                excl_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                                                  'Alasan': 'Data historis valid setelah dipotong < 6 bulan.'}); continue

                            best_ensemble, w_a, w_g = ensemble_bobot(best_arima, best_gb, aktual_h)
                            best_ensemble = best_ensemble.astype(float)

                            rmse_arima    = float(np.sqrt(mean_squared_error(aktual_h, best_arima)))
                            mape_arima    = mape_aman(aktual_h, best_arima)
                            rmse_gb       = float(np.sqrt(mean_squared_error(aktual_h, best_gb)))
                            mape_gb       = mape_aman(aktual_h, best_gb)
                            rmse_ensemble = float(np.sqrt(mean_squared_error(aktual_h, best_ensemble)))
                            mape_ensemble = mape_aman(aktual_h, best_ensemble)

                            mape_dict     = {'ARIMA': mape_arima, 'Gradient Boosting': mape_gb, 'Ensemble': mape_ensemble}
                            rmse_dict     = {'ARIMA': rmse_arima, 'Gradient Boosting': rmse_gb, 'Ensemble': rmse_ensemble}
                            best_mdl      = min(mape_dict, key=mape_dict.get)
                            pred_terpilih = {'ARIMA': best_arima, 'Gradient Boosting': best_gb,
                                             'Ensemble': best_ensemble}[best_mdl]
                            rmse_terpilih = rmse_dict[best_mdl]

                            metrics_list.append({
                                'EQUIP NAME':               unit,
                                'NAMA_MASTER_TERPETAKAN':   mapped_name,
                                'Jenis_Alat':               spec_jenis,
                                'Type_Merk':                spec_type,
                                'Capacity':                 spec_cap,
                                'Horse_Power':              spec_hp,
                                'Model_Terpilih_Unit':      best_mdl,
                                'RMSE_ARIMA':               round(rmse_arima, 2),
                                'MAPE_ARIMA (%)':           round(mape_arima, 2),
                                'RMSE_GB':                  round(rmse_gb, 2),
                                'MAPE_GB (%)':              round(mape_gb, 2),
                                'RMSE_Ensemble':            round(rmse_ensemble, 2),
                                'MAPE_Ensemble (%)':        round(mape_ensemble, 2),
                                'MAPE_Terpilih (%)':        round(mape_dict[best_mdl], 2),
                                'RMSE_Terpilih':            round(rmse_terpilih, 2),
                                'Bobot_ARIMA':              w_a,
                                'Bobot_GB':                 w_g,
                                'Standar Pabrik Konsumsi BBM Per Jam':      std_per_jam,
                                'Standar Pabrik Konsumsi BBM Min (L/Jam)':  std_min,
                                'Standar Pabrik Konsumsi BBM Max (L/Jam)':  std_max,
                                'Konsumsi BBM Aktual Per Jam':              round(aktual_l_per_jam, 4) if aktual_l_per_jam else None,
                                'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': round(deviasi_pct, 2) if deviasi_pct is not None else None,
                                'Kategori Efisiensi':                       kategorikan_deviasi(deviasi_pct),
                                'Sumber Data Standar Pabrik':               std_sumber,
                                'Catatan Standar Pabrik':                   std_catatan,
                            })

                            all_actual.extend(aktual_h.tolist())
                            all_arima.extend(best_arima.tolist())
                            all_gb.extend(best_gb.tolist())
                            all_ens.extend(best_ensemble.tolist())

                            for i, period in enumerate(df_u_test.index):
                                std_per_bulan = (round(float(pred_terpilih[i]) * std_min, 2)
                                                 if std_min else None)

                                results_combined.append({
                                    'EQUIP NAME':             unit,
                                    'NAMA_MASTER_TERPETAKAN': mapped_name,
                                    'Jenis_Alat':             spec_jenis,
                                    'Type_Merk':              spec_type,
                                    'Capacity':               spec_cap,
                                    'Horse_Power':            spec_hp,
                                    'Bulan':                  str(period),
                                    'Aktual_HM':              round(float(aktual_h[i]), 2),
                                    'Aktual_LITER':           round(float(aktual_l[i]), 2),
                                    'Prediksi_HM_ARIMA':          round(float(best_arima[i]), 2),
                                    'Prediksi_LITER_ARIMA':       round(float(best_arima[i]) * true_ratio, 2),
                                    'Prediksi_HM_GB':             round(float(best_gb[i]), 2),
                                    'Prediksi_LITER_GB':          round(float(best_gb[i]) * true_ratio, 2),
                                    'Prediksi_HM_Ensemble':       round(float(best_ensemble[i]), 2),
                                    'Prediksi_LITER_Ensemble':    round(float(best_ensemble[i]) * true_ratio, 2),
                                    'Model_Terpilih_Unit':         best_mdl,
                                    'Prediksi_HM_Terpilih':       round(float(pred_terpilih[i]), 2),
                                    'Prediksi_LITER_Terpilih':    round(float(pred_terpilih[i]) * true_ratio, 2),
                                    'Standar Pabrik Minimum Konsumsi BBM Per Jam':   std_min,
                                    'Standar Pabrik Minimum Konsumsi BBM Per Bulan': std_per_bulan,
                                    'Konsumsi BBM Aktual Per Jam':                   round(aktual_l_per_jam, 4) if aktual_l_per_jam else None,
                                    'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': round(deviasi_pct, 2) if deviasi_pct is not None else None,
                                    'Kategori Efisiensi':                            kategorikan_deviasi(deviasi_pct),
                                    'Sumber Data Standar Pabrik':                    std_sumber,
                                    'MAPE_ARIMA (%)':       round(mape_arima, 2),
                                    'RMSE_ARIMA':           round(rmse_arima, 2),
                                    'MAPE_GB (%)':          round(mape_gb, 2),
                                    'RMSE_GB':              round(rmse_gb, 2),
                                    'MAPE_Ensemble (%)':    round(mape_ensemble, 2),
                                    'RMSE_Ensemble':        round(rmse_ensemble, 2),
                                    'MAPE_Terpilih (%)':    round(mape_dict[best_mdl], 2),
                                    'RMSE_Terpilih':        round(rmse_terpilih, 2),
                                    'Bobot_ARIMA':          w_a,
                                    'Bobot_GB':             w_g,
                                })

                        except Exception as e:
                            excl_list.append({'EQUIP NAME': unit,
                                              'NAMA_MASTER_TERPETAKAN': mapped_name if mapped_name else '-',
                                              'Alasan': f'Gagal diproses: {str(e)}'})
                            continue

                    prog_bar.empty()

                    df_komparasi = pd.DataFrame(results_combined)
                    df_metrik    = pd.DataFrame(metrics_list)
                    df_excl      = pd.DataFrame(excl_list)

                    out_fcst = io.BytesIO()
                    with pd.ExcelWriter(out_fcst, engine='openpyxl') as writer:
                        df_komparasi.to_excel(writer, sheet_name='Komparasi_Model', index=False)

                        if not df_komparasi.empty:
                            cols_terpilih = [
                                'EQUIP NAME', 'NAMA_MASTER_TERPETAKAN',
                                'Jenis_Alat', 'Type_Merk', 'Capacity', 'Horse_Power',
                                'Bulan', 'Aktual_HM', 'Aktual_LITER',
                                'Model_Terpilih_Unit',
                                'Prediksi_HM_Terpilih', 'Prediksi_LITER_Terpilih',
                                'Standar Pabrik Minimum Konsumsi BBM Per Jam',
                                'Standar Pabrik Minimum Konsumsi BBM Per Bulan',
                                'Konsumsi BBM Aktual Per Jam',
                                'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)',
                                'Kategori Efisiensi',
                                'Sumber Data Standar Pabrik',
                                'MAPE_Terpilih (%)', 'RMSE_Terpilih',
                            ]
                            cols_ok_tp = [c for c in cols_terpilih if c in df_komparasi.columns]
                            df_tp = df_komparasi[cols_ok_tp].copy()
                            df_tp[df_tp['MAPE_Terpilih (%)'] < 35].to_excel(writer, sheet_name='Akurasi_Bagus_Under35', index=False)
                            df_tp[df_tp['MAPE_Terpilih (%)'] >= 35].to_excel(writer, sheet_name='Akurasi_Rendah_Over35', index=False)

                        df_metrik.to_excel(writer, sheet_name='Metrik_Per_Unit', index=False)

                        if not df_metrik.empty:
                            cols_deviasi = [
                                'EQUIP NAME', 'NAMA_MASTER_TERPETAKAN',
                                'Jenis_Alat', 'Type_Merk', 'Capacity', 'Horse_Power',
                                'Model_Terpilih_Unit', 'MAPE_Terpilih (%)',
                                'Standar Pabrik Konsumsi BBM Per Jam',
                                'Standar Pabrik Konsumsi BBM Min (L/Jam)',
                                'Standar Pabrik Konsumsi BBM Max (L/Jam)',
                                'Konsumsi BBM Aktual Per Jam',
                                'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)',
                                'Kategori Efisiensi',
                                'Sumber Data Standar Pabrik',
                                'Catatan Standar Pabrik',
                            ]
                            cols_ok_dev = [c for c in cols_deviasi if c in df_metrik.columns]
                            df_deviasi  = df_metrik[cols_ok_dev].copy()
                            if 'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)' in df_deviasi.columns:
                                df_deviasi = df_deviasi.sort_values(
                                    'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)', ascending=False)
                            df_deviasi.to_excel(writer, sheet_name='Deviasi_Standar_Pabrik', index=False)

                        if not df_excl.empty:
                            df_excl.to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)

                    st.session_state.fcst_hm_df_komparasi  = df_komparasi
                    st.session_state.fcst_hm_df_metrik     = df_metrik
                    st.session_state.fcst_hm_df_excl       = df_excl
                    st.session_state.fcst_hm_out_file      = out_fcst.getvalue()
                    st.session_state.fcst_hm_processed     = True

                except ImportError:
                    st.error("Library `pmdarima` tidak ditemukan. Jalankan: `pip install pmdarima`")
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")
        else:
            if not f_train_files:
                st.warning("Mohon upload minimal 1 file Data Train.")
            if not f_test_files:
                st.warning("Mohon upload minimal 1 file Data Test.")
            if not f_master:
                st.warning("Mohon upload file Detail Alat Berat.")

    # DASHBOARD HASIL
    if st.session_state.fcst_hm_processed:
        df_komparasi   = st.session_state.fcst_hm_df_komparasi
        df_metrik      = st.session_state.fcst_hm_df_metrik
        df_excl        = st.session_state.fcst_hm_df_excl
        df_monthly_raw = st.session_state.fcst_hm_df_monthly_raw

        st.success("✅ Proses Forecast Selesai!")

        # KPI CARDS
        total_berhasil = df_metrik['EQUIP NAME'].nunique() if not df_metrik.empty else 0
        total_excl_n   = len(df_excl)
        unit_under35   = int((df_metrik['MAPE_Terpilih (%)'] < 35).sum()) if not df_metrik.empty else 0
        unit_over35    = int((df_metrik['MAPE_Terpilih (%)'] >= 35).sum()) if not df_metrik.empty else 0

        kc1, kc2, kc3, kc4, kc5 = st.columns(5)
        kc1.metric("Total Unit Diproses",       f"{total_berhasil + total_excl_n} unit")
        kc2.metric("Berhasil Dimodelkan",        f"{total_berhasil} unit")
        kc3.metric("Prediksi Low Error (< 35%)", f"{unit_under35} unit")
        kc4.metric("Prediksi High Error (≥ 35%)",f"{unit_over35} unit")
        kc5.metric("Unit Di-exclude",            f"{total_excl_n} unit")

        # TABS
        tab_ovr, tab_mape, tab_tren, tab_unit, tab_excl, tab_dl = st.tabs([
            "📊 Overview",
            "📈 Tingkat Error Prediksi",
            "📉 Tren & Proyeksi",
            "📋 Unit Termodelkan",
            "🚫 Unit Di-exclude",
            "📥 Download",
        ])

        # TAB 1: OVERVIEW
        with tab_ovr:
            col_pie, col_bar = st.columns([1, 1])

            with col_pie:
                pie_data = pd.DataFrame({
                    'Status': ['Prediksi Low Error (< 35%)', 'Prediksi High Error (≥ 35%)', 'Di-exclude'],
                    'Jumlah': [unit_under35, unit_over35, total_excl_n]
                })
                fig_pie = px.pie(pie_data, names='Status', values='Jumlah', hole=0.45,
                                 color='Status',
                                 color_discrete_map={
                                     'Prediksi Low Error (< 35%)':  '#009943',
                                     'Prediksi High Error (≥ 35%)': '#E3000F',
                                     'Di-exclude':                  '#E6E6E6'
                                 }, title="Pembagian Status Unit")
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                fig_pie.update_layout(showlegend=False, margin=dict(t=40, b=10), height=360)
                st.plotly_chart(fig_pie, use_container_width=True)

            with col_bar:
                if not df_metrik.empty and 'Jenis_Alat' in df_metrik.columns:
                    df_ovr = df_metrik.copy()
                    df_ovr['Status_Ovr'] = df_ovr['MAPE_Terpilih (%)'].apply(
                        lambda x: 'Low Error' if x < 35 else 'High Error')
                    df_ovr_grp = df_ovr.groupby(['Jenis_Alat', 'Status_Ovr']).size().reset_index(name='Jumlah')
                    fig_bar_ovr = px.bar(df_ovr_grp, x='Jenis_Alat', y='Jumlah',
                                         color='Status_Ovr', barmode='stack',
                                         color_discrete_map={
                                             'Low Error':  '#009943',
                                             'High Error': '#E3000F'
                                         },
                                         title="Jumlah Unit per Jenis Alat",
                                         text='Jumlah')
                    fig_bar_ovr.update_traces(textposition='inside')
                    fig_bar_ovr.update_layout(
                        xaxis_tickangle=-30, height=360,
                        margin=dict(t=40, b=10),
                        legend=dict(orientation='h', yanchor='top',
                                    y=-0.25, xanchor='center', x=0.5,
                                    title_text=''),
                        xaxis_title='', yaxis_title='Jumlah Unit'
                    )
                    st.plotly_chart(fig_bar_ovr, use_container_width=True)

            if not df_metrik.empty and 'Jenis_Alat' in df_metrik.columns:
                st.markdown("**Ringkasan per Jenis Alat:**")
                df_sum = df_metrik.groupby('Jenis_Alat').agg(
                    Jumlah_Unit=('EQUIP NAME', 'count'),
                    Low_Error=('MAPE_Terpilih (%)', lambda x: (x < 35).sum()),
                    High_Error=('MAPE_Terpilih (%)', lambda x: (x >= 35).sum()),
                ).reset_index()
                df_sum.columns = ['Jenis Alat', 'Total Unit', 'Prediksi Low Error', 'Prediksi High Error']
                st.dataframe(df_sum.sort_values('Total Unit', ascending=False).reset_index(drop=True),
                             use_container_width=True, hide_index=True)

        # TAB 2: TINGKAT ERROR PREDIKSI
        with tab_mape:
            if not df_metrik.empty:
                mape_view = st.radio("Lihat per:", ["Per Unit", "Per Jenis Alat"],
                                     horizontal=True, key="mape_view_mode")

                if mape_view == "Per Unit":
                    df_ms = df_metrik.sort_values('MAPE_Terpilih (%)').copy()
                    df_ms['Status'] = df_ms['MAPE_Terpilih (%)'].apply(
                        lambda x: 'Prediksi Low Error (< 35%)' if x < 35 else 'Prediksi High Error (≥ 35%)')
                    fig_mape = px.bar(df_ms, x='EQUIP NAME', y='MAPE_Terpilih (%)',
                                      color='Status',
                                      color_discrete_map={
                                          'Prediksi Low Error (< 35%)':  '#009943',
                                          'Prediksi High Error (≥ 35%)': '#E3000F'
                                      },
                                      title="Tingkat Error Prediksi per Unit")
                    fig_mape.add_hline(y=35, line_dash="dash", line_color="orange",
                                       annotation_text="Threshold 35%",
                                       annotation_position="top right")
                    fig_mape.update_layout(xaxis_tickangle=-45, height=480,
                                           xaxis_title='Nama Unit',
                                           yaxis_title='Nilai Error Unit (%)',
                                           margin=dict(b=10, t=50),
                                           legend=dict(orientation='h', yanchor='top',
                                                       y=-0.25, xanchor='center', x=0.5,
                                                       title_text=''))
                    st.plotly_chart(fig_mape, use_container_width=True)

                else:
                    if 'Jenis_Alat' in df_metrik.columns:
                        df_jenis = df_metrik.groupby('Jenis_Alat').agg(
                            MAPE_Rata=('MAPE_Terpilih (%)', 'mean'),
                            Jumlah_Unit=('EQUIP NAME', 'count')
                        ).reset_index()
                        df_jenis['Status'] = df_jenis['MAPE_Rata'].apply(
                            lambda x: 'Prediksi Low Error (< 35%)' if x < 35 else 'Prediksi High Error (≥ 35%)')
                        fig_jenis = px.bar(df_jenis, x='Jenis_Alat', y='MAPE_Rata',
                                           color='Status', text='Jumlah_Unit',
                                           color_discrete_map={
                                               'Prediksi Low Error (< 35%)':  '#009943',
                                               'Prediksi High Error (≥ 35%)': '#E3000F'
                                           },
                                           title="Rata-rata Tingkat Error Prediksi per Jenis Alat Berat")
                        fig_jenis.update_traces(texttemplate='%{text} unit', textposition='outside')
                        fig_jenis.add_hline(y=35, line_dash="dash", line_color="orange",
                                            annotation_text="Threshold 35%",
                                            annotation_position="top right")
                        fig_jenis.update_layout(xaxis_tickangle=-30, height=450,
                                                xaxis_title='Jenis Alat Berat',
                                                yaxis_title='Rata-rata Nilai Error Jenis Alat Berat (%)',
                                                margin=dict(t=60, b=60),
                                                showlegend=False)
                        st.plotly_chart(fig_jenis, use_container_width=True)

        # TAB 3: TREN & PROYEKSI
        with tab_tren:
            if not df_komparasi.empty and df_monthly_raw is not None:
                col_sel0, col_sel1, col_sel2 = st.columns([1, 2, 1])
                with col_sel0:
                    tren_view = st.radio("Lihat per:", ["Per Unit", "Per Jenis Alat"],
                                         key="tren_view_mode", horizontal=False)
                with col_sel1:
                    if tren_view == "Per Unit":
                        unit_list_chart = sorted(df_komparasi['EQUIP NAME'].unique().tolist())
                        selected_unit   = st.selectbox("Pilih Unit:", unit_list_chart, key="fcst_unit_sel")
                        selected_jenis_tren = None
                    else:
                        jenis_list_tren = sorted(df_komparasi['Jenis_Alat'].dropna().unique().tolist()) \
                                          if 'Jenis_Alat' in df_komparasi.columns else []
                        selected_jenis_tren = st.selectbox("Pilih Jenis Alat:", jenis_list_tren, key="fcst_jenis_sel")
                        selected_unit = None
                with col_sel2:
                    chart_mode = st.radio("Tampilkan:", ["Hour Meter (HM)", "Kebutuhan BBM (Liter)"],
                                          key="fcst_chart_mode", horizontal=False)

                if tren_view == "Per Unit" and selected_unit:
                    df_hist_unit = df_monthly_raw[
                        (df_monthly_raw['EQUIP NAME'] == selected_unit) &
                        (df_monthly_raw['TAHUN_BULAN'] <= batas_train_akhir)
                    ].copy().sort_values('TAHUN_BULAN')

                    df_aktual_unit = df_komparasi[
                        df_komparasi['EQUIP NAME'] == selected_unit
                    ].copy().sort_values('Bulan')
                    chart_title_suffix = selected_unit

                else:
                    df_hist_unit = df_monthly_raw[
                        (df_monthly_raw['EQUIP NAME'].isin(
                            df_komparasi[df_komparasi['Jenis_Alat'] == selected_jenis_tren]['EQUIP NAME'].unique()
                        )) &
                        (df_monthly_raw['TAHUN_BULAN'] <= batas_train_akhir)
                    ].groupby('TAHUN_BULAN')[['HM', 'LITER']].sum().reset_index()

                    df_aktual_unit = df_komparasi[
                        df_komparasi['Jenis_Alat'] == selected_jenis_tren
                    ].groupby('Bulan').agg(
                        Aktual_HM=('Aktual_HM', 'sum'),
                        Aktual_LITER=('Aktual_LITER', 'sum'),
                        Prediksi_HM_Terpilih=('Prediksi_HM_Terpilih', 'sum'),
                        Prediksi_LITER_Terpilih=('Prediksi_LITER_Terpilih', 'sum'),
                    ).reset_index().sort_values('Bulan')
                    chart_title_suffix = selected_jenis_tren

                if chart_mode == "Hour Meter (HM)":
                    y_hist, y_akt = 'HM',    'Aktual_HM'
                    y_tp          = 'Prediksi_HM_Terpilih'
                    y_label       = 'Hour Meter (HM)'
                    chart_title   = f"Tren Hour Meter: {chart_title_suffix}"
                else:
                    y_hist, y_akt = 'LITER', 'Aktual_LITER'
                    y_tp          = 'Prediksi_LITER_Terpilih'
                    y_label       = 'Konsumsi BBM (Liter)'
                    chart_title   = f"Tren Kebutuhan BBM: {chart_title_suffix}"

                fig_line = go.Figure()

                if not df_hist_unit.empty:
                    fig_line.add_trace(go.Scatter(
                        x=df_hist_unit['TAHUN_BULAN'].tolist(),
                        y=[float(v) for v in df_hist_unit[y_hist].tolist()],
                        mode='lines+markers', name='Aktual Data Train',
                        line=dict(color='#636efa', width=2), marker=dict(size=6)
                    ))

                if not df_hist_unit.empty and not df_aktual_unit.empty:
                    last_val   = float(df_hist_unit[y_hist].iloc[-1])
                    last_bln   = str(df_hist_unit['TAHUN_BULAN'].iloc[-1])
                    bulan_2025 = df_aktual_unit['Bulan'].tolist()

                    def safe_float_list(series):
                        return [float(v) for v in series.tolist()]

                    fig_line.add_trace(go.Scatter(
                        x=[last_bln] + bulan_2025,
                        y=[last_val] + safe_float_list(df_aktual_unit[y_akt]),
                        mode='lines+markers', name='Aktual Data Test',
                        line=dict(color='#00cc96', width=2), marker=dict(size=7, symbol='square')
                    ))

                    fig_line.add_trace(go.Scatter(
                        x=[last_bln] + bulan_2025,
                        y=[last_val] + safe_float_list(df_aktual_unit[y_tp]),
                        mode='lines+markers', name='Hasil Forecast',
                        line=dict(color='#d62728', width=3), marker=dict(size=8, symbol='diamond')
                    ))

                if not df_hist_unit.empty:
                    fig_line.add_vline(x=batas_train_akhir, line_dash='dash', line_color='gray')
                    fig_line.add_annotation(x=batas_train_akhir, y=1, yref='paper',
                                            text='Train | Test', showarrow=False,
                                            xanchor='left', yanchor='bottom',
                                            font=dict(color='gray'))

                fig_line.update_layout(
                    title=dict(text=chart_title, y=0.97, x=0),
                    yaxis_title=y_label, xaxis_title='Bulan',
                    height=500,
                    legend=dict(orientation='h', yanchor='top',
                                y=-0.18, xanchor='center', x=0.5),
                    hovermode='x unified',
                    margin=dict(t=50, b=20)
                )
                fig_line.update_yaxes(rangemode='tozero')
                st.plotly_chart(fig_line, use_container_width=True)

                # Info card & tabel detail HANYA untuk mode Per Unit
                if tren_view == "Per Unit" and selected_unit and not df_aktual_unit.empty and not df_metrik.empty:
                    kat_ef  = df_aktual_unit['Kategori Efisiensi'].iloc[0] \
                              if 'Kategori Efisiensi' in df_aktual_unit.columns else '-'
                    sel_pct = df_aktual_unit['Selisih Konsumsi Aktual vs Standar Min Pabrik (%)'].iloc[0] \
                              if 'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)' in df_aktual_unit.columns else None

                    met2, met3 = st.columns(2)
                    met2.metric("Kategori Efisiensi BBM", kat_ef)
                    met3.metric("Selisih vs Std Min Pabrik",
                                f"{sel_pct:+.1f}%" if sel_pct is not None else "-")

                    st.markdown("**Detail Prediksi vs Aktual per Bulan:**")
                    cols_det = [
                        'Bulan', 'Aktual_HM', 'Aktual_LITER',
                        'Prediksi_HM_Terpilih', 'Prediksi_LITER_Terpilih',
                        'Standar Pabrik Minimum Konsumsi BBM Per Jam',
                        'Standar Pabrik Minimum Konsumsi BBM Per Bulan',
                        'Konsumsi BBM Aktual Per Jam',
                        'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)',
                        'Kategori Efisiensi',
                    ]
                    cols_ok = [c for c in cols_det if c in df_aktual_unit.columns]
                    df_detail_show = df_aktual_unit[cols_ok].reset_index(drop=True)
                    df_detail_show.rename(columns={
                        'Aktual_HM':                                         'HM Aktual',
                        'Aktual_LITER':                                      'Liter BBM Aktual',
                        'Prediksi_HM_Terpilih':                              'Prediksi HM',
                        'Prediksi_LITER_Terpilih':                           'Prediksi Liter BBM',
                        'Standar Pabrik Minimum Konsumsi BBM Per Jam':       'Std Min Pabrik (L/Jam)',
                        'Standar Pabrik Minimum Konsumsi BBM Per Bulan':     'Std Min Per Bulan (L)',
                        'Konsumsi BBM Aktual Per Jam':                       'Aktual (L/Jam)',
                        'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': 'Selisih vs Std Min (%)',
                        'Kategori Efisiensi':                                'Kategori Efisiensi',
                    }, inplace=True)
                    st.dataframe(df_detail_show, use_container_width=True)

        # TAB 4: UNIT TERMODELKAN
        with tab_unit:
            if not df_metrik.empty:
                cf1, cf2 = st.columns(2)
                with cf1:
                    f_status = st.selectbox("Filter Status:",
                                            ["Semua", "Prediksi Low Error (MAPE < 35%)", "Prediksi High Error (MAPE ≥ 35%)"],
                                            key="f_status")
                with cf2:
                    f_model = st.selectbox("Filter Model:",
                                           ["Semua"] + sorted(df_metrik['Model_Terpilih_Unit'].unique().tolist()),
                                           key="f_model")

                df_md = df_metrik.copy()
                if f_status == "Prediksi Low Error (MAPE < 35%)":
                    df_md = df_md[df_md['MAPE_Terpilih (%)'] < 35]
                elif f_status == "Prediksi High Error (MAPE ≥ 35%)":
                    df_md = df_md[df_md['MAPE_Terpilih (%)'] >= 35]
                if f_model != "Semua":
                    df_md = df_md[df_md['Model_Terpilih_Unit'] == f_model]

                df_md.rename(columns={
                    'EQUIP NAME':               'Nama Unit',
                    'NAMA_MASTER_TERPETAKAN':   'Nama Master',
                    'Model_Terpilih_Unit':       'Model Terpilih',
                    'Jenis_Alat':               'Jenis Alat',
                    'Type_Merk':                'Type/Merk',
                    'Capacity':                 'Capacity',
                    'Horse_Power':              'Horse Power',
                    'Standar Pabrik Konsumsi BBM Min (L/Jam)':           'Std Min Pabrik (L/Jam)',
                    'Konsumsi BBM Aktual Per Jam':                       'Aktual (L/Jam)',
                    'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': 'Selisih vs Std Min (%)',
                    'Kategori Efisiensi':                                'Kategori Efisiensi',
                }, inplace=True)

                cols_to_show = [
                    'Nama Unit', 'Jenis Alat', 'Type/Merk', 'Capacity', 'Horse Power',
                    'Std Min Pabrik (L/Jam)', 'Aktual (L/Jam)',
                    'Selisih vs Std Min (%)', 'Kategori Efisiensi',
                ]
                cols_available = [c for c in cols_to_show if c in df_md.columns]
                st.dataframe(df_md[cols_available].sort_values('Nama Unit').reset_index(drop=True),
                             use_container_width=True)

        # TAB 5: UNIT DI-EXCLUDE
        with tab_excl:
            if not df_excl.empty:
                dist_al = df_excl['Alasan'].value_counts().reset_index()
                dist_al.columns = ['Alasan Di-exclude', 'Jumlah']
                fig_excl = px.bar(dist_al, x='Jumlah', y='Alasan Di-exclude',
                                  orientation='h', text='Jumlah',
                                  color_discrete_sequence=['#d62728'],
                                  title="Distribusi Alasan Unit Di-exclude")
                fig_excl.update_traces(textposition='outside')
                fig_excl.update_layout(yaxis=dict(categoryorder='total ascending', automargin=True),
                                       showlegend=False, height=400,
                                       margin=dict(t=40, b=20, l=300, r=20))
                st.plotly_chart(fig_excl, use_container_width=True)

                df_excl_show = df_excl[['EQUIP NAME', 'Alasan']].copy()
                df_excl_show.rename(columns={'EQUIP NAME': 'Nama Unit'}, inplace=True)
                st.dataframe(df_excl_show.reset_index(drop=True), use_container_width=True)
            else:
                st.success("Tidak ada unit yang di-exclude!")

        # TAB 6: DOWNLOAD
        with tab_dl:
            st.download_button(
                label="📥 Download Hasil Forecast (.xlsx)",
                data=st.session_state.fcst_hm_out_file,
                file_name="Hasil_Forecast_HM_BBM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.caption("Sheet: Komparasi_Model | Akurasi_Bagus_Under35 | Akurasi_Rendah_Over35 | "
                       "Metrik_Per_Unit | Deviasi_Standar_Pabrik | Unit_Dikecualikan")