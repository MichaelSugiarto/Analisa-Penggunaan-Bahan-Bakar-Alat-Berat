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

warnings.filterwarnings('ignore')

# 1. KONFIGURASI FILE & PATH
FILE_HASIL_TRUCKING     = "HasilTrucking.xlsx"
FILE_HASIL_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_BBM_RAW            = "BBM AAB.xlsx"
FILE_HAULAGE_RAW        = "HAULAGE OKT-DES 2025 (Copy).xlsx"
FILE_DOORING_REVISI     = "DOORING_WITH_DISTANCE_REVISI.xlsx"
FILE_MASTER_REF         = "cost & bbm 2022 sd 2025 HP & Type.xlsx"

if "etl_step1_processed" not in st.session_state:
    st.session_state.etl_step1_processed = False
    st.session_state.out_dooring_file    = None

if "etl_step2_processed" not in st.session_state:
    st.session_state.etl_step2_processed = False
    st.session_state.out_truck_file      = None
    st.session_state.out_nontruck_file   = None

if "forecast_processed" not in st.session_state:
    st.session_state.forecast_processed = False
    st.session_state.fcst_df_res        = None
    st.session_state.fcst_df_final      = None
    st.session_state.fcst_out_file      = None

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

# 3. FUNGSI UTILITIES
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    name = name.replace("FORKLIFT", "FORKLIF")
    return re.sub(r'[^A-Z0-9]', '', name)

def get_smart_match(raw_name, master_dict):
    raw_clean = clean_unit_name(raw_name)
    raw_upper = str(raw_name).upper().strip()
    if raw_clean in master_dict: return raw_clean
    if "L 8477 UUC" in raw_upper:
        target = clean_unit_name("L 9902 UR / S75")
        if target in master_dict: return target
    if "EX." in raw_upper or "EX " in raw_upper:
        parts = raw_upper.split("EX.") if "EX." in raw_upper else raw_upper.split("EX ")
        if len(parts) > 1:
            candidate = clean_unit_name(parts[-1].replace(")", "").strip())
            if candidate in master_dict: return candidate
            for k in master_dict:
                if candidate in k: return k
    if "(" in raw_upper:
        candidate = clean_unit_name(raw_upper.split("(")[0])
        if candidate in master_dict: return candidate
        for k in master_dict:
            if candidate in k: return k
    return None

# 4. LOGIKA PROSES DATA: NON-TRUCKING
@st.cache_data(show_spinner=False)
def process_alat_berat():
    if not os.path.exists(FILE_HASIL_NON_TRUCKING):
        st.warning(f"File {FILE_HASIL_NON_TRUCKING} tidak ditemukan.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    try:
        df_agg     = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Total_Agregat')
        df_monthly = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Data_Bulanan')
        try:
            df_missing = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Unit_Inaktif')
            rename_missing = {
                'Unit_Name': 'Nama Unit', 'Jenis_Alat': 'Jenis', 'Type_Merk': 'Type/Merk',
                'Horse_Power': 'Horse Power', 'Capacity': 'Capacity (Ton)',
                'LITER': 'Total Pengisian BBM (L)', 'Total_Ton': 'Total Berat Angkutan (Ton)',
                'Total Pengisian BBM': 'Total Pengisian BBM (L)'
            }
            df_missing.rename(columns=rename_missing, inplace=True)
        except:
            df_missing = pd.DataFrame()

        df_agg.columns     = df_agg.columns.str.strip()
        df_monthly.columns = df_monthly.columns.str.strip()
        df_agg['Capacity_Num']     = df_agg['Capacity'].fillna(0).astype(float).astype(int)
        df_monthly['Capacity_Num'] = df_monthly['Capacity'].fillna(0).astype(float).astype(int)

        def get_benchmark_group(jenis, cap):
            jenis = str(jenis).upper()
            if 'FORKLIFT' in jenis:
                if 3 <= cap <= 8 or cap == 0: return 'Forklift (Capacity 3-8)'
                elif cap >= 10: return 'Forklift (Capacity 10, 28, 32)'
                return 'Forklift (Lainnya)'
            elif 'REACH STACKER' in jenis: return 'Reach Stacker'
            elif 'LOADER' in jenis: return 'Top Loader & Side Loader'
            elif 'CRANE' in jenis:
                if cap >= 70: return 'Crane (Capacity 75, 80, 127)'
                return 'Crane (Lainnya)'
            elif 'TRONTON' in jenis: return 'Tronton'
            elif 'TRAILER' in jenis or 'HEAD' in jenis: return 'Trailer/Head'
            return 'Lainnya'

        df_agg['Benchmark_Group']     = df_agg.apply(lambda r: get_benchmark_group(r.get('Jenis_Alat',''), r.get('Capacity_Num',0)), axis=1)
        df_monthly['Benchmark_Group'] = df_monthly.apply(lambda r: get_benchmark_group(r.get('Jenis_Alat',''), r.get('Capacity_Num',0)), axis=1)

        for df in [df_agg, df_monthly]:
            if 'Total_Ton' in df.columns and 'LITER' in df.columns:
                mask = df['Total_Ton'] > 0
                df.loc[mask, 'Fuel Ratio (L/Ton)'] = df.loc[mask, 'LITER'] / df.loc[mask, 'Total_Ton']
                df.loc[~mask, 'Fuel Ratio (L/Ton)'] = 0

        benchmark = df_agg[df_agg['Total_Ton'] > 0].groupby('Benchmark_Group')['Fuel Ratio (L/Ton)'].median().reset_index()
        benchmark.rename(columns={'Fuel Ratio (L/Ton)': 'Benchmark (L/Ton)'}, inplace=True)
        df_agg = pd.merge(df_agg, benchmark, on='Benchmark_Group', how='left')

        def get_status(row):
            if row['Total_Ton'] <= 0: return "Inaktif"
            return "Efisien" if row['Fuel Ratio (L/Ton)'] <= row['Benchmark (L/Ton)'] else "Boros"

        df_agg['Status'] = df_agg.apply(get_status, axis=1)
        df_agg['Potensi Pemborosan BBM (L)'] = df_agg.apply(
            lambda r: (r['Fuel Ratio (L/Ton)'] - r['Benchmark (L/Ton)']) * r['Total_Ton'] if r['Status'] == 'Boros' else 0, axis=1)

        rename_map = {
            'Unit_Name': 'Nama Unit', 'Jenis_Alat': 'Jenis', 'Type_Merk': 'Type/Merk',
            'Horse_Power': 'Horse Power', 'Capacity': 'Capacity (Ton)',
            'LITER': 'Total Pengisian BBM (L)', 'Total_Ton': 'Total Berat Angkutan (Ton)'
        }
        df_agg.rename(columns=rename_map, inplace=True)
        df_monthly.rename(columns=rename_map, inplace=True)

        if not df_missing.empty:
            if 'Capacity (Ton)' not in df_missing.columns:
                temp_agg = df_agg[['Nama Unit', 'Capacity (Ton)']].drop_duplicates()
                if 'Nama Unit' in df_missing.columns:
                    df_missing = pd.merge(df_missing, temp_agg, on='Nama Unit', how='left')
                    df_missing['Capacity (Ton)'] = df_missing['Capacity (Ton)'].fillna(0)

        return df_agg, df_monthly, df_missing
    except Exception as e:
        st.error(f"Error memproses data Non-Trucking: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# 5. LOGIKA PROSES DATA: TRUCKING
@st.cache_data(show_spinner=False)
def process_trucking():
    master_dict = {}
    if os.path.exists(FILE_MASTER_REF):
        try:
            df_map    = pd.read_excel(FILE_MASTER_REF, sheet_name='Sheet2', header=1)
            col_name  = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
            col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
            col_type  = next((c for c in df_map.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
            col_loc   = next((c for c in df_map.columns if 'LOKASI' in str(c).upper() or 'DES 2025' in str(c).upper()), df_map.columns[2])
            col_hp    = next((c for c in df_map.columns if 'HP' in str(c).upper() or 'HORSE' in str(c).upper()), None)
            if col_name:
                for _, row in df_map.iterrows():
                    u_name = str(row[col_name]).strip().upper()
                    jenis  = str(row[col_jenis]).strip().upper() if col_jenis else ""
                    if "TRONTON" in jenis or "TRAILER" in jenis or "HEAD" in jenis:
                        c_id = clean_unit_name(u_name)
                        master_dict[c_id] = {
                            'Real_Name': u_name, 'Jenis': jenis,
                            'Type/Merk': str(row[col_type]).strip() if col_type else "-",
                            'Lokasi': str(row[col_loc]).strip() if col_loc else "-",
                            'Horse Power': row[col_hp] if col_hp else "-", 'Capacity': 40
                        }
        except Exception as e:
            st.error(f"Gagal membaca Master File: {e}")

    df_trucking = pd.DataFrame()
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_raw    = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='HASIL_ANALISA')
            valid_rows = []
            for _, row in df_raw.iterrows():
                raw_name  = str(row['Nama_Unit']) if 'Nama_Unit' in row else str(row.get('EQUIP NAME', ''))
                match_key = get_smart_match(raw_name, master_dict)
                if match_key:
                    meta = master_dict[match_key]
                    valid_rows.append({
                        'Nama Unit': meta['Real_Name'], 'Jenis': meta['Jenis'], 'Type/Merk': meta['Type/Merk'],
                        'Lokasi': meta['Lokasi'], 'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40,
                        'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Berat Angkutan (Ton)': row.get('Total_Ton', 0),
                        'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0),
                        'Fuel Ratio (L/Ton*Km)': row.get('L_per_TonKm', 0)
                    })
            df_trucking = pd.DataFrame(valid_rows)
            if not df_trucking.empty:
                col_ratio = 'Fuel Ratio (L/Ton*Km)'
                col_work  = 'Total Kerja (Ton*Km)'
                median_ratio = df_trucking[df_trucking[col_ratio] > 0][col_ratio].median()
                df_trucking['Benchmark (L/Ton*Km)'] = median_ratio
                df_trucking['Status'] = df_trucking.apply(
                    lambda x: "Efisien" if x[col_ratio] <= x['Benchmark (L/Ton*Km)'] else "Boros", axis=1)
                df_trucking['Potensi Pemborosan BBM (L)'] = df_trucking.apply(
                    lambda r: (r[col_ratio] - r['Benchmark (L/Ton*Km)']) * r[col_work] if r['Status'] == 'Boros' else 0, axis=1)
        except Exception as e:
            st.error(f"Gagal memproses data trucking utama: {e}")

    df_monthly_trucking = pd.DataFrame()
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_monthly_raw = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='Data_Bulanan')
            monthly_list   = []
            for _, row in df_monthly_raw.iterrows():
                raw_name  = str(row['Nama_Unit'])
                match_key = get_smart_match(raw_name, master_dict)
                if match_key:
                    meta = master_dict[match_key]
                    monthly_list.append({
                        'Nama Unit': meta['Real_Name'], 'Bulan': str(row['Bulan']).capitalize(),
                        'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0),
                        'Jenis': meta['Jenis'], 'Type/Merk': meta['Type/Merk'],
                        'Lokasi': meta['Lokasi'], 'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40
                    })
            if monthly_list: df_monthly_trucking = pd.DataFrame(monthly_list)
        except Exception:
            pass

    df_missing_truck = pd.DataFrame()
    list_audit = []
    if os.path.exists(FILE_HASIL_TRUCKING):
        for sheet in ['OPS_TANPA_BBM', 'BBM_TANPA_OPS', 'GAGAL_MAPPING']:
            try:
                df_aud = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name=sheet)
                col_n  = 'Nama Unit' if 'Nama Unit' in df_aud.columns else ('Nama_Unit' if 'Nama_Unit' in df_aud.columns else 'Kode_Lambung')
                for _, row in df_aud.iterrows():
                    raw_u     = str(row.get(col_n, ''))
                    match_key = get_smart_match(raw_u, master_dict)
                    if match_key:
                        meta = master_dict[match_key]
                        list_audit.append({
                            'Nama Unit': meta['Real_Name'], 'Jenis': meta['Jenis'],
                            'Type/Merk': meta['Type/Merk'], 'Lokasi': meta['Lokasi'],
                            'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40,
                            'Total Pengisian BBM (L)': row.get('LITER', 0),
                            'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0),
                            'Keterangan': f"Inaktif ({sheet})"
                        })
            except:
                pass
    if list_audit: df_missing_truck = pd.DataFrame(list_audit)
    return df_trucking, df_monthly_trucking, df_missing_truck

# 6. SIDEBAR & MENU NAVIGASI
st.sidebar.subheader("Menu Navigasi")
category_filter = st.sidebar.radio(
    "Pilih Fitur Aplikasi:",
    ["Forecast Data"]
    #["Analisa Trucking", "Analisa Non-Trucking", "Forecast Data"]
)

st.sidebar.markdown("---")

# BAGIAN A: MENU FORECASTING (ARIMA + GRADIENT BOOSTING + ENSEMBLE)
if category_filter == "Forecast Data":
    st.header("Forecast Hour Meter & Kebutuhan BBM")

    #PANDUAN FORMAT & TEMPLATE
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

        st.markdown("**Konsumsi BBM Standar Pabrik (Opsional)**")
        st.markdown(
            "- Format file wajib **`.xlsx`**.\n"
            "- Hanya upload **1 file** (jika diupload).\n"
            "- File ini **tidak dibuat manual**, file ini dihasilkan secara otomatis dari script "
            "`forecastHMArimaGBEnsemble.py` saat dijalankan dengan `MODE = 'crawl'`. "
            "Selama format output script crawling tidak diubah, file yang dihasilkan dapat "
            "langsung digunakan tanpa modifikasi apapun.\n"
            "- Data harus berada di **sheet pertama** file, apapun nama sheet-nya.\n"
            "- Jika file ini tidak diupload, kolom standar pabrik pada hasil forecast akan dikosongkan "
            "dan perbandingan vs standar pabrik tidak akan tersedia."
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
        f_standar = st.file_uploader(
            "4. Konsumsi BBM Standar Pabrik (Opsional, 1 file)",
            type=["xlsx"], accept_multiple_files=False, key="fcst_standar"
        )
    if f_standar is None:
        st.info("ℹ️ File standar pabrik tidak diupload. Kolom standar pabrik akan dikosongkan.")

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

                    # LOAD STANDAR PABRIK
                    standar_per_key = {}
                    if f_standar is not None:
                        try:
                            f_standar.seek(0)
                            df_sp = pd.read_excel(f_standar)
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
                                standar_per_key[key] = {
                                    'Standar Pabrik Konsumsi BBM Per Jam':     row.get('Standar Pabrik Konsumsi BBM Per Jam'),
                                    'Standar Pabrik Konsumsi BBM Min (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Min (L/Jam)'),
                                    'Standar Pabrik Konsumsi BBM Max (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Max (L/Jam)'),
                                    'Sumber Data Standar Pabrik':              row.get('Sumber Data Standar Pabrik', '-'),
                                    'Catatan Standar Pabrik':                  row.get('Catatan Standar Pabrik', '-'),
                                }
                        except Exception as e_sp:
                            st.warning(f"Gagal membaca file standar pabrik: {e_sp}")

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
                        mode='lines+markers', name='Aktual 2023-2024',
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
                        mode='lines+markers', name='Aktual 2025',
                        line=dict(color='#00cc96', width=2), marker=dict(size=7, symbol='square')
                    ))

                    fig_line.add_trace(go.Scatter(
                        x=[last_bln] + bulan_2025,
                        y=[last_val] + safe_float_list(df_aktual_unit[y_tp]),
                        mode='lines+markers', name='Proyeksi 2025',
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

# BAGIAN B: DASHBOARD UTAMA (TRUCKING / NON-TRUCKING)
else:
    BIAYA_PER_LITER = st.sidebar.number_input("Biaya Bahan Bakar (Rp/Liter)", min_value=0, value=6800, step=100)

    df_active_raw = pd.DataFrame()
    df_monthly    = pd.DataFrame()
    df_missing    = pd.DataFrame()

    if category_filter == "Analisa Trucking":
        with st.spinner("Memproses Data Trucking..."):
            df_active_raw, df_monthly, df_missing = process_trucking()
            mode_label  = "Trucking"
            ratio_label = "L/Ton*Km"
            work_col    = "Total Kerja (Ton*Km)"
    else:
        with st.spinner("Memuat Data Non-Trucking..."):
            df_active_raw, df_monthly, df_missing = process_alat_berat()
            mode_label  = "Non-Trucking"
            ratio_label = "L/Ton"
            work_col    = "Total Berat Angkutan (Ton)"

    if not df_active_raw.empty:

        df_inaktif_from_active = df_active_raw[
            (df_active_raw['Total Pengisian BBM (L)'] <= 0) |
            (df_active_raw[work_col] <= 0)
        ].copy()

        if not df_inaktif_from_active.empty:
            col_bbm    = 'Total Pengisian BBM (L)'
            conditions = [
                (df_inaktif_from_active[col_bbm] <= 0) & (df_inaktif_from_active[work_col] <= 0),
                (df_inaktif_from_active[work_col] <= 0),
                (df_inaktif_from_active[col_bbm] <= 0)
            ]
            choices = ["Tidak ada aktivitas", "Unit tidak melakukan aktivitas kerja", "Unit tidak pernah mengisi BBM"]
            df_inaktif_from_active['Keterangan'] = np.select(conditions, choices, default="-")

        list_inaktif = []
        if not df_inaktif_from_active.empty: list_inaktif.append(df_inaktif_from_active)

        if not df_missing.empty:
            if 'Total Pengisian BBM' in df_missing.columns:
                df_missing.rename(columns={'Total Pengisian BBM': 'Total Pengisian BBM (L)'}, inplace=True)
            m_bbm_col  = 'Total Pengisian BBM (L)'
            m_work_col = work_col
            if m_bbm_col in df_missing.columns and m_work_col in df_missing.columns:
                m_conds = [
                    (df_missing[m_bbm_col] <= 0) & (df_missing[m_work_col] <= 0),
                    (df_missing[m_work_col] <= 0),
                    (df_missing[m_bbm_col] <= 0)
                ]
                m_choices = ["Tidak ada aktivitas", "Unit tidak melakukan aktivitas kerja", "Unit tidak pernah mengisi BBM"]
                df_missing['Keterangan'] = np.select(m_conds, m_choices, default="Inaktif (Sumber: File Audit)")
            else:
                df_missing['Keterangan'] = "Inaktif (Sumber: File Audit)"
            list_inaktif.append(df_missing)

        df_inaktif_all = pd.concat(list_inaktif, ignore_index=True) if list_inaktif else pd.DataFrame()
        df_active      = df_active_raw[(df_active_raw['Total Pengisian BBM (L)'] > 0) & (df_active_raw[work_col] > 0)].copy()
        df_full_filter = pd.concat([df_active, df_inaktif_all], ignore_index=True) if not df_inaktif_all.empty else df_active

        st.sidebar.markdown("---")
        st.sidebar.header("Filter Data")
        lokasi_list   = ["Semua"] + sorted(df_full_filter['Lokasi'].dropna().unique().tolist())
        selected_lok  = st.sidebar.selectbox("📍 Filter Lokasi", lokasi_list)
        jenis_list    = ["Semua"] + sorted(df_full_filter['Jenis'].dropna().unique().tolist())
        selected_jen  = st.sidebar.selectbox("🚜 Filter Jenis", jenis_list)
        type_list     = ["Semua"] + sorted(df_full_filter['Type/Merk'].dropna().astype(str).unique().tolist())
        selected_type = st.sidebar.selectbox("🏷️ Filter Type/Merk", type_list)

        st.markdown("### 🛑 Daftar Unit Inaktif")
        st.caption("Unit yang terdeteksi tidak aktif karena tidak ada pengisian BBM atau tidak ada aktivitas kerja")
        df_inaktif_f = df_inaktif_all.copy()
        if not df_inaktif_f.empty:
            if selected_lok  != "Semua": df_inaktif_f = df_inaktif_f[df_inaktif_f['Lokasi'] == selected_lok]
            if selected_jen  != "Semua": df_inaktif_f = df_inaktif_f[df_inaktif_f['Jenis'] == selected_jen]
            if selected_type != "Semua": df_inaktif_f = df_inaktif_f[df_inaktif_f['Type/Merk'] == selected_type]
            if not df_inaktif_f.empty:
                if mode_label == "Trucking":
                    cols_in = ['Nama Unit','Jenis','Type/Merk','Lokasi','Horse Power','Capacity (Feet)','Total Pengisian BBM (L)',work_col,'Keterangan']
                    fmt_in  = {'Capacity (Feet)': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}'}
                else:
                    cols_in = ['Nama Unit','Jenis','Type/Merk','Lokasi','Horse Power','Capacity (Ton)','Total Pengisian BBM (L)',work_col,'Keterangan']
                    fmt_in  = {'Capacity (Ton)': '{:.0f}', 'Horse Power': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}'}
                cols_show = [c for c in cols_in if c in df_inaktif_f.columns]
                st.dataframe(df_inaktif_f[cols_show].style.format(fmt_in, na_rep="-"), use_container_width=True)
            else:
                st.success("Tidak ada unit inaktif untuk kombinasi filter ini.")
        else:
            st.success("Seluruh unit beroperasi aktif.")
        st.markdown("---")

        st.markdown("### 🔍 Cari Data Spesifik (Unit Aktif)")
        cs1, cs2 = st.columns([1, 3])
        with cs1:  search_cat = st.selectbox("Cari Berdasarkan:", ["Nama Unit"])
        with cs2:  search_q   = st.text_input(f"Ketik {search_cat}:", "")

        df_filtered = df_active.copy()
        if selected_lok  != "Semua": df_filtered = df_filtered[df_filtered['Lokasi'] == selected_lok]
        if selected_jen  != "Semua": df_filtered = df_filtered[df_filtered['Jenis'] == selected_jen]
        if selected_type != "Semua": df_filtered = df_filtered[df_filtered['Type/Merk'] == selected_type]
        if search_q:
            df_filtered = df_filtered[df_filtered['Nama Unit'].str.contains(search_q, case=False, na=False)]

        df_filtered['Total Biaya BBM'] = df_filtered['Total Pengisian BBM (L)'] * BIAYA_PER_LITER

        if not df_monthly.empty:
            df_monthly_filtered = df_monthly[df_monthly['Nama Unit'].isin(df_filtered['Nama Unit'])].copy()
        else:
            df_monthly_filtered = pd.DataFrame()

        if mode_label == "Trucking":
            total_bbm   = df_filtered['Total Pengisian BBM (L)'].sum()
            total_kerja = df_filtered['Total Kerja (Ton*Km)'].sum()
            total_biaya = df_filtered['Total Biaya BBM'].sum()
            total_aktif = len(df_filtered)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Unit Aktif",     f"{total_aktif} Unit")
            c2.metric("Total Kerja (Ton*Km)", f"{total_kerja:,.0f}")
            c3.metric("Total Pengisian BBM",  f"{total_bbm:,.0f} L")
            c4.metric("Total Biaya BBM (Rp)", f"Rp {total_biaya:,.0f}")
        else:
            total_bbm   = df_filtered['Total Pengisian BBM (L)'].sum()
            total_ton   = df_filtered['Total Berat Angkutan (Ton)'].sum()
            total_biaya = df_filtered['Total Biaya BBM'].sum()
            total_aktif = len(df_filtered)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Unit Aktif",       f"{total_aktif} Unit")
            c2.metric("Total Tonase Container", f"{total_ton:,.0f} Ton")
            c3.metric("Total Pengisian BBM",    f"{total_bbm:,.0f} L")
            c4.metric("Total Biaya BBM (Rp)",   f"Rp {total_biaya:,.0f}")

        st.markdown("---")

        tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview Data", "📈 Efisiensi Setiap Unit", "📍 Korelasi Beban & BBM", "💸 Unit Terboros"])

        ratio_col = 'Fuel Ratio (L/Ton*Km)' if mode_label == "Trucking" else 'Fuel Ratio (L/Ton)'
        bm_col    = 'Benchmark (L/Ton*Km)'  if mode_label == "Trucking" else 'Benchmark (L/Ton)'

        def highlight_fuel_ratio(row):
            styles = [''] * len(row)
            for i, col in enumerate(row.index):
                if col == ratio_col:
                    val, bm = row[col], row[bm_col]
                    if pd.notna(val) and pd.notna(bm) and bm > 0:
                        styles[i] = ('background-color: #d62728; color: white; font-weight: bold;'
                                     if val > bm else
                                     'background-color: #2ca02c; color: white; font-weight: bold;')
            return styles

        with tab1:
            st.subheader(f"Data Detail {mode_label}")
            sort_opts = ["Fuel Ratio (Tertinggi)", "Fuel Ratio (Terendah)", "Total Kerja (Tertinggi)", "Total Pengisian BBM (L) (Tertinggi)"]
            sort_by   = st.selectbox("Sort by:", sort_opts)
            if sort_by == "Fuel Ratio (Tertinggi)":                df_filtered = df_filtered.sort_values(ratio_col, ascending=False)
            elif sort_by == "Fuel Ratio (Terendah)":               df_filtered = df_filtered.sort_values(ratio_col, ascending=True)
            elif sort_by == "Total Kerja (Tertinggi)":             df_filtered = df_filtered.sort_values(work_col,  ascending=False)
            elif sort_by == "Total Pengisian BBM (L) (Tertinggi)": df_filtered = df_filtered.sort_values('Total Pengisian BBM (L)', ascending=False)

            if mode_label == "Trucking":
                cols_s = ['Nama Unit','Jenis','Lokasi','Horse Power','Capacity (Feet)','Total Pengisian BBM (L)','Total Biaya BBM','Total Berat Angkutan (Ton)','Total Kerja (Ton*Km)','Benchmark (L/Ton*Km)','Fuel Ratio (L/Ton*Km)','Potensi Pemborosan BBM (L)']
                fmt_d  = {'Capacity (Feet)':'{:.0f}','Total Pengisian BBM (L)':'{:,.0f}','Total Biaya BBM':'Rp {:,.0f}','Total Berat Angkutan (Ton)':'{:,.0f}','Total Kerja (Ton*Km)':'{:,.0f}','Benchmark (L/Ton*Km)':'{:.4f}','Fuel Ratio (L/Ton*Km)':'{:.4f}','Potensi Pemborosan BBM (L)':'{:,.0f}'}
                st.dataframe(df_filtered[cols_s].style.apply(highlight_fuel_ratio, axis=1).format(fmt_d))
            else:
                cols_s = ['Nama Unit','Jenis','Type/Merk','Horse Power','Capacity (Ton)','Lokasi','Total Pengisian BBM (L)','Total Biaya BBM','Total Berat Angkutan (Ton)','Benchmark (L/Ton)','Fuel Ratio (L/Ton)','Potensi Pemborosan BBM (L)']
                fmt_d  = {'Total Pengisian BBM (L)':'{:,.0f}','Total Biaya BBM':'Rp {:,.0f}','Total Berat Angkutan (Ton)':'{:,.0f}','Benchmark (L/Ton)':'{:.4f}','Fuel Ratio (L/Ton)':'{:.4f}','Potensi Pemborosan BBM (L)':'{:,.0f}'}
                cols_ok = [c for c in cols_s if c in df_filtered.columns]
                st.dataframe(df_filtered[cols_ok].style.apply(highlight_fuel_ratio, axis=1).format(fmt_d, na_rep="-"))

        with tab2:
            st.subheader(f"Efisiensi BBM per Unit ({mode_label})")
            if not df_filtered.empty:
                df_eff = df_filtered[[ratio_col, 'Nama Unit', bm_col, 'Status']].dropna(subset=[ratio_col])
                df_eff = df_eff[df_eff[ratio_col] > 0].sort_values(ratio_col, ascending=False)
                fig_eff = px.bar(df_eff, x='Nama Unit', y=ratio_col,
                                 color='Status', color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                 title=f"Fuel Ratio per Unit ({ratio_label})")
                if not df_eff.empty:
                    fig_eff.add_hline(y=df_eff[bm_col].mean(), line_dash="dash", line_color="blue",
                                      annotation_text=f"Benchmark: {df_eff[bm_col].mean():.4f}")
                fig_eff.update_layout(xaxis_tickangle=-45, height=500)
                st.plotly_chart(fig_eff, use_container_width=True)

        with tab3:
            st.subheader("Korelasi Beban Kerja & BBM")
            if not df_filtered.empty and work_col in df_filtered.columns:
                fig_sc = px.scatter(df_filtered, x=work_col, y='Total Pengisian BBM (L)',
                                    color='Status', hover_name='Nama Unit', size='Total Pengisian BBM (L)',
                                    color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                    title=f"Korelasi {work_col} vs Total BBM")
                st.plotly_chart(fig_sc, use_container_width=True)

        with tab4:
            st.subheader("Unit dengan Potensi Pemborosan BBM Tertinggi")
            df_boros = df_filtered[df_filtered['Status'] == 'Boros'].sort_values('Potensi Pemborosan BBM (L)', ascending=False)
            if not df_boros.empty:
                fig_boros = px.bar(df_boros.head(20), x='Nama Unit', y='Potensi Pemborosan BBM (L)',
                                   color='Potensi Pemborosan BBM (L)',
                                   color_continuous_scale='Reds',
                                   title="Top 20 Unit dengan Potensi Pemborosan BBM Terbesar")
                fig_boros.update_layout(xaxis_tickangle=-45, height=500)
                st.plotly_chart(fig_boros, use_container_width=True)

                total_boros = df_boros['Potensi Pemborosan BBM (L)'].sum()
                biaya_boros = total_boros * BIAYA_PER_LITER
                b1, b2 = st.columns(2)
                b1.metric("Total Potensi Pemborosan (L)",  f"{total_boros:,.0f} L")
                b2.metric("Total Potensi Pemborosan (Rp)", f"Rp {biaya_boros:,.0f}")
            else:
                st.success("Tidak ada unit yang terdeteksi boros berdasarkan filter saat ini.")