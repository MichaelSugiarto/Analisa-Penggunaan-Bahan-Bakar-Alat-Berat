import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import re
import warnings
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
import io
import time
import requests
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

warnings.filterwarnings('ignore')

# ==============================================================================
# 1. KONFIGURASI FILE & PATH
# ==============================================================================
FILE_HASIL_TRUCKING = "HasilTrucking.xlsx" 
FILE_HASIL_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_BBM_RAW = "BBM AAB.xlsx"
FILE_HAULAGE_RAW = "HAULAGE OKT-DES 2025 (Copy).xlsx"
FILE_DOORING_REVISI = "DOORING_WITH_DISTANCE_REVISI.xlsx" 
FILE_MASTER_REF = "cost & bbm 2022 sd 2025 HP & Type.xlsx"

# Inisialisasi Session State agar tombol download dan grafik tidak hilang saat diklik
if "etl_step1_processed" not in st.session_state:
    st.session_state.etl_step1_processed = False
    st.session_state.out_dooring_file = None

if "etl_step2_processed" not in st.session_state:
    st.session_state.etl_step2_processed = False
    st.session_state.out_truck_file = None
    st.session_state.out_nontruck_file = None

if "forecast_processed" not in st.session_state:
    st.session_state.forecast_processed = False
    st.session_state.fcst_df_res = None
    st.session_state.fcst_df_final = None
    st.session_state.fcst_out_file = None

# ==============================================================================
# 2. SETUP HALAMAN
# ==============================================================================
st.set_page_config(page_title="Dashboard Efisiensi BBM", layout="wide")
st.title("Dashboard BBM Alat Berat")

# ==============================================================================
# 3. FUNGSI UTILITIES 
# ==============================================================================
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

# ==============================================================================
# 4. LOGIKA PROSES DATA: NON-TRUCKING 
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_alat_berat():
    if not os.path.exists(FILE_HASIL_NON_TRUCKING):
        st.warning(f"File {FILE_HASIL_NON_TRUCKING} tidak ditemukan. Jalankan Proses Data Awal terlebih dahulu.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        df_agg = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Total_Agregat')
        df_monthly = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Data_Bulanan')
        try:
            df_missing = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Unit_Inaktif')
            rename_missing = {
                'Unit_Name': 'Nama Unit', 'Jenis_Alat': 'Jenis', 'Type_Merk': 'Type/Merk',
                'Horse_Power': 'Horse Power', 'Capacity': 'Capacity (Ton)', 'LITER': 'Total Pengisian BBM (L)',
                'Total_Ton': 'Total Berat Angkutan (Ton)', 'Total Pengisian BBM': 'Total Pengisian BBM (L)' 
            }
            df_missing.rename(columns=rename_missing, inplace=True)
        except:
            df_missing = pd.DataFrame()
        
        df_agg.columns = df_agg.columns.str.strip()
        df_monthly.columns = df_monthly.columns.str.strip()
        df_agg['Capacity_Num'] = df_agg['Capacity'].fillna(0).astype(float).astype(int)
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
                elif cap >= 40: return 'Crane (Capacity 41)'
                return 'Crane (Lainnya)'
            return 'Lainnya'

        df_agg['Benchmark_Group'] = df_agg.apply(lambda r: get_benchmark_group(r['Jenis_Alat'], r['Capacity_Num']), axis=1)
        df_agg['Capacity'] = df_agg['Capacity_Num']
        df_monthly['Capacity'] = df_monthly['Capacity_Num']
        
        df_agg['Fuel Ratio (L/Ton)'] = np.where(df_agg['Total_Ton'] > 0, df_agg['LITER'] / df_agg['Total_Ton'], 0)
        df_monthly['Fuel Ratio (L/Ton)'] = np.where(df_monthly['Total_Ton'] > 0, df_monthly['LITER'] / df_monthly['Total_Ton'], 0)
        
        benchmark = df_agg[df_agg['Total_Ton'] > 0].groupby('Benchmark_Group')['Fuel Ratio (L/Ton)'].median().reset_index()
        benchmark.rename(columns={'Fuel Ratio (L/Ton)': 'Benchmark (L/Ton)'}, inplace=True)
        df_agg = pd.merge(df_agg, benchmark, on='Benchmark_Group', how='left')
        
        def get_status(row):
            if row['Total_Ton'] <= 0: return "Inaktif"
            return "Efisien" if row['Fuel Ratio (L/Ton)'] <= row['Benchmark (L/Ton)'] else "Boros"
            
        df_agg['Status'] = df_agg.apply(get_status, axis=1)
        df_agg['Potensi Pemborosan BBM (L)'] = df_agg.apply(
            lambda r: (r['Fuel Ratio (L/Ton)'] - r['Benchmark (L/Ton)']) * r['Total_Ton'] if r['Status'] == 'Boros' else 0, axis=1
        )

        rename_map = {
            'Unit_Name': 'Nama Unit', 'Jenis_Alat': 'Jenis', 'Type_Merk': 'Type/Merk',
            'Horse_Power': 'Horse Power', 'Capacity': 'Capacity (Ton)', 'LITER': 'Total Pengisian BBM (L)', 'Total_Ton': 'Total Berat Angkutan (Ton)'
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

# ==============================================================================
# 5. LOGIKA PROSES DATA: TRUCKING 
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_trucking():
    master_dict = {}
    if os.path.exists(FILE_MASTER_REF):
        try:
            df_map = pd.read_excel(FILE_MASTER_REF, sheet_name='Sheet2', header=1)
            col_name = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
            col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
            col_type = next((c for c in df_map.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
            col_loc = next((c for c in df_map.columns if 'LOKASI' in str(c).upper() or 'DES 2025' in str(c).upper()), df_map.columns[2])
            col_hp = next((c for c in df_map.columns if 'HP' in str(c).upper() or 'HORSE' in str(c).upper()), None)
            
            if col_name:
                for _, row in df_map.iterrows():
                    u_name = str(row[col_name]).strip().upper()
                    jenis = str(row[col_jenis]).strip().upper() if col_jenis else ""
                    if "TRONTON" in jenis or "TRAILER" in jenis or "HEAD" in jenis:
                        c_id = clean_unit_name(u_name)
                        master_dict[c_id] = {
                            'Real_Name': u_name, 'Jenis': jenis, 'Type/Merk': str(row[col_type]).strip() if col_type else "-",
                            'Lokasi': str(row[col_loc]).strip() if col_loc else "-", 'Horse Power': row[col_hp] if col_hp else "-", 'Capacity': 40 
                        }
        except Exception as e:
            st.error(f"Gagal membaca Master File: {e}")
    
    df_trucking = pd.DataFrame()
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_raw = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='HASIL_ANALISA')
            valid_rows = []
            for _, row in df_raw.iterrows():
                raw_name = str(row['Nama_Unit']) if 'Nama_Unit' in row else str(row.get('EQUIP NAME', ''))
                match_key = get_smart_match(raw_name, master_dict)
                if match_key:
                    meta = master_dict[match_key]
                    valid_rows.append({
                        'Nama Unit': meta['Real_Name'], 'Jenis': meta['Jenis'], 'Type/Merk': meta['Type/Merk'], 'Lokasi': meta['Lokasi'],
                        'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40, 'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Berat Angkutan (Ton)': row.get('Total_Ton', 0), 'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0), 
                        'Fuel Ratio (L/Ton*Km)': row.get('L_per_TonKm', 0) 
                    })
            df_trucking = pd.DataFrame(valid_rows)
            if not df_trucking.empty:
                col_ratio = 'Fuel Ratio (L/Ton*Km)'
                col_work = 'Total Kerja (Ton*Km)'
                median_ratio = df_trucking[df_trucking[col_ratio] > 0][col_ratio].median()
                df_trucking['Benchmark (L/Ton*Km)'] = median_ratio
                df_trucking['Status'] = df_trucking.apply(lambda x: "Efisien" if x[col_ratio] <= x['Benchmark (L/Ton*Km)'] else "Boros", axis=1)
                df_trucking['Potensi Pemborosan BBM (L)'] = df_trucking.apply(lambda r: (r[col_ratio] - r['Benchmark (L/Ton*Km)']) * r[col_work] if r['Status'] == 'Boros' else 0, axis=1)
        except Exception as e:
            st.error(f"Gagal memproses data trucking utama: {e}")

    df_monthly_trucking = pd.DataFrame()
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_monthly_raw = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='Data_Bulanan')
            monthly_list = []
            for _, row in df_monthly_raw.iterrows():
                raw_name = str(row['Nama_Unit'])
                match_key = get_smart_match(raw_name, master_dict)
                if match_key:
                    meta = master_dict[match_key]
                    monthly_list.append({
                        'Nama Unit': meta['Real_Name'], 'Bulan': str(row['Bulan']).capitalize(), 'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0), 'Jenis': meta['Jenis'], 'Type/Merk': meta['Type/Merk'],
                        'Lokasi': meta['Lokasi'], 'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40
                    })
            if monthly_list: df_monthly_trucking = pd.DataFrame(monthly_list)
        except Exception as e: pass

    df_missing_truck = pd.DataFrame()
    list_audit = []
    if os.path.exists(FILE_HASIL_TRUCKING):
        for sheet in ['OPS_TANPA_BBM', 'BBM_TANPA_OPS', 'GAGAL_MAPPING']:
            try:
                df_aud = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name=sheet)
                col_n = 'Nama Unit' if 'Nama Unit' in df_aud.columns else ('Nama_Unit' if 'Nama_Unit' in df_aud.columns else 'Kode_Lambung')
                for _, row in df_aud.iterrows():
                    raw_u = str(row.get(col_n, ''))
                    match_key = get_smart_match(raw_u, master_dict)
                    if match_key:
                        meta = master_dict[match_key]
                        list_audit.append({
                            'Nama Unit': meta['Real_Name'], 'Jenis': meta['Jenis'], 'Type/Merk': meta['Type/Merk'], 'Lokasi': meta['Lokasi'],
                            'Horse Power': meta['Horse Power'], 'Capacity (Feet)': 40, 'Total Pengisian BBM (L)': row.get('LITER', 0),
                            'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0), 'Keterangan': f"Inaktif ({sheet})"
                        })
            except: pass
            
    if list_audit: df_missing_truck = pd.DataFrame(list_audit)
    return df_trucking, df_monthly_trucking, df_missing_truck

# ==============================================================================
# 6. SIDEBAR & MENU NAVIGASI
# ==============================================================================
st.sidebar.subheader("Menu Navigasi")
category_filter = st.sidebar.radio(
    "Pilih Fitur Aplikasi:", 
    ["Analisa Trucking", "Analisa Non-Trucking", "Proses Data Awal", "Proses Forecast"]
)

st.sidebar.markdown("---")

# ------------------------------------------------------------------------------
# BAGIAN A: MENU INPUT DATA MENTAH (ETL) DENGAN 2 LANGKAH
# ------------------------------------------------------------------------------
if category_filter == "Proses Data Awal":
    st.header("Proses Data Awal")
    st.markdown("Bagian ini bertujuan untuk memproses file data raw menjadi **HasilTrucking.xlsx** dan **HasilNonTrucking.xlsx** yang diperlukan untuk melakukan forecasting")
    
    # ---------------------------------------------------------
    # LANGKAH 1: OSRM DOORING (MENGGUNAKAN LOGIKA ASLI USER)
    # ---------------------------------------------------------
    st.markdown("### 🔹 Langkah 1: Perhitungan Jarak & Berat Keperluan Dooring")
    st.info("Upload file kebutuhan Dooring untuk menghitung jarak yang ditempuh menggunakan OSRM, mengonversi berat berdasarkan jenis kontainer, dan memberikan output dalam format .xlsx untuk proses selanjutnya.")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1: f_dooring_mentah = st.file_uploader("Upload DOORING OKT-DES 2025 (Copy).xlsx", type=["xlsx", "csv"], key="drg_step1")
    with col_d2: f_tlp_mentah = st.file_uploader("Upload dashboard TLP okt-des (S1L Trucking).xlsx", type=["xlsx", "csv"], key="tlp_step1")
    
    if st.button("🗺️ 1. Hitung Jarak & Berat Dooring"):
        if f_dooring_mentah and f_tlp_mentah:
            with st.spinner("Menghitung Jarak OSRM dan Berat... (Jika pertama kali, akan memakan waktu agak lama)"):
                try:
                    f_dooring_mentah.seek(0)
                    f_tlp_mentah.seek(0)
                    
                    # 1. BACA TLP (AMBIL SEMUA SHEET)
                    tlp_sheets = pd.ExcelFile(f_tlp_mentah)
                    tlp_frames = []
                    for sht in tlp_sheets.sheet_names:
                        df_sht = pd.read_excel(tlp_sheets, sheet_name=sht)
                        col_sopt_tlp = 'SOPT NO' if 'SOPT NO' in df_sht.columns else ('SOPT_NO' if 'SOPT_NO' in df_sht.columns else None)
                        col_alamat_tlp = 'PICKUP / DELIVERY ADDRESS' if 'PICKUP / DELIVERY ADDRESS' in df_sht.columns else ('DOORING_ADDRESS' if 'DOORING_ADDRESS' in df_sht.columns else None)
                        if col_sopt_tlp and col_alamat_tlp:
                            df_sht = df_sht[[col_sopt_tlp, col_alamat_tlp]].rename(columns={col_sopt_tlp: 'SOPT_NO', col_alamat_tlp: 'ALAMAT'})
                            tlp_frames.append(df_sht)
                    df_tlp = pd.concat(tlp_frames, ignore_index=True) if tlp_frames else pd.DataFrame(columns=['SOPT_NO', 'ALAMAT'])
                    
                    # 2. BACA DOORING & MERGE (BACA SEMUA SHEET)
                    door_sheets = pd.ExcelFile(f_dooring_mentah)
                    door_frames = []
                    for sht in door_sheets.sheet_names:
                        df_sht = pd.read_excel(door_sheets, sheet_name=sht)
                        col_sopt_door = 'NO. SOPT 1' if 'NO. SOPT 1' in df_sht.columns else ('NO SOPT' if 'NO SOPT' in df_sht.columns else None)
                        if col_sopt_door:
                            df_sht['SOPT_NO_CLEAN'] = df_sht[col_sopt_door].astype(str).str.strip()
                            door_frames.append(df_sht)
                            
                    df_door = pd.concat(door_frames, ignore_index=True) if door_frames else pd.DataFrame()
                    
                    df_tlp['SOPT_NO_CLEAN'] = df_tlp['SOPT_NO'].astype(str).str.strip()
                    df_tlp_unique = df_tlp.dropna(subset=['SOPT_NO_CLEAN', 'ALAMAT']).drop_duplicates(subset=['SOPT_NO_CLEAN'])
                    df_merged = pd.merge(df_door, df_tlp_unique[['SOPT_NO_CLEAN', 'ALAMAT']], on='SOPT_NO_CLEAN', how='left')
                    
                    # Jika alamat tidak ada di TLP, maka biarkan kosong
                    df_merged['ALAMAT_FINAL'] = df_merged['ALAMAT']

                    # 3. FUNGSI OSRM DARI jarakDooringOSRMv2.py
                    DEPO_LAT = -7.2145
                    DEPO_LON = 112.7238
                    
                    def apply_manual_fixes(raw_address):
                        addr_upper = str(raw_address).upper()
                        addr_upper = addr_upper.replace("[NO.KM](HTTP://NO.KM/)", "NO.KM")
                        fixes = {
                            "VJ5J 4PF, MADURAN": "JL. RAYA ROOMO, MADURAN, ROOMO, KEC. MANYAR, KABUPATEN GRESIK, JAWA TIMUR",
                            "VJC7 HWW, TENGER": "TENGER, ROOMO, KEC. MANYAR, KABUPATEN GRESIK, JAWA TIMUR"
                        }
                        for key, correct_addr in fixes.items():
                            if key in addr_upper: return correct_addr, True 
                        return raw_address, False

                    def check_manual_bypass(raw_address):
                        addr_clean = str(raw_address).upper().replace('\n', ' ')
                        addr_clean = re.sub(r'(?i)\bINDONESIA\b', '', addr_clean)
                        addr_clean = re.sub(r'^[^\w]+|[^\w]+$', '', addr_clean).strip()
                        
                        EXACT_MATCH_COORDS = {
                            "KOTA SBY, JAWA TIMUR": (-7.2504, 112.7688), "SURABAYA, JAWA TIMUR": (-7.2504, 112.7688),
                            "KABUPATEN MOJOKERTO, JAWA TIMUR": (-7.5458, 112.4939), "MOJOKERTO REGENCY, EAST JAVA": (-7.5458, 112.4939),
                            "MOJOKERTO, JAWA TIMUR": (-7.5458, 112.4939), "PASURUAN, JAWA TIMUR": (-7.6453, 112.8208),
                            "KEC. GRESIK, KABUPATEN GRESIK, JAWA TIMUR": (-7.1610, 112.6515), "KABUPATEN GRESIK, JAWA TIMUR": (-7.1610, 112.6515)
                        }
                        if addr_clean in EXACT_MATCH_COORDS:
                            return EXACT_MATCH_COORDS[addr_clean], f"{addr_clean}"
                            
                        KEYWORD_COORDS = {
                            "DAENDELS 64-65": (-6.8778, 112.3550), "MASJID NURUL JANNAH": (-7.1585, 112.6465),         
                            "KEMIRI SEWU, PANDAAN": (-7.6440, 112.7160), "DUA KELINCI": (-6.8126, 111.0818),                 
                            "GARUDA FOOD JAYA": (-7.3685, 112.6322), "MOJOSARI-PACET KM. 6,5": (-7.5540, 112.5385),      
                            "NESTL INDONESIA - GEMPOL": (-7.5780, 112.6900), "RUNGKUT INDUSTRI I NO.16": (-7.3220, 112.7530),    
                            "RUNGKUT INDUSTRI RAYA NO.19": (-7.3320, 112.7660), "MADIUN - SURABAYA, BANJARAGUNG": (-7.4910, 112.4280),
                            "MANYARSIDORUKUN": (-7.1350, 112.6250), "PANDANREJO, REJOSO": (-7.6600, 112.9550),          
                            "SOFTEX INDONESIA SIDOARJO": (-7.4565, 112.7355), "RANGKAH KIDUL": (-7.4565, 112.7355),               
                            "DUMAR INDUSTRI": (-7.2405, 112.6685), "RAYA DLANGGU NO.KM 19": (-7.5800, 112.5000),
                            "SAWUNGGALING NO.24": (-7.3615, 112.6865), "BANJARAGUNG, KRIAN": (-7.4116, 112.5855),
                            "BUMI MASPION": (-7.2033, 112.6481), "ROMOKALISARI I": (-7.2033, 112.6481),
                            "POLEREJO, PURWOSARI": (-7.7475, 112.7335), "TJIWI KIMIA": (-7.4361, 112.4727),
                            "JL. TJ. TEMBAGA": (-7.2140, 112.7300), "JALAN NILAM TIMUR": (-7.2210, 112.7300), 
                            "JL. KALIANGET": (-7.2280, 112.7320), "DUPAK RUKUN": (-7.2410, 112.7150),
                            "KALIANAK BARAT": (-7.2240, 112.6870), "MARGOMULYO III": (-7.2450, 112.6750),
                            "PERGUDANGAN MARGOMULYO PERMAI": (-7.2450, 112.6750), "MARGOMULYO NO.65": (-7.2350, 112.6750),
                            "MARGOMULYO NO.44": (-7.2350, 112.6750), "MARGOMULYO GG.SENTONG": (-7.2510, 112.6720),
                            "JL. TANJUNGSARI": (-7.2600, 112.6850), "JL. RAYA MASTRIP": (-7.3370, 112.6950), 
                            "JL. PANJANG JIWO": (-7.3200, 112.7660), "KALI RUNGKUT": (-7.3253, 112.7663), 
                            "KENDANGSARI": (-7.3220, 112.7500), "TENGGILIS MEJOYO": (-7.3220, 112.7500), 
                            "RUNGKUT LOR": (-7.3200, 112.7660), "LINGKAR TIMUR": (-7.4560, 112.7350), 
                            "WEDI, KEC. GEDANGAN": (-7.3911, 112.7369), "JL. BERBEK INDUSTRI": (-7.3450, 112.7500), 
                            "JL. RAYA BUDURAN": (-7.4200, 112.7200), "JL. RAYA TROSOBO": (-7.3800, 112.6600), 
                            "JL. TAMBAK SAWAH": (-7.3600, 112.7600), "JL. RAYA MOJOKERTO SURABAYA": (-7.3850, 112.6700), 
                            "JL.RAYA GILANG": (-7.3750, 112.6800), "KARANGREJO, PASURUAN": (-7.6976, 112.8805), 
                            "LATEK": (-7.6045, 112.8251), "REJOSO": (-7.6534, 112.9360), 
                            "KIG RAYA SELATAN": (-7.1700, 112.6450), "GRESIK - BABAT": (-7.1180, 112.4250), 
                            "DUSUN WATES, CANGKIR": (-7.3600, 112.6100), "GUBERNUR SURYO": (-7.1650, 112.6550),
                            "RAYA KRIKILAN": (-7.3619, 112.6300), "PELEM WATU": (-7.3100, 112.5950),
                            "RAYA ROOMO": (-7.1450, 112.6400), "RAYA SEMBAYAT": (-7.1100, 112.6050),
                            "RAYA SUKOMULYO": (-7.1450, 112.6400), "BANYUTAMI": (-7.1000, 112.5800),
                            "MASPION": (-7.1250, 112.6200), "RAYA UTARA, BLOK. M/1": (-7.1450, 112.6400),
                            "KRAJAN, SUMENGKO": (-7.3750, 112.5350), "WINGS SURYA, DUSUN WATES": (-7.3600, 112.6100),
                            "NGORO": (-7.5683, 112.6171), "BARENG PROYEK": (-7.5950, 112.6950),
                            "RAYA BEJI": (-7.5950, 112.7550), "CANGKRINGMALANG": (-7.5950, 112.7350),
                            "RAYA GEMPOL": (-7.5850, 112.6950), "MALANG NO.KM 40, KECINCANG": (-7.6150, 112.6950),
                            "WICAKSONO, BANGLE": (-7.6050, 112.7250), "KALIPUTIH, SUMBERSUKO": (-7.6050, 112.6850),
                            "GROGOLAN, WINONG": (-7.5850, 112.6850)
                        }
                        for keyword, coords in KEYWORD_COORDS.items():
                            if keyword in addr_clean:
                                return coords, f"{keyword}"
                        return None, None

                    def get_osrm_distance(lat1, lon1, lat2, lon2):
                        try:
                            url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=false"
                            res = requests.get(url, timeout=10)
                            data = res.json()
                            if data.get('code') == 'Ok':
                                return data['routes'][0]['distance'] / 1000 
                        except Exception: pass
                        return 0.0

                    def build_fallback_queries(raw_address):
                        clean_addr = str(raw_address).replace('\n', ' ').strip().upper()
                        clean_addr = re.sub(r'\b\d{5}\b', '', clean_addr)
                        clean_addr = re.sub(r'(?i)\bINDONESIA\b', '', clean_addr)
                        clean_addr = re.sub(r'(?i)\bKELET-JVA\b', '', clean_addr)
                        clean_addr = re.sub(r'^[A-Z0-9]{4}\+[A-Z0-9]*\s*\,?', '', clean_addr)
                        clean_addr = re.sub(r'^[A-Z0-9]{4}\s[A-Z0-9]{3}\s*\,', '', clean_addr) 
                        clean_addr = re.sub(r'^[^\w]+', '', clean_addr)
                        parts = [p.strip() for p in clean_addr.split(',') if p.strip()]
                        queries = []
                        if parts: queries.append(", ".join(parts))                  
                        if len(parts) > 2: queries.append(", ".join(parts[1:]))     
                        if len(parts) >= 3: queries.append(", ".join(parts[-3:]))   
                        if len(parts) >= 2: queries.append(", ".join(parts[-2:]))   
                        return [q.title() for q in queries]

                    # 4. IMPLEMENTASI CACHE & API
                    cache_file = "Database_Jarak_Pabrik.xlsx"
                    if os.path.exists(cache_file):
                        df_cache = pd.read_excel(cache_file)
                    else:
                        df_cache = pd.DataFrame(columns=['ALAMAT_FINAL', 'Jarak_PP_Km', 'Alasan', 'ALAMAT YANG DIPAKAI'])
                        
                    # Filter alamat yang valid saja (bukan NaN)
                    unique_addrs = df_merged['ALAMAT_FINAL'].dropna().unique()
                    new_addrs = [addr for addr in unique_addrs if addr not in df_cache['ALAMAT_FINAL'].values and str(addr).strip() != ""]
                    
                    if len(new_addrs) > 0:
                        geolocator = Nominatim(user_agent="spil_routing_app")
                        new_data = []
                        
                        progress_text = "Memanggil API OSRM untuk alamat baru..."
                        my_bar = st.progress(0, text=progress_text)
                        
                        for i, addr in enumerate(new_addrs):
                            address_to_geocode, is_revised = apply_manual_fixes(addr)
                            
                            dist_out, dist_in, dist_pp = 0.0, 0.0, 0.0
                            status, lat, lon = "Gagal", None, None
                            used_address = "-"
                            
                            bypass_coords, bypass_keyword = check_manual_bypass(address_to_geocode)
                            
                            if bypass_coords:
                                lat, lon = bypass_coords
                                status = "Sukses (Bypass Manual)"
                                used_address = bypass_keyword 
                            else:
                                queries_to_try = build_fallback_queries(address_to_geocode)
                                if is_revised: 
                                    queries_to_try.insert(0, address_to_geocode)
                                    
                                for q_idx, q in enumerate(queries_to_try):
                                    try:
                                        location = geolocator.geocode(q, timeout=10)
                                        if location:
                                            lat, lon = location.latitude, location.longitude
                                            used_address = q 
                                            
                                            if is_revised and q_idx == 0: status = "Sukses (Manual Translasi)"
                                            elif q_idx == 0: status = "Sukses (Spesifik OSM)"
                                            elif q_idx == 1: status = "Sukses (Tanpa Gedung OSM)"
                                            elif q_idx == 2: status = "Sukses (Kecamatan OSM)"
                                            else: status = "Sukses (Kota/Kab OSM)"
                                            break 
                                    except Exception:
                                        status = "Timeout Nominatim"
                                    time.sleep(1)
                                    
                                if not lat:
                                    status = "Alamat Tidak Ditemukan"
                                    
                            if lat and lon:
                                geo_dist = geodesic((DEPO_LAT, DEPO_LON), (lat, lon)).kilometers
                                d_out = get_osrm_distance(DEPO_LAT, DEPO_LON, lat, lon)
                                d_in  = get_osrm_distance(lat, lon, DEPO_LAT, DEPO_LON)
                                
                                if d_out > 0 and d_in > 0:
                                    if d_out > (geo_dist * 5) or d_out > 700: 
                                        status = "Peringatan: Koordinat Nyasar Jauh"
                                        dist_pp = 0.0
                                    else:
                                        dist_pp = d_out + d_in
                                else:
                                    status = "Gagal Routing OSRM"
                                    
                            new_data.append({
                                'ALAMAT_FINAL': addr, 'Jarak_PP_Km': dist_pp, 'Alasan': status, 'ALAMAT YANG DIPAKAI': used_address
                            })
                            my_bar.progress((i + 1) / len(new_addrs), text=progress_text)
                            
                        df_cache = pd.concat([df_cache, pd.DataFrame(new_data)], ignore_index=True)
                        df_cache.to_excel(cache_file, index=False)
                        
                    df_final_dooring = pd.merge(df_merged, df_cache, on='ALAMAT_FINAL', how='left')
                    
                    # 5. MENGHITUNG BERAT & FORMAT CSV AKHIR
                    def get_weight_info(size_cont):
                        size_str = str(size_cont).upper().strip()
                        W_20_EMPTY = 2.3; W_20_FULL = 27.0; W_40_EMPTY = 3.8; W_40_FULL = 32.0
                        if 'COMBO' in size_str: return (W_20_FULL * 2), (W_20_EMPTY * 2)
                        multiplier = 1
                        if '2X' in size_str or '2 X' in size_str: multiplier = 2
                        elif '1X' in size_str or '1 X' in size_str: multiplier = 1
                        if '40' in size_str: return (W_40_FULL * multiplier), (W_40_EMPTY * multiplier)
                        elif '20' in size_str: return (W_20_FULL * multiplier), (W_20_EMPTY * multiplier)
                        return W_20_FULL, W_20_EMPTY
                        
                    csv_data = []
                    for idx, row in df_final_dooring.iterrows():
                        size_cont = row.get('SIZE CONT', row.get('SIZE CONT ', ''))
                        b_full, b_empty = get_weight_info(size_cont)
                        
                        # Jika alamat tidak ada di TLP, kosongi alamat dan tulis keterangan
                        if pd.isna(row.get('ALAMAT')) or str(row.get('ALAMAT')).strip() == "":
                            jarak_pp = 0.0
                            alasan = "SOPT Tidak Ada Alamat"
                            alamat_pabrik = ""
                            alamat_dipakai = ""
                        else:
                            jarak_pp = float(row.get('Jarak_PP_Km', 0.0)) if pd.notna(row.get('Jarak_PP_Km')) else 0.0
                            alasan = row.get('Alasan', '')
                            alamat_pabrik = row.get('ALAMAT_FINAL', '')
                            alamat_dipakai = row.get('ALAMAT YANG DIPAKAI', '')
                            
                        total_berat = b_full + b_empty
                        tonkm = ((jarak_pp / 2) * b_full) + ((jarak_pp / 2) * b_empty) if jarak_pp > 0 else 0
                        
                        csv_data.append({
                            'NO': idx + 1, 'BULAN': row.get('BULAN', row.get('Bulan', '')),
                            'LAMBUNG': row.get('LAMBUNG', row.get('PLAT NUMBER', '')),
                            'NOPOL': row.get('NOPOL', row.get('LAMBUNG', '')),
                            'SIZE CONT': size_cont, 'SOPT_NO': row.get('SOPT_NO_CLEAN', ''),
                            'AREA START': 'YON', 'AREA AMBIL EMPTY': 'YON',
                            'AREA PABRIK': alamat_pabrik,
                            'ALAMAT YANG DIPAKAI': alamat_dipakai,
                            'AREA BONGKAR': 'YON', 'AREA AKHIR': 'YON',
                            'Jarak_PP_Km': jarak_pp, 'Total_Berat_Ton': total_berat,
                            'TonKm_Dooring': tonkm, 'Alasan': alasan
                        })
                        
                    df_export_dooring = pd.DataFrame(csv_data)
                    
                    out_door = io.BytesIO()
                    with pd.ExcelWriter(out_door, engine='openpyxl') as writer:
                        df_export_dooring.to_excel(writer, index=False)
                    st.session_state.out_dooring_file = out_door.getvalue()
                    
                    st.session_state.etl_step1_processed = True
                    st.success("✅ Langkah 1 Selesai!")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses Langkah 1: {e}")
        else:
            st.warning("Mohon unggah Dooring & TLP mentah untuk memulai Langkah 1.")

    if st.session_state.etl_step1_processed:
        st.download_button(
            label="📥 Download 'Jarak dan Berat Dooring.xlsx' (Hasil Langkah 1)", 
            data=st.session_state.out_dooring_file, 
            file_name="Jarak dan Berat Dooring.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")

    # ---------------------------------------------------------
    # LANGKAH 2: GABUNG SELURUH DATA (TRUCKING & NON TRUCKING)
    # ---------------------------------------------------------
    st.markdown("### 🔹 Langkah 2: Gabung BBM, Master, Haulage & Alat Berat")
    st.info("Upload file hasil perhitungan Dooring dari Langkah 1 beserta file lainnya untuk membuat file hasil Trucking dan Non-Trucking.")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        f_bbm_raw = st.file_uploader("1. Upload BBM AAB.xlsx", type=["xlsx", "csv"], key="bbm_step2")
        f_dooring_raw = st.file_uploader("4. Upload Jarak & Berat Dooring.xlsx", type=["xlsx", "csv"], key="drg_step2")
    with col2:
        f_master_raw = st.file_uploader("2. Upload cost & bbm 2022 sd 2025 HP & Type.xlsx", type=["xlsx", "csv"], key="mst_step2")
        f_nontruck_raw = st.file_uploader("5. Upload Summary Alat Berat.xlsx", type=["xlsx", "csv"], key="nt_step2")
    with col3:
        f_haulage_raw = st.file_uploader("3. Upload HAULAGE OKT-DES 2025 (Copy).xlsx", type=["xlsx", "csv"], key="hlg_step2")
        
    if st.button("🔄 2. Buat File Hasil Trucking dan Non-Trucking"):
        if f_bbm_raw and f_master_raw and f_haulage_raw and f_dooring_raw and f_nontruck_raw:
            with st.spinner("Memproses seluruh data operasional (Mencocokkan Logika Asli)..."):
                try:
                    # ==============================================================================
                    # ANALISA TRUCKING KODE ASLI
                    # ==============================================================================
                    def extract_lambung_code_tr(raw_kpi):
                        s = str(raw_kpi).upper().strip()
                        match = re.search(r'\((.*?)\)', s)
                        if match: return match.group(1).strip()
                        parts = s.split()
                        if len(parts) > 1:
                            last_part = parts[-1]
                            if len(last_part) < 6: return last_part
                        return s.replace("-", "").strip()

                    master_data_map_tr = {}
                    master_keys_set_tr = set()
                    
                    f_master_raw.seek(0)
                    df_map_tr = pd.read_excel(f_master_raw, sheet_name='Sheet2', header=1)
                    col_name_tr = next((c for c in df_map_tr.columns if 'NAMA' in str(c).upper()), None)
                    df_map_tr.dropna(subset=[col_name_tr], inplace=True)
                    
                    for _, row in df_map_tr.iterrows():
                        u_name = str(row[col_name_tr]).strip().upper()
                        c_id = clean_unit_name(u_name)
                        if c_id:
                            master_data_map_tr[c_id] = u_name
                            master_keys_set_tr.add(c_id)

                    def get_master_match_tr(raw_name):
                        raw_name = str(raw_name).strip().upper()
                        c_raw = clean_unit_name(raw_name)
                        if c_raw in master_data_map_tr: return master_data_map_tr[c_raw]
                        if " (" in raw_name:
                            b_paren = clean_unit_name(raw_name.split(" (")[0])
                            if b_paren in master_data_map_tr: return master_data_map_tr[b_paren]
                        if "EX." in raw_name or "EX " in raw_name:
                            after_ex = raw_name.split("EX.")[-1] if "EX." in raw_name else raw_name.split("EX ")[-1]
                            c_after = clean_unit_name(after_ex.replace(")", ""))
                            if c_after in master_data_map_tr: return master_data_map_tr[c_after]
                        return raw_name 

                    lambung_to_std_name = {} 
                    all_bbm_data = []

                    f_bbm_raw.seek(0)
                    xls_bbm = pd.ExcelFile(f_bbm_raw)
                    for sheet_name in xls_bbm.sheet_names:
                        sht_up = sheet_name.upper()
                        if not any(x in sht_up for x in ['OKT', 'OCT', 'NOV']): continue
                            
                        df_cek = pd.read_excel(xls_bbm, sheet_name=sheet_name, header=None, nrows=3)
                        if df_cek.empty: continue
                        a1_val = str(df_cek.iloc[0, 0]).strip().upper()
                        if "EQUIP" not in a1_val and "NAMA" not in a1_val: continue 
                            
                        df_full = pd.read_excel(xls_bbm, sheet_name=sheet_name, header=None)
                        unit_names_row = df_full.iloc[0].ffill() 
                        group_kpi_row = df_full.iloc[1].ffill()  
                        headers = df_full.iloc[2]                
                        data_rows = df_full.iloc[3:]             
                        
                        def is_valid_date_row(val):
                            if pd.isna(val): return False
                            if isinstance(val, (int, float)) and 1 <= val <= 31: return True
                            val_str = str(val).strip()
                            if re.match(r'^\d{4}-\d{2}-\d{2}', val_str): return True 
                            match = re.match(r'^(\d{1,2})', val_str) 
                            if match and 1 <= int(match.group(1)) <= 31: return True
                            return False
                            
                        valid_mask = data_rows.iloc[:, 0].apply(is_valid_date_row)
                        valid_data_rows = data_rows[valid_mask]
                        bln_str = 'Oktober' if 'OKT' in sht_up or 'OCT' in sht_up else 'November'
                        
                        for col in range(1, df_full.shape[1]):
                            header_str = str(headers.iloc[col]).strip().upper()
                            if 'LITER' in header_str:
                                raw_equip = str(unit_names_row.iloc[col]).strip().upper()
                                raw_kpi = str(group_kpi_row.iloc[col]).strip().upper()
                                if "TOTAL" in raw_equip or "UNNAMED" in raw_equip or raw_equip == "NAN": continue
                                
                                is_trucking = 0
                                if any(k in raw_kpi for k in ['TRAILER', 'TRONTON']) or any(k in raw_equip for k in ['TRAILER', 'TRONTON']): is_trucking = 1
                                if any(k in raw_equip for k in ['MOBIL', 'STOORING', 'MEKANIK']) or any(k in raw_kpi for k in ['MOBIL', 'STOORING', 'MEKANIK']): is_trucking = 0
                                
                                lambung = extract_lambung_code_tr(raw_kpi)
                                lambung_clean = clean_unit_name(lambung)
                                std_name = get_master_match_tr(raw_equip)
                                
                                if lambung_clean: lambung_to_std_name[lambung_clean] = std_name
                                lambung_to_std_name[clean_unit_name(raw_equip)] = std_name 
                                
                                vals = pd.to_numeric(valid_data_rows.iloc[:, col], errors='coerce').sum()
                                if vals > 0:
                                    all_bbm_data.append({'Nama_Unit': std_name, 'Unit_Clean': clean_unit_name(std_name), 'LITER': vals, 'Bulan': bln_str, 'Is_Trucking': is_trucking})

                    if all_bbm_data:
                        df_bbm_total = pd.DataFrame(all_bbm_data).groupby(['Nama_Unit', 'Unit_Clean']).agg({'LITER': 'sum', 'Is_Trucking': 'max'}).reset_index()
                        df_bbm_monthly = pd.DataFrame(all_bbm_data).groupby(['Unit_Clean', 'Bulan'])['LITER'].sum().reset_index()
                    else:
                        df_bbm_total = pd.DataFrame(columns=['Nama_Unit', 'Unit_Clean', 'LITER', 'Is_Trucking'])
                        df_bbm_monthly = pd.DataFrame(columns=['Unit_Clean', 'Bulan', 'LITER'])

                    def get_haulage_distance(rute, depo, jenis_cont, exim):
                        rute_up, depo_up, cont_up, exim_up = str(rute).upper().strip(), str(depo).upper().strip(), str(jenis_cont).upper().strip(), str(exim).upper().strip()
                        if "BERLIAN - DEPO" in rute_up and depo_up == "UDATIN": return 0.500
                        if "BERLIAN - DEPO" in rute_up and depo_up == "MIRAH": return 1.100
                        if "BERLIAN - DEPO" in rute_up and depo_up == "UNGGUL": return 0.750
                        if "BERLIAN - DEPO" in rute_up and depo_up == "I POWER": return 0.900
                        if "BERLIAN - DEPO" in rute_up and depo_up == "MARUMAN": return 0.270
                        if "DEPO - BERLIAN" in rute_up and depo_up == "UDATIN": return 0.850
                        if "DEPO - BERLIAN" in rute_up and depo_up == "MIRAH": return 2.800
                        if "DEPO - BERLIAN" in rute_up and depo_up == "UNGGUL": return 2.200
                        if "DEPO - BERLIAN" in rute_up and depo_up == "I POWER": return 0.650
                        if "BERLIAN - KALIANAK" in rute_up and exim_up == "EXIM" and ("20FT" in cont_up or "40FT" in cont_up): return 11.0
                        if "DEPO - KALIANAK" in rute_up and "GENSET" in cont_up: return 5.150
                        if "DEPO - KALIANAK" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 9.350
                        if "KALIANAK - BERLIAN" in rute_up and "GENSET" in cont_up: return 6.400
                        if "KALIANAK - DEPO" in rute_up and "GENSET" in cont_up: return 4.975
                        if "KALIANAK - DEPO" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 8.975
                        if "KALIANAK - TELUK LAMONG" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 9.200
                        if "KALIANAK-TAMBAK LANGON" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 2.700
                        if "NILAM - KALIANAK" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 10.300
                        if "TAMBAK LANGON - KALIANAK" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 3.300
                        if "TELUK LAMONG - KALIANAK" in rute_up and ("20FT" in cont_up or "40FT" in cont_up): return 8.600
                        if "BERLIAN - DEPO" in rute_up and exim_up == "EXIM": return 2.775
                        if "DEPO - BERLIAN" in rute_up and exim_up == "EXIM": return 2.575
                        distances = {
                            "BERLIAN - DEPO": 2.775, "BERLIAN - TAMBAK LANGON": 11.5, "BERLIAN - TERMINAL TELUK LAMONG": 18.1,
                            "DEPO - BERLIAN": 2.575, "DEPO - DEPO": 2.825, "DEPO - ICT": 7.65, "DEPO - KALOG": 2.275, "DEPO - NILAM": 2.375,
                            "DEPO PERAK - TAMBAK LANGON": 10.025, "DEPO PERAK - TELUK LAMONG": 16.6, "ICT - DEPO": 4.55,
                            "ICT/TPS - TAMBAK LANGON": 13.2, "ICT/TPS - TELUK LAMONG": 19.8, "JAMRUD - DEPO": 5.6, "KADE - KADE": 2.0,
                            "KALOG - BERLIAN": 3.4, "KALOG - DEPO": 3.575, "KALOG - NILAM": 3.7, "KALOG - TAMBAK LANGON": 9.4, "KALOG - TELUK LAMONG": 16.0,
                            "MIRAH - DEPO": 4.975, "NILAM - DEPO": 2.325, "NILAM - TAMBAK LANGON": 10.8, "NILAM - TERMINAL TELUK LAMONG": 17.4,
                            "OPER SHIFT BERLIAN-DEPO": 2.775, "OPER SHIFT DEPO - BERLIAN": 2.575, "OPER SHIFT JAMRUD - DEPO": 5.6,
                            "OPER SHIFT NILAM-DEPO": 2.325, "TAMBAK LANGON - BERLIAN": 10.2, "TAMBAK LANGON - DEPO PERAK": 8.875,
                            "TAMBAK LANGON - NILAM": 10.5, "TAMBAK LANGON - TTL": 6.7, "TELUK LAMONG - DEPO PERAK": 14.325,
                            "TELUK LAMONG - ICT/TPS": 20.6, "TELUK LAMONG - KALOG": 14.3, "TERMINAL TELUK LAMONG - BERLIAN": 15.5,
                            "TERMINAL TELUK LAMONG - NILAM": 15.9, "TTL - TAMBAK LANGON": 6.4
                        }
                        for key, val in distances.items():
                            if key in rute_up: return val
                        return 0.0

                    def get_haulage_weight(kegiatan, jenis_cont, is_full):
                        kegiatan, jenis_cont, is_full = str(kegiatan).upper().strip(), str(jenis_cont).upper().strip(), str(is_full).upper().strip()
                        if "MOVE ARMADA" in kegiatan or "KIR" in kegiatan or "LOWBED" in jenis_cont or "DOLLY" in jenis_cont: return 0.0
                        if "GENSET" in jenis_cont: return 2.3 
                        if "20FT" in jenis_cont: return 27.0 if "FULL" in is_full else 2.3
                        if "40FT" in jenis_cont: return 32.0 if "FULL" in is_full else 3.8
                        return 0.0

                    unmapped_haulage = []
                    f_haulage_raw.seek(0)
                    try:
                        df_haulage = pd.read_excel(f_haulage_raw)
                        df_haulage = df_haulage.rename(columns={'JENIS\nCONT': 'JENIS CONT'})
                        col_lambung = 'NO LAMBUNG' if 'NO LAMBUNG' in df_haulage.columns else 'LAMBUNG'
                        
                        if 'TANGGAL' in df_haulage.columns: df_haulage['Bulan'] = df_haulage['TANGGAL'].dt.month.map({10: 'Oktober', 11: 'November', 12: 'Desember'})
                        else: df_haulage['Bulan'] = 'Oktober'
                        df_haulage = df_haulage[df_haulage['Bulan'].isin(['Oktober', 'November'])]
                        
                        df_haulage['LAMBUNG_CLEAN'] = df_haulage[col_lambung].astype(str).apply(lambda x: clean_unit_name(extract_lambung_code_tr(x)))
                        df_haulage['Nama_Unit'] = df_haulage['LAMBUNG_CLEAN'].map(lambung_to_std_name).fillna(df_haulage['LAMBUNG_CLEAN'])
                        df_haulage['Unit_Clean'] = df_haulage['Nama_Unit'].apply(clean_unit_name)
                        
                        df_haulage['Jarak_Km'] = df_haulage.apply(lambda row: get_haulage_distance(row.get('RUTE', ''), row.get('DEPO', ''), row.get('JENIS CONT', ''), row.get('EXIM', '')), axis=1)
                        df_haulage['Berat_Ton'] = df_haulage.apply(lambda row: get_haulage_weight(row.get('Jenis Kegiatan', ''), row.get('JENIS CONT', ''), row.get('FULL/EMPTY', '')), axis=1)
                        df_haulage['TonKm'] = df_haulage['Jarak_Km'] * df_haulage['Berat_Ton'] 
                        
                        unmapped_haulage = df_haulage[~df_haulage['LAMBUNG_CLEAN'].isin(lambung_to_std_name.keys())]['LAMBUNG_CLEAN'].unique().tolist()
                        df_haul_agregat = df_haulage.groupby('Unit_Clean').agg({'Berat_Ton': 'sum', 'TonKm': 'sum', 'Jarak_Km': 'sum'}).reset_index().rename(columns={'Jarak_Km': 'Km_Haulage', 'Berat_Ton': 'Ton_Haulage', 'TonKm': 'TonKm_Haulage'})
                        df_haul_monthly = df_haulage.groupby(['Unit_Clean', 'Bulan']).agg({'Berat_Ton': 'sum', 'TonKm': 'sum', 'Jarak_Km': 'sum'}).reset_index().rename(columns={'Jarak_Km': 'Km_Haulage', 'Berat_Ton': 'Ton_Haulage', 'TonKm': 'TonKm_Haulage'})
                    except Exception as e:
                        df_haul_agregat = pd.DataFrame(columns=['Unit_Clean', 'Ton_Haulage', 'TonKm_Haulage', 'Km_Haulage'])
                        df_haul_monthly = pd.DataFrame(columns=['Unit_Clean', 'Bulan', 'Ton_Haulage', 'TonKm_Haulage', 'Km_Haulage'])

                    unmapped_dooring = []
                    f_dooring_raw.seek(0)
                    try:
                        df_dooring = pd.read_excel(f_dooring_raw)
                        if 'BULAN' in df_dooring.columns:
                            df_dooring['Bulan'] = df_dooring['BULAN'].astype(str).str.upper().map({'OKTOBER': 'Oktober', 'NOVEMBER': 'November', 'DESEMBER': 'Desember'}).fillna(df_dooring['BULAN'])
                        else:
                            df_dooring['Bulan'] = 'Oktober'
                        df_dooring = df_dooring[df_dooring['Bulan'].isin(['Oktober', 'November'])]
                        
                        col_lamb_door = 'LAMBUNG' if 'LAMBUNG' in df_dooring.columns else ('NO LAMBUNG' if 'NO LAMBUNG' in df_dooring.columns else df_dooring.columns[0])
                        df_dooring['LAMBUNG_CLEAN'] = df_dooring[col_lamb_door].astype(str).apply(lambda x: clean_unit_name(extract_lambung_code_tr(x)))
                        df_dooring['Nama_Unit'] = df_dooring['LAMBUNG_CLEAN'].map(lambung_to_std_name).fillna(df_dooring['LAMBUNG_CLEAN'])
                        df_dooring['Unit_Clean'] = df_dooring['Nama_Unit'].apply(clean_unit_name)
                        
                        unmapped_dooring = df_dooring[~df_dooring['LAMBUNG_CLEAN'].isin(lambung_to_std_name.keys())]['LAMBUNG_CLEAN'].unique().tolist()
                        
                        col_berat_door = 'Total_Berat_Ton' if 'Total_Berat_Ton' in df_dooring.columns else 'Berat_Ton'
                        col_tonkm_door = 'TonKm_Dooring' if 'TonKm_Dooring' in df_dooring.columns else 'TonKm'
                        col_jarak_door = 'Jarak_PP_Km' if 'Jarak_PP_Km' in df_dooring.columns else 'Jarak_Km'
                        
                        df_door_agregat = df_dooring.groupby('Unit_Clean').agg({col_berat_door: 'sum', col_tonkm_door: 'sum', col_jarak_door: 'sum'}).reset_index().rename(columns={col_jarak_door: 'Km_Dooring', col_berat_door: 'Ton_Dooring', col_tonkm_door: 'TonKm_Dooring'})
                        df_door_monthly = df_dooring.groupby(['Unit_Clean', 'Bulan']).agg({col_berat_door: 'sum', col_tonkm_door: 'sum', col_jarak_door: 'sum'}).reset_index().rename(columns={col_jarak_door: 'Km_Dooring', col_berat_door: 'Ton_Dooring', col_tonkm_door: 'TonKm_Dooring'})
                    except Exception as e:
                        df_door_agregat = pd.DataFrame(columns=['Unit_Clean', 'Km_Dooring', 'Ton_Dooring', 'TonKm_Dooring'])
                        df_door_monthly = pd.DataFrame(columns=['Unit_Clean', 'Bulan', 'Km_Dooring', 'Ton_Dooring', 'TonKm_Dooring'])

                    df_ops = pd.merge(df_haul_agregat, df_door_agregat, on='Unit_Clean', how='outer').fillna(0)
                    df_ops['Total_Km'] = df_ops.get('Km_Haulage', 0) + df_ops.get('Km_Dooring', 0)
                    df_ops['Total_TonKm'] = df_ops.get('TonKm_Haulage', 0) + df_ops.get('TonKm_Dooring', 0)
                    df_ops['Total_Ton'] = df_ops.get('Ton_Haulage', 0) + df_ops.get('Ton_Dooring', 0)

                    df_audit = pd.merge(df_ops, df_bbm_total, on='Unit_Clean', how='outer', indicator=True)
                    for c in ['Km_Haulage', 'Ton_Haulage', 'TonKm_Haulage', 'Km_Dooring', 'Ton_Dooring', 'TonKm_Dooring', 'Total_Km', 'Total_Ton', 'Total_TonKm', 'LITER']:
                        if c in df_audit.columns: df_audit[c] = df_audit[c].fillna(0)

                    if 'Nama_Unit' in df_bbm_total.columns:
                        clean_to_std = dict(zip(df_bbm_total['Unit_Clean'], df_bbm_total['Nama_Unit']))
                        df_audit['Nama_Unit'] = df_audit['Unit_Clean'].map(clean_to_std).fillna(df_audit['Unit_Clean'])
                    elif 'Unit_Clean' in df_audit.columns:
                        df_audit['Nama_Unit'] = df_audit['Unit_Clean']
                    else:
                        df_audit['Nama_Unit'] = "UNKNOWN"

                    df_sukses = df_audit[df_audit['_merge'] == 'both'].copy()
                    df_sukses['L_per_TonKm'] = np.where(df_sukses['Total_TonKm'] > 0, df_sukses['LITER'] / df_sukses['Total_TonKm'], 0)
                    df_sukses['Km_per_L'] = np.where(df_sukses['LITER'] > 0, df_sukses['Total_Km'] / df_sukses['LITER'], 0)

                    df_no_bbm = df_audit[df_audit['_merge'] == 'left_only'].copy()
                    df_no_bbm['Keterangan'] = "Ada Kegiatan Dooring/Haulage, Tapi Data BBM Tidak Ditemukan"

                    df_idle = df_audit[df_audit['_merge'] == 'right_only'].copy()
                    if 'Is_Trucking' in df_idle.columns:
                        df_idle = df_idle[df_idle['Is_Trucking'] == 1]
                    df_idle['Keterangan'] = "Ada Pengisian BBM, Tapi Tidak Ada Kegiatan Dooring/Haulage (Truk Sedang Idle/Standby)"

                    df_monthly_merged = pd.merge(df_bbm_monthly, df_door_monthly, on=['Unit_Clean', 'Bulan'], how='outer').fillna(0)
                    df_monthly_merged = pd.merge(df_monthly_merged, df_haul_monthly, on=['Unit_Clean', 'Bulan'], how='outer').fillna(0)

                    if 'Nama_Unit' in df_bbm_total.columns:
                        df_monthly_merged['Nama_Unit'] = df_monthly_merged['Unit_Clean'].map(clean_to_std).fillna(df_monthly_merged['Unit_Clean'])
                    elif 'Unit_Clean' in df_monthly_merged.columns:
                        df_monthly_merged['Nama_Unit'] = df_monthly_merged['Unit_Clean']
                    else:
                        df_monthly_merged['Nama_Unit'] = "UNKNOWN"

                    df_monthly_merged['Total_Ton'] = df_monthly_merged.get('Ton_Haulage', 0) + df_monthly_merged.get('Ton_Dooring', 0)
                    df_monthly_merged['Total_TonKm'] = df_monthly_merged.get('TonKm_Haulage', 0) + df_monthly_merged.get('TonKm_Dooring', 0)
                    df_monthly_merged['Total_Km'] = df_monthly_merged.get('Km_Haulage', 0) + df_monthly_merged.get('Km_Dooring', 0)
                    df_monthly_merged['Fuel_Ratio'] = np.where(df_monthly_merged['Total_TonKm'] > 0, df_monthly_merged['LITER'] / df_monthly_merged['Total_TonKm'], 0)

                    valid_units = set(df_sukses['Nama_Unit']) if not df_sukses.empty else set()
                    df_monthly_out = df_monthly_merged[df_monthly_merged['Nama_Unit'].isin(valid_units)]
                    
                    if not df_monthly_out.empty:
                        df_monthly_out = df_monthly_out[['Nama_Unit', 'Bulan', 'LITER', 'Total_Ton', 'Total_TonKm', 'Fuel_Ratio', 'Total_Km']].sort_values(['Nama_Unit', 'Bulan'])
                        df_pivot_tr = df_monthly_out.pivot_table(index='Nama_Unit', columns='Bulan', values='Fuel_Ratio').reset_index().fillna(0)
                        if 'Oktober' in df_pivot_tr.columns and 'November' in df_pivot_tr.columns:
                            df_pivot_tr['Status_Tren'] = df_pivot_tr.apply(lambda x: "MEMBAIK (Makin Irit)" if 0 < x['November'] < x['Oktober'] else ("MEMBURUK (Makin Boros)" if x['November'] > x['Oktober'] > 0 else "STABIL"), axis=1)
                    else:
                        df_pivot_tr = pd.DataFrame()

                    df_unmapped = pd.DataFrame({'Kode_Lambung': list(set(unmapped_dooring + unmapped_haulage)), 'Keterangan': "Lambung dari Dooring/Haulage ini TIDAK ditemukan di GROUP KPI file BBM AAB."})
                    
                    out_truck = io.BytesIO()
                    with pd.ExcelWriter(out_truck, engine='openpyxl') as writer:
                        if not df_sukses.empty: df_sukses[['Nama_Unit','Total_Km','Total_Ton','Total_TonKm','LITER','L_per_TonKm','Km_per_L']].to_excel(writer, sheet_name='HASIL_ANALISA', index=False)
                        else: pd.DataFrame(columns=['Nama_Unit','Total_Km','Total_Ton','Total_TonKm','LITER','L_per_TonKm','Km_per_L']).to_excel(writer, sheet_name='HASIL_ANALISA', index=False)
                        if not df_monthly_out.empty: df_monthly_out.to_excel(writer, sheet_name='Data_Bulanan', index=False)
                        else: pd.DataFrame().to_excel(writer, sheet_name='Data_Bulanan', index=False)
                        if not df_pivot_tr.empty: df_pivot_tr.to_excel(writer, sheet_name='Laporan_Tren', index=False)
                        if not df_no_bbm.empty: df_no_bbm[['Nama_Unit','Total_Km','Total_Ton','Total_TonKm','Keterangan']].to_excel(writer, sheet_name='OPS_TANPA_BBM', index=False)
                        if not df_idle.empty: df_idle[['Nama_Unit','LITER','Keterangan']].to_excel(writer, sheet_name='BBM_TANPA_OPS', index=False)
                        if not df_unmapped.empty: df_unmapped.to_excel(writer, sheet_name='GAGAL_MAPPING', index=False)
                    st.session_state.out_truck_file = out_truck.getvalue()

                    # ==============================================================================
                    # 2. ANALISA NON-TRUCKING KODE ASLI
                    # ==============================================================================
                    MAP_CABANG = {
                        'TANGKIANG': 'CABANG LUWUK', 'AMBON': 'CABANG AMBON', 'BAU BAU': 'BAU-BAU',
                        'MANOKWARI': 'DEPO MANOKWARI', 'MERAK': 'DEPO MERAK', 'MARUNI': 'DEPO MARUNI'
                    }
                    STATUS_WEIGHT_TYPE = {
                        'MTA': 'empty', 'STR': 'full', 'MTB': 'full', 'STF': 'full', 'FAC': 'full', 'MAS': 'empty', 
                        'FOB': 'full', 'MOB': 'empty', 'FXD': 'full', 'MXD': 'empty', 'FTL': 'full', 'MTL': 'empty', 'FIT': 'full', 'MIT': 'empty'
                    }
                    WEIGHT_DICT = {
                        ('20 Feet', 'empty'): 2300, ('20 Feet', 'full'): 27000,
                        ('40 Feet', 'empty'): 3800, ('40 Feet', 'full'): 32000
                    }
                    EXCLUDED_BRANCHES = ['JAKARTA', 'SURABAYA']

                    master_data_map_nt = {} 
                    master_keys_set_nt = set()
                    
                    f_master_raw.seek(0)
                    df_map_nt = pd.read_excel(f_master_raw, sheet_name='Sheet2', header=1)
                    col_name_nt = next((c for c in df_map_nt.columns if 'NAMA' in str(c).upper()), None)
                    col_jenis_nt = next((c for c in df_map_nt.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name_nt), None)
                    col_type_nt = next((c for c in df_map_nt.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
                    col_hp_nt = next((c for c in df_map_nt.columns if any(k == str(c).upper() for k in ['HP', 'HORSE POWER'])), None)
                    col_cap_nt = next((c for c in df_map_nt.columns if any(k in str(c).upper() for k in ['CAP', 'KAPASITAS'])), None)
                    col_loc_nt = next((c for c in df_map_nt.columns if 'LOKASI' in str(c).upper() or 'DES 2025' in str(c).upper()), df_map_nt.columns[2])

                    df_map_nt.dropna(subset=[col_name_nt], inplace=True)
                    df_map_nt['Unit_Original'] = df_map_nt[col_name_nt].astype(str).str.strip().str.upper()
                    df_map_nt = df_map_nt[~df_map_nt['Unit_Original'].str.contains('DUMMY', na=False)]
                    df_map_nt = df_map_nt[~df_map_nt['Unit_Original'].str.contains('FALCON', na=False)]

                    for _, row in df_map_nt.iterrows():
                        u_name = str(row['Unit_Original'])
                        c_id = clean_unit_name(u_name)
                        if c_id:
                            cap_val = 0
                            try:
                                match = re.search(r"(\d+(\.\d+)?)", str(row[col_cap_nt]))
                                if match: cap_val = int(float(match.group(1)) + 0.5)
                            except: pass

                            t_merk = str(row[col_type_nt]).strip().upper() if pd.notna(row[col_type_nt]) else "-"
                            jenis = str(row[col_jenis_nt]).strip().upper() if pd.notna(row[col_jenis_nt]) else "-"
                            is_trucking = True if any(x in jenis for x in ['TRAILER', 'TRONTON', 'TRUCK']) else False

                            master_data_map_nt[c_id] = {
                                'Unit_Name': u_name, 'Jenis_Alat': jenis, 'Type_Merk': t_merk, 
                                'Horse_Power': row[col_hp_nt] if pd.notna(row[col_hp_nt]) else "0", 
                                'Capacity': cap_val, 'Lokasi': str(row[col_loc_nt]).strip().upper(),
                                'Is_Trucking': is_trucking
                            }
                            master_keys_set_nt.add(c_id)

                    sheets_bbm = {'JAN': 'Januari', 'FEB': 'Februari', 'MAR': 'Maret', 'APR': 'April', 'MEI': 'Mei', 'JUN': 'Juni', 'JUL': 'Juli', 'AGT': 'Agustus', 'SEP': 'September', 'OKT': 'Oktober', 'NOV': 'November'}
                    list_df_bbm_nt = []
                    
                    f_bbm_raw.seek(0)
                    xls_bbm_nt = pd.ExcelFile(f_bbm_raw)

                    def get_master_match_nt(raw_name):
                        raw_name = str(raw_name).strip().upper()
                        if "WIND RIVER" in raw_name: return clean_unit_name("TOP LOADER BOSS") if clean_unit_name("TOP LOADER BOSS") in master_data_map_nt else None
                        if "FL RENTAL 01" in raw_name and "TIMIKA" not in raw_name: return clean_unit_name("FL RENTAL 01 TIMIKA") if clean_unit_name("FL RENTAL 01 TIMIKA") in master_data_map_nt else None
                        if "TOBATI" in raw_name: return clean_unit_name("TOP LOADER KALMAR 35 T/TOBATI") if clean_unit_name("TOP LOADER KALMAR 35 T/TOBATI") in master_data_map_nt else None
                        if "L 8477 UUC" in raw_name: return clean_unit_name("L 9902 UR / S75") if clean_unit_name("L 9902 UR / S75") in master_data_map_nt else None
                            
                        c_raw = clean_unit_name(raw_name)
                        if c_raw in master_data_map_nt: return c_raw
                        if " (" in raw_name:
                            b_paren = clean_unit_name(raw_name.split(" (")[0])
                            if b_paren in master_data_map_nt: return b_paren
                        if "EX." in raw_name or "EX " in raw_name:
                            after_ex = raw_name.split("EX.")[-1] if "EX." in raw_name else raw_name.split("EX ")[-1]
                            c_after = clean_unit_name(after_ex.replace(")", ""))
                            if c_after in master_data_map_nt: return c_after
                            for m_key in master_keys_set_nt:
                                if c_after != "" and c_after in m_key: return m_key
                        for p in raw_name.replace("(", "/").replace(")", "/").split("/"):
                            cl = clean_unit_name(p)
                            if cl in master_data_map_nt: return cl
                        return None

                    for sht, bln in sheets_bbm.items():
                        if sht in xls_bbm_nt.sheet_names:
                            df_temp = pd.read_excel(xls_bbm_nt, sheet_name=sht, header=None)
                            unit_names_row = df_temp.iloc[0].ffill()
                            headers = df_temp.iloc[2]
                            data = df_temp.iloc[3:]
                            
                            for col in range(1, df_temp.shape[1]):
                                header_str = str(headers[col]).strip().upper()
                                if header_str in ['LITER', 'QTY']:
                                    raw_unit_name = str(unit_names_row[col]).strip().upper()
                                    if raw_unit_name == "" or "UNNAMED" in raw_unit_name or "TOTAL" in raw_unit_name: continue
                                    if raw_unit_name.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL', 'GROUP')): continue
                                    
                                    matched_id = get_master_match_nt(raw_unit_name)
                                    if matched_id:
                                        vals = pd.to_numeric(data[col], errors='coerce').sum()
                                        if vals > 0:
                                            list_df_bbm_nt.append({'Unit_Name': master_data_map_nt[matched_id]['Unit_Name'], 'Bulan': bln, 'LITER': vals})

                    df_bbm_all_nt = pd.DataFrame(list_df_bbm_nt).groupby(['Unit_Name', 'Bulan'])['LITER'].sum().reset_index() if list_df_bbm_nt else pd.DataFrame(columns=['Unit_Name', 'Bulan', 'LITER'])

                    def get_allowed_tasks(jenis, cap):
                        jenis = str(jenis).upper()
                        if 'FORKLIFT' in jenis:
                            if cap == 0 or (3 <= cap <= 8): return ['STR', 'STF', 'MTB']
                            elif cap >= 10: return ['STR', 'STF', 'MTB', 'FAC', 'MAS'] 
                        elif 'REACH STACKER' in jenis: return ['FOB', 'MOB', 'FXD', 'MXD', 'FAC', 'MAS']
                        elif 'TOP LOADER' in jenis or 'SIDE LOADER' in jenis: return ['FOB', 'MOB', 'FXD', 'MXD', 'FAC', 'MAS']
                        elif 'CRANE' in jenis:
                            if cap >= 70: return ['FOB', 'MOB', 'FXD', 'MXD']
                            elif cap >= 40: return ['FOB', 'MOB', 'FXD', 'MXD', 'FAC', 'MAS']
                        return []

                    f_nontruck_raw.seek(0)
                    xls_summary = pd.ExcelFile(f_nontruck_raw)
                    all_cabang_data = []

                    for sheet in xls_summary.sheet_names:
                        cabang_summary = str(sheet).upper().strip()
                        if cabang_summary in EXCLUDED_BRANCHES: continue 
                            
                        lokasi_bbm = MAP_CABANG.get(cabang_summary, cabang_summary)
                        unit_cabang_dict = {k:v for k,v in master_data_map_nt.items() if not v['Is_Trucking'] and clean_unit_name(lokasi_bbm) in clean_unit_name(str(v['Lokasi']))}
                        if not unit_cabang_dict: continue 
                            
                        try:
                            df_sheet = pd.read_excel(xls_summary, sheet_name=sheet, header=2).dropna(subset=['Bulan'])
                            df_sheet = df_sheet[df_sheet['Bulan'] != 'Grand Total']
                            raw_df = pd.read_excel(xls_summary, sheet_name=sheet, header=None)
                            clean_status = [str(x).split(' ')[0].upper() for x in raw_df.iloc[2].fillna(method='ffill').values]
                            header_size = raw_df.iloc[3].values
                            
                            for idx, row in df_sheet.iterrows():
                                bulan = str(row['Bulan']).strip().capitalize()
                                task_counts = {}
                                for col_idx in range(1, len(row)):
                                    if col_idx >= len(clean_status): break
                                    stat, size, val = clean_status[col_idx], str(header_size[col_idx]).strip(), row.iloc[col_idx]
                                    if pd.notna(val) and val != '' and stat != 'TOTAL':
                                        task_counts[(stat, size)] = task_counts.get((stat, size), 0) + float(val)
                                        
                                for u_id, unit in unit_cabang_dict.items():
                                    allowed = get_allowed_tasks(unit['Jenis_Alat'], unit['Capacity'])
                                    total_berat = 0
                                    for (stat, size), count in task_counts.items():
                                        if stat in allowed:
                                            rekan = sum(1 for k, u_lain in unit_cabang_dict.items() if stat in get_allowed_tasks(u_lain['Jenis_Alat'], u_lain['Capacity']))
                                            if rekan > 0: total_berat += (count / rekan) * WEIGHT_DICT.get((size, STATUS_WEIGHT_TYPE.get(stat, 'empty')), 0)
                                    
                                    if total_berat > 0:
                                        all_cabang_data.append({'Unit_Name': unit['Unit_Name'], 'Bulan': bulan, 'Total_Ton': total_berat / 1000})
                        except: pass

                    df_tonase = pd.DataFrame(all_cabang_data).groupby(['Unit_Name', 'Bulan'])['Total_Ton'].sum().reset_index() if all_cabang_data else pd.DataFrame(columns=['Unit_Name', 'Bulan', 'Total_Ton'])
                    df_bbm_ab = df_bbm_all_nt[df_bbm_all_nt['Unit_Name'].isin([v['Unit_Name'] for v in master_data_map_nt.values() if not v['Is_Trucking']])] if not df_bbm_all_nt.empty else pd.DataFrame(columns=['Unit_Name', 'Bulan', 'LITER'])

                    df_monthly_nt = pd.merge(df_tonase, df_bbm_ab, on=['Unit_Name', 'Bulan'], how='outer').fillna(0)
                    for k in ['Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity', 'Lokasi']:
                        df_monthly_nt[k] = df_monthly_nt['Unit_Name'].apply(lambda n: master_data_map_nt[clean_unit_name(n)][k] if clean_unit_name(n) in master_data_map_nt else "-")

                    df_monthly_nt['Fuel_Ratio'] = np.where(df_monthly_nt['Total_Ton'] > 0, df_monthly_nt['LITER'] / df_monthly_nt['Total_Ton'], 0)
                    df_total_nt = df_monthly_nt.groupby(['Unit_Name', 'Lokasi', 'Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity']).agg({'Total_Ton': 'sum', 'LITER': 'sum'}).reset_index()

                    active_names = df_total_nt['Unit_Name'].unique() if not df_total_nt.empty else []
                    inaktif_list = []
                    for v in master_data_map_nt.values():
                        if not v['Is_Trucking'] and v['Unit_Name'] not in active_names:
                            inaktif_list.append({'Nama Unit': v['Unit_Name'], 'Jenis': v['Jenis_Alat'], 'Type/Merk': v['Type_Merk'], 'Lokasi': v['Lokasi'], 'Total Pengisian BBM': 0, 'Total Berat Angkutan (Ton)': 0, 'Keterangan': 'Tidak ada aktivitas (BBM 0 & Tonase 0)'})

                    out_nontruck = io.BytesIO()
                    with pd.ExcelWriter(out_nontruck, engine='openpyxl') as writer:
                        if not df_total_nt.empty:
                            df_total_nt.to_excel(writer, sheet_name='Total_Agregat', index=False)
                        else:
                            pd.DataFrame(columns=['Unit_Name','Lokasi','Jenis_Alat','Type_Merk','Horse_Power','Capacity','Total_Ton','LITER']).to_excel(writer, sheet_name='Total_Agregat', index=False)
                            
                        if not df_monthly_nt.empty:
                            df_monthly_nt.to_excel(writer, sheet_name='Data_Bulanan', index=False)
                        else:
                            pd.DataFrame(columns=['Unit_Name','Bulan','Total_Ton','LITER','Jenis_Alat','Type_Merk','Horse_Power','Capacity','Lokasi','Fuel_Ratio']).to_excel(writer, sheet_name='Data_Bulanan', index=False)
                            
                        if inaktif_list: pd.DataFrame(inaktif_list).to_excel(writer, sheet_name='Unit_Inaktif', index=False)
                    st.session_state.out_nontruck_file = out_nontruck.getvalue()
                    
                    # Tandai bahwa proses telah berhasil dan selesai
                    st.session_state.etl_step2_processed = True
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses data Langkah 2: {e}")
        else:
            st.warning("Mohon unggah semua file mentah Langkah 2 terlebih dahulu untuk memulai.")

    if st.session_state.etl_step2_processed:
        st.success("✅ Langkah 2 selesai!")
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📥 Download HasilTrucking.xlsx", 
                data=st.session_state.out_truck_file, 
                file_name="HasilTrucking.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_dl2:
            st.download_button(
                label="📥 Download HasilNonTrucking.xlsx", 
                data=st.session_state.out_nontruck_file, 
                file_name="HasilNonTrucking.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ------------------------------------------------------------------------------
# BAGIAN B: MENU FORECASTING
# ------------------------------------------------------------------------------
elif category_filter == "Proses Forecast":
    st.header("Forecast Penggunaan BBM")
    
    col1, col2 = st.columns(2)
    with col1:
        file_truck = st.file_uploader("1. Upload HasilTrucking.xlsx", type=["xlsx"])
        file_bbm = st.file_uploader("3. Upload BBM AAB.xlsx", type=["xlsx"])
    with col2:
        file_nontruck = st.file_uploader("2. Upload HasilNonTrucking.xlsx", type=["xlsx"])
        file_master = st.file_uploader("4. Upload Master cost & bbm 2022 sd 2025 HP & Type.xlsx", type=["xlsx"])
        
    if st.button("🚀 Jalankan Proses Forecast"):
        if file_truck and file_nontruck and file_bbm and file_master:
            with st.spinner("Memproses data historis dan melatih model Machine Learning..."):
                try:
                    def extract_lambung_code(raw_kpi):
                        s = str(raw_kpi).upper().strip()
                        match = re.search(r'\((.*?)\)', s)
                        if match: return match.group(1).strip()
                        parts = s.split()
                        if len(parts) > 1:
                            last_part = parts[-1]
                            if len(last_part) < 6: return last_part
                        return s.replace("-", "").strip()

                    def get_standar_pabrik_liter(jenis, merk, cap, hp, hm_pred):
                        jenis = str(jenis).upper()
                        merk = str(merk).upper()
                        try: cap = float(cap)
                        except: cap = 0.0
                        try: hp = float(hp)
                        except: hp = 0.0
                        
                        base_l_per_hm = 0.0
                        
                        if hp > 0:
                            if 'REACH STACKER' in jenis: load_factor = 0.055
                            elif 'FORKLIFT' in jenis: load_factor = 0.045
                            elif 'LOADER' in jenis: load_factor = 0.050
                            elif 'CRANE' in jenis: load_factor = 0.040
                            elif 'TRAILER' in jenis or 'HEAD' in jenis: load_factor = 0.020 
                            elif 'TRONTON' in jenis: load_factor = 0.018
                            else: load_factor = 0.04
                            
                            if 'KALMAR' in merk or 'VOLVO' in merk or 'SCANIA' in merk: load_factor *= 0.95 
                            elif 'SANY' in merk or 'WEICHAI' in merk: load_factor *= 1.05 
                            elif 'HINO' in merk or 'ISUZU' in merk: load_factor *= 0.98 
                            
                            base_l_per_hm = hp * load_factor
                            
                        else:
                            if 'REACH STACKER' in jenis:
                                if 'KALMAR' in merk: base_l_per_hm = 16.5
                                elif 'SANY' in merk: base_l_per_hm = 15.0
                                elif 'KONE' in merk: base_l_per_hm = 17.0
                                elif 'LINDE' in merk: base_l_per_hm = 16.0
                                else: base_l_per_hm = 16.0
                            elif 'FORKLIFT' in jenis:
                                if cap >= 25: base_l_per_hm = 9.0
                                elif cap >= 10: base_l_per_hm = 6.5
                                elif cap >= 5: base_l_per_hm = 4.0
                                else: base_l_per_hm = 2.5
                            elif 'LOADER' in jenis: 
                                if 'KALMAR' in merk: base_l_per_hm = 14.5
                                elif 'SANY' in merk: base_l_per_hm = 13.5
                                else: base_l_per_hm = 14.0
                            elif 'CRANE' in jenis:
                                if cap >= 70: base_l_per_hm = 14.0
                                elif cap >= 40: base_l_per_hm = 10.0
                                else: base_l_per_hm = 8.0
                            elif 'TRAILER' in jenis or 'HEAD' in jenis:
                                if 'HINO' in merk: base_l_per_hm = 5.5
                                elif 'ISUZU' in merk: base_l_per_hm = 5.0
                                elif 'UD' in merk or 'NISSAN' in merk: base_l_per_hm = 5.8
                                elif 'MERCEDES' in merk or 'BENZ' in merk: base_l_per_hm = 6.0
                                else: base_l_per_hm = 5.5
                            elif 'TRONTON' in jenis:
                                base_l_per_hm = 4.5
                            else:
                                base_l_per_hm = 5.0 
                                
                        min_l_per_hm = base_l_per_hm * 0.85
                        max_l_per_hm = base_l_per_hm * 1.15
                        
                        return hm_pred * min_l_per_hm, hm_pred * max_l_per_hm

                    master_data_map = {}
                    master_keys_set = set()
                    df_map = pd.read_excel(file_master, sheet_name='Sheet2', header=1)
                    col_name = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
                    col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
                    col_tipe = next((c for c in df_map.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
                    col_cap = next((c for c in df_map.columns if any(k in str(c).upper() for k in ['CAP', 'KAPASITAS'])), None)
                    col_hp = next((c for c in df_map.columns if any(k in str(c).upper() for k in ['HP', 'HORSE POWER'])), None)
                    
                    df_map.dropna(subset=[col_name], inplace=True)
                    
                    for _, row in df_map.iterrows():
                        u_name = str(row[col_name]).strip().upper()
                        c_id = clean_unit_name(u_name)
                        if c_id:
                            cap_val = 0.0
                            try:
                                if pd.notna(row[col_cap]):
                                    match_cap = re.search(r"(\d+(\.\d+)?)", str(row[col_cap]))
                                    if match_cap: cap_val = float(match_cap.group(1))
                            except: pass
                            
                            hp_val = 0.0
                            try:
                                if col_hp and pd.notna(row[col_hp]):
                                    match_hp = re.search(r"(\d+(\.\d+)?)", str(row[col_hp]))
                                    if match_hp: hp_val = float(match_hp.group(1))
                            except: pass
                            
                            master_data_map[c_id] = {
                                'Unit_Name': u_name,
                                'Jenis_Alat': str(row[col_jenis]).strip().upper() if pd.notna(row[col_jenis]) else "OTHERS",
                                'Merk': str(row[col_tipe]).strip().upper() if pd.notna(row[col_tipe]) else "-",
                                'Capacity': cap_val,
                                'HP': hp_val
                            }
                            master_keys_set.add(c_id)

                    def get_master_match(raw_name):
                        raw_name = str(raw_name).strip().upper()
                        c_raw = clean_unit_name(raw_name)
                        if c_raw in master_data_map: return c_raw
                        if " (" in raw_name:
                            b_paren = clean_unit_name(raw_name.split(" (")[0])
                            if b_paren in master_data_map: return b_paren
                        if "EX." in raw_name or "EX " in raw_name:
                            after_ex = raw_name.split("EX.")[-1] if "EX." in raw_name else raw_name.split("EX ")[-1]
                            c_after = clean_unit_name(after_ex.replace(")", ""))
                            if c_after in master_data_map: return c_after
                            for m_key in master_keys_set:
                                if c_after != "" and c_after in m_key: return m_key
                        return c_raw 

                    hm_data_store = {}
                    xls = pd.ExcelFile(file_bbm)
                    for sheet_name in xls.sheet_names:
                        sht_up = sheet_name.upper()
                        df_cek = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                        if df_cek.empty: continue
                        a1_val = str(df_cek.iloc[0, 0]).strip().upper()
                        if "EQUIP" not in a1_val and "NAMA" not in a1_val: continue 
                            
                        df_full = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                        r_eq, r_kpi, r_hd = 0, 1, 2
                        for idx in range(min(10, len(df_full))):
                            row_str = df_full.iloc[idx].astype(str).str.upper()
                            if row_str.str.contains('EQUIP NAME').any(): r_eq = idx
                            if row_str.str.contains('GROUP KPI').any(): r_kpi = idx
                            if row_str.str.contains('HM').any(): r_hd = idx
                                
                        unit_names_row = df_full.iloc[r_eq].ffill() 
                        group_kpi_row = df_full.iloc[r_kpi].ffill()  
                        headers = df_full.iloc[r_hd]                
                        data_rows = df_full.iloc[r_hd+1:]             
                        
                        def is_valid_date_row(val):
                            if pd.isna(val): return False
                            if isinstance(val, (int, float)) and 1 <= val <= 31: return True
                            val_str = str(val).strip()
                            if re.match(r'^\d{4}-\d{2}-\d{2}', val_str): return True 
                            match = re.match(r'^(\d{1,2})', val_str) 
                            if match and 1 <= int(match.group(1)) <= 31: return True
                            return False
                            
                        valid_mask = data_rows.iloc[:, 0].apply(is_valid_date_row)
                        valid_data_rows = data_rows[valid_mask]
                        dates_raw = valid_data_rows.iloc[:, 0]
                        month_idx = {'JAN':1, 'FEB':2, 'MAR':3, 'APR':4, 'MEI':5, 'JUN':6, 'JUL':7, 'AGT':8, 'SEP':9, 'OKT':10, 'NOV':11, 'DES': 12}
                        cur_month = next((v for k,v in month_idx.items() if k in sht_up), 1)
                        
                        dates_parsed = [pd.Timestamp(year=2025, month=cur_month, day=int(d)) if isinstance(d, (int, float)) else pd.to_datetime(d, dayfirst=True, errors='coerce') for d in dates_raw]
                        dates_series = pd.Series(dates_parsed)
                        
                        for col in range(1, df_full.shape[1]):
                            header_str = str(headers.iloc[col]).strip().upper()
                            is_hm = (header_str == 'HM')
                            if is_hm:
                                raw_equip = str(unit_names_row.iloc[col]).strip().upper()
                                raw_kpi = str(group_kpi_row.iloc[col]).strip().upper()
                                if "TOTAL" in raw_equip or "UNNAMED" in raw_equip or raw_equip == "NAN": continue
                                
                                lambung = extract_lambung_code(raw_kpi)
                                std_match = get_master_match(raw_equip)
                                final_clean_key = std_match if std_match in master_data_map else clean_unit_name(lambung if lambung else raw_equip)
                                
                                vals = pd.to_numeric(valid_data_rows.iloc[:, col], errors='coerce')
                                temp = pd.DataFrame({'Date': dates_series, 'Value': vals.values}).dropna(subset=['Date'])
                                
                                if final_clean_key not in hm_data_store: hm_data_store[final_clean_key] = []
                                hm_data_store[final_clean_key].append(temp)

                    def resolve_unit_type(clean_ops, master_map):
                        utype = "OTHERS"
                        if clean_ops in master_map: 
                            t = str(master_map[clean_ops]['Jenis_Alat']).upper()
                            if "TRONTON" in t: utype = "TRONTON"
                            elif "TRAILER" in t or "HEAD" in t: utype = "TRAILER"
                            elif "REACH" in t or "STACKER" in t or "SMV" in t: utype = "REACH STACKER"
                            elif "FORKLIFT" in t: utype = "FORKLIFT"
                            elif "CRANE" in t: utype = "CRANE"
                            elif "SIDE" in t: utype = "SIDE LOADER"
                            elif "TOP" in t: utype = "TOP LOADER"
                        if utype != "OTHERS": return utype
                        if "TRONTON" in clean_ops: return "TRONTON"
                        if "TRAILER" in clean_ops or re.search(r'L\s*\d+', clean_ops): return "TRAILER"
                        return "OTHERS"

                    def calculate_monthly_hm_for_unit(df_list):
                        if not df_list: return pd.DataFrame()
                        df = pd.concat(df_list, ignore_index=True).sort_values('Date')
                        df['HM_Clean'] = pd.to_numeric(df['Value'], errors='coerce').replace(0, np.nan).ffill().fillna(0)
                        df['Delta_HM'] = df['HM_Clean'].diff().fillna(0)
                        df.loc[df['Delta_HM'] < 0, 'Delta_HM'] = 0
                        df.loc[df['Delta_HM'] > 100, 'Delta_HM'] = 0 
                        df['Month_Num'] = df['Date'].dt.month
                        return df.groupby('Month_Num')['Delta_HM'].sum().reset_index()

                    df_tr_raw = pd.read_excel(file_truck, sheet_name='Data_Bulanan')
                    df_tr_raw = df_tr_raw[df_tr_raw['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
                    tr_list = []
                    for _, r in df_tr_raw.iterrows():
                        raw_u = str(r['Nama_Unit'])
                        clean_u = clean_unit_name(raw_u)
                        m_num = {'Oktober': 10, 'November': 11}.get(r['Bulan'].strip())
                        hm_val = 0
                        if clean_u in hm_data_store:
                            res_hm = calculate_monthly_hm_for_unit(hm_data_store[clean_u])
                            if not res_hm.empty:
                                row_hm = res_hm[res_hm['Month_Num'] == m_num]
                                if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                        tr_list.append({'Category': 'TRUCKING', 'Type': resolve_unit_type(clean_u, master_data_map), 'Unit_Clean': raw_u, 'Month_Num': m_num, 'LITER': r.get('LITER',0), 'Workload': r.get('Total_TonKm', 0), 'Ton': 0, 'HM': hm_val})
                    
                    df_nt_raw = pd.read_excel(file_nontruck, sheet_name='Data_Bulanan')
                    df_nt_raw = df_nt_raw[df_nt_raw['Bulan'].isin(['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November'])]
                    m_map_id = {'Januari':1, 'Februari':2, 'Maret':3, 'April':4, 'Mei':5, 'Juni':6, 'Juli':7, 'Agustus':8, 'September':9, 'Oktober':10, 'November':11}
                    nt_list = []
                    col_u = 'Unit_Name' if 'Unit_Name' in df_nt_raw.columns else 'Nama_Unit'
                    for _, r in df_nt_raw.iterrows():
                        raw_u = str(r[col_u])
                        clean_u = clean_unit_name(raw_u)
                        utype = resolve_unit_type(clean_u, master_data_map)
                        if utype in ['REACH STACKER','FORKLIFT','CRANE','SIDE LOADER','TOP LOADER']:
                            m_num = m_map_id.get(r['Bulan'].strip())
                            hm_val = 0
                            if clean_u in hm_data_store:
                                res_hm = calculate_monthly_hm_for_unit(hm_data_store[clean_u])
                                if not res_hm.empty:
                                    row_hm = res_hm[res_hm['Month_Num'] == m_num]
                                    if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                            nt_list.append({'Category': 'NON-TRUCKING', 'Type': utype, 'Unit_Clean': raw_u, 'Month_Num': m_num, 'LITER': r.get('LITER',0), 'Ton': r.get('Total_Ton', 0), 'Workload': 0, 'HM': hm_val})
                    
                    df_final = pd.concat([pd.DataFrame(tr_list), pd.DataFrame(nt_list)], ignore_index=True)
                    configs = [{'cat': 'TRUCKING', 'type': 'TRAILER', 'preds': ['Workload', 'HM'], 'range': [10, 11]}, {'cat': 'TRUCKING', 'type': 'TRONTON', 'preds': ['Workload', 'HM'], 'range': [10, 11]}, {'cat': 'NON-TRUCKING', 'type': 'REACH STACKER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}, {'cat': 'NON-TRUCKING', 'type': 'FORKLIFT', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}, {'cat': 'NON-TRUCKING', 'type': 'CRANE', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}, {'cat': 'NON-TRUCKING', 'type': 'SIDE LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}, {'cat': 'NON-TRUCKING', 'type': 'TOP LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}]
                    forecast_detail_list = []
                    
                    for cfg in configs:
                        sub = df_final[(df_final['Category']==cfg['cat']) & (df_final['Type']==cfg['type']) & (df_final['Month_Num'].isin(cfg['range']))]
                        if sub.empty: continue
                        agg = sub.groupby('Month_Num').agg({'LITER': 'sum', 'Unit_Clean': 'nunique', **{p: 'sum' for p in cfg['preds']}}).reset_index()
                        
                        valid_model = False
                        rm = LinearRegression()
                        r2_model = None
                        
                        if len(agg) >= 2:
                            for col in cfg['preds'] + ['LITER']: agg[f'{col}_Per_Unit'] = agg[col] / agg['Unit_Clean']
                            X_train = agg[[f'{p}_Per_Unit' for p in cfg['preds']]]
                            y_train = agg['LITER_Per_Unit']
                            if X_train.sum().sum() > 0:
                                rm.fit(X_train, y_train)
                                r2_model = r2_score(y_train, rm.predict(X_train))
                                valid_model = True

                        unique_units = sub['Unit_Clean'].unique()
                        
                        for unit in unique_units:
                            unit_history = sub[sub['Unit_Clean'] == unit]
                            unit_pred_act = {}
                            
                            for p in cfg['preds']:
                                act_series = unit_history[['Month_Num', p]].dropna()
                                if len(act_series) >= 2 and act_series[p].sum() > 0:
                                    tm = LinearRegression()
                                    tm.fit(act_series[['Month_Num']], act_series[p])
                                    pred_val = tm.predict([[12]])[0]
                                    unit_pred_act[p] = max(0, pred_val)
                                elif len(act_series) == 1:
                                    unit_pred_act[p] = act_series[p].values[0]
                                else:
                                    unit_pred_act[p] = 0
                                    
                            pred_wl_ton = unit_pred_act.get('Workload', unit_pred_act.get('Ton', 0))
                            pred_hm_final = unit_pred_act.get('HM', 0)
                            
                            c_unit = clean_unit_name(unit)
                            meta_unit = master_data_map.get(c_unit, {})
                            merk_unit = meta_unit.get('Merk', '-')
                            cap_unit = meta_unit.get('Capacity', 0.0)
                            hp_unit = meta_unit.get('HP', 0.0)
                            
                            std_pabrik_min, std_pabrik_max = get_standar_pabrik_liter(cfg['type'], merk_unit, cap_unit, hp_unit, pred_hm_final)
                            
                            expected_liter_normal = 0
                            use_ml = False
                            
                            if valid_model: 
                                X_pred = [unit_pred_act.get(p, 0) for p in cfg['preds']]
                                ml_pred = rm.predict([X_pred])[0]
                                if ml_pred > 0:
                                    expected_liter_normal = ml_pred
                                    use_ml = True
                                    
                            ratios_eff, actual_lits = [], []
                            for _, row in unit_history.iterrows():
                                actual_lit = row['LITER']
                                if pd.isna(actual_lit): actual_lit = 0
                                
                                if actual_lit > 0:
                                    actual_lits.append(actual_lit)
                                    if use_ml:
                                        X_hist = [row[p] for p in cfg['preds']]
                                        normal_hist = rm.predict([X_hist])[0]
                                        if normal_hist > 0:
                                            ratios_eff.append(actual_lit / normal_hist)
                                            
                            eff_factor = 1.0
                            final_forecast = 0
                            note = "Data Kurang / ML Gagal"
                            error_margin_val = None
                            
                            if use_ml and expected_liter_normal > 0:
                                if ratios_eff:
                                    eff_factor = np.mean(ratios_eff)
                                    eff_factor = max(0.80, min(1.20, eff_factor))
                                    if len(actual_lits) > 1 and np.mean(actual_lits) > 0:
                                        cv = np.std(actual_lits) / np.mean(actual_lits)
                                        error_margin_val = round(cv * 100, 2)
                                    else:
                                        error_margin_val = 0.0
                                final_forecast = expected_liter_normal * eff_factor
                                
                                if ratios_eff:
                                    trend_str = "Boros" if eff_factor > 1.05 else ("Irit" if eff_factor < 0.95 else "Standard")
                                    note = f"ML Regression -> {trend_str}"
                                else:
                                    note = "ML Regression -> Tidak Ada History BBM Valid"
                            else:
                                if pred_hm_final == 0:
                                    note = "Tidak Ada Prediksi Aktivitas (HM=0)"
                                else:
                                    note = "Data Historis Tidak Cukup / ML Gagal (Anomali)"
                                    
                            r2_val = round(r2_model, 2) if use_ml else None
                            
                            forecast_detail_list.append({
                                'Category': cfg['cat'], 
                                'Type': cfg['type'], 
                                'Unit_Name': unit,
                                'Forecast_HM_Dec': round(pred_hm_final, 2), 
                                'Forecast_Workload_Ton_Dec': round(pred_wl_ton, 2), 
                                'Expected_BBM_Normal': round(expected_liter_normal, 2), 
                                'Standar_BBM_Pabrik_Min': round(std_pabrik_min, 2),
                                'Standar_BBM_Pabrik_Max': round(std_pabrik_max, 2),
                                'Correction_Factor': round(eff_factor, 2),
                                'Forecast_BBM_Dec': round(final_forecast, 2), 
                                'Akurasi_General_R2': r2_val,
                                'Error_Fluktuasi_Unit_Pct': error_margin_val,
                                'Note': note
                            })
                    
                    df_res = pd.DataFrame(forecast_detail_list)
                    
                    cols_order = [
                        'Category', 'Type', 'Unit_Name', 'Forecast_HM_Dec', 'Forecast_Workload_Ton_Dec',
                        'Expected_BBM_Normal', 'Standar_BBM_Pabrik_Min', 'Standar_BBM_Pabrik_Max', 
                        'Correction_Factor', 'Forecast_BBM_Dec', 'Akurasi_General_R2', 
                        'Error_Fluktuasi_Unit_Pct', 'Note'
                    ]
                    df_res = df_res[[c for c in cols_order if c in df_res.columns]]
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, index=False, sheet_name='Forecast_Desember')
                    
                    # Simpan hasil ke Session State
                    st.session_state.fcst_df_res = df_res
                    st.session_state.fcst_df_final = df_final
                    st.session_state.fcst_out_file = output.getvalue()
                    st.session_state.forecast_processed = True
                    
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses data: {e}")
        else:
            st.warning("Mohon unggah keempat file di atas terlebih dahulu untuk memulai.")

    # Jika data forecast sudah pernah diproses, tampilkan UI ini
    if st.session_state.forecast_processed:
        st.success("✅ Proses Forecast Selesai!")
        
        df_res = st.session_state.fcst_df_res
        df_final = st.session_state.fcst_df_final

        # ===================================================================================
        # AUDIT KELENGKAPAN DATA (DATA COMPLETENESS)
        # ===================================================================================
        st.markdown("### 📊 Kelengkapan Data untuk Forecast")
        
        # Logika pemisahan unit berhasil vs gagal (0 BBM di bulan depan)
        df_gagal = df_res[df_res['Forecast_BBM_Dec'] == 0].copy()
        df_berhasil = df_res[df_res['Forecast_BBM_Dec'] > 0].copy()
        
        total_unit = len(df_res)
        total_gagal = len(df_gagal)
        total_berhasil = len(df_berhasil)
        pct_gagal = (total_gagal / total_unit * 100) if total_unit > 0 else 0
        
        gagal_trucking = len(df_gagal[df_gagal['Category'] == 'TRUCKING'])
        gagal_nontrucking = len(df_gagal[df_gagal['Category'] == 'NON-TRUCKING'])
        
        # 1. Metric Cards
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total Populasi Unit", f"{total_unit} Unit")
        col_m2.metric("Persentase Unit Tidak Masuk Analisis", f"{pct_gagal:.1f}%")
        col_m3.metric("Rincian Unit yang Tidak Masuk Analisis", f"Truck: {gagal_trucking} | Non-Truck: {gagal_nontrucking}")
        
        # 2. Layout untuk Donut Chart & Expander
        col_chart, col_exp = st.columns([1, 2])
        
        with col_chart:
            # Donut Chart (Kesehatan Data)
            pie_data = pd.DataFrame({
                'Status': ['Berhasil di-Forecast', 'Data Kosong/Anomali'],
                'Jumlah': [total_berhasil, total_gagal]
            })
            fig_pie = px.pie(pie_data, names='Status', values='Jumlah', hole=0.5,
                             color='Status', color_discrete_map={'Berhasil di-Forecast': '#2ca02c', 'Data Kosong/Anomali': '#d62728'})
            fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_exp:
            st.markdown("<br>", unsafe_allow_html=True) # Spacer agar sejajar
            with st.expander("📋 Daftar unit yang tidak masuk analisis (Data Kosong/Anomali)", expanded=True):
                if not df_gagal.empty:
                    # Menampilkan tabel detail untuk evaluasi manajemen
                    st.dataframe(df_gagal[['Category', 'Type', 'Unit_Name', 'Note']].reset_index(drop=True), use_container_width=True)
                else:
                    st.success("Luar biasa! 100% data unit berhasil dianalisa tanpa ada yang kosong.")

        st.markdown("---")
        # ===================================================================================

        # -----------------------------------------------------------------------------------
        # GRAFIK TREN HISTORIS VS FORECAST PER UNIT
        # -----------------------------------------------------------------------------------
        st.markdown("### 📈 Tren Penggunaan BBM per Unit (Historis vs Forecast)")
        try:
            # 1. Siapkan data historis
            df_hist = df_final[['Unit_Clean', 'Month_Num', 'LITER']].copy()
            df_hist.rename(columns={'Unit_Clean': 'Nama Unit'}, inplace=True)
            
            # 2. Siapkan data forecast
            df_fcst = df_res[['Unit_Name', 'Forecast_BBM_Dec']].copy()
            df_fcst.rename(columns={'Unit_Name': 'Nama Unit', 'Forecast_BBM_Dec': 'LITER'}, inplace=True)
            df_fcst['Month_Num'] = 12
            
            unit_list = sorted(df_hist['Nama Unit'].unique().tolist())
            selected_unit_fcst = st.selectbox("Pilih Unit untuk melihat visualisasi tren historis ke bulan Desember:", unit_list)
            
            # 3. Filter data berdasarkan unit yang dipilih
            df_plot_hist = df_hist[df_hist['Nama Unit'] == selected_unit_fcst].groupby('Month_Num', as_index=False)['LITER'].sum()
            df_plot_hist['Tipe'] = 'Historis'
            
            df_plot_fcst = df_fcst[df_fcst['Nama Unit'] == selected_unit_fcst].groupby('Month_Num', as_index=False)['LITER'].sum()
            df_plot_fcst['Tipe'] = 'Forecast'
            
            # 4. Trik agar garis menyambung: Tambahkan titik terakhir historis ke data forecast
            if not df_plot_hist.empty:
                last_hist = df_plot_hist[df_plot_hist['Month_Num'] == df_plot_hist['Month_Num'].max()].copy()
                last_hist['Tipe'] = 'Forecast'
                df_plot_fcst = pd.concat([last_hist, df_plot_fcst], ignore_index=True)
                
            df_plot_combined = pd.concat([df_plot_hist, df_plot_fcst], ignore_index=True)
            
            month_map = {1:'Januari', 2:'Februari', 3:'Maret', 4:'April', 5:'Mei', 6:'Juni', 7:'Juli', 8:'Agustus', 9:'September', 10:'Oktober', 11:'November', 12:'Desember (Forecast)'}
            df_plot_combined['Bulan'] = df_plot_combined['Month_Num'].map(month_map)
            df_plot_combined = df_plot_combined.sort_values(['Month_Num', 'Tipe'], ascending=[True, False])
            
            # 5. Plotting dengan warna yang dibedakan (Biru untuk Historis, Merah untuk Forecast)
            fig_fcst = px.line(df_plot_combined, x='Bulan', y='LITER', color='Tipe', markers=True, 
                               title=f"Tren Penggunaan BBM (Liter): {selected_unit_fcst}",
                               color_discrete_map={'Historis': '#1f77b4', 'Forecast': '#d62728'})
            
            fig_fcst.update_traces(line=dict(width=3), marker=dict(size=8), 
                                   text=df_plot_combined['LITER'].apply(lambda x: f"{x:,.0f}"), 
                                   textposition="top center")
            
            fig_fcst.update_layout(yaxis_title="Total Konsumsi BBM (Liter)", xaxis_title="Bulan")
            fig_fcst.update_yaxes(rangemode="tozero")
            
            st.plotly_chart(fig_fcst, use_container_width=True)
        except Exception as e:
            st.warning(f"Gagal memuat grafik tren: {e}")
        # -----------------------------------------------------------------------------------

        st.dataframe(df_res, use_container_width=True)
        
        st.download_button(label="📥 Download Hasil Forecast (.xlsx)", data=st.session_state.fcst_out_file, file_name="Forecast_Desember.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------------------------------------------------------
# BAGIAN C: DASHBOARD UTAMA (TRUCKING / NON-TRUCKING)
# ------------------------------------------------------------------------------
else:
    BIAYA_PER_LITER = st.sidebar.number_input("Biaya Bahan Bakar (Rp/Liter)", min_value=0, value=6800, step=100)

    df_active_raw = pd.DataFrame()
    df_monthly = pd.DataFrame()
    df_missing = pd.DataFrame()

    if category_filter == "Analisa Trucking":
        with st.spinner("Memproses Data Trucking..."):
            df_active_raw, df_monthly, df_missing = process_trucking()
            mode_label = "Trucking"
            ratio_label = "L/Ton*Km"
            work_col = "Total Kerja (Ton*Km)"
    else:
        with st.spinner("Memuat Data Non-Trucking..."):
            df_active_raw, df_monthly, df_missing = process_alat_berat()
            mode_label = "Non-Trucking"
            ratio_label = "L/Ton"
            work_col = "Total Berat Angkutan (Ton)"

    if not df_active_raw.empty:
        
        # --- PROSES UNIT INAKTIF (DENGAN KETERANGAN DINAMIS) ---
        df_inaktif_from_active = df_active_raw[
            (df_active_raw['Total Pengisian BBM (L)'] <= 0) | 
            (df_active_raw[work_col] <= 0)
        ].copy()
        
        if not df_inaktif_from_active.empty:
            col_bbm = 'Total Pengisian BBM (L)'
            
            conditions = [
                (df_inaktif_from_active[col_bbm] <= 0) & (df_inaktif_from_active[work_col] <= 0),
                (df_inaktif_from_active[work_col] <= 0),
                (df_inaktif_from_active[col_bbm] <= 0)
            ]
            choices = [
                "Tidak ada aktivitas",
                "Unit tidak melakukan aktivitas kerja",
                "Unit tidak pernah mengisi BBM"
            ]
            df_inaktif_from_active['Keterangan'] = np.select(conditions, choices, default="-")

        list_inaktif = []
        if not df_inaktif_from_active.empty: list_inaktif.append(df_inaktif_from_active)
        
        if not df_missing.empty:
            if 'Total Pengisian BBM' in df_missing.columns:
                 df_missing.rename(columns={'Total Pengisian BBM': 'Total Pengisian BBM (L)'}, inplace=True)
            
            m_bbm_col = 'Total Pengisian BBM (L)'
            m_work_col = work_col 
            
            if m_bbm_col in df_missing.columns and m_work_col in df_missing.columns:
                 m_conds = [
                    (df_missing[m_bbm_col] <= 0) & (df_missing[m_work_col] <= 0),
                    (df_missing[m_work_col] <= 0),
                    (df_missing[m_bbm_col] <= 0)
                 ]
                 m_choices = [
                    "Tidak ada aktivitas",
                    "Unit tidak melakukan aktivitas kerja",
                    "Unit tidak pernah mengisi BBM"
                 ]
                 df_missing['Keterangan'] = np.select(m_conds, m_choices, default="Inaktif (Sumber: File Audit)")
            else:
                 df_missing['Keterangan'] = "Inaktif (Sumber: File Audit)"

            list_inaktif.append(df_missing)

        df_inaktif_all = pd.concat(list_inaktif, ignore_index=True) if list_inaktif else pd.DataFrame()
        
        df_active = df_active_raw[(df_active_raw['Total Pengisian BBM (L)'] > 0) & (df_active_raw[work_col] > 0)].copy()
        df_full_for_filter = pd.concat([df_active, df_inaktif_all], ignore_index=True) if not df_inaktif_all.empty else df_active

        st.sidebar.markdown("---")
        st.sidebar.header("Filter Data")
        
        lokasi_list = ["Semua"] + sorted(df_full_for_filter['Lokasi'].dropna().unique().tolist())
        selected_lokasi = st.sidebar.selectbox("📍 Filter Lokasi", lokasi_list)
        
        jenis_list = ["Semua"] + sorted(df_full_for_filter['Jenis'].dropna().unique().tolist())
        selected_jenis = st.sidebar.selectbox("🚜 Filter Jenis", jenis_list)
        
        type_list = ["Semua"] + sorted(df_full_for_filter['Type/Merk'].dropna().astype(str).unique().tolist())
        selected_type = st.sidebar.selectbox("🏷️ Filter Type/Merk", type_list)
        
        st.markdown("### 🛑 Daftar Unit Inaktif")
        st.caption("Unit yang terdeteksi tidak aktif karena tidak ada pengisian BBM atau tidak ada aktivitas kerja yang berlangsung")
        
        df_inaktif_filtered = df_inaktif_all.copy()
        if not df_inaktif_filtered.empty:
            if selected_lokasi != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Lokasi'] == selected_lokasi]
            if selected_jenis != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Jenis'] == selected_jenis]
            if selected_type != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Type/Merk'] == selected_type]
            
            if not df_inaktif_filtered.empty:
                if mode_label == "Trucking":
                    cols_inaktif = ['Nama Unit', 'Jenis', 'Type/Merk', 'Lokasi', 'Horse Power', 'Capacity (Feet)', 'Total Pengisian BBM (L)', work_col, 'Keterangan']
                    fmt_dict_inaktif = {'Capacity (Feet)': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}'}
                else:
                    cols_inaktif = ['Nama Unit', 'Jenis', 'Type/Merk', 'Lokasi', 'Horse Power', 'Capacity (Ton)', 'Total Pengisian BBM (L)', work_col, 'Keterangan']
                    fmt_dict_inaktif = {'Capacity (Ton)': '{:.0f}', 'Horse Power': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}'}
                    
                cols_to_show = [c for c in cols_inaktif if c in df_inaktif_filtered.columns]
                st.dataframe(df_inaktif_filtered[cols_to_show].style.format(fmt_dict_inaktif, na_rep="-"), use_container_width=True)
            else:
                st.success("Tidak ada unit inaktif untuk kombinasi filter ini.")
        else:
            st.success("Seluruh unit beroperasi aktif dan masuk ke dalam perhitungan dashboard.")
            
        st.markdown("---")
        
        st.markdown("### 🔍 Cari Data Spesifik (Unit Aktif)")
        c_search1, c_search2 = st.columns([1, 3])
        with c_search1:
            search_category = st.selectbox("Cari Berdasarkan:", ["Nama Unit"])
        with c_search2:
            search_query = st.text_input(f"Ketik {search_category}:", "")

        df_filtered = df_active.copy()
        if selected_lokasi != "Semua": df_filtered = df_filtered[df_filtered['Lokasi'] == selected_lokasi]
        if selected_jenis != "Semua": df_filtered = df_filtered[df_filtered['Jenis'] == selected_jenis]
        if selected_type != "Semua": df_filtered = df_filtered[df_filtered['Type/Merk'] == selected_type]

        if search_query:
            if search_category == "Nama Unit": df_filtered = df_filtered[df_filtered['Nama Unit'].str.contains(search_query, case=False, na=False)]
        
        df_filtered['Total Biaya BBM'] = df_filtered['Total Pengisian BBM (L)'] * BIAYA_PER_LITER
        
        if not df_monthly.empty:
            df_monthly_filtered = df_monthly[df_monthly['Nama Unit'].isin(df_filtered['Nama Unit'])].copy()
        else:
            df_monthly_filtered = pd.DataFrame()

        if mode_label == "Trucking":
            total_bbm = df_filtered['Total Pengisian BBM (L)'].sum()
            total_kerja = df_filtered['Total Kerja (Ton*Km)'].sum()
            total_biaya = df_filtered['Total Biaya BBM'].sum()
            total_unit_aktif = len(df_filtered)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Unit Aktif", f"{total_unit_aktif} Unit")
            c2.metric("Total Kerja (Ton*Km)", f"{total_kerja:,.0f}")
            c3.metric("Total Pengisian BBM", f"{total_bbm:,.0f} L")
            c4.metric("Total Biaya BBM (Rp)", f"Rp {total_biaya:,.0f}")
        else:
            total_bbm = df_filtered['Total Pengisian BBM (L)'].sum()
            total_ton = df_filtered['Total Berat Angkutan (Ton)'].sum()
            total_biaya = df_filtered['Total Biaya BBM'].sum()
            total_unit_aktif = len(df_filtered)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Unit Aktif", f"{total_unit_aktif} Unit")
            c2.metric("Total Tonase Container", f"{total_ton:,.0f} Ton")
            c3.metric("Total Pengisian BBM", f"{total_bbm:,.0f} L")
            c4.metric("Total Biaya BBM (Rp)", f"Rp {total_biaya:,.0f}")

        st.markdown("---")

        tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview Data", "📈 Efisiensi Setiap Unit", "📍 Korelasi Beban & BBM", "💸 Unit Terboros"])

        # TAB 1: OVERVIEW
        with tab1:
            st.subheader(f"Data Detail {mode_label}")
            
            sort_options = ["Fuel Ratio (Tertinggi)", "Fuel Ratio (Terendah)", "Total Kerja (Tertinggi)", "Total Pengisian BBM (L) (Tertinggi)"]
            sort_by = st.selectbox("Sort by:", sort_options)
            
            ratio_col = 'Fuel Ratio (L/Ton*Km)' if mode_label == "Trucking" else 'Fuel Ratio (L/Ton)'
            bm_col = 'Benchmark (L/Ton*Km)' if mode_label == "Trucking" else 'Benchmark (L/Ton)'

            if sort_by == "Fuel Ratio (Tertinggi)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=False)
            elif sort_by == "Fuel Ratio (Terendah)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=True)
            elif sort_by == "Total Kerja (Tertinggi)": df_filtered = df_filtered.sort_values(by=work_col, ascending=False)
            elif sort_by == "Total Pengisian BBM (L) (Tertinggi)": df_filtered = df_filtered.sort_values(by='Total Pengisian BBM (L)', ascending=False)
            
            def highlight_fuel_ratio(row):
                styles = [''] * len(row)
                for i, col in enumerate(row.index):
                    if col == ratio_col:
                        val = row[col]
                        bm = row[bm_col]
                        if pd.notna(val) and pd.notna(bm) and bm > 0:
                            if val > bm:
                                styles[i] = 'background-color: #d62728; color: white; font-weight: bold;' 
                            else:
                                styles[i] = 'background-color: #2ca02c; color: white; font-weight: bold;'
                return styles

            if mode_label == "Trucking":
                cols_show = ['Nama Unit', 'Jenis', 'Lokasi', 'Horse Power', 'Capacity (Feet)', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Total Kerja (Ton*Km)', 'Benchmark (L/Ton*Km)', 'Fuel Ratio (L/Ton*Km)', 'Potensi Pemborosan BBM (L)']
                format_dict = {'Capacity (Feet)': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Total Kerja (Ton*Km)': '{:,.0f}', 'Benchmark (L/Ton*Km)': '{:.4f}', 'Fuel Ratio (L/Ton*Km)': '{:.4f}', 'Potensi Pemborosan BBM (L)': '{:,.0f}'}
                st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))
            else:
                cols_show = ['Nama Unit', 'Jenis', 'Type/Merk', 'Horse Power', 'Capacity (Ton)', 'Lokasi', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Benchmark (L/Ton)', 'Fuel Ratio (L/Ton)', 'Potensi Pemborosan BBM (L)']
                format_dict = {'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Benchmark (L/Ton)': '{:.4f}', 'Fuel Ratio (L/Ton)': '{:.4f}', 'Potensi Pemborosan BBM (L)': '{:,.0f}'}
                st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))

            # TREN BULANAN
            st.markdown("---")
            st.subheader("📈 Tren Kinerja Bulanan Setiap Unit")
            if not df_monthly_filtered.empty:
                unit_list_trend = sorted(df_monthly_filtered['Nama Unit'].unique().tolist())
                selected_unit_trend = st.selectbox("Pilih Unit untuk melihat tren:", unit_list_trend)
                
                trend_data_unit = df_monthly_filtered[df_monthly_filtered['Nama Unit'] == selected_unit_trend]
                month_order = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
                trend_data_unit['Bulan'] = pd.Categorical(trend_data_unit['Bulan'], categories=month_order, ordered=True)
                
                if mode_label == "Trucking":
                    trend_data = trend_data_unit.groupby('Bulan', as_index=False).agg({'Total Kerja (Ton*Km)': 'sum', 'Total Pengisian BBM (L)': 'sum'})
                    trend_data = trend_data.dropna()
                    trend_data['Ratio'] = np.where(trend_data['Total Kerja (Ton*Km)'] > 0, trend_data['Total Pengisian BBM (L)'] / trend_data['Total Kerja (Ton*Km)'], 0)
                    
                    c_trend1, c_trend2 = st.columns(2)
                    with c_trend1:
                        fig_trend_work = px.bar(trend_data, x='Bulan', y='Total Kerja (Ton*Km)', text_auto='.2s', title=f"Tren Kerja (Ton*Km): {selected_unit_trend}", color_discrete_sequence=['#1f77b4'])
                        st.plotly_chart(fig_trend_work, use_container_width=True)
                    with c_trend2:
                        fig_trend_ratio = px.line(trend_data, x='Bulan', y='Ratio', markers=True, title=f"Tren Efisiensi (L/Ton*Km): {selected_unit_trend}", color_discrete_sequence=['#ff7f0e'])
                        fig_trend_ratio.update_yaxes(rangemode="tozero")
                        st.plotly_chart(fig_trend_ratio, use_container_width=True)
                else:
                    trend_data = trend_data_unit.groupby('Bulan', as_index=False).agg({'Total Berat Angkutan (Ton)': 'sum', 'Total Pengisian BBM (L)': 'sum'})
                    trend_data = trend_data.dropna()
                    trend_data['Ratio'] = np.where(trend_data['Total Berat Angkutan (Ton)'] > 0, trend_data['Total Pengisian BBM (L)'] / trend_data['Total Berat Angkutan (Ton)'], 0)
                    
                    c_trend1, c_trend2 = st.columns(2)
                    with c_trend1:
                        fig_trend_ton = px.bar(trend_data, x='Bulan', y='Total Berat Angkutan (Ton)', text_auto='.2s', title=f"Tren Berat Angkutan: {selected_unit_trend}", color_discrete_sequence=['#1f77b4'])
                        st.plotly_chart(fig_trend_ton, use_container_width=True)
                    with c_trend2:
                        fig_trend_ratio = px.line(trend_data, x='Bulan', y='Ratio', markers=True, title=f"Tren Efisiensi (L/Ton): {selected_unit_trend}", color_discrete_sequence=['#ff7f0e'])
                        fig_trend_ratio.update_yaxes(rangemode="tozero")
                        st.plotly_chart(fig_trend_ratio, use_container_width=True)

        # TAB 2: BAR CHART
        with tab2:
            st.subheader(f"Peringkat Efisiensi BBM ({ratio_label})")
            df_chart = df_filtered.sort_values(ratio_col, ascending=False)
            
            fig_bar = px.bar(df_chart, x='Nama Unit', y=ratio_col, color='Status',
                             custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                             title=f"Fuel Ratio Seluruh Unit ({ratio_label})",
                             color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'})
            
            fig_bar.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                              "<b>Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # TAB 3: SCATTER PLOT
        with tab3:
            st.subheader("Korelasi Beban Kerja vs BBM")
            max_bubble_size = 45 
            
            if mode_label == "Trucking":
                vals = df_filtered[ratio_col].fillna(0)
                if vals.max() > vals.min():
                     size_col = 10 + ((vals - vals.min()) / (vals.max() - vals.min())) * (max_bubble_size - 10)
                else:
                     size_col = vals.apply(lambda x: 20)

                fig_scat = px.scatter(df_filtered, x='Total Kerja (Ton*Km)', y='Total Pengisian BBM (L)', color='Status',
                                    custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                                    size=size_col, 
                                    size_max=max_bubble_size, opacity=0.65,
                                    color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                    title="Korelasi Beban Kerja (Total Kerja (Ton*Km)) vs Total Pengisian BBM (L)")
            else:
                size_col_ton = df_filtered[ratio_col].apply(lambda x: x if x > 0 else 0.0001)
                
                fig_scat = px.scatter(df_filtered, x='Total Berat Angkutan (Ton)', y='Total Pengisian BBM (L)', color='Status',
                                        custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                                        size=size_col_ton,
                                        size_max=max_bubble_size, opacity=0.65,
                                        color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                        title="Korelasi Total Berat Angkutan (Ton) vs Total Pengisian BBM (L)")
            
            fig_scat.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                              "<b>Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
            )
            st.plotly_chart(fig_scat, use_container_width=True)

        # TAB 4: UNIT TERBOROS
        with tab4:
            st.subheader("💸 Top 10 Unit dengan Pemborosan Tertinggi")
            
            if mode_label == "Trucking":
                df_waste = df_filtered[df_filtered['Status'] == 'Boros'].copy()
                df_waste = df_waste.sort_values(by='Potensi Pemborosan BBM (L)', ascending=False).head(10)
                
                if not df_waste.empty:
                    fig_waste = px.bar(df_waste, x='Nama Unit', y='Potensi Pemborosan BBM (L)', text_auto='.0f',
                                    title="Estimasi Pemborosan (Liter BBM)", color_discrete_sequence=['#d62728'])
                    st.plotly_chart(fig_waste, use_container_width=True)
                    st.write(df_waste[['Nama Unit', 'Total Kerja (Ton*Km)', 'Fuel Ratio (L/Ton*Km)', 'Benchmark (L/Ton*Km)', 'Potensi Pemborosan BBM (L)']])
                else:
                    st.success("Tidak ada unit yang tergolong BOROS saat ini.")
            else:
                df_waste = df_filtered[df_filtered['Potensi Pemborosan BBM (L)'] > 0].sort_values(by='Potensi Pemborosan BBM (L)', ascending=False).head(10)
                
                if not df_waste.empty:
                    fig_waste = px.bar(df_waste, x='Nama Unit', y='Potensi Pemborosan BBM (L)', text_auto='.0f',
                                    title="Estimasi Pemborosan (Liter BBM)", color_discrete_sequence=['#d62728'])
                    st.plotly_chart(fig_waste, use_container_width=True)
                    st.write(df_waste[['Nama Unit', 'Total Berat Angkutan (Ton)', 'Fuel Ratio (L/Ton)', 'Benchmark (L/Ton)', 'Potensi Pemborosan BBM (L)']])
                else:
                    st.success("Tidak ada unit yang tergolong BOROS saat ini.")

    elif not df_active_raw.empty and df_filtered.empty:
        st.warning("⚠️ Tidak ada unit yang cocok dengan kombinasi filter & pencarian Anda.")