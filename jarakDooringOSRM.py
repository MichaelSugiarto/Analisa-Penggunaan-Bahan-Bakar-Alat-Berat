import pandas as pd
import requests
import json
import time
import re
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic

# ==============================================================================
# KONFIGURASI & MANUAL FIX
# ==============================================================================
FILE_DOORING = "DOORING OKT-DES 2025 (Copy).xlsx" 
FILE_TLP = "dashboard TLP okt-des (S1L Trucking).xlsx"
OUTPUT_FILE = "DOORING_WITH_DISTANCE.xlsx"

# OSRM ENGINE
OSRM_URL = "http://router.project-osrm.org/route/v1/driving"

# DEPO: Laksda M. Nasir (Sesuai Request: Depo Yon)
DEPO_LAT = -7.2145
DEPO_LON = 112.7238

# --- DAFTAR PERBAIKAN ALAMAT MANUAL (Sesuai Chat Terakhir) ---
MANUAL_FIX = {
    "GARUDA FOOD JAYA PT": "DUSUN LARANGAN, KRIKILAN, GRESIK REGENCY, EAST JAVA, INDONESIA",
    "JALAN RAYA DAENDELS 64-65 KM": "TANJUNG PAKIS, KEMANTREN, PACIRAN, LAMONGAN REGENCY, EAST JAVA 62264, INDONESIA",
    "JALAN RAYA SURABAYA - MALANG KM 48.5": "KALI TENGAH, KARANG JATI, KEC. PANDAAN, PASURUAN, JAWA TIMUR 67156, INDONESIA",
    "JL. GEMPOL - MOJOKERTO NO.102": "NGORO, SEDATI, KEC. NGORO, KABUPATEN MOJOKERTO, JAWA TIMUR 61385, INDONESIA",
    "JL. RAYA CANGKRINGMALANG, TUREN": "TUREN, CANGKRINGMALANG, KEC. BEJI, PASURUAN, JAWA TIMUR 67154, INDONESIA",
    "JL. RAYA GRESIK - BABAT, GAJAH": "GAJAH, REJOSARI, KEC. DEKET, KABUPATEN LAMONGAN, JAWA TIMUR, INDONESIA",
    "JL. RAYA GRESIK - BABAT, REJOSARI": "REJOSARI, KEC. DEKET, KABUPATEN LAMONGAN, JAWA TIMUR, INDONESIA",
    "JL. RAYA MADIUN - SURABAYA": "BANJARAGUNG, KRIAN, MOJOKERTO, EAST JAVA, INDONESIA",
    "JL. RAYA MOJOKERTO SURABAYA NO.KM. 19": "BRINGIN WETAN, BRINGINBENDO, KEC. TAMAN, KABUPATEN SIDOARJO, JAWA TIMUR 61257, INDONESIA",
    "JL. RAYA SURABAYA - MALANG NO.52": "JERUKUWIK, NGADIMULYO, KEC. SUKOREJO, PASURUAN, JAWA TIMUR 67161, INDONESIA",
    "JL. RAYA SURABAYA - MALANG NO.KM 40": "KECINCANG, NGERONG, KEC. GEMPOL, PASURUAN, JAWA TIMUR 67155, INDONESIA",
    "JL. RAYA SURABAYA - MALANG, TAMBAK": "TAMBAK, NGADIMULYO, KEC. SUKOREJO, PASURUAN, JAWA TIMUR, INDONESIA",
    "KEMIRI SEWU, PANDAAN": "KEMIRI SEWU, KEC. PANDAAN, PASURUAN, JAWA TIMUR",
    "KRAJAN, SUMENGKO": "KRAJAN, SUMENGKO, KEC. WRINGINANOM, KABUPATEN GRESIK, JAWA TIMUR",
    "MANYARSIDORUKUN": "MANYAR SIDO RUKUN, MANYARSIDORUKUN, KEC. MANYAR, KABUPATEN GRESIK, JAWA TIMUR",
    "MASJID NURUL JANNAH": "NGIPIK, KARANGPOH, GRESIK REGENCY, EAST JAVA, INDONESIA",
    "P. T. PABRIK KERTAS TJIWI KIMIA": "KRAMAT, KRAMAT TEMENGGUNG, SIDOARJO REGENCY, EAST JAVA, INDONESIA",
    "POLEREJO, PURWOSARI": "POLEREJO, PURWOSARI, KEC. PURWOSARI, PASURUAN, JAWA TIMUR",
    "PT. AJINOMOTO INDONESIA": "GEDONG, MLIRIP, MOJOKERTO, EAST JAVA, INDONESIA",
    "PT. SOFTEX INDONESIA": "RANGKAH KIDUL, SIDOARJO REGENCY, EAST JAVA, INDONESIA",
    "PT. WINGS SURYA": "DUSUN WATES, CANGKIR, DRIYOREJO, GRESIK REGENCY, EAST JAVA, INDONESIA",
    "RANGKAH KIDUL, SIDOARJO SUB-DISTRCIT": "RANGKAH KIDUL, KEC. SIDOARJO, KABUPATEN SIDOARJO, JAWA TIMUR",
    # Tambahan umum jika perlu
    "PABRIK KERTAS TJIWI": "KRAMAT, KRAMAT TEMENGGUNG, SIDOARJO REGENCY, EAST JAVA, INDONESIA",
}

# ==============================================================================
# FUNGSI BANTUAN
# ==============================================================================
geolocator = Nominatim(user_agent="spil_project_manual_fix_v19", timeout=20)
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5)

def clean_address_string(addr):
    return str(addr).upper().replace("\n", ", ").replace(";", ", ").replace("  ", " ").strip()

def get_coordinates_smart(address):
    if pd.isna(address) or str(address).strip() in ["-", "", "nan"]:
        return None, None, "Alamat Kosong"

    raw_addr = clean_address_string(address)

    # --- CEK MANUAL FIX DULU ---
    # Jika alamat mengandung kata kunci salah, ganti dengan yang benar
    for wrong_key, correct_val in MANUAL_FIX.items():
        if wrong_key in raw_addr:
            raw_addr = correct_val # Override dengan alamat manual
            break # Stop loop, pakai ini

    # --- STRATEGI 1: EXACT ---
    try:
        loc = geocode(f"{raw_addr}, Jawa Timur")
        if loc: return loc.latitude, loc.longitude, "Akurat/Manual"
    except: pass 

    # --- STRATEGI 2: HAPUS GEDUNG ---
    keywords_trash = ["PT.", "CV.", "PABRIK", "MASJID", "GEREJA", "GUDANG", "DEPO", "TOKO"]
    clean_name = raw_addr
    for k in keywords_trash:
        if k in clean_name:
            parts = clean_name.split(',')
            if len(parts) > 1:
                clean_name = ", ".join(parts[1:])
            break  
    if clean_name != raw_addr:
        try:
            loc = geocode(f"{clean_name}, Jawa Timur")
            if loc: return loc.latitude, loc.longitude, "Estimasi (Tanpa Gedung)"
        except: pass

    # --- STRATEGI 3: SPLIT WILAYAH ---
    parts = raw_addr.split(',')
    if len(parts) > 1:
        try:
            loc = geocode(", ".join(parts[1:]) + ", Jawa Timur")
            if loc: return loc.latitude, loc.longitude, "Estimasi (Wilayah)"
        except: pass

    # --- STRATEGI 4: KOTA/KAB ---
    try:
        keywords = [p for p in parts if "SURABAYA" in p or "SIDOARJO" in p or "GRESIK" in p or "MOJOKERTO" in p or "PASURUAN" in p or "LAMONGAN" in p or "MALANG" in p or "TUBAN" in p]
        query_4 = f"{keywords[0]}, Jawa Timur" if keywords else f"{parts[-1]}, Jawa Timur"
        loc = geocode(query_4)
        if loc: return loc.latitude, loc.longitude, "General (Kota/Kab)"
    except: pass

    return None, None, "Gagal Geocoding"

def get_osrm_distance(lat1, lon1, lat2, lon2):
    if not lat1 or not lat2: return 0
    url = f"{OSRM_URL}/{lon1},{lat1};{lon2},{lat2}?overview=false"
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200 and r.json()['code'] == 'Ok':
            return r.json()['routes'][0]['distance'] / 1000
        return 0
    except: return 0

def get_weight_info(size_cont_str):
    s = str(size_cont_str).upper().strip()
    if "40" in s: return 32.0, 3.8
    elif "COMBO" in s or "2X20" in s: return 54.0, 4.6
    else: return 27.0, 2.3

def read_all_sheets_and_normalize(file_path, file_type):
    print(f"   -> Membaca {file_type}...")
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None) 
        df_list = []
        for sheet_name, df in all_sheets.items():
            df.columns = [str(c).strip().upper() for c in df.columns]
            df['SOURCE_SHEET'] = sheet_name
            
            if file_type == 'DOORING':
                if 'NO. SOPT 1' in df.columns: df['NO_SOPT_FIX'] = df['NO. SOPT 1']
                elif 'NO SOPT' in df.columns: df['NO_SOPT_FIX'] = df['NO SOPT']
                else: continue
                
                if 'SIZE CONT' in df.columns: df['SIZE_CONT_FIX'] = df['SIZE CONT']
                else: df['SIZE_CONT_FIX'] = "20"
                
                # Simpan AREA START jika ada
                if 'AREA START' in df.columns: df['AREA_START_FIX'] = df['AREA START']
                else: df['AREA_START_FIX'] = "YON"

            elif file_type == 'TLP':
                if 'SOPT NO' in df.columns: df['SOPT_NO_REF'] = df['SOPT NO']
                elif 'SOPT_NO' in df.columns: df['SOPT_NO_REF'] = df['SOPT_NO']
                else: continue
                
                if 'DOORING_ADDRESS' in df.columns: df['ALAMAT_REF'] = df['DOORING_ADDRESS']
                elif 'PICKUP / DELIVERY ADDRESS' in df.columns: df['ALAMAT_REF'] = df['PICKUP / DELIVERY ADDRESS']
                else: df['ALAMAT_REF'] = None
            
            df_list.append(df)
        return pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    except Exception as e:
        print(f"Error reading {file_type}: {e}")
        return pd.DataFrame()

# ==============================================================================
# MAIN PROCESS
# ==============================================================================
if __name__ == "__main__":
    print("1. Membaca & Menstandarisasi Data...")
    df_dooring = read_all_sheets_and_normalize(FILE_DOORING, 'DOORING')
    df_tlp = read_all_sheets_and_normalize(FILE_TLP, 'TLP')
    
    if df_dooring.empty or df_tlp.empty:
        print("[CRITICAL] File gagal dibaca.")
    else:
        print("2. Menggabungkan Data...")
        df_ref = df_tlp[['SOPT_NO_REF', 'ALAMAT_REF']].drop_duplicates(subset=['SOPT_NO_REF'])
        df_merge = pd.merge(df_dooring, df_ref, left_on='NO_SOPT_FIX', right_on='SOPT_NO_REF', how='left')
        
        print("3. Memulai Perhitungan...")
        results_jarak = []
        results_alasan = []
        results_total_berat = []
        results_tonkm = []
        
        total_data = len(df_merge)
        
        for idx, row in df_merge.iterrows():
            alamat = row['ALAMAT_REF']
            size_cont = row['SIZE_CONT_FIX']
            
            dist_pp = 0
            alasan = ""
            
            if pd.isna(alamat):
                alasan = "SOPT Tidak Ditemukan"
            else:
                try:
                    lat_dest, lon_dest, status_geo = get_coordinates_smart(alamat)
                except:
                    lat_dest = None
                    alasan = "Error Koneksi"
                
                if lat_dest:
                    if geodesic((DEPO_LAT, DEPO_LON), (lat_dest, lon_dest)).km > 400:
                        alasan = "Salah Geocode (Kejauhan)"
                    else:
                        d_out = get_osrm_distance(DEPO_LAT, DEPO_LON, lat_dest, lon_dest)
                        d_in = get_osrm_distance(lat_dest, lon_dest, DEPO_LAT, DEPO_LON)
                        time.sleep(0.2)
                        
                        if d_out > 0 and d_in > 0:
                            dist_pp = d_out + d_in
                            alasan = f"Sukses ({status_geo})"
                        else:
                            alasan = f"Gagal Routing ({status_geo})"
                else:
                    if not alasan: alasan = status_geo

            # Perhitungan Berat & TonKm
            b_full, b_empty = get_weight_info(size_cont)
            total_berat = b_full + b_empty
            tonkm_val = (dist_pp/2 * b_full) + (dist_pp/2 * b_empty) if dist_pp > 0 else 0
                
            results_jarak.append(dist_pp)
            results_alasan.append(alasan)
            results_total_berat.append(total_berat)
            results_tonkm.append(tonkm_val)
            
            if (idx + 1) % 50 == 0:
                print(f"   -> {idx+1}/{total_data} | {dist_pp:.1f}km | {alasan}")

        # --- UPDATE KOLOM SESUAI REQUEST ---
        df_merge['Jarak_PP_Km'] = results_jarak
        df_merge['Total_Berat_Ton'] = results_total_berat
        df_merge['TonKm_Dooring'] = results_tonkm
        df_merge['Alasan'] = results_alasan
        
        # 1. Kolom Baru/Rename Value Pasti (YON)
        df_merge['AREA AKHIR'] = 'YON'
        df_merge['AREA AMBIL EMPTY'] = 'YON'  
        df_merge['AREA BONGKAR'] = 'YON'      
        
        # 2. AREA PABRIK diisi Alamat
        df_merge['AREA PABRIK'] = df_merge['ALAMAT_REF']
        
        # 3. AREA START
        if 'AREA_START_FIX' in df_merge.columns:
            df_merge['AREA START'] = df_merge['AREA_START_FIX'].fillna('YON')
        else:
            df_merge['AREA START'] = 'YON'

        # 4. Final Columns Selection
        cols_final = [
            'NO', 'BULAN', 'LAMBUNG', 'NOPOL', 'SIZE CONT', 'SOPT_NO', 
            'AREA START', 'AREA AMBIL EMPTY', 'AREA PABRIK', 'AREA BONGKAR', 'AREA AKHIR',
            'Jarak_PP_Km', 'Total_Berat_Ton', 'TonKm_Dooring', 'Alasan'
        ]
        
        df_merge['NO'] = df_merge.index + 1
        df_merge['SOPT_NO'] = df_merge['NO_SOPT_FIX']
        
        # Mapping kolom asli
        final_df = pd.DataFrame()
        for col in cols_final:
            if col == 'SIZE CONT': final_df[col] = df_merge['SIZE_CONT_FIX']
            elif col in df_merge.columns: final_df[col] = df_merge[col]
            else: final_df[col] = "-"
            
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSELESAI! Hasil disimpan di: {OUTPUT_FILE}")