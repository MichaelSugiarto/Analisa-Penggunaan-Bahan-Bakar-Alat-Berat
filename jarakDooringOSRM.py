import pandas as pd
import requests
import json
import time
import re
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic

# ==============================================================================
# KONFIGURASI
# ==============================================================================
FILE_DOORING = "DOORING OKT-DES 2025 (Copy).xlsx" 
FILE_TLP = "dashboard TLP okt-des (S1L Trucking).xlsx"
OUTPUT_FILE = "DOORING_WITH_DISTANCE.xlsx"

# ENGINE: OSRM PUBLIC
OSRM_URL = "http://router.project-osrm.org/route/v1/driving"

# DEPO: Laksda M. Nasir (Sesuai Request Terakhir)
DEPO_LAT = -7.2145
DEPO_LON = 112.7238

# ==============================================================================
# FUNGSI BANTUAN
# ==============================================================================
geolocator = Nominatim(user_agent="spil_project_multisheet_v17", timeout=20)
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5)

def clean_address_string(addr):
    s = str(addr).upper().replace("\n", ", ").replace(";", ", ").replace("  ", " ").strip()
    return s

def get_coordinates_smart(address):
    """Geocoding Hybrid (Smart Split + Anti-Gedung)"""
    if pd.isna(address) or str(address).strip() in ["-", "", "nan"]:
        return None, None, "Alamat Kosong"

    raw_addr = clean_address_string(address)
    
    # 1. EXACT SEARCH
    try:
        loc = geocode(f"{raw_addr}, Jawa Timur")
        if loc: return loc.latitude, loc.longitude, "Akurat (Lengkap)"
    except: pass 

    # 2. HAPUS NAMA PT/GEDUNG
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

    # 3. POTONG BAGIAN DEPAN
    parts = raw_addr.split(',')
    if len(parts) > 1:
        try:
            loc = geocode(", ".join(parts[1:]) + ", Jawa Timur")
            if loc: return loc.latitude, loc.longitude, "Estimasi (Wilayah)"
        except: pass

    # 4. CARI KOTA/KABUPATEN
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
    """
    Return: (berat_full_ton, berat_empty_ton)
    """
    s = str(size_cont_str).upper().strip()
    # Logic 40ft
    if "40" in s: 
        return 32.0, 3.8
    # Logic Combo / 2x20
    elif "COMBO" in s or "2X20" in s: 
        return 54.0, 4.6
    # Logic 20ft (Default)
    else: 
        return 27.0, 2.3

def read_all_sheets_and_normalize(file_path, file_type):
    """
    Membaca semua sheet dan menstandarisasi nama kolom
    file_type: 'DOORING' atau 'TLP'
    """
    print(f"   -> Membaca {file_type}: {file_path}...")
    try:
        # Baca semua sheet
        all_sheets = pd.read_excel(file_path, sheet_name=None) 
        df_list = []
        
        for sheet_name, df in all_sheets.items():
            # Standardize column names (strip spaces, upper case)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # Tambahkan kolom penanda sheet (bulan)
            df['SOURCE_SHEET'] = sheet_name
            
            # --- LOGIKA NORMALISASI KOLOM ---
            if file_type == 'DOORING':
                # Target: 'NO_SOPT_FIX', 'SIZE_CONT_FIX'
                # Cari NO SOPT
                if 'NO. SOPT 1' in df.columns:
                    df['NO_SOPT_FIX'] = df['NO. SOPT 1']
                elif 'NO SOPT' in df.columns: # Handle Sheet DES
                    df['NO_SOPT_FIX'] = df['NO SOPT']
                else:
                    print(f"      [Warning] Sheet '{sheet_name}' tidak punya kolom SOPT. Skip.")
                    continue
                
                # Cari Size Cont
                if 'SIZE CONT' in df.columns:
                    df['SIZE_CONT_FIX'] = df['SIZE CONT']
                else:
                    df['SIZE_CONT_FIX'] = "20" # Default
                    
            elif file_type == 'TLP':
                # Target: 'SOPT_NO_REF', 'ALAMAT_REF'
                # Cari SOPT NO
                if 'SOPT NO' in df.columns:
                    df['SOPT_NO_REF'] = df['SOPT NO']
                elif 'SOPT_NO' in df.columns:
                    df['SOPT_NO_REF'] = df['SOPT_NO']
                else:
                    print(f"      [Warning] Sheet '{sheet_name}' tidak punya kolom SOPT NO. Skip.")
                    continue
                
                # Cari Alamat
                if 'DOORING_ADDRESS' in df.columns: # Sheet Okt
                    df['ALAMAT_REF'] = df['DOORING_ADDRESS']
                elif 'PICKUP / DELIVERY ADDRESS' in df.columns: # Sheet Nov, Des
                    df['ALAMAT_REF'] = df['PICKUP / DELIVERY ADDRESS']
                else:
                    df['ALAMAT_REF'] = None
            
            df_list.append(df)
            
        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            print(f"      -> Total data {file_type}: {len(combined_df)} baris (Gabungan semua sheet)")
            return combined_df
        else:
            return pd.DataFrame()
            
    except Exception as e:
        print(f"Error reading {file_type}: {e}")
        return pd.DataFrame()

# ==============================================================================
# MAIN PROCESS
# ==============================================================================
if __name__ == "__main__":
    print("1. Membaca & Menstandarisasi Data (Multi-Sheet)...")
    
    # BACA DOORING (Okt, Nov, Des)
    df_dooring = read_all_sheets_and_normalize(FILE_DOORING, 'DOORING')
    
    # BACA TLP (Okt, Nov, Des)
    df_tlp = read_all_sheets_and_normalize(FILE_TLP, 'TLP')
    
    if df_dooring.empty or df_tlp.empty:
        print("[CRITICAL] Salah satu file kosong atau gagal dibaca. Stop.")
    else:
        print("2. Menggabungkan Data (Merge SOPT)...")
        # Siapkan referensi TLP (hapus duplikat SOPT agar VLOOKUP aman)
        df_ref = df_tlp[['SOPT_NO_REF', 'ALAMAT_REF']].drop_duplicates(subset=['SOPT_NO_REF'])
        
        # Merge
        df_merge = pd.merge(df_dooring, df_ref, 
                            left_on='NO_SOPT_FIX', 
                            right_on='SOPT_NO_REF', 
                            how='left')
        
        print("3. Memulai Perhitungan (Hybrid Smart Geocoding + OSRM)...")
        
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
            
            # --- GEOCODING & ROUTING ---
            if pd.isna(alamat):
                alasan = "SOPT Tidak Ditemukan / Alamat Kosong di TLP"
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
                        
                        time.sleep(0.2) # Rate limit protection
                        
                        if d_out > 0 and d_in > 0:
                            dist_pp = d_out + d_in
                            alasan = f"Sukses ({status_geo})"
                        else:
                            alasan = f"Gagal Routing OSRM ({status_geo})"
                else:
                    if not alasan: alasan = status_geo

            # --- PERHITUNGAN BERAT & TONKM ---
            b_full, b_empty = get_weight_info(size_cont)
            total_berat_display = b_full + b_empty
            
            # Rumus: (Jarak/2 * Full) + (Jarak/2 * Empty)
            if dist_pp > 0:
                half_dist = dist_pp / 2
                tonkm_val = (half_dist * b_full) + (half_dist * b_empty)
            else:
                tonkm_val = 0
                
            results_jarak.append(dist_pp)
            results_alasan.append(alasan)
            results_total_berat.append(total_berat_display)
            results_tonkm.append(tonkm_val)
            
            if (idx + 1) % 50 == 0:
                print(f"   -> {idx+1}/{total_data} | {dist_pp:.1f}km | {alasan}")

        # --- OUTPUT ---
        df_merge['Jarak_PP_Km'] = results_jarak
        df_merge['Total_Berat_Ton'] = results_total_berat
        df_merge['TonKm_Dooring'] = results_tonkm
        df_merge['Alasan'] = results_alasan
        
        # Rapikan Kolom Output
        # Gunakan kolom asli jika ada, jika tidak pakai "-"
        # Prioritas kolom output: Sesuai file Excel Asli Dooring
        target_cols = [
            'NO', 'BULAN', 'LAMBUNG', 'NOPOL', 'SIZE CONT', 'NO_SOPT_FIX', # Use fixed SOPT
            'ALAMAT_REF', # Alamat hasil merge
            'AREA START', 'AREA AMBIL EMPTY 1', 'AREA PABRIK', 'AREA BONGKAR 1',
            'Jarak_PP_Km', 'Total_Berat_Ton', 'TonKm_Dooring', 'Alasan'
        ]
        
        # Buat kolom NO dummy
        df_merge['NO'] = df_merge.index + 1
        
        # Mapping nama kolom agar sesuai request user
        final_df = pd.DataFrame()
        
        for col in target_cols:
            if col == 'NO_SOPT_FIX':
                final_df['SOPT_NO'] = df_merge['NO_SOPT_FIX']
            elif col == 'ALAMAT_REF':
                final_df['DOORING_ADDRESS'] = df_merge['ALAMAT_REF']
            elif col == 'SIZE CONT':
                final_df['SIZE CONT'] = df_merge['SIZE_CONT_FIX']
            elif col in df_merge.columns:
                final_df[col] = df_merge[col]
            else:
                # Coba cari variasi nama kolom di data asli
                found = False
                for c_orig in df_merge.columns:
                    if col.replace(" 1", "") in c_orig: # Misal AREA AMBIL EMPTY 1 -> AREA AMBIL EMPTY
                        final_df[col] = df_merge[c_orig]
                        found = True
                        break
                if not found:
                    final_df[col] = "-"
                    
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSELESAI! Hasil disimpan di: {OUTPUT_FILE}")