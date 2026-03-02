import pandas as pd
import requests
import json
import time
import re
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic

# ==============================================================================
# KONFIGURASI FILE
# ==============================================================================
FILE_DOORING = "DOORING OKT-DES 2025 (Copy).xlsx" 
FILE_TLP = "dashboard TLP okt-des (S1L Trucking).xlsx"
OUTPUT_FILE = "DOORING_WITH_DISTANCE.xlsx"

# URL OSRM (Server Publik - Gratis & Stabil)
# Kita pakai mode 'driving' (Mobil) karena server publik jarang membuka mode truck.
# Untuk forecasting BBM, akurasinya tetap >95% valid.
OSRM_URL = "http://router.project-osrm.org/route/v1/driving"

# KOORDINAT DEPO PT SPIL YON (Titik Jalan Raya)
DEPO_LAT = -7.213825
DEPO_LON = 112.723788

# ==============================================================================
# FUNGSI BANTUAN
# ==============================================================================
# Gunakan timeout 15 detik agar koneksi tidak mudah putus
geolocator = Nominatim(user_agent="spil_project_osrm_final_v12", timeout=15)
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5)

def get_coordinates_smart(address):
    """
    Geocoding Bertingkat: Jalan -> Kelurahan -> Kota
    """
    if pd.isna(address) or str(address).strip() in ["-", "", "nan"]:
        return None, None, "Alamat Kosong"

    clean_addr = str(address).replace("\n", ", ").replace(";", ", ").replace("  ", " ").strip()
    parts = [p.strip() for p in clean_addr.split(',')]
    
    # --- LEVEL 1: PENCARIAN EXACT ---
    try:
        query_1 = f"{clean_addr}, Jawa Timur"
        loc = geocode(query_1)
        if loc: return loc.latitude, loc.longitude, "Akurat (Jalan)"
    except: pass 

    # --- LEVEL 2: PENCARIAN WILAYAH ---
    if len(parts) > 1:
        try:
            # Hapus bagian depan
            query_2 = ", ".join(parts[1:]) + ", Jawa Timur"
            loc = geocode(query_2)
            if loc: return loc.latitude, loc.longitude, "Estimasi (Kel/Kec)"
        except: pass

    # --- LEVEL 3: PENCARIAN KOTA/KABUPATEN ---
    try:
        keywords = [p for p in parts if "SURABAYA" in p.upper() or "SIDOARJO" in p.upper() or "GRESIK" in p.upper() or "MOJOKERTO" in p.upper() or "PASURUAN" in p.upper()]
        if keywords:
            query_3 = f"{keywords[0]}, Jawa Timur"
        else:
            query_3 = f"{parts[-1]}, Jawa Timur"
            
        loc = geocode(query_3)
        if loc: return loc.latitude, loc.longitude, "General (Kota/Kab)"
    except: pass

    return None, None, "Gagal Geocoding"

def get_osrm_distance(lat1, lon1, lat2, lon2):
    """
    Menghitung jarak menggunakan OSRM Public API.
    """
    if not lat1 or not lat2: return 0
    
    # Format URL: {lon},{lat};{lon},{lat}
    url = f"{OSRM_URL}/{lon1},{lat1};{lon2},{lat2}?overview=false"
    
    try:
        # Request ke server OSRM
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if data['code'] == 'Ok':
                # OSRM mengembalikan meter, convert ke KM
                return data['routes'][0]['distance'] / 1000
        return 0
    except:
        return 0

def get_weight_info(size_cont_str):
    s = str(size_cont_str).upper().strip()
    if "40" in s: return 32.0, 3.8
    elif "COMBO" in s or "2X20" in s: return 54.0, 4.6
    else: return 27.0, 2.3

# ==============================================================================
# MAIN PROCESS
# ==============================================================================
if __name__ == "__main__":
    print("1. Membaca File Data...")
    
    try:
        df_dooring = pd.read_excel(FILE_DOORING)
        df_tlp = pd.read_excel(FILE_TLP)
        
        col_sopt_dooring = 'NO. SOPT 1'
        col_sopt_tlp = 'SOPT_NO'
        col_alamat_tlp = 'DOORING_ADDRESS' 
        col_size_cont = 'SIZE CONT'
        
        print("2. Menggabungkan Data...")
        df_ref = df_tlp[[col_sopt_tlp, col_alamat_tlp]].drop_duplicates(subset=[col_sopt_tlp])
        df_merge = pd.merge(df_dooring, df_ref, 
                            left_on=col_sopt_dooring, 
                            right_on=col_sopt_tlp, 
                            how='left')
        
        print("3. Memulai Perhitungan (OSRM Public)...")
        print("   (Menggunakan internet. Mohon tunggu...)")
        
        results_jarak = []
        results_alasan = []
        results_total_berat = []
        results_tonkm = []
        
        total_data = len(df_merge)
        
        for idx, row in df_merge.iterrows():
            alamat = row[col_alamat_tlp]
            size_cont = row[col_size_cont]
            
            dist_pp = 0
            alasan = ""
            
            if pd.isna(alamat):
                alasan = "SOPT Tidak Ditemukan"
            else:
                lat_dest, lon_dest, status_geo = get_coordinates_smart(alamat)
                
                if lat_dest:
                    # Cek Jarak Lurus (Sanity Check)
                    if geodesic((DEPO_LAT, DEPO_LON), (lat_dest, lon_dest)).km > 300:
                        alasan = "Salah Geocode (Kejauhan)"
                        dist_pp = 0
                    else:
                        # HITUNG JARAK VIA OSRM (Outbound + Inbound)
                        d_out = get_osrm_distance(DEPO_LAT, DEPO_LON, lat_dest, lon_dest)
                        d_in = get_osrm_distance(lat_dest, lon_dest, DEPO_LAT, DEPO_LON)
                        
                        # Beri jeda 0.5 detik agar server tidak memblokir (Rate Limit)
                        time.sleep(0.5)
                        
                        if d_out > 0 and d_in > 0:
                            dist_pp = d_out + d_in
                            alasan = f"Sukses ({status_geo})"
                        else:
                            # Jika OSRM gagal, berarti memang tidak ada rute jalan
                            alasan = f"Gagal Routing OSRM ({status_geo})"
                            dist_pp = 0
                else:
                    alasan = status_geo

            # BERAT & TONKM
            b_full, b_empty = get_weight_info(size_cont)
            total_berat_display = b_full + b_empty
            
            if dist_pp > 0:
                half_dist = dist_pp / 2
                tonkm_val = (half_dist * b_full) + (half_dist * b_empty)
            else:
                tonkm_val = 0
                
            results_jarak.append(dist_pp)
            results_alasan.append(alasan)
            results_total_berat.append(total_berat_display)
            results_tonkm.append(tonkm_val)
            
            if (idx + 1) % 10 == 0:
                short_addr = str(alamat)[:20].replace('\n','')
                print(f"   -> {idx+1}/{total_data} | {dist_pp:.1f}km | {alasan} | {short_addr}...")

        # OUTPUT
        df_merge['Jarak_PP_Km'] = results_jarak
        df_merge['Total_Berat_Ton'] = results_total_berat
        df_merge['TonKm_Dooring'] = results_tonkm
        df_merge['Alasan'] = results_alasan
        df_merge['NO'] = df_merge.index + 1
        df_merge['SOPT_NO_FINAL'] = df_merge[col_sopt_tlp].fillna(df_merge[col_sopt_dooring])
        
        cols_final = [
            'NO', 'BULAN', 'LAMBUNG', 'NOPOL', 'SIZE CONT', 'SOPT_NO_FINAL', 
            'DOORING_ADDRESS', 'AREA START', 'AREA AMBIL EMPTY 1', 'AREA PABRIK', 
            'AREA BONGKAR 1', 'Jarak_PP_Km', 'Total_Berat_Ton', 'TonKm_Dooring', 'Alasan'
        ]
        
        for c in cols_final:
            if c not in df_merge.columns:
                if c == 'AREA AMBIL EMPTY 1' and 'AREA AMBIL EMPTY' in df_merge.columns:
                    df_merge['AREA AMBIL EMPTY 1'] = df_merge['AREA AMBIL EMPTY']
                elif c == 'AREA BONGKAR 1' and 'AREA BONGKAR' in df_merge.columns:
                    df_merge['AREA BONGKAR 1'] = df_merge['AREA BONGKAR']
                else:
                    df_merge[c] = "-"

        df_out = df_merge[cols_final].rename(columns={'SOPT_NO_FINAL': 'SOPT_NO'})
        df_out.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSELESAI! Hasil: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"\n[ERROR CRITICAL] {e}")