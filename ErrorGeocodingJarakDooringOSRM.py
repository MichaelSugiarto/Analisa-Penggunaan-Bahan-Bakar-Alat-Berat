import pandas as pd
import requests
import time
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic

# ==============================================================================
# KONFIGURASI FILE
# ==============================================================================
# Gunakan file hasil terakhir yang masih ada error-nya
INPUT_FILE = "DOORING_WITH_DISTANCE.xlsx" 
OUTPUT_FILE = "DOORING_WITH_DISTANCE_REVISI.xlsx"

# OSRM ENGINE
OSRM_URL = "http://router.project-osrm.org/route/v1/driving"

# DEPO: Laksda M. Nasir
DEPO_LAT = -7.2145
DEPO_LON = 112.7238

# --- DAFTAR PERBAIKAN ALAMAT MANUAL (BARU) ---
# Format: "POTONGAN ALAMAT DI EXCEL": "ALAMAT BARU YANG BENAR"
MANUAL_FIX_UPDATE = {
    "JALAN RAYA DAENDELS": "TJ. PAKIS, KEMANTREN, KEC. PACIRAN, KABUPATEN LAMONGAN, JAWA TIMUR",
    "GARUDA FOOD JAYA PT": "DUSUN LARANGAN, KRIKILAN, KEC. DRIYOREJO, KABUPATEN GRESIK, JAWA TIMUR",
    "MASJID NURUL JANNAH": "KARANGPOH, NGIPIK, KEC. GRESIK, KABUPATEN GRESIK, JAWA TIMUR",
    "PT. WINGS SURYA": "DUSUN WATES, CANGKIR, KEC. DRIYOREJO, KABUPATEN GRESIK, JAWA TIMUR",
    "JL. RAYA CANGKRINGMALANG, TUREN": "TUREN, CANGKRINGMALANG, KEC. BEJI, PASURUAN, JAWA TIMUR",
    # Anda bisa menambahkan alamat lain di sini jika ada lagi
}

# ==============================================================================
# FUNGSI BANTUAN (SAMA SEPERTI SEBELUMNYA)
# ==============================================================================
geolocator = Nominatim(user_agent="spil_project_fixer_v1", timeout=10)
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5)

def get_coordinates_smart(address):
    if pd.isna(address): return None, None, "Alamat Kosong"
    
    # Coba cari persis
    try:
        loc = geocode(f"{address}, Jawa Timur")
        if loc: return loc.latitude, loc.longitude, "Akurat (Revisi Manual)"
    except: pass
    
    # Coba potong-potong jika masih gagal
    parts = str(address).split(',')
    if len(parts) > 1:
        try:
            loc = geocode(f"{parts[-2]}, {parts[-1]}, Jawa Timur")
            if loc: return loc.latitude, loc.longitude, "Estimasi (Kota/Kab)"
        except: pass
        
    return None, None, "Gagal Geocoding"

def get_osrm_distance(lat1, lon1, lat2, lon2):
    if not lat1 or not lat2: return 0
    url = f"{OSRM_URL}/{lon1},{lat1};{lon2},{lat2}?overview=false"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200 and r.json()['code'] == 'Ok':
            return r.json()['routes'][0]['distance'] / 1000
    except: pass
    return 0

def get_weight_info(size_cont_str):
    s = str(size_cont_str).upper().strip()
    if "40" in s: return 32.0, 3.8
    elif "COMBO" in s or "2X20" in s: return 54.0, 4.6
    else: return 27.0, 2.3

# ==============================================================================
# MAIN PROCESS (HANYA YANG GAGAL)
# ==============================================================================
if __name__ == "__main__":
    print(f"1. Membaca File: {INPUT_FILE}...")
    df = pd.read_excel(INPUT_FILE)
    
    # Filter: Cari baris yang Jarak-nya 0
    # (Kita pakai index agar bisa update langsung ke DataFrame asli)
    error_indices = df[df['Jarak_PP_Km'] == 0].index
    
    print(f"   -> Ditemukan {len(error_indices)} baris yang masih error/0 km.")
    print("2. Memulai Perbaikan (Hanya baris yang error)...")
    
    count_fixed = 0
    
    for idx in error_indices:
        alamat_lama = str(df.at[idx, 'AREA PABRIK']).upper()
        alamat_baru = alamat_lama # Default sama
        
        # Cek apakah alamat ini ada di daftar MANUAL_FIX_UPDATE
        found_fix = False
        for k, v in MANUAL_FIX_UPDATE.items():
            if k in alamat_lama:
                alamat_baru = v
                found_fix = True
                break
        
        # Jika ada perbaikan, atau kita mau coba geocode ulang (siapa tau tadi cuma error koneksi)
        if found_fix:
            print(f"   [Baris {idx+1}] Memperbaiki: {k}...")
            
            # 1. Update Alamat di DataFrame
            df.at[idx, 'AREA PABRIK'] = alamat_baru
            
            # 2. Geocoding Ulang
            lat, lon, status = get_coordinates_smart(alamat_baru)
            
            if lat:
                # 3. Hitung Jarak Ulang
                d_out = get_osrm_distance(DEPO_LAT, DEPO_LON, lat, lon)
                d_in = get_osrm_distance(lat, lon, DEPO_LAT, DEPO_LON)
                
                if d_out > 0 and d_in > 0:
                    dist_pp = d_out + d_in
                    
                    # 4. Hitung TonKm Ulang
                    size_cont = df.at[idx, 'SIZE CONT']
                    b_full, b_empty = get_weight_info(size_cont)
                    tonkm = (dist_pp/2 * b_full) + (dist_pp/2 * b_empty)
                    
                    # 5. SIMPAN KE DATAFRAME
                    df.at[idx, 'Jarak_PP_Km'] = dist_pp
                    df.at[idx, 'TonKm_Dooring'] = tonkm
                    df.at[idx, 'Alasan'] = f"Sukses Revisi ({status})"
                    
                    count_fixed += 1
                    print(f"      -> SUKSES! Jarak: {dist_pp:.1f} km")
                else:
                    df.at[idx, 'Alasan'] = "Gagal Routing (Revisi)"
                    print("      -> Gagal Routing OSRM")
            else:
                df.at[idx, 'Alasan'] = "Gagal Geocoding (Revisi)"
                print("      -> Gagal Geocoding")
                
            # Jeda agar aman
            time.sleep(1)

    print(f"\nSELESAI! Berhasil memperbaiki {count_fixed} data.")
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"File hasil revisi disimpan di: {OUTPUT_FILE}")