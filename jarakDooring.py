import pandas as pd
import requests
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import time

# --- KONFIGURASI ---
FILE_DOORING = "DOORING OKT-DES 2025 (Copy).xlsx"
FILE_TLP = "dashboard TLP okt-des (S1L Trucking).xlsx"
OUTPUT_FILE = "DOORING_WITH_DISTANCE.xlsx"

# Koordinat Depo PT SPIL YON (Titik Start & Akhir Default)
# Pastikan koordinat ini akurat. Contoh ini di Surabaya.
DEPO_LAT = -7.2088  # Contoh Latitude
DEPO_LON = 112.7168 # Contoh Longitude

# --- FUNGSI GEOCODING (ALAMAT -> KOORDINAT) ---
geolocator = Nominatim(user_agent="trucking_analytics_app")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

def get_coordinates(address):
    try:
        if pd.isna(address) or address == "-":
            return None, None
        location = geocode(address)
        if location:
            return location.latitude, location.longitude
        return None, None
    except:
        return None, None

# --- FUNGSI VALHALLA (KOORDINAT -> JARAK) ---
# Asumsi Valhalla jalan di localhost port 8002
VALHALLA_URL = "http://localhost:8002/route"

def get_valhalla_distance(lat_start, lon_start, lat_end, lon_end):
    if None in [lat_start, lon_start, lat_end, lon_end]:
        return 0
    
    json_body = {
        "locations": [
            {"lat": lat_start, "lon": lon_start},
            {"lat": lat_end, "lon": lon_end},
            {"lat": lat_start, "lon": lon_start} # Kembali ke Depo
        ],
        "costing": "truck",
        "units": "km"
    }
    
    try:
        response = requests.post(VALHALLA_URL, json=json_body)
        data = response.json()
        # Mengambil total jarak dari trip summary
        return data['trip']['summary']['length']
    except:
        return 0

# --- PROSES UTAMA ---
def process_distances():
    print("Membaca Data...")
    # Load Dooring (Gabung Sheet OKT, NOV) - Abaikan DES sesuai instruksi (hanya ada BBM Okt-Nov)
    df_dooring = pd.concat([
        pd.read_excel(FILE_DOORING, sheet_name='OKT').assign(Bulan='Oktober'),
        pd.read_excel(FILE_DOORING, sheet_name='NOV').assign(Bulan='November')
    ])
    
    # Load TLP Reference
    df_tlp_okt = pd.read_excel(FILE_TLP, sheet_name='oktober')
    df_tlp_nov = pd.read_excel(FILE_TLP, sheet_name='november')
    
    # Standardisasi Nama Kolom TLP agar mudah di-merge
    # Okt: SOPT_NO, DOORING_ADDRESS
    # Nov: SOPT NO, PICKUP / DELIVERY ADDRESS
    df_tlp_okt = df_tlp_okt[['SOPT_NO', 'DOORING_ADDRESS']].rename(
        columns={'SOPT_NO': 'NO_SOPT', 'DOORING_ADDRESS': 'ALAMAT_PABRIK'}
    )
    df_tlp_nov = df_tlp_nov[['SOPT NO', 'PICKUP / DELIVERY ADDRESS']].rename(
        columns={'SOPT NO': 'NO_SOPT', 'PICKUP / DELIVERY ADDRESS': 'ALAMAT_PABRIK'}
    )
    df_ref = pd.concat([df_tlp_okt, df_tlp_nov]).drop_duplicates(subset=['NO_SOPT'])
    
    # Merge Data
    print("Menggabungkan Data SOPT...")
    df_merge = pd.merge(df_dooring, df_ref, left_on='NO. SOPT 1', right_on='NO_SOPT', how='left')
    
    # List SOPT Tidak Terdeteksi
    missing_sopt = df_merge[df_merge['ALAMAT_PABRIK'].isna()]['NO. SOPT 1'].unique()
    print(f"Jumlah SOPT Tidak Ditemukan Alamatnya: {len(missing_sopt)}")
    print("Contoh SOPT Missing:", missing_sopt[:5])
    
    # Hitung Jarak (Looping) - Note: Ini akan memakan waktu
    print("Memulai Perhitungan Jarak (Geocoding + Valhalla)...")
    distances = []
    
    for index, row in df_merge.iterrows():
        alamat = row['ALAMAT_PABRIK']
        # 1. Geocoding
        lat_pabrik, lon_pabrik = get_coordinates(alamat)
        
        # 2. Routing (Depo -> Pabrik -> Depo)
        km = get_valhalla_distance(DEPO_LAT, DEPO_LON, lat_pabrik, lon_pabrik)
        distances.append(km)
        
        if index % 10 == 0: print(f"Processed {index} rows...")

    df_merge['JARAK_KM_VALHALLA'] = distances
    
    # Simpan Hasil
    df_merge.to_excel(OUTPUT_FILE, index=False)
    print(f"Selesai! File disimpan di {OUTPUT_FILE}")

# Uncomment baris di bawah ini jika environment Python lokal sudah siap
# process_distances()