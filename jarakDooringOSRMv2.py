import pandas as pd
import requests
import time
import re
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# KONFIGURASI FILE & LOKASI
# ==============================================================================
FILE_DOORING = "DOORING OKT-DES 2025 (Copy).xlsx" 
FILE_TLP = "dashboard TLP okt-des (S1L Trucking).xlsx"
OUTPUT_FILE = "DOORING_FINAL_DISTANCE.xlsx"

OSRM_URL = "http://router.project-osrm.org/route/v1/driving"

# Titik Nol: Depo SPIL Yon Laksda M Nasir
DEPO_LAT = -7.2145
DEPO_LON = 112.7238

# ==============================================================================
# 1. KAMUS TRANSLASI ALAMAT 
# ==============================================================================
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

# ==============================================================================
# 2. BYPASS KOORDINAT MANUAL (Presisi Tinggi Anti U-Turn)
# ==============================================================================
EXACT_MATCH_COORDS = {
    "KOTA SBY, JAWA TIMUR": (-7.2504, 112.7688), 
    "SURABAYA, JAWA TIMUR": (-7.2504, 112.7688),
    "KABUPATEN MOJOKERTO, JAWA TIMUR": (-7.5458, 112.4939), 
    "MOJOKERTO REGENCY, EAST JAVA": (-7.5458, 112.4939),
    "MOJOKERTO, JAWA TIMUR": (-7.5458, 112.4939),
    "PASURUAN, JAWA TIMUR": (-7.6453, 112.8208),
    "KEC. GRESIK, KABUPATEN GRESIK, JAWA TIMUR": (-7.1610, 112.6515), 
    "KABUPATEN GRESIK, JAWA TIMUR": (-7.1610, 112.6515)
}

# Diurutkan dari yang paling SPESIFIK (Jalan/Pabrik) ke yang paling UMUM (Kawasan)
KEYWORD_COORDS = {
    "DAENDELS 64-65": (-6.8778, 112.3550),              # Paciran Lamongan (~134 km PP)
    "MASJID NURUL JANNAH": (-7.1585, 112.6465),         # Petrokimia Gresik (~43.4 km PP)
    "KEMIRI SEWU, PANDAAN": (-7.6440, 112.7160),        # Pandaan (~124 km PP)
    "DUA KELINCI": (-6.8126, 111.0818),                 # Pati Jateng (~625 km PP)
    "GARUDA FOOD JAYA": (-7.3685, 112.6322),            # Driyorejo (~74.3 km PP)
    "MOJOSARI-PACET KM. 6,5": (-7.5540, 112.5385),      # Kutorejo Mojokerto (~134 km PP)
    "NESTL INDONESIA - GEMPOL": (-7.5780, 112.6900),    # Gempol (~96.5 km PP)
    "RUNGKUT INDUSTRI I NO.16": (-7.3220, 112.7530),    # Tenggilis (~51 km PP)
    "RUNGKUT INDUSTRI RAYA NO.19": (-7.3320, 112.7660), # Rungkut Kidul (~51 km PP)
    "MADIUN - SURABAYA, BANJARAGUNG": (-7.4910, 112.4280), # Krian/Puri Mojokerto (~186 km PP)
    "MANYARSIDORUKUN": (-7.1350, 112.6250),             # Manyar Gresik (~66 km PP)
    "PANDANREJO, REJOSO": (-7.6600, 112.9550),          # Pasuruan (~166 km PP)
    "SOFTEX INDONESIA SIDOARJO": (-7.4565, 112.7355),   # Sidoarjo (~73 km PP)
    "RANGKAH KIDUL": (-7.4565, 112.7355),               # Sidoarjo (~73 km PP)
    "DUMAR INDUSTRI": (-7.2405, 112.6685),
    "RAYA DLANGGU NO.KM 19": (-7.5800, 112.5000),
    "SAWUNGGALING NO.24": (-7.3615, 112.6865),
    "BANJARAGUNG, KRIAN": (-7.4116, 112.5855),
    "BUMI MASPION": (-7.2033, 112.6481),
    "ROMOKALISARI I": (-7.2033, 112.6481),
    "POLEREJO, PURWOSARI": (-7.7475, 112.7335),
    "TJIWI KIMIA": (-7.4361, 112.4727),
    "JL. TJ. TEMBAGA": (-7.2140, 112.7300), 
    "JALAN NILAM TIMUR": (-7.2210, 112.7300), 
    "JL. KALIANGET": (-7.2280, 112.7320), 
    "DUPAK RUKUN": (-7.2410, 112.7150),
    "KALIANAK BARAT": (-7.2240, 112.6870),
    "MARGOMULYO III": (-7.2450, 112.6750),
    "PERGUDANGAN MARGOMULYO PERMAI": (-7.2450, 112.6750),
    "MARGOMULYO NO.65": (-7.2350, 112.6750),
    "MARGOMULYO NO.44": (-7.2350, 112.6750),
    "MARGOMULYO GG.SENTONG": (-7.2510, 112.6720),
    "JL. TANJUNGSARI": (-7.2600, 112.6850), 
    "JL. RAYA MASTRIP": (-7.3370, 112.6950), 
    "JL. PANJANG JIWO": (-7.3200, 112.7660), 
    "KALI RUNGKUT": (-7.3253, 112.7663), 
    "KENDANGSARI": (-7.3220, 112.7500), 
    "TENGGILIS MEJOYO": (-7.3220, 112.7500), 
    "RUNGKUT LOR": (-7.3200, 112.7660),
    "LINGKAR TIMUR": (-7.4560, 112.7350), 
    "WEDI, KEC. GEDANGAN": (-7.3911, 112.7369), 
    "JL. BERBEK INDUSTRI": (-7.3450, 112.7500), 
    "JL. RAYA BUDURAN": (-7.4200, 112.7200), 
    "JL. RAYA TROSOBO": (-7.3800, 112.6600), 
    "JL. TAMBAK SAWAH": (-7.3600, 112.7600), 
    "JL. RAYA MOJOKERTO SURABAYA": (-7.3850, 112.6700), 
    "JL.RAYA GILANG": (-7.3750, 112.6800), 
    "KARANGREJO, PASURUAN": (-7.6976, 112.8805), 
    "LATEK": (-7.6045, 112.8251), 
    "REJOSO": (-7.6534, 112.9360), 
    "KIG RAYA SELATAN": (-7.1700, 112.6450), 
    "GRESIK - BABAT": (-7.1180, 112.4250), 
    "DUSUN WATES, CANGKIR": (-7.3600, 112.6100),
    "GUBERNUR SURYO": (-7.1650, 112.6550),
    "RAYA KRIKILAN": (-7.3619, 112.6300),
    "PELEM WATU": (-7.3100, 112.5950),
    "RAYA ROOMO": (-7.1450, 112.6400),
    "RAYA SEMBAYAT": (-7.1100, 112.6050),
    "RAYA SUKOMULYO": (-7.1450, 112.6400),
    "BANYUTAMI": (-7.1000, 112.5800),
    "MASPION": (-7.1250, 112.6200),
    "RAYA UTARA, BLOK. M/1": (-7.1450, 112.6400),
    "KRAJAN, SUMENGKO": (-7.3750, 112.5350),
    "WINGS SURYA, DUSUN WATES": (-7.3600, 112.6100),
    "NGORO": (-7.5683, 112.6171),               
    "BARENG PROYEK": (-7.5950, 112.6950),
    "RAYA BEJI": (-7.5950, 112.7550),
    "CANGKRINGMALANG": (-7.5950, 112.7350),
    "RAYA GEMPOL": (-7.5850, 112.6950),
    "MALANG NO.KM 40, KECINCANG": (-7.6150, 112.6950),
    "WICAKSONO, BANGLE": (-7.6050, 112.7250),
    "KALIPUTIH, SUMBERSUKO": (-7.6050, 112.6850),
    "GROGOLAN, WINONG": (-7.5850, 112.6850),
}

def check_manual_bypass(raw_address):
    addr_clean = str(raw_address).upper().replace('\n', ' ')
    addr_clean = re.sub(r'(?i)\bINDONESIA\b', '', addr_clean)
    addr_clean = re.sub(r'^[^\w]+|[^\w]+$', '', addr_clean).strip()
    
    if addr_clean in EXACT_MATCH_COORDS:
        return EXACT_MATCH_COORDS[addr_clean], f"{addr_clean}"
        
    for keyword, coords in KEYWORD_COORDS.items():
        if keyword in addr_clean:
            return coords, f"{keyword}"
            
    return None, None

# ==============================================================================
# FUNGSI BANTUAN & PERHITUNGAN BERAT
# ==============================================================================
def get_osrm_distance(lat1, lon1, lat2, lon2):
    try:
        url = f"{OSRM_URL}/{lon1},{lat1};{lon2},{lat2}?overview=false"
        res = requests.get(url, timeout=10)
        data = res.json()
        if data.get('code') == 'Ok':
            return data['routes'][0]['distance'] / 1000 
    except Exception: pass
    return 0.0

def get_weight_info(size_cont):
    """
    Menghitung berat FULL dan EMPTY (dalam satuan TON).
    """
    size_str = str(size_cont).upper().strip()
    
    W_20_EMPTY = 2.3
    W_20_FULL = 27.0
    W_40_EMPTY = 3.8
    W_40_FULL = 32.0
    
    if 'COMBO' in size_str:
        return (W_20_FULL * 2), (W_20_EMPTY * 2)
        
    multiplier = 1
    if '2X' in size_str or '2 X' in size_str:
        multiplier = 2
    elif '1X' in size_str or '1 X' in size_str:
        multiplier = 1
        
    if '40' in size_str:
        return (W_40_FULL * multiplier), (W_40_EMPTY * multiplier)
    elif '20' in size_str:
        return (W_20_FULL * multiplier), (W_20_EMPTY * multiplier)
        
    return W_20_FULL, W_20_EMPTY

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

# ==============================================================================
# PROSES UTAMA
# ==============================================================================
print("1. Membaca dan Menyinkronkan File Excel...")

tlp_oct = pd.read_excel(FILE_TLP, sheet_name='oktober')
if 'DOORING_ADDRESS' in tlp_oct.columns: tlp_oct = tlp_oct.rename(columns={'DOORING_ADDRESS': 'ALAMAT'})
tlp_nov = pd.read_excel(FILE_TLP, sheet_name='november')
if 'SOPT NO' in tlp_nov.columns: tlp_nov = tlp_nov.rename(columns={'SOPT NO': 'SOPT_NO'})
if 'PICKUP / DELIVERY ADDRESS' in tlp_nov.columns: tlp_nov = tlp_nov.rename(columns={'PICKUP / DELIVERY ADDRESS': 'ALAMAT'})
tlp_des = pd.read_excel(FILE_TLP, sheet_name='desember')
if 'SOPT NO' in tlp_des.columns: tlp_des = tlp_des.rename(columns={'SOPT NO': 'SOPT_NO'})
if 'PICKUP / DELIVERY ADDRESS' in tlp_des.columns: tlp_des = tlp_des.rename(columns={'PICKUP / DELIVERY ADDRESS': 'ALAMAT'})

df_tlp_clean = pd.concat([tlp_oct[['SOPT_NO', 'ALAMAT']], tlp_nov[['SOPT_NO', 'ALAMAT']], tlp_des[['SOPT_NO', 'ALAMAT']]], ignore_index=True)
df_tlp_clean = df_tlp_clean.dropna(subset=['SOPT_NO', 'ALAMAT']).drop_duplicates(subset=['SOPT_NO'])

door_oct = pd.read_excel(FILE_DOORING, sheet_name='OKT').rename(columns={'NO. SOPT 1': 'SOPT_NO'})
door_nov = pd.read_excel(FILE_DOORING, sheet_name='NOV').rename(columns={'NO. SOPT 1': 'SOPT_NO'})
door_des = pd.read_excel(FILE_DOORING, sheet_name='DES').rename(columns={'NO SOPT': 'SOPT_NO'})

df_dooring = pd.concat([door_oct, door_nov, door_des], ignore_index=True)
df_merge = pd.merge(df_dooring, df_tlp_clean, on='SOPT_NO', how='left')

# ==============================================================================
# PROSES GEOCODING & OSRM
# ==============================================================================
unique_raw_addresses = df_merge['ALAMAT'].dropna().unique()
print(f"   -> Ditemukan {len(unique_raw_addresses)} alamat TLP unik.")
print("2. Memulai Geocoding & Routing OSRM...")

geolocator = Nominatim(user_agent="dooring_routing_trucking")
address_dict = {}
seen_coords = {}      
cached_distances = {} 

for i, raw_address in enumerate(unique_raw_addresses, 1):
    print(f"   [{i}/{len(unique_raw_addresses)}] Memproses: {str(raw_address)[:40]}...")
    
    address_to_geocode, is_revised = apply_manual_fixes(raw_address)
    
    dist_out, dist_in, dist_pp = 0.0, 0.0, 0.0
    status, lat, lon = "Gagal", None, None
    used_address = "-"
    
    bypass_coords, bypass_keyword = check_manual_bypass(address_to_geocode)
    
    if bypass_coords:
        lat, lon = bypass_coords
        status = f"Sukses (Bypass Manual)"
        used_address = bypass_keyword 
    else:
        queries_to_try = build_fallback_queries(address_to_geocode)
        if is_revised: queries_to_try.insert(0, address_to_geocode)
            
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
            except:
                status = "Timeout Nominatim"
            time.sleep(1)
            
        if not lat:
            status = "Alamat Tidak Ditemukan"
            
    if lat and lon:
        coord_key = (round(lat, 4), round(lon, 4))
        
        if coord_key in cached_distances:
            dist_pp = cached_distances[coord_key]
            if "Bypass" not in status:
                status = f"Sukses (Titik Kembar dgn {seen_coords[coord_key][:15]})"
        else:
            if "Bypass" not in status:
                seen_coords[coord_key] = str(used_address)
            
            geo_dist = geodesic((DEPO_LAT, DEPO_LON), (lat, lon)).kilometers
            d_out = get_osrm_distance(DEPO_LAT, DEPO_LON, lat, lon)
            d_in  = get_osrm_distance(lat, lon, DEPO_LAT, DEPO_LON)
            
            # Toleransi maksimal dinaikkan menjadi 700km sekali jalan
            if d_out > 0 and d_in > 0:
                if d_out > (geo_dist * 5) or d_out > 700: 
                    status = "Peringatan: Koordinat Nyasar Jauh"
                    dist_pp = 0
                else:
                    dist_pp = d_out + d_in
                    cached_distances[coord_key] = dist_pp 
            else:
                status = "Gagal Routing OSRM"
                    
    address_dict[raw_address] = {
        'Jarak_PP_Km': dist_pp, 'Status': status, 'Used_Address': used_address
    }

# ==============================================================================
# PENYUSUNAN LAPORAN AKHIR
# ==============================================================================
print("\n3. Menyusun Data ke Format Final...")
results_pp, results_tonkm, results_ton, results_alasan, results_used_addr = [], [], [], [], []

if 'SIZE CONT ' in df_merge.columns:
    df_merge['SIZE CONT'] = df_merge['SIZE CONT'].fillna(df_merge['SIZE CONT '])

for idx, row in df_merge.iterrows():
    addr = row['ALAMAT'] 
    size = row['SIZE CONT'] if 'SIZE CONT' in df_merge.columns else "20"
    
    # -------------------------------------------------------------------------
    # PERBAIKAN: BERAT SELALU DIHITUNG WALAU ALAMAT TIDAK ADA/TIDAK KETEMU
    # -------------------------------------------------------------------------
    b_full, b_empty = get_weight_info(size)
    total_berat_sekali_jalan = b_full + b_empty
    
    if pd.isna(addr):
        results_pp.append(0)
        results_ton.append(total_berat_sekali_jalan)
        results_tonkm.append(0)
        results_alasan.append("SOPT Tidak Ada Alamat")
        results_used_addr.append("-")
        continue
        
    info = address_dict.get(addr, {})
    if not info:
        results_pp.append(0)
        results_ton.append(total_berat_sekali_jalan)
        results_tonkm.append(0)
        results_alasan.append("Alamat Tidak Ditemukan")
        results_used_addr.append("-")
        continue

    dist_pp = info['Jarak_PP_Km']
    
    # Hitung TonKm
    tonkm = ((dist_pp/2) * b_full) + ((dist_pp/2) * b_empty) if dist_pp > 0 else 0
        
    results_pp.append(dist_pp)
    results_ton.append(total_berat_sekali_jalan)
    results_tonkm.append(tonkm)
    results_alasan.append(info['Status'])
    results_used_addr.append(info['Used_Address'])

df_merge['Jarak_PP_Km'] = results_pp
df_merge['Total_Berat_Ton'] = results_ton
df_merge['TonKm_Dooring'] = results_tonkm
df_merge['Alasan'] = results_alasan
df_merge['ALAMAT YANG DIPAKAI'] = results_used_addr

df_merge['AREA START'] = 'YON'
df_merge['AREA AMBIL EMPTY'] = 'YON'
df_merge['AREA PABRIK'] = df_merge['ALAMAT'] 
df_merge['AREA BONGKAR'] = 'YON'
df_merge['AREA AKHIR'] = 'YON'

df_merge['LAMBUNG'] = df_merge['LAMBUNG'].fillna(df_merge['PLAT NUMBER'] if 'PLAT NUMBER' in df_merge.columns else "")
df_merge['NOPOL'] = df_merge['NOPOL'] if 'NOPOL' in df_merge.columns else df_merge['LAMBUNG']

cols_final = [
    'BULAN', 'LAMBUNG', 'NOPOL', 'SIZE CONT', 'SOPT_NO', 
    'AREA START', 'AREA AMBIL EMPTY', 'AREA PABRIK', 'ALAMAT YANG DIPAKAI', 'AREA BONGKAR', 'AREA AKHIR',
    'Jarak_PP_Km', 'Total_Berat_Ton', 'TonKm_Dooring', 'Alasan'
]
for col in cols_final:
    if col not in df_merge.columns: df_merge[col] = ""

df_final = df_merge[cols_final]
df_final.insert(0, 'NO', range(1, 1 + len(df_final)))

df_final.to_excel(OUTPUT_FILE, index=False)
print(f"SELESAI! {len(df_final)} baris data diproses. File tersimpan sebagai: {OUTPUT_FILE}")