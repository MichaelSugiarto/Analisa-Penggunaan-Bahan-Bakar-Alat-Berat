import requests
import json
from geopy.geocoders import Nominatim

# KONFIGURASI
VALHALLA_URL = "http://localhost:8002/route"

# 1. Titik Depo Saat Ini (Jl. Laksda M. Nasir)
DEPO_CURRENT = (-7.213825, 112.723788)

# 2. Titik Alternatif (Pintu Tol Dupak / Jalan Tol) - Pasti bisa dilewati truk
DEPO_ALT_TOLL = (-7.240726, 112.722883) 

def test_routing(lat_start, lon_start, lat_end, lon_end, label, mode="truck"):
    print(f"\n--- TEST ROUTE: {label} (Mode: {mode}) ---")
    payload = {
        "locations": [{"lat": lat_start, "lon": lon_start}, {"lat": lat_end, "lon": lon_end}],
        "costing": mode,
        "units": "km"
    }
    try:
        r = requests.post(VALHALLA_URL, json=payload, timeout=5)
        if r.status_code == 200:
            print(f"✅ SUKSES! Jarak: {r.json()['trip']['summary']['length']} km")
        else:
            # INI YANG PENTING: Kita lihat pesan error aslinya
            print(f"❌ GAGAL. Status: {r.status_code}")
            print(f"   Pesan Error Valhalla: {r.text}")
    except Exception as e:
        print(f"⚠️ Error Koneksi: {e}")

# ==========================================
# MULAI DIAGNOSA
# ==========================================
print("=== DIAGNOSA VALHALLA & ALAMAT ===")

# TEST 1: Cek Kesehatan Valhalla (Depo ke Titik Dekat 500 meter)
# Jika ini gagal, berarti Titik Start Depo "Terkunci" di peta.
print("\n[1] Cek Titik Start Depo (Current)")
test_routing(DEPO_CURRENT[0], DEPO_CURRENT[1], -7.218000, 112.724000, "Depo ke Tetangga (500m)", "auto")

# TEST 2: Cek Alamat "JL. TANJUNGSARI" (Yang tadi error)
print("\n[2] Cek Geocoding Alamat Bermasalah")
geolocator = Nominatim(user_agent="debug_spil_diagnostic")
alamat_test = "JL. TANJUNGSARI NO.14, SURABAYA"
try:
    loc = geolocator.geocode(alamat_test)
    if loc:
        print(f"✅ Geocoding Sukses!")
        print(f"   Alamat di Peta: {loc.address}")
        print(f"   Koordinat: {loc.latitude}, {loc.longitude}")
        
        # TEST 3: Coba Routing ke Sana
        print("\n[3] Coba Routing ke Tanjungsari")
        # Pake Depo Current
        test_routing(DEPO_CURRENT[0], DEPO_CURRENT[1], loc.latitude, loc.longitude, "Depo Current -> Tanjungsari", "truck")
        # Pake Depo Alternatif (Tol)
        test_routing(DEPO_ALT_TOLL[0], DEPO_ALT_TOLL[1], loc.latitude, loc.longitude, "Depo Alternatif -> Tanjungsari", "truck")
    else:
        print("❌ Geocoding GAGAL untuk alamat tersebut.")
        print("   Tips: Coba ganti 'JL. TANJUNGSARI' jadi 'JL. TJ. SARI'")
except Exception as e:
    print(f"Error Geocoding: {e}")