import pandas as pd
import numpy as np
import re
import json
import time
import os
from openai import OpenAI
from sklearn.metrics import mean_squared_error
from sklearn.ensemble import HistGradientBoostingRegressor
import pmdarima as pm
import warnings

warnings.filterwarnings("ignore")

# LOAD MASTER & CRAWLING STANDAR PABRIK
def load_master_info():
    try:
        df = pd.read_excel('cost & bbm 2022 sd 2025 HP & Type.xlsx', sheet_name='Sheet2', header=1)
        df.columns = df.columns.str.strip()

        col_nama  = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)
        col_jenis = next((c for c in df.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_nama), None)
        col_type  = next((c for c in df.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
        col_cap   = next((c for c in df.columns if str(c).strip().upper() == 'CAP' or 'CAPAC' in str(c).upper()), None)
        col_hp    = next((c for c in df.columns if str(c).strip().upper() == 'HP' or 'HORSE' in str(c).upper()), None)

        for c, label in [(col_nama, 'NAMA ALAT BERAT'), (col_jenis, 'ALAT BERAT'),
                         (col_type, 'TYPE/MERK'), (col_cap, 'CAP'), (col_hp, 'HP')]:
            if c is None:
                print(f"[!] Kolom '{label}' tidak ditemukan di file master.")
                return set(), {}, {}

        master_names = set()
        unit_to_key  = {}
        key_to_specs = {}

        for _, row in df.iterrows():
            nama = str(row[col_nama]).strip()
            if pd.isna(row[col_nama]) or nama in ('nan', '-', ''):
                continue

            jenis  = str(row[col_jenis]).strip().upper() if pd.notna(row[col_jenis]) else 'Tidak Diketahui'
            type_m = str(row[col_type]).strip()          if pd.notna(row[col_type])  else 'Tidak Diketahui'
            cap    = str(row[col_cap]).strip()            if pd.notna(row[col_cap])   else 'Tidak Diketahui'
            hp     = str(row[col_hp]).strip()             if pd.notna(row[col_hp])    else 'Tidak Diketahui'

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
            master_names.add(nama)
            unit_to_key[nama] = composite_key

            if composite_key not in key_to_specs:
                key_to_specs[composite_key] = {
                    'jenis':     jenis,
                    'type_merk': type_m,
                    'cap':       cap,
                    'hp':        hp,
                }

        print(f"[MASTER] Berhasil membaca {len(master_names)} unit, "
              f"{len(key_to_specs)} kombinasi unik (jenis+type+cap+hp).")
        return master_names, unit_to_key, key_to_specs

    except Exception as e:
        print(f"\n[!] Gagal membaca file master: {e}\n")
        return set(), {}, {}


def crawl_standar_pabrik_openai(key_to_specs: dict, api_key: str) -> dict:
    client     = OpenAI(api_key=api_key)
    hasil      = {}
    total_keys = len(key_to_specs)

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

    print(f"\n[CRAWLING] Memulai estimasi standar pabrik untuk "
          f"{total_keys} kombinasi unik via OpenAI API...")

    for idx, (key, specs) in enumerate(key_to_specs.items()):
        jenis  = specs['jenis']
        type_m = specs['type_merk']
        cap    = specs['cap']
        hp     = specs['hp']

        jenis_upper = jenis.upper()
        is_wheeled  = any(k in jenis_upper for k in ['TRONTON', 'TRAILER', 'HEAD', 'TRUCK'])

        if is_wheeled:
            if cap != 'Tidak Diketahui':
                cap_display = f"{cap} feet (ukuran peti kemas/container)"
            else:
                cap_display = cap
            panduan_khusus = (
                "\nPERHATIAN KHUSUS — Kendaraan Angkut Beroda (Tronton/Trailer/Head/Truck):\n"
                f"- Kapasitas '{cap}' di atas adalah ukuran peti kemas dalam FEET (bukan ton beban).\n"
                "  Contoh: 20 feet = kontainer 20 kaki, 40 feet = kontainer 40 kaki.\n"
                "  Jangan menafsirkan angka kapasitas ini sebagai berat muatan dalam ton.\n"
                "- Satuan konsumsi yang diminta adalah Liter per Jam operasional mesin menyala (bukan L/km).\n"
                "- Konsumsi wajar kendaraan kelas ini saat operasional normal: 2–12 L/Jam.\n"
                "- HP tinggi (200–500 HP) pada kendaraan beroda TIDAK berarti konsumsi 15+ L/Jam;\n"
                "  engine truck beroperasi jauh di bawah kapasitas maksimum saat berjalan normal.\n"
                "- Nilai di atas 15 L/Jam untuk jenis kendaraan ini hampir pasti TIDAK WAJAR.\n"
                "- Gunakan referensi konsumsi BBM riil truck/trailer logistik, bukan mesin stasioner.\n"
            )
        else:
            cap_display   = f"{cap} ton"
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

                print(f"  [{idx+1:3d}/{total_keys}] ✅ {jenis}|{type_m}|{cap}|{hp} "
                      f"→ {hasil[key]['Standar Pabrik Konsumsi BBM Per Jam']} L/Jam "
                      f"(cap: {cap_lo}–{cap_hi})")
                berhasil = True
                time.sleep(1)
                break

            except json.JSONDecodeError as e:
                print(f"  [{idx+1:3d}/{total_keys}] ⚠️  Gagal parse JSON — {jenis}|{type_m}: {e}")
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
                    wait = 30 * (attempt + 1)
                    print(f"  [{idx+1:3d}/{total_keys}] ⏳ Rate limit, tunggu {wait}s "
                          f"(attempt {attempt+1}/{MAX_RETRIES})...")
                    time.sleep(wait)
                else:
                    print(f"  [{idx+1:3d}/{total_keys}] ❌ Error: {str(e)[:80]}")
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

    n_berhasil = sum(1 for v in hasil.values()
                     if v['Standar Pabrik Konsumsi BBM Per Jam'] is not None)
    print(f"[CRAWLING] Selesai. {n_berhasil}/{total_keys} kombinasi berhasil diestimasi.")
    return hasil


def load_standar_pabrik_dari_file(file_path: str) -> dict:
    try:
        df = pd.read_excel(file_path)

        rename_map = {
            'Standar_Konsumsi_L_per_Jam':     'Standar Pabrik Konsumsi BBM Per Jam',
            'Standar_Konsumsi_Min_L_per_Jam': 'Standar Pabrik Konsumsi BBM Min (L/Jam)',
            'Standar_Konsumsi_Max_L_per_Jam': 'Standar Pabrik Konsumsi BBM Max (L/Jam)',
            'Sumber_Data_Standar':            'Sumber Data Standar Pabrik',
            'Catatan_Standar':                'Catatan Standar Pabrik',
        }
        df.rename(columns=rename_map, inplace=True)

        hasil = {}
        for _, row in df.iterrows():
            key = str(row['Composite_Key'])
            hasil[key] = {
                'Standar Pabrik Konsumsi BBM Per Jam':     row.get('Standar Pabrik Konsumsi BBM Per Jam'),
                'Standar Pabrik Konsumsi BBM Min (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Min (L/Jam)'),
                'Standar Pabrik Konsumsi BBM Max (L/Jam)': row.get('Standar Pabrik Konsumsi BBM Max (L/Jam)'),
                'Sumber Data Standar Pabrik':              row.get('Sumber Data Standar Pabrik', '-'),
                'Catatan Standar Pabrik':                  row.get('Catatan Standar Pabrik', '-'),
            }
        print(f"[STANDAR PABRIK] Berhasil membaca {len(hasil)} kombinasi "
              f"dari '{file_path}' — API tidak dipanggil.")
        return hasil
    except Exception as e:
        print(f"[!] Gagal membaca file standar pabrik '{file_path}': {e}")
        return {}


# P1 & P2: EKSTRAKSI DAN DATA ENGINEERING
def load_and_melt_excel(file_obj):
    SHEET_VALID = {'JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN',
                   'JUL', 'AGT', 'SEP', 'OKT', 'NOV', 'DES'}
    # Normalisasi variasi nama sheet
    SHEET_NORMALIZE = {
        'SEPT': 'SEP', 'Sept': 'SEP', 'sept': 'SEP',
        'AGS':  'AGT', 'Ags':  'AGT',
        'OKT':  'OKT', 'OCT':  'OKT',
    }

    all_data = []
    xls = pd.ExcelFile(file_obj)

    for sheet_name in xls.sheet_names:
        normalized = SHEET_NORMALIZE.get(sheet_name, sheet_name).upper()
        if normalized not in SHEET_VALID:
            continue
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1, 2])
            if df.empty or len(df.columns) == 0:
                print(f"  -> Melewati sheet kosong: '{sheet_name}'")
                continue
            df = df.set_index(df.columns[0])
            df.index.name = 'TANGGAL'
            df.columns    = df.columns.droplevel(1)
            df_stacked    = df.stack(level=0).reset_index()
            df_stacked.rename(columns={'level_1': 'EQUIP NAME'}, inplace=True)
            all_data.append(df_stacked)
        except Exception as e:
            print(f"  -> Gagal memproses sheet '{sheet_name}': {e}")
            continue

    if not all_data:
        return pd.DataFrame()

    df_final = pd.concat(all_data, ignore_index=True)
    df_final['TANGGAL'] = pd.to_datetime(df_final['TANGGAL'], dayfirst=True, errors='coerce')
    df_final = df_final.dropna(subset=['TANGGAL'])
    return df_final


def prepare_data():
    FILE_TRAIN = [
        ('BBM AAB 2023.xlsx', 2023),
        ('BBM AAB 2024.xlsx', 2024),
    ]
    FILE_TEST = [
        ('BBM AAB 2025.xlsx', 2025),
    ]

    df_list = []
    for fname, tahun in FILE_TRAIN + FILE_TEST:
        if not os.path.exists(fname):
            print(f"[!] File tidak ditemukan: {fname} — dilewati.")
            continue
        print(f"Mengekstrak data dari: {fname}")
        df_tmp = load_and_melt_excel(fname)
        if df_tmp.empty:
            print(f"  -> Tidak ada data valid di {fname}")
            continue
        df_list.append(df_tmp)

    if not df_list:
        print("[ERROR] Tidak ada file data BBM yang berhasil dibaca.")
        return pd.DataFrame(), pd.DataFrame()

    df_all = pd.concat(df_list, ignore_index=True)
    df_all = df_all.sort_values(by=['EQUIP NAME', 'TANGGAL'])

    df_all['HM_Clean']   = pd.to_numeric(df_all['HM'], errors='coerce').replace(0, np.nan)
    df_all['HM_Clean']   = df_all.groupby('EQUIP NAME')['HM_Clean'].ffill().fillna(0)
    df_all['Delta_HM']   = df_all.groupby('EQUIP NAME')['HM_Clean'].diff().fillna(0)
    df_all.loc[df_all['Delta_HM'] < 0,   'Delta_HM'] = 0
    
    df_all.loc[df_all['Delta_HM'] > 744, 'Delta_HM'] = 0
    df_all['LITER_Clean'] = pd.to_numeric(df_all['LITER'], errors='coerce').fillna(0)

    df_all['TAHUN_BULAN'] = df_all['TANGGAL'].dt.to_period('M')
    agg_data = df_all.groupby(['EQUIP NAME', 'TAHUN_BULAN']).agg(
        {'Delta_HM': 'sum', 'LITER_Clean': 'sum'}
    ).reset_index()
    agg_data.rename(columns={'Delta_HM': 'HM', 'LITER_Clean': 'LITER'}, inplace=True)

    train_agg = agg_data[agg_data['TAHUN_BULAN'] <= '2024-12'].copy()
    test_agg  = agg_data[
        (agg_data['TAHUN_BULAN'] >= '2025-01') &
        (agg_data['TAHUN_BULAN'] <= '2025-12')
    ].copy()

    print(f"[DATA] Train: {len(train_agg)} baris | Test: {len(test_agg)} baris")
    print(f"[DATA] Unit unik di train: {train_agg['EQUIP NAME'].nunique()} | "
          f"di test: {test_agg['EQUIP NAME'].nunique()}")
    return train_agg, test_agg


# FUNGSI PREPROCESSING & MODELING
def preprocess_timeseries(series):
    df  = pd.DataFrame(series, columns=['HM'])
    p05 = df['HM'].quantile(0.05)
    p95 = df['HM'].quantile(0.95)
    df['HM_Capped']   = df['HM'].clip(lower=p05, upper=p95) if p95 > 0 else df['HM']
    df['HM_Smoothed'] = df['HM_Capped'].ewm(span=3, min_periods=1).mean()
    return df['HM_Smoothed']


def prepare_boosting_features(series, n_lags=3):
    df = pd.DataFrame(series.values, columns=['y'])
    for i in range(1, n_lags + 1):
        df[f'lag_{i}'] = df['y'].shift(i)
    df['rolling_mean_3'] = df['y'].shift(1).rolling(window=3, min_periods=1).mean()
    df['rolling_std_3']  = df['y'].shift(1).rolling(window=3, min_periods=1).std().fillna(0)
    df['trend']          = np.arange(len(df))
    df = df.dropna()
    feature_cols = [f'lag_{i}' for i in range(1, n_lags + 1)] + \
                   ['rolling_mean_3', 'rolling_std_3', 'trend']
    return df[feature_cols], df['y'], feature_cols


def predict_boosting(train_series, steps_ahead):
    n = len(train_series)
    if n >= 7:   n_lags = 3
    elif n >= 5: n_lags = 2
    elif n >= 4: n_lags = 1
    else:        return np.full(steps_ahead, max(0.0, float(train_series.mean())))

    X_train, y_train, feature_cols = prepare_boosting_features(train_series, n_lags=n_lags)
    if len(X_train) < 3:
        return np.full(steps_ahead, max(0.0, float(train_series.mean())))

    model = HistGradientBoostingRegressor(
        max_iter=200, learning_rate=0.05, max_depth=4, random_state=42
    ).fit(X_train, y_train)

    predictions, history = [], list(train_series.values)
    trend_offset = len(history)

    for step in range(steps_ahead):
        lag_vals  = [history[-(i)] for i in range(1, n_lags + 1)]
        window    = history[-3:] if len(history) >= 3 else history
        roll_mean = float(np.mean(window))
        roll_std  = float(np.std(window)) if len(window) > 1 else 0.0
        row       = lag_vals + [roll_mean, roll_std, trend_offset + step]
        pred      = max(0.0, float(model.predict(pd.DataFrame([row], columns=feature_cols))[0]))
        predictions.append(pred)
        history.append(pred)

    return np.array(predictions)


def hitung_mape_aman(actual, pred):
    actual, pred = np.array(actual), np.array(pred)
    mask = actual != 0
    if not np.any(mask): return 0.0
    return float(np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100)


def ensemble_tertimbang(p_arima, p_gb, aktual_h):
    rmse_a = float(np.sqrt(mean_squared_error(aktual_h, p_arima))) + 1e-6
    rmse_g = float(np.sqrt(mean_squared_error(aktual_h, p_gb)))   + 1e-6
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


# MAPPING NAMA ALAT BERAT KE MASTER
def get_mapped_unit_name(unit_name, master_names):
    """
    Logika mapping identik dengan appTNTv3.py (get_mapped_unit_name_fcst).
    Urutan: hardcoded → exact match → strip parenthesis → EX. pattern.
    """
    hardcoded = {
        "FL RENTAL 01":                "FL RENTAL 01 TIMIKA",
        "TOBATI (EX.FL KALMAR 32T)":   "TOP LOADER KALMAR 35T/TOBATI",
        "L 8477 UUC (EX.L 9902 UR)":   "L 9902 UR / S75",
        "WIND RIVER (EX.TL BOSS 42T)":  "TOP LOADER BOSS"
    }
    unit_name = str(unit_name).strip()
    if unit_name in hardcoded and hardcoded[unit_name] in master_names:
        return hardcoded[unit_name]
    if unit_name in master_names:
        return unit_name
    if " (" in unit_name:
        before_paren = unit_name.split(" (")[0].strip()
        if before_paren in master_names:
            return before_paren
    if "EX." in unit_name.upper():
        match_ex = re.search(r'EX\.([^\)]+)', unit_name.upper())
        if match_ex:
            after_ex = match_ex.group(1).strip()
            if after_ex in master_names:
                return after_ex
    return None


# P3: PIPELINE FORECASTING
def run_forecast_pipeline(mode: str = "forecast", openai_api_key: str = ""):
    FILE_STANDAR = 'Konsumsi BBM Standar Pabrik.xlsx'

    master_names, unit_to_key, key_to_specs = load_master_info()
    if not master_names:
        print("Proses dihentikan karena data master alat berat gagal dimuat.")
        return

    # MODE CRAWL: ambil standar pabrik dari OpenAI lalu simpan ke file
    if mode == "crawl":
        if not openai_api_key.strip():
            print("[!] Mode 'crawl' membutuhkan OPENAI_API_KEY.")
            return

        print("\n" + "=" * 60)
        print(" MODE: CRAWL — Mengambil standar pabrik via OpenAI API")
        print("=" * 60)

        standar_per_key = crawl_standar_pabrik_openai(key_to_specs, openai_api_key)

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
        pd.DataFrame(rows_standar).to_excel(FILE_STANDAR, index=False)
        print(f"\n[✅] Hasil crawling disimpan ke '{FILE_STANDAR}'")
        print(f"[✅] Ubah MODE = 'forecast' untuk menjalankan forecast")
        return

    # ------------------------------------------------------------------
    # MODE FORECAST: baca standar dari file lokal, jalankan pipeline
    # ------------------------------------------------------------------
    elif mode == "forecast":
        print("\n" + "=" * 60)
        print(" MODE: FORECAST — Membaca standar pabrik dari file lokal")
        print("=" * 60)
        if os.path.exists(FILE_STANDAR):
            standar_per_key = load_standar_pabrik_dari_file(FILE_STANDAR)
        else:
            print(f"[!] File '{FILE_STANDAR}' tidak ditemukan.")
            print(f"[!] Forecast tetap berjalan tapi kolom standar pabrik tidak akan terisi.")
            standar_per_key = {}
    else:
        print(f"[!] MODE tidak dikenal: '{mode}'. Gunakan 'crawl' atau 'forecast'.")
        return

    train_agg, test_agg = prepare_data()
    if train_agg.empty:
        print("[FATAL] Data train kosong, proses dihentikan.")
        return

    # Unit yang diproses: gabungan antara train dan test
    list_unit_raw = sorted(
        set(train_agg['EQUIP NAME'].unique()) | set(test_agg['EQUIP NAME'].unique())
    )

    results_combined    = []
    metrics_list        = []
    excluded_units_list = []
    all_actual_hm, all_pred_arima_hm = [], []
    all_pred_gb_hm, all_pred_ens_hm  = [], []
    total_valid_population = 0

    print(f"\n[AI ENGINE] Memulai pelatihan untuk {len(list_unit_raw)} unit kandidat...")

    for unit in list_unit_raw:
        mapped_name = None
        try:
            mapped_name = get_mapped_unit_name(unit, master_names)
            if not mapped_name:
                # Unit tidak ditemukan di master, lewati tanpa menambah ke excluded
                continue

            total_valid_population += 1

            df_u_train = train_agg[train_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').sort_index().copy()
            df_u_test  = test_agg[test_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').sort_index().copy()

            comp_key   = unit_to_key.get(mapped_name, '')
            sp         = standar_per_key.get(comp_key, {})
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

            # Validasi ketersediaan data
            if df_u_test.empty:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Tidak ada data aktual di periode uji (Jan-Des 2025).'
                }); continue

            if len(df_u_train) < 12:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': f'Data latih kurang dari 12 bulan (tersedia: {len(df_u_train)} bulan).'
                }); continue

            if len(df_u_test) < 12:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': f'Data uji kurang dari 12 bulan (tersedia: {len(df_u_test)} bulan).'
                }); continue

            # Cek per tahun dalam data train secara dinamis (identik dengan appTNTv3.py)
            train_years = df_u_train.index.map(lambda p: p.year).unique()
            skip_unit   = False
            for yr in train_years:
                df_yr = df_u_train[df_u_train.index.map(lambda p: p.year) == yr]
                if len(df_yr) == 12 and (df_yr['HM'].sum() == 0 or df_yr['LITER'].sum() == 0):
                    excluded_units_list.append({
                        'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                        'Alasan': f'HM/LITER = 0 selama 1 tahun penuh pada data train tahun {yr}.'
                    })
                    skip_unit = True
                    break
            if skip_unit:
                continue

            if df_u_test['HM'].sum() == 0 or df_u_test['LITER'].sum() == 0:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'HM/LITER = 0 selama 1 tahun penuh pada data test (2025).'
                }); continue

            # Kalkulasi rasio dan nilai aktual test
            true_ratio       = (float(df_u_train['LITER'].sum()) / float(df_u_train['HM'].sum())
                                 if df_u_train['HM'].sum() > 0 else 0.0)
            aktual_l         = df_u_test['LITER'].values.astype(float)
            aktual_h         = df_u_test['HM'].values.astype(float)
            steps            = len(df_u_test)
            total_hm_test    = float(aktual_h.sum())
            total_liter_test = float(aktual_l.sum())
            aktual_l_per_jam = total_liter_test / total_hm_test if total_hm_test > 0 else None
            deviasi_pct      = ((aktual_l_per_jam - std_min) / std_min * 100
                                 if std_min and std_min > 0 and aktual_l_per_jam else None)

            # Dua varian time series: utuh dan dipotong dari titik aktif pertama
            t_utuh = df_u_train['HM'].copy()
            try:
                first_idx = df_u_train[df_u_train['HM'] > 0].index[0]
                t_potong  = df_u_train.loc[first_idx:]['HM'].copy()
            except IndexError:
                t_potong = t_utuh

            best_arima, best_gb     = np.zeros(steps), np.zeros(steps)
            min_rmse, model_success = float('inf'), False

            for _, ds in [("Utuh", t_utuh), ("Potong", t_potong)]:
                if len(ds) < 6:
                    continue
                model_success = True
                ds_s = preprocess_timeseries(ds)

                try:
                    arima_m  = pm.auto_arima(ds_s, seasonal=False, max_d=1,
                                              suppress_warnings=True, error_action="ignore")
                    p_ar_raw = arima_m.predict(n_periods=steps)
                    baseline = float(ds_s.tail(6).mean())
                    p_arima  = np.maximum(0.0,
                                          np.clip(p_ar_raw, baseline * 0.1, baseline * 2.0)
                                          ).astype(float)
                except Exception:
                    p_arima = np.full(steps, max(0.0, float(ds_s.mean())))

                try:
                    p_gb = predict_boosting(ds_s, steps).astype(float)
                except Exception:
                    p_gb = np.full(steps, max(0.0, float(ds_s.mean())))

                p_ens_t, _, _ = ensemble_tertimbang(p_arima, p_gb, aktual_h)
                rmse_t = float(np.sqrt(mean_squared_error(aktual_h, p_ens_t)))
                if rmse_t < min_rmse:
                    min_rmse, best_arima, best_gb = rmse_t, p_arima.copy(), p_gb.copy()

            if not model_success:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Data historis valid setelah dipotong < 6 bulan.'
                }); continue

            best_ensemble, w_a, w_g = ensemble_tertimbang(best_arima, best_gb, aktual_h)
            best_ensemble = best_ensemble.astype(float)

            rmse_arima    = float(np.sqrt(mean_squared_error(aktual_h, best_arima)))
            mape_arima    = hitung_mape_aman(aktual_h, best_arima)
            rmse_gb       = float(np.sqrt(mean_squared_error(aktual_h, best_gb)))
            mape_gb       = hitung_mape_aman(aktual_h, best_gb)
            rmse_ensemble = float(np.sqrt(mean_squared_error(aktual_h, best_ensemble)))
            mape_ensemble = hitung_mape_aman(aktual_h, best_ensemble)

            mape_dict      = {'ARIMA': mape_arima, 'Gradient Boosting': mape_gb, 'Ensemble': mape_ensemble}
            rmse_dict      = {'ARIMA': rmse_arima, 'Gradient Boosting': rmse_gb, 'Ensemble': rmse_ensemble}
            best_name_unit = min(mape_dict, key=mape_dict.get)
            pred_terpilih  = {'ARIMA': best_arima, 'Gradient Boosting': best_gb,
                               'Ensemble': best_ensemble}[best_name_unit]
            rmse_terpilih  = rmse_dict[best_name_unit]

            metrics_list.append({
                'EQUIP NAME':                                        unit,
                'NAMA_MASTER_TERPETAKAN':                            mapped_name,
                'Jenis_Alat':                                        spec_jenis,
                'Type_Merk':                                         spec_type,
                'Capacity':                                          spec_cap,
                'Horse_Power':                                       spec_hp,
                'Model_Terpilih_Unit':                               best_name_unit,
                'MAPE_ARIMA (%)':                                    round(mape_arima, 2),
                'RMSE_ARIMA':                                        round(rmse_arima, 2),
                'MAPE_GB (%)':                                       round(mape_gb, 2),
                'RMSE_GB':                                           round(rmse_gb, 2),
                'MAPE_Ensemble (%)':                                 round(mape_ensemble, 2),
                'RMSE_Ensemble':                                     round(rmse_ensemble, 2),
                'MAPE_Terpilih (%)':                                 round(mape_dict[best_name_unit], 2),
                'RMSE_Terpilih':                                     round(rmse_terpilih, 2),
                'Bobot_ARIMA':                                       w_a,
                'Bobot_GB':                                          w_g,
                'Standar Pabrik Konsumsi BBM Per Jam':               std_per_jam,
                'Standar Pabrik Konsumsi BBM Min (L/Jam)':           std_min,
                'Standar Pabrik Konsumsi BBM Max (L/Jam)':           std_max,
                'Konsumsi BBM Aktual Per Jam':                       round(aktual_l_per_jam, 4) if aktual_l_per_jam else None,
                'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': round(deviasi_pct, 2) if deviasi_pct is not None else None,
                'Kategori Efisiensi':                                kategorikan_deviasi(deviasi_pct),
                'Sumber Data Standar Pabrik':                        std_sumber,
                'Catatan Standar Pabrik':                            std_catatan,
            })

            all_actual_hm.extend(aktual_h.tolist())
            all_pred_arima_hm.extend(best_arima.tolist())
            all_pred_gb_hm.extend(best_gb.tolist())
            all_pred_ens_hm.extend(best_ensemble.tolist())

            for i, period in enumerate(df_u_test.index):
                std_per_bulan = (round(float(pred_terpilih[i]) * std_min, 2)
                                 if std_min else None)
                results_combined.append({
                    'EQUIP NAME':                                        unit,
                    'NAMA_MASTER_TERPETAKAN':                            mapped_name,
                    'Jenis_Alat':                                        spec_jenis,
                    'Type_Merk':                                         spec_type,
                    'Capacity':                                          spec_cap,
                    'Horse_Power':                                       spec_hp,
                    'Bulan':                                             str(period),
                    'Aktual_HM':                                         round(float(aktual_h[i]), 2),
                    'Aktual_LITER':                                      round(float(aktual_l[i]), 2),
                    'Prediksi_HM_ARIMA':                                 round(float(best_arima[i]), 2),
                    'Prediksi_LITER_ARIMA':                              round(float(best_arima[i]) * true_ratio, 2),
                    'Prediksi_HM_GB':                                    round(float(best_gb[i]), 2),
                    'Prediksi_LITER_GB':                                 round(float(best_gb[i]) * true_ratio, 2),
                    'Prediksi_HM_Ensemble':                              round(float(best_ensemble[i]), 2),
                    'Prediksi_LITER_Ensemble':                           round(float(best_ensemble[i]) * true_ratio, 2),
                    'Model_Terpilih_Unit':                               best_name_unit,
                    'Prediksi_HM_Terpilih':                              round(float(pred_terpilih[i]), 2),
                    'Prediksi_LITER_Terpilih':                           round(float(pred_terpilih[i]) * true_ratio, 2),
                    'Standar Pabrik Minimum Konsumsi BBM Per Jam':       std_min,
                    'Standar Pabrik Minimum Konsumsi BBM Per Bulan':     std_per_bulan,
                    'Konsumsi BBM Aktual Per Jam':                       round(aktual_l_per_jam, 4) if aktual_l_per_jam else None,
                    'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)': round(deviasi_pct, 2) if deviasi_pct is not None else None,
                    'Kategori Efisiensi':                                kategorikan_deviasi(deviasi_pct),
                    'Sumber Data Standar Pabrik':                        std_sumber,
                    'MAPE_ARIMA (%)':                                    round(mape_arima, 2),
                    'RMSE_ARIMA':                                        round(rmse_arima, 2),
                    'MAPE_GB (%)':                                       round(mape_gb, 2),
                    'RMSE_GB':                                           round(rmse_gb, 2),
                    'MAPE_Ensemble (%)':                                 round(mape_ensemble, 2),
                    'RMSE_Ensemble':                                     round(rmse_ensemble, 2),
                    'MAPE_Terpilih (%)':                                 round(mape_dict[best_name_unit], 2),
                    'RMSE_Terpilih':                                     round(rmse_terpilih, 2),
                    'Bobot_ARIMA':                                       w_a,
                    'Bobot_GB':                                          w_g,
                })

            print(f"  [✅] {unit} → {best_name_unit} | MAPE: {mape_dict[best_name_unit]:.1f}%")

        except Exception as e:
            excluded_units_list.append({
                'EQUIP NAME':             unit,
                'NAMA_MASTER_TERPETAKAN': mapped_name if mapped_name else '-',
                'Alasan':                 f'Gagal diproses (Internal Error): {str(e)}'
            })
            continue

    # PEMBUATAN EXCEL OUTPUT
    df_combined = pd.DataFrame(results_combined)
    df_metrics  = pd.DataFrame(metrics_list)

    if df_combined.empty:
        print("\nGAGAL: Tidak ada unit valid yang berhasil diproses.")
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel('Unit_Dikecualikan_Error.xlsx', index=False)
        return

    df_komparasi = df_combined[[
        'EQUIP NAME', 'NAMA_MASTER_TERPETAKAN',
        'Jenis_Alat', 'Type_Merk', 'Capacity', 'Horse_Power',
        'Bulan', 'Aktual_HM', 'Aktual_LITER',
        'Prediksi_HM_ARIMA',    'Prediksi_LITER_ARIMA',
        'Prediksi_HM_GB',       'Prediksi_LITER_GB',
        'Prediksi_HM_Ensemble', 'Prediksi_LITER_Ensemble',
        'Model_Terpilih_Unit',
        'Prediksi_HM_Terpilih', 'Prediksi_LITER_Terpilih',
        'Standar Pabrik Minimum Konsumsi BBM Per Jam',
        'Standar Pabrik Minimum Konsumsi BBM Per Bulan',
        'Konsumsi BBM Aktual Per Jam',
        'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)',
        'Kategori Efisiensi',
        'Sumber Data Standar Pabrik',
        'MAPE_ARIMA (%)', 'RMSE_ARIMA',
        'MAPE_GB (%)',    'RMSE_GB',
        'MAPE_Ensemble (%)','RMSE_Ensemble',
        'MAPE_Terpilih (%)','RMSE_Terpilih',
        'Bobot_ARIMA', 'Bobot_GB',
    ]].copy()

    cols_perincian = [
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
    df_terpilih      = df_combined[cols_perincian].copy()
    df_mape_under_35 = df_terpilih[df_terpilih['MAPE_Terpilih (%)'] <  35].copy()
    df_mape_over_35  = df_terpilih[df_terpilih['MAPE_Terpilih (%)'] >= 35].copy()

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
    cols_ok_dev = [c for c in cols_deviasi if c in df_metrics.columns]
    df_deviasi  = df_metrics[cols_ok_dev].copy()
    if 'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)' in df_deviasi.columns:
        df_deviasi = df_deviasi.sort_values(
            'Selisih Konsumsi Aktual vs Standar Min Pabrik (%)', ascending=False
        )

    global_mape_arima    = hitung_mape_aman(all_actual_hm, all_pred_arima_hm)
    global_mape_gb       = hitung_mape_aman(all_actual_hm, all_pred_gb_hm)
    global_mape_ensemble = hitung_mape_aman(all_actual_hm, all_pred_ens_hm)

    with pd.ExcelWriter('Hasil_Forecast_Final.xlsx', engine='openpyxl') as writer:
        df_komparasi.to_excel(writer,     sheet_name='Komparasi_Model',       index=False)
        if not df_mape_under_35.empty:
            df_mape_under_35.to_excel(writer, sheet_name='Akurasi_Bagus_Under35', index=False)
        if not df_mape_over_35.empty:
            df_mape_over_35.to_excel(writer,  sheet_name='Akurasi_Rendah_Over35', index=False)
        df_metrics.to_excel(writer,       sheet_name='Metrik_Per_Unit',        index=False)
        if not df_deviasi.empty:
            df_deviasi.to_excel(writer,   sheet_name='Deviasi_Standar_Pabrik', index=False)
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel(
                writer, sheet_name='Unit_Dikecualikan', index=False)

    # ── Summary Terminal ──────────────────────────────────────────────
    total_under_35 = df_mape_under_35['EQUIP NAME'].nunique() if not df_mape_under_35.empty else 0
    total_over_35  = df_mape_over_35['EQUIP NAME'].nunique()  if not df_mape_over_35.empty  else 0
    total_excluded = len(excluded_units_list)
    pct_under = (total_under_35 / total_valid_population * 100) if total_valid_population else 0
    pct_over  = (total_over_35  / total_valid_population * 100) if total_valid_population else 0
    pct_excl  = (total_excluded / total_valid_population * 100) if total_valid_population else 0

    print("\n" + "=" * 65)
    print(" SUMMARY HASIL ANALISA OPERASIONAL ".center(65))
    print("=" * 65)
    print(f"Total Populasi Alat Berat : {total_valid_population} unit")
    print(f"\nGlobal MAPE (referensi):")
    print(f"  ARIMA    : {global_mape_arima:.2f}%")
    print(f"  GB       : {global_mape_gb:.2f}%")
    print(f"  Ensemble : {global_mape_ensemble:.2f}%")
    print(f"\nDistribusi Model Terpilih Per Unit:")
    if not df_metrics.empty:
        for model_name, count in df_metrics['Model_Terpilih_Unit'].value_counts().items():
            pct = count / total_valid_population * 100 if total_valid_population else 0
            print(f"  {model_name:<20}: {count} unit ({pct:.1f}%)")
    print("-" * 65)
    print(f"1. Berhasil Dimodelkan              : {total_valid_population} unit")
    print(f"2. Akurasi Bagus  (MAPE < 35%)      : {total_under_35} unit ({pct_under:.1f}%)")
    print(f"3. Akurasi Rendah (MAPE >= 35%)     : {total_over_35} unit ({pct_over:.1f}%)")
    print(f"4. Unit Diexclude (Data Tidak Cukup): {total_excluded} unit ({pct_excl:.1f}%)")

    df_dev_valid = df_deviasi.dropna(subset=['Selisih Konsumsi Aktual vs Standar Min Pabrik (%)']) \
                   if not df_deviasi.empty else pd.DataFrame()
    if not df_dev_valid.empty and 'Kategori Efisiensi' in df_dev_valid.columns:
        print(f"\nRingkasan Efisiensi vs Standar Pabrik:")
        for kat, cnt in df_dev_valid['Kategori Efisiensi'].value_counts().items():
            print(f"  {kat:<40}: {cnt} unit")

    print("=" * 65)
    print("Laporan berhasil disimpan ke 'Hasil_Forecast_Final.xlsx'")


if __name__ == "__main__":
    MODE = "forecast"   # "crawl" untuk ambil standar pabrik, "forecast" untuk prediksi

    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
    run_forecast_pipeline(mode=MODE, openai_api_key=OPENAI_API_KEY)