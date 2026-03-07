import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import re
import warnings

warnings.filterwarnings('ignore')

# --- CONFIGURATION ---
FILE_BBM = "BBM AAB.xlsx"
FILE_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_TRUCKING = "HasilTrucking.xlsx"
FILE_MASTER = "cost & bbm 2022 sd 2025 HP & Type.xlsx"
OUTPUT_EXCEL = "Forecast_Detail_Data.xlsx"

# --- MAPPINGS ---
MONTH_MAP_ID = {
    'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4, 'Mei': 5, 'Juni': 6,
    'Juli': 7, 'Agustus': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
}
MONTH_MAP_EN = {
    'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MEI': 5, 'JUN': 6,
    'JUL': 7, 'AGT': 8, 'SEP': 9, 'OKT': 10, 'NOV': 11
}

# --- HELPER FUNCTIONS ---
def clean_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    return re.sub(r'[^A-Z0-9]', '', name)

def get_smart_match(raw_name, master_keys):
    """Mencocokkan nama dengan aturan spesifik user."""
    raw_upper = str(raw_name).upper().strip()
    clean_raw = clean_name(raw_upper)
    
    # 1. Exact Match
    if clean_raw in master_keys: return clean_raw
    
    # 2. Hardcoded Rules
    if "FL RENTAL 01" in raw_upper: return clean_name("FL RENTAL 01 TIMIKA")
    if "TOBATI" in raw_upper and "KALMAR" in raw_upper: return clean_name("TOP LOADER KALMAR 35T/TOBATI")
    if "L 8477 UUC" in raw_upper: return clean_name("L 9902 UR / S75")
    if "WIND RIVER" in raw_upper: return clean_name("TOP LOADER BOSS")
    
    # 3. "EX." Logic (Ambil setelah EX.)
    if "EX." in raw_upper or "EX " in raw_upper:
        parts = re.split(r'EX\.|EX\s', raw_upper)
        if len(parts) > 1:
            candidate = parts[1].split(')')[0].strip()
            cand_clean = clean_name(candidate)
            # Cek di Master
            if cand_clean in master_keys: return cand_clean
            # Cek partial match
            for k in master_keys:
                if cand_clean in k: return k

    # 4. Bracket Logic (Ambil sebelum kurung)
    if "(" in raw_upper:
        candidate = raw_upper.split("(")[0].strip()
        cand_clean = clean_name(candidate)
        if cand_clean in master_keys: return cand_clean
    
    return None # Return None jika tidak ketemu di master

def infer_type_fallback(name):
    """Tebak jenis jika tidak ada di Master (Fallback)."""
    n = str(name).upper()
    if "TRONTON" in n: return "TRONTON"
    if re.search(r'L\s*\d+', n) or "TRAILER" in n: return "TRAILER"
    if "REACH" in n or "STACKER" in n or "SMV" in n or "KALMAR" in n: return "REACH STACKER"
    if "FORKLIFT" in n or "FL " in n: return "FORKLIFT"
    if "CRANE" in n: return "CRANE"
    if "SIDE" in n: return "SIDE LOADER"
    if "TOP" in n: return "TOP LOADER"
    return "OTHERS"

print("=== START FORECASTING (HM UNTUK SEMUA) ===\n")

# ==============================================================================
# 1. LOAD MASTER DATA (KAMUS TIPE)
# ==============================================================================
print("1. Loading Master Data...")
unit_type_map = {}
master_keys = set()

try:
    df_m = pd.read_excel(FILE_MASTER, sheet_name='Sheet2', header=1)
    col_n = next((c for c in df_m.columns if 'NAMA' in str(c).upper()), None)
    col_t = next((c for c in df_m.columns if any(x in str(c).upper() for x in ['ALAT','JENIS'])), None)
    
    if col_n and col_t:
        for _, r in df_m.iterrows():
            c_name = clean_name(r[col_n])
            raw_t = str(r[col_t]).upper()
            
            t = "OTHERS"
            if "TRONTON" in raw_t: t = "TRONTON"
            elif "TRAILER" in raw_t or "HEAD" in raw_t: t = "TRAILER"
            elif "REACH" in raw_t: t = "REACH STACKER"
            elif "FORKLIFT" in raw_t: t = "FORKLIFT"
            elif "CRANE" in raw_t: t = "CRANE"
            elif "SIDE" in raw_t: t = "SIDE LOADER"
            elif "TOP" in raw_t: t = "TOP LOADER"
            
            if c_name:
                unit_type_map[c_name] = t
                master_keys.add(c_name)
except Exception as e: print(f"   [ERROR] Master File: {e}")

# ==============================================================================
# 2. EXTRACT HM DARI BBM AAB (UNTUK TRUCKING & NON-TRUCKING)
# ==============================================================================
print("2. Extracting HM from BBM AAB (All Units)...")
hm_data = []

try:
    xls = pd.ExcelFile(FILE_BBM)
    for sheet in xls.sheet_names:
        m_key = next((m for m in MONTH_MAP_EN if m in sheet.upper()), None)
        if not m_key: continue
        m_num = MONTH_MAP_EN[m_key]
        
        df_head = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=20)
        # Cari baris HM (atau LITER untuk patokan)
        h_rows = df_head[df_head.astype(str).apply(lambda x: x.str.contains(r'^HM$', case=False)).any(axis=1)].index
        if len(h_rows) == 0: continue
        h_idx = h_rows[0]
        
        df_full = pd.read_excel(xls, sheet_name=sheet, header=None)
        
        # Cari Baris Nama Unit (Scan ke atas)
        row_equip = None
        for r in range(h_idx-1, -1, -1):
            if df_full.iloc[r].notna().sum() > 1:
                row_equip = df_full.iloc[r].ffill(); break
        if row_equip is None: continue
        
        df_vals = df_full.iloc[h_idx+1:]
        
        for c in range(df_full.shape[1]):
            val_head = str(df_full.iloc[h_idx, c]).strip().upper()
            if val_head == 'HM':
                u_name = str(row_equip.iloc[c]).strip()
                if "TOTAL" in u_name.upper() or "NAN" in u_name.upper(): continue
                
                # Resolve Name (Gunakan Smart Match agar cocok dengan file lain)
                u_clean = get_smart_match(u_name, master_keys)
                if not u_clean: u_clean = clean_name(u_name) # Fallback clean
                
                # Ambil Nilai HM
                raw_hm = pd.to_numeric(df_vals.iloc[:, c], errors='coerce').dropna()
                hm_val = 0
                if not raw_hm.empty:
                    if raw_hm.mean() > 500: hm_val = raw_hm.max() - raw_hm.min() # Kumulatif
                    else: hm_val = raw_hm.sum() # Harian
                
                if hm_val > 0:
                    hm_data.append({'Month_Num': m_num, 'Unit_Clean': u_clean, 'HM': hm_val})

except Exception as e: print(f"   [ERROR] HM Extraction: {e}")

df_hm = pd.DataFrame(hm_data)
if not df_hm.empty:
    # Agregat jika ada duplikat unit per bulan
    df_hm = df_hm.groupby(['Month_Num', 'Unit_Clean'])['HM'].sum().reset_index()
    print(f"   -> {len(df_hm)} HM records extracted.")
else:
    print("   -> No HM data found.")

# ==============================================================================
# 3. PREPARE TRUCKING DATA (OKT-NOV)
# ==============================================================================
print("3. Preparing Trucking Data...")
df_truck = pd.DataFrame()
try:
    df_truck = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    # Filter Okt-Nov
    df_truck = df_truck[df_truck['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
    df_truck['Month_Num'] = df_truck['Bulan'].map(MONTH_MAP_ID)
    df_truck['Category'] = 'TRUCKING'
    
    # Clean Name & Resolve Type
    df_truck['Unit_Clean'] = df_truck['Nama_Unit'].apply(lambda x: get_smart_match(x, master_keys) or clean_name(x))
    
    def resolve_truck_type(u_clean):
        if u_clean in unit_type_map: return unit_type_map[u_clean]
        return "TRAILER" # Default Trucking
        
    df_truck['Type'] = df_truck['Unit_Clean'].apply(resolve_truck_type)
    
    # Rename & Select Cols
    df_truck = df_truck.rename(columns={'Nama_Unit': 'Unit_Name', 'Total_TonKm': 'TonKm', 'Total_Ton': 'Ton'})
    
    # Merge HM
    if not df_hm.empty:
        df_truck = pd.merge(df_truck, df_hm, on=['Month_Num', 'Unit_Clean'], how='left')
        df_truck['HM'] = df_truck['HM'].fillna(0)
    else:
        df_truck['HM'] = 0
        
    # Standardize Cols
    df_truck = df_truck[['Category', 'Type', 'Unit_Name', 'Unit_Clean', 'Month_Num', 'LITER', 'TonKm', 'Ton', 'HM']]

except Exception as e: print(f"   [ERROR] Trucking Prep: {e}")

# ==============================================================================
# 4. PREPARE NON-TRUCKING DATA (JAN-NOV)
# ==============================================================================
print("4. Preparing Non-Trucking Data...")
df_nt = pd.DataFrame()
try:
    df_nt = pd.read_excel(FILE_NON_TRUCKING, sheet_name='Data_Bulanan')
    # Filter Jan-Nov
    valid_m = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November']
    df_nt = df_nt[df_nt['Bulan'].isin(valid_m)]
    df_nt['Month_Num'] = df_nt['Bulan'].map(MONTH_MAP_ID)
    df_nt['Category'] = 'NON-TRUCKING'
    
    # Clean Name & Type
    col_u = 'Unit_Name' if 'Unit_Name' in df_nt.columns else 'Nama_Unit'
    df_nt['Unit_Clean'] = df_nt[col_u].apply(lambda x: get_smart_match(x, master_keys) or clean_name(x))
    
    # Gunakan Jenis_Alat dari file NonTrucking jika ada, atau map dari master
    if 'Jenis_Alat' in df_nt.columns:
        df_nt['Type'] = df_nt['Jenis_Alat'].apply(lambda x: x.upper() if pd.notna(x) else "OTHERS")
        # Normalize specific names
        df_nt['Type'] = df_nt['Type'].replace({
            'REACH STACKER': 'REACH STACKER', 'SMV': 'REACH STACKER', 
            'FL': 'FORKLIFT', 'SL': 'SIDE LOADER'
        })
    else:
        df_nt['Type'] = df_nt['Unit_Clean'].apply(lambda x: unit_type_map.get(x, infer_type_fallback(x)))

    # Rename & Cols
    df_nt = df_nt.rename(columns={col_u: 'Unit_Name', 'Total_Ton': 'Ton'})
    
    # Merge HM
    if not df_hm.empty:
        df_nt = pd.merge(df_nt, df_hm, on=['Month_Num', 'Unit_Clean'], how='left')
        df_nt['HM'] = df_nt['HM'].fillna(0)
    else:
        df_nt['HM'] = 0
        
    df_nt['TonKm'] = 0 # Placeholder for merge compatibility
    df_nt = df_nt[['Category', 'Type', 'Unit_Name', 'Unit_Clean', 'Month_Num', 'LITER', 'TonKm', 'Ton', 'HM']]

except Exception as e: print(f"   [ERROR] Non-Trucking Prep: {e}")

# ==============================================================================
# 5. GABUNG SEMUA DATA & EXPORT SOURCE
# ==============================================================================
df_final_data = pd.concat([df_truck, df_nt], ignore_index=True)

# ==============================================================================
# 6. FORECAST ENGINE
# ==============================================================================
print("5. Running Forecasts...")
forecast_results = []

# Definisi Forecast Config
forecast_configs = [
    # TRUCKING (Okt-Nov) -> Predictors: TonKm & HM
    {'cat': 'TRUCKING', 'type': 'TRAILER', 'range': [10, 11], 'preds': ['TonKm', 'HM']},
    {'cat': 'TRUCKING', 'type': 'TRONTON', 'range': [10, 11], 'preds': ['TonKm', 'HM']},
    
    # NON-TRUCKING (Jan-Nov) -> Predictors: Ton & HM
    {'cat': 'NON-TRUCKING', 'type': 'REACH STACKER', 'range': range(1, 12), 'preds': ['Ton', 'HM']},
    {'cat': 'NON-TRUCKING', 'type': 'FORKLIFT', 'range': range(1, 12), 'preds': ['Ton', 'HM']},
    {'cat': 'NON-TRUCKING', 'type': 'CRANE', 'range': range(1, 12), 'preds': ['Ton', 'HM']},
    {'cat': 'NON-TRUCKING', 'type': 'SIDE LOADER', 'range': range(1, 12), 'preds': ['Ton', 'HM']},
    {'cat': 'NON-TRUCKING', 'type': 'TOP LOADER', 'range': range(1, 12), 'preds': ['Ton', 'HM']}
]

for cfg in forecast_configs:
    # Filter Data
    sub = df_final_data[
        (df_final_data['Category'] == cfg['cat']) & 
        (df_final_data['Type'].str.contains(cfg['type'], case=False)) & # Flexible matching
        (df_final_data['Month_Num'].isin(cfg['range']))
    ]
    
    # Aggregate per Month
    agg = sub.groupby('Month_Num')[cfg['preds'] + ['LITER']].sum().reset_index().sort_values('Month_Num')
    
    # Cek Validitas
    if len(agg) < 2:
        status = 'Data Kurang (<2 bln)'
        fc_liter = 0
        trend_info = "-"
        formula = "-"
    else:
        # 1. Forecast Predictors (X) -> Time Series
        dec_X = []
        trend_list = []
        tm = LinearRegression()
        
        for p in cfg['preds']:
            # Cek apakah predictor ini semua 0? (Misal HM trucking tidak ada)
            if agg[p].sum() == 0:
                dec_X.append(0)
                trend_list.append(f"{p}: No Data")
            else:
                tm.fit(agg[['Month_Num']], agg[p])
                proj = max(0, tm.predict([[12]])[0])
                dec_X.append(proj)
                t_dir = "Naik" if tm.coef_[0] > 0 else "Turun"
                trend_list.append(f"{p}: {t_dir} ({proj:,.0f})")
        
        # 2. Forecast Liter (Y) -> Regression
        rm = LinearRegression()
        rm.fit(agg[cfg['preds']], agg['LITER'])
        fc_liter = max(0, rm.predict([dec_X])[0])
        
        status = 'Sukses'
        trend_info = ", ".join(trend_list)
        
        # Formula String
        formula = f"{rm.intercept_:,.0f}"
        for i, p in enumerate(cfg['preds']):
            formula += f" + ({rm.coef_[i]:.2f} * {p})"

    forecast_results.append({
        'Category': cfg['cat'],
        'Type': cfg['type'],
        'Status': status,
        'Forecast_LITER': fc_liter,
        'Trend_Info': trend_info,
        'Rumus_Estimasi': formula
    })

# --- 7. EXPORT ---
print("6. Exporting Results...")
df_res = pd.DataFrame(forecast_results)

with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
    df_res.to_excel(writer, sheet_name='Forecast_Result', index=False)
    df_final_data.to_excel(writer, sheet_name='Data_Source_Detail', index=False)

print("\n=== FORECAST RESULT ===")
print(df_res[['Type', 'Status', 'Forecast_LITER']])
print(f"\nFile Detail tersimpan: {OUTPUT_EXCEL}")