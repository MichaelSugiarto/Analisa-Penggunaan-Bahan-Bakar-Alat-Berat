import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import re
import os
import warnings

warnings.filterwarnings('ignore')

# ==============================================================================
# KONFIGURASI FILE
# ==============================================================================
FILE_BBM = "BBM AAB.xlsx"
FILE_MASTER = "cost & bbm 2022 sd 2025 HP & Type.xlsx"
FILE_TRUCKING = "HasilTrucking.xlsx"
FILE_NON_TRUCKING = "HasilNonTrucking.xlsx"
OUTPUT_EXCEL = "Forecast_Detail_Data.xlsx"

print("=== START FORECASTING (FIXED: LITER CALCULATION & TRONTON DETECTION) ===\n")

# ==============================================================================
# FUNGSI HELPER
# ==============================================================================
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    name = name.replace("FORKLIFT", "FORKLIF")
    name = re.sub(r'\s+', ' ', name)
    return name

def normalize_type(t_str):
    t = str(t_str).upper()
    if "TRONTON" in t: return "TRONTON"
    if "TRAILER" in t or "HEAD" in t: return "TRAILER"
    if "REACH" in t or "STACKER" in t or "SMV" in t: return "REACH STACKER"
    if "FORKLIFT" in t: return "FORKLIFT"
    if "CRANE" in t: return "CRANE"
    if "SIDE" in t: return "SIDE LOADER"
    if "TOP" in t: return "TOP LOADER"
    return "OTHERS"

# ==============================================================================
# 1. LOAD MASTER DATA
# ==============================================================================
print("1. Loading Master Data...")
master_data_map = {}
if os.path.exists(FILE_MASTER):
    try:
        df_map = pd.read_excel(FILE_MASTER, sheet_name='Sheet2', header=1)
        col_map = {}
        for c in df_map.columns:
            c_str = str(c).upper()
            if 'NAMA' in c_str and 'ALAT' in c_str: col_map[c] = 'Unit_Name'
            elif 'LOKASI' in c_str: col_map[c] = 'Lokasi'
            elif 'JENIS' in c_str: col_map[c] = 'Jenis_Alat'
        df_map.rename(columns=col_map, inplace=True)
        if 'Unit_Name' in df_map.columns:
            for _, row in df_map.iterrows():
                if pd.notna(row['Unit_Name']):
                    raw_n = str(row['Unit_Name']).strip()
                    clean_id = clean_unit_name(raw_n)
                    master_data_map[clean_id] = {
                        'Unit_Name': raw_n,
                        'Jenis_Alat': row.get('Jenis_Alat', 'OTHERS')
                    }
        print(f"   -> {len(master_data_map)} unit ter-load dari Master.")
    except Exception as e: print(f"   [ERROR] Load Master: {e}")

# ==============================================================================
# 2. BACA BBM AAB (MODIFIKASI: BACA LITER & FILTER TOTAL)
# ==============================================================================
print("2. Membaca BBM AAB (HM & LITER with Total Filter)...")
target_sheets = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']
# Dictionary untuk menyimpan: { 'UNIT_NAME': {'HM': [df...], 'LITER': [df...]} }
bbm_data_store = {} 

if os.path.exists(FILE_BBM):
    xls = pd.ExcelFile(FILE_BBM)
    for sheet in target_sheets:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            unit_names_row = df.iloc[0].ffill()
            headers = df.iloc[2]
            dates = df.iloc[3:, 0]
            
            for col in range(1, df.shape[1]):
                header_str = str(headers[col]).strip().upper()
                
                # Deteksi Kolom HM dan LITER (Termasuk variasi nama kolom)
                is_hm = (header_str == 'HM')
                is_liter = (header_str in ['LITER', 'KELUAR', 'PEMAKAIAN', 'SOLAR'])
                
                if is_hm or is_liter:
                    raw_u = str(unit_names_row[col]).strip().upper()
                    if raw_u == "" or "UNNAMED" in raw_u or "EQUIP NAME" in raw_u or "TOTAL" in raw_u: continue
                    if raw_u.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL')): continue
                    
                    vals = pd.to_numeric(df.iloc[3:, col], errors='coerce')
                    clean_key = clean_unit_name(raw_u)
                    
                    if clean_key not in bbm_data_store: 
                        bbm_data_store[clean_key] = {'HM': [], 'LITER': []}
                    
                    temp = pd.DataFrame({'Date': dates, 'Value': vals})
                    
                    # --- FILTER PENTING: BUANG BARIS TOTAL/AVG ---
                    # 1. Drop jika tanggal kosong
                    temp.dropna(subset=['Date'], inplace=True)
                    # 2. Cek apakah ada string "Total" atau "Avg" di kolom tanggal
                    mask_summary = temp['Date'].astype(str).str.contains(r'TOTAL|AVG|AVERAGE|GRAND', case=False, na=False)
                    temp = temp[~mask_summary]
                    # 3. Konversi ke Datetime (yang gagal/string aneh jadi NaT)
                    temp['Date'] = pd.to_datetime(temp['Date'], dayfirst=True, errors='coerce')
                    # 4. Drop NaT
                    temp.dropna(subset=['Date'], inplace=True)
                    # ---------------------------------------------
                    
                    if is_hm:
                        bbm_data_store[clean_key]['HM'].append(temp)
                    else:
                        bbm_data_store[clean_key]['LITER'].append(temp)

print(f"   -> Terbaca {len(bbm_data_store)} unit unik di BBM AAB (HM & Liter terpisah).")

# ==============================================================================
# 3. FUNGSI LOGIC (MATCHING & CALCULATION)
# ==============================================================================
def find_matching_hm_unit(ops_unit_name, bbm_keys):
    clean_ops = clean_unit_name(ops_unit_name)
    if "FL RENTAL 01" in clean_ops and "TIMIKA" not in clean_ops:
        t = clean_unit_name("FL RENTAL 01 TIMIKA")
        if t in bbm_keys: return t
    if "TOBATI" in clean_ops and "KALMAR" in clean_ops:
        for k in bbm_keys:
            if "TOBATI" in k and "KALMAR" in k: return k
    if "L 8477 UUC" in clean_ops:
        for k in bbm_keys:
            if "L 9902 UR" in k: return k
    if "BOSS" in clean_ops and "TOP" in clean_ops:
        for k in bbm_keys:
            if "WIND RIVER" in k: return k
    
    if clean_ops in bbm_keys: return clean_ops
    for k in bbm_keys:
        if len(clean_ops) > 3 and clean_ops in k: return k
    for k in bbm_keys:
        if "EX." in k or " EX " in k:
            p = re.split(r'EX\.| EX ', k)
            if len(p) > 1:
                ae = clean_unit_name(p[1].replace(')', '').strip())
                if ae == clean_ops or (len(ae)>3 and ae in clean_ops): return k
    for k in bbm_keys:
        if "(" in k:
            bb = clean_unit_name(k.split("(")[0].strip())
            if bb == clean_ops or (len(bb)>3 and bb in clean_ops): return k
    return None

def resolve_unit_type_fixed(raw_u, m_key, master_data_map):
    # FIX: TRONTON DETECTION
    raw_u_upper = str(raw_u).upper()
    if "PA 8511 AH" in raw_u_upper or "PA8511AH" in clean_unit_name(raw_u): return "TRONTON"
    if "L 8568 UK" in raw_u_upper or "L8568UK" in clean_unit_name(raw_u): return "TRONTON"

    if m_key and m_key in master_data_map:
        return normalize_type(master_data_map[m_key]['Jenis_Alat'])
    
    # Fallback Logic
    if "TRONTON" in raw_u_upper: return "TRONTON"
    if re.search(r'L\s*\d+', raw_u_upper) or "TRAILER" in raw_u_upper: return "TRAILER"
    if "REACH" in raw_u_upper or "STACKER" in raw_u_upper: return "REACH STACKER"
    if "FORKLIFT" in raw_u_upper: return "FORKLIFT"
    if "CRANE" in raw_u_upper: return "CRANE"
    if "SIDE" in raw_u_upper: return "SIDE LOADER"
    if "TOP" in raw_u_upper: return "TOP LOADER"
    
    return "OTHERS"

# Hitung HM (Delta)
def calculate_monthly_hm(df_list):
    if not df_list: return pd.DataFrame()
    df = pd.concat(df_list, ignore_index=True).sort_values('Date')
    df['HM_Clean'] = df['Value'].replace(0, np.nan).ffill().fillna(0)
    df['Delta_HM'] = df['HM_Clean'].diff().fillna(0)
    df.loc[df['Delta_HM'] < 0, 'Delta_HM'] = 0
    df.loc[df['Delta_HM'] > 100, 'Delta_HM'] = 0
    df['Month_Num'] = df['Date'].dt.month
    return df.groupby('Month_Num')['Delta_HM'].sum().reset_index()

# Hitung LITER (Sum Murni)
def calculate_monthly_liter(df_list):
    if not df_list: return pd.DataFrame()
    df = pd.concat(df_list, ignore_index=True).sort_values('Date')
    df['Value'] = df['Value'].fillna(0)
    df['Month_Num'] = df['Date'].dt.month
    return df.groupby('Month_Num')['Value'].sum().reset_index()

# ==============================================================================
# 4. PREPARE DATA
# ==============================================================================
print("3. Mengolah Data (Trucking & Non-Trucking) & Recalculate Liter...")

# --- TRUCKING ---
df_truck = pd.DataFrame()
if os.path.exists(FILE_TRUCKING):
    df_tr_raw = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    df_tr_raw = df_tr_raw[df_tr_raw['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
    tr_list = []
    month_map_rev = {'Oktober': 10, 'November': 11}
    bbm_keys = list(bbm_data_store.keys())
    
    for _, r in df_tr_raw.iterrows():
        raw_u = str(r['Nama_Unit'])
        m_key = find_matching_hm_unit(raw_u, bbm_keys)
        utype = resolve_unit_type_fixed(raw_u, m_key, master_data_map)
        
        if utype in ['TRAILER', 'TRONTON']:
            m_num = month_map_rev.get(r['Bulan'].strip())
            
            hm_val = 0
            liter_val = 0 # Default 0 (akan diisi dari Raw Data)
            
            if m_key:
                # Get HM
                if bbm_data_store[m_key]['HM']:
                    res_hm = calculate_monthly_hm(bbm_data_store[m_key]['HM'])
                    if not res_hm.empty:
                        row_hm = res_hm[res_hm['Month_Num'] == m_num]
                        if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                
                # Get LITER (RECALCULATED)
                if bbm_data_store[m_key]['LITER']:
                    res_lit = calculate_monthly_liter(bbm_data_store[m_key]['LITER'])
                    if not res_lit.empty:
                        row_lit = res_lit[res_lit['Month_Num'] == m_num]
                        if not row_lit.empty: liter_val = row_lit['Value'].values[0]
            
            # Jika Liter dari RAW tetap 0, bisa fallback ke report r['LITER']
            # Tapi user minta perbaiki perhitungan, jadi kita percaya RAW dulu.
            
            tr_list.append({
                'Category': 'TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 
                'LITER': liter_val, # Using Recalculated Value
                'Workload': r['Total_TonKm'],
                'Ton': 0, 'HM': hm_val
            })
    df_truck = pd.DataFrame(tr_list)

# --- NON-TRUCKING ---
df_nt = pd.DataFrame()
if os.path.exists(FILE_NON_TRUCKING):
    df_nt_raw = pd.read_excel(FILE_NON_TRUCKING, sheet_name='Data_Bulanan')
    valid = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November']
    df_nt_raw = df_nt_raw[df_nt_raw['Bulan'].isin(valid)]
    m_map_id = {'Januari':1, 'Februari':2, 'Maret':3, 'April':4, 'Mei':5, 'Juni':6,
                'Juli':7, 'Agustus':8, 'September':9, 'Oktober':10, 'November':11}
    col_u = 'Unit_Name' if 'Unit_Name' in df_nt_raw.columns else 'Nama_Unit'
    nt_list = []
    bbm_keys = list(bbm_data_store.keys())
    
    for _, r in df_nt_raw.iterrows():
        raw_u = str(r[col_u])
        m_key = find_matching_hm_unit(raw_u, bbm_keys)
        
        utype = "OTHERS"
        if 'Jenis_Alat' in df_nt_raw.columns and pd.notna(r['Jenis_Alat']):
            utype = normalize_type(r['Jenis_Alat'])
        elif m_key and m_key in master_data_map:
             utype = normalize_type(master_data_map[m_key]['Jenis_Alat'])
             
        if utype in ['REACH STACKER','FORKLIFT','CRANE','SIDE LOADER','TOP LOADER']:
            m_num = m_map_id.get(r['Bulan'])
            hm_val = 0
            liter_val = 0
            
            if m_key:
                # Get HM
                if bbm_data_store[m_key]['HM']:
                    res_hm = calculate_monthly_hm(bbm_data_store[m_key]['HM'])
                    if not res_hm.empty:
                        row_hm = res_hm[res_hm['Month_Num'] == m_num]
                        if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                
                # Get LITER (RECALCULATED)
                if bbm_data_store[m_key]['LITER']:
                    res_lit = calculate_monthly_liter(bbm_data_store[m_key]['LITER'])
                    if not res_lit.empty:
                        row_lit = res_lit[res_lit['Month_Num'] == m_num]
                        if not row_lit.empty: liter_val = row_lit['Value'].values[0]
            
            nt_list.append({
                'Category': 'NON-TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 
                'LITER': liter_val, # Using Recalculated Value
                'Ton': r['Total_Ton'],
                'Workload': 0, 'HM': hm_val
            })
    df_nt = pd.DataFrame(nt_list)

# ==============================================================================
# 5. FORECASTING ENGINE
# ==============================================================================
print("4. Menjalankan Forecasting Hybrid (Filter Active)...")
df_final = pd.concat([df_truck, df_nt], ignore_index=True)

# Tambahkan Status Analisa
df_final['Status_Analisa'] = 'Tidak Masuk Analisa'
condition_active = (df_final['LITER'] > 0) & ((df_final['Workload'] > 0) | (df_final['Ton'] > 0))
df_final.loc[condition_active, 'Status_Analisa'] = 'Masuk Analisa'

configs = [
    {'cat': 'TRUCKING', 'type': 'TRAILER', 'preds': ['Workload', 'HM'], 'range': [10, 11]},
    {'cat': 'TRUCKING', 'type': 'TRONTON', 'preds': ['Workload', 'HM'], 'range': [10, 11]},
    {'cat': 'NON-TRUCKING', 'type': 'REACH STACKER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'FORKLIFT', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'CRANE', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'SIDE LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'TOP LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}
]

forecast_detail_list = []

for cfg in configs:
    sub_full = df_final[(df_final['Category']==cfg['cat']) & (df_final['Type']==cfg['type']) & (df_final['Month_Num'].isin(cfg['range']))]
    
    # Hanya data valid untuk membangun Tren
    sub_analysis = sub_full[sub_full['Status_Analisa'] == 'Masuk Analisa']
    
    agg = sub_analysis.groupby('Month_Num').agg({
        'LITER': 'sum', 'Unit_Clean': 'nunique', **{p: 'sum' for p in cfg['preds']}
    }).reset_index()
    
    forecast_base = 0
    valid_model = False
    
    if len(agg) >= 2:
        for col in cfg['preds'] + ['LITER']:
            agg[f'{col}_Per_Unit'] = agg[col] / agg['Unit_Clean']
            
        pred_activity = {}
        tm = LinearRegression()
        for p in cfg['preds']:
            col_avg = f'{p}_Per_Unit'
            if agg[col_avg].sum() == 0: pred_activity[p] = 0
            else:
                tm.fit(agg[['Month_Num']], agg[col_avg])
                pred_activity[p] = max(0, tm.predict([[12]])[0])
        
        rm = LinearRegression()
        X_train = agg[[f'{p}_Per_Unit' for p in cfg['preds']]]
        y_train = agg['LITER_Per_Unit']
        rm.fit(X_train, y_train)
        
        X_pred = [pred_activity[p] for p in cfg['preds']]
        forecast_base = max(0, rm.predict([X_pred])[0])
        valid_model = True

    # Hitung per unit
    unique_units = sub_full['Unit_Clean'].unique()
    for unit in unique_units:
        unit_history = sub_full[sub_full['Unit_Clean'] == unit]
        
        ratios = []
        for _, row in unit_history.iterrows():
            if row['Status_Analisa'] == 'Masuk Analisa':
                avg_row = agg[agg['Month_Num'] == row['Month_Num']]
                if not avg_row.empty and avg_row['LITER_Per_Unit'].values[0] > 0:
                    ratios.append(row['LITER'] / avg_row['LITER_Per_Unit'].values[0])
        
        correction_factor = 1.0
        if ratios:
            correction_factor = np.mean(ratios)
            correction_factor = max(0.5, min(2.0, correction_factor))
            
        final_forecast = (forecast_base * correction_factor) if valid_model else 0
        
        note = "Data Kurang/Tidak Masuk Analisa"
        if valid_model and ratios:
            trend_str = "Standard"
            if correction_factor > 1.05: trend_str = f"Boros ({correction_factor:.2f}x)"
            elif correction_factor < 0.95: trend_str = f"Irit ({correction_factor:.2f}x)"
            note = f"General -> {trend_str}"
        elif valid_model and not ratios:
            note = "Tidak ada history valid, pakai rata-rata general"
        
        forecast_detail_list.append({
            'Category': cfg['cat'], 'Type': cfg['type'], 'Unit_Name': unit,
            'Correction_Factor': round(correction_factor, 2),
            'Forecast_LITER_Dec': round(final_forecast, 2), 'Note': note
        })

df_res_detail = pd.DataFrame(forecast_detail_list)
with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
    df_res_detail.to_excel(writer, sheet_name='Forecast_Per_Unit', index=False)
    df_final.to_excel(writer, sheet_name='Data_Source_Detail', index=False)

print("\n=== CONTOH HASIL FORECAST PER UNIT (HYBRID) ===")
print(df_res_detail.head(10))
print(f"\nFile Detail tersimpan: {OUTPUT_EXCEL}")