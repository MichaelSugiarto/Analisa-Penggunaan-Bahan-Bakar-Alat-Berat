import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
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

print("=== START FORECASTING (MEMUNCULKAN BASE FORECAST LITER) ===\n")

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
# 2. BACA BBM AAB 
# ==============================================================================
print("2. Membaca BBM AAB...")
target_sheets = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']
hm_data_store = {}
liter_data_store = {}

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
                
                is_hm = (header_str == 'HM')
                is_liter = (header_str in ['LITER', 'KELUAR', 'PEMAKAIAN'])
                
                if is_hm or is_liter:
                    raw_u = str(unit_names_row[col]).strip().upper()
                    if raw_u == "" or "UNNAMED" in raw_u or "TOTAL" in raw_u: continue
                    if raw_u.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL')): continue
                    
                    vals = pd.to_numeric(df.iloc[3:, col], errors='coerce')
                    clean_key = clean_unit_name(raw_u)
                    
                    temp = pd.DataFrame({'Date': dates, 'Value': vals})
                    
                    temp.dropna(subset=['Date'], inplace=True)
                    date_str_col = temp['Date'].astype(str).str.upper()
                    mask_exclude = date_str_col.str.contains('TOTAL|AVG|AVERAGE|GRAND', na=False)
                    temp = temp[~mask_exclude]
                    
                    temp['Date'] = pd.to_datetime(temp['Date'], dayfirst=True, errors='coerce')
                    temp.dropna(subset=['Date'], inplace=True)
                    
                    if is_hm:
                        if clean_key not in hm_data_store: hm_data_store[clean_key] = []
                        hm_data_store[clean_key].append(temp)
                    else:
                        if clean_key not in liter_data_store: liter_data_store[clean_key] = []
                        liter_data_store[clean_key].append(temp)

print(f"   -> Terbaca {len(hm_data_store)} unit HM dan {len(liter_data_store)} unit LITER.")

# ==============================================================================
# 3. FUNGSI LOGIC 
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

def resolve_unit_type(unit_name, matched_bbm_key, master_map):
    clean_ops = clean_unit_name(unit_name)
    if "PA 8511 AH" in clean_ops: return "TRONTON"
    if "L 8568 UK" in clean_ops: return "TRONTON"

    utype = "OTHERS"
    if clean_ops in master_map:
        utype = normalize_type(master_map[clean_ops]['Jenis_Alat'])
    if utype == "OTHERS" and matched_bbm_key and matched_bbm_key in master_map:
        utype = normalize_type(master_map[matched_bbm_key]['Jenis_Alat'])
        
    if utype == "OTHERS":
        target_names = [unit_name, matched_bbm_key]
        for name in target_names:
            if name and ("EX." in name or " EX " in name):
                parts = re.split(r'EX\.| EX ', name)
                if len(parts) > 1:
                    old_name = clean_unit_name(parts[1].replace(')', '').strip())
                    if old_name in master_map:
                        t = normalize_type(master_map[old_name]['Jenis_Alat'])
                        if t != "OTHERS":
                            utype = t
                            break

    if utype != "OTHERS": return utype

    if "TRONTON" in clean_ops: return "TRONTON"
    if "TRAILER" in clean_ops or re.search(r'L\s*\d+', clean_ops): return "TRAILER"
    return "OTHERS"

def calculate_monthly_hm_for_unit(df_list):
    if not df_list: return pd.DataFrame()
    df = pd.concat(df_list, ignore_index=True).sort_values('Date')
    df['HM_Clean'] = df['Value'].replace(0, np.nan).ffill().fillna(0)
    df['Delta_HM'] = df['HM_Clean'].diff().fillna(0)
    df.loc[df['Delta_HM'] < 0, 'Delta_HM'] = 0
    df.loc[df['Delta_HM'] > 100, 'Delta_HM'] = 0
    df['Month_Num'] = df['Date'].dt.month
    return df.groupby('Month_Num')['Delta_HM'].sum().reset_index()

def calculate_monthly_liter_for_unit(df_list):
    if not df_list: return pd.DataFrame()
    df = pd.concat(df_list, ignore_index=True).sort_values('Date')
    df['Value'] = df['Value'].fillna(0)
    df['Month_Num'] = df['Date'].dt.month
    return df.groupby('Month_Num')['Value'].sum().reset_index()

# ==============================================================================
# 4. PREPARE DATA
# ==============================================================================
print("3. Mengolah Data (Trucking & Non-Trucking)...")
df_truck = pd.DataFrame()
if os.path.exists(FILE_TRUCKING):
    df_tr_raw = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    df_tr_raw = df_tr_raw[df_tr_raw['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
    tr_list = []
    month_map_rev = {'Oktober': 10, 'November': 11}
    bbm_keys = list(set(list(hm_data_store.keys()) + list(liter_data_store.keys())))
    
    for _, r in df_tr_raw.iterrows():
        raw_u = str(r['Nama_Unit'])
        m_key = find_matching_hm_unit(raw_u, bbm_keys)
        utype = resolve_unit_type(raw_u, m_key, master_data_map)
        
        if utype in ['TRAILER', 'TRONTON']:
            m_num = month_map_rev.get(r['Bulan'].strip())
            hm_val = 0
            liter_val = r['LITER'] if 'LITER' in df_tr_raw.columns and pd.notna(r['LITER']) else 0 
            
            if m_key:
                if m_key in hm_data_store:
                    res_hm = calculate_monthly_hm_for_unit(hm_data_store[m_key])
                    if not res_hm.empty:
                        row_hm = res_hm[res_hm['Month_Num'] == m_num]
                        if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                
                if m_key in liter_data_store:
                    res_lit = calculate_monthly_liter_for_unit(liter_data_store[m_key])
                    if not res_lit.empty:
                        row_lit = res_lit[res_lit['Month_Num'] == m_num]
                        if not row_lit.empty: 
                            calc_lit = row_lit['Value'].values[0]
                            if calc_lit > 0: liter_val = calc_lit
            
            tr_list.append({
                'Category': 'TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 'LITER': liter_val, 'Workload': r.get('Total_TonKm', 0),
                'Ton': 0, 'HM': hm_val
            })
    df_truck = pd.DataFrame(tr_list)

df_nt = pd.DataFrame()
if os.path.exists(FILE_NON_TRUCKING):
    df_nt_raw = pd.read_excel(FILE_NON_TRUCKING, sheet_name='Data_Bulanan')
    valid = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November']
    df_nt_raw = df_nt_raw[df_nt_raw['Bulan'].isin(valid)]
    m_map_id = {'Januari':1, 'Februari':2, 'Maret':3, 'April':4, 'Mei':5, 'Juni':6,
                'Juli':7, 'Agustus':8, 'September':9, 'Oktober':10, 'November':11}
    col_u = 'Unit_Name' if 'Unit_Name' in df_nt_raw.columns else 'Nama_Unit'
    nt_list = []
    bbm_keys = list(set(list(hm_data_store.keys()) + list(liter_data_store.keys())))
    
    for _, r in df_nt_raw.iterrows():
        raw_u = str(r[col_u])
        m_key = find_matching_hm_unit(raw_u, bbm_keys)
        utype = resolve_unit_type(raw_u, m_key, master_data_map)
        
        if utype == "OTHERS" and 'Jenis_Alat' in df_nt_raw.columns and pd.notna(r['Jenis_Alat']):
            utype = normalize_type(r['Jenis_Alat'])
             
        if utype in ['REACH STACKER','FORKLIFT','CRANE','SIDE LOADER','TOP LOADER']:
            m_num = m_map_id.get(r['Bulan'])
            hm_val = 0
            liter_val = r['LITER'] if 'LITER' in df_nt_raw.columns and pd.notna(r['LITER']) else 0
            
            if m_key:
                if m_key in hm_data_store:
                    res_hm = calculate_monthly_hm_for_unit(hm_data_store[m_key])
                    if not res_hm.empty:
                        row_hm = res_hm[res_hm['Month_Num'] == m_num]
                        if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
                        
                if m_key in liter_data_store:
                    res_lit = calculate_monthly_liter_for_unit(liter_data_store[m_key])
                    if not res_lit.empty:
                        row_lit = res_lit[res_lit['Month_Num'] == m_num]
                        if not row_lit.empty: 
                            calc_lit = row_lit['Value'].values[0]
                            if calc_lit > 0: liter_val = calc_lit
            
            nt_list.append({
                'Category': 'NON-TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 'LITER': liter_val, 'Ton': r.get('Total_Ton', 0),
                'Workload': 0, 'HM': hm_val
            })
    df_nt = pd.DataFrame(nt_list)

# ==============================================================================
# 5. FORECASTING ENGINE DENGAN MEMUNCULKAN BASE LITER
# ==============================================================================
print("4. Menjalankan Forecasting Hybrid & Hitung Akurasi...")
df_final = pd.concat([df_truck, df_nt], ignore_index=True)

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
    sub_full = df_final[
        (df_final['Category']==cfg['cat']) & 
        (df_final['Type']==cfg['type']) & 
        (df_final['Month_Num'].isin(cfg['range']))
    ]
    
    sub_analysis = sub_full[sub_full['Status_Analisa'] == 'Masuk Analisa']
    
    agg = sub_analysis.groupby('Month_Num').agg({
        'LITER': 'sum', 'Unit_Clean': 'nunique', **{p: 'sum' for p in cfg['preds']}
    }).reset_index()
    
    forecast_base = 0
    valid_model = False
    r2_model = 0.0 
    
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
        
        y_pred_train = rm.predict(X_train)
        if len(y_train) > 1 and np.var(y_train) > 0:
            r2_model = r2_score(y_train, y_pred_train)
        
        X_pred = [pred_activity[p] for p in cfg['preds']]
        forecast_base = max(0, rm.predict([X_pred])[0])
        valid_model = True

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
            
        ratio_variance = np.std(ratios) if len(ratios) > 1 else 0.0
            
        final_forecast = (forecast_base * correction_factor) if valid_model else 0
        
        note = "Data Kurang/Tidak Masuk Analisa"
        if valid_model:
            if ratios:
                trend_str = "Standard"
                if correction_factor > 1.05: trend_str = f"Boros"
                elif correction_factor < 0.95: trend_str = f"Irit"
                note = f"General -> {trend_str}"
            else:
                note = "Tidak ada history valid, pakai rata-rata general"
                
        error_margin_val = round(ratio_variance * 100, 2) if len(ratios) > 1 else None
        r2_val = round(r2_model, 2) if valid_model else None
        
        forecast_detail_list.append({
            'Category': cfg['cat'], 
            'Type': cfg['type'], 
            'Unit_Name': unit,
            'Base_Forecast_LITER': round(forecast_base, 2), # <--- KOLOM BARU YANG ANDA CARI
            'Correction_Factor': round(correction_factor, 2),
            'Forecast_LITER_Dec': round(final_forecast, 2), 
            'Akurasi_General_R2': r2_val,
            'Error_Fluktuasi_Unit_Pct': error_margin_val,
            'Note': note
        })

# ==============================================================================
# 6. EXPORT HASIL
# ==============================================================================
df_res_detail = pd.DataFrame(forecast_detail_list)
with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
    df_res_detail.to_excel(writer, sheet_name='Forecast_Per_Unit', index=False)
    df_final.to_excel(writer, sheet_name='Data_Source_Detail', index=False)

print("\n=== CONTOH HASIL FORECAST PER UNIT (DENGAN BASE LITER) ===")
print(df_res_detail.head(5))
print(f"\nFile Detail tersimpan: {OUTPUT_EXCEL}")