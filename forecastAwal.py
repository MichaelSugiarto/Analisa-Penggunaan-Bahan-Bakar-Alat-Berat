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

print("=== START FORECASTING (FINAL FIX: CONSTRAINED CAUSAL MODEL) ===\n")

# ==============================================================================
# FUNGSI HELPER 
# ==============================================================================
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    name = name.replace("FORKLIFT", "FORKLIF")
    name = re.sub(r'\s+', ' ', name)
    return name

def extract_lambung_code(raw_kpi):
    s = str(raw_kpi).upper().strip()
    match = re.search(r'\((.*?)\)', s)
    if match: return match.group(1).strip()
    parts = s.split()
    if len(parts) > 1:
        last_part = parts[-1]
        if len(last_part) < 6: return last_part
    return s.replace("-", "").strip()

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
print("1. Loading Master Data untuk standarisasi nama...")
master_data_map = {}
master_keys_set = set()
if os.path.exists(FILE_MASTER):
    try:
        df_map = pd.read_excel(FILE_MASTER, sheet_name='Sheet2', header=1)
        col_name = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
        col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
        df_map.dropna(subset=[col_name], inplace=True)
        
        for _, row in df_map.iterrows():
            u_name = str(row[col_name]).strip().upper()
            c_id = clean_unit_name(u_name)
            if c_id:
                master_data_map[c_id] = {
                    'Unit_Name': u_name,
                    'Jenis_Alat': str(row[col_jenis]).strip().upper() if pd.notna(row[col_jenis]) else "OTHERS"
                }
                master_keys_set.add(c_id)
    except Exception as e: print(f"   [ERROR] Load Master: {e}")

def get_master_match(raw_name):
    raw_name = str(raw_name).strip().upper()
    c_raw = clean_unit_name(raw_name)
    if c_raw in master_data_map: return c_raw
    if " (" in raw_name:
        b_paren = clean_unit_name(raw_name.split(" (")[0])
        if b_paren in master_data_map: return b_paren
    if "EX." in raw_name or "EX " in raw_name:
        after_ex = raw_name.split("EX.")[-1] if "EX." in raw_name else raw_name.split("EX ")[-1]
        c_after = clean_unit_name(after_ex.replace(")", ""))
        if c_after in master_data_map: return c_after
        for m_key in master_keys_set:
            if c_after != "" and c_after in m_key: return m_key
    return c_raw # Fallback ke clean name

# ==============================================================================
# 2. BACA BBM AAB (KHUSUS UNTUK MENGAMBIL HM SECARA SERIES)
# ==============================================================================
print("2. Membaca BBM AAB (Mengekstrak HM & Liter)...")
hm_data_store = {}
liter_data_store = {}

if os.path.exists(FILE_BBM):
    xls = pd.ExcelFile(FILE_BBM)
    for sheet_name in xls.sheet_names:
        sht_up = sheet_name.upper()
        # Ambil semua bulan (karena Non-Trucking butuh dari Jan-Nov)
        
        df_cek = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
        if df_cek.empty: continue
        a1_val = str(df_cek.iloc[0, 0]).strip().upper()
        if "EQUIP" not in a1_val and "NAMA" not in a1_val: continue 
            
        df_full = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        unit_names_row = df_full.iloc[0].ffill() 
        group_kpi_row = df_full.iloc[1].ffill()  
        headers = df_full.iloc[2]                
        data_rows = df_full.iloc[3:]             
        
        def is_valid_date_row(val):
            if pd.isna(val): return False
            if isinstance(val, (int, float)) and 1 <= val <= 31: return True
            val_str = str(val).strip()
            if re.match(r'^\d{4}-\d{2}-\d{2}', val_str): return True 
            match = re.match(r'^(\d{1,2})', val_str) 
            if match and 1 <= int(match.group(1)) <= 31: return True
            return False
            
        valid_mask = data_rows.iloc[:, 0].apply(is_valid_date_row)
        valid_data_rows = data_rows[valid_mask]
        
        # Ekstrak kolom tanggal untuk TimeSeries HM
        dates_raw = valid_data_rows.iloc[:, 0]
        
        # Berikan tahun dummy jika hanya angka tanggal agar bisa diparse sebagai datetime
        month_idx = {'JAN':1, 'FEB':2, 'MAR':3, 'APR':4, 'MEI':5, 'JUN':6, 'JUL':7, 'AGT':8, 'SEP':9, 'OKT':10, 'NOV':11}
        cur_month = next((v for k,v in month_idx.items() if k in sht_up), 1)
        
        dates_parsed = []
        for d in dates_raw:
            if isinstance(d, (int, float)): dates_parsed.append(pd.Timestamp(year=2025, month=cur_month, day=int(d)))
            else: dates_parsed.append(pd.to_datetime(d, dayfirst=True, errors='coerce'))
            
        dates_series = pd.Series(dates_parsed)
        
        for col in range(1, df_full.shape[1]):
            header_str = str(headers.iloc[col]).strip().upper()
            is_hm = (header_str == 'HM')
            is_liter = ('LITER' in header_str)
            
            if is_hm or is_liter:
                raw_equip = str(unit_names_row.iloc[col]).strip().upper()
                raw_kpi = str(group_kpi_row.iloc[col]).strip().upper()
                
                if "TOTAL" in raw_equip or "UNNAMED" in raw_equip or raw_equip == "NAN": continue
                
                # Ekstrak Lambung dan Standarisasi Nama
                lambung = extract_lambung_code(raw_kpi)
                std_match = get_master_match(raw_equip)
                final_clean_key = std_match if std_match in master_data_map else clean_unit_name(lambung if lambung else raw_equip)
                
                # Ambil Array Data (Series per Hari)
                vals = pd.to_numeric(valid_data_rows.iloc[:, col], errors='coerce')
                
                temp = pd.DataFrame({'Date': dates_series, 'Value': vals.values})
                temp.dropna(subset=['Date'], inplace=True)
                
                if is_hm:
                    if final_clean_key not in hm_data_store: hm_data_store[final_clean_key] = []
                    hm_data_store[final_clean_key].append(temp)
                else:
                    if final_clean_key not in liter_data_store: liter_data_store[final_clean_key] = []
                    liter_data_store[final_clean_key].append(temp)

# ==============================================================================
# 3. FUNGSI LOGIC BANTUAN
# ==============================================================================
def resolve_unit_type(clean_ops, master_map):
    utype = "OTHERS"
    if clean_ops in master_map: 
        utype = normalize_type(master_map[clean_ops]['Jenis_Alat'])
        
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
    df.loc[df['Delta_HM'] > 100, 'Delta_HM'] = 0 # Sanity Check HM
    df['Month_Num'] = df['Date'].dt.month
    return df.groupby('Month_Num')['Delta_HM'].sum().reset_index()

# ==============================================================================
# 4. PREPARE DATA (MENGGABUNGKAN HASIL TRUCKING & NON-TRUCKING)
# ==============================================================================
print("3. Mengolah Data (Trucking & Non-Trucking)...")
df_truck = pd.DataFrame()
if os.path.exists(FILE_TRUCKING):
    df_tr_raw = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    # Filter Hanya Oktober & November yang relevan untuk Truk
    df_tr_raw = df_tr_raw[df_tr_raw['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
    tr_list = []
    month_map_rev = {'Oktober': 10, 'November': 11}
    
    for _, r in df_tr_raw.iterrows():
        raw_u = str(r['Nama_Unit'])
        clean_u = clean_unit_name(raw_u)
        utype = resolve_unit_type(clean_u, master_data_map)
        
        m_num = month_map_rev.get(r['Bulan'].strip())
        hm_val = 0
        liter_val = r['LITER'] if 'LITER' in df_tr_raw.columns and pd.notna(r['LITER']) else 0 
        
        # Tarik data HM dari store yang sudah dicleaning
        if clean_u in hm_data_store:
            res_hm = calculate_monthly_hm_for_unit(hm_data_store[clean_u])
            if not res_hm.empty:
                row_hm = res_hm[res_hm['Month_Num'] == m_num]
                if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
        
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
    
    for _, r in df_nt_raw.iterrows():
        raw_u = str(r[col_u])
        clean_u = clean_unit_name(raw_u)
        utype = resolve_unit_type(clean_u, master_data_map)
             
        if utype in ['REACH STACKER','FORKLIFT','CRANE','SIDE LOADER','TOP LOADER']:
            m_num = m_map_id.get(r['Bulan'].strip())
            hm_val = 0
            
            # Tarik HM untuk Alat Berat
            if clean_u in hm_data_store:
                res_hm = calculate_monthly_hm_for_unit(hm_data_store[clean_u])
                if not res_hm.empty:
                    row_hm = res_hm[res_hm['Month_Num'] == m_num]
                    if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
            
            liter_val = r['LITER'] if 'LITER' in df_nt_raw.columns and pd.notna(r['LITER']) else 0 
            
            nt_list.append({
                'Category': 'NON-TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 'LITER': liter_val, 'Ton': r.get('Total_Ton', 0),
                'Workload': 0, 'HM': hm_val
            })
    df_nt = pd.DataFrame(nt_list)

# ==============================================================================
# 5. FORECASTING ENGINE (CONSTRAINED CAUSAL MODEL)
# ==============================================================================
print("4. Menjalankan Forecasting Engine...")
df_final = pd.concat([df_truck, df_nt], ignore_index=True)

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
    sub = df_final[(df_final['Category']==cfg['cat']) & (df_final['Type']==cfg['type']) & (df_final['Month_Num'].isin(cfg['range']))]
    if sub.empty: continue
    
    agg = sub.groupby('Month_Num').agg({'LITER': 'sum', 'Unit_Clean': 'nunique', **{p: 'sum' for p in cfg['preds']}}).reset_index()
    valid_model = False
    rm = LinearRegression()
    
    if len(agg) >= 2:
        for col in cfg['preds'] + ['LITER']:
            agg[f'{col}_Per_Unit'] = agg[col] / agg['Unit_Clean']
            
        X_train = agg[[f'{p}_Per_Unit' for p in cfg['preds']]]
        y_train = agg['LITER_Per_Unit']
        
        if X_train.sum().sum() > 0:
            rm.fit(X_train, y_train)
            r2_model = r2_score(y_train, rm.predict(X_train))
            valid_model = True

    unique_units = sub['Unit_Clean'].unique()
    
    for unit in unique_units:
        unit_history = sub[sub['Unit_Clean'] == unit]
        unit_pred_act = {}
        
        # Forecast Aktivitas Unit untuk Bulan Desember
        for p in cfg['preds']:
            act_series = unit_history[['Month_Num', p]].dropna()
            if len(act_series) >= 2 and act_series[p].sum() > 0:
                tm = LinearRegression()
                tm.fit(act_series[['Month_Num']], act_series[p])
                pred_val = tm.predict([[12]])[0]
                unit_pred_act[p] = max(0, pred_val)
            elif len(act_series) == 1:
                unit_pred_act[p] = act_series[p].values[0]
            else:
                unit_pred_act[p] = 0
                
        # Hitung Normal LITER berdasarkan Regresi General Armada
        expected_liter_normal = 0
        if valid_model:
            X_pred = [unit_pred_act.get(p, 0) for p in cfg['preds']]
            expected_liter_normal = max(0, rm.predict([X_pred])[0])
            
        # Constrained Efficiency Factor (Batas: 0.8x s/d 1.2x Rata-rata Normal)
        ratios_eff = []
        for _, row in unit_history.iterrows():
            m_num = row['Month_Num']
            X_hist = [row[p] for p in cfg['preds']]
            if valid_model and sum(X_hist) > 0:
                normal_hist = max(1, rm.predict([X_hist])[0])
                actual_lit = row['LITER']
                ratios_eff.append(actual_lit / normal_hist)
                
        eff_factor = 1.0
        final_forecast = 0
        note = "Data Kurang"
        ratio_variance = 0
        
        if valid_model and expected_liter_normal > 0:
            if ratios_eff:
                eff_factor = np.mean(ratios_eff)
                ratio_variance = np.std(ratios_eff)
                # Batas aman faktor irit/boros maksimal 20% deviasi dari Standar Armada
                eff_factor = max(0.80, min(1.20, eff_factor))
                
            final_forecast = expected_liter_normal * eff_factor
            
            if ratios_eff:
                trend_str = "Standard"
                if eff_factor > 1.05: trend_str = f"Boros"
                elif eff_factor < 0.95: trend_str = f"Irit"
                note = f"General -> {trend_str}"
            else:
                note = "Tidak ada history valid, pakai rata-rata general"
                
        error_margin_val = round(ratio_variance * 100, 2) if len(ratios_eff) > 1 else None
        r2_val = round(r2_model, 2) if valid_model else None
        
        pred_wl_ton = unit_pred_act.get('Workload', unit_pred_act.get('Ton', 0))
        
        forecast_detail_list.append({
            'Category': cfg['cat'], 
            'Type': cfg['type'], 
            'Unit_Name': unit,
            'Forecast_HM_Dec': round(unit_pred_act.get('HM', 0), 2), 
            'Forecast_Workload_Ton_Dec': round(pred_wl_ton, 2), 
            'Expected_LITER_Normal': round(expected_liter_normal, 2), 
            'Correction_Factor': round(eff_factor, 2),
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

print("\n=== CONTOH HASIL FORECAST PER UNIT (CONSTRAINED MODEL) ===")
print(df_res_detail[['Unit_Name', 'Forecast_HM_Dec', 'Forecast_Workload_Ton_Dec', 'Expected_LITER_Normal', 'Forecast_LITER_Dec']].head(5))
print(f"\nFile Detail tersimpan: {OUTPUT_EXCEL}")