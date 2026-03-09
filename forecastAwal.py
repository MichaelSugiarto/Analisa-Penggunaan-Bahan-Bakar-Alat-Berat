import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import re
import os
import warnings

warnings.filterwarnings('ignore')

# --- KONFIGURASI FILE ---
FILE_BBM = "BBM AAB.xlsx"
FILE_MASTER = "cost & bbm 2022 sd 2025 HP & Type.xlsx"
FILE_TRUCKING = "HasilTrucking.xlsx"
FILE_NON_TRUCKING = "HasilNonTrucking.xlsx"
OUTPUT_EXCEL = "Forecast_Detail_Data.xlsx"

# --- FUNGSI HELPER ---
def clean_unit_name(name):
    """Membersihkan nama unit"""
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

print("=== START FORECASTING (PER UNIT BASIS) ===\n")

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
    except Exception as e:
        print(f"   [ERROR] Load Master: {e}")

# ==============================================================================
# 2. BACA BBM AAB
# ==============================================================================
print("2. Membaca BBM AAB...")
target_sheets = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']
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
                if header_str == 'HM':
                    raw_unit_name_bbm = str(unit_names_row[col]).strip().upper()
                    if raw_unit_name_bbm == "" or "UNNAMED" in raw_unit_name_bbm or "EQUIP NAME" in raw_unit_name_bbm or "TOTAL" in raw_unit_name_bbm:
                        continue
                    if raw_unit_name_bbm.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL')):
                        continue
                    
                    vals = pd.to_numeric(df.iloc[3:, col], errors='coerce')
                    clean_key_bbm = clean_unit_name(raw_unit_name_bbm)
                    
                    if clean_key_bbm not in bbm_data_store:
                        bbm_data_store[clean_key_bbm] = []
                    
                    temp_df = pd.DataFrame({'Date': dates, 'Value': vals})
                    temp_df.dropna(subset=['Date'], inplace=True)
                    temp_df['Date'] = pd.to_datetime(temp_df['Date'], dayfirst=True, errors='coerce')
                    temp_df.dropna(subset=['Date'], inplace=True)
                    bbm_data_store[clean_key_bbm].append(temp_df)

print(f"   -> Terbaca {len(bbm_data_store)} unit unik dari file BBM AAB.")

# ==============================================================================
# 3. FUNGSI MATCHING (FINAL FIXED)
# ==============================================================================
def find_matching_hm_unit(ops_unit_name, bbm_keys):
    clean_ops = clean_unit_name(ops_unit_name)
    
    # RULE 1: SPECIAL
    if "FL RENTAL 01" in clean_ops and "TIMIKA" not in clean_ops:
        target = clean_unit_name("FL RENTAL 01 TIMIKA")
        if target in bbm_keys: return target
    if "TOBATI" in clean_ops and "KALMAR" in clean_ops:
        for k in bbm_keys:
            if "TOBATI" in k and "KALMAR" in k: return k
    if "L 8477 UUC" in clean_ops:
        for k in bbm_keys:
            if "L 9902 UR" in k: return k
    if "BOSS" in clean_ops and "TOP" in clean_ops: 
        for k in bbm_keys:
            if "WIND RIVER" in k: return k
            
    # RULE 2: EXACT
    if clean_ops in bbm_keys: return clean_ops
    # RULE 3: CONTAINS
    for bbm_k in bbm_keys:
        if len(clean_ops) > 3 and clean_ops in bbm_k: return bbm_k
    # RULE 4: EX
    for bbm_k in bbm_keys:
        if "EX." in bbm_k or " EX " in bbm_k:
            parts = re.split(r'EX\.| EX ', bbm_k)
            if len(parts) > 1:
                after_ex = clean_unit_name(parts[1].replace(')', '').strip())
                if after_ex == clean_ops: return bbm_k
                if len(after_ex) > 3 and after_ex in clean_ops: return bbm_k
    # RULE 5: BRACKET
    for bbm_k in bbm_keys:
        if "(" in bbm_k:
            before_bracket = clean_unit_name(bbm_k.split("(")[0].strip())
            if before_bracket == clean_ops: return bbm_k
            if len(before_bracket) > 3 and before_bracket in clean_ops: return bbm_k

    return None

# ==============================================================================
# 4. HITUNG HM
# ==============================================================================
def calculate_monthly_hm_for_unit(df_list):
    if not df_list: return pd.DataFrame()
    df_concat = pd.concat(df_list, ignore_index=True)
    df_concat.sort_values('Date', inplace=True)
    df_concat['HM_Clean'] = df_concat['Value'].replace(0, np.nan).ffill().fillna(0)
    df_concat['Delta_HM'] = df_concat['HM_Clean'].diff().fillna(0)
    df_concat.loc[df_concat['Delta_HM'] < 0, 'Delta_HM'] = 0
    df_concat.loc[df_concat['Delta_HM'] > 100, 'Delta_HM'] = 0
    df_concat['Month_Num'] = df_concat['Date'].dt.month
    return df_concat.groupby('Month_Num')['Delta_HM'].sum().reset_index()

# ==============================================================================
# 5. DATA TRUCKING
# ==============================================================================
print("3. Mengolah Data Trucking...")
df_truck = pd.DataFrame()
if os.path.exists(FILE_TRUCKING):
    df_tr_raw = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    df_tr_raw = df_tr_raw[df_tr_raw['Bulan'].astype(str).str.strip().isin(['Oktober', 'November'])]
    tr_list = []
    month_map_rev = {'Oktober': 10, 'November': 11}
    bbm_keys_list = list(bbm_data_store.keys())
    
    for _, r in df_tr_raw.iterrows():
        raw_u = str(r['Nama_Unit'])
        matched_bbm_key = find_matching_hm_unit(raw_u, bbm_keys_list)
        utype = "OTHERS"
        if matched_bbm_key and matched_bbm_key in master_data_map:
            utype = normalize_type(master_data_map[matched_bbm_key]['Jenis_Alat'])
        else:
            if "TRONTON" in raw_u.upper(): utype = "TRONTON"
            elif re.search(r'L\s*\d+', raw_u.upper()) or "TRAILER" in raw_u.upper(): utype = "TRAILER"
        
        if utype in ['TRAILER', 'TRONTON']:
            m_num = month_map_rev.get(r['Bulan'].strip())
            hm_val = 0
            if matched_bbm_key:
                df_hm_res = calculate_monthly_hm_for_unit(bbm_data_store[matched_bbm_key])
                if not df_hm_res.empty:
                    row_hm = df_hm_res[df_hm_res['Month_Num'] == m_num]
                    if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
            
            tr_list.append({
                'Category': 'TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 'LITER': r['LITER'], 'Workload': r['Total_TonKm'],
                'Ton': 0, 'HM': hm_val
            })
    df_truck = pd.DataFrame(tr_list)

# ==============================================================================
# 6. DATA NON-TRUCKING
# ==============================================================================
print("4. Mengolah Data Non-Trucking...")
df_nt = pd.DataFrame()
if os.path.exists(FILE_NON_TRUCKING):
    df_nt_raw = pd.read_excel(FILE_NON_TRUCKING, sheet_name='Data_Bulanan')
    valid_months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November']
    df_nt_raw = df_nt_raw[df_nt_raw['Bulan'].isin(valid_months)]
    m_map_id = {'Januari':1, 'Februari':2, 'Maret':3, 'April':4, 'Mei':5, 'Juni':6,
                'Juli':7, 'Agustus':8, 'September':9, 'Oktober':10, 'November':11}
    nt_list = []
    bbm_keys_list = list(bbm_data_store.keys())
    
    col_u = 'Unit_Name' if 'Unit_Name' in df_nt_raw.columns else 'Nama_Unit'
    for _, r in df_nt_raw.iterrows():
        raw_u = str(r[col_u])
        matched_bbm_key = find_matching_hm_unit(raw_u, bbm_keys_list)
        utype = "OTHERS"
        if 'Jenis_Alat' in df_nt_raw.columns and pd.notna(r['Jenis_Alat']):
            utype = normalize_type(r['Jenis_Alat'])
        elif matched_bbm_key and matched_bbm_key in master_data_map:
             utype = normalize_type(master_data_map[matched_bbm_key]['Jenis_Alat'])
             
        if utype in ['REACH STACKER','FORKLIFT','CRANE','SIDE LOADER','TOP LOADER']:
            m_num = m_map_id.get(r['Bulan'])
            hm_val = 0
            if matched_bbm_key:
                df_hm_res = calculate_monthly_hm_for_unit(bbm_data_store[matched_bbm_key])
                if not df_hm_res.empty:
                    row_hm = df_hm_res[df_hm_res['Month_Num'] == m_num]
                    if not row_hm.empty: hm_val = row_hm['Delta_HM'].values[0]
            
            nt_list.append({
                'Category': 'NON-TRUCKING', 'Type': utype, 'Unit_Clean': raw_u,
                'Month_Num': m_num, 'LITER': r['LITER'], 'Ton': r['Total_Ton'],
                'Workload': 0, 'HM': hm_val
            })
    df_nt = pd.DataFrame(nt_list)

# ==============================================================================
# 7. FORECAST (PER UNIT BASIS)
# ==============================================================================
print("5. Forecasting (Metode: Average Per Unit)...")
df_final = pd.concat([df_truck, df_nt], ignore_index=True)
results = []
configs = [
    {'cat': 'TRUCKING', 'type': 'TRAILER', 'preds': ['Workload', 'HM'], 'range': [10, 11]},
    {'cat': 'TRUCKING', 'type': 'TRONTON', 'preds': ['Workload', 'HM'], 'range': [10, 11]},
    {'cat': 'NON-TRUCKING', 'type': 'REACH STACKER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'FORKLIFT', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'CRANE', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'SIDE LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)},
    {'cat': 'NON-TRUCKING', 'type': 'TOP LOADER', 'preds': ['Ton', 'HM'], 'range': range(1, 12)}
]

for cfg in configs:
    # 1. Filter Data
    sub = df_final[(df_final['Category']==cfg['cat']) & (df_final['Type']==cfg['type']) & (df_final['Month_Num'].isin(cfg['range']))]
    
    # 2. Agregasi TOTAL per Bulan, tapi hitung juga JUMLAH UNIT (nunique)
    agg = sub.groupby('Month_Num').agg({
        'LITER': 'sum',
        'Unit_Clean': 'nunique', # Hitung ada berapa unit aktif bulan itu
        **{p: 'sum' for p in cfg['preds']}
    }).reset_index()
    
    if len(agg) < 2:
        results.append({'Type': cfg['type'], 'Status': 'Data Kurang', 'Forecast_LITER_Per_Unit': 0, 'Trend_Info': '-'})
    else:
        # 3. Konversi ke Rata-Rata PER UNIT
        # Agar forecastnya adalah "Kebutuhan Solar untuk 1 Unit"
        for col in cfg['preds'] + ['LITER']:
            agg[f'{col}_Per_Unit'] = agg[col] / agg['Unit_Clean']
            
        dec_X = []
        info = []
        tm = LinearRegression()
        
        # Forecast Predictor (Average Activity per Unit)
        for p in cfg['preds']:
            col_avg = f'{p}_Per_Unit'
            if agg[col_avg].sum() == 0:
                dec_X.append(0)
                info.append(f"{p}:NoData")
            else:
                tm.fit(agg[['Month_Num']], agg[col_avg])
                proj = max(0, tm.predict([[12]])[0]) # Prediksi bulan ke-12
                dec_X.append(proj)
                trend = "Naik" if tm.coef_[0]>0 else "Turun"
                info.append(f"{p}:{trend}")
        
        # Forecast Fuel (Average Fuel per Unit)
        # Kita regresi variabel Average Activity -> Average Fuel
        rm = LinearRegression()
        X_train = agg[[f'{p}_Per_Unit' for p in cfg['preds']]]
        y_train = agg['LITER_Per_Unit']
        
        rm.fit(X_train, y_train)
        fc = max(0, rm.predict([dec_X])[0])
        
        # Format Rumus
        formula = f"{rm.intercept_:,.0f}" + "".join([f" + ({rm.coef_[i]:.2f}*{p}_Avg)" for i,p in enumerate(cfg['preds'])])
        
        results.append({
            'Category': cfg['cat'], 
            'Type': cfg['type'], 
            'Status': 'Sukses',
            'Forecast_LITER_Per_Unit': fc, # Ini utk 1 unit
            'Trend_Info': ", ".join(info), 
            'Rumus_Per_Unit': formula
        })

df_res = pd.DataFrame(results)
with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
    df_res.to_excel(writer, sheet_name='Forecast_Result', index=False)
    df_final.to_excel(writer, sheet_name='Data_Source_Detail', index=False)

print("\n=== HASIL FORECASTING (PER 1 UNIT) ===")
print(df_res[['Type', 'Status', 'Forecast_LITER_Per_Unit', 'Trend_Info']])
print(f"\nFile Detail tersimpan: {OUTPUT_EXCEL}")