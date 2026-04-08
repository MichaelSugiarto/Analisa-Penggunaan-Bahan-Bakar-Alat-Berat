import pandas as pd
import numpy as np
import re
import os
import datetime
import warnings
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_percentage_error, mean_squared_error
from statsmodels.tsa.holtwinters import ExponentialSmoothing

warnings.filterwarnings('ignore')

# ==============================================================================
# KONFIGURASI FILE
# ==============================================================================
FILE_TRUCKING = "HasilTrucking.xlsx"
FILE_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_BBM_TEST = "BBM AAB Jan-Mar 2026.xlsx" 
OUTPUT_EXCEL = "Hasil_Eksperimen_Forecast_Semua_Unit.xlsx"

print("=== START BULK FORECASTING: TRUCKING & NON-TRUCKING ===\n")

# ==============================================================================
# FUNGSI HELPER & KAMUS ALIAS
# ==============================================================================
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    
    # KAMUS ALIAS (Penyatuan Unit Kembar)
    if "FL RENTAL 01" in name: name = "FL RENTAL 01"
    elif "TOBATI" in name: name = "TOBATI"
    elif "8477 UUC" in name or "9902 UR" in name: name = "L9902UR"
    elif "WIND RIVER" in name or "TOP LOADER BOSS" in name: name = "WIND RIVER"
        
    name = name.replace("FORKLIFT", "FORKLIF")
    name = re.sub(r'[^A-Z0-9]', '', name) 
    return name

def get_bulan_angka(teks):
    teks = str(teks).lower()
    if 'jan' in teks: return 1
    elif 'feb' in teks: return 2
    elif 'mar' in teks: return 3
    elif 'apr' in teks: return 4
    elif 'mei' in teks: return 5
    elif 'jun' in teks: return 6
    elif 'jul' in teks: return 7
    elif 'agu' in teks or 'agt' in teks: return 8
    elif 'sep' in teks: return 9
    elif 'okt' in teks: return 10
    elif 'nov' in teks: return 11
    elif 'des' in teks: return 12
    return 0

def is_valid_date(val):
    if pd.isna(val): return False
    if isinstance(val, (int, float)) and 1 <= val <= 31: return True
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)): return True
    val_str = str(val).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', val_str): return True 
    if re.match(r'^\d{2}-\d{2}-\d{4}', val_str): return True 
    match = re.match(r'^(\d{1,2})', val_str) 
    if match and 1 <= int(match.group(1)) <= 31: return True
    return False

# ==============================================================================
# 1. LOAD DATA TRAIN 2025 (TRUCKING & NON-TRUCKING)
# ==============================================================================
print("1. Mengumpulkan data historis (Train) seluruh armada...")
df_train_list = []

# --- BACA TRUCKING ---
if os.path.exists(FILE_TRUCKING):
    df_tr = pd.read_excel(FILE_TRUCKING, sheet_name='Data_Bulanan')
    df_tr['Unit_Clean'] = df_tr['Nama_Unit'].apply(clean_unit_name)
    df_tr['Bulan_Angka'] = df_tr['Bulan'].apply(get_bulan_angka)
    df_tr.rename(columns={'Nama_Unit': 'Unit_Asli', 'Total_TonKm': 'Workload', 'Total_Km': 'HM_or_Km'}, inplace=True)
    df_train_list.append(df_tr[['Unit_Clean', 'Unit_Asli', 'Bulan_Angka', 'HM_or_Km', 'Workload', 'LITER']])

# --- BACA NON-TRUCKING ---
if os.path.exists(FILE_NON_TRUCKING):
    df_nt = pd.read_excel(FILE_NON_TRUCKING, sheet_name='Data_Bulanan')
    col_u = 'Unit_Name' if 'Unit_Name' in df_nt.columns else 'Nama_Unit'
    df_nt['Unit_Clean'] = df_nt[col_u].apply(clean_unit_name)
    df_nt['Bulan_Angka'] = df_nt['Bulan'].apply(get_bulan_angka)
    df_nt['HM_or_Km'] = df_nt['Total_Ton'] # Pengganti HM sementara
    df_nt.rename(columns={col_u: 'Unit_Asli', 'Total_Ton': 'Workload'}, inplace=True)
    df_train_list.append(df_nt[['Unit_Clean', 'Unit_Asli', 'Bulan_Angka', 'HM_or_Km', 'Workload', 'LITER']])

if not df_train_list:
    print("ERROR: File Trucking & Non-Trucking tidak ditemukan!")
    exit()

df_train_all = pd.concat(df_train_list, ignore_index=True)
df_train_all.fillna(0, inplace=True)

# ==============================================================================
# 2. LOAD DATA TEST AKTUAL JAN-MAR 2026
# ==============================================================================
print("2. Mengekstraksi data Test aktual 2026 untuk evaluasi...")
test_data_dict = {}

if os.path.exists(FILE_BBM_TEST):
    xls_test = pd.ExcelFile(FILE_BBM_TEST)
    for sheet in xls_test.sheet_names:
        b_angka = get_bulan_angka(sheet)
        if b_angka == 0: continue
            
        df_temp = pd.read_excel(xls_test, sheet_name=sheet, header=None)
        
        # Hardcode index agar 100% aman dari teks pengganggu di baris data
        r_eq, r_hd = 0, 2
        
        unit_names_row = df_temp.iloc[r_eq].ffill() 
        headers_row = df_temp.iloc[r_hd]
        
        mask_valid_dates = df_temp.iloc[:, 0].apply(is_valid_date)
        data_rows = df_temp[mask_valid_dates]
        
        for col in range(1, df_temp.shape[1]):
            head_unit = clean_unit_name(str(unit_names_row.iloc[col]))
            head_metric = str(headers_row.iloc[col]).strip().upper()
            
            if head_unit and 'LITER' in head_metric:
                vals = pd.to_numeric(data_rows.iloc[:, col], errors='coerce').fillna(0)
                tot_liter = vals.sum()
                
                if head_unit not in test_data_dict:
                    test_data_dict[head_unit] = {}
                test_data_dict[head_unit][b_angka] = test_data_dict[head_unit].get(b_angka, 0) + tot_liter

# ==============================================================================
# 3. PROSES MACHINE LEARNING & FILTERING
# ==============================================================================
print("3. Memulai proses Machine Learning ke seluruh armada...\n")
unique_units = df_train_all['Unit_Clean'].unique()

results = []
excluded_units = [] # Variabel baru untuk menampung unit yang HM/LITER = 0
unit_berhasil = 0

for unit_id in unique_units:
    df_u = df_train_all[df_train_all['Unit_Clean'] == unit_id].copy()
    unit_asli = df_u['Unit_Asli'].iloc[0]
    
    df_u = df_u.groupby('Bulan_Angka')[['HM_or_Km', 'Workload', 'LITER']].sum().reset_index()
    df_u = df_u.sort_values(by='Bulan_Angka').reset_index(drop=True)
    
    # --- PENGECEKAN FILTER (HM = 0 atau LITER = 0) ---
    tot_aktivitas = df_u['HM_or_Km'].sum()
    tot_liter = df_u['LITER'].sum()
    
    if tot_aktivitas == 0 or tot_liter == 0 or len(df_u) < 4:
        alasan = "Aktivitas (HM/KM) atau Konsumsi BBM = 0 sepanjang 2025" if len(df_u) >= 4 else "Data kurang dari 4 bulan"
        excluded_units.append({
            'Nama_Unit': unit_asli,
            'Total_Aktivitas_2025': tot_aktivitas,
            'Total_LITER_2025': tot_liter,
            'Alasan_Dikecualikan': alasan
        })
        continue
        
    try:
        # --- TAHAP 1: HOLT'S DAMPED TREND (Prediksi Dinamis & Aman) ---
        # trend='add' mengaktifkan pembacaan tren naik/turun
        # damped_trend=True memberikan 'rem' agar tren tidak meledak di masa depan
        model_b1 = ExponentialSmoothing(df_u['HM_or_Km'], trend='add', damped_trend=True, seasonal=None, initialization_method="estimated").fit()
        pred_b1 = model_b1.forecast(3).values
        
        model_b2 = ExponentialSmoothing(df_u['Workload'], trend='add', damped_trend=True, seasonal=None, initialization_method="estimated").fit()
        pred_b2 = model_b2.forecast(3).values
        
        # --- TAHAP 2: REGRESI BBM ANTI-BLOWUP ---
        X_train = df_u[['HM_or_Km', 'Workload']]
        y_train = df_u['LITER']
        
        # LinearRegression dengan parameter pengunci intercept
        mlr = LinearRegression(fit_intercept=False, positive=True)
        mlr.fit(X_train, y_train)
        
        X_test_pred = pd.DataFrame({'HM_or_Km': pred_b1, 'Workload': pred_b2})
        pred_liter = mlr.predict(X_test_pred)
        pred_liter = [max(0, p) for p in pred_liter]
        
        # --- TAHAP 3: EVALUASI ---
        b_test = [1, 2, 3]
        act_liter = [test_data_dict.get(unit_id, {}).get(b, 0) for b in b_test]
        
        mape_val = "N/A"
        rmse_val = 0
        status = "Menunggu Data 2026"
        
        if sum(act_liter) > 0:
            act_for_mape = [a for a in act_liter if a > 0]
            pred_for_mape = [p for a, p in zip(act_liter, pred_liter) if a > 0]
            
            if act_for_mape:
                mape = mean_absolute_percentage_error(act_for_mape, pred_for_mape)
                mape_val = f"{mape * 100:.2f}%"
            
            rmse_val = np.sqrt(mean_squared_error(act_liter, pred_liter))
            status = "Terverifikasi (Ada Data Test)"

        results.append({
            'Nama_Unit': unit_asli,
            'Prediksi_Jan_26': round(pred_liter[0], 2),
            'Prediksi_Feb_26': round(pred_liter[1], 2),
            'Prediksi_Mar_26': round(pred_liter[2], 2),
            'Total_Prediksi_3Bln': round(sum(pred_liter), 2),
            'Aktual_Jan_26': round(act_liter[0], 2),
            'Aktual_Feb_26': round(act_liter[1], 2),
            'Aktual_Mar_26': round(act_liter[2], 2),
            'Total_Aktual_3Bln': round(sum(act_liter), 2),
            'Selisih_Liter (RMSE)': round(rmse_val, 2),
            'Error_Prediksi (MAPE)': mape_val,
            'Status': status
        })
        unit_berhasil += 1
        
    except Exception as e:
        continue

# ==============================================================================
# 4. EXPORT KE EXCEL (MULTI-SHEET)
# ==============================================================================
df_res = pd.DataFrame(results)
df_exc = pd.DataFrame(excluded_units)

# Menggunakan ExcelWriter untuk membuat Sub-Sheet / Tab di Excel
with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
    if not df_res.empty:
        df_res = df_res.sort_values(by='Total_Aktual_3Bln', ascending=False)
        df_res.to_excel(writer, sheet_name='Forecast_Valid', index=False)
    
    if not df_exc.empty:
        df_exc.to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)

if not df_res.empty:
    print(f"✅ BERHASIL! {unit_berhasil} Armada (Trucking & Alat Berat) selesai di-forecast.")
    print(f"✅ {len(df_exc)} Armada dikeluarkan karena tidak ada aktivitas/BBM.")
    print(f"📁 Laporan tersimpan di: {OUTPUT_EXCEL}")
else:
    print("GAGAL: Tidak ada unit yang memenuhi syarat untuk di-forecast.")