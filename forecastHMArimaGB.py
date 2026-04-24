import pandas as pd
import numpy as np
import re
from sklearn.metrics import mean_squared_error
from sklearn.ensemble import HistGradientBoostingRegressor
import pmdarima as pm
import warnings

warnings.filterwarnings("ignore")

# ==========================================
# FASE 1 & 2: EKSTRAKSI DAN DATA ENGINEERING
# ==========================================
def load_and_melt_excel(file_path, target_sheets):
    print(f"Mengekstrak data dari: {file_path}")
    all_data = []
    xls = pd.read_excel(file_path, sheet_name=None, header=[0, 1, 2])
    
    for sheet_name, df in xls.items():
        if sheet_name not in target_sheets:
            continue
        if df.empty or len(df.columns) == 0:
            print(f"  -> Melewati sheet kosong: '{sheet_name}'")
            continue
        try:
            df = df.set_index(df.columns[0]) 
            df.index.name = 'TANGGAL'
            df.columns = df.columns.droplevel(1) 
            df_stacked = df.stack(level=0).reset_index()
            df_stacked.rename(columns={'level_1': 'EQUIP NAME'}, inplace=True)
            all_data.append(df_stacked)
        except Exception as e:
            print(f"  -> Gagal memproses sheet '{sheet_name}': {e}")
            continue
            
    if not all_data:
        return pd.DataFrame()
        
    df_final = pd.concat(all_data, ignore_index=True)
    df_final['TANGGAL'] = pd.to_datetime(df_final['TANGGAL'], format='%d-%m-%Y', errors='coerce')
    df_final = df_final.dropna(subset=['TANGGAL'])
    return df_final

def prepare_data():
    sheets_2023 = ['01. Jan', '02. Feb', '03. Mar', '04. Apr', '05. Mei', '06. Jun', '07. Jul', '08. Aug', '09. Sep', '10. Okt', '11. Nov', '12. Des']
    sheets_2024_25 = ['JANUARI 24', 'FEB 24', 'maret 24', 'April 24', 'Mei 24', 'Juni 24', 'Juli 24', 'agt 24', 'sept 24', 'okt 24', 'nov 24', 'des 24', 'Des 25']
    sheets_2025 = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']
    
    df_2023 = load_and_melt_excel('BBM AAB 2023.xlsx', target_sheets=sheets_2023)
    df_2024_des25 = load_and_melt_excel('BBM AAB 2024 & Des 2025.xlsx', target_sheets=sheets_2024_25)
    df_jan_nov_25 = load_and_melt_excel('BBM AAB Jan-Nov 2025.xlsx', target_sheets=sheets_2025)
    
    df_all = pd.concat([df_2023, df_2024_des25, df_jan_nov_25], ignore_index=True)
    df_all = df_all.sort_values(by=['EQUIP NAME', 'TANGGAL'])
    
    df_all['HM_Clean'] = pd.to_numeric(df_all['HM'], errors='coerce').replace(0, np.nan)
    df_all['HM_Clean'] = df_all.groupby('EQUIP NAME')['HM_Clean'].ffill().fillna(0)
    df_all['Delta_HM'] = df_all.groupby('EQUIP NAME')['HM_Clean'].diff().fillna(0)
    df_all.loc[df_all['Delta_HM'] < 0, 'Delta_HM'] = 0
    df_all.loc[df_all['Delta_HM'] > 100, 'Delta_HM'] = 0 
    df_all['LITER_Clean'] = pd.to_numeric(df_all['LITER'], errors='coerce').fillna(0)
    
    df_all['TAHUN_BULAN'] = df_all['TANGGAL'].dt.to_period('M')
    agg_data = df_all.groupby(['EQUIP NAME', 'TAHUN_BULAN']).agg({'Delta_HM': 'sum', 'LITER_Clean': 'sum'}).reset_index()
    agg_data.rename(columns={'Delta_HM': 'HM', 'LITER_Clean': 'LITER'}, inplace=True)
    
    train_agg = agg_data[agg_data['TAHUN_BULAN'] <= '2024-12']
    test_agg = agg_data[agg_data['TAHUN_BULAN'] >= '2025-01']
    return train_agg, test_agg

# ==========================================
# FUNGSI PREPROCESSING & MODELING
# ==========================================
def preprocess_timeseries(series):
    df = pd.DataFrame(series)
    df.columns = ['HM']
    p95 = df['HM'].quantile(0.95)
    df['HM_Capped'] = np.minimum(df['HM'], p95) if p95 > 0 else df['HM']
    df['HM_Smoothed'] = df['HM_Capped'].rolling(window=2, min_periods=1).mean()
    return df['HM_Smoothed']

def prepare_boosting_features(series):
    df = pd.DataFrame(series)
    df.columns = ['y']
    df['lag_1'], df['lag_2'], df['lag_3'] = df['y'].shift(1), df['y'].shift(2), df['y'].shift(3)
    df = df.dropna()
    return df[['lag_1', 'lag_2', 'lag_3']], df['y']

def predict_boosting(train_series, steps_ahead):
    X_train, y_train = prepare_boosting_features(train_series)
    if len(X_train) < 3: return np.zeros(steps_ahead)
    model = HistGradientBoostingRegressor(random_state=42).fit(X_train, y_train)
    predictions, curr = [], train_series.tail(3).values[::-1].tolist()
    for _ in range(steps_ahead):
        pred = max(0, model.predict(pd.DataFrame([curr], columns=['lag_1','lag_2','lag_3']))[0])
        predictions.append(pred)
        curr = [pred] + curr[:2]
    return np.array(predictions)

def hitung_mape_aman(actual, pred):
    actual, pred = np.array(actual), np.array(pred)
    mask = actual != 0
    if not np.any(mask): 
        return 0.0 
    return np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100

# ==========================================
# MAPPING NAMA ALAT BERAT KE MASTER
# ==========================================
def load_master_names():
    try:
        master_df = pd.read_excel('cost & bbm 2022 sd 2025 HP & Type.xlsx', header=1)
        master_names = set(master_df['NAMA ALAT BERAT'].dropna().astype(str).str.strip())
        return master_names
    except Exception as e:
        print("\n[!] Gagal membaca file 'cost & bbm 2022 sd 2025 HP & Type.xlsx'.")
        print(f"Detail error: {e}\n")
        return set()

def get_mapped_unit_name(unit_name, master_names):
    hardcoded = {
        "FL RENTAL 01": "FL RENTAL 01 TIMIKA",
        "TOBATI (EX.FL KALMAR 32T)": "TOP LOADER KALMAR 35T/TOBATI",
        "L 8477 UUC (EX.L 9902 UR)": "L 9902 UR / S75",
        "WIND RIVER (EX.TL BOSS 42T)": "TOP LOADER BOSS"
    }
    
    if unit_name in hardcoded and hardcoded[unit_name] in master_names:
        return hardcoded[unit_name]
        
    if unit_name in master_names:
        return unit_name
        
    if " (" in unit_name:
        before_paren = unit_name.split(" (")[0].strip()
        if before_paren in master_names:
            return before_paren
            
    if "EX." in unit_name:
        match_ex = re.search(r'EX\.([^\)]+)', unit_name)
        if match_ex:
            after_ex = match_ex.group(1).strip()
            if after_ex in master_names:
                return after_ex
                
    return None

# ==========================================
# FASE 3: PIPELINE FORECASTING
# ==========================================
def run_forecast_pipeline():
    train_agg, test_agg = prepare_data()
    list_unit_raw = train_agg['EQUIP NAME'].unique()
    
    master_names = load_master_names()
    if not master_names:
        print("Proses dihentikan karena data master alat berat gagal dimuat.")
        return
    
    results_combined = []
    metrics_list = []
    excluded_units_list = []
    
    all_actual_hm = []
    all_pred_arima_hm = []
    all_pred_gb_hm = []
    
    total_valid_population = 0
    
    print(f"\n[AI ENGINE] Memulai pelatihan: ARIMA vs Gradient Boosting...")
    
    for unit in list_unit_raw:
        try:
            # ---------------------------------------------------------
            # MAPPING MASTER FILE
            # ---------------------------------------------------------
            mapped_name = get_mapped_unit_name(unit, master_names)
            if not mapped_name:
                continue
                
            total_valid_population += 1
                
            df_u_train = train_agg[train_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
            df_u_test = test_agg[test_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
            
            # ---------------------------------------------------------
            # FILTER LOGIS DATA CACAT (UPDATE FILTER 1 TAHUN = 0)
            # ---------------------------------------------------------
            if df_u_test.empty:
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'Tidak ada data aktual di tahun uji (2025).'})
                continue
                
            if len(df_u_train) < 12:
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'Data Latih (Train) kurang dari 12 bulan.'})
                continue
                
            if len(df_u_test) < 12:
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'Data Uji (Test) kurang dari 12 bulan.'})
                continue
                
            # Filter 1: Cek HM atau LITER = 0 selama 1 tahun (12 bln berturut-turut) di Data Latih (2023-2024)
            rolling_hm_train = df_u_train['HM'].rolling(window=12).sum()
            rolling_liter_train = df_u_train['LITER'].rolling(window=12).sum()
            if (rolling_hm_train == 0).any() or (rolling_liter_train == 0).any():
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'HM atau LITER bernilai 0 selama 1 tahun di Data Latih.'})
                continue
                
            # Filter 2: Cek HM atau LITER = 0 secara total di Data Uji (Test 2025 -> 1 tahun)
            if df_u_test['HM'].sum() == 0 or df_u_test['LITER'].sum() == 0:
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'HM atau LITER Aktual bernilai 0 selama 1 tahun di Data Uji (2025).'})
                continue
                
            # Filter 3 (Opsional Pengaman): Cek Pensiun/Rusak mendadak di 3 bulan akhir 2024
            last_3_months = df_u_train.tail(3)
            if len(last_3_months) == 3 and (last_3_months['HM'].sum() == 0 or last_3_months['LITER'].sum() == 0):
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'Alat pensiun/rusak mendadak (HM atau LITER 0 selama 3 bulan di akhir 2024).'})
                continue
            
            # --- Rasio Liter/HM Sederhana ---
            true_ratio = df_u_train['LITER'].sum() / df_u_train['HM'].sum() if df_u_train['HM'].sum() > 0 else 0
            aktual_l, aktual_h = df_u_test['LITER'].values, df_u_test['HM'].values
            steps = len(df_u_test)

            # --- Preprocessing 1: Eksperimen Pemotongan (Truncation) ---
            t_utuh = df_u_train['HM'].copy()
            try:
                first_idx = df_u_train[df_u_train['HM'] > 0].index[0]
                t_potong = df_u_train.loc[first_idx:]['HM'].copy()
            except IndexError:
                t_potong = t_utuh
            
            best_arima = np.zeros(steps)
            best_gb = np.zeros(steps)
            min_rmse = float('inf')
            model_success = False
            
            # A/B Testing
            for ds_name, ds in [("Utuh", t_utuh), ("Potong", t_potong)]:
                if len(ds) < 6: 
                    continue
                model_success = True
                
                ds_s = preprocess_timeseries(ds)
                
                try:
                    p_arima = np.maximum(0, pm.auto_arima(ds_s, seasonal=False, suppress_warnings=True, error_action="ignore").predict(n_periods=steps).values)
                except Exception:
                    p_arima = np.zeros(steps)
                    
                try:
                    p_gb = predict_boosting(ds_s, steps)
                except Exception:
                    p_gb = np.zeros(steps)
                    
                rmse_ds = np.sqrt(mean_squared_error(aktual_h, (p_arima + p_gb)/2))
                if rmse_ds < min_rmse: 
                    min_rmse, best_arima, best_gb = rmse_ds, p_arima, p_gb

            if not model_success:
                excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': 'Data historis valid setelah dipotong kurang dari 6 bulan.'})
                continue

            # --- Kalkulasi Metrik ---
            rmse_arima = np.sqrt(mean_squared_error(aktual_h, best_arima))
            mape_arima = hitung_mape_aman(aktual_h, best_arima)
            
            rmse_gb = np.sqrt(mean_squared_error(aktual_h, best_gb))
            mape_gb = hitung_mape_aman(aktual_h, best_gb)
            
            metrics_list.append({
                'EQUIP NAME': unit,
                'NAMA_MASTER_TERPETAKAN': mapped_name,
                'RMSE_HM_ARIMA': round(rmse_arima, 2),
                'MAPE_HM_ARIMA (%)': round(mape_arima, 2),
                'RMSE_HM_GB': round(rmse_gb, 2),
                'MAPE_HM_GB (%)': round(mape_gb, 2)
            })

            all_actual_hm.extend(aktual_h)
            all_pred_arima_hm.extend(best_arima)
            all_pred_gb_hm.extend(best_gb)

            # --- Simpan Hasil ---
            for i, period in enumerate(df_u_test.index):
                pred_liter_a = best_arima[i] * true_ratio
                pred_liter_g = best_gb[i] * true_ratio
                
                results_combined.append({
                    'EQUIP NAME': unit, 
                    'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Bulan': str(period), 
                    'Aktual_HM': round(aktual_h[i], 2), 
                    'Aktual_LITER': round(aktual_l[i], 2), 
                    'Prediksi_HM_ARIMA': round(best_arima[i], 2), 
                    'Prediksi_LITER_ARIMA': round(pred_liter_a, 2),
                    'Prediksi_HM_GB': round(best_gb[i], 2), 
                    'Prediksi_LITER_GB': round(pred_liter_g, 2)
                })
                
        except Exception as e:
            excluded_units_list.append({'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name, 'Alasan': f'Gagal diproses (Internal Error): {str(e)}'})
            continue

    # ========================================================
    # KEPUTUSAN & PEMBUATAN EXCEL 3 SUBSHEET
    # ========================================================
    df_combined = pd.DataFrame(results_combined)
    df_metrics = pd.DataFrame(metrics_list)
    
    if df_combined.empty:
        print("\nGAGAL: Tidak ada unit valid yang berhasil diproses.")
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel('Unit_Dikecualikan_Error.xlsx', index=False)
        return

    df_combined = df_combined.merge(df_metrics, on=['EQUIP NAME', 'NAMA_MASTER_TERPETAKAN'], how='left')

    global_mape_arima = hitung_mape_aman(all_actual_hm, all_pred_arima_hm)
    global_mape_gb = hitung_mape_aman(all_actual_hm, all_pred_gb_hm)
    best_name = "Gradient Boosting" if global_mape_gb < global_mape_arima else "ARIMA"
    
    df_terpilih = df_combined[['EQUIP NAME', 'NAMA_MASTER_TERPETAKAN', 'Bulan', 'Aktual_HM', 'Aktual_LITER']].copy()
    if best_name == "ARIMA":
        df_terpilih['Prediksi_HM'] = df_combined['Prediksi_HM_ARIMA']
        df_terpilih['Prediksi_LITER'] = df_combined['Prediksi_LITER_ARIMA']
        df_terpilih['MAPE_Unit'] = df_combined['MAPE_HM_ARIMA (%)']
    else:
        df_terpilih['Prediksi_HM'] = df_combined['Prediksi_HM_GB']
        df_terpilih['Prediksi_LITER'] = df_combined['Prediksi_LITER_GB']
        df_terpilih['MAPE_Unit'] = df_combined['MAPE_HM_GB (%)']
        
    df_terpilih.insert(3, 'Model_Terpilih_Global', best_name)

    # --- PEMISAHAN BERDASARKAN THRESHOLD MAPE 35% ---
    df_mape_under_35 = df_terpilih[df_terpilih['MAPE_Unit'] < 35].copy()
    df_mape_over_35 = df_terpilih[df_terpilih['MAPE_Unit'] >= 35].copy()

    # Ekspor ke Excel
    with pd.ExcelWriter('Hasil_Forecast_Final.xlsx') as writer:
        if not df_mape_under_35.empty:
            df_mape_under_35.to_excel(writer, sheet_name='Akurasi_Bagus_Under35', index=False)
        if not df_mape_over_35.empty:
            df_mape_over_35.to_excel(writer, sheet_name='Akurasi_Rendah_Over35', index=False)
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)
            
    # ========================================================
    # OUTPUT SUMMARY KE TERMINAL
    # ========================================================
    total_under_35 = len(df_mape_under_35['EQUIP NAME'].unique()) if not df_mape_under_35.empty else 0
    total_over_35 = len(df_mape_over_35['EQUIP NAME'].unique()) if not df_mape_over_35.empty else 0
    total_excluded = len(excluded_units_list)
    
    if total_valid_population == 0:
        pct_under = pct_over = pct_excl = 0.0
    else:
        pct_under = (total_under_35 / total_valid_population) * 100
        pct_over = (total_over_35 / total_valid_population) * 100
        pct_excl = (total_excluded / total_valid_population) * 100
    
    print("\n" + "="*55)
    print(" SUMMARY HASIL ANALISA OPERASIONAL SPIL ".center(55))
    print("="*55)
    print(f"Total Populasi Alat Berat (Terpetakan Master): {total_valid_population} unit")
    print("-" * 55)
    print(f"1. Unit Akurasi Bagus (< 35%)   : {total_under_35} unit ({pct_under:.1f}%)")
    print(f"2. Unit Akurasi Rendah (>= 35%) : {total_over_35} unit ({pct_over:.1f}%)")
    print(f"3. Unit Diexclude (Data Cacat)  : {total_excluded} unit ({pct_excl:.1f}%)")
    print("="*55)
    print(f"Laporan berhasil disimpan ke 'Hasil_Forecast_Final.xlsx'")

if __name__ == "__main__":
    run_forecast_pipeline()