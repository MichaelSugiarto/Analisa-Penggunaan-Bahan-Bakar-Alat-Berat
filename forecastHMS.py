import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error
import pmdarima as pm
import warnings

warnings.filterwarnings("ignore")

# ==========================================
# FASE 1 & 2: EKSTRAKSI DAN DATA ENGINEERING
# ==========================================
def load_and_melt_excel(file_path):
    print(f"Mengekstrak data dari: {file_path}")
    all_data = []
    xls = pd.read_excel(file_path, sheet_name=None, header=[0, 1, 2])
    for sheet_name, df in xls.items():
        df = df.set_index(df.columns[0]) 
        df.index.name = 'TANGGAL'
        df.columns = df.columns.droplevel(1) 
        df_stacked = df.stack(level=0).reset_index()
        df_stacked.rename(columns={'level_1': 'EQUIP NAME'}, inplace=True)
        all_data.append(df_stacked)
    df_final = pd.concat(all_data, ignore_index=True)
    df_final['TANGGAL'] = pd.to_datetime(df_final['TANGGAL'], format='%d-%m-%Y', errors='coerce')
    df_final = df_final.dropna(subset=['TANGGAL'])
    return df_final

def prepare_data():
    df_2024 = load_and_melt_excel('BBM AAB 2024 & Des 2025.xlsx')
    df_2025 = load_and_melt_excel('BBM AAB.xlsx')
    df_2026 = load_and_melt_excel('BBM AAB Jan-Mar 2026.xlsx')
    
    df_all = pd.concat([df_2024, df_2025, df_2026], ignore_index=True)
    df_all = df_all.sort_values(by=['EQUIP NAME', 'TANGGAL'])
    
    df_all['HM_Clean'] = pd.to_numeric(df_all['HM'], errors='coerce').replace(0, np.nan)
    df_all['HM_Clean'] = df_all.groupby('EQUIP NAME')['HM_Clean'].ffill().fillna(0)
    
    df_all['Delta_HM'] = df_all.groupby('EQUIP NAME')['HM_Clean'].diff().fillna(0)
    df_all.loc[df_all['Delta_HM'] < 0, 'Delta_HM'] = 0
    df_all.loc[df_all['Delta_HM'] > 100, 'Delta_HM'] = 0 
    
    df_all['LITER_Clean'] = pd.to_numeric(df_all['LITER'], errors='coerce').fillna(0)
    
    df_all['TAHUN_BULAN'] = df_all['TANGGAL'].dt.to_period('M')
    agg_data = df_all.groupby(['EQUIP NAME', 'TAHUN_BULAN']).agg({
        'Delta_HM': 'sum',
        'LITER_Clean': 'sum'
    }).reset_index()
    agg_data.rename(columns={'Delta_HM': 'HM', 'LITER_Clean': 'LITER'}, inplace=True)
    
    train_agg = agg_data[agg_data['TAHUN_BULAN'] <= '2025-12']
    test_agg = agg_data[agg_data['TAHUN_BULAN'] >= '2026-01']
    
    df_eksogen = pd.read_excel('Eksogen Hari Kerja Efektif.xlsx', sheet_name='Sheet1')
    df_eksogen['TAHUN_BULAN'] = pd.to_datetime(df_eksogen['Tahun-Bulan']).dt.to_period('M')
    df_eksogen = df_eksogen[['TAHUN_BULAN', 'Hari Kerja Efektif (With Sabtu)']]
    df_eksogen.set_index('TAHUN_BULAN', inplace=True)
    
    return train_agg, test_agg, df_eksogen

# ==========================================
# FASE 3: PIPELINE FORECASTING & MACHINE LEARNING
# ==========================================
def apply_seasonal_fallback(pred_array, df_train_capped, test_periods):
    fallback_preds = []
    for period in test_periods:
        target_month = period.month
        hist_month = df_train_capped[df_train_capped.index.month == target_month]
        if len(hist_month) > 0:
            fallback_preds.append(hist_month.mean())
        else:
            fallback_preds.append(df_train_capped.mean()) 
            
    fallback_preds = np.array(fallback_preds)
    hist_max = df_train_capped.max()
    hist_mean = df_train_capped.mean()
    
    if np.std(pred_array) < 1 or (hist_max > 0 and np.mean(pred_array) > 1.5 * hist_max) or (hist_mean == 0 and np.mean(pred_array) > 10):
        return fallback_preds 
    return pred_array

def run_forecast_pipeline():
    train_agg, test_agg, exog_data = prepare_data()
    list_unit = train_agg['EQUIP NAME'].unique()
    
    results_sarima = []
    results_sarimax = []
    error_sarima_list = []
    error_sarimax_list = []
    excluded_units_list = []
    
    print("\nMemulai training: Pure Seasonal Auto-ARIMA & Cumulative Regression...")
    
    for unit in list_unit:
        try:
            df_unit_train = train_agg[train_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
            df_unit_test = test_agg[test_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
            
            # --- 1. SMART FILTER ---
            # SYARAT WAJIB PURE SEASONAL: Data harus genap 24 bulan atau lebih
            if len(df_unit_train) < 24:
                excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': 'Data historis latih kurang dari 24 bulan (Syarat wajib Pure Seasonal).'})
                continue
            if df_unit_test.empty:
                excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': 'Tidak ada data aktual di tahun 2026.'})
                continue
            if df_unit_train['LITER'].sum() == 0:
                excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': 'Total historis LITER BBM adalah 0.'})
                continue
            last_3_months = df_unit_train.tail(3)
            if len(last_3_months) == 3:
                if (last_3_months['HM'].sum() == 0) or (last_3_months['LITER'].sum() == 0):
                    excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': 'Alat mati suri (HM/LITER 0) berturut-turut di akhir 2025.'})
                    continue
            if df_unit_test['HM'].sum() == 0 and df_unit_test['LITER'].sum() == 0:
                excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': 'Alat tidak digunakan sama sekali (HM/LITER 0) pada periode testing 2026.'})
                continue
                
            # REGIME SHIFT DETECTOR DIHILANGKAN SECARA SENGAJA UNTUK PURE SEASONAL
                    
            exog_train = exog_data.loc[df_unit_train.index]
            exog_test = exog_data.loc[df_unit_test.index]
            
            # --- OUTLIER CAPPING ---
            Q1 = df_unit_train['HM'].quantile(0.25)
            Q3 = df_unit_train['HM'].quantile(0.75)
            IQR = Q3 - Q1
            upper_bound = Q3 + 1.5 * IQR
            df_unit_train['HM_Capped'] = np.minimum(df_unit_train['HM'], upper_bound) if upper_bound > 0 else df_unit_train['HM']
            steps_ahead = len(df_unit_test)
            test_periods = df_unit_test.index
            
            # --- 3. TAHAP 1: PREDIKSI HM DENGAN PURE SEASONAL AUTO-ARIMA ---
            # Model SARIMA (Pure Seasonal Forced)
            try:
                model_sarima = pm.auto_arima(df_unit_train['HM_Capped'], X=None, seasonal=True, m=12, stepwise=True, suppress_warnings=True, error_action="ignore")
                raw_pred_sarima = np.maximum(0, model_sarima.predict(n_periods=steps_ahead))
                pred_hm_sarima = apply_seasonal_fallback(raw_pred_sarima.values, df_unit_train['HM_Capped'], test_periods)
                
                model_sarimax = pm.auto_arima(df_unit_train['HM_Capped'], X=exog_train, seasonal=True, m=12, stepwise=True, suppress_warnings=True, error_action="ignore")
                raw_pred_sarimax = np.maximum(0, model_sarimax.predict(n_periods=steps_ahead, X=exog_test))
                pred_hm_sarimax = apply_seasonal_fallback(raw_pred_sarimax.values, df_unit_train['HM_Capped'], test_periods)
            except Exception as e:
                # Jika secara matematis Auto-ARIMA musiman gagal (misal data terlalu konstan), lempar ke exclude
                excluded_units_list.append({'EQUIP NAME': unit, 'Alasan': f'Gagal kalkulasi Seasonal Auto-ARIMA: Datanya tidak mendukung pola musiman.'})
                continue
            
            # --- 4. TAHAP 2: CUMULATIVE REGRESSION ---
            cum_hm = df_unit_train['HM'].cumsum().values.reshape(-1, 1)
            cum_liter = df_unit_train['LITER'].cumsum().values
            true_ratio = 0
            if np.sum(cum_hm) > 0:
                lr = LinearRegression(fit_intercept=False, positive=True)
                lr.fit(cum_hm, cum_liter)
                true_ratio = lr.coef_[0]
            
            pred_liter_sarima = pred_hm_sarima * true_ratio
            pred_liter_sarimax = pred_hm_sarimax * true_ratio
            aktual_liter = df_unit_test['LITER'].values
            aktual_hm = df_unit_test['HM'].values
            
            rmse_sarima = np.sqrt(mean_squared_error(aktual_liter, pred_liter_sarima))
            rmse_sarimax = np.sqrt(mean_squared_error(aktual_liter, pred_liter_sarimax))
            error_sarima_list.append(rmse_sarima)
            error_sarimax_list.append(rmse_sarimax)
            
            for i, period in enumerate(df_unit_test.index):
                sel_hm_sa = pred_hm_sarima[i] - aktual_hm[i]
                sel_lit_sa = pred_liter_sarima[i] - aktual_liter[i]
                results_sarima.append({
                    'EQUIP NAME': unit, 'Bulan': str(period),
                    'Aktual_HM': round(aktual_hm[i], 2), 'Prediksi_HM': round(pred_hm_sarima[i], 2),
                    'Selisih_HM': round(sel_hm_sa, 2), 'Persentase_Selisih_HM (%)': round((sel_hm_sa / aktual_hm[i] * 100) if aktual_hm[i] != 0 else 0, 2),
                    'Rasio_Sejati (L/HM)': round(true_ratio, 2),
                    'Aktual_LITER': round(aktual_liter[i], 2), 'Prediksi_LITER': round(pred_liter_sarima[i], 2),
                    'Selisih_LITER': round(sel_lit_sa, 2), 'Persentase_Selisih_LITER (%)': round((sel_lit_sa / aktual_liter[i] * 100) if aktual_liter[i] != 0 else 0, 2)
                })
                
                sel_hm_sx = pred_hm_sarimax[i] - aktual_hm[i]
                sel_lit_sx = pred_liter_sarimax[i] - aktual_liter[i]
                results_sarimax.append({
                    'EQUIP NAME': unit, 'Bulan': str(period),
                    'Aktual_HM': round(aktual_hm[i], 2), 'Prediksi_HM': round(pred_hm_sarimax[i], 2),
                    'Selisih_HM': round(sel_hm_sx, 2), 'Persentase_Selisih_HM (%)': round((sel_hm_sx / aktual_hm[i] * 100) if aktual_hm[i] != 0 else 0, 2),
                    'Rasio_Sejati (L/HM)': round(true_ratio, 2),
                    'Aktual_LITER': round(aktual_liter[i], 2), 'Prediksi_LITER': round(pred_liter_sarimax[i], 2),
                    'Selisih_LITER': round(sel_lit_sx, 2), 'Persentase_Selisih_LITER (%)': round((sel_lit_sx / aktual_liter[i] * 100) if aktual_liter[i] != 0 else 0, 2)
                })
        except Exception as e:
            continue

    if not results_sarima:
        print("PERINGATAN: Semua data tersaring (di-exclude).")
        # Simpan file yang dikecualikan agar user bisa lihat alasannya
        with pd.ExcelWriter('Hasil_Forecast_Final.xlsx') as writer:
            pd.DataFrame(excluded_units_list).to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)
        return
        
    avg_rmse_sarima = np.mean(error_sarima_list) if error_sarima_list else 0
    avg_rmse_sarimax = np.mean(error_sarimax_list) if error_sarimax_list else 0
    best_model_name = "SARIMAX (Pure Seasonal + Cumulative)" if avg_rmse_sarimax < avg_rmse_sarima else "SARIMA (Pure Seasonal + Cumulative)"
    final_results = results_sarimax if "SARIMAX" in best_model_name else results_sarima
        
    df_hasil = pd.DataFrame(final_results)
    df_hasil.insert(2, 'Model_Terpilih_Global', best_model_name)
    
    unit_metrics = []
    for unit in df_hasil['EQUIP NAME'].unique():
        sub = df_hasil[df_hasil['EQUIP NAME'] == unit]
        rmse = np.sqrt(mean_squared_error(sub['Aktual_LITER'], sub['Prediksi_LITER']))
        sum_aktual = sub['Aktual_LITER'].sum()
        wmape = (sub['Aktual_LITER'] - sub['Prediksi_LITER']).abs().sum() / sum_aktual if sum_aktual > 0 else 0
        unit_metrics.append({'EQUIP NAME': unit, 'RMSE_Keseluruhan': round(rmse, 2), 'WMAPE_Keseluruhan (%)': round(wmape * 100, 2)})
        
    df_hasil = df_hasil.merge(pd.DataFrame(unit_metrics), on='EQUIP NAME', how='left')
    
    with pd.ExcelWriter('Hasil_Forecast_Final.xlsx') as writer:
        df_hasil.to_excel(writer, sheet_name='Hasil_Forecast', index=False)
        pd.DataFrame(excluded_units_list).to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)
    
    print("\nProses Selesai! File 'Hasil_Forecast_Final.xlsx' berhasil diperbarui.")

if __name__ == "__main__":
    run_forecast_pipeline()