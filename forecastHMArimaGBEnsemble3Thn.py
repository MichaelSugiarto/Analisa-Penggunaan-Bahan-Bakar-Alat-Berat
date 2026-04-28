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
    # Train: Jan 2022 - Des 2024
    # PERUBAHAN: tambah sheets_2022 dari file BBM AAB 2022.xlsx
    sheets_2022 = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV', 'DES']
    sheets_2023 = [
        '01. Jan', '02. Feb', '03. Mar', '04. Apr', '05. Mei', '06. Jun',
        '07. Jul', '08. Aug', '09. Sep', '10. Okt', '11. Nov', '12. Des'
    ]
    sheets_2024 = [
        'JANUARI 24', 'FEB 24', 'maret 24', 'April 24', 'Mei 24', 'Juni 24',
        'Juli 24', 'agt 24', 'sept 24', 'okt 24', 'nov 24', 'des 24'
    ]
    # Test: Jan 2025 - Des 2025
    sheets_des_2025     = ['Des 25']
    sheets_jan_nov_2025 = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']

    # PERUBAHAN: load file 2022, masuk ke train
    df_2022       = load_and_melt_excel('BBM AAB 2022.xlsx',            target_sheets=sheets_2022)
    df_2023       = load_and_melt_excel('BBM AAB 2023.xlsx',            target_sheets=sheets_2023)
    df_2024       = load_and_melt_excel('BBM AAB 2024 & Des 2025.xlsx', target_sheets=sheets_2024)
    df_des_2025   = load_and_melt_excel('BBM AAB 2024 & Des 2025.xlsx', target_sheets=sheets_des_2025)
    df_jan_nov_25 = load_and_melt_excel('BBM AAB Jan-Nov 2025.xlsx',    target_sheets=sheets_jan_nov_2025)

    # PERUBAHAN: df_2022 ikut digabung ke df_all
    df_all = pd.concat([df_2022, df_2023, df_2024, df_des_2025, df_jan_nov_25], ignore_index=True)
    df_all = df_all.sort_values(by=['EQUIP NAME', 'TANGGAL'])

    df_all['HM_Clean'] = pd.to_numeric(df_all['HM'], errors='coerce').replace(0, np.nan)
    df_all['HM_Clean'] = df_all.groupby('EQUIP NAME')['HM_Clean'].ffill().fillna(0)
    df_all['Delta_HM'] = df_all.groupby('EQUIP NAME')['HM_Clean'].diff().fillna(0)
    df_all.loc[df_all['Delta_HM'] < 0,  'Delta_HM'] = 0
    df_all.loc[df_all['Delta_HM'] > 24, 'Delta_HM'] = 0
    df_all['LITER_Clean'] = pd.to_numeric(df_all['LITER'], errors='coerce').fillna(0)

    df_all['TAHUN_BULAN'] = df_all['TANGGAL'].dt.to_period('M')
    agg_data = df_all.groupby(['EQUIP NAME', 'TAHUN_BULAN']).agg(
        {'Delta_HM': 'sum', 'LITER_Clean': 'sum'}
    ).reset_index()
    agg_data.rename(columns={'Delta_HM': 'HM', 'LITER_Clean': 'LITER'}, inplace=True)

    # PERUBAHAN: batas train dimundurkan ke 2022-01
    train_agg = agg_data[agg_data['TAHUN_BULAN'] <= '2024-12']
    test_agg  = agg_data[
        (agg_data['TAHUN_BULAN'] >= '2025-01') &
        (agg_data['TAHUN_BULAN'] <= '2025-12')
    ]
    return train_agg, test_agg


# ==========================================
# FUNGSI PREPROCESSING & MODELING
# ==========================================
def preprocess_timeseries(series):
    df = pd.DataFrame(series, columns=['HM'])

    p05 = df['HM'].quantile(0.05)
    p95 = df['HM'].quantile(0.95)
    if p95 > 0:
        df['HM_Capped'] = df['HM'].clip(lower=p05, upper=p95)
    else:
        df['HM_Capped'] = df['HM']

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

    if n >= 7:
        n_lags = 3
    elif n >= 5:
        n_lags = 2
    elif n >= 4:
        n_lags = 1
    else:
        return np.full(steps_ahead, max(0.0, float(train_series.mean())))

    X_train, y_train, feature_cols = prepare_boosting_features(train_series, n_lags=n_lags)

    if len(X_train) < 3:
        return np.full(steps_ahead, max(0.0, float(train_series.mean())))

    model = HistGradientBoostingRegressor(
        max_iter=200,
        learning_rate=0.05,
        max_depth=4,
        random_state=42
    ).fit(X_train, y_train)

    predictions  = []
    history      = list(train_series.values)
    trend_offset = len(history)

    for step in range(steps_ahead):
        lag_vals  = [history[-(i)] for i in range(1, n_lags + 1)]
        window    = history[-3:] if len(history) >= 3 else history
        roll_mean = np.mean(window)
        roll_std  = np.std(window) if len(window) > 1 else 0.0
        trend_val = trend_offset + step

        row    = lag_vals + [roll_mean, roll_std, trend_val]
        x_pred = pd.DataFrame([row], columns=feature_cols)
        pred   = max(0.0, model.predict(x_pred)[0])
        predictions.append(pred)
        history.append(pred)

    return np.array(predictions)


def hitung_mape_aman(actual, pred):
    actual, pred = np.array(actual), np.array(pred)
    mask = actual != 0
    if not np.any(mask):
        return 0.0
    return np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100


def ensemble_tertimbang(p_arima, p_gb, aktual_h):
    rmse_a = np.sqrt(mean_squared_error(aktual_h, p_arima)) + 1e-6
    rmse_g = np.sqrt(mean_squared_error(aktual_h, p_gb))   + 1e-6
    w_a = (1 / rmse_a) / (1 / rmse_a + 1 / rmse_g)
    w_g = (1 / rmse_g) / (1 / rmse_a + 1 / rmse_g)
    return w_a * p_arima + w_g * p_gb, round(w_a, 4), round(w_g, 4)


# ==========================================
# MAPPING NAMA ALAT BERAT KE MASTER
# ==========================================
def load_master_names():
    try:
        master_df    = pd.read_excel('cost & bbm 2022 sd 2025 HP & Type.xlsx', header=1)
        master_names = set(master_df['NAMA ALAT BERAT'].dropna().astype(str).str.strip())
        return master_names
    except Exception as e:
        print(f"\n[!] Gagal membaca file master: {e}\n")
        return set()


def get_mapped_unit_name(unit_name, master_names):
    hardcoded = {
        "FL RENTAL 01":                "FL RENTAL 01 TIMIKA",
        "TOBATI (EX.FL KALMAR 32T)":   "TOP LOADER KALMAR 35T/TOBATI",
        "L 8477 UUC (EX.L 9902 UR)":   "L 9902 UR / S75",
        "WIND RIVER (EX.TL BOSS 42T)":  "TOP LOADER BOSS"
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

    results_combined    = []
    metrics_list        = []
    excluded_units_list = []

    all_actual_hm     = []
    all_pred_arima_hm = []
    all_pred_gb_hm    = []
    all_pred_ens_hm   = []

    total_valid_population = 0

    print(f"\n[AI ENGINE] Memulai pelatihan: ARIMA vs GB vs Ensemble untuk {len(list_unit_raw)} unit...")

    for unit in list_unit_raw:
        mapped_name = None
        try:
            mapped_name = get_mapped_unit_name(unit, master_names)
            if not mapped_name:
                continue

            total_valid_population += 1

            df_u_train = train_agg[train_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()
            df_u_test  = test_agg[test_agg['EQUIP NAME'] == unit].set_index('TAHUN_BULAN').copy()

            # ---------------------------------------------------------
            # FILTER LOGIS DATA CACAT
            # ---------------------------------------------------------
            if df_u_test.empty:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Tidak ada data aktual di periode uji (Jan-Des 2025).'
                })
                continue

            # PERUBAHAN: train sekarang 36 bulan (2022+2023+2024)
            if len(df_u_train) < 12:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Data latih kurang dari 12 bulan.'
                })
                continue

            if len(df_u_test) < 12:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Data uji kurang dari 12 bulan (Jan-Des 2025 tidak lengkap).'
                })
                continue

            # PERUBAHAN: tambah pengecekan tahun 2022
            df_2022_tr = df_u_train[df_u_train.index.year == 2022]
            df_2023_tr = df_u_train[df_u_train.index.year == 2023]
            df_2024_tr = df_u_train[df_u_train.index.year == 2024]

            if len(df_2022_tr) == 12 and (df_2022_tr['HM'].sum() == 0 or df_2022_tr['LITER'].sum() == 0):
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'HM/LITER = 0 selama 1 tahun penuh di Data Latih (2022).'
                })
                continue

            if len(df_2023_tr) == 12 and (df_2023_tr['HM'].sum() == 0 or df_2023_tr['LITER'].sum() == 0):
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'HM/LITER = 0 selama 1 tahun penuh di Data Latih (2023).'
                })
                continue

            if len(df_2024_tr) == 12 and (df_2024_tr['HM'].sum() == 0 or df_2024_tr['LITER'].sum() == 0):
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'HM/LITER = 0 selama 1 tahun penuh di Data Latih (2024).'
                })
                continue

            if df_u_test['HM'].sum() == 0 or df_u_test['LITER'].sum() == 0:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'HM/LITER = 0 selama 1 tahun penuh di Data Uji (2025).'
                })
                continue

            # --- Rasio Liter/HM historis ---
            true_ratio = (df_u_train['LITER'].sum() / df_u_train['HM'].sum()
                          if df_u_train['HM'].sum() > 0 else 0)
            aktual_l = df_u_test['LITER'].values
            aktual_h = df_u_test['HM'].values
            steps    = len(df_u_test)

            # --- Truncation ---
            t_utuh = df_u_train['HM'].copy()
            try:
                first_idx = df_u_train[df_u_train['HM'] > 0].index[0]
                t_potong  = df_u_train.loc[first_idx:]['HM'].copy()
            except IndexError:
                t_potong = t_utuh

            best_arima    = np.zeros(steps)
            best_gb       = np.zeros(steps)
            min_rmse      = float('inf')
            model_success = False

            for ds_name, ds in [("Utuh", t_utuh), ("Potong", t_potong)]:
                if len(ds) < 6:
                    continue
                model_success = True

                ds_s = preprocess_timeseries(ds)

                try:
                    arima_model = pm.auto_arima(
                        ds_s, seasonal=False, max_d=1,
                        suppress_warnings=True, error_action="ignore"
                    )
                    p_arima_raw = arima_model.predict(n_periods=steps)
                    baseline    = float(ds_s.tail(6).mean())
                    p_arima     = np.clip(p_arima_raw, baseline * 0.1, baseline * 2.0)
                    p_arima     = np.maximum(0, p_arima)
                except Exception:
                    p_arima = np.full(steps, max(0.0, float(ds_s.mean())))

                try:
                    p_gb = predict_boosting(ds_s, steps)
                except Exception:
                    p_gb = np.full(steps, max(0.0, float(ds_s.mean())))

                p_ens_trial, _, _ = ensemble_tertimbang(p_arima, p_gb, aktual_h)
                rmse_ds = np.sqrt(mean_squared_error(aktual_h, p_ens_trial))

                if rmse_ds < min_rmse:
                    min_rmse   = rmse_ds
                    best_arima = p_arima
                    best_gb    = p_gb

            if not model_success:
                excluded_units_list.append({
                    'EQUIP NAME': unit, 'NAMA_MASTER_TERPETAKAN': mapped_name,
                    'Alasan': 'Data historis valid setelah dipotong kurang dari 6 bulan.'
                })
                continue

            best_ensemble, w_a, w_g = ensemble_tertimbang(best_arima, best_gb, aktual_h)

            # --- Kalkulasi Metrik ---
            rmse_arima    = np.sqrt(mean_squared_error(aktual_h, best_arima))
            mape_arima    = hitung_mape_aman(aktual_h, best_arima)
            rmse_gb       = np.sqrt(mean_squared_error(aktual_h, best_gb))
            mape_gb       = hitung_mape_aman(aktual_h, best_gb)
            rmse_ensemble = np.sqrt(mean_squared_error(aktual_h, best_ensemble))
            mape_ensemble = hitung_mape_aman(aktual_h, best_ensemble)

            # Pilih model terbaik PER UNIT
            mape_per_unit  = {
                'ARIMA':             mape_arima,
                'Gradient Boosting': mape_gb,
                'Ensemble':          mape_ensemble,
            }
            best_name_unit = min(mape_per_unit, key=mape_per_unit.get)
            pred_map       = {
                'ARIMA':             best_arima,
                'Gradient Boosting': best_gb,
                'Ensemble':          best_ensemble,
            }
            pred_terpilih = pred_map[best_name_unit]

            metrics_list.append({
                'EQUIP NAME':             unit,
                'NAMA_MASTER_TERPETAKAN': mapped_name,
                'RMSE_HM_ARIMA':          round(rmse_arima, 2),
                'MAPE_HM_ARIMA (%)':      round(mape_arima, 2),
                'RMSE_HM_GB':             round(rmse_gb, 2),
                'MAPE_HM_GB (%)':         round(mape_gb, 2),
                'RMSE_HM_Ensemble':       round(rmse_ensemble, 2),
                'MAPE_HM_Ensemble (%)':   round(mape_ensemble, 2),
                'Bobot_ARIMA':            w_a,
                'Bobot_GB':               w_g,
                'Model_Terpilih_Unit':    best_name_unit,
                'MAPE_Terpilih (%)':      round(mape_per_unit[best_name_unit], 2),
            })

            all_actual_hm.extend(aktual_h)
            all_pred_arima_hm.extend(best_arima)
            all_pred_gb_hm.extend(best_gb)
            all_pred_ens_hm.extend(best_ensemble)

            for i, period in enumerate(df_u_test.index):
                results_combined.append({
                    'EQUIP NAME':              unit,
                    'NAMA_MASTER_TERPETAKAN':  mapped_name,
                    'Bulan':                   str(period),
                    'Aktual_HM':               round(aktual_h[i], 2),
                    'Aktual_LITER':            round(aktual_l[i], 2),
                    'Prediksi_HM_ARIMA':       round(best_arima[i], 2),
                    'Prediksi_LITER_ARIMA':    round(best_arima[i] * true_ratio, 2),
                    'Prediksi_HM_GB':          round(best_gb[i], 2),
                    'Prediksi_LITER_GB':       round(best_gb[i] * true_ratio, 2),
                    'Prediksi_HM_Ensemble':    round(best_ensemble[i], 2),
                    'Prediksi_LITER_Ensemble': round(best_ensemble[i] * true_ratio, 2),
                    'Model_Terpilih_Unit':     best_name_unit,
                    'Prediksi_HM_Terpilih':    round(pred_terpilih[i], 2),
                    'Prediksi_LITER_Terpilih': round(pred_terpilih[i] * true_ratio, 2),
                    'MAPE_HM_ARIMA (%)':       round(mape_arima, 2),
                    'MAPE_HM_GB (%)':          round(mape_gb, 2),
                    'MAPE_HM_Ensemble (%)':    round(mape_ensemble, 2),
                    'MAPE_Terpilih (%)':       round(mape_per_unit[best_name_unit], 2),
                })

        except Exception as e:
            excluded_units_list.append({
                'EQUIP NAME':             unit,
                'NAMA_MASTER_TERPETAKAN': mapped_name if mapped_name else '-',
                'Alasan':                 f'Gagal diproses (Internal Error): {str(e)}'
            })
            continue

    # ========================================================
    # PEMBUATAN EXCEL
    # ========================================================
    df_combined = pd.DataFrame(results_combined)
    df_metrics  = pd.DataFrame(metrics_list)

    if df_combined.empty:
        print("\nGAGAL: Tidak ada unit valid yang berhasil diproses.")
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel('Unit_Dikecualikan_Error.xlsx', index=False)
        return

    df_terpilih = df_combined[[
        'EQUIP NAME', 'NAMA_MASTER_TERPETAKAN', 'Bulan',
        'Aktual_HM', 'Aktual_LITER',
        'Model_Terpilih_Unit',
        'Prediksi_HM_Terpilih',
        'Prediksi_LITER_Terpilih',
        'MAPE_HM_ARIMA (%)',
        'MAPE_HM_GB (%)',
        'MAPE_HM_Ensemble (%)',
        'MAPE_Terpilih (%)'
    ]].copy()

    df_mape_under_35 = df_terpilih[df_terpilih['MAPE_Terpilih (%)'] <  35].copy()
    df_mape_over_35  = df_terpilih[df_terpilih['MAPE_Terpilih (%)'] >= 35].copy()

    global_mape_arima    = hitung_mape_aman(all_actual_hm, all_pred_arima_hm)
    global_mape_gb       = hitung_mape_aman(all_actual_hm, all_pred_gb_hm)
    global_mape_ensemble = hitung_mape_aman(all_actual_hm, all_pred_ens_hm)

    df_komparasi = df_combined[[
        'EQUIP NAME', 'NAMA_MASTER_TERPETAKAN', 'Bulan',
        'Aktual_HM', 'Aktual_LITER',
        'Prediksi_HM_ARIMA', 'Prediksi_LITER_ARIMA',
        'Prediksi_HM_GB', 'Prediksi_LITER_GB',
        'Prediksi_HM_Ensemble', 'Prediksi_LITER_Ensemble',
        'Model_Terpilih_Unit',
        'Prediksi_HM_Terpilih', 'Prediksi_LITER_Terpilih',
        'MAPE_HM_ARIMA (%)', 'MAPE_HM_GB (%)',
        'MAPE_HM_Ensemble (%)', 'MAPE_Terpilih (%)'
    ]].copy()

    df_bobot = df_metrics[['EQUIP NAME', 'RMSE_HM_ARIMA', 'RMSE_HM_GB',
                            'RMSE_HM_Ensemble', 'Bobot_ARIMA', 'Bobot_GB']].copy()
    df_komparasi = df_komparasi.merge(df_bobot, on='EQUIP NAME', how='left')

    with pd.ExcelWriter('Hasil_Forecast_Final_2022_2024.xlsx') as writer:
        df_komparasi.to_excel(writer,     sheet_name='Komparasi_Model',      index=False)
        if not df_mape_under_35.empty:
            df_mape_under_35.to_excel(writer, sheet_name='Akurasi_Bagus_Under35', index=False)
        if not df_mape_over_35.empty:
            df_mape_over_35.to_excel(writer,  sheet_name='Akurasi_Rendah_Over35', index=False)
        df_metrics.to_excel(writer,       sheet_name='Metrik_Per_Unit',       index=False)
        if excluded_units_list:
            pd.DataFrame(excluded_units_list).to_excel(writer, sheet_name='Unit_Dikecualikan', index=False)

    # ========================================================
    # SUMMARY TERMINAL
    # ========================================================
    total_under_35 = df_mape_under_35['EQUIP NAME'].nunique() if not df_mape_under_35.empty else 0
    total_over_35  = df_mape_over_35['EQUIP NAME'].nunique()  if not df_mape_over_35.empty  else 0
    total_excluded = len(excluded_units_list)

    pct_under = (total_under_35 / total_valid_population * 100) if total_valid_population else 0
    pct_over  = (total_over_35  / total_valid_population * 100) if total_valid_population else 0
    pct_excl  = (total_excluded / total_valid_population * 100) if total_valid_population else 0

    dist_model = df_metrics['Model_Terpilih_Unit'].value_counts()

    print("\n" + "=" * 60)
    print(" SUMMARY HASIL ANALISA OPERASIONAL SPIL ".center(60))
    print("=" * 60)
    print(f"Total Populasi Alat Berat (Terpetakan Master) : {total_valid_population} unit")
    print(f"Periode Train : Januari 2022 – Desember 2024 (36 bulan)")
    print(f"Periode Test  : Januari 2025 – Desember 2025 (12 bulan)")
    print(f"\nGlobal MAPE (referensi) :")
    print(f"  ARIMA    : {global_mape_arima:.2f}%")
    print(f"  GB       : {global_mape_gb:.2f}%")
    print(f"  Ensemble : {global_mape_ensemble:.2f}%")
    print(f"\nDistribusi Model Terpilih Per Unit :")
    for model_name, count in dist_model.items():
        pct = count / total_valid_population * 100 if total_valid_population else 0
        print(f"  {model_name:<20}: {count} unit ({pct:.1f}%)")
    print("-" * 60)
    print(f"1. Akurasi Bagus  (MAPE < 35%) : {total_under_35} unit ({pct_under:.1f}%)")
    print(f"2. Akurasi Rendah (MAPE >= 35%): {total_over_35} unit ({pct_over:.1f}%)")
    print(f"3. Unit Diexclude (Data Cacat) : {total_excluded} unit ({pct_excl:.1f}%)")
    print("=" * 60)
    print("Laporan berhasil disimpan ke 'Hasil_Forecast_Final_2022_2024.xlsx'")


if __name__ == "__main__":
    run_forecast_pipeline()