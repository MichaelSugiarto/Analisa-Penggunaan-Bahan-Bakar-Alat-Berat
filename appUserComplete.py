import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import warnings

warnings.filterwarnings('ignore')

# ==============================================================================
# 1. KONFIGURASI HALAMAN
# ==============================================================================
st.set_page_config(
    page_title="BBM Alat Berat",
    layout="wide"
)

st.title("Dashboard Monitoring Efisiensi BBM Alat Berat")

# ==============================================================================
# 2. SIDEBAR: INPUT DATA MENTAH
# ==============================================================================
#st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2830/2830305.png", width=50)
st.sidebar.title("Input Data Untuk Analisa")
st.sidebar.caption("Program akan memproses data menjadi Laporan Benchmark dan Laporan Tren Bulanan")

master_file = st.sidebar.file_uploader("1. Upload Master Data (cost & bbm 2022 sd 2025 HP & Type.xlsx)", type=['xlsx'])
bbm_file = st.sidebar.file_uploader("2. Upload Transaksi BBM Mentah (BBM AAB.xlsx)", type=['xlsx'])

mulai_proses = st.sidebar.button("Mulai Proses Analisa", type="primary", use_container_width=True)

st.sidebar.markdown("---")

# ==============================================================================
# 3. FUNGSI PEMROSESAN DATA (GABUNGAN JUPYTER + STREAMLIT)
# ==============================================================================
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    name = name.replace("FORKLIFT", "FORKLIF")
    return re.sub(r'[^A-Z0-9]', '', name)

@st.cache_data(show_spinner=False)
def process_raw_data(file_master, file_bbm):
    master_data_map = {} 
    master_keys_set = set()

    # --- A. BACA MASTER DATA ---
    df_map = pd.read_excel(file_master, sheet_name='Sheet2', header=1)
    
    col_name = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
    col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
    col_type = next((c for c in df_map.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
    col_hp = next((c for c in df_map.columns if any(k == str(c).upper() for k in ['HP', 'HORSE POWER'])), None)
    col_cap = next((c for c in df_map.columns if any(k in str(c).upper() for k in ['CAP', 'KAPASITAS'])), None)
    col_loc = 'DES 2025' if 'DES 2025' in df_map.columns else df_map.columns[2]

    rename_dict = {
        col_name: 'Unit_Original', 
        col_jenis: 'Jenis_Alat', 
        col_hp: 'Horse_Power', 
        col_cap: 'Capacity_Raw', 
        col_loc: 'Lokasi'
    }
    if col_type:
        rename_dict[col_type] = 'Type_Merk'

    df_map.rename(columns=rename_dict, inplace=True)
    
    if 'Type_Merk' not in df_map.columns:
        df_map['Type_Merk'] = "-"

    df_map.dropna(subset=['Unit_Original'], inplace=True)
    df_map['Unit_ID'] = df_map['Unit_Original'].apply(clean_unit_name)
    df_map = df_map[~df_map['Unit_Original'].astype(str).str.upper().str.contains('DUMMY', na=False)]
    df_map = df_map[~df_map['Unit_Original'].astype(str).str.upper().str.contains('FALCON', na=False)]
    df_map['Horse_Power'] = pd.to_numeric(df_map['Horse_Power'], errors='coerce').fillna(0)

    for _, row in df_map.iterrows():
        clean_id = row['Unit_ID']
        if clean_id:
            u_name = str(row['Unit_Original']).strip().upper()
            cap_val = 0
            
            if clean_id == clean_unit_name("L 9025 US"):
                cap_val = 40
            else:
                try:
                    raw_cap = str(row['Capacity_Raw'])
                    match = re.search(r"(\d+(\.\d+)?)", raw_cap)
                    if match:
                        val_float = float(match.group(1))
                        cap_val = int(val_float + 0.5)
                except: pass
                
                if cap_val == 0:
                    try:
                        match_name = re.search(r"(\d+(\.\d+)?)\s*(T|TON|K)", u_name)
                        if match_name:
                            val_float = float(match_name.group(1))
                            cap_val = int(val_float + 0.5)
                    except: pass

            # [UPDATE] Fix Typo Type/Merk secara spesifik
            t_merk = str(row['Type_Merk']).strip().upper()
            
            # Ganti MITSUBHISI (Typo H) menjadi MITSUBISHI
            t_merk = t_merk.replace("MITSUBHISI", "MITSUBISHI")
            
            # Ganti ITSUBISHI (Kurang M) menjadi MITSUBISHI
            if t_merk == "ITSUBISHI":
                t_merk = "MITSUBISHI"
            elif t_merk.startswith("ITSUBISHI "):
                t_merk = "MITSUBISHI " + t_merk[10:]
            elif " ITSUBISHI" in t_merk:
                t_merk = t_merk.replace(" ITSUBISHI", " MITSUBISHI")

            master_data_map[clean_id] = {
                'Unit_Name': row['Unit_Original'],
                'Jenis_Alat': row['Jenis_Alat'],
                'Type_Merk': t_merk,
                'Horse_Power': row['Horse_Power'], 
                'Capacity': cap_val,
                'Lokasi': row['Lokasi']
            }
            master_keys_set.add(clean_id)

    # --- B. BACA DATA TRANSAKSI BBM MENTAH ---
    raw_data_list = []
    xls = pd.ExcelFile(file_bbm)
    target_sheets = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV', 'DES']
    
    for sheet in target_sheets:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            unit_names_row = df.iloc[0].ffill()
            headers = df.iloc[2]
            dates = df.iloc[3:, 0]
            
            for col in range(1, df.shape[1]):
                header_str = str(headers[col]).strip().upper()
                if header_str in ['HM', 'LITER', 'KELUAR', 'PEMAKAIAN']:
                    raw_unit_name = str(unit_names_row[col]).strip().upper()
                    if raw_unit_name == "" or "UNNAMED" in raw_unit_name or "TOTAL" in raw_unit_name: continue
                    if raw_unit_name.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL')): continue
                    
                    clean_trx_id = clean_unit_name(raw_unit_name)
                    matched_id = None
                    
                    # Manual Mapping
                    if "FL RENTAL 01" in raw_unit_name and "TIMIKA" not in raw_unit_name:
                        matched_id = clean_unit_name("FL RENTAL 01 TIMIKA") if clean_unit_name("FL RENTAL 01 TIMIKA") in master_data_map else None
                    elif "TOBATI" in raw_unit_name and "KALMAR 32T" in raw_unit_name:
                        matched_id = clean_unit_name("TOP LOADER KALMAR 35T/TOBATI") if clean_unit_name("TOP LOADER KALMAR 35T/TOBATI") in master_data_map else None
                    elif "L 8477 UUC" in raw_unit_name:
                        matched_id = clean_unit_name("L 9902 UR / S75") if clean_unit_name("L 9902 UR / S75") in master_data_map else None
                    elif "L 9054 UT" in raw_unit_name:
                        matched_id = clean_unit_name("L 9054 UT") if clean_unit_name("L 9054 UT") in master_data_map else None
                    
                    # Auto Mapping
                    if not matched_id and clean_trx_id in master_data_map: matched_id = clean_trx_id
                    if not matched_id and "EX." in raw_unit_name:
                        try:
                            clean_after = clean_unit_name(raw_unit_name.split("EX.")[-1].replace(")", "").strip())
                            if clean_after in master_data_map: matched_id = clean_after
                            elif clean_after:
                                for k in master_keys_set:
                                    if clean_after in k: matched_id = k; break
                        except: pass
                    if not matched_id and " (" in raw_unit_name:
                        try:
                            clean_before = clean_unit_name(raw_unit_name.split(" (")[0].strip())
                            if clean_before in master_data_map: matched_id = clean_before
                        except: pass

                    # Ekstrak Data
                    if matched_id:
                        metric_type = 'HM' if header_str == 'HM' else 'LITER'
                        vals = pd.to_numeric(df.iloc[3:, col], errors='coerce')
                        info = master_data_map[matched_id]
                        temp_df = pd.DataFrame({
                            'Date': dates, 'Unit_Name': info['Unit_Name'], 
                            'Jenis_Alat': info['Jenis_Alat'], 
                            'Type_Merk': info['Type_Merk'],
                            'Horse_Power': info['Horse_Power'],
                            'Capacity': info['Capacity'], 'Lokasi': info['Lokasi'],
                            'Metric': metric_type, 'Value': vals
                        })
                        temp_df.dropna(subset=['Value', 'Date'], inplace=True)
                        if not temp_df.empty: raw_data_list.append(temp_df)

    # --- C. KALKULASI DELTA HM & PIVOT ---
    if not raw_data_list: return None, None, None
    df_all = pd.concat(raw_data_list, ignore_index=True)
    df_all['Date'] = pd.to_datetime(df_all['Date'], dayfirst=True, errors='coerce')
    df_all.dropna(subset=['Date'], inplace=True)
    
    # Simpan data mentah untuk Grafik Tren Bulanan
    df_trend_raw = df_all.copy()

    # Pivot Total
    df_pivot = df_all.pivot_table(index=['Unit_Name', 'Lokasi', 'Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity', 'Date'], columns='Metric', values='Value', aggfunc='sum').reset_index()
    if 'HM' not in df_pivot.columns: df_pivot['HM'] = 0
    if 'LITER' not in df_pivot.columns: df_pivot['LITER'] = 0
    df_pivot['HM'], df_pivot['LITER'] = df_pivot['HM'].fillna(0), df_pivot['LITER'].fillna(0)
    df_pivot.sort_values(by=['Unit_Name', 'Date'], inplace=True)
    
    # Hitung Delta HM
    df_pivot['HM_Clean'] = df_pivot['HM'].replace(0, np.nan).groupby(df_pivot['Unit_Name']).ffill().fillna(0)
    df_pivot['Delta_HM'] = df_pivot.groupby('Unit_Name')['HM_Clean'].diff().fillna(0)
    df_pivot.loc[(df_pivot['Delta_HM'] < 0) | (df_pivot['Delta_HM'] > 100), 'Delta_HM'] = 0 
    
    # --- D. BENCHMARK & STATUS ---
    final_stats = df_pivot.groupby(['Unit_Name', 'Lokasi', 'Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity']).agg({'LITER': 'sum', 'Delta_HM': 'sum'}).reset_index()
    final_stats.rename(columns={'LITER': 'Total_Liter', 'Delta_HM': 'Total_HM_Work'}, inplace=True)
    final_stats['Fuel_Ratio'] = final_stats.apply(lambda row: row['Total_Liter'] / row['Total_HM_Work'] if row['Total_HM_Work'] > 0 else 0, axis=1)
    
    df_valid = final_stats[(final_stats['Total_HM_Work'] > 0) & (final_stats['Total_Liter'] > 0)].copy()
    benchmark_stats = df_valid.groupby('Horse_Power')['Fuel_Ratio'].median().reset_index()
    benchmark_stats.rename(columns={'Fuel_Ratio': 'Group_Benchmark_Median'}, inplace=True)
    
    df_final = pd.merge(final_stats, benchmark_stats, on='Horse_Power', how='left')
    
    def get_status(row):
        if row['Total_HM_Work'] <= 0 or row['Total_Liter'] <= 0: return "INAKTIF"
        return "EFISIEN" if row['Fuel_Ratio'] <= row['Group_Benchmark_Median'] else "BOROS"

    df_final['Performance_Status'] = df_final.apply(get_status, axis=1)
    df_final['Potensi_Pemborosan_Liter'] = 0
    boros_mask = df_final['Performance_Status'] == "BOROS"
    df_final.loc[boros_mask, 'Potensi_Pemborosan_Liter'] = ((df_final.loc[boros_mask, 'Fuel_Ratio'] - df_final.loc[boros_mask, 'Group_Benchmark_Median']) * df_final.loc[boros_mask, 'Total_HM_Work'])
    
    df_final['Fuel_Ratio'] = df_final['Fuel_Ratio'].round(2)
    df_final['Group_Benchmark_Median'] = df_final['Group_Benchmark_Median'].round(2)
    df_final['Potensi_Pemborosan_Liter'] = df_final['Potensi_Pemborosan_Liter'].round(2)

    df_active = df_final[df_final['Performance_Status'] != "INAKTIF"].copy()
    df_inactive = df_final[df_final['Performance_Status'] == "INAKTIF"].copy()

    # --- E. GENERATE DATA TREN BULANAN ---
    df_trend_raw['Month_Year'] = df_trend_raw['Date'].dt.to_period('M').astype(str)
    df_pivot_trend = df_trend_raw.pivot_table(index=['Unit_Name', 'Month_Year', 'Date'], columns='Metric', values='Value', aggfunc='sum').reset_index()
    if 'HM' not in df_pivot_trend.columns: df_pivot_trend['HM'] = 0
    if 'LITER' not in df_pivot_trend.columns: df_pivot_trend['LITER'] = 0
    df_pivot_trend.sort_values(by=['Unit_Name', 'Date'], inplace=True)
    
    df_pivot_trend['HM_Clean'] = df_pivot_trend['HM'].replace(0, np.nan).groupby(df_pivot_trend['Unit_Name']).ffill().fillna(0)
    df_pivot_trend['Delta_HM'] = df_pivot_trend.groupby('Unit_Name')['HM_Clean'].diff().fillna(0)
    df_pivot_trend.loc[(df_pivot_trend['Delta_HM'] < 0) | (df_pivot_trend['Delta_HM'] > 100), 'Delta_HM'] = 0 
    
    trend_monthly = df_pivot_trend.groupby(['Unit_Name', 'Month_Year']).agg({'LITER': 'sum', 'Delta_HM': 'sum'}).reset_index()
    trend_monthly['Fuel_Ratio'] = trend_monthly.apply(lambda r: r['LITER'] / r['Delta_HM'] if r['Delta_HM'] > 0 else 0, axis=1)
    trend_monthly.rename(columns={'Month_Year': 'Bulan'}, inplace=True)

    return df_active, df_inactive, trend_monthly


# ==============================================================================
# JALANKAN PROSES JIKA TOMBOL DITEKAN
# ==============================================================================
if 'df_unit' not in st.session_state:
    st.session_state['df_unit'] = None
    st.session_state['df_inaktif'] = None
    st.session_state['df_trend'] = None

if mulai_proses:
    if master_file and bbm_file:
        with st.spinner("Processing data yang diberikan (estimasi 10-20 detik)..."):
            df_active, df_inactive, df_trend = process_raw_data(master_file, bbm_file)
            st.session_state['df_unit'] = df_active
            st.session_state['df_inaktif'] = df_inactive
            st.session_state['df_trend'] = df_trend
        st.success("Data selesai diproses!")
    else:
        st.error("Upload kedua file terlebih dahulu sebelum memulai proses.")

df_unit = st.session_state['df_unit']
df_inaktif = st.session_state['df_inaktif']
df_trend_global = st.session_state['df_trend']

# --- FUNGSI FORMAT SATUAN (TON/FEET) DENGAN HANDLING ANGKA 0 ---
def format_capacity_with_unit(row):
    cap = row.get('Capacity', 0)
    try:
        cap_float = float(cap)
        if pd.isna(cap_float) or cap_float == 0:
            return "0"  
        cap_val = int(cap_float)
    except:
        return str(cap)

    jenis = str(row.get('Jenis_Alat', '')).upper()
    if jenis in ['CRANE', 'FORKLIFT', 'REACH STACKER', 'SIDE LOADER', 'TOP LOADER']: return f"{cap_val} Ton"
    elif jenis in ['TRONTON', 'TRAILER']: return f"{cap_val} Feet"
    return str(cap_val)

# ==============================================================================
# 4. KONTEN UTAMA DASHBOARD
# ==============================================================================
if df_unit is not None:
    # --- PENCARIAN UNIT ---
    st.subheader("Cari Data Spesifik")
    
    # Dropdown Pemilihan Kategori (Hanya Nama Unit & Horse Power)
    search_category = st.selectbox("Pilih Kategori Pencarian:", ["Nama Unit", "Horse Power"])
    
    # Input Pencarian
    search_keyword = st.text_input(f"Ketik {search_category}:", key="search_keyword", placeholder=f"Cari {search_category}...").upper()
    
    if search_keyword:
        mask_active = pd.Series([False] * len(df_unit))
        mask_inactive = pd.Series([False] * (len(df_inaktif) if df_inaktif is not None else 0))
        valid_search = True

        if search_category == "Nama Unit":
            mask_active = df_unit['Unit_Name'].astype(str).str.contains(search_keyword, na=False)
            if df_inaktif is not None: mask_inactive = df_inaktif['Unit_Name'].astype(str).str.contains(search_keyword, na=False)
        elif search_category == "Horse Power":
            try:
                float(search_keyword) 
                mask_active = df_unit['Horse_Power'].astype(str).str.contains(search_keyword, na=False)
                if df_inaktif is not None: mask_inactive = df_inaktif['Horse_Power'].astype(str).str.contains(search_keyword, na=False)
            except ValueError:
                st.warning("⚠️ Untuk pencarian Horse Power, mohon masukkan angka.")
                valid_search = False

        if valid_search:
            res_active = df_unit[mask_active].copy()
            res_active['Status'] = 'AKTIF'
            
            res_inactive = pd.DataFrame()
            if df_inaktif is not None:
                res_inactive = df_inaktif[mask_inactive].copy()
                res_inactive['Status'] = 'INAKTIF'
                
            res_all = pd.concat([res_active, res_inactive], ignore_index=True)
            
            if not res_all.empty:
                st.info(f"Ditemukan {len(res_all)} Unit:")
                if 'Fuel_Ratio' in res_all.columns: res_all.sort_values(by='Fuel_Ratio', ascending=False, inplace=True)
                
                cols_to_show = ['Unit_Name', 'Jenis_Alat', 'Type_Merk', 'Status', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Group_Benchmark_Median', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']
                for c in cols_to_show:
                    if c not in res_all.columns: res_all[c] = 0 if c in ['Total_Liter', 'Total_HM_Work', 'Fuel_Ratio', 'Potensi_Pemborosan_Liter', 'Group_Benchmark_Median'] else "-"
                
                df_search_display = res_all[cols_to_show].copy()
                
                # Penyesuaian Teks Untuk Unit Inaktif dan Nilai Benchmark
                df_search_display['Capacity'] = df_search_display.apply(format_capacity_with_unit, axis=1)
                
                df_search_display['Group_Benchmark_Median'] = df_search_display['Group_Benchmark_Median'].apply(
                    lambda x: "None" if pd.isna(x) or x == 0 else f"{float(x):.2f}"
                )
                
                def format_fuel_ratio(row):
                    if row['Status'] == 'INAKTIF': return "Tidak Ada Karena Unit Inaktif"
                    return f"{float(row['Fuel_Ratio']):.2f}"

                def format_pemborosan(row):
                    if row['Status'] == 'INAKTIF': return "Tidak Ada Karena Unit Inaktif"
                    return f"{float(row['Potensi_Pemborosan_Liter']):,.0f}"

                def format_status_bbm(row):
                    if row['Status'] == 'INAKTIF': return "UNIT INAKTIF"
                    return row['Performance_Status']

                df_search_display['Fuel_Ratio'] = df_search_display.apply(format_fuel_ratio, axis=1)
                df_search_display['Potensi_Pemborosan_Liter'] = df_search_display.apply(format_pemborosan, axis=1)
                df_search_display['Performance_Status'] = df_search_display.apply(format_status_bbm, axis=1) 
                
                rename_map_search = {'Unit_Name': 'Unit', 'Type_Merk': 'Type/Merk', 'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Group_Benchmark_Median': 'Benchmark', 'Performance_Status': 'Status_BBM', 'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'}
                df_search_display.rename(columns=rename_map_search, inplace=True)

                # Fix konsistensi warna teks pencarian
                def highlight_search(row):
                    status_bbm = str(row['Status_BBM']).upper()
                    if status_bbm == 'EFISIEN':
                        return [f'background-color: #2ca02c; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]
                    elif status_bbm == 'BOROS':
                        return [f'background-color: #d62728; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]
                    else:
                        return ['' for _ in row.index]

                st.dataframe(df_search_display.style.format({'Horse_Power': '{:.0f}', 'Total_Pengisian_BBM': '{:,.0f}', 'Total_Jam_Kerja': '{:,.0f}'}).apply(highlight_search, axis=1))
            else:
                st.warning("Unit Tidak Ditemukan.")

    st.markdown("---")

    # --- Sidebar Filters ---
    st.sidebar.subheader("Filter Dashboard")

    loc_options = ["Semua"] + sorted(df_unit['Lokasi'].unique().tolist())
    selected_loc = st.sidebar.selectbox("Pilih Lokasi:", loc_options)

    type_options = ["Semua"] + sorted(df_unit['Jenis_Alat'].unique().tolist())
    selected_type = st.sidebar.selectbox("Pilih Jenis Alat:", type_options)

    type_merk_options = ["Semua"] + sorted(df_unit['Type_Merk'].astype(str).unique().tolist())
    selected_type_merk = st.sidebar.selectbox("Pilih Type/Merk:", type_merk_options)

    st.sidebar.subheader("Biaya Bahan Bakar")
    harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_alat')

    # --- FILTER FINAL BERDASARKAN LOKASI & JENIS & TYPE/MERK ---
    df_active = df_unit.copy()
    df_inactive_show = df_inaktif.copy() if df_inaktif is not None else pd.DataFrame()

    if selected_loc != "Semua":
        df_active = df_active[df_active['Lokasi'] == selected_loc]
        if not df_inactive_show.empty:
            df_inactive_show = df_inactive_show[df_inactive_show['Lokasi'] == selected_loc]

    if selected_type != "Semua":
        df_active = df_active[df_active['Jenis_Alat'] == selected_type]
        if not df_inactive_show.empty:
            df_inactive_show = df_inactive_show[df_inactive_show['Jenis_Alat'] == selected_type]

    if selected_type_merk != "Semua":
        df_active = df_active[df_active['Type_Merk'] == selected_type_merk]
        if not df_inactive_show.empty:
            df_inactive_show = df_inactive_show[df_inactive_show['Type_Merk'] == selected_type_merk]

    # --- MAIN CONTENT ---
    # [UPDATE] Menambahkan variabel selected_type_merk ke subheader
    st.subheader(f"Analisa Kategori: {selected_loc} - {selected_type} - {selected_type_merk}")

    if not df_inactive_show.empty:
        with st.expander(f"⚠️ {len(df_inactive_show)} Unit Tidak Masuk Analisa (Inaktif)"):
            df_inactive_display = df_inactive_show[['Unit_Name', 'Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work']].copy()
            df_inactive_display['Capacity'] = df_inactive_display.apply(format_capacity_with_unit, axis=1)
            st.dataframe(df_inactive_display.rename(columns={'Type_Merk': 'Type/Merk', 'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Unit_Name': 'Unit'}))
            
    if df_active.empty:
        st.warning(f"Tidak ada unit aktif untuk kategori {selected_loc} - {selected_type}.")
        st.stop()

    # --- KPI CALCULATIONS ---
    total_waste = df_active['Potensi_Pemborosan_Liter'].sum()
    total_loss_rp = total_waste * harga_solar

    df_active.sort_values('Fuel_Ratio', ascending=True, inplace=True)
    best_unit = df_active.iloc[0]

    if len(df_active) > 1:
        worst_unit = df_active.iloc[-1]
        worst_txt = f"{worst_unit['Unit_Name']}"
        worst_val = f"({worst_unit['Fuel_Ratio']:.2f} L/Jam)"
    else: worst_txt = "-"; worst_val = ""

    # --- METRICS ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Populasi Aktif", f"{len(df_active)} Unit")
    m2.metric("Estimasi Kerugian", f"Rp {total_loss_rp:,.0f}", help=f"{total_waste:,.0f} Liter Terbuang")
    m3.metric(f"Unit Teririt: {best_unit['Unit_Name']}", f"{best_unit['Fuel_Ratio']:.2f} L/Jam")

    st.markdown("---")

    # --- TABS ---
    tab_a, tab_b, tab_c, tab_d = st.tabs(["📋 Overview Data", "📊 Efisiensi Setiap Unit", "📉 Persebaran Efisiensi Setiap Unit", "⛽ Unit Terboros"])

    # Tab A: Data Detail
    with tab_a:
        st.subheader("Detail Unit Aktif")
        st.info(f"**Total Pemborosan**: **{total_waste:,.0f} Liter** setara dengan **Rp {total_loss_rp:,.0f}**")
        
        df_display_active = df_active[['Unit_Name', 'Jenis_Alat', 'Type_Merk', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Group_Benchmark_Median', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']].copy()
        df_display_active.sort_values(by='Fuel_Ratio', ascending=False, inplace=True)
        df_display_active['Capacity'] = df_display_active.apply(format_capacity_with_unit, axis=1)
        
        rename_map_active = {'Unit_Name': 'Unit', 'Type_Merk': 'Type/Merk', 'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Group_Benchmark_Median': 'Benchmark', 'Performance_Status': 'Status_BBM', 'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'}
        df_display_active.rename(columns=rename_map_active, inplace=True)

        # Fix konsistensi warna teks
        def highlight_status(row):
            status = str(row['Status_BBM']).upper()
            if status == 'EFISIEN':
                return [f'background-color: #2ca02c; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]
            elif status == 'BOROS':
                return [f'background-color: #d62728; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]
            else:
                return ['' for _ in row.index]

        st.dataframe(df_display_active.style.format({'Horse_Power': '{:.0f}', 'Total_Pengisian_BBM': '{:,.0f}', 'Total_Jam_Kerja': '{:,.0f}', 'Fuel_Ratio': '{:.2f}', 'Benchmark': '{:.2f}', 'Potensi_Pemborosan_BBM': '{:,.0f}'}).apply(highlight_status, axis=1))
        
        st.markdown("---")
        st.markdown("### Efisiensi BBM Bulanan Setiap Unit")
        
        list_unit_active = df_display_active['Unit'].unique().tolist()
        
        if list_unit_active and df_trend_global is not None:
            selected_unit_active = st.selectbox("Pilih Unit yang Diinginkan:", list_unit_active, key='sb_active')
            df_trend_filtered = df_trend_global[df_trend_global['Unit_Name'] == selected_unit_active]
            
            unit_benchmark = df_active[df_active['Unit_Name'] == selected_unit_active]['Group_Benchmark_Median'].iloc[0] if not df_active[df_active['Unit_Name'] == selected_unit_active].empty else 0

            if not df_trend_filtered.empty:
                fig_trend = px.line(df_trend_filtered, x='Bulan', y='Fuel_Ratio', markers=True, 
                                    title=f"Pergerakan Fuel Ratio {selected_unit_active} (Bulanan)",
                                    labels={'Bulan': 'Bulan', 'Fuel_Ratio': 'Fuel Ratio'})
                
                all_y = df_trend_filtered['Fuel_Ratio'].tolist() + [unit_benchmark]
                min_y, max_y = min(all_y), max(all_y)
                padding = (max_y - min_y) * 0.2 if max_y > min_y else (max_y * 0.2 if max_y > 0 else 1.0)
                fig_trend.update_yaxes(range=[max(0, min_y - padding), max_y + padding])

                fig_trend.add_hline(y=unit_benchmark, line_dash="dash", line_color="red", annotation_text=f"Benchmark: {unit_benchmark:.2f} L/Jam", annotation_position="top left", annotation_font_color="white")
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.warning("Data tren bulanan tidak tersedia untuk unit ini.")

    # Tab B: Peringkat
    with tab_b:
        st.subheader("Peringkat Efisiensi Setiap Unit")
        df_plot_bar = df_active.rename(columns={'Unit_Name': 'Unit'})
        
        # Bar chart warna kategori & urutan terkecil ke terbesar
        fig_bar = px.bar(df_plot_bar, x='Unit', y='Fuel_Ratio', color='Performance_Status',
                         color_discrete_map={'EFISIEN': '#2ca02c', 'BOROS': '#d62728'},
                         text_auto='.2f', 
                         title=f"Konsumsi BBM (Liter/Jam)", 
                         labels={'Fuel_Ratio': 'Fuel Ratio', 'Lokasi': 'Lokasi', 'Horse_Power': 'Horse Power', 'Group_Benchmark_Median': 'Benchmark'}, 
                         hover_data={'Lokasi': True, 'Horse_Power': True, 'Group_Benchmark_Median': ':.2f'})
        
        fig_bar.update_layout(xaxis={'categoryorder':'array', 'categoryarray': df_plot_bar['Unit']})
        
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # Tab C: Scatter
    with tab_c:
        st.subheader("Jam Kerja vs BBM")
        color_map_status = {"EFISIEN": "#2ca02c", "BOROS": "#d62728"}
        labels_map = {'Total_HM_Work': 'Total_Jam_Kerja', 'Total_Liter': 'Total_Pengisian_BBM', 'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM', 'Performance_Status': 'Status_BBM', 'Unit_Name': 'Unit', 'Lokasi': 'Lokasi', 'Group_Benchmark_Median': 'Benchmark'}

        # Menyiapkan kolom size
        df_active['Scatter_Size'] = df_active['Potensi_Pemborosan_Liter'].apply(lambda x: 1000 + x if x > 0 else 1000)

        fig_scat = px.scatter(df_active, x='Total_HM_Work', y='Total_Liter', color='Performance_Status', size='Scatter_Size', hover_name='Unit_Name', 
                              hover_data={'Performance_Status': False, 'Total_HM_Work': ':,.0f', 'Total_Liter': ':,.0f', 'Scatter_Size': False, 'Fuel_Ratio': ':.2f', 'Group_Benchmark_Median': ':.2f', 'Potensi_Pemborosan_Liter': ':,.0f', 'Lokasi': True, 'Horse_Power': False}, 
                              color_discrete_map=color_map_status, labels=labels_map, title="Sebaran Efisiensi Setiap Unit")
        st.plotly_chart(fig_scat, use_container_width=True)

    # Tab D: Analisa Pemborosan
    with tab_d:
        st.subheader("Kontribusi Pemborosan Terbesar")
        df_boros = df_active[df_active['Potensi_Pemborosan_Liter'] > 0].copy()
        df_boros.rename(columns={'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM', 'Unit_Name': 'Unit'}, inplace=True)
        
        if not df_boros.empty:
            df_boros.sort_values('Potensi_Pemborosan_BBM', ascending=True, inplace=True)
            fig_waste = px.bar(df_boros.tail(10), x='Potensi_Pemborosan_BBM', y='Unit', orientation='h', title="Unit dengan Potensi Pemborosan Tertinggi (Liter)", text_auto='.0f', color_discrete_sequence=['#c0392b'], labels={'Potensi_Pemborosan_BBM': 'Potensi Pemborosan BBM (Liter)', 'Lokasi': 'Lokasi'}, hover_data=['Lokasi'])
            st.plotly_chart(fig_waste, use_container_width=True)
        else:
            st.success("Tidak ada unit yang terindikasi boros dalam kategori ini.")

elif not master_file and not bbm_file:
    st.info("Silakan upload file berisi data yang dibutuhkan pada menu sebelah kiri untuk memulai analisa.")