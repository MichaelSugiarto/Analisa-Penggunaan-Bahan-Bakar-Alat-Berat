import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import re
import warnings

warnings.filterwarnings('ignore')

# ==============================================================================
# 1. KONFIGURASI FILE & PATH
# ==============================================================================
FILE_HASIL_TRUCKING = "HasilTrucking.xlsx" 
FILE_HASIL_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_BBM_RAW = "BBM AAB.xlsx"
FILE_HAULAGE_RAW = "HAULAGE OKT-DES 2025 (Copy).xlsx"
FILE_DOORING_REVISI = "DOORING_WITH_DISTANCE_REVISI.xlsx" 
FILE_MASTER_REF = "cost & bbm 2022 sd 2025 HP & Type.xlsx"

# ==============================================================================
# 2. SETUP HALAMAN
# ==============================================================================
st.set_page_config(page_title="Dashboard Efisiensi BBM", layout="wide")
st.title("Dashboard Monitoring Efisiensi BBM")

# ==============================================================================
# 3. FUNGSI UTILITIES 
# ==============================================================================
def clean_unit_name(name):
    if pd.isna(name): return ""
    name = str(name).upper().strip()
    name = name.replace("FORKLIFT", "FORKLIF")
    return re.sub(r'[^A-Z0-9]', '', name)

def get_smart_match(raw_name, master_dict):
    """Mencocokkan nama dari file operasional dengan nama resmi di Master File."""
    raw_clean = clean_unit_name(raw_name)
    raw_upper = str(raw_name).upper().strip()

    if raw_clean in master_dict: return raw_clean
    
    if "L 8477 UUC" in raw_upper:
        target = clean_unit_name("L 9902 UR / S75")
        if target in master_dict: return target

    if "EX." in raw_upper or "EX " in raw_upper:
        parts = raw_upper.split("EX.") if "EX." in raw_upper else raw_upper.split("EX ")
        if len(parts) > 1:
            candidate = clean_unit_name(parts[-1].replace(")", "").strip())
            if candidate in master_dict: return candidate
            for k in master_dict:
                if candidate in k: return k

    if "(" in raw_upper:
        candidate = clean_unit_name(raw_upper.split("(")[0])
        if candidate in master_dict: return candidate
        for k in master_dict:
            if candidate in k: return k
    
    return None

# ==============================================================================
# 4. LOGIKA PROSES DATA: NON-TRUCKING 
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_alat_berat():
    if not os.path.exists(FILE_HASIL_NON_TRUCKING):
        st.warning(f"File {FILE_HASIL_NON_TRUCKING} tidak ditemukan.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        df_agg = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Total_Agregat')
        df_monthly = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Data_Bulanan')
        try:
            df_missing = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Unit_Inaktif')
            rename_missing = {
                'Unit_Name': 'Nama Unit',
                'Jenis_Alat': 'Jenis',
                'Type_Merk': 'Type/Merk',
                'Horse_Power': 'Horse Power',
                'Capacity': 'Capacity (Ton)',
                'LITER': 'Total Pengisian BBM (L)',
                'Total_Ton': 'Total Berat Angkutan (Ton)',
                'Total Pengisian BBM': 'Total Pengisian BBM (L)' 
            }
            df_missing.rename(columns=rename_missing, inplace=True)
        except:
            df_missing = pd.DataFrame()
        
        df_agg.columns = df_agg.columns.str.strip()
        df_monthly.columns = df_monthly.columns.str.strip()

        df_agg['Capacity_Num'] = df_agg['Capacity'].fillna(0).astype(float).astype(int)
        df_monthly['Capacity_Num'] = df_monthly['Capacity'].fillna(0).astype(float).astype(int)

        def get_benchmark_group(jenis, cap):
            jenis = str(jenis).upper()
            if 'FORKLIFT' in jenis:
                if 3 <= cap <= 8 or cap == 0: return 'Forklift (Capacity 3-8)'
                elif cap >= 10: return 'Forklift (Capacity 10, 28, 32)'
                return 'Forklift (Lainnya)'
            elif 'REACH STACKER' in jenis: return 'Reach Stacker'
            elif 'LOADER' in jenis: return 'Top Loader & Side Loader'
            elif 'CRANE' in jenis:
                if cap >= 70: return 'Crane (Capacity 75, 80, 127)'
                elif cap >= 40: return 'Crane (Capacity 41)'
                return 'Crane (Lainnya)'
            return 'Lainnya'

        df_agg['Benchmark_Group'] = df_agg.apply(lambda r: get_benchmark_group(r['Jenis_Alat'], r['Capacity_Num']), axis=1)
        
        df_agg['Capacity'] = df_agg['Capacity_Num']
        df_monthly['Capacity'] = df_monthly['Capacity_Num']
        
        df_agg['Fuel Ratio (L/Ton)'] = np.where(df_agg['Total_Ton'] > 0, df_agg['LITER'] / df_agg['Total_Ton'], 0)
        df_monthly['Fuel Ratio (L/Ton)'] = np.where(df_monthly['Total_Ton'] > 0, df_monthly['LITER'] / df_monthly['Total_Ton'], 0)
        
        benchmark = df_agg[df_agg['Total_Ton'] > 0].groupby('Benchmark_Group')['Fuel Ratio (L/Ton)'].median().reset_index()
        benchmark.rename(columns={'Fuel Ratio (L/Ton)': 'Benchmark (L/Ton)'}, inplace=True)
        df_agg = pd.merge(df_agg, benchmark, on='Benchmark_Group', how='left')
        
        def get_status(row):
            if row['Total_Ton'] <= 0: return "Inaktif"
            return "Efisien" if row['Fuel Ratio (L/Ton)'] <= row['Benchmark (L/Ton)'] else "Boros"
            
        df_agg['Status'] = df_agg.apply(get_status, axis=1)
        df_agg['Potensi Pemborosan BBM (L)'] = df_agg.apply(
            lambda r: (r['Fuel Ratio (L/Ton)'] - r['Benchmark (L/Ton)']) * r['Total_Ton'] if r['Status'] == 'Boros' else 0, axis=1
        )

        rename_map = {
            'Unit_Name': 'Nama Unit',
            'Jenis_Alat': 'Jenis',
            'Type_Merk': 'Type/Merk',
            'Horse_Power': 'Horse Power',
            'Capacity': 'Capacity (Ton)',
            'LITER': 'Total Pengisian BBM (L)',
            'Total_Ton': 'Total Berat Angkutan (Ton)'
        }
        df_agg.rename(columns=rename_map, inplace=True)
        df_monthly.rename(columns=rename_map, inplace=True)

        if not df_missing.empty:
            if 'Capacity (Ton)' not in df_missing.columns:
                temp_agg = df_agg[['Nama Unit', 'Capacity (Ton)']].drop_duplicates()
                if 'Nama Unit' in df_missing.columns:
                    df_missing = pd.merge(df_missing, temp_agg, on='Nama Unit', how='left')
                    df_missing['Capacity (Ton)'] = df_missing['Capacity (Ton)'].fillna(0)

        return df_agg, df_monthly, df_missing
    except Exception as e:
        st.error(f"Error memproses data Non-Trucking: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==============================================================================
# 5. LOGIKA PROSES DATA: TRUCKING (FINAL)
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_trucking():
    # A. LOAD MASTER DATA
    master_dict = {}
    if os.path.exists(FILE_MASTER_REF):
        try:
            df_map = pd.read_excel(FILE_MASTER_REF, sheet_name='Sheet2', header=1)
            
            col_name = next((c for c in df_map.columns if 'NAMA' in str(c).upper()), None)
            col_jenis = next((c for c in df_map.columns if 'ALAT' in str(c).upper() and 'BERAT' in str(c).upper() and c != col_name), None)
            col_type = next((c for c in df_map.columns if 'TYPE' in str(c).upper() or 'MERK' in str(c).upper()), None)
            col_loc = next((c for c in df_map.columns if 'LOKASI' in str(c).upper() or 'DES 2025' in str(c).upper()), df_map.columns[2])
            col_hp = next((c for c in df_map.columns if 'HP' in str(c).upper() or 'HORSE' in str(c).upper()), None)
            
            if col_name:
                for _, row in df_map.iterrows():
                    u_name = str(row[col_name]).strip().upper()
                    jenis = str(row[col_jenis]).strip().upper() if col_jenis else ""
                    
                    if "TRONTON" in jenis or "TRAILER" in jenis or "HEAD" in jenis:
                        c_id = clean_unit_name(u_name)
                        master_dict[c_id] = {
                            'Real_Name': u_name,
                            'Jenis': jenis,
                            'Type/Merk': str(row[col_type]).strip() if col_type else "-",
                            'Lokasi': str(row[col_loc]).strip() if col_loc else "-",
                            'Horse Power': row[col_hp] if col_hp else "-",
                            'Capacity': 40 # Force 40 Feet
                        }
        except Exception as e:
            st.error(f"Gagal membaca Master File: {e}")
    
    # B. PROSES DATA UTAMA (AGREGAT)
    df_trucking = pd.DataFrame()
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_raw = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='HASIL_ANALISA')
            
            valid_rows = []
            for _, row in df_raw.iterrows():
                raw_name = str(row['Nama_Unit']) if 'Nama_Unit' in row else str(row.get('EQUIP NAME', ''))
                match_key = get_smart_match(raw_name, master_dict)
                
                if match_key:
                    meta = master_dict[match_key]
                    valid_rows.append({
                        'Nama Unit': meta['Real_Name'],
                        'Jenis': meta['Jenis'],
                        'Type/Merk': meta['Type/Merk'],
                        'Lokasi': meta['Lokasi'],
                        'Horse Power': meta['Horse Power'],
                        'Capacity (Feet)': 40, 
                        'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Berat Angkutan (Ton)': row.get('Total_Ton', 0),
                        'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0), 
                        'Fuel Ratio (L/Ton*Km)': row.get('L_per_TonKm', 0) 
                    })
            
            df_trucking = pd.DataFrame(valid_rows)
            
            if not df_trucking.empty:
                col_ratio = 'Fuel Ratio (L/Ton*Km)'
                col_work = 'Total Kerja (Ton*Km)'
                
                median_ratio = df_trucking[df_trucking[col_ratio] > 0][col_ratio].median()
                df_trucking['Benchmark (L/Ton*Km)'] = median_ratio
                
                df_trucking['Status'] = df_trucking.apply(
                    lambda x: "Efisien" if x[col_ratio] <= x['Benchmark (L/Ton*Km)'] else "Boros", 
                    axis=1
                )
                df_trucking['Potensi Pemborosan BBM (L)'] = df_trucking.apply(
                    lambda r: (r[col_ratio] - r['Benchmark (L/Ton*Km)']) * r[col_work] if r['Status'] == 'Boros' else 0, 
                    axis=1
                )

        except Exception as e:
            st.error(f"Gagal memproses data trucking utama: {e}")

    # C. PROSES DATA BULANAN (LOAD FROM FILE)
    df_monthly_trucking = pd.DataFrame()
    
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_monthly_raw = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name='Data_Bulanan')
            
            monthly_list = []
            for _, row in df_monthly_raw.iterrows():
                # Pastikan nama unit sesuai dengan master
                raw_name = str(row['Nama_Unit'])
                match_key = get_smart_match(raw_name, master_dict)
                
                if match_key:
                    meta = master_dict[match_key]
                    monthly_list.append({
                        'Nama Unit': meta['Real_Name'],
                        'Bulan': str(row['Bulan']).capitalize(), # Standardize Bulan
                        'Total Pengisian BBM (L)': row.get('LITER', 0),
                        'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0),
                        'Jenis': meta['Jenis'],
                        'Type/Merk': meta['Type/Merk'],
                        'Lokasi': meta['Lokasi'],
                        'Horse Power': meta['Horse Power'],
                        'Capacity (Feet)': 40
                    })
            
            if monthly_list:
                df_monthly_trucking = pd.DataFrame(monthly_list)

        except Exception as e:
            st.warning(f"Gagal memuat data bulanan Trucking (Sheet Data_Bulanan): {e}")

    # D. DATA AUDIT (INAKTIF)
    df_missing_truck = pd.DataFrame()
    list_audit = []
    
    if os.path.exists(FILE_HASIL_TRUCKING):
        for sheet in ['OPS_TANPA_BBM', 'BBM_TANPA_OPS', 'GAGAL_MAPPING']:
            try:
                df_aud = pd.read_excel(FILE_HASIL_TRUCKING, sheet_name=sheet)
                col_n = 'Nama Unit' if 'Nama Unit' in df_aud.columns else ('Nama_Unit' if 'Nama_Unit' in df_aud.columns else 'Kode_Lambung')
                
                for _, row in df_aud.iterrows():
                    raw_u = str(row.get(col_n, ''))
                    match_key = get_smart_match(raw_u, master_dict)
                    
                    if match_key:
                        meta = master_dict[match_key]
                        list_audit.append({
                            'Nama Unit': meta['Real_Name'],
                            'Jenis': meta['Jenis'],
                            'Type/Merk': meta['Type/Merk'],
                            'Lokasi': meta['Lokasi'],
                            'Horse Power': meta['Horse Power'],
                            'Capacity (Feet)': 40,
                            'Total Pengisian BBM (L)': row.get('LITER', 0),
                            'Total Kerja (Ton*Km)': row.get('Total_TonKm', 0),
                            'Keterangan': f"Inaktif ({sheet})"
                        })
            except: pass
            
    if list_audit:
        df_missing_truck = pd.DataFrame(list_audit)

    return df_trucking, df_monthly_trucking, df_missing_truck

# ==============================================================================
# 6. SIDEBAR & FILTER
# ==============================================================================
st.sidebar.subheader("Filter Dashboard")
category_filter = st.sidebar.radio("Pilih Kategori Unit:", ["Trucking", "Non-Trucking"])

st.sidebar.markdown("---")
BIAYA_PER_LITER = st.sidebar.number_input("Biaya Bahan Bakar (Rp/Liter)", min_value=0, value=6800, step=100)

df_active_raw = pd.DataFrame()
df_monthly = pd.DataFrame()
df_missing = pd.DataFrame()

if category_filter == "Trucking":
    with st.spinner("Memproses Data Trucking..."):
        df_active_raw, df_monthly, df_missing = process_trucking()
        mode_label = "Trucking"
        ratio_label = "L/Ton*Km"
        work_col = "Total Kerja (Ton*Km)"
else:
    with st.spinner("Memuat Data Non-Trucking..."):
        df_active_raw, df_monthly, df_missing = process_alat_berat()
        mode_label = "Non-Trucking"
        ratio_label = "L/Ton"
        work_col = "Total Berat Angkutan (Ton)"

if not df_active_raw.empty:
    
    # --- PROSES UNIT INAKTIF (DENGAN KETERANGAN DINAMIS) ---
    df_inaktif_from_active = df_active_raw[
        (df_active_raw['Total Pengisian BBM (L)'] <= 0) | 
        (df_active_raw[work_col] <= 0)
    ].copy()
    
    if not df_inaktif_from_active.empty:
        col_bbm = 'Total Pengisian BBM (L)'
        
        # LOGIKA SERAGAM UNTUK TRUCKING & NON-TRUCKING
        conditions = [
            (df_inaktif_from_active[col_bbm] <= 0) & (df_inaktif_from_active[work_col] <= 0),
            (df_inaktif_from_active[work_col] <= 0),
            (df_inaktif_from_active[col_bbm] <= 0)
        ]
        choices = [
            "Tidak ada aktivitas",
            "Unit tidak melakukan aktivitas kerja",
            "Unit tidak pernah mengisi BBM"
        ]
        df_inaktif_from_active['Keterangan'] = np.select(conditions, choices, default="-")

    list_inaktif = []
    if not df_inaktif_from_active.empty: list_inaktif.append(df_inaktif_from_active)
    
    if not df_missing.empty:
        if 'Total Pengisian BBM' in df_missing.columns:
             df_missing.rename(columns={'Total Pengisian BBM': 'Total Pengisian BBM (L)'}, inplace=True)
        
        # Terapkan logika keterangan yang sama untuk data df_missing
        m_bbm_col = 'Total Pengisian BBM (L)'
        m_work_col = work_col 
        
        if m_bbm_col in df_missing.columns and m_work_col in df_missing.columns:
             m_conds = [
                (df_missing[m_bbm_col] <= 0) & (df_missing[m_work_col] <= 0),
                (df_missing[m_work_col] <= 0),
                (df_missing[m_bbm_col] <= 0)
             ]
             m_choices = [
                "Tidak ada aktivitas",
                "Unit tidak melakukan aktivitas kerja",
                "Unit tidak pernah mengisi BBM"
             ]
             df_missing['Keterangan'] = np.select(m_conds, m_choices, default="Inaktif (Sumber: File Audit)")
        else:
             df_missing['Keterangan'] = "Inaktif (Sumber: File Audit)"

        list_inaktif.append(df_missing)

    df_inaktif_all = pd.concat(list_inaktif, ignore_index=True) if list_inaktif else pd.DataFrame()
    
    # DATA AKTIF UTAMA
    df_active = df_active_raw[(df_active_raw['Total Pengisian BBM (L)'] > 0) & (df_active_raw[work_col] > 0)].copy()
    df_full_for_filter = pd.concat([df_active, df_inaktif_all], ignore_index=True) if not df_inaktif_all.empty else df_active

    st.sidebar.markdown("---")
    st.sidebar.header("Filter Data")
    
    lokasi_list = ["Semua"] + sorted(df_full_for_filter['Lokasi'].dropna().unique().tolist())
    selected_lokasi = st.sidebar.selectbox("📍 Filter Lokasi", lokasi_list)
    
    jenis_list = ["Semua"] + sorted(df_full_for_filter['Jenis'].dropna().unique().tolist())
    selected_jenis = st.sidebar.selectbox("🚜 Filter Jenis", jenis_list)
    
    type_list = ["Semua"] + sorted(df_full_for_filter['Type/Merk'].dropna().astype(str).unique().tolist())
    selected_type = st.sidebar.selectbox("🏷️ Filter Type/Merk", type_list)
    
    # ==============================================================================
    # TABEL UNIT INAKTIF
    # ==============================================================================
    st.markdown("### 🛑 Daftar Unit Inaktif")
    st.caption("Unit yang terdeteksi tidak aktif karena tidak ada pengisian BBM atau tidak ada aktivitas kerja yang berlangsung")
    
    df_inaktif_filtered = df_inaktif_all.copy()
    if not df_inaktif_filtered.empty:
        if selected_lokasi != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Lokasi'] == selected_lokasi]
        if selected_jenis != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Jenis'] == selected_jenis]
        if selected_type != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Type/Merk'] == selected_type]
        
        if not df_inaktif_filtered.empty:
            if mode_label == "Trucking":
                cols_inaktif = ['Nama Unit', 'Jenis', 'Type/Merk', 'Lokasi', 'Horse Power', 'Capacity (Feet)', 'Total Pengisian BBM (L)', work_col, 'Keterangan']
                fmt_dict_inaktif = {'Capacity (Feet)': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}'}
            else:
                cols_inaktif = ['Nama Unit', 'Jenis', 'Type/Merk', 'Lokasi', 'Horse Power', 'Capacity (Ton)', 'Total Pengisian BBM (L)', work_col, 'Keterangan']
                fmt_dict_inaktif = {'Capacity (Ton)': '{:.0f}', 'Horse Power': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}'}
                
            cols_to_show = [c for c in cols_inaktif if c in df_inaktif_filtered.columns]
            st.dataframe(df_inaktif_filtered[cols_to_show].style.format(fmt_dict_inaktif, na_rep="-"), use_container_width=True)
        else:
            st.success("Tidak ada unit inaktif untuk kombinasi filter ini.")
    else:
        st.success("Seluruh unit beroperasi aktif dan masuk ke dalam perhitungan dashboard.")
        
    st.markdown("---")
    
    # ==============================================================================
    # PENCARIAN & MAIN CONTENT
    # ==============================================================================
    st.markdown("### 🔍 Cari Data Spesifik (Unit Aktif)")
    c_search1, c_search2 = st.columns([1, 3])
    with c_search1:
        search_category = st.selectbox("Cari Berdasarkan:", ["Nama Unit"])
    with c_search2:
        search_query = st.text_input(f"Ketik {search_category}:", "")

    df_filtered = df_active.copy()
    if selected_lokasi != "Semua": df_filtered = df_filtered[df_filtered['Lokasi'] == selected_lokasi]
    if selected_jenis != "Semua": df_filtered = df_filtered[df_filtered['Jenis'] == selected_jenis]
    if selected_type != "Semua": df_filtered = df_filtered[df_filtered['Type/Merk'] == selected_type]

    if search_query:
        if search_category == "Nama Unit": df_filtered = df_filtered[df_filtered['Nama Unit'].str.contains(search_query, case=False, na=False)]
    
    df_filtered['Total Biaya BBM'] = df_filtered['Total Pengisian BBM (L)'] * BIAYA_PER_LITER
    
    if not df_monthly.empty:
        df_monthly_filtered = df_monthly[df_monthly['Nama Unit'].isin(df_filtered['Nama Unit'])].copy()
    else:
        df_monthly_filtered = pd.DataFrame()

    if mode_label == "Trucking":
        total_bbm = df_filtered['Total Pengisian BBM (L)'].sum()
        total_kerja = df_filtered['Total Kerja (Ton*Km)'].sum()
        total_biaya = df_filtered['Total Biaya BBM'].sum()
        total_unit_aktif = len(df_filtered)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Unit Aktif", f"{total_unit_aktif} Unit")
        c2.metric("Total Kerja (Ton*Km)", f"{total_kerja:,.0f}")
        c3.metric("Total Pengisian BBM", f"{total_bbm:,.0f} L")
        c4.metric("Total Biaya BBM (Rp)", f"Rp {total_biaya:,.0f}")
    else:
        total_bbm = df_filtered['Total Pengisian BBM (L)'].sum()
        total_ton = df_filtered['Total Berat Angkutan (Ton)'].sum()
        total_biaya = df_filtered['Total Biaya BBM'].sum()
        total_unit_aktif = len(df_filtered)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Unit Aktif", f"{total_unit_aktif} Unit")
        c2.metric("Total Tonase Container", f"{total_ton:,.0f} Ton")
        c3.metric("Total Pengisian BBM", f"{total_bbm:,.0f} L")
        c4.metric("Total Biaya BBM (Rp)", f"Rp {total_biaya:,.0f}")

    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview Data", "📈 Efisiensi Setiap Unit", "📍 Korelasi Beban & BBM", "💸 Unit Terboros"])

    # TAB 1: OVERVIEW
    with tab1:
        st.subheader(f"Data Detail {mode_label}")
        
        sort_options = ["Fuel Ratio (Tertinggi)", "Fuel Ratio (Terendah)", "Total Kerja (Tertinggi)", "Total Pengisian BBM (L) (Tertinggi)"]
        sort_by = st.selectbox("Sort by:", sort_options)
        
        ratio_col = 'Fuel Ratio (L/Ton*Km)' if mode_label == "Trucking" else 'Fuel Ratio (L/Ton)'
        bm_col = 'Benchmark (L/Ton*Km)' if mode_label == "Trucking" else 'Benchmark (L/Ton)'

        if sort_by == "Fuel Ratio (Tertinggi)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=False)
        elif sort_by == "Fuel Ratio (Terendah)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=True)
        elif sort_by == "Total Kerja (Tertinggi)": df_filtered = df_filtered.sort_values(by=work_col, ascending=False)
        elif sort_by == "Total Pengisian BBM (L) (Tertinggi)": df_filtered = df_filtered.sort_values(by='Total Pengisian BBM (L)', ascending=False)
        
        def highlight_fuel_ratio(row):
            styles = [''] * len(row)
            for i, col in enumerate(row.index):
                if col == ratio_col:
                    val = row[col]
                    bm = row[bm_col]
                    if pd.notna(val) and pd.notna(bm) and bm > 0:
                        if val > bm:
                            styles[i] = 'background-color: #d62728; color: white; font-weight: bold;' 
                        else:
                            styles[i] = 'background-color: #2ca02c; color: white; font-weight: bold;'
            return styles

        if mode_label == "Trucking":
            cols_show = ['Nama Unit', 'Jenis', 'Lokasi', 'Horse Power', 'Capacity (Feet)', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Total Kerja (Ton*Km)', 'Benchmark (L/Ton*Km)', 'Fuel Ratio (L/Ton*Km)', 'Potensi Pemborosan BBM (L)']
            format_dict = {'Capacity (Feet)': '{:.0f}', 'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Total Kerja (Ton*Km)': '{:,.0f}', 'Benchmark (L/Ton*Km)': '{:.4f}', 'Fuel Ratio (L/Ton*Km)': '{:.4f}', 'Potensi Pemborosan BBM (L)': '{:,.0f}'}
            st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))
        else:
            cols_show = ['Nama Unit', 'Jenis', 'Type/Merk', 'Horse Power', 'Capacity (Ton)', 'Lokasi', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Benchmark (L/Ton)', 'Fuel Ratio (L/Ton)', 'Potensi Pemborosan BBM (L)']
            format_dict = {'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Benchmark (L/Ton)': '{:.4f}', 'Fuel Ratio (L/Ton)': '{:.4f}', 'Potensi Pemborosan BBM (L)': '{:,.0f}'}
            st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))

        # TREN BULANAN
        st.markdown("---")
        st.subheader("📈 Tren Kinerja Bulanan Setiap Unit")
        if not df_monthly_filtered.empty:
            unit_list_trend = sorted(df_monthly_filtered['Nama Unit'].unique().tolist())
            selected_unit_trend = st.selectbox("Pilih Unit untuk melihat tren:", unit_list_trend)
            
            trend_data_unit = df_monthly_filtered[df_monthly_filtered['Nama Unit'] == selected_unit_trend]
            month_order = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
            trend_data_unit['Bulan'] = pd.Categorical(trend_data_unit['Bulan'], categories=month_order, ordered=True)
            
            if mode_label == "Trucking":
                trend_data = trend_data_unit.groupby('Bulan', as_index=False).agg({'Total Kerja (Ton*Km)': 'sum', 'Total Pengisian BBM (L)': 'sum'})
                trend_data = trend_data.dropna()
                trend_data['Ratio'] = np.where(trend_data['Total Kerja (Ton*Km)'] > 0, trend_data['Total Pengisian BBM (L)'] / trend_data['Total Kerja (Ton*Km)'], 0)
                
                c_trend1, c_trend2 = st.columns(2)
                with c_trend1:
                    fig_trend_work = px.bar(trend_data, x='Bulan', y='Total Kerja (Ton*Km)', text_auto='.2s', title=f"Tren Kerja (Ton*Km): {selected_unit_trend}", color_discrete_sequence=['#1f77b4'])
                    st.plotly_chart(fig_trend_work, use_container_width=True)
                with c_trend2:
                    fig_trend_ratio = px.line(trend_data, x='Bulan', y='Ratio', markers=True, title=f"Tren Efisiensi (L/Ton*Km): {selected_unit_trend}", color_discrete_sequence=['#ff7f0e'])
                    fig_trend_ratio.update_yaxes(rangemode="tozero")
                    st.plotly_chart(fig_trend_ratio, use_container_width=True)
            else:
                trend_data = trend_data_unit.groupby('Bulan', as_index=False).agg({'Total Berat Angkutan (Ton)': 'sum', 'Total Pengisian BBM (L)': 'sum'})
                trend_data = trend_data.dropna()
                trend_data['Ratio'] = np.where(trend_data['Total Berat Angkutan (Ton)'] > 0, trend_data['Total Pengisian BBM (L)'] / trend_data['Total Berat Angkutan (Ton)'], 0)
                
                c_trend1, c_trend2 = st.columns(2)
                with c_trend1:
                    fig_trend_ton = px.bar(trend_data, x='Bulan', y='Total Berat Angkutan (Ton)', text_auto='.2s', title=f"Tren Berat Angkutan: {selected_unit_trend}", color_discrete_sequence=['#1f77b4'])
                    st.plotly_chart(fig_trend_ton, use_container_width=True)
                with c_trend2:
                    fig_trend_ratio = px.line(trend_data, x='Bulan', y='Ratio', markers=True, title=f"Tren Efisiensi (L/Ton): {selected_unit_trend}", color_discrete_sequence=['#ff7f0e'])
                    fig_trend_ratio.update_yaxes(rangemode="tozero")
                    st.plotly_chart(fig_trend_ratio, use_container_width=True)

    # TAB 2: BAR CHART
    with tab2:
        st.subheader(f"Peringkat Efisiensi BBM ({ratio_label})")
        df_chart = df_filtered.sort_values(ratio_col, ascending=False)
        
        fig_bar = px.bar(df_chart, x='Nama Unit', y=ratio_col, color='Status',
                         custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                         title=f"Fuel Ratio Seluruh Unit ({ratio_label})",
                         color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'})
        
        fig_bar.update_traces(
            hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                          "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                          "<b>Jenis:</b> %{customdata[2]}<br>" +
                          "<b>Lokasi:</b> %{customdata[3]}<br>" +
                          "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                          "<b>Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # TAB 3: SCATTER PLOT
    with tab3:
        st.subheader("Korelasi Beban Kerja vs BBM")
        max_bubble_size = 45 
        
        if mode_label == "Trucking":
            vals = df_filtered[ratio_col].fillna(0)
            if vals.max() > vals.min():
                 size_col = 10 + ((vals - vals.min()) / (vals.max() - vals.min())) * (max_bubble_size - 10)
            else:
                 size_col = vals.apply(lambda x: 20)

            fig_scat = px.scatter(df_filtered, x='Total Kerja (Ton*Km)', y='Total Pengisian BBM (L)', color='Status',
                                custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                                size=size_col, 
                                size_max=max_bubble_size, opacity=0.65,
                                color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                title="Korelasi Beban Kerja (Total Kerja (Ton*Km)) vs Total Pengisian BBM (L)")
        else:
            size_col_ton = df_filtered[ratio_col].apply(lambda x: x if x > 0 else 0.0001)
            
            fig_scat = px.scatter(df_filtered, x='Total Berat Angkutan (Ton)', y='Total Pengisian BBM (L)', color='Status',
                                    custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                                    size=size_col_ton,
                                    size_max=max_bubble_size, opacity=0.65,
                                    color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                    title="Korelasi Total Berat Angkutan (Ton) vs Total Pengisian BBM (L)")
        
        fig_scat.update_traces(
            hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                          "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                          "<b>Jenis:</b> %{customdata[2]}<br>" +
                          "<b>Lokasi:</b> %{customdata[3]}<br>" +
                          "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                          "<b>Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
        )
        st.plotly_chart(fig_scat, use_container_width=True)

    # TAB 4: UNIT TERBOROS
    with tab4:
        st.subheader("💸 Top 10 Unit dengan Pemborosan Tertinggi")
        
        if mode_label == "Trucking":
            df_waste = df_filtered[df_filtered['Status'] == 'Boros'].copy()
            df_waste = df_waste.sort_values(by='Potensi Pemborosan BBM (L)', ascending=False).head(10)
            
            if not df_waste.empty:
                fig_waste = px.bar(df_waste, x='Nama Unit', y='Potensi Pemborosan BBM (L)', text_auto='.0f',
                                title="Estimasi Pemborosan (Liter BBM)", color_discrete_sequence=['#d62728'])
                st.plotly_chart(fig_waste, use_container_width=True)
                st.write(df_waste[['Nama Unit', 'Total Kerja (Ton*Km)', 'Fuel Ratio (L/Ton*Km)', 'Benchmark (L/Ton*Km)', 'Potensi Pemborosan BBM (L)']])
            else:
                st.success("Tidak ada unit yang tergolong BOROS saat ini.")
        else:
            df_waste = df_filtered[df_filtered['Potensi Pemborosan BBM (L)'] > 0].sort_values(by='Potensi Pemborosan BBM (L)', ascending=False).head(10)
            
            if not df_waste.empty:
                fig_waste = px.bar(df_waste, x='Nama Unit', y='Potensi Pemborosan BBM (L)', text_auto='.0f',
                                title="Estimasi Pemborosan (Liter BBM)", color_discrete_sequence=['#d62728'])
                st.plotly_chart(fig_waste, use_container_width=True)
                st.write(df_waste[['Nama Unit', 'Total Berat Angkutan (Ton)', 'Fuel Ratio (L/Ton)', 'Benchmark (L/Ton)', 'Potensi Pemborosan BBM (L)']])
            else:
                st.success("Tidak ada unit yang tergolong BOROS saat ini.")

elif not df_active_raw.empty and df_filtered.empty:
    st.warning("⚠️ Tidak ada unit yang cocok dengan kombinasi filter & pencarian Anda.")