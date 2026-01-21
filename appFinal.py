import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re

# ==============================================================================
# 1. KONFIGURASI HALAMAN
# ==============================================================================
st.set_page_config(
    page_title="Dashboard BBM Alat Berat",
    #page_icon="🚜",
    layout="wide"
)

st.title("Dashboard Monitoring Efisiensi BBM")

# ==============================================================================
# 2. LOAD DATA UTAMA (SUPER ROBUST MAPPING)
# ==============================================================================
@st.cache_data
def load_data():
    data_kpi = None
    data_unit = None
    data_inaktif = None
    
    # Dictionary Mapping
    map_loc = {}
    map_cap = {} 
    
    # Fungsi Pembersih Nama Ekstrem (Hanya Huruf & Angka)
    def clean_key(text):
        if pd.isna(text): return ""
        text = str(text).upper().strip()
        # [UPDATE] Fix Typo agar konsisten
        text = text.replace("FORKLIFT", "FORKLIF")
        # Buang semua simbol & spasi, uppercase
        return re.sub(r'[^A-Z0-9]', '', text)

    # 0. Load Master Data (cost & bbm 2022 sd 2025.xlsx)
    try:
        df_master = pd.read_excel('cost & bbm 2022 sd 2025.xlsx', header=1)
        
        # Cari Kolom
        col_name_master = next((c for c in df_master.columns if 'NAMA' in str(c).upper()), None)
        cap_keywords = ['KAPASITAS', 'CAPACITY', 'CAP', 'TON', 'CLASS']
        col_cap_master = next((c for c in df_master.columns if any(k in str(c).upper() for k in cap_keywords)), None)
        col_loc_master = 'DES 2025' # Hardcoded sesuai request
        
        if col_name_master and col_loc_master in df_master.columns:
            for _, row in df_master.iterrows():
                u_name = str(row[col_name_master]).strip().upper()
                u_key = clean_key(u_name) # Key Super Bersih
                
                loc_val = str(row[col_loc_master])
                
                # Parsing Capacity
                cap_val = 0
                
                # [UPDATE: MANUAL FIX KAPASITAS L 9025 US]
                if u_key == clean_key("L 9025 US"):
                    cap_val = 40
                else:
                    # Cara A: Dari Kolom Kapasitas
                    if col_cap_master:
                        try:
                            raw_cap = str(row[col_cap_master])
                            match = re.search(r"(\d+(\.\d+)?)", raw_cap)
                            if match:
                                val_float = float(match.group(1))
                                cap_val = int(val_float + 0.5)
                        except:
                            pass
                    
                    # Cara B: Backup dari Nama Unit
                    if cap_val == 0:
                        try:
                            match_name = re.search(r"(\d+(\.\d+)?)\s*(T|TON|K)", u_name)
                            if match_name:
                                val_float = float(match_name.group(1))
                                cap_val = int(val_float + 0.5)
                        except:
                            pass
                
                # Simpan ke Map (Key Asli & Key Bersih)
                if u_name not in map_loc: map_loc[u_name] = loc_val
                if u_key not in map_loc:  map_loc[u_key] = loc_val
                
                if cap_val > 0:
                    if u_name not in map_cap: map_cap[u_name] = cap_val
                    if u_key not in map_cap:  map_cap[u_key] = cap_val
                    
    except Exception as e:
        # print(f"Error Master: {e}") 
        pass

    # 1. Load Data KPI
    try:
        data_kpi = pd.read_excel('Laporan_Benchmark_BBM.xlsx')
    except FileNotFoundError:
        pass

    # 2. Load Data Unit Aktif & Inaktif
    possible_files = ['Benchmark_Per_Alat_Berat_Data_Baru2.xlsx']
    for f in possible_files:
        try:
            data_unit = pd.read_excel(f, sheet_name='Rapor_Unit_Aktif')
            
            # [UPDATE] Rounding Capacity pada Data Unit Aktif
            if 'Capacity' in data_unit.columns:
                data_unit['Capacity'] = pd.to_numeric(data_unit['Capacity'], errors='coerce').fillna(0)
                data_unit['Capacity'] = data_unit['Capacity'].apply(lambda x: int(x + 0.5))

            data_inaktif = pd.read_excel(f, sheet_name='Unit_Inaktif')
            break 
        except FileNotFoundError:
            continue
        except Exception:
            continue
    
    # 3. PERBAIKAN: ISI DATA KOSONG PADA INAKTIF
    if data_inaktif is not None:
        
        # Helper untuk mencari data
        def get_master_data(unit_name, map_dict, default_val):
            # Cek 1: Nama Persis
            clean_name = str(unit_name).strip().upper()
            if clean_name in map_dict:
                return map_dict[clean_name]
            
            # Cek 2: Nama Super Bersih (Tanpa spasi/simbol)
            super_clean = clean_key(unit_name)
            if super_clean in map_dict:
                return map_dict[super_clean]
            
            return default_val

        def fix_inaktif_row(row):
            # Fix Lokasi
            curr_loc = str(row.get('Lokasi', '-'))
            if curr_loc in ['-', 'nan', 'None', '', '0', 'NaT']:
                new_loc = get_master_data(row['Unit_Name'], map_loc, "-")
                row['Lokasi'] = new_loc
            
            # Fix Capacity
            curr_cap = pd.to_numeric(row.get('Capacity', 0), errors='coerce')
            if pd.isna(curr_cap) or curr_cap == 0:
                new_cap = get_master_data(row['Unit_Name'], map_cap, 0)
                row['Capacity'] = new_cap
            else:
                # Rounding jika ada nilai desimal
                row['Capacity'] = int(curr_cap + 0.5)
                
            return row

        # Terapkan Fix baris per baris
        if 'Lokasi' not in data_inaktif.columns: data_inaktif['Lokasi'] = "-"
        if 'Capacity' not in data_inaktif.columns: data_inaktif['Capacity'] = 0
        
        data_inaktif = data_inaktif.apply(fix_inaktif_row, axis=1)

    return data_kpi, data_unit, data_inaktif

df_kpi, df_unit, df_inaktif = load_data()

# ==============================================================================
# FUNGSI LOAD DATA TREN (UPDATED: SMART MATCHING)
# ==============================================================================
@st.cache_data
def load_monthly_data(unit_name_target):
    """
    Mengambil data tren bulanan dengan logika pencocokan cerdas 
    (Manual Map, EX, Kurung, Exclude) sesuai request.
    """
    file_path = 'Laporan_Tren_Efisiensi_Bulanan_Fix.xlsx'
    
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    try:
        df_trend = pd.read_excel(file_path)
        
        # Helper Clean
        def clean_key_trend(text):
            if pd.isna(text): return ""
            text = str(text).upper().strip()
            text = text.replace("FORKLIFT", "FORKLIF")
            return re.sub(r'[^A-Z0-9]', '', text)

        target_clean = clean_key_trend(unit_name_target)
        
        # --- FUNGSI MATCHING BARIS PER BARIS ---
        def is_match(row_unit_name):
            raw_name = str(row_unit_name).strip().upper()
            
            # 1. EXCLUDE JUNK
            if raw_name.startswith(('GENSET', 'KOMPRESSOR', 'MESIN', 'TANGKI', 'SPBU', 'MOBIL')):
                return False
            
            clean_row_id = clean_key_trend(raw_name)
            
            # 2. MANUAL MAPPING (PRIORITAS)
            # FL RENTAL
            if "FL RENTAL 01" in raw_name and "TIMIKA" not in raw_name:
                if target_clean == clean_key_trend("FL RENTAL 01 TIMIKA"): return True
            # TOBATI
            elif "TOBATI" in raw_name and "KALMAR 32T" in raw_name:
                if target_clean == clean_key_trend("TOP LOADER KALMAR 35T/TOBATI"): return True
            # L 8477 UUC
            elif "L 8477 UUC" in raw_name:
                if target_clean == clean_key_trend("L 9902 UR / S75"): return True
            # L 9054 UT
            elif "L 9054 UT" in raw_name:
                if target_clean == clean_key_trend("L 9054 UT"): return True

            # 3. EXACT MATCH (CLEAN)
            if clean_row_id == target_clean:
                return True
                
            # 4. LOGIKA SEBELUM KURUNG " ("
            if " (" in raw_name:
                try:
                    part_before = raw_name.split(" (")[0].strip()
                    if clean_key_trend(part_before) == target_clean: return True
                except: pass

            # 5. LOGIKA SETELAH "EX."
            if "EX." in raw_name:
                try:
                    part_after_ex = raw_name.split("EX.")[-1].replace(")", "").strip()
                    clean_after = clean_key_trend(part_after_ex)
                    
                    if clean_after == target_clean: return True
                    # Partial Match
                    if clean_after and clean_after in target_clean: return True
                except: pass
                
            return False

        # Terapkan Filter
        mask = df_trend['Unit'].apply(is_match)
        row_data = df_trend[mask]
        
        if not row_data.empty:
            # Ambil yang pertama jika ada duplikat
            row_data = row_data.iloc[[0]]
            
            month_cols = [c for c in df_trend.columns if str(c).startswith('2025')]
            
            df_melt = row_data.melt(
                id_vars=['Unit'], 
                value_vars=month_cols, 
                var_name='Bulan', 
                value_name='Fuel_Ratio'
            )
            df_melt['Fuel_Ratio'] = pd.to_numeric(df_melt['Fuel_Ratio'], errors='coerce').fillna(0)
            return df_melt
            
    except Exception as e:
        pass
    
    return pd.DataFrame()

# ==============================================================================
# 3. STANDARDIZASI DATA
# ==============================================================================
if df_inaktif is not None:
    if 'Lokasi' not in df_inaktif.columns: df_inaktif['Lokasi'] = "-"
    # Capacity sudah dihandle di load_data

# ==============================================================================
# 4. SIDEBAR: NAVIGASI
# ==============================================================================
st.sidebar.header("Penentuan Benchmark")

analysis_mode = st.sidebar.radio(
    "Pilih Benchmark Analisa:",
    ["Group KPI", "Jenis Alat & Kapasitas"]
)

st.sidebar.markdown("---")

# ==============================================================================
# MODE 1: GROUP KPI
# ==============================================================================
if analysis_mode == "Group KPI":
    
    if df_kpi is None:
        st.error("⚠️ File 'Laporan_Benchmark_BBM.xlsx' tidak ditemukan.")
        st.stop()

    # --- Sidebar Filters ---
    st.sidebar.subheader("Filter Group KPI")
    groups = sorted(df_kpi['Benchmark_Group'].astype(str).unique())
    selected_group = st.sidebar.selectbox("Pilih Benchmark Group:", groups)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("Biaya Bahan Bakar")
    harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_kpi')
    
    # --- PENCARIAN UNIT ---
    st.subheader("Cari Unit Spesifik")
    search_kpi = st.text_input("Ketik Nama Unit:", key="search_kpi", placeholder="Cari unit...").upper()
    if search_kpi:
        res_kpi = df_kpi[df_kpi['Unit'].astype(str).str.contains(search_kpi, na=False)]
        if not res_kpi.empty:
            st.info(f"Ditemukan {len(res_kpi)} Unit:")
            
            cols_to_show = ['Unit', 'Category', 'Total_Solar_Liter', 'Total_Jam', 'Rata_Rata_Efisiensi']
            bench_col = None
            if 'Benchmark_Median' in res_kpi.columns:
                bench_col = 'Benchmark_Median'
                cols_to_show.append(bench_col)
            
            rename_map = {
                'Total_Solar_Liter': 'Total_Pengisian_BBM',
                'Total_Jam': 'Total_Jam_Kerja',
                'Rata_Rata_Efisiensi': 'Fuel_Ratio'
            }
            if bench_col:
                rename_map[bench_col] = 'Benchmark'
                
            st.dataframe(res_kpi[cols_to_show].rename(columns=rename_map))
        else:
            st.warning("Unit Tidak Ditemukan.")
    st.markdown("---")

    # --- MAIN CONTENT ---
    st.subheader(f"Analisa Group KPI: {selected_group}")
    
    col_durasi = next((c for c in df_kpi.columns if 'Total_Jam' in c or 'HM' in c), 'Total_Jam')
    df_view = df_kpi[df_kpi['Benchmark_Group'] == selected_group].copy()
    
    # Cleaning Status BBM
    if 'Status_BBM' in df_view.columns:
        df_view['Status_BBM'] = df_view['Status_BBM'].astype(str).str.replace(r'\s*\(.*?\)', '', regex=True)

    # Unit Inaktif Check
    if df_inaktif is not None:
        df_inactive_kpi = df_inaktif[df_inaktif['Benchmark_Group'] == selected_group]
        if not df_inactive_kpi.empty:
             with st.expander(f"⚠️ {len(df_inactive_kpi)} Unit Tidak Masuk Analisa"):
                st.warning(f"Unit berikut masuk dalam grup **{selected_group}** tetapi memiliki Total HM=0 atau BBM=0.")
                st.dataframe(df_inactive_kpi[['Unit_Name', 'Lokasi', 'Total_Liter', 'Total_HM_Work']]
                             .rename(columns={'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja'}))

    # KPI Metrics
    total_solar = df_view['Total_Solar_Liter'].sum()
    total_jam = df_view[col_durasi].sum()
    avg_eff = total_solar / total_jam if total_jam > 0 else 0
    populasi = df_view['Unit'].nunique()
    
    total_waste_liter = df_view['Potensi_Pemborosan_Liter'].sum() if 'Potensi_Pemborosan_Liter' in df_view.columns else 0
    estimasi_rugi_rp = total_waste_liter * harga_solar
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Populasi", f"{populasi} Unit")
    c2.metric("Total BBM yang Digunakan", f"{total_solar:,.0f} Liter")
    c3.metric("Benchmark (Median)", f"{avg_eff:.2f} L/Jam")
    c4.metric("Estimasi Kerugian", f"Rp {estimasi_rugi_rp:,.0f}", help=f"{total_waste_liter:,.0f} Liter Terbuang")
    
    st.markdown("---")
    
    # --- TABS ---
    tab0, tab1, tab2 = st.tabs(["📋 Overview Data", "📉 Persebaran Efisiensi Setiap Unit", "⛽ Top 10 Unit Terboros"])
    
    with tab0:
        st.subheader(f"Detail Data: {selected_group}")
        
        cols_kpi_show = [c for c in ['Unit', 'Category', 'Total_Solar_Liter', col_durasi, 'Rata_Rata_Efisiensi', 'Status_BBM', 'Potensi_Pemborosan_Liter'] if c in df_view.columns]
        
        # [UPDATE] Sort by Fuel_Ratio Ascending (Efisien -> Boros)
        df_display_kpi = df_view[cols_kpi_show].sort_values('Rata_Rata_Efisiensi', ascending=True).copy()
        
        rename_map_kpi = {
            'Total_Solar_Liter': 'Total_Pengisian_BBM', 
            col_durasi: 'Total_Jam_Kerja',
            'Rata_Rata_Efisiensi': 'Fuel_Ratio',
            'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'
        }
        df_display_kpi.rename(columns=rename_map_kpi, inplace=True)
        
        def highlight_status_kpi(row):
            status = str(row.get('Status_BBM', '')).upper()
            color = ''
            if 'EFISIEN' in status:
                color = '#2ca02c' 
            elif 'BOROS' in status:
                color = '#d62728'
            
            return [f'background-color: {color}; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]

        st.dataframe(
            df_display_kpi.style.format({
                'Total_Pengisian_BBM': '{:,.0f}', 
                'Total_Jam_Kerja': '{:,.0f}', 
                'Fuel_Ratio': '{:.2f}', 
                'Potensi_Pemborosan_BBM': '{:,.0f}'
            })
            .apply(highlight_status_kpi, axis=1)
        )
        
        st.markdown("---")
        st.markdown("### Efisiensi BBM Bulanan Setiap Unit")
        
        list_unit_kpi = df_display_kpi['Unit'].unique().tolist()
        
        if list_unit_kpi:
            selected_unit_kpi = st.selectbox("Pilih Unit yang Diinginkan:", list_unit_kpi, key='sb_kpi')
            
            df_trend = load_monthly_data(selected_unit_kpi)
            
            if not df_trend.empty:
                fig_trend = px.line(df_trend, x='Bulan', y='Fuel_Ratio', markers=True, title=f"Pergerakan Fuel Ratio {selected_unit_kpi} (Jan-Nov)")
                
                # Auto Scaling Y Axis
                all_y = df_trend['Fuel_Ratio'].tolist() + [avg_eff]
                min_y, max_y = min(all_y), max(all_y)
                padding = (max_y - min_y) * 0.2 if max_y > min_y else (max_y * 0.2 if max_y > 0 else 1.0)
                fig_trend.update_yaxes(range=[max(0, min_y - padding), max_y + padding])

                fig_trend.add_hline(
                    y=avg_eff, 
                    line_dash="dash", 
                    line_color="red",
                    annotation_text=f"Benchmark: {avg_eff:.2f} L/Jam", 
                    annotation_position="top left", 
                    annotation_font_color="white",
                )
                
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.warning("Data tren bulanan tidak tersedia untuk unit ini di file laporan.")

    # TAB 1: SCATTER MATRIX
    with tab1:
        color_col = 'Status_BBM' if 'Status_BBM' in df_view.columns else None
        color_map = {"EFISIEN": "#2ca02c", "BOROS": "#d62728"} if color_col else None
        
        labels_map = {
            col_durasi: 'Total_Jam_Kerja',
            'Total_Solar_Liter': 'Total_Pengisian_BBM', 
            'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM',
            'Status_BBM': 'Status_BBM',
            'Unit': 'Unit'
        }

        hover_data_kpi = {col_durasi: True, 'Total_Solar_Liter': True}
        if 'Potensi_Pemborosan_Liter' in df_view.columns:
            hover_data_kpi['Potensi_Pemborosan_Liter'] = ':,.0f'
            
        fig = px.scatter(
            df_view, 
            x=col_durasi, 
            y="Total_Solar_Liter", 
            color=color_col, 
            size="Total_Solar_Liter", 
            hover_name="Unit", 
            hover_data=hover_data_kpi,
            color_discrete_map=color_map, 
            labels=labels_map, 
            title="Sebaran Efisiensi Setiap Unit"
        )
        st.plotly_chart(fig, use_container_width=True)
        
    with tab2:
        if 'Potensi_Pemborosan_Liter' in df_view.columns:
            df_waste = df_view[df_view['Potensi_Pemborosan_Liter'] > 0].sort_values('Potensi_Pemborosan_Liter', ascending=False).head(10)
            
            df_waste.rename(columns={'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'}, inplace=True)
            
            if not df_waste.empty:
                fig_bar = px.bar(
                    df_waste, 
                    x='Unit', 
                    y='Potensi_Pemborosan_BBM', 
                    text_auto='.0f', 
                    title="Top 10 Unit Boros", 
                    color_discrete_sequence=['#c0392b'], # Calm Red
                    labels={'Potensi_Pemborosan_BBM': 'Potensi Pemborosan BBM (Liter)'}
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.success("Tidak ada unit boros.")

# ==============================================================================
# MODE 2: JENIS ALAT & KAPASITAS
# ==============================================================================
elif analysis_mode == "Jenis Alat & Kapasitas":
    
    if df_unit is None:
        st.error("⚠️ File Analisa Alat Berat tidak ditemukan.")
        st.stop()
        
    # --- PENCARIAN UNIT ---
    st.subheader("Cari Unit Spesifik")
    search_unit = st.text_input("Ketik Nama Unit:", key="search_unit", placeholder="Cari unit...").upper()
    if search_unit:
        res_active = df_unit[df_unit['Unit_Name'].astype(str).str.contains(search_unit, na=False)].copy()
        res_active['Status'] = 'AKTIF'
        
        res_inactive = pd.DataFrame()
        if df_inaktif is not None:
            res_inactive = df_inaktif[df_inaktif['Unit_Name'].astype(str).str.contains(search_unit, na=False)].copy()
            res_inactive['Status'] = 'INAKTIF'
            
        res_all = pd.concat([res_active, res_inactive], ignore_index=True)
        
        if not res_all.empty:
            st.info(f"Ditemukan {len(res_all)} Unit:")
            
            cols = ['Unit_Name', 'Jenis_Alat', 'Capacity', 'Lokasi', 'Status', 'Fuel_Ratio', 'Total_Liter']
            if 'Group_Benchmark_Median' in res_all.columns:
                cols.append('Group_Benchmark_Median')
            
            rename_map_search = {
                'Group_Benchmark_Median': 'Benchmark',
                'Total_Liter': 'Total_Pengisian_BBM',
                'Unit_Name': 'Unit'
            }
                
            st.dataframe(res_all[cols].rename(columns=rename_map_search))
        else:
            st.warning("Unit Tidak Ditemukan.")
    
    st.markdown("---")

    # --- Sidebar Filters ---
    st.sidebar.subheader("Filter Spesifik")
    
    # 1. Pilih Jenis
    jenis_list = sorted(df_unit['Jenis_Alat'].astype(str).unique())
    selected_jenis = st.sidebar.selectbox("1. Pilih Jenis Alat:", jenis_list)
    
    # 2. Pilih Kapasitas (LOGIKA DYNAMIC DROPDOWN)
    
    df_check_active = df_unit[df_unit['Jenis_Alat'] == selected_jenis].copy()
    
    df_check_inactive = pd.DataFrame()
    if df_inaktif is not None:
        df_check_inactive = df_inaktif[df_inaktif['Jenis_Alat'] == selected_jenis].copy()
        
    def has_data(min_val, max_val=None, exact_val=None):
        if exact_val is not None:
            act = (df_check_active['Capacity'] == exact_val).any()
            inact = (df_check_inactive['Capacity'] == exact_val).any() if not df_check_inactive.empty else False
            return act or inact
        else:
            cond_act = (df_check_active['Capacity'] >= min_val)
            if max_val is not None:
                cond_act = cond_act & (df_check_active['Capacity'] < max_val)
            act = cond_act.any()
            
            inact = False
            if not df_check_inactive.empty:
                cond_inact = (df_check_inactive['Capacity'] >= min_val)
                if max_val is not None:
                    cond_inact = cond_inact & (df_check_inactive['Capacity'] < max_val)
                inact = cond_inact.any()
            return act or inact

    # Generate Dropdown Options
    cap_options = []
    
    if selected_jenis in ['FORKLIFT', 'SIDE LOADER']:
        if has_data(0, 5): cap_options.append('Di Bawah 5 Ton')
        if has_data(5, 10): cap_options.append('Di Bawah 10 Ton')
        if has_data(10, 15): cap_options.append('Di Bawah 15 Ton')
        if has_data(15): cap_options.append('15 Ton ke Atas') # Renamed
        
    elif selected_jenis == 'CRANE':
        if has_data(0, 100): cap_options.append('Di Bawah 100 Ton')
        if has_data(100): cap_options.append('Di Atas 100 Ton')
        
    elif selected_jenis in ['REACH STACKER', 'TOP LOADER']:
        if has_data(35): cap_options.append('Di Atas 35 Ton')
        
    elif selected_jenis in ['TRAILER', 'TRONTON']:
        if has_data(0, exact_val=40): cap_options.append('40 Ton')
        
    else:
        cap_options = sorted(df_check_active['Capacity'].unique().astype(str).tolist())
        
    if not cap_options:
        st.warning(f"Tidak ada data unit (Aktif/Inaktif) untuk jenis {selected_jenis}")
        st.stop()
        
    selected_cap_filter = st.sidebar.selectbox("2. Pilih Kategori Kapasitas:", cap_options)
    
    # 3. Input Harga Solar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Biaya Bahan Bakar")
    harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_alat')
    
    # --- FILTER FINAL ---
    df_active = pd.DataFrame()
    df_inactive_show = pd.DataFrame()
    
    def apply_filter(df_target):
        if df_target.empty: return df_target
        if selected_cap_filter == 'Di Bawah 5 Ton':
            return df_target[df_target['Capacity'] < 5]
        elif selected_cap_filter == 'Di Bawah 10 Ton':
            return df_target[(df_target['Capacity'] >= 5) & (df_target['Capacity'] < 10)]
        elif selected_cap_filter == 'Di Bawah 15 Ton':
            return df_target[(df_target['Capacity'] >= 10) & (df_target['Capacity'] < 15)]
        elif selected_cap_filter == '15 Ton ke Atas':
            return df_target[df_target['Capacity'] >= 15]
        elif selected_cap_filter == 'Di Bawah 100 Ton':
            return df_target[df_target['Capacity'] < 100]
        elif selected_cap_filter == 'Di Atas 100 Ton':
            return df_target[df_target['Capacity'] >= 100]
        elif selected_cap_filter == 'Di Atas 35 Ton':
            return df_target[df_target['Capacity'] > 35]
        elif selected_cap_filter == '40 Ton':
            return df_target[df_target['Capacity'] == 40]
        else:
            try:
                val = float(selected_cap_filter)
                return df_target[df_target['Capacity'] == val]
            except:
                return df_target

    df_active = apply_filter(df_check_active).copy()
    df_inactive_show = apply_filter(df_check_inactive).copy()
    
    # --- MAIN CONTENT ---
    st.subheader(f"Analisa: {selected_jenis} - {selected_cap_filter}")
    
    if not df_inactive_show.empty:
        with st.expander(f"⚠️ {len(df_inactive_show)} Unit Tidak Masuk Analisa"):
            st.dataframe(df_inactive_show[['Unit_Name', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work']]
                          .rename(columns={'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Unit_Name': 'Unit'}))
            
    if df_active.empty:
        st.warning(f"Tidak ada unit aktif untuk kategori {selected_jenis} {selected_cap_filter}.")
        st.stop()

    # --- KPI CALCULATIONS ---
    benchmark_val = df_active['Group_Benchmark_Median'].iloc[0] if 'Group_Benchmark_Median' in df_active.columns else 0
    total_waste = df_active['Potensi_Pemborosan_Liter'].sum()
    total_loss_rp = total_waste * harga_solar
    
    # [UPDATE] Sort by Fuel_Ratio Ascending (Efisien -> Boros)
    df_active.sort_values('Fuel_Ratio', ascending=True, inplace=True)
    best_unit = df_active.iloc[0]
    
    if len(df_active) > 1:
        worst_unit = df_active.iloc[-1]
        worst_txt = f"{worst_unit['Unit_Name']}"
        worst_val = f"({worst_unit['Fuel_Ratio']:.2f} L/Jam)"
    else:
        worst_txt = "-"
        worst_val = ""

    # --- METRICS ---
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Populasi Aktif", f"{len(df_active)} Unit")
    m2.metric("Benchmark (Median)", f"{benchmark_val:.2f} L/Jam")
    m3.metric("Estimasi Kerugian", f"Rp {total_loss_rp:,.0f}", help=f"{total_waste:,.0f} Liter Terbuang")
    m4.metric(f"Unit Teririt: {best_unit['Unit_Name']}", f"{best_unit['Fuel_Ratio']:.2f} L/Jam")
    
    st.markdown("---")
    
    # --- TABS ---
    tab_a, tab_b, tab_c, tab_d = st.tabs(["📋 Overview Data", "📊 Efisiensi Setiap Unit", "📉 Persebaran Efisiensi Setiap Unit", "⛽ Unit Terboros"])
    
    # Tab A: Data Detail (FITUR PILIH UNIT -> GRAFIK)
    with tab_a:
        st.subheader("Detail Unit Aktif")
        st.info(f"**Total Pemborosan**: **{total_waste:,.0f} Liter** setara dengan **Rp {total_loss_rp:,.0f}**")
        
        # [PERBAIKAN] RENAME KOLOM MODE 2
        df_display_active = df_active[['Unit_Name', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']].copy()
        
        rename_map_active = {
            'Unit_Name': 'Unit',
            'Total_Liter': 'Total_Pengisian_BBM',
            'Total_HM_Work': 'Total_Jam_Kerja',
            'Performance_Status': 'Status_BBM',
            'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'
        }
        df_display_active.rename(columns=rename_map_active, inplace=True)

        def highlight_status(row):
            color = '#2ca02c' if row['Status_BBM'] == 'EFISIEN' else ('#d62728' if row['Status_BBM'] == 'BOROS' else '')
            return [f'background-color: {color}; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]

        st.dataframe(
            df_display_active.style.format({
                'Capacity': '{:.0f}',
                'Total_Pengisian_BBM': '{:,.0f}', 
                'Total_Jam_Kerja': '{:,.0f}', 
                'Fuel_Ratio': '{:.2f}', 
                'Potensi_Pemborosan_BBM': '{:,.0f}'
            })
            .apply(highlight_status, axis=1)
        )
        
        st.markdown("---")
        st.markdown("### Efisiensi BBM Bulanan Setiap Unit")
        
        # Daftar Unit Aktif
        list_unit_active = df_display_active['Unit'].unique().tolist() # Use 'Unit'
        
        if list_unit_active:
            selected_unit_active = st.selectbox("Pilih Unit yang Diinginkan:", list_unit_active, key='sb_active')
            
            df_trend = load_monthly_data(selected_unit_active)
            
            if not df_trend.empty:
                fig_trend = px.line(df_trend, x='Bulan', y='Fuel_Ratio', markers=True, title=f"Pergerakan Fuel Ratio {selected_unit_active} (Jan-Nov)")
                
                # Auto Scaling Y Axis
                all_y = df_trend['Fuel_Ratio'].tolist() + [benchmark_val]
                min_y, max_y = min(all_y), max(all_y)
                padding = (max_y - min_y) * 0.2 if max_y > min_y else (max_y * 0.2 if max_y > 0 else 1.0)
                fig_trend.update_yaxes(range=[max(0, min_y - padding), max_y + padding])

                fig_trend.add_hline(
                    y=benchmark_val, 
                    line_dash="dash", 
                    line_color="red",
                    annotation_text=f"Benchmark: {benchmark_val:.2f} L/Jam", 
                    annotation_position="top left",
                    annotation_font_color="white",
                )
                
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.warning("Data tren bulanan tidak tersedia untuk unit ini di file laporan.")
    
    # Tab B: Peringkat
    with tab_b:
        st.subheader("Peringkat Efisiensi Setiap Unit")
        
        df_plot_bar = df_active.rename(columns={'Unit_Name': 'Unit'})
        
        fig_bar = px.bar(
            df_plot_bar, x='Unit', y='Fuel_Ratio', color='Fuel_Ratio',
            color_continuous_scale='RdYlGn_r', text_auto='.2f',
            title=f"Konsumsi BBM (Liter/Jam)",
            labels={'Fuel_Ratio': 'Fuel Ratio'}
        )
        
        fig_bar.add_hline(
            y=benchmark_val, 
            line_dash="dash",
            line_color="white",
            line_width=2,
            annotation_text=f"Benchmark: {benchmark_val:.2f} L/Jam",
            annotation_position="top left",
            annotation_font_color="white",
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # Tab C: Scatter
    with tab_c:
        st.subheader("Jam Kerja vs BBM")
        
        color_map_status = {"EFISIEN": "#2ca02c", "BOROS": "#d62728"}
        
        # Mapping Label untuk Scatter
        labels_map = {
            'Total_HM_Work': 'Total_Jam_Kerja',
            'Total_Liter': 'Total_Pengisian_BBM',
            'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM',
            'Performance_Status': 'Status_BBM',
            'Unit_Name': 'Unit'
        }

        fig_scat = px.scatter(
            df_active, 
            x='Total_HM_Work', 
            y='Total_Liter', 
            color='Performance_Status', 
            size='Total_Liter',
            hover_name='Unit_Name', 
            hover_data={
                'Performance_Status': False,
                'Fuel_Ratio': ':.2f',
                'Potensi_Pemborosan_Liter': ':,.0f'
            },
            color_discrete_map=color_map_status,
            labels=labels_map, 
            title="Sebaran Efisiensi Setiap Unit"
        )
        
        st.plotly_chart(fig_scat, use_container_width=True)

    # Tab D: Analisa Pemborosan (SECTION BARU)
    with tab_d:
        st.subheader("Kontribusi Pemborosan Terbesar")
        
        # Filter hanya yang boros (> 0)
        df_boros = df_active[df_active['Potensi_Pemborosan_Liter'] > 0].copy()
        
        # Rename kolom sebelum plotting
        df_boros.rename(columns={'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM', 'Unit_Name': 'Unit'}, inplace=True)
        
        if not df_boros.empty:
            df_boros.sort_values('Potensi_Pemborosan_BBM', ascending=True, inplace=True) # Ascending for horizontal bar
            
            fig_waste = px.bar(
                df_boros.tail(10), # Top 10
                x='Potensi_Pemborosan_BBM',
                y='Unit', 
                orientation='h',
                title="Unit dengan Potensi Pemborosan Tertinggi (Liter)",
                text_auto='.0f',
                color_discrete_sequence=['#c0392b'], # Calm Red
                labels={'Potensi_Pemborosan_BBM': 'Potensi Pemborosan BBM (Liter)'}
            )
            fig_waste.update_layout(xaxis_title="Potensi Pemborosan (Liter)", yaxis_title="Nama Unit")
            st.plotly_chart(fig_waste, use_container_width=True)
        else:
            st.success("✅ Tidak ada unit yang terindikasi boros dalam kategori ini.")