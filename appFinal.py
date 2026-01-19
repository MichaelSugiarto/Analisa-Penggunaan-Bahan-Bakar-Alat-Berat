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
# 2. LOAD DATA UTAMA
# ==============================================================================
@st.cache_data
def load_data():
    data_kpi = None
    data_unit = None
    data_inaktif = None
    map_loc = {}
    
    # 0. Load Master Data
    try:
        df_master = pd.read_excel('cost & bbm 2022 sd 2025.xlsx', header=1)
        if 'NAMA ALAT BERAT' in df_master.columns and 'DES 2025' in df_master.columns:
            for _, row in df_master.iterrows():
                u_name = str(row['NAMA ALAT BERAT']).strip().upper()
                loc = row['DES 2025']
                map_loc[u_name] = loc
                u_norm = " ".join(u_name.replace('/', ' ').replace('.', ' ').split())
                map_loc[u_norm] = loc
    except Exception:
        pass 

    # 1. Load Data KPI
    try:
        data_kpi = pd.read_excel('Laporan_Benchmark_BBM.xlsx')
    except FileNotFoundError:
        pass

    # 2. Load Data Unit Aktif
    possible_files = ['Benchmark_Per_Alat_Berat_Data_Baru.xlsx']
    for f in possible_files:
        try:
            data_unit = pd.read_excel(f, sheet_name='Rapor_Unit_Aktif')
            data_inaktif = pd.read_excel(f, sheet_name='Unit_Inaktif')
            break 
        except FileNotFoundError:
            continue
        except Exception:
            continue
    
    # Apply Location Map
    if data_inaktif is not None and map_loc:
        def fill_loc(row):
            current_loc = str(row.get('Lokasi', '-'))
            if current_loc in ['-', 'nan', 'None', '']:
                unit_name = str(row['Unit_Name']).strip().upper()
                if unit_name in map_loc: return map_loc[unit_name]
                unit_norm = " ".join(unit_name.replace('/', ' ').replace('.', ' ').split())
                if unit_norm in map_loc: return map_loc[unit_norm]
                return "-"
            return current_loc
        
        if 'Lokasi' not in data_inaktif.columns: data_inaktif['Lokasi'] = "-"
        data_inaktif['Lokasi'] = data_inaktif.apply(fill_loc, axis=1)

    return data_kpi, data_unit, data_inaktif

df_kpi, df_unit, df_inaktif = load_data()

# ==============================================================================
# FUNGSI LOAD DATA TREN (DARI FILE LAPORAN FIX)
# ==============================================================================
@st.cache_data
def load_monthly_data(unit_name_target):
    file_path = 'Laporan_Tren_Efisiensi_Bulanan_Fix.xlsx'
    
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    try:
        df_trend = pd.read_excel(file_path)
        
        def normalize(name):
            if pd.isna(name): return ""
            name = str(name).upper().strip()
            name = re.sub(r'[/\-._]', ' ', name)
            return " ".join(name.split())

        target_clean = normalize(unit_name_target)
        df_trend['Unit_Clean'] = df_trend['Unit'].apply(normalize)
        
        row_data = df_trend[df_trend['Unit_Clean'] == target_clean]
        
        if not row_data.empty:
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
    if 'Capacity' not in df_inaktif.columns: df_inaktif['Capacity'] = "-" 

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
    c2.metric("Total Solar", f"{total_solar:,.0f} Liter")
    c3.metric("Benchmark (Median)", f"{avg_eff:.2f} L/Jam")
    c4.metric("Estimasi Kerugian", f"Rp {estimasi_rugi_rp:,.0f}", help=f"{total_waste_liter:,.0f} Liter Terbuang")
    
    st.markdown("---")
    
    # --- TABS ---
    tab0, tab1, tab2 = st.tabs(["📋 Overview Data", "📉 Persebaran Efisiensi Setiap Unit", "⛽ Top 10 Unit Terboros"])
    
    with tab0:
        st.subheader(f"Detail Data: {selected_group}")
        
        cols_kpi_show = [c for c in ['Unit', 'Category', 'Total_Solar_Liter', col_durasi, 'Rata_Rata_Efisiensi', 'Status_BBM', 'Potensi_Pemborosan_Liter'] if c in df_view.columns]
        
        df_display_kpi = df_view[cols_kpi_show].sort_values('Total_Solar_Liter', ascending=False).copy()
        
        rename_map_kpi = {
            'Total_Solar_Liter': 'Total_Penggunaan_BBM', 
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
                'Total_Penggunaan_BBM': '{:,.0f}', 
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
                
                # [UPDATE] Posisi Teks Top Left & Background Putih
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
            'Total_Solar_Liter': 'Total_Penggunaan_BBM', 
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
            title="Sebaran Efisiensi Unit"
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
    
    # 2. Pilih Kapasitas (LOGIKA BARU SESUAI REQUEST)
    df_jenis_filtered = df_unit[df_unit['Jenis_Alat'] == selected_jenis]
    cap_options = []
    
    if selected_jenis in ['FORKLIFT', 'SIDE LOADER']:
        cap_options = ['Di Bawah 5 Ton', 'Di Bawah 10 Ton', 'Di Bawah 15 Ton']
    elif selected_jenis == 'CRANE':
        cap_options = ['Di Bawah 100 Ton', 'Di Atas 100 Ton']
    elif selected_jenis in ['REACH STACKER', 'TOP LOADER']:
        cap_options = ['Di Atas 35 Ton']
    else:
        cap_options = sorted(df_jenis_filtered['Capacity'].unique().astype(str).tolist())
        
    selected_cap_filter = st.sidebar.selectbox("2. Pilih Kategori Kapasitas:", cap_options)
    
    # 3. Input Harga Solar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Biaya Bahan Bakar")
    harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_alat')
    
    # --- FILTER FINAL ---
    df_active = pd.DataFrame()
    df_inactive_show = pd.DataFrame()
    df_jenis_filtered['Capacity'] = pd.to_numeric(df_jenis_filtered['Capacity'], errors='coerce').fillna(0)
    
    if selected_cap_filter == 'Di Bawah 5 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] < 5].copy()
    elif selected_cap_filter == 'Di Bawah 10 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] < 10].copy()
    elif selected_cap_filter == 'Di Bawah 15 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] < 15].copy()
    elif selected_cap_filter == 'Di Bawah 100 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] < 100].copy()
    elif selected_cap_filter == 'Di Atas 100 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] >= 100].copy()
    elif selected_cap_filter == 'Di Atas 35 Ton':
        df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] > 35].copy()
    else:
        try:
             val = float(selected_cap_filter)
             df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] == val].copy()
        except:
             df_active = df_jenis_filtered.copy()
    
    # Filter Inaktif Logic
    if df_inaktif is not None:
        df_inactive_show = df_inaktif[df_inaktif['Jenis_Alat'] == selected_jenis].copy()
    
    # --- MAIN CONTENT ---
    st.subheader(f"Analisa: {selected_jenis} - {selected_cap_filter}")
    
    if not df_inactive_show.empty:
        with st.expander(f"⚠️ {len(df_inactive_show)} Unit Tidak Masuk Analisa (Inaktif pada jenis ini)"):
            st.dataframe(df_inactive_show[['Unit_Name', 'Lokasi', 'Total_Liter', 'Total_HM_Work']]
                         .rename(columns={'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Unit_Name': 'Unit'}))
            
    if df_active.empty:
        st.warning(f"Tidak ada unit aktif untuk kategori {selected_jenis} {selected_cap_filter}.")
        st.stop()

    # --- KPI CALCULATIONS ---
    benchmark_val = df_active['Group_Benchmark_Median'].iloc[0] if 'Group_Benchmark_Median' in df_active.columns else 0
    total_waste = df_active['Potensi_Pemborosan_Liter'].sum()
    total_loss_rp = total_waste * harga_solar
    
    df_active.sort_values('Fuel_Ratio', inplace=True)
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
    m4.metric(f"Teririt: {best_unit['Unit_Name']}", f"{best_unit['Fuel_Ratio']:.2f} L/Jam")
    
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
            color = '#2ca02c' if row['Status_BBM'] == 'EFISIEN' else '#d62728' 
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
        list_unit_active = df_display_active['Unit'].unique().tolist()
        
        if list_unit_active:
            selected_unit_active = st.selectbox("Pilih Unit yang Diinginkan:", list_unit_active, key='sb_active')
            
            df_trend = load_monthly_data(selected_unit_active)
            
            if not df_trend.empty:
                fig_trend = px.line(df_trend, x='Bulan', y='Fuel_Ratio', markers=True, title=f"Pergerakan Fuel Ratio {selected_unit_active} (Jan-Nov)")
                
                # [UPDATE] Posisi Teks Top Left & Background Putih
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
        
        # [UPDATE] Posisi Teks Top Left & Background Putih
        fig_bar.add_hline(
            y=benchmark_val, 
            line_dash="dash",
            line_color="white",
            line_width=2,
            annotation_text=f"Benchmark: {benchmark_val:.2f} L/Jam",
            annotation_position="top left",
            annotation_font_color="white",
            annotation_bgcolor="rgba(0, 0, 0, 0.5)" # Hitam transparan untuk chart background gelap
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # Tab C: Scatter
    with tab_c:
        st.subheader("Peta Jam Kerja vs BBM")
        
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
            title="Sebaran Efisiensi Unit"
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