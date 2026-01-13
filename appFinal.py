import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

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
# 2. LOAD DATA
# ==============================================================================
@st.cache_data
def load_data():
    data_kpi = None
    data_unit = None
    data_inaktif = None
    
    # 1. Load Data untuk Mode KPI (File Lama)
    try:
        data_kpi = pd.read_excel('Laporan_Benchmark_BBM.xlsx')
    except FileNotFoundError:
        pass

    # 2. Load Data untuk Mode Jenis Alat (File Baru)
    possible_files = ['Analisa_Benchmark_Alat_Berat.xlsx', 'Benchmark_Per_Alat_Berat_Data_Baru.xlsx']
    
    for f in possible_files:
        try:
            data_unit = pd.read_excel(f, sheet_name='Rapor_Unit_Aktif')
            data_inaktif = pd.read_excel(f, sheet_name='Unit_Inaktif')
            break 
        except FileNotFoundError:
            continue
        except Exception:
            continue
        
    return data_kpi, data_unit, data_inaktif

df_kpi, df_unit, df_inaktif = load_data()

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
            st.dataframe(res_kpi)
        else:
            st.warning("Unit Tidak Ditemukan.")
    st.markdown("---")

    # --- MAIN CONTENT ---
    st.subheader(f"Analisa Group KPI: {selected_group}")
    
    col_durasi = next((c for c in df_kpi.columns if 'Total_Jam' in c or 'HM' in c), 'Total_Jam')
    df_view = df_kpi[df_kpi['Benchmark_Group'] == selected_group].copy()
    
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
    c3.metric("Benchmark Median", f"{avg_eff:.2f} L/Jam")
    c4.metric("Estimasi Kerugian", f"Rp {estimasi_rugi_rp:,.0f}", help=f"{total_waste_liter:,.0f} Liter Terbuang")
    
    st.markdown("---")
    
    # --- TABS ---
    tab0, tab1, tab2 = st.tabs(["📋 Overview Data", "📉 Persebaran Efisiensi Setiap Unit", "💰 Top 10 Unit Pemborosan"])
    
    with tab0:
        st.subheader(f"Detail Data: {selected_group}")
        cols_kpi_show = [c for c in ['Unit', 'Category', 'Total_Solar_Liter', col_durasi, 'Rata_Rata_Efisiensi', 'Status_BBM', 'Potensi_Pemborosan_Liter'] if c in df_view.columns]
        st.dataframe(
            df_view[cols_kpi_show].sort_values('Total_Solar_Liter', ascending=False)
            .style.format({'Total_Solar_Liter': '{:,.0f}', col_durasi: '{:,.0f}', 'Rata_Rata_Efisiensi': '{:.2f}', 'Potensi_Pemborosan_Liter': '{:,.0f}'})
        )

    # TAB 1: SCATTER MATRIX (UPDATE: Tambah Hover Potensi Pemborosan)
    with tab1:
        color_col = 'Status_BBM' if 'Status_BBM' in df_view.columns else None
        color_map = {"EFISIEN (Hijau)": "#2ca02c", "BOROS (Merah)": "#d62728"} if color_col else None
        
        # Siapkan Hover Data
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
            title="Sebaran Efisiensi Unit"
        )
        st.plotly_chart(fig, use_container_width=True)
        
    with tab2:
        if 'Potensi_Pemborosan_Liter' in df_view.columns:
            df_waste = df_view[df_view['Potensi_Pemborosan_Liter'] > 0].sort_values('Potensi_Pemborosan_Liter', ascending=False).head(10)
            if not df_waste.empty:
                fig_bar = px.bar(df_waste, x='Unit', y='Potensi_Pemborosan_Liter', text_auto='.0f', title="Top 10 Unit Boros", color_discrete_sequence=['#d62728'])
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
            st.dataframe(res_all[['Unit_Name', 'Jenis_Alat', 'Capacity', 'Lokasi', 'Status', 'Fuel_Ratio', 'Total_Liter']])
        else:
            st.warning("Unit Tidak Ditemukan.")
    
    st.markdown("---")

    # --- Sidebar Filters ---
    st.sidebar.subheader("Filter Spesifik")
    
    # 1. Pilih Jenis
    jenis_list = sorted(df_unit['Jenis_Alat'].astype(str).unique())
    selected_jenis = st.sidebar.selectbox("1. Pilih Jenis Alat:", jenis_list)
    
    # 2. Pilih Kapasitas
    df_jenis_filtered = df_unit[df_unit['Jenis_Alat'] == selected_jenis]
    cap_list = sorted(df_jenis_filtered['Capacity'].unique())
    selected_cap = st.sidebar.selectbox("2. Pilih Kapasitas (Ton):", cap_list)
    
    # 3. Input Harga Solar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Biaya Bahan Bakar")
    harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_alat')
    
    # --- FILTER FINAL ---
    df_active = df_jenis_filtered[df_jenis_filtered['Capacity'] == selected_cap].copy()
    
    # Filter Inaktif
    df_inactive_show = pd.DataFrame()
    if df_inaktif is not None:
        # Filter Jenis
        df_in_temp = df_inaktif[df_inaktif['Jenis_Alat'] == selected_jenis].copy()
        if not df_in_temp.empty:
            # Cek apakah ada kolom Capacity
            if 'Capacity' in df_in_temp.columns and df_in_temp['Capacity'].iloc[0] != "-":
                df_inactive_show = df_in_temp[df_in_temp['Capacity'].astype(str) == str(selected_cap)]
            else:
                # Fallback ke Benchmark Group string matching
                search_str = f"({selected_cap}T)"
                df_inactive_show = df_in_temp[df_in_temp['Benchmark_Group'].astype(str).str.contains(search_str, regex=False)]
    
    # --- MAIN CONTENT ---
    st.subheader(f"Analisa: {selected_jenis} - Kapasitas {selected_cap} Ton")
    
    if not df_inactive_show.empty:
        with st.expander(f"⚠️ {len(df_inactive_show)} Unit Tidak Masuk Analisa"):
            st.warning(f"Unit berikut tercatat sebagai **{selected_jenis} {selected_cap}T** tetapi memiliki Total HM=0 atau BBM=0.")
            st.dataframe(df_inactive_show[['Unit_Name', 'Lokasi', 'Total_Liter', 'Total_HM_Work']])
            
    if df_active.empty:
        st.warning(f"Tidak ada unit aktif untuk kategori {selected_jenis} {selected_cap} Ton.")
        st.stop()

    # --- KPI CALCULATIONS ---
    avg_ratio_group = df_active['Total_Liter'].sum() / df_active['Total_HM_Work'].sum()
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
    m2.metric("Benchmark Berdasarkan Filter", f"{avg_ratio_group:.2f} L/Jam")
    m3.metric("Estimasi Kerugian", f"Rp {total_loss_rp:,.0f}", help=f"{total_waste:,.0f} Liter Terbuang")
    m4.metric(f"Teririt: {best_unit['Unit_Name']}", f"{best_unit['Fuel_Ratio']:.2f} L/Jam")
    
    st.markdown("---")
    
    # --- TABS ---
    tab_a, tab_b, tab_c = st.tabs(["📋 Overview Data", "📊 Efisiensi Setiap Unit", "📉 Persebaran Efisiensi Setiap Unit"])
    
    # Tab A: Data Detail
    with tab_a:
        st.subheader("Detail Unit Aktif")
        st.info(f"**Total Pemborosan**: **{total_waste:,.0f} Liter** setara dengan **Rp {total_loss_rp:,.0f}**")
        
        st.dataframe(
            df_active[['Unit_Name', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']]
            .style.format({
                'Total_Liter': '{:,.0f}', 
                'Total_HM_Work': '{:,.0f}', 
                'Fuel_Ratio': '{:.2f}',
                'Potensi_Pemborosan_Liter': '{:,.0f}'
            })
            .background_gradient(subset=['Fuel_Ratio'], cmap='RdYlGn_r')
        )
    
    # Tab B: Peringkat
    with tab_b:
        st.subheader("Peringkat Efisiensi Setiap Unit")
        fig_bar = px.bar(
            df_active, x='Unit_Name', y='Fuel_Ratio', color='Fuel_Ratio',
            color_continuous_scale='RdYlGn_r', text_auto='.2f',
            title=f"Konsumsi BBM (Liter/Jam)"
        )
        fig_bar.add_hline(y=avg_ratio_group, line_dash="dash", annotation_text="Benchmark")
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # Tab C: Scatter (REVISI: STATUS COLOR HIJAU/MERAH SAJA)
    with tab_c:
        st.subheader("Peta Jam Kerja vs BBM")
        
        # Mapping Warna agar Efisien=Hijau, Boros=Merah (Tanpa Gradasi Angka)
        color_map_status = {"EFISIEN": "#2ca02c", "BOROS": "#d62728"}
        
        fig_scat = px.scatter(
            df_active, 
            x='Total_HM_Work', 
            y='Total_Liter', 
            color='Performance_Status', # Ganti color jadi Status Kategorikal
            size='Total_Liter',
            hover_name='Unit_Name', 
            hover_data={
                'Performance_Status': False,
                'Fuel_Ratio': ':.2f',
                'Potensi_Pemborosan_Liter': ':,.0f'
            },
            color_discrete_map=color_map_status,
            title="Sebaran Efisiensi Unit"
        )
        
        # Garis Benchmark DIHAPUS sesuai permintaan
        
        st.plotly_chart(fig_scat, use_container_width=True)