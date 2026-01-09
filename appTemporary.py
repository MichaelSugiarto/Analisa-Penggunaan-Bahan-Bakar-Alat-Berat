import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ==============================================================================
# 1. KONFIGURASI HALAMAN
# ==============================================================================
st.set_page_config(
    page_title="Dashboard BBM Alat Berat",
    page_icon="🚜",
    layout="wide"
)

st.title("Dashboard Monitoring Efisiensi BBM")
#st.markdown("Pusat analisa performa alat berat dengan dua metode pendekatan Benchmark.")

# ==============================================================================
# 2. LOAD DATA
# ==============================================================================
@st.cache_data
def load_data():
    data_kpi = None
    data_unit = None
    
    # 1. Load Data untuk Mode KPI (File Lama)
    try:
        data_kpi = pd.read_excel('Laporan_Benchmark_BBM.xlsx')
    except FileNotFoundError:
        pass

    # 2. Load Data untuk Mode Jenis Alat (File Baru - Rapor Unit)
    try:
        data_unit = pd.read_excel('Analisa_Benchmark_Per_Alat_Berat.xlsx', sheet_name='Rapor_Per_Unit')
    except FileNotFoundError:
        pass
        
    return data_kpi, data_unit

df_kpi, df_unit = load_data()

# ==============================================================================
# 3. SIDEBAR: KONFIGURASI GLOBAL
# ==============================================================================
st.sidebar.header("Pengaturan Benchmark")

# Pilihan Mode (Diganti "Benchmark Group (KPI)" -> "Group KPI")
analysis_mode = st.sidebar.radio(
    "Pilih Benchmark:",
    ["Group KPI", "Jenis Alat Berat"]
)

st.sidebar.markdown("---")

# Inisialisasi variabel agar bisa dipakai di bawah
selected_group = None
harga_solar = 0

# --- LOGIC SIDEBAR KHUSUS GROUP KPI ---
if analysis_mode == "Group KPI":
    if df_kpi is not None:
        # 1. FILTER GROUP (DILETAKKAN DI ATAS SECTION BIAYA)
        groups = sorted(df_kpi['Benchmark_Group'].astype(str).unique())
        selected_group = st.sidebar.selectbox("Pilih Benchmark Group:", groups)
        
        st.sidebar.markdown("---")
        
        # 2. INPUT BIAYA (DI BAWAH FILTER)
        st.sidebar.subheader("Biaya Bahan Bakar")
        harga_solar = st.sidebar.number_input(
            "Harga (IDR/Liter):", 
            min_value=0, 
            value=6800, 
            step=100,
            help="Digunakan untuk menghitung estimasi kerugian."
        )
        st.sidebar.markdown("---")

# ==============================================================================
# MODE A: GROUP KPI (LOGIKA KPI)
# ==============================================================================
if analysis_mode == "Group KPI":
    
    if df_kpi is None:
        st.error("⚠️ File 'Laporan_Benchmark_BBM.xlsx' tidak ditemukan.")
        st.stop()

    st.subheader("Analisa Berdasarkan Group KPI")
    
    # --- FILTER (Sudah dilakukan di Sidebar) ---
    col_durasi = next((c for c in df_kpi.columns if 'Total_Jam' in c or 'HM' in c), 'Total_Jam')
    
    # Filter Data berdasarkan selected_group dari sidebar
    df_view = df_kpi[df_kpi['Benchmark_Group'] == selected_group].copy()
    
    # --- KPI CARDS ---
    total_solar = df_view['Total_Solar_Liter'].sum()
    total_jam = df_view[col_durasi].sum()
    avg_eff = total_solar / total_jam if total_jam > 0 else 0
    populasi = df_view['Unit'].nunique()
    
    # Hitung Kerugian
    total_waste_liter = 0
    if 'Potensi_Pemborosan_Liter' in df_view.columns:
        total_waste_liter = df_view['Potensi_Pemborosan_Liter'].sum()
    
    estimasi_rugi_rp = total_waste_liter * harga_solar
    
    # Tampilkan 4 Metric
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Populasi Group", f"{populasi} Unit")
    c2.metric("Total Solar", f"{total_solar:,.0f} Liter")
    c3.metric("Benchmark Median", f"{avg_eff:.2f} L/Jam", help="Benchmark dari tipe filter yang dipilih")
    
    # Estimasi Kerugian (TANPA PANAH / DELTA)
    # Info liter terbuang dipindah ke help/tooltip agar tidak memunculkan panah
    c4.metric(
        "Estimasi Kerugian", 
        f"Rp {estimasi_rugi_rp:,.0f}", 
        help=f"{total_waste_liter:,.0f} Liter Terbuang"
    )
    st.markdown("---")
    
    # --- VISUALISASI ---
    tab0, tab1, tab2 = st.tabs(["📋 Overview Data", "📉 Scatter Matrix", "💰 Potensi Pemborosan"])
    
    # TAB 0: Overview (Tabel)
    with tab0:
        st.subheader(f"Data Detail: {selected_group}")
        cols_show = ['Unit', 'Category', 'Total_Solar_Liter', col_durasi, 'Rata_Rata_Efisiensi', 'Status_BBM', 'Potensi_Pemborosan_Liter']
        cols_final = [c for c in cols_show if c in df_view.columns]
        
        st.dataframe(
            df_view[cols_final].sort_values('Total_Solar_Liter', ascending=False)
            .style.format({
                'Total_Solar_Liter': '{:,.0f}',
                col_durasi: '{:,.0f}',
                'Rata_Rata_Efisiensi': '{:.2f}',
                'Potensi_Pemborosan_Liter': '{:,.0f}'
            })
        )

    # TAB 1: Scatter
    with tab1:
        st.subheader(f"Sebaran Efisiensi: {selected_group}")
        if 'Status_BBM' in df_view.columns:
            color_col = 'Status_BBM'
            color_map = {"EFISIEN (Hijau)": "#2ca02c", "BOROS (Merah)": "#d62728"}
        else:
            color_col = None
            color_map = None

        fig_scat = px.scatter(
            df_view,
            x=col_durasi,
            y="Total_Solar_Liter",
            color=color_col,
            size="Total_Solar_Liter",
            hover_name="Unit",
            color_discrete_map=color_map,
            title="Total Jam Kerja vs Total Solar"
        )
        st.plotly_chart(fig_scat, use_container_width=True)
        
    # TAB 2: Waste
    with tab2:
        st.subheader("Analisa Pemborosan (Liter)")
        if 'Potensi_Pemborosan_Liter' in df_view.columns:
            df_waste = df_view[df_view['Potensi_Pemborosan_Liter'] > 0].sort_values('Potensi_Pemborosan_Liter', ascending=False)
            
            if not df_waste.empty:
                fig_bar = px.bar(
                    df_waste.head(10),
                    x='Unit',
                    y='Potensi_Pemborosan_Liter',
                    text_auto='.0f',
                    title="Top 10 Unit Paling Boros (vs Median Group)",
                    color_discrete_sequence=['#d62728']
                )
                fig_bar.update_layout(yaxis_title="Liter Terbuang")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.success("✅ Tidak ada unit yang terdeteksi boros di grup ini.")
        else:
            st.warning("Kolom 'Potensi_Pemborosan_Liter' tidak ditemukan.")

# ==============================================================================
# MODE B: PER JENIS ALAT BERAT (UNIT vs UNIT)
# ==============================================================================
elif analysis_mode == "Jenis Alat Berat":
    
    if df_unit is None:
        st.error("⚠️ File 'Analisa_Benchmark_Alat_Berat.xlsx' tidak ditemukan.")
        st.stop()
        
    col_liter = 'Total_Liter'
    col_hm = 'Total_HM_Work'
    col_ratio = 'Fuel_Ratio'
    
    st.subheader("Analisa Per Jenis Alat Berat")
    
    # --- FILTER (SINGLE SELECT) ---
    jenis_list = sorted(df_unit['Jenis_Alat'].astype(str).unique())
    selected_jenis = st.sidebar.selectbox("Pilih Jenis Alat Berat:", jenis_list)
    
    # Filter Data
    df_active = df_unit[df_unit['Jenis_Alat'] == selected_jenis].copy()
    
    # Pisahkan Unit Aktif vs Inaktif
    df_visual = df_active[
        (df_active[col_liter] > 0) & 
        (df_active[col_hm] > 0)
    ].copy()
    
    df_inactive = df_active[~df_active['Unit_Name'].isin(df_visual['Unit_Name'])]
    
    # --- NOTES / WARNING JIKA ADA DATA 0 ---
    if not df_inactive.empty:
        with st.expander(f"⚠️ Notes: {len(df_inactive)} Unit Tidak Ditampilkan", expanded=True):
            st.warning(
                f"Unit berikut memiliki **Total HM = 0** atau **Total BBM = 0** pada kategori **{selected_jenis}**. "
                f"Data ini dikecualikan dari grafik dan perhitungan ranking."
            )
            st.dataframe(
                df_inactive[['Unit_Name', col_liter, col_hm]]
                .style.format({col_liter: '{:,.0f}', col_hm: '{:,.0f}'})
            )

    if df_visual.empty:
        st.warning(f"Tidak ada unit aktif untuk jenis alat: {selected_jenis}")
        st.stop()

    # --- HITUNG BENCHMARK INTERNAL ---
    total_liter_group = df_visual[col_liter].sum()
    total_hm_group = df_visual[col_hm].sum()
    
    avg_ratio_group = total_liter_group / total_hm_group if total_hm_group > 0 else 0

    # --- KPI LOGIC (BEST vs WORST) ---
    df_visual.sort_values(col_ratio, inplace=True)
    
    populasi_aktif = len(df_visual)
    
    best_unit = df_visual.iloc[0]
    best_name = best_unit['Unit_Name']
    best_val = best_unit[col_ratio]
    
    # Logika jika populasi cuma 1
    if populasi_aktif > 1:
        worst_unit = df_visual.iloc[-1]
        worst_name = worst_unit['Unit_Name']
        worst_val = worst_unit[col_ratio]
    else:
        worst_name = "-"
        worst_val = 0

    # TAMPILKAN KPI (TANPA PANAH/DELTA)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Populasi Aktif", f"{populasi_aktif} Unit")
    
    # Ubah Label sesuai request
    c2.metric("Benchmark Median", f"{avg_ratio_group:.2f} L/Jam", help="Benchmark dari tipe filter yang dipilih")
    
    # Unit Teririt (Hanya Nama, Angka di Label/Help)
    c3.metric(
        label=f"Unit Teririt ({best_val:.2f} L/Jam)", 
        value=f"{best_name}"
    )
    
    # Unit Terboros
    worst_label_val = f"({worst_val:.2f} L/Jam)" if worst_name != "-" else ""
    c4.metric(
        label=f"Unit Terboros {worst_label_val}", 
        value=f"{worst_name}"
    )
    st.markdown("---")
    
    # --- VISUALISASI ---
    tab_a, tab_b, tab_c = st.tabs(["📋 Overview Data", "📊 Peringkat Unit", "📉 Peta Aktivitas"])
    
    # TAB A: Overview Tabel
    with tab_a:
        st.subheader(f"Overview Unit: {selected_jenis}")
        st.dataframe(
            df_visual[['Unit_Name', col_liter, col_hm, col_ratio, 'Performance_Status']]
            .sort_values(col_ratio)
            .style.format({
                col_liter: '{:,.0f}', 
                col_hm: '{:,.0f}', 
                col_ratio: '{:.2f}'
            })
            .background_gradient(subset=[col_ratio], cmap='RdYlGn_r')
        )
    
    # TAB B: Peringkat
    with tab_b:
        st.subheader(f"Peringkat Efisiensi: {selected_jenis}")
        
        # Grafik
        fig_bar = px.bar(
            df_visual,
            x='Unit_Name',
            y=col_ratio,
            color=col_ratio,
            color_continuous_scale='RdYlGn_r',
            text_auto='.2f',
            title=f"Konsumsi BBM per Unit (Liter/Jam)"
        )
        fig_bar.add_hline(y=avg_ratio_group, line_dash="dash", line_color="black", annotation_text="Benchmark")
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # TAB C: Scatter
    with tab_c:
        st.subheader("Scatter Plot: Jam Kerja vs Total BBM")
        
        fig_scat = px.scatter(
            df_visual,
            x=col_hm,
            y=col_liter,
            color=col_ratio,
            size=col_ratio,
            hover_name='Unit_Name',
            color_continuous_scale='RdYlGn_r',
            title="Posisi Unit berdasarkan Jam Kerja"
        )
        x_max = df_visual[col_hm].max() * 1.1
        fig_scat.add_trace(go.Scatter(
            x=[0, x_max],
            y=[0, x_max * avg_ratio_group],
            mode='lines',
            name='Garis Benchmark',
            line=dict(color='gray', dash='dash')
        ))
        
        st.plotly_chart(fig_scat, use_container_width=True)