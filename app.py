import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Dashboard BBM Alat Berat",
    page_icon="⛽",
    layout="wide"
)

st.title("Dashboard Monitoring Efisiensi BBM")
st.markdown("### Analisa Performa, Tren, dan Pemborosan Bahan Bakar")

# 2. LOAD DATA
@st.cache_data
def load_data():
    try:
        # Load Data
        df_bench = pd.read_excel('Laporan_Benchmark_BBM.xlsx')
        df_trend = pd.read_excel('Laporan_Tren_Efisiensi_Bulanan_Fix.xlsx')
        return df_bench, df_trend
    except FileNotFoundError:
        return None, None

df_bench, df_trend = load_data()

# Cek File
if df_bench is None or df_trend is None:
    st.error("File Excel tidak ditemukan! Pastikan file 'Laporan_Benchmark_BBM.xlsx' dan 'Laporan_Tren_Efisiensi_Bulanan_Fix.xlsx' ada di folder yang sama.")
    st.stop()

# DETEKSI KOLOM DURASI OTOMATIS
col_durasi = next((c for c in df_bench.columns if 'Total_Jam' in c), None)

if not col_durasi:
    st.error("Gagal mendeteksi kolom Jam Kerja (Total_Jam atau Total_Jam/KM). Cek file Excel Anda.")
    st.stop()

# 3. SIDEBAR (FILTER & SIMULASI)
st.sidebar.header("Panel Kontrol")

# Filter Kategori
all_categories = df_bench['Category'].unique()
selected_category = st.sidebar.multiselect(
    "Filter Kategori Alat:",
    options=all_categories,
    default=all_categories,
    help="Pilih satu atau lebih kategori alat untuk memfilter data yang tampil di Dashboard."
)

st.sidebar.markdown("---")

# Simulasi Harga
st.sidebar.header("Simulasi Bahan Bakar")
harga_bbm = st.sidebar.number_input(
    "Harga BBM per Liter (Rp):", 
    min_value=0, 
    value=6800, 
    step=500,
    format="%d",
    help="Masukkan harga BBM yang digunakan. Angka ini berfungsi untuk menghitung estimasi kerugian yang ada."
)

# Terapkan Filter
df_filtered = df_bench[df_bench['Category'].isin(selected_category)]

# 4. TAB UTAMA
tab1, tab2, tab3 = st.tabs(["📊 Overview", "🔍 Analisa Per Unit", "📈 Scatter Plot Efisiensi"])

# TAB 1: OVERVIEW
with tab1:
    st.header("Ringkasan Performa Armada")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total_unit = len(df_filtered)
    total_boros = df_filtered[df_filtered['Status_BBM'].str.contains('BOROS', na=False)].shape[0]
    total_liter_loss = df_filtered['Potensi_Pemborosan_Liter'].sum()
    total_uang_loss = total_liter_loss * harga_bbm
    
    col1.metric(
        "Total Unit Dianalisa", 
        f"{total_unit} Unit",
        help="Jumlah total unit yang memiliki data valid (Jam Kerja > 0 dan Konsumsi BBM > 0) sesuai filter kategori."
    )
    
    col2.metric(
        "Unit Status BOROS", 
        f"{total_boros} Unit", 
        delta_color="inverse",
        help="Jumlah unit yang rata-rata konsumsi BBM-nya lebih tinggi daripada nilai Benchmark grupnya."
    )
    
    col3.metric(
        "Potensi BBM Terbuang", 
        f"{total_liter_loss:,.0f} Liter",
        help="Total liter BBM yang terbuang. Rumus: (Aktual - Benchmark) x Jam Kerja, hanya untuk unit Boros."
    )
    
    col4.metric(
        "Estimasi Kerugian (Rp)", 
        f"Rp {total_uang_loss:,.0f}",
        help=f"Nilai kerugian finansial berdasarkan potensi BBM terbuang dikalikan harga input (Rp {harga_bbm:,})."
    )
    
    st.markdown("---")
    
    st.subheader("Top 10 Unit Paling Boros")
    st.info("Daftar ini diurutkan berdasarkan **Potensi Pemborosan Liter** terbesar.")
    
    top_10 = df_filtered.sort_values(by='Potensi_Pemborosan_Liter', ascending=False).head(10)
    
    st.dataframe(
        top_10,
        column_order=['Unit', 'Group_KPI', 'Total_Solar_Liter', col_durasi, 'Rata_Rata_Efisiensi', 'Benchmark_Median', 'Potensi_Pemborosan_Liter', 'Status_BBM'],
        column_config={
            'Unit': st.column_config.TextColumn("Unit", help="Nama Unit"),
            'Group_KPI': st.column_config.TextColumn("Group", help="Kelompok KPI Unit"),
            'Total_Solar_Liter': st.column_config.NumberColumn("Total BBM (L)", help="Total Konsumsi BBM (Liter)"),
            col_durasi: st.column_config.NumberColumn("Durasi Kerja (Jam)", help="Total Jam Kerja (HM) atau Jarak Tempuh (KM)"),
            'Rata_Rata_Efisiensi': st.column_config.NumberColumn("Aktual (L/Jam)", format="%.2f", help="Konsumsi BBM Rata-rata Unit"),
            'Benchmark_Median': st.column_config.NumberColumn("Benchmark (L)", format="%.2f", help="Nilai Penentu Efisiensi Untuk Suatu Grup"),
            'Potensi_Pemborosan_Liter': st.column_config.NumberColumn("Pemborosan (L)", help="Selisih Liter yang Terbuang Melebihi Benchmark. Rumus: (Aktual - Benchmark) x Jam Kerja"),
            'Status_BBM': st.column_config.TextColumn("Status", help="Indikator Boros/Efisien-nya Unit")
        },
        use_container_width=True,
        hide_index=True
    )

# TAB 2: ANALISA PER UNIT
with tab2:
    st.header("Analisa Detail Per Unit")
    
    list_unit = sorted(df_filtered['Unit'].unique())
    if list_unit:
        col_sel1, col_sel2 = st.columns([1, 3])
        with col_sel1:
            pilih_unit = st.selectbox("Pilih Unit:", list_unit, help="Pilih unit spesifik untuk melihat grafik trennya.")
        
        trend_data = df_trend[df_trend['Unit'] == pilih_unit]
        bench_data = df_filtered[df_filtered['Unit'] == pilih_unit]
        
        if not bench_data.empty:
            nilai_benchmark = bench_data['Benchmark_Median'].values[0]
            status_sekarang = bench_data['Status_BBM'].values[0]
            
            if "BOROS" in status_sekarang:
                st.error(f"Status Saat Ini: **{status_sekarang}**")
            else:
                st.success(f"Status Saat Ini: **{status_sekarang}**")
            
            # Grafik Tren
            if not trend_data.empty:
                cols_bulan = [c for c in trend_data.columns if '20' in str(c)]
                y_vals = trend_data[cols_bulan].values.flatten()
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=cols_bulan, y=y_vals, mode='lines+markers', name='Aktual', line=dict(color='blue', width=3)))
                fig.add_trace(go.Scatter(x=cols_bulan, y=[nilai_benchmark]*len(cols_bulan), mode='lines', name=f'Benchmark ({nilai_benchmark})', line=dict(color='red', width=2, dash='dash')))
                
                fig.update_layout(
                    title=f"Tren Efisiensi: {pilih_unit}", 
                    xaxis_title="Bulan", 
                    yaxis_title="Konsumsi BBM (Liter/Jam)",
                    hovermode="x unified"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                status_tren = trend_data['Status_Tren'].values[0]
            else:
                st.warning("Data tren tidak tersedia.")
    else:
        st.warning("Tidak ada unit yang sesuai filter.")

# TAB 3: SCATTER PLOT
with tab3:
    st.header("Scatter Plot Efisiensi BBM Alat Berat")
    st.markdown("""
    **Cara Membaca Grafik:**
    * **Sumbu X:** Total Jam Kerja
    * **Sumbu Y:** Total Penggunaan BBM
    * **Titik Hijau:** Unit Efisien
    * **Titik Merah:** Unit Boros
    * **Ukuran Titik:** Semakin besar titik, semakin besar potensi pemborosan atau penghematan BBM
    """)
    
    # 1. Filter Data
    scatter_data = df_filtered[df_filtered[col_durasi] > 0].copy()
    
    if not scatter_data.empty:
        # 2. Hitung Nilai Visual
        scatter_data['Gap'] = scatter_data['Rata_Rata_Efisiensi'] - scatter_data['Benchmark_Median']
        scatter_data['Nilai_Visual'] = (scatter_data['Gap'] * scatter_data[col_durasi]).abs()
        scatter_data['Ukuran_Final'] = scatter_data['Nilai_Visual'] + 50

        # 3. Sorting
        scatter_data.sort_values(by='Status_BBM', ascending=False, inplace=True) 

        # 4. Tooltip Informatif
        def buat_teks_tooltip(row):
            base_text = (
                f"<b>{row['Unit']}</b><br>"
                f"Grup: {row['Group_KPI']}<br>"
                f"Penggunaan: {row['Rata_Rata_Efisiensi']} L/Jam<br>"
                f"Benchmark: {row['Benchmark_Median']} L/Jam<br>"
            )
            if "BOROS" in row['Status_BBM']:
                return base_text + f"🔴 Pemborosan: {row['Potensi_Pemborosan_Liter']:,.0f} Liter"
            else:
                hemat = (row['Benchmark_Median'] - row['Rata_Rata_Efisiensi']) * row[col_durasi]
                return base_text + f"🟢 Penghematan: {hemat:,.0f} Liter"

        scatter_data['Info_Lengkap'] = scatter_data.apply(buat_teks_tooltip, axis=1)

        fig_scatter = px.scatter(
            scatter_data,
            x=col_durasi,
            y="Total_Solar_Liter",
            color="Status_BBM",
            custom_data=['Info_Lengkap'],
            size="Ukuran_Final", 
            color_discrete_map={
                "EFISIEN (Hijau)": "#2ca02c", 
                "BOROS (Merah)": "#d62728"
            },
            size_max=70, 
            height=650
        )
        
        fig_scatter.update_traces(hovertemplate="%{customdata[0]}<extra></extra>")
        
        fig_scatter.update_layout(
            xaxis_title=f"Total Durasi Kerja ({col_durasi})",
            yaxis_title="Total Konsumsi BBM (Liter)",
            legend_title="Status Efisiensi"
        )
        
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.warning("Tidak ada data valid untuk Scatter Plot")