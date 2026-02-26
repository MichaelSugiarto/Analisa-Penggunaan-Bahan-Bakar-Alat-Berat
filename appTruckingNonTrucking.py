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
FILE_HASIL_TRUCKING = "HASIL_ANALISA_TRUCKING_OKT_NOV.xlsx" 
FILE_HASIL_NON_TRUCKING = "HasilNonTrucking.xlsx"
FILE_REKAP_BBM_ALL = "Rekap_BBM_All.xlsx" 

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

# ==============================================================================
# 4. LOGIKA PROSES DATA: NON-TRUCKING 
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_alat_berat():
    if not os.path.exists(FILE_HASIL_NON_TRUCKING):
        st.warning(f"File {FILE_HASIL_NON_TRUCKING} tidak ditemukan. Jalankan script Jupyter terbaru.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        df_agg = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Total_Agregat')
        df_monthly = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Data_Bulanan')
        try:
            df_missing = pd.read_excel(FILE_HASIL_NON_TRUCKING, sheet_name='Unit_Inaktif')
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

        return df_agg, df_monthly, df_missing
    except Exception as e:
        st.error(f"Error memproses data Non-Trucking: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==============================================================================
# 5. LOGIKA PROSES DATA: TRUCKING 
# ==============================================================================
@st.cache_data(show_spinner=False)
def process_trucking():
    if os.path.exists(FILE_HASIL_TRUCKING):
        try:
            df_trucking = pd.read_excel(FILE_HASIL_TRUCKING)
            
            # --- PENARIKAN DATA LITER BBM MURNI DARI JUPYTER (TRONTON & TRAILER) ---
            if os.path.exists(FILE_REKAP_BBM_ALL):
                df_bbm_all = pd.read_excel(FILE_REKAP_BBM_ALL)
                df_bbm_trucking = df_bbm_all[df_bbm_all['Bulan'].isin(['Oktober', 'November'])].copy()
                bbm_sum = df_bbm_trucking.groupby('Unit_Name')['LITER'].sum().reset_index()
                
                bbm_sum['Unit_ID'] = bbm_sum['Unit_Name'].apply(clean_unit_name)
                df_trucking['Unit_ID'] = df_trucking['Unit_Name'].apply(clean_unit_name)
                
                if 'LITER' in df_trucking.columns:
                    df_trucking.drop(columns=['LITER'], inplace=True)
                    
                df_trucking = pd.merge(df_trucking, bbm_sum[['Unit_ID', 'LITER']], on='Unit_ID', how='left')
                df_trucking['LITER'] = df_trucking['LITER'].fillna(0)
                df_trucking.drop(columns=['Unit_ID'], inplace=True)

            if 'Total_TonKm' in df_trucking.columns and 'LITER' in df_trucking.columns:
                df_trucking['Fuel Ratio'] = np.where(df_trucking['Total_TonKm'] > 0, df_trucking['LITER'] / df_trucking['Total_TonKm'], 0)
            elif 'L_per_TonKm' in df_trucking.columns:
                df_trucking['Fuel Ratio'] = df_trucking['L_per_TonKm']
                
            if 'Benchmark' not in df_trucking.columns and 'Fuel Ratio' in df_trucking.columns:
                 df_trucking['Benchmark'] = df_trucking['Fuel Ratio'].median()
            
            if 'Fuel Ratio' in df_trucking.columns and 'Benchmark' in df_trucking.columns:
                df_trucking['Status'] = df_trucking.apply(lambda x: "Efisien" if x['Fuel Ratio'] <= x['Benchmark'] else "Boros", axis=1)
            else:
                df_trucking['Status'] = "Inaktif"
            
            df_trucking['Potensi Pemborosan BBM'] = 0
            if 'Lokasi' not in df_trucking.columns: df_trucking['Lokasi'] = "Trucking Pool"
            if 'Jenis_Alat' not in df_trucking.columns: df_trucking['Jenis_Alat'] = "Trucking"
            if 'Type_Merk' not in df_trucking.columns: df_trucking['Type/Merk'] = "Trucking"
            
            rename_truck_map = {
                'Unit_Name': 'Nama Unit',
                'Jenis_Alat': 'Jenis',
                'LITER': 'Total Pengisian BBM (L)',
                'Total_Ton': 'Total Berat Angkutan (Ton)',
            }
            df_trucking.rename(columns=rename_truck_map, inplace=True)
            
            df_monthly_trucking = pd.DataFrame() 
            df_missing_truck = pd.DataFrame()
            return df_trucking, df_monthly_trucking, df_missing_truck
        except Exception as e: 
            st.error(f"Error membaca {FILE_HASIL_TRUCKING}: {e}")
            pass 
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==============================================================================
# 6. SIDEBAR & FILTER
# ==============================================================================
st.sidebar.subheader("Filter Kategori Unit")
category_filter = st.sidebar.radio("Pilih Mode:", ["Trucking (Tronton/Trailer)", "Non-Trucking"])

st.sidebar.markdown("---")
BIAYA_PER_LITER = st.sidebar.number_input("Biaya Bahan Bakar (Rp/Liter)", min_value=0, value=6800, step=100)

df_active_raw = pd.DataFrame()
df_monthly = pd.DataFrame()
df_missing = pd.DataFrame()

if category_filter == "Trucking (Tronton/Trailer)":
    with st.spinner("Memproses Data Trucking..."):
        df_active_raw, df_monthly, df_missing = process_trucking()
        mode_label = "Trucking"
        ratio_label = "L/Ton-Km"
        work_col = "Total_TonKm"
else:
    with st.spinner("Memuat Data Non-Trucking..."):
        df_active_raw, df_monthly, df_missing = process_alat_berat()
        mode_label = "Non-Trucking"
        ratio_label = "L/Ton"
        work_col = "Total Berat Angkutan (Ton)"

if not df_active_raw.empty:
    
    # --- PISAHKAN DATA INAKTIF DARI DASHBOARD UTAMA ---
    df_inaktif_from_active = df_active_raw[(df_active_raw['Total Pengisian BBM (L)'] <= 0) | (df_active_raw[work_col] <= 0)].copy()
    if not df_inaktif_from_active.empty:
        df_inaktif_from_active['Keterangan'] = np.where(
            df_inaktif_from_active[work_col] <= 0,
            "Unit tidak melakukan aktivitas kerja",
            "Unit tidak pernah mengisi BBM"
        )
    
    list_inaktif = []
    if not df_inaktif_from_active.empty: list_inaktif.append(df_inaktif_from_active)
    
    if not df_missing.empty:
        # Menyamakan nama kolom dari Jupyter agar sejajar saat digabungkan dengan df_active_raw
        df_missing.rename(columns={'Total Pengisian BBM': 'Total Pengisian BBM (L)'}, inplace=True)
        
        # Mengubah keterangan spesifik yang ditarik dari script jupyter
        df_missing['Keterangan'] = df_missing['Keterangan'].replace(
            "Tidak ada aktivitas (BBM 0 & Tonase 0)", "Unit tidak digunakan"
        )
        list_inaktif.append(df_missing)

    df_inaktif_all = pd.concat(list_inaktif, ignore_index=True) if list_inaktif else pd.DataFrame()
    
    # df_active sekarang MURNI HANYA berisi unit yang aktif
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
    # TABEL UNIT INAKTIF (SEBELUM DASHBOARD UTAMA)
    # ==============================================================================
    st.markdown("### 🛑 Daftar Unit Inaktif")
    st.caption("Unit yang terdeteksi tidak aktif karena tidak ada pengisian BBM atau tidak ada aktivitas kerja yang berlangsung")
    
    df_inaktif_filtered = df_inaktif_all.copy()
    if not df_inaktif_filtered.empty:
        if selected_lokasi != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Lokasi'] == selected_lokasi]
        if selected_jenis != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Jenis'] == selected_jenis]
        if selected_type != "Semua": df_inaktif_filtered = df_inaktif_filtered[df_inaktif_filtered['Type/Merk'] == selected_type]
        
        if not df_inaktif_filtered.empty:
            cols_inaktif = ['Nama Unit', 'Jenis', 'Type/Merk', 'Lokasi', 'Total Pengisian BBM (L)', work_col, 'Keterangan']
            cols_to_show = [c for c in cols_inaktif if c in df_inaktif_filtered.columns]
            st.dataframe(df_inaktif_filtered[cols_to_show], use_container_width=True)
        else:
            st.success("Tidak ada unit inaktif untuk kombinasi filter ini.")
    else:
        st.success("Seluruh unit beroperasi aktif dan masuk ke dalam perhitungan dashboard.")
        
    st.markdown("---")
    
    # ==============================================================================
    # PENCARIAN SPESIFIK & APPLY FILTER KE DASHBOARD
    # ==============================================================================
    st.markdown("### 🔍 Cari Data Spesifik (Unit Aktif)")
    c_search1, c_search2 = st.columns([1, 3])
    with c_search1:
        search_category = st.selectbox("Cari Berdasarkan:", ["Nama Unit"])
    with c_search2:
        search_query = st.text_input(f"Ketik {search_category}:", "")

    # APPLY SEMUA FILTER KE DATA AKTIF DASHBOARD
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

else:
    df_filtered = pd.DataFrame()
    df_monthly_filtered = pd.DataFrame()

# ==============================================================================
# 7. MAIN DASHBOARD CONTENT
# ==============================================================================
if not df_filtered.empty:
    
    if mode_label == "Trucking":
        total_bbm = df_filtered['Total Pengisian BBM (L)'].sum()
        total_kerja = df_filtered['Total_TonKm'].sum()
        total_biaya = df_filtered['Total Biaya BBM'].sum()
        total_unit_aktif = len(df_filtered)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Unit Aktif", f"{total_unit_aktif} Unit")
        c2.metric("Total Kerja (Ton-Km)", f"{total_kerja:,.0f}")
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
        
        sort_options = ["Fuel Ratio (Tertinggi)", "Fuel Ratio (Terendah)", "Total Berat Angkutan (Tertinggi)", "Total Pengisian BBM (L) (Tertinggi)"]
        sort_by = st.selectbox("Sort by:", sort_options)
        
        work_col = 'Total_TonKm' if mode_label == "Trucking" else 'Total Berat Angkutan (Ton)'
        ratio_col = 'Fuel Ratio' if mode_label == "Trucking" else 'Fuel Ratio (L/Ton)'
        bm_col = 'Benchmark' if mode_label == "Trucking" else 'Benchmark (L/Ton)'

        if sort_by == "Fuel Ratio (Tertinggi)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=False)
        elif sort_by == "Fuel Ratio (Terendah)": df_filtered = df_filtered.sort_values(by=ratio_col, ascending=True)
        elif sort_by == "Total Berat Angkutan (Tertinggi)": df_filtered = df_filtered.sort_values(by=work_col, ascending=False)
        elif sort_by == "Total Pengisian BBM (L) (Tertinggi)": df_filtered = df_filtered.sort_values(by='Total Pengisian BBM (L)', ascending=False)
        
        def highlight_fuel_ratio(row):
            styles = [''] * len(row)
            for i, col in enumerate(row.index):
                if col == 'Fuel Ratio (L/Ton)' or col == 'Fuel Ratio':
                    val = row[col]
                    bm = row[bm_col]
                    if pd.notna(val) and pd.notna(bm) and bm > 0:
                        if val > bm:
                            styles[i] = 'background-color: #d62728; color: white; font-weight: bold;' 
                        else:
                            styles[i] = 'background-color: #2ca02c; color: white; font-weight: bold;'
            return styles

        if mode_label == "Trucking":
            cols_show = ['Nama Unit', 'Lokasi', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Total_TonKm', 'Benchmark', 'Fuel Ratio']
            format_dict = {'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Total_TonKm': '{:,.0f}', 'Benchmark': '{:.4f}', 'Fuel Ratio': '{:.4f}'}
            st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))
        else:
            cols_show = ['Nama Unit', 'Jenis', 'Type/Merk', 'Horse Power', 'Capacity (Ton)', 'Lokasi', 'Total Pengisian BBM (L)', 'Total Biaya BBM', 'Total Berat Angkutan (Ton)', 'Benchmark (L/Ton)', 'Fuel Ratio (L/Ton)', 'Potensi Pemborosan BBM (L)']
            format_dict = {'Total Pengisian BBM (L)': '{:,.0f}', 'Total Biaya BBM': 'Rp {:,.0f}', 'Total Berat Angkutan (Ton)': '{:,.0f}', 'Benchmark (L/Ton)': '{:.4f}', 'Fuel Ratio (L/Ton)': '{:.4f}', 'Potensi Pemborosan BBM (L)': '{:,.0f}'}
            st.dataframe(df_filtered[cols_show].style.apply(highlight_fuel_ratio, axis=1).format(format_dict))

        # TREN BULANAN PER UNIT
        st.markdown("---")
        st.subheader("📈 Tren Kinerja Bulanan Setiap Unit")
        if not df_monthly_filtered.empty and mode_label == "Non-Trucking":
            unit_list_trend = sorted(df_monthly_filtered['Nama Unit'].unique().tolist())
            selected_unit_trend = st.selectbox("Pilih Unit untuk melihat tren:", unit_list_trend)
            
            trend_data_unit = df_monthly_filtered[df_monthly_filtered['Nama Unit'] == selected_unit_trend]
            
            month_order = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
            trend_data_unit['Bulan'] = pd.Categorical(trend_data_unit['Bulan'], categories=month_order, ordered=True)
            
            trend_data = trend_data_unit.groupby('Bulan', as_index=False).agg({'Total Berat Angkutan (Ton)': 'sum', 'Total Pengisian BBM (L)': 'sum'})
            trend_data = trend_data.dropna()
            
            trend_data['Fuel Ratio (L/Ton)'] = np.where(trend_data['Total Berat Angkutan (Ton)'] > 0, trend_data['Total Pengisian BBM (L)'] / trend_data['Total Berat Angkutan (Ton)'], 0)
            
            c_trend1, c_trend2 = st.columns(2)
            with c_trend1:
                fig_trend_ton = px.bar(trend_data, x='Bulan', y='Total Berat Angkutan (Ton)', text_auto='.2s', title=f"Tren Berat Angkutan: {selected_unit_trend}", color_discrete_sequence=['#1f77b4'])
                st.plotly_chart(fig_trend_ton, use_container_width=True)
            with c_trend2:
                fig_trend_ratio = px.line(trend_data, x='Bulan', y='Fuel Ratio (L/Ton)', markers=True, title=f"Tren Efisiensi (L/Ton): {selected_unit_trend}", color_discrete_sequence=['#ff7f0e'])
                fig_trend_ratio.update_yaxes(rangemode="tozero")
                
                bm_val = df_filtered[df_filtered['Nama Unit'] == selected_unit_trend][bm_col].iloc[0] if not df_filtered[df_filtered['Nama Unit'] == selected_unit_trend].empty else 0
                if bm_val > 0:
                    fig_trend_ratio.add_hline(y=bm_val, line_dash="dash", line_color="red", annotation_text=f"Benchmark: {bm_val:.4f}", annotation_position="bottom right")
                
                st.plotly_chart(fig_trend_ratio, use_container_width=True)
        elif mode_label == "Trucking":
            st.info("Data tren bulanan spesifik belum tersedia untuk mode Trucking di versi ini.")

    # TAB 2: BAR CHART
    with tab2:
        st.subheader(f"Peringkat Efisiensi BBM ({ratio_label})")
        df_chart = df_filtered.sort_values(ratio_col, ascending=False)
        
        fig_bar = px.bar(df_chart, x='Nama Unit', y=ratio_col, color='Status',
                         custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', bm_col, ratio_col],
                         title=f"Fuel Ratio Seluruh Unit ({ratio_label})",
                         color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'})
        
        if mode_label == "Non-Trucking":
            fig_bar.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark (L/Ton):</b> %{customdata[4]:.4f}<br>" +
                              "<b>Fuel Ratio (L/Ton):</b> %{customdata[5]:.4f}<extra></extra>"
            )
        else:
            fig_bar.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                              "<b>Fuel Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
            )
        st.plotly_chart(fig_bar, use_container_width=True)

    # TAB 3: SCATTER PLOT
    with tab3:
        st.subheader("Korelasi Beban Kerja vs BBM")
        
        max_bubble_size = 45 
        
        if mode_label == "Trucking":
            size_col = df_filtered['Fuel Ratio'].apply(lambda x: x if x > 0 else 0.0001)
            
            fig_scat = px.scatter(df_filtered, x='Total_TonKm', y='Total Pengisian BBM (L)', color='Status',
                                custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', 'Benchmark', 'Fuel Ratio'],
                                size=size_col, 
                                size_max=max_bubble_size, opacity=0.65,
                                color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                title="Korelasi Beban Kerja (Total Ton-Km) vs Total Pengisian BBM (L)")
            fig_scat.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark:</b> %{customdata[4]:.4f}<br>" +
                              "<b>Fuel Ratio:</b> %{customdata[5]:.4f}<extra></extra>"
            )
            st.plotly_chart(fig_scat, use_container_width=True)
        else:
            size_col_ton = df_filtered['Fuel Ratio (L/Ton)'].apply(lambda x: x if x > 0 else 0.0001)
            
            fig_scat_ton = px.scatter(df_filtered, x='Total Berat Angkutan (Ton)', y='Total Pengisian BBM (L)', color='Status',
                                    custom_data=['Status', 'Nama Unit', 'Jenis', 'Lokasi', 'Benchmark (L/Ton)', 'Fuel Ratio (L/Ton)'],
                                    size=size_col_ton,
                                    size_max=max_bubble_size, opacity=0.65,
                                    color_discrete_map={'Efisien': '#2ca02c', 'Boros': '#d62728'},
                                    title="Korelasi Total Berat Angkutan (Ton) vs Total Pengisian BBM (L)")
            fig_scat_ton.update_traces(
                hovertemplate="<b>Status:</b> %{customdata[0]}<br>" +
                              "<b>Nama Unit:</b> %{customdata[1]}<br>" +
                              "<b>Jenis:</b> %{customdata[2]}<br>" +
                              "<b>Lokasi:</b> %{customdata[3]}<br>" +
                              "<b>Benchmark (L/Ton):</b> %{customdata[4]:.4f}<br>" +
                              "<b>Fuel Ratio (L/Ton):</b> %{customdata[5]:.4f}<extra></extra>"
            )
            st.plotly_chart(fig_scat_ton, use_container_width=True)

    # TAB 4: PEMBOROSAN
    with tab4:
        st.subheader("Analisa Pemborosan BBM")
        if mode_label == "Trucking":
            st.info("Fitur 'Potensi Pemborosan BBM' belum tersedia untuk Trucking karena Benchmark masih dinamis berdasarkan Ton-Km.")
        else:
            df_boros = df_filtered[df_filtered['Potensi Pemborosan BBM (L)'] > 0].sort_values('Potensi Pemborosan BBM (L)', ascending=False)
            if not df_boros.empty:
                df_boros['Potensi Kerugian Rp'] = df_boros['Potensi Pemborosan BBM (L)'] * BIAYA_PER_LITER
                fig_waste = px.bar(df_boros, x='Potensi Pemborosan BBM (L)', y='Nama Unit', orientation='h',
                                   custom_data=['Nama Unit', 'Jenis', 'Lokasi', 'Potensi Pemborosan BBM (L)', 'Potensi Kerugian Rp'],
                                   title="Unit Terboros dalam Penggunaan BBM", 
                                   color_discrete_sequence=['#c0392b'])
                # Format Urutan Hover Analisa Pemborosan
                fig_waste.update_traces(
                    hovertemplate="<b>Nama Unit:</b> %{customdata[0]}<br>" +
                                  "<b>Jenis:</b> %{customdata[1]}<br>" +
                                  "<b>Lokasi:</b> %{customdata[2]}<br>" +
                                  "<b>Potensi Pemborosan BBM (L):</b> %{customdata[3]:,.0f} L<br>" +
                                  "<b>Potensi Kerugian:</b> Rp %{customdata[4]:,.0f}<extra></extra>"
                )
                st.plotly_chart(fig_waste, use_container_width=True)
            else:
                st.success("Tidak ada unit yang terindikasi boros signifikan dibandingkan median jenisnya.")

elif not df_active_raw.empty and df_filtered.empty:
    st.warning("⚠️ Tidak ada unit yang cocok dengan kombinasi filter & pencarian Anda.")
else:
    st.error("Data tidak ditemukan! Pastikan script Jupyter telah dijalankan dan menghasilkan file bulanan.")