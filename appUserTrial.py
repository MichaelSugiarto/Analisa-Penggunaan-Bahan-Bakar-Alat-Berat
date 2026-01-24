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
# [BARU] FITUR UPLOAD FILE DI SIDEBAR
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2830/2830305.png", width=50)
st.sidebar.title("Input Data Analisa")

master_file = st.sidebar.file_uploader("1. Upload Master Data Unit (Excel)", type=['xlsx'])
benchmark_file = st.sidebar.file_uploader("2. Upload Laporan Benchmark (Excel)", type=['xlsx'])
trend_file = st.sidebar.file_uploader("3. Upload Laporan Tren Bulanan (Excel - Opsional)", type=['xlsx'])

st.sidebar.markdown("---")

# Hentikan eksekusi jika file utama belum di-upload
if not master_file or not benchmark_file:
    st.info("👈 Silakan upload file Master Data dan Laporan Benchmark di menu sebelah kiri untuk memulai analisa.")
    st.stop()

# ==============================================================================
# 2. LOAD DATA UTAMA (SUPER ROBUST MAPPING - DARI FILE UPLOAD)
# ==============================================================================
@st.cache_data
def load_data(uploaded_master, uploaded_benchmark):
    data_unit = None
    data_inaktif = None
    
    # Dictionary Mapping
    map_loc = {}
    map_cap = {} 
    map_hp = {} 
    
    # Fungsi Pembersih Nama Ekstrem (Hanya Huruf & Angka)
    def clean_key(text):
        if pd.isna(text): return ""
        text = str(text).upper().strip()
        text = text.replace("FORKLIFT", "FORKLIF")
        # Buang semua simbol & spasi, uppercase
        return re.sub(r'[^A-Z0-9]', '', text)

    # 0. Load Master Data dari File Upload
    try:
        df_master = pd.read_excel(uploaded_master, sheet_name='Sheet2', header=1)
        
        # Cari Kolom
        col_name_master = next((c for c in df_master.columns if 'NAMA' in str(c).upper()), None)
        cap_keywords = ['KAPASITAS', 'CAPACITY', 'CAP', 'TON', 'CLASS']
        col_cap_master = next((c for c in df_master.columns if any(k in str(c).upper() for k in cap_keywords)), None)
        
        # Cari Kolom HP
        hp_keywords = ['HP', 'HORSE POWER', 'TENAGA']
        col_hp_master = next((c for c in df_master.columns if any(k == str(c).upper() for k in hp_keywords)), None)
        
        col_loc_master = 'DES 2025' 
        
        if col_name_master and col_loc_master in df_master.columns:
            for _, row in df_master.iterrows():
                u_name = str(row[col_name_master]).strip().upper()
                u_key = clean_key(u_name) 
                
                loc_val = str(row[col_loc_master])
                
                # Parsing Capacity
                cap_val = 0
                if u_key == clean_key("L 9025 US"):
                    cap_val = 40
                else:
                    if col_cap_master:
                        try:
                            raw_cap = str(row[col_cap_master])
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

                # Parsing HP
                hp_val = 0
                if col_hp_master:
                    try:
                        hp_val = pd.to_numeric(row[col_hp_master], errors='coerce')
                        if pd.isna(hp_val): hp_val = 0
                    except: pass
                
                # Simpan ke Map (Key Asli & Key Bersih)
                if u_name not in map_loc: map_loc[u_name] = loc_val
                if u_key not in map_loc:  map_loc[u_key] = loc_val
                
                if cap_val > 0:
                    if u_name not in map_cap: map_cap[u_name] = cap_val
                    if u_key not in map_cap:  map_cap[u_key] = cap_val
                    
                if hp_val > 0:
                    if u_name not in map_hp: map_hp[u_name] = hp_val
                    if u_key not in map_hp:  map_hp[u_key] = hp_val
                    
    except Exception as e:
        pass

    # 1. Load Data Unit Aktif & Inaktif dari File Upload
    try:
        data_unit = pd.read_excel(uploaded_benchmark, sheet_name='Unit_Aktif')
        
        # Rounding Capacity pada Data Unit Aktif
        if 'Capacity' in data_unit.columns:
            data_unit['Capacity'] = pd.to_numeric(data_unit['Capacity'], errors='coerce').fillna(0)
            data_unit['Capacity'] = data_unit['Capacity'].apply(lambda x: int(x + 0.5))
            
        # Pastikan HP Numerik
        if 'Horse_Power' in data_unit.columns:
            data_unit['Horse_Power'] = pd.to_numeric(data_unit['Horse_Power'], errors='coerce').fillna(0)

        data_inaktif = pd.read_excel(uploaded_benchmark, sheet_name='Unit_Inaktif')
    except Exception as e:
        pass
    
    # 2. PERBAIKAN: ISI DATA KOSONG PADA INAKTIF
    if data_inaktif is not None:
        
        # Helper untuk mencari data
        def get_master_data(unit_name, map_dict, default_val):
            clean_name = str(unit_name).strip().upper()
            if clean_name in map_dict: return map_dict[clean_name]
            super_clean = clean_key(unit_name)
            if super_clean in map_dict: return map_dict[super_clean]
            return default_val

        def fix_inaktif_row(row):
            # Fix Lokasi
            curr_loc = str(row.get('Lokasi', '-'))
            if curr_loc in ['-', 'nan', 'None', '', '0', 'NaT']:
                row['Lokasi'] = get_master_data(row['Unit_Name'], map_loc, "-")
            
            # Fix Capacity
            curr_cap = pd.to_numeric(row.get('Capacity', 0), errors='coerce')
            if pd.isna(curr_cap) or curr_cap == 0:
                row['Capacity'] = get_master_data(row['Unit_Name'], map_cap, 0)
            else:
                row['Capacity'] = int(curr_cap + 0.5)
                
            # Fix HP
            curr_hp = pd.to_numeric(row.get('Horse_Power', 0), errors='coerce')
            if pd.isna(curr_hp) or curr_hp == 0:
                row['Horse_Power'] = get_master_data(row['Unit_Name'], map_hp, 0)
                
            return row

        if 'Lokasi' not in data_inaktif.columns: data_inaktif['Lokasi'] = "-"
        if 'Capacity' not in data_inaktif.columns: data_inaktif['Capacity'] = 0
        if 'Horse_Power' not in data_inaktif.columns: data_inaktif['Horse_Power'] = 0
        
        data_inaktif = data_inaktif.apply(fix_inaktif_row, axis=1)

    return data_unit, data_inaktif

df_unit, df_inaktif = load_data(master_file, benchmark_file)

# ==============================================================================
# FUNGSI LOAD DATA TREN (DARI FILE UPLOAD)
# ==============================================================================
@st.cache_data
def load_monthly_data(unit_name_target, uploaded_trend):
    if uploaded_trend is None:
        return pd.DataFrame()
        
    try:
        df_trend = pd.read_excel(uploaded_trend)
        
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
            if "FL RENTAL 01" in raw_name and "TIMIKA" not in raw_name:
                if target_clean == clean_key_trend("FL RENTAL 01 TIMIKA"): return True
            elif "TOBATI" in raw_name and "KALMAR 32T" in raw_name:
                if target_clean == clean_key_trend("TOP LOADER KALMAR 35T/TOBATI"): return True
            elif "L 8477 UUC" in raw_name:
                if target_clean == clean_key_trend("L 9902 UR / S75"): return True
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
                    if clean_after and clean_after in target_clean: return True
                except: pass
                
            return False

        mask = df_trend['Unit'].apply(is_match)
        row_data = df_trend[mask]
        
        if not row_data.empty:
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

# ==============================================================================
# 4. KONTEN UTAMA: ANALISA BERDASARKAN HP
# ==============================================================================
if df_unit is None:
    st.error("⚠️ File Benchmark belum di-upload atau format tidak sesuai.")
    st.stop()

# --- FUNGSI PENAMBAHAN SATUAN CAPACITY (TON / FEET) ---
def format_capacity_with_unit(row):
    cap = row.get('Capacity', 0)
    if pd.isna(cap) or cap == 0 or cap == "-":
        return str(cap)
        
    jenis = str(row.get('Jenis_Alat', '')).upper()
    try:
        cap_val = int(float(cap))
        # Logika pembagian satuan berdasarkan jenis alat
        if jenis in ['CRANE', 'FORKLIFT', 'REACH STACKER', 'SIDE LOADER', 'TOP LOADER']:
            return f"{cap_val} Ton"
        elif jenis in ['TRONTON', 'TRAILER']:
            return f"{cap_val} Feet"
        return str(cap_val)
    except:
        return str(cap)

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
        
        # Urutkan dari Fuel Ratio Terboros ke Terefisien
        if 'Fuel_Ratio' in res_all.columns:
            res_all.sort_values(by='Fuel_Ratio', ascending=False, inplace=True)
        
        # Kolom disamakan dengan Detail Unit Aktif + kolom Status
        cols_to_show = ['Unit_Name', 'Jenis_Alat', 'Status', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Group_Benchmark_Median', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']
        
        # Pastikan kolom ada untuk data inaktif
        for c in cols_to_show:
            if c not in res_all.columns:
                res_all[c] = 0 if c in ['Total_Liter', 'Total_HM_Work', 'Fuel_Ratio', 'Potensi_Pemborosan_Liter', 'Group_Benchmark_Median'] else "-"
        
        df_search_display = res_all[cols_to_show].copy()
        
        # Terapkan Satuan Ton/Feet pada kolom Capacity
        df_search_display['Capacity'] = df_search_display.apply(format_capacity_with_unit, axis=1)
        
        # Rename kolom persis seperti di Detail Unit Aktif
        rename_map_search = {
            'Unit_Name': 'Unit',
            'Total_Liter': 'Total_Pengisian_BBM',
            'Total_HM_Work': 'Total_Jam_Kerja',
            'Group_Benchmark_Median': 'Benchmark',
            'Performance_Status': 'Status_BBM',
            'Potensi_Pemborosan_Liter': 'Potensi_Pemborosan_BBM'
        }
        df_search_display.rename(columns=rename_map_search, inplace=True)

        # Fungsi highlight warna Merah/Hijau
        def highlight_search(row):
            status_bbm = str(row['Status_BBM']).upper()
            color = '#2ca02c' if status_bbm == 'EFISIEN' else ('#d62728' if status_bbm == 'BOROS' else '')
            return [f'background-color: {color}; color: white' if col == 'Fuel_Ratio' else '' for col in row.index]

        # Tampilkan tabel dengan format yang persis sama, 'Capacity' dihilangkan dari style.format karena sudah berbentuk string
        st.dataframe(
            df_search_display.style.format({
                'Horse_Power': '{:.0f}',
                'Total_Pengisian_BBM': '{:,.0f}', 
                'Total_Jam_Kerja': '{:,.0f}', 
                'Fuel_Ratio': '{:.2f}', 
                'Benchmark': '{:.2f}', 
                'Potensi_Pemborosan_BBM': '{:,.0f}'
            })
            .apply(highlight_search, axis=1)
        )
    else:
        st.warning("Unit Tidak Ditemukan.")

st.markdown("---")

# --- Sidebar Filters ---
st.sidebar.subheader("Filter Benchmark")

# 1. Pilih HP 
hp_active = df_unit['Horse_Power'].dropna().unique().tolist()
hp_inactive = df_inaktif['Horse_Power'].dropna().unique().tolist() if df_inaktif is not None and not df_inaktif.empty else []
all_hp = sorted(list(set(hp_active + hp_inactive)))

if not all_hp:
    st.warning("Tidak ada data unit (Aktif/Inaktif) dengan Horse Power yang tersedia.")
    st.stop()

# Format nama HP di dropdown
hp_options = [f"{int(hp)} HP" for hp in all_hp if hp > 0]
if not hp_options:
    st.warning("Data HP bernilai > 0 kosong.")
    st.stop()

selected_hp_str = st.sidebar.selectbox("1. Pilih Kategori HP (Horse Power):", hp_options)
selected_hp_val = float(selected_hp_str.replace(" HP", ""))

# 2. Input Harga Solar
st.sidebar.markdown("---")
st.sidebar.subheader("Biaya Bahan Bakar")
harga_solar = st.sidebar.number_input("Harga Solar (IDR):", value=6800, step=100, key='solar_alat')

# --- FILTER FINAL BERDASARKAN HP SAJA ---
df_active = df_unit[df_unit['Horse_Power'] == selected_hp_val].copy()
df_inactive_show = df_inaktif[df_inaktif['Horse_Power'] == selected_hp_val].copy() if df_inaktif is not None and not df_inaktif.empty else pd.DataFrame()

# --- MAIN CONTENT ---
st.subheader(f"Analisa Kategori: {selected_hp_str}")

if not df_inactive_show.empty:
    with st.expander(f"⚠️ {len(df_inactive_show)} Unit Tidak Masuk Analisa"):
        # Terapkan Satuan Ton/Feet untuk tabel unit inaktif
        df_inactive_display = df_inactive_show[['Unit_Name', 'Jenis_Alat', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work']].copy()
        df_inactive_display['Capacity'] = df_inactive_display.apply(format_capacity_with_unit, axis=1)
        
        st.dataframe(df_inactive_display.rename(columns={'Total_Liter': 'Total_Pengisian_BBM', 'Total_HM_Work': 'Total_Jam_Kerja', 'Unit_Name': 'Unit'}))
        
if df_active.empty:
    st.warning(f"Tidak ada unit aktif untuk kategori {selected_hp_str}.")
    st.stop()

# --- KPI CALCULATIONS ---
benchmark_val = df_active['Group_Benchmark_Median'].iloc[0] if 'Group_Benchmark_Median' in df_active.columns else 0
total_waste = df_active['Potensi_Pemborosan_Liter'].sum()
total_loss_rp = total_waste * harga_solar

# Sort by Fuel_Ratio Ascending (Efisien -> Boros)
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

# Tab A: Data Detail
with tab_a:
    st.subheader("Detail Unit Aktif")
    st.info(f"**Total Pemborosan**: **{total_waste:,.0f} Liter** setara dengan **Rp {total_loss_rp:,.0f}**")
    
    # Menampilkan Jenis Alat, Horse_Power dan Capacity
    df_display_active = df_active[['Unit_Name', 'Jenis_Alat', 'Horse_Power', 'Capacity', 'Lokasi', 'Total_Liter', 'Total_HM_Work', 'Fuel_Ratio', 'Performance_Status', 'Potensi_Pemborosan_Liter']].copy()
    
    # Urutkan dari Fuel Ratio Terboros ke Terefisien
    df_display_active.sort_values(by='Fuel_Ratio', ascending=False, inplace=True)
    
    # Terapkan Satuan Ton/Feet untuk tabel Unit Aktif
    df_display_active['Capacity'] = df_display_active.apply(format_capacity_with_unit, axis=1)
    
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

    # 'Capacity' dihilangkan dari style.format karena datanya sudah berbentuk string
    st.dataframe(
        df_display_active.style.format({
            'Horse_Power': '{:.0f}',
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
        
        # [BARU] Pass trend_file yang di-upload ke fungsi ini
        df_trend = load_monthly_data(selected_unit_active, trend_file)
        
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
            st.warning("Data tren bulanan tidak tersedia atau file Tren belum di-upload.")

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

# Tab D: Analisa Pemborosan
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
        st.success("Tidak ada unit yang terindikasi boros dalam kategori ini.")