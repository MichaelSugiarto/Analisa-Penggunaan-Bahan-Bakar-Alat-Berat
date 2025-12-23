import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ==========================================
# 1. KONFIGURASI FILE & FUNGSI LOAD DATA
# ==========================================

# Cukup tulis nama file Excel utamanya saja
excel_filename = 'BBM AAB.xlsx' 

def preprocess_excel_data(filename):
    print(f"Sedang membaca file Excel: {filename}...")
    print("Ini mungkin memakan waktu beberapa detik...")
    
    try:
        # sheet_name=None artinya membaca SEMUA sheet sekaligus menjadi Dictionary
        # header=None agar kita bisa menangkap baris ke-0 dan ke-2 manual seperti sebelumnya
        all_sheets = pd.read_excel(filename, sheet_name=None, header=None)
    except Exception as e:
        print(f"Error membaca file Excel: {e}")
        return pd.DataFrame()

    all_data = []
    
    # Loop untuk setiap Sheet yang ditemukan (Jan, Feb, Mar, dst)
    for sheet_name, df in all_sheets.items():
        print(f"Memproses Sheet: {sheet_name}")
        
        try:
            # === LOGIKA DATA CLEANING (SAMA SEPERTI SEBELUMNYA) ===
            
            # Baris 0: Nama Unit (Merged Cells -> perlu ffill)
            unit_names = df.iloc[0].ffill()
            
            # Baris 2: Metrik (HM, LITER, KELUAR, dll)
            metrics = df.iloc[2]
            
            # Baris 3 ke bawah: Data Transaksi
            dates = df.iloc[3:, 0]
            
            # Loop setiap kolom mulai index 1
            for col in range(1, df.shape[1]):
                unit = unit_names[col]
                metric = metrics[col]
                
                if pd.notna(unit) and pd.notna(metric):
                    metric = str(metric).strip().upper()
                    
                    if metric in ['HM', 'KM', 'LITER', 'KELUAR']:
                        values = df.iloc[3:, col]
                        
                        temp_df = pd.DataFrame({
                            'Date': dates,
                            'Month': sheet_name, # Nama sheet (misal 'JAN') jadi nama bulan
                            'Unit': unit,
                            'Metric': metric,
                            'Value': values
                        })
                        all_data.append(temp_df)
                        
        except Exception as e:
            print(f"Gagal memproses sheet {sheet_name}: {e}")
            
    if not all_data:
        return pd.DataFrame()
    
    big_df = pd.concat(all_data, ignore_index=True)
    big_df['Value'] = pd.to_numeric(big_df['Value'], errors='coerce')
    
    return big_df

# ==========================================
# 2. EKSEKUSI (Gunakan Fungsi Baru Ini)
# ==========================================
df_long = preprocess_excel_data(excel_filename)

# Cek hasil data
print("\nData berhasil digabungkan:")
print(df_long.head())

# ==========================================
# 2. TRANSFORMASI DATA (PIVOT TABLE)
# ==========================================
# Ubah jadi format: 1 Baris = 1 Hari per Unit, dengan kolom HM, LITER, KELUAR
df_tidy = df_long.pivot_table(
    index=['Date', 'Month', 'Unit'], 
    columns='Metric', 
    values='Value', 
    aggfunc='sum'
).reset_index()

# Isi NaN dengan 0 agar bisa dihitung
for col in ['HM', 'LITER', 'KELUAR']:
    if col in df_tidy.columns:
        df_tidy[col] = df_tidy[col].fillna(0)

# Kategorisasi Jenis Unit (Otomatis berdasarkan nama)
def get_category(name):
    name = str(name).upper()
    if any(x in name for x in ['TANGKI', 'SPBU', 'BUNKER', 'GENSET']):
        return 'Storage'
    elif any(x in name for x in ['KALMAR', 'LINDE', 'CRANE', 'SMV', 'KONECRANES']):
        return 'Alat Berat' # Menggunakan HM
    elif name.startswith('L ') or name.startswith('B '): 
        return 'Truk' # Biasanya pakai KM (jika ada datanya)
    else:
        return 'Lainnya'

df_tidy['Category'] = df_tidy['Unit'].apply(get_category)

# ==========================================
# 3. ANALISA 1: EFISIENSI (LITER PER HOUR)
# ==========================================
# Pastikan urutan data benar berdasarkan Tanggal per Unit
df_tidy['Date'] = pd.to_datetime(df_tidy['Date'], format='%d-%m-%Y', errors='coerce')
df_tidy = df_tidy.sort_values(by=['Unit', 'Date'])

# Hitung Delta HM (Jam Kerja Harian) = HM Hari Ini - HM Kemarin
df_tidy['Prev_HM'] = df_tidy.groupby('Unit')['HM'].shift(1)
df_tidy['Delta_HM'] = df_tidy['HM'] - df_tidy['Prev_HM']

# Validasi Delta HM (Hapus nilai negatif/reset atau nilai tidak wajar > 24 jam)
df_tidy.loc[(df_tidy['Delta_HM'] < 0) | (df_tidy['Delta_HM'] > 24), 'Delta_HM'] = np.nan

# Hitung LPH (Liter Per Hour)
# Hindari pembagian dengan 0
df_tidy['LPH'] = np.where(df_tidy['Delta_HM'] > 0, df_tidy['LITER'] / df_tidy['Delta_HM'], 0)

# Filter data Alat Berat untuk analisa efisiensi
alat_berat_df = df_tidy[df_tidy['Category'] == 'Alat Berat'].copy()

# ==========================================
# 4. VISUALISASI HASIL ANALISA
# ==========================================

# A. Grafik Distribusi Efisiensi (Boxplot) - Mencari Alat Boros
plt.figure(figsize=(12, 6))
# Ambil Top 15 unit paling aktif
top_units = alat_berat_df.groupby('Unit')['LITER'].sum().nlargest(15).index
sns.boxplot(data=alat_berat_df[alat_berat_df['Unit'].isin(top_units)], x='Unit', y='LPH')
plt.title('Distribusi Konsumsi BBM (Liter/Jam) - Top 15 Alat Berat')
plt.xlabel('Nama Unit')
plt.ylabel('Liter / Jam (LPH)')
plt.xticks(rotation=45, ha='right')
plt.ylim(0, 50) # Batasi axis Y agar outlier ekstrim tidak merusak grafik
plt.tight_layout()
plt.show()

# B. Grafik Scatter (Produktivitas vs Konsumsi)
plt.figure(figsize=(10, 6))
# Filter data valid (HM > 0 dan Liter > 0)
scatter_data = alat_berat_df[(alat_berat_df['Delta_HM'] > 0) & (alat_berat_df['LITER'] > 0)]
sns.scatterplot(data=scatter_data, x='Delta_HM', y='LITER', hue='Unit', alpha=0.6, legend=False)
plt.title('Produktivitas (Jam Kerja) vs Konsumsi BBM (Liter)')
plt.xlabel('Jam Kerja Harian (HM)')
plt.ylabel('Konsumsi BBM Harian (Liter)')
plt.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show()

# ==========================================
# 5. ANALISA 2: REKONSILIASI STOK (STORAGE VS UNIT)
# ==========================================
# Membandingkan Total 'KELUAR' dari Tangki vs Total 'LITER' masuk ke Unit
# Jika Total Keluar > Total Masuk = Shrinkage/Loss wajar
# Jika Total Keluar < Total Masuk = Anomali (Ada BBM "siluman" atau lupa catat pengeluaran)

monthly_recon = df_tidy.groupby(['Month', 'Category'])[['KELUAR', 'LITER']].sum().reset_index()

# Pivot agar mudah dibandingkan
total_out_storage = df_tidy[df_tidy['Category'] == 'Storage'].groupby('Month')['KELUAR'].sum()
total_in_equipment = df_tidy[df_tidy['Category'] != 'Storage'].groupby('Month')['LITER'].sum()

recon_df = pd.DataFrame({
    'Total Keluar Gudang (Liter)': total_out_storage, 
    'Total Masuk Unit (Liter)': total_in_equipment
})

# Urutkan bulan
bulan_order = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGT', 'SEP', 'OKT', 'NOV']
recon_df = recon_df.reindex(bulan_order)

# Plot Bar Chart Rekonsiliasi
recon_df.plot(kind='bar', figsize=(12, 6))
plt.title('Rekonsiliasi BBM Bulanan: Keluar Gudang vs Masuk Unit')
plt.ylabel('Total Liter')
plt.xticks(rotation=45)
plt.grid(axis='y', linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show()

# Tampilkan Tabel Data Rekonsiliasi
print("\n=== DATA REKONSILIASI BULANAN ===")
print(recon_df)

# Tampilkan Top 5 Unit Paling Boros (Rata-rata LPH Tertinggi)
print("\n=== TOP 5 ALAT BERAT PALING BOROS (AVG LPH) ===")
print(alat_berat_df.groupby('Unit')['LPH'].mean().sort_values(ascending=False).head(5))