import sqlite3
import pandas as pd

print("--- DIAGNOSA SIBAPER ---")

# 1. Cek isi Database SQLite
try:
    conn = sqlite3.connect('sibaper.db')
    c = conn.cursor()
    
    c.execute("SELECT COUNT(*) FROM barang")
    jumlah_barang = c.fetchone()[0]
    print(f"1. Jumlah barang di database: {jumlah_barang}")
    
    c.execute("SELECT * FROM settings")
    pengaturan = c.fetchall()
    print(f"2. Isi tabel settings di database: {pengaturan}")
except Exception as e:
    print(f"Error Database: {e}")

print("------------------------")

# 2. Cek file Excel
try:
    df = pd.read_excel('Persediaan barang.xlsx').dropna(subset=['Jml Barang'])
    bbm_rows = df[df['Uraian (Nama Barang)'].str.contains("Bahan Bakar|Pelumas", case=False, na=False)]
    print(f"3. Total BBM terbaca dari Excel: {bbm_rows['Jml Barang'].sum()} Liter")
except Exception as e:
    print(f"Error Excel: {e}")