import streamlit as st
import pandas as pd
import datetime
import io
import uuid
import sqlite3
import os
import plotly.express as px

# --- ⚙️ KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SiBAPER Dashboard", layout="wide", initial_sidebar_state="expanded")

# Atasi FutureWarning Pandas
pd.set_option('future.no_silent_downcasting', True)

# ==========================================
# 0. DATABASE INITIALIZATION (SQLITE)
# ==========================================
@st.cache_resource
def init_connection():
    return sqlite3.connect('sibaper.db', check_same_thread=False)

conn = init_connection()

def init_db():
    c = conn.cursor()
    # Tabel Utama
    c.execute('''CREATE TABLE IF NOT EXISTS barang (
                    kode TEXT, nama TEXT PRIMARY KEY, kategori TEXT, 
                    jumlah REAL, satuan TEXT, last_updated TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS req_barang (
                    id TEXT PRIMARY KEY, tanggal TEXT, unit TEXT, barang TEXT, 
                    jumlah REAL, satuan TEXT, keterangan TEXT, status TEXT, alasan_tolak TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS req_bbm (
                    id TEXT PRIMARY KEY, tanggal TEXT, unit TEXT, kegiatan TEXT, 
                    kendaraan TEXT, liter REAL, jumlah_pembelian REAL, 
                    status TEXT, link_struk TEXT, status_laporan TEXT, alasan_tolak TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS laporan_bbm (
                    no_laporan TEXT PRIMARY KEY, no_st TEXT, tanggal_struk TEXT, 
                    jenis_bbm TEXT, jumlah_liter REAL, total_rp REAL, 
                    link_struk TEXT, status_laporan TEXT, alasan_tolak TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS log_restok (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, tanggal TEXT, 
                    kategori TEXT, item TEXT, jumlah REAL, link_bukti TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
    conn.commit()

init_db()

# ==========================================
# 1. LOAD DATA DARI EXCEL & DUMMY DATA
# ==========================================
def generate_dummy_data(c):
    today = datetime.date.today()
    
    # 1. Dummy Data Permohonan Barang
    c.execute("INSERT INTO req_barang VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              (f"PRM-{str(uuid.uuid4())[:4].upper()}", (today - datetime.timedelta(days=5)).strftime("%Y-%m-%d"), "Irban Wilayah I", "Bahan Cetak", 5, "Rim", "Pengawasan Reguler", "Disetujui", "-"))
    c.execute("INSERT INTO req_barang VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              (f"PRM-{str(uuid.uuid4())[:4].upper()}", (today - datetime.timedelta(days=1)).strftime("%Y-%m-%d"), "Irban Wilayah II", "Kertas dan Cover", 10, "Rim", "Stok ATK Ruangan", "Menunggu", "-"))
    c.execute("INSERT INTO req_barang VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              (f"PRM-{str(uuid.uuid4())[:4].upper()}", (today - datetime.timedelta(days=2)).strftime("%Y-%m-%d"), "Subag UK", "Benda Pos", 5, "Pcs", "Arsip Dokumen", "Ditolak", "Stok Habis"))
    c.execute("INSERT INTO req_barang VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              (f"PRM-{str(uuid.uuid4())[:4].upper()}", today.strftime("%Y-%m-%d"), "Irban Wilayah III", "Alat Listrik", 2, "Pcs", "aaaa", "Menunggu", "-"))

    # 2. Dummy Data Permohonan & Pelaporan BBM
    # Case A: Selesai (Sudah lapor struk dan divalidasi)
    c.execute("INSERT INTO req_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("ST/001/2026", (today - datetime.timedelta(days=10)).strftime("%Y-%m-%d"), "Irban Wilayah I", "Monev Desa A", "DK 1111 A", 10, 118000, "Disetujui", "https://gdrive.com/struk1", "Selesai", "-"))
    c.execute("INSERT INTO laporan_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("LAP-A1B2", "ST/001/2026", (today - datetime.timedelta(days=9)).strftime("%Y-%m-%d"), "Pertamax", 10, 118000, "https://gdrive.com/struk1", "Selesai", "-"))

    # Case B: Menunggu Verifikasi Struk (Admin Bagian sudah lapor, nunggu Superadmin klik)
    c.execute("INSERT INTO req_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("ST/002/2026", (today - datetime.timedelta(days=4)).strftime("%Y-%m-%d"), "Irban Wilayah II", "Rakor Provinsi", "DK 2222 B", 15, 177000, "Disetujui", "https://gdrive.com/struk2", "Menunggu Verifikasi", "-"))
    c.execute("INSERT INTO laporan_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("LAP-C3D4", "ST/002/2026", (today - datetime.timedelta(days=2)).strftime("%Y-%m-%d"), "Pertamax", 15, 177000, "https://gdrive.com/struk2", "Menunggu Verifikasi", "-"))

    # Case C: Disetujui ST, Tapi Belum Lapor Struk (Blokir permohonan baru untuk unit ini)
    c.execute("INSERT INTO req_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("ST/003/2026", (today - datetime.timedelta(days=2)).strftime("%Y-%m-%d"), "Subag UK", "Distribusi Surat", "DK 3333 C", 8, 94400, "Disetujui", "", "Belum Dilaporkan", "-"))

    # Case D: Menunggu Persetujuan ST Baru
    c.execute("INSERT INTO req_bbm VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
              ("ST/004/2026", today.strftime("%Y-%m-%d"), "Irban Wilayah IV", "Kunjungan Lapangan", "DK 4444 D", 20, 236000, "Menunggu", "", "Belum Dilaporkan", "-"))
    
    # 3. Dummy Data Log Restok
    c.execute("INSERT INTO log_restok (tanggal, kategori, item, jumlah, link_bukti) VALUES (?, ?, ?, ?, ?)",
              ((today - datetime.timedelta(days=7)).strftime("%Y-%m-%d"), "Restok BBM", "BBM", 500, "https://gdrive.com/do1"))
    
def load_initial_data():
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM barang")
    if c.fetchone()[0] == 0:
        excel_path = 'data/Persediaan barang.xlsx'
        if os.path.exists(excel_path):
            try:
                df_excel = pd.read_excel(excel_path).dropna(subset=['Jml Barang'])
                bbm_rows = df_excel[df_excel['Uraian (Nama Barang)'].str.contains("Bahan Bakar|Pelumas", case=False, na=False)]
                total_bbm_awal = float(bbm_rows['Jml Barang'].sum()) if not bbm_rows.empty else 1000.0
                
                c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('total_liter_bbm', ?)", (str(total_bbm_awal),))
                
                df_excel = df_excel[~df_excel['Uraian (Nama Barang)'].str.contains("Bahan Bakar|Pelumas", case=False, na=False)]
                today = datetime.date.today().strftime("%Y-%m-%d")
                for _, row in df_excel.iterrows():
                    c.execute("INSERT OR IGNORE INTO barang (kode, nama, kategori, jumlah, satuan, last_updated) VALUES (?, ?, ?, ?, ?, ?)",
                              (str(row['Kode']), str(row['Uraian (Nama Barang)']), 'Umum', float(row['Jml Barang']), 'Pcs', today))
                
                # Masukkan Dummy Data
                generate_dummy_data(c)
                conn.commit()
            except Exception as e:
                st.error(f"Gagal membaca Excel: {e}")
                c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('total_liter_bbm', '1000.0')")
                generate_dummy_data(c)
                conn.commit()
        else:
            st.warning("File 'Persediaan barang.xlsx' tidak ditemukan di folder 'data/'. Menggunakan saldo default 1000 L.")
            c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('total_liter_bbm', '1000.0')")
            # Masukkan Dummy Data karena database kosong
            generate_dummy_data(c)
            conn.commit()

load_initial_data()

# Fungsi Reset Database Manual
def reset_database():
    c = conn.cursor()
    c.execute("DELETE FROM barang")
    c.execute("DELETE FROM req_barang")
    c.execute("DELETE FROM req_bbm")
    c.execute("DELETE FROM laporan_bbm")
    c.execute("DELETE FROM log_restok")
    c.execute("DELETE FROM settings")
    conn.commit()
    load_initial_data()
    st.success("✅ Database berhasil di-reset, disinkronkan dengan Excel, dan Dummy Data telah dimasukkan!")
    st.rerun()

# ==========================================
# 2. SISTEM AUTENTIKASI
# ==========================================
USERS = {
    "superadmin": ["admin123", "superadmin", "Super Admin"],
    "irban1": ["user123", "admin_bagian", "Irban Wilayah I"],
    "irban2": ["user123", "admin_bagian", "Irban Wilayah II"],
    "irban3": ["user123", "admin_bagian", "Irban Wilayah III"],
    "irban4": ["user123", "admin_bagian", "Irban Wilayah IV"],
    "irban5": ["user123", "admin_bagian", "Irban Wilayah V"],
    "subag_uk": ["user123", "admin_bagian", "Subag UK"]
}

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username, st.session_state.role, st.session_state.unit_name = "", "", ""

def login():
    st.title("🔐 Login SiBAPER")
    with st.form("login_form"):
        st.caption("🔴 *Wajib Diisi*")
        username = st.text_input("Username *")
        password = st.text_input("Password *", type="password")
        if st.form_submit_button("Masuk", type="primary"):
            if username in USERS and USERS[username][0] == password:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = USERS[username][1]
                st.session_state.unit_name = USERS[username][2]
                st.rerun()
            else:
                st.error("Username atau Password salah!")

def logout():
    st.session_state.logged_in = False
    st.rerun()

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2: login()
    st.stop()


# --- 🧭 NAVIGASI SIDEBAR ---
st.sidebar.title("🏢 SiBAPER")
st.sidebar.markdown(f"**Halo, {st.session_state.unit_name}**")
st.sidebar.markdown("---")

if st.session_state.role == "superadmin":
    menu = st.sidebar.radio("Navigasi Super Admin", ["Beranda", "Restok Barang", "Tambah Jenis Barang", "Restok BBM", "Verifikasi Barang", "Verifikasi BBM", "Laporan"])
else:
    menu = st.sidebar.radio("Navigasi Admin Bagian", ["Dashboard Unit", "Permohonan Barang", "Permohonan BBM", "Pelaporan BBM", "Riwayat Unit"])

st.sidebar.markdown("---")
if st.sidebar.button("Keluar (Logout)"): logout()

HARGA_BBM = 11800


# ==========================================
# MENU ADMIN BAGIAN
# ==========================================
if st.session_state.role == "admin_bagian":
    my_unit = st.session_state.unit_name
    
    if menu == "Dashboard Unit":
        st.title(f"Dashboard {my_unit}")
        df_req_brg = pd.read_sql_query("SELECT * FROM req_barang WHERE unit = ?", conn, params=(my_unit,))
        df_req_bbm = pd.read_sql_query("SELECT * FROM req_bbm WHERE unit = ?", conn, params=(my_unit,))
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Permohonan Barang", len(df_req_brg))
        c2.metric("Total Surat Tugas BBM", len(df_req_bbm))
        c3.metric("BBM Menunggu Pelaporan", len(df_req_bbm[(df_req_bbm['status'] == 'Disetujui') & (df_req_bbm['link_struk'] == '')]))
        
        st.markdown("---")
        st.subheader("📦 Status Terbaru Barang")
        st.dataframe(df_req_brg.tail(5), use_container_width=True, hide_index=True)
        st.subheader("⛽ Status Terbaru BBM")
        st.dataframe(df_req_bbm.tail(5), use_container_width=True, hide_index=True)
        
    elif menu == "Permohonan Barang":
        st.title("Permohonan Barang")
        with st.form("form_ajukan_barang"):
            st.caption("🔴 *Wajib Diisi*")
            col1, col2 = st.columns([1, 1])
            with col1:
                tanggal = st.date_input("Tanggal Permohonan *")
                kegiatan = st.text_input("Kegiatan / Keperluan *")
            with col2:
                df_barang = pd.read_sql_query("SELECT nama, satuan FROM barang", conn)
                barang_diminta = st.selectbox("Pilih Barang *", ["-- Pilih Barang --"] + df_barang['nama'].tolist())
                jumlah_diminta = st.number_input("Jumlah Diminta *", min_value=1.0, step=1.0)
                
            if st.form_submit_button("Kirim Permohonan", type="primary"):
                if not kegiatan or barang_diminta == "-- Pilih Barang --":
                    st.error("Kegiatan dan Barang wajib diisi!")
                else:
                    satuan_barang = df_barang[df_barang['nama'] == barang_diminta]['satuan'].values[0]
                    req_id = f"PRM-{str(uuid.uuid4())[:4].upper()}"
                    c = conn.cursor()
                    c.execute("INSERT INTO req_barang (id, tanggal, unit, barang, jumlah, satuan, keterangan, status, alasan_tolak) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                              (req_id, tanggal.strftime("%Y-%m-%d"), my_unit, barang_diminta, jumlah_diminta, satuan_barang, kegiatan, 'Menunggu', '-'))
                    conn.commit()
                    st.success("✅ Permohonan barang berhasil dikirim!")

    elif menu == "Permohonan BBM":
        st.title("Permohonan BBM")
        active_bbm = pd.read_sql_query("SELECT * FROM req_bbm WHERE unit = ? AND status_laporan != 'Selesai' AND status != 'Ditolak'", conn, params=(my_unit,))
        
        if not active_bbm.empty:
            st.error("⚠️ Selesaikan permohonan BBM sebelumnya (hingga laporan diverifikasi) sebelum mengajukan baru.")
            st.dataframe(active_bbm[['id', 'tanggal', 'kegiatan', 'status', 'status_laporan']], hide_index=True)
        else:
            with st.form("form_ajukan_bbm"):
                st.caption("🔴 *Wajib Diisi*")
                no_st = st.text_input("Nomor Surat Tugas *")
                kegiatan = st.text_area("Tujuan / Kegiatan *")
                kendaraan = st.text_input("Plat Kendaraan *")
                tanggal = st.date_input("Tanggal Kegiatan *")
                liter_diminta = st.number_input("Jumlah Liter *", min_value=1.0, step=1.0)
                
                st.text_input(f"Total Biaya (Rp {HARGA_BBM}/L)", value=f"Rp {(liter_diminta * HARGA_BBM):,.0f}".replace(",", "."), disabled=True)
                
                if st.form_submit_button("Kirim Permohonan", type="primary"):
                    if not no_st or not kegiatan or not kendaraan:
                        st.error("Semua field bertanda * wajib diisi!")
                    else:
                        c = conn.cursor()
                        c.execute("INSERT INTO req_bbm (id, tanggal, unit, kegiatan, kendaraan, liter, jumlah_pembelian, status, link_struk, status_laporan, alasan_tolak) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                  (no_st, tanggal.strftime("%Y-%m-%d"), my_unit, kegiatan, kendaraan, liter_diminta, liter_diminta * HARGA_BBM, 'Menunggu', '', 'Belum Dilaporkan', '-'))
                        conn.commit()
                        st.success("✅ Permohonan BBM berhasil diajukan!")
                        st.rerun()
                        
    elif menu == "Pelaporan BBM":
        st.title("Pelaporan BBM (Upload Struk)")
        df_pending = pd.read_sql_query("SELECT * FROM req_bbm WHERE unit = ? AND status = 'Disetujui' AND (link_struk = '' OR status_laporan = 'Ditolak')", conn, params=(my_unit,))
        
        if df_pending.empty:
            st.success("🎉 Tidak ada pelaporan BBM yang tertunda.")
        else:
            with st.form("form_lapor_bbm"):
                st.caption("🔴 *Wajib Diisi*")
                col1, col2 = st.columns(2)
                with col1:
                    pilihan_st = st.selectbox("Nomor Surat Tugas *", df_pending['id'].tolist())
                    tgl_struk = st.date_input("Tanggal Struk BBM *")
                    link_struk = st.text_input("Link Google Drive Struk *")
                with col2:
                    jenis_bbm = st.selectbox("Jenis BBM *", ["Pertamax", "Dexlite", "Pertamina Dex"])
                    
                    default_liter = df_pending[df_pending['id'] == pilihan_st]['liter'].values[0] if pilihan_st else 1.0
                    liter_aktual = st.number_input("Jumlah Liter Aktual *", min_value=1.0, value=float(default_liter))
                    st.text_input("Total Pembelian (Rp)", value=f"Rp {(liter_aktual * HARGA_BBM):,.0f}".replace(",", "."), disabled=True)
                
                if st.form_submit_button("Simpan Laporan", type="primary"):
                    if not link_struk:
                        st.error("Link Google Drive wajib diisi!")
                    else:
                        c = conn.cursor()
                        c.execute("UPDATE req_bbm SET link_struk = ?, status_laporan = 'Menunggu Verifikasi' WHERE id = ?", (link_struk, pilihan_st))
                        no_lap = f"LAP-{str(uuid.uuid4())[:4]}"
                        c.execute("INSERT INTO laporan_bbm (no_laporan, no_st, tanggal_struk, jenis_bbm, jumlah_liter, total_rp, link_struk, status_laporan, alasan_tolak) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                  (no_lap, pilihan_st, tgl_struk.strftime("%Y-%m-%d"), jenis_bbm, liter_aktual, liter_aktual * HARGA_BBM, link_struk, 'Menunggu Verifikasi', '-'))
                        conn.commit()
                        st.success("✅ Struk berhasil dilaporkan!")
                        st.rerun()

    elif menu == "Riwayat Unit":
        st.title(f"Riwayat Permohonan {my_unit}")
        tab1, tab2 = st.tabs(["Riwayat Barang", "Riwayat BBM"])
        
        with tab1:
            df_brg = pd.read_sql_query("SELECT * FROM req_barang WHERE unit = ?", conn, params=(my_unit,))
            if not df_brg.empty:
                df_brg.insert(0, 'No', range(1, len(df_brg) + 1))
            st.dataframe(df_brg, hide_index=True, use_container_width=True)
            
        with tab2:
            df_bbm = pd.read_sql_query("SELECT * FROM req_bbm WHERE unit = ?", conn, params=(my_unit,))
            if not df_bbm.empty:
                df_bbm = df_bbm.rename(columns={'id': 'No ST'})
                df_bbm.insert(0, 'No', range(1, len(df_bbm) + 1))
            st.dataframe(df_bbm, hide_index=True, use_container_width=True)


# ==========================================
# MENU SUPER ADMIN
# ==========================================
elif st.session_state.role == "superadmin":
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key = 'total_liter_bbm'")
    res = c.fetchone()
    total_liter_bbm = float(res[0]) if res else 1000.0
    
    if menu == "Beranda":
        c1, c2 = st.columns([8, 2])
        c1.title("Beranda (Data Real Stok)")
        if c2.button("🔄 Reset Database", help="Hapus semua data, ulang dari Excel awal, dan generate Dummy Data"):
            reset_database()
            
        df_barang = pd.read_sql_query("SELECT * FROM barang", conn)
        df_req_brg = pd.read_sql_query("SELECT * FROM req_barang", conn)
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Jenis Barang", len(df_barang))
        col2.metric("Sisa BBM (Liter)", f"{total_liter_bbm:,.1f}".replace(",", "."))
        col3.metric("Estimasi Nilai BBM", f"Rp {total_liter_bbm * HARGA_BBM:,.0f}".replace(",", "."))
        col4.metric("Menunggu Verifikasi", len(df_req_brg[df_req_brg['status'] == 'Menunggu']))
        
        st.markdown("---")
        st.subheader("📈 Total Permintaan Per Unit (Barang vs BBM)")
        
        units = ["Irban Wilayah I", "Irban Wilayah II", "Irban Wilayah III", "Irban Wilayah IV", "Irban Wilayah V", "Subag UK"]
        df_units = pd.DataFrame({'Unit': units})
        
        df_chart_brg = pd.read_sql_query("SELECT unit as Unit, COUNT(*) as 'Barang' FROM req_barang GROUP BY unit", conn)
        df_chart_bbm = pd.read_sql_query("SELECT unit as Unit, COUNT(*) as 'BBM' FROM req_bbm GROUP BY unit", conn)
        
        df_chart = df_units.merge(df_chart_brg, on='Unit', how='left').merge(df_chart_bbm, on='Unit', how='left')
        df_chart['Barang'] = df_chart['Barang'].fillna(0).astype(int)
        df_chart['BBM'] = df_chart['BBM'].fillna(0).astype(int)
        
        # Plotly Grouped Bar Chart (Unstacked)
        df_melt = df_chart.melt(id_vars='Unit', value_vars=['Barang', 'BBM'], var_name='Kategori', value_name='Jumlah Permintaan')
        fig = px.bar(df_melt, x='Unit', y='Jumlah Permintaan', color='Kategori', barmode='group',
                     color_discrete_map={'Barang': '#1f77b4', 'BBM': '#ff7f0e'})
        st.plotly_chart(fig, use_container_width=True)

    elif menu == "Restok Barang":
        st.title("Restok Barang")
        with st.form("form_restok_barang"):
            st.caption("🔴 *Wajib Diisi*")
            col1, col2 = st.columns(2)
            df_barang = pd.read_sql_query("SELECT nama FROM barang", conn)
            with col1:
                pilihan_barang = st.selectbox("Pilih Barang *", ["-- Pilih --"] + df_barang['nama'].tolist())
                jumlah_masuk = st.number_input("Jumlah Tambahan *", min_value=1.0, step=1.0)
            with col2:
                link_bukti = st.text_input("Link Google Drive (Bukti) *")
                
            if st.form_submit_button("Simpan Restok", type="primary"):
                if pilihan_barang != "-- Pilih --" and link_bukti:
                    c.execute("UPDATE barang SET jumlah = jumlah + ?, last_updated = ? WHERE nama = ?", 
                              (jumlah_masuk, datetime.date.today().strftime("%Y-%m-%d"), pilihan_barang))
                    c.execute("INSERT INTO log_restok (tanggal, kategori, item, jumlah, link_bukti) VALUES (?, ?, ?, ?, ?)",
                              (datetime.date.today().strftime("%Y-%m-%d"), "Restok Barang", pilihan_barang, jumlah_masuk, link_bukti))
                    conn.commit()
                    st.success(f"✅ Stok {pilihan_barang} berhasil ditambah!")
                else:
                    st.error("Isi semua data restok!")

    elif menu == "Tambah Jenis Barang":
        st.title("Tambah Jenis Barang Baru")
        with st.form("form_tambah_jenis"):
            st.caption("🔴 *Wajib Diisi*")
            col1, col2 = st.columns(2)
            with col1:
                kode = st.text_input("Kode Barang *")
                nama_barang = st.text_input("Nama Barang *")
                kategori = st.selectbox("Kategori *", ["ATK", "Bahan Habis Pakai", "Lainnya"])
            with col2:
                jumlah = st.number_input("Stok Awal *", min_value=0.0, step=1.0)
                satuan = st.selectbox("Satuan *", ["Rim", "Pcs", "Kotak", "Botol", "Roll"])
                
            if st.form_submit_button("Simpan Jenis Baru", type="primary"):
                if kode and nama_barang:
                    c.execute("INSERT INTO barang VALUES (?, ?, ?, ?, ?, ?)", (kode, nama_barang, kategori, jumlah, satuan, datetime.date.today().strftime("%Y-%m-%d")))
                    conn.commit()
                    st.success(f"✅ {nama_barang} ditambahkan!")

    elif menu == "Restok BBM":
        st.title("Restok Saldo BBM (Liter)")
        st.info(f"Sisa Saldo: **{total_liter_bbm:,.1f} Liter**".replace(",", "."))
        
        with st.form("form_restok_bbm"):
            st.caption("🔴 *Wajib Diisi*")
            liter_masuk = st.number_input("Jumlah Tambahan (Liter) *", min_value=1.0, step=10.0)
            link_bukti = st.text_input("Link Bukti GDrive *")
            
            if st.form_submit_button("Tambahkan Saldo", type="primary"):
                if link_bukti:
                    c.execute("UPDATE settings SET value = value + ? WHERE key = 'total_liter_bbm'", (liter_masuk,))
                    c.execute("INSERT INTO log_restok (tanggal, kategori, item, jumlah, link_bukti) VALUES (?, ?, ?, ?, ?)",
                              (datetime.date.today().strftime("%Y-%m-%d"), "Restok BBM", "BBM", liter_masuk, link_bukti))
                    conn.commit()
                    st.success(f"✅ Ditambahkan {liter_masuk} Liter BBM!")
                    st.rerun()

    elif menu == "Verifikasi Barang":
        st.title("Verifikasi Permintaan Barang")
        df_menunggu = pd.read_sql_query("SELECT * FROM req_barang WHERE status = 'Menunggu'", conn)
        
        if df_menunggu.empty:
            st.info("🎉 Tidak ada permintaan yang menunggu verifikasi.")
        else:
            st.dataframe(df_menunggu, use_container_width=True, hide_index=True)
            st.markdown("---")
            
            selected_id = st.selectbox("🔍 Pilih ID Permintaan untuk diproses", df_menunggu['id'].tolist())
            req = df_menunggu[df_menunggu['id'] == selected_id].iloc[0]
            
            st.markdown(f"**Detail Permintaan: {req['unit']}**")
            st.write(f"Barang: **{req['barang']}** | Jumlah: **{req['jumlah']} {req['satuan']}**")
            
            c.execute("SELECT jumlah FROM barang WHERE nama = ?", (req['barang'],))
            res = c.fetchone()
            stok_tersedia = res[0] if res else 0.0
            
            if not res:
                st.warning(f"⚠️ Barang '{req['barang']}' tidak terdaftar di master data. Stok dianggap 0.")
            else:
                st.info(f"Stok Tersedia Saat Ini: **{stok_tersedia} {req['satuan']}**")
            
            can_approve = True
            if req['jumlah'] > stok_tersedia:
                st.error("⚠️ Stok tidak mencukupi untuk memenuhi permintaan ini! Permohonan harus ditolak.")
                can_approve = False
                
            alasan = st.text_input("Alasan Penolakan (Opsional / Wajib jika tolak)", key="alasan_brg")
            
            col1, col2 = st.columns(2)
            if col1.button("✅ Setujui", type="primary", disabled=not can_approve, use_container_width=True):
                c.execute("UPDATE req_barang SET status = 'Disetujui', alasan_tolak = '-' WHERE id = ?", (req['id'],))
                c.execute("UPDATE barang SET jumlah = jumlah - ? WHERE nama = ?", (req['jumlah'], req['barang']))
                conn.commit()
                st.success("Disetujui!")
                st.rerun()
                
            if col2.button("❌ Tolak", use_container_width=True):
                if not alasan: alasan = "Ditolak oleh Admin"
                c.execute("UPDATE req_barang SET status = 'Ditolak', alasan_tolak = ? WHERE id = ?", (alasan, req['id']))
                conn.commit()
                st.warning("Ditolak!")
                st.rerun()

    elif menu == "Verifikasi BBM":
        st.title("Verifikasi BBM")
        tab1, tab2 = st.tabs(["1. Validasi Surat Tugas (Baru)", "2. Validasi Laporan (Struk)"])
        
        with tab1:
            df_st = pd.read_sql_query("SELECT * FROM req_bbm WHERE status = 'Menunggu'", conn)
            if df_st.empty:
                st.info("Tidak ada Surat Tugas baru.")
            else:
                st.dataframe(df_st, use_container_width=True, hide_index=True)
                st.markdown("---")
                
                selected_st = st.selectbox("🔍 Pilih Nomor ST untuk diproses", df_st['id'].tolist())
                req_st = df_st[df_st['id'] == selected_st].iloc[0]
                
                st.write(f"Unit: **{req_st['unit']}** | Diminta: **{req_st['liter']} Liter**")
                
                can_approve = True
                if req_st['liter'] > total_liter_bbm:
                    st.error("⚠️ Saldo BBM Sistem tidak mencukupi!")
                    can_approve = False
                    
                alasan_st = st.text_input("Alasan Penolakan", key="alasan_st")
                
                c1, c2 = st.columns(2)
                if c1.button("✅ Setujui ST", type="primary", disabled=not can_approve, use_container_width=True):
                    c.execute("UPDATE settings SET value = value - ? WHERE key = 'total_liter_bbm'", (req_st['liter'],))
                    c.execute("UPDATE req_bbm SET status = 'Disetujui', alasan_tolak = '-' WHERE id = ?", (req_st['id'],))
                    conn.commit()
                    st.rerun()
                if c2.button("❌ Tolak ST", use_container_width=True):
                    if not alasan_st: alasan_st = "Ditolak"
                    c.execute("UPDATE req_bbm SET status = 'Ditolak', alasan_tolak = ? WHERE id = ?", (alasan_st, req_st['id']))
                    conn.commit()
                    st.rerun()

        with tab2:
            df_lap = pd.read_sql_query("SELECT * FROM laporan_bbm WHERE status_laporan = 'Menunggu Verifikasi'", conn)
            if df_lap.empty:
                st.info("Tidak ada laporan struk baru.")
            else:
                st.dataframe(df_lap, use_container_width=True, hide_index=True)
                st.markdown("---")
                
                selected_lap = st.selectbox("🔍 Pilih Laporan", df_lap['no_laporan'].tolist())
                req_lap = df_lap[df_lap['no_laporan'] == selected_lap].iloc[0]
                
                st.write(f"Nomor ST: **{req_lap['no_st']}** | Liter Aktual: **{req_lap['jumlah_liter']}**")
                st.markdown(f"📄 [Lihat Bukti Struk]({req_lap['link_struk']})")
                alasan_lap = st.text_input("Alasan Penolakan Struk", key="alasan_lap")
                
                c1, c2 = st.columns(2)
                if c1.button("✅ Validasi Sesuai", type="primary", use_container_width=True):
                    c.execute("UPDATE laporan_bbm SET status_laporan = 'Selesai' WHERE no_laporan = ?", (req_lap['no_laporan'],))
                    c.execute("UPDATE req_bbm SET status_laporan = 'Selesai' WHERE id = ?", (req_lap['no_st'],))
                    conn.commit()
                    st.rerun()
                if c2.button("❌ Tolak Struk", use_container_width=True):
                    if not alasan_lap: alasan_lap = "Struk tidak valid"
                    c.execute("UPDATE laporan_bbm SET status_laporan = 'Ditolak', alasan_tolak = ? WHERE no_laporan = ?", (alasan_lap, req_lap['no_laporan']))
                    c.execute("UPDATE req_bbm SET status_laporan = 'Ditolak', link_struk = '' WHERE id = ?", (req_lap['no_st'],))
                    conn.commit()
                    st.rerun()

    elif menu == "Laporan":
        st.title("Laporan & Export")
        tab1, tab2, tab3 = st.tabs(["Stok Barang", "Log Restok", "Penggunaan BBM"])
        
        with tab1:
            df_b = pd.read_sql_query("SELECT * FROM barang", conn)
            st.dataframe(df_b, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(pd.read_sql_query("SELECT * FROM log_restok", conn), use_container_width=True, hide_index=True)
        with tab3:
            df_penggunaan_bbm = pd.read_sql_query("SELECT * FROM laporan_bbm", conn)
            st.dataframe(df_penggunaan_bbm, use_container_width=True, hide_index=True)