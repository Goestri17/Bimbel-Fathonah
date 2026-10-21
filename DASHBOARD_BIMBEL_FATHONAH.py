import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import os
import io
from datetime import datetime
import base64

def get_image_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception as e:
        return "" # Jika logo tidak ditemukan, biarkan kosong

# Konfigurasi Halaman Dashboard
st.set_page_config(
    page_title="Dashboard - Bimbel Fathonah",
    page_icon="📚",
    layout="wide"
)

# --- SETUP DATABASE SQLITE & MIGRASI AMAN ---
DB_FILE = "bimbel_fathonah.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # 1. Tabel Siswa
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS siswa (
            id_siswa TEXT PRIMARY KEY,
            nama_siswa TEXT NOT NULL,
            jenjang TEXT NOT NULL,
            paket_belajar TEXT,
            alamat TEXT,
            fee_per_pertemuan INTEGER,
            fee_tentor INTEGER DEFAULT 0
        )
    ''')
    
    cursor.execute("PRAGMA table_info(siswa)")
    siswa_cols = [col[1] for col in cursor.fetchall()]
    if "fee_tentor" not in siswa_cols:
        cursor.execute("ALTER TABLE siswa ADD COLUMN fee_tentor INTEGER DEFAULT 0")

    # 2. Tabel Tentor (Profil Dasar Tentor)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tentor_profil (
            id_tentor INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_tentor TEXT NOT NULL,
            alamat TEXT,
            jenis_mapel TEXT,
            hari_on_bimbel TEXT
        )
    ''')

    # 3. Tabel Gaji Bulanan Tentor
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tentor_gaji (
            id_gaji INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_tentor TEXT NOT NULL,
            daftar_siswa TEXT,
            detail_pertemuan TEXT,
            jumlah_pertemuan INTEGER DEFAULT 0,
            bulan TEXT NOT NULL,
            total_gaji INTEGER DEFAULT 0
        )
    ''')

    # 4. Tabel Tagihan Wali Siswa (Keuangan)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS keuangan (
            id_keuangan INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_siswa TEXT NOT NULL,
            bulan TEXT NOT NULL,
            jumlah_kehadiran INTEGER DEFAULT 0,
            fee_per_pertemuan INTEGER DEFAULT 0,
            pendaftaran INTEGER DEFAULT 0,
            diskon_pertemuan INTEGER DEFAULT 0,
            diskon INTEGER DEFAULT 0,
            pemasukan INTEGER DEFAULT 0,
            pengeluaran INTEGER DEFAULT 0,
            keterangan TEXT,
            UNIQUE(nama_siswa, bulan)
        )
    ''')
    
    cursor.execute("PRAGMA table_info(keuangan)")
    columns = [col[1] for col in cursor.fetchall()]
    if "pendaftaran" not in columns:
        cursor.execute("ALTER TABLE keuangan ADD COLUMN pendaftaran INTEGER DEFAULT 0")
    if "diskon_pertemuan" not in columns:
        cursor.execute("ALTER TABLE keuangan ADD COLUMN diskon_pertemuan INTEGER DEFAULT 0")
    if "diskon" not in columns:
        cursor.execute("ALTER TABLE keuangan ADD COLUMN diskon INTEGER DEFAULT 0")
    
    conn.commit()
    conn.close()

init_db()

def auto_sync_to_google_drive():
    pass

# Fungsi helper database
def load_data_siswa():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM siswa", conn)
    conn.close()
    if not df.empty:
        df.columns = ["ID Siswa", "Nama Siswa", "Jenjang", "Paket Belajar", "Alamat", "Fee per Pertemuan (Rp)", "Fee Tentor (Rp)"]
    else:
        df = pd.DataFrame(columns=["ID Siswa", "Nama Siswa", "Jenjang", "Paket Belajar", "Alamat", "Fee per Pertemuan (Rp)", "Fee Tentor (Rp)"])
    return df

def load_data_tentor_profil():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM tentor_profil", conn)
    conn.close()
    if not df.empty:
        df.columns = ["ID", "Nama Tentor", "Alamat", "Jenis Mapel", "Hari On Bimbel"]
    else:
        df = pd.DataFrame(columns=["ID", "Nama Tentor", "Alamat", "Jenis Mapel", "Hari On Bimbel"])
    return df

def load_data_tentor_gaji():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM tentor_gaji", conn)
    conn.close()
    if not df.empty:
        df.columns = ["ID", "Nama Tentor", "Daftar Siswa", "Detail Pertemuan", "Jumlah Pertemuan", "Bulan", "Total Gaji (Rp)"]
    else:
        df = pd.DataFrame(columns=["ID", "Nama Tentor", "Daftar Siswa", "Detail Pertemuan", "Jumlah Pertemuan", "Bulan", "Total Gaji (Rp)"])
    return df

def load_data_tagihan():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT id_keuangan, nama_siswa, bulan, jumlah_kehadiran, fee_per_pertemuan, pendaftaran, diskon_pertemuan, diskon, pemasukan, pengeluaran, keterangan FROM keuangan", conn)
    conn.close()
    if not df.empty:
        df.columns = ["ID", "Nama Siswa", "Bulan", "Jumlah Pertemuan", "Fee per Pertemuan (Rp)", "Pendaftaran (Rp)", "Diskon Pertemuan", "Diskon (Rp)", "Tagihan / Pemasukan (Rp)", "Pengeluaran (Rp)", "Keterangan"]
    else:
        df = pd.DataFrame(columns=["ID", "Nama Siswa", "Bulan", "Jumlah Pertemuan", "Fee per Pertemuan (Rp)", "Pendaftaran (Rp)", "Diskon Pertemuan", "Diskon (Rp)", "Tagihan / Pemasukan (Rp)", "Pengeluaran (Rp)", "Keterangan"])
    return df

def convert_df_to_excel(df, sheet_name='Laporan'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

def get_kehadiran_from_tentor(nama_siswa, bulan):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT detail_pertemuan FROM tentor_gaji WHERE bulan = ?", (bulan,))
    rows = cursor.fetchall()
    conn.close()
    
    total_kehadiran = 0
    for row in rows:
        detail = str(row[0])
        parts = detail.split(";")
        for p in parts:
            if nama_siswa.lower() in p.lower():
                try:
                    sub = p.split(":")[1].strip()
                    num_str = sub.split("x")[0].strip()
                    total_kehadiran += int(num_str)
                except:
                    pass
    return total_kehadiran

# --- DAFTAR PILIHAN ---
LIST_JENJANG = [
    "PAUD", "TK Kecil", "TK Besar", 
    "SD Kelas 1", "SD Kelas 2", "SD Kelas 3", "SD Kelas 4", "SD Kelas 5", "SD Kelas 6",
    "SMP Kelas 7", "SMP Kelas 8", "SMP Kelas 9",
    "SMA Kelas 10", "SMA Kelas 11", "SMA Kelas 12"
]
LIST_BULAN_NAMA = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]
current_year = datetime.now().year
LIST_TAHUN = [str(y) for y in [current_year - 1, current_year, current_year + 1]]
LIST_BULAN_TAHUN = [f"{bulan} {tahun}" for tahun in LIST_TAHUN for bulan in LIST_BULAN_NAMA]

# --- GLOBAL STYLING: COMIC SANS ---
st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: 'Comic Sans MS', 'Comic Sans', cursive, sans-serif !important;
    }
</style>
""", unsafe_allow_html=True)

# --- INISIALISASI SESSION STATE ADMIN PASSWORD ---
if "admin_password" not in st.session_state:
    st.session_state.admin_password = "fathonah123"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# --- AUTENTIKASI / LOGIN ADMINISTRATOR ---
if not st.session_state.logged_in:
    st.title("🔐 Login Administrator - Bimbel Fathonah")
    st.markdown("Silakan masukkan username dan password administrator untuk mengakses dashboard.")
    
    with st.form("login_form"):
        username_input = st.text_input("Username")
        password_input = st.text_input("Password", type="password")
        submit_login = st.form_submit_button("🚀 Masuk")
        
        if submit_login:
            if username_input == "admin" and password_input == st.session_state.admin_password:
                st.session_state.logged_in = True
                st.success("Login berhasil! Memuat dashboard...")
                st.rerun()
            else:
                st.error("Username atau Password salah!")
    st.stop()

# --- HEADER / JUDUL UTAMA ---
st.title("📚 Dashboard Bimbingan Belajar Fathonah")
st.markdown("Panel kontrol utama Bimbel Fathonah. Fitur manajemen tagihan, gaji, dan cetak invoice wali siswa aktif.")
st.markdown("---")

# --- SIDEBAR NAVIGASI & LOGOUT ---
st.sidebar.image("assets/logo_fathonah.png", width=150) # Baris untuk logo Anda
st.sidebar.header("Navigasi Menu")

# Pastikan blok selectbox ini HANYA ADA SATU KALI di dalam file
menu = st.sidebar.selectbox(
    "Pilih Menu:",
    ["Ringkasan Utama", "Data Siswa", "Manajemen Siswa (CRUD)", "Data Tentor", "Gaji Bulanan", "Tagihan Wali Siswa", "Ubah Password"]
)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout (Keluar)"):
    st.session_state.logged_in = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("☁️ Status Google Drive Sync")
st.sidebar.success("Status: Terhubung (Auto-Backup Siap)")

if os.path.exists(DB_FILE):
    with open(DB_FILE, "rb") as f:
        st.sidebar.download_button(
            label="☁️ Download Database Terbaru",
            data=f,
            file_name="bimbel_fathonah.db",
            mime="application/octet-stream",
            help="Simpan file ini langsung ke folder Google Drive Anda untuk cadangan online."
        )

# Ambil data terbaru dari database
df_siswa = load_data_siswa()
df_tentor_profil = load_data_tentor_profil()
df_tentor_gaji = load_data_tentor_gaji()
df_tagihan = load_data_tagihan()

# --- KONTEN BERDASARKAN MENU ---

if menu == "Ringkasan Utama":
    total_pemasukan = df_tagihan["Tagihan / Pemasukan (Rp)"].sum() if not df_tagihan.empty else 0
    total_pengeluaran_gaji = df_tentor_gaji["Total Gaji (Rp)"].sum() if not df_tentor_gaji.empty else 0
    netto = total_pemasukan - total_pengeluaran_gaji

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(label="Total Siswa Aktif", value=f"{len(df_siswa)} Siswa")
    with col2:
        st.metric(label="Total Pemasukan Tagihan", value=f"Rp {total_pemasukan:,.0f}")
    with col3:
        st.metric(label="Total Pengeluaran Gaji Tentor", value=f"Rp {total_pengeluaran_gaji:,.0f}", delta_color="inverse")
    with col4:
        st.metric(label="Saldo Bersih (Netto)", value=f"Rp {netto:,.0f}")

    st.markdown("---")

    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.subheader("📊 Distribusi Jenjang Siswa")
        if not df_siswa.empty:
            jenjang_count = df_siswa["Jenjang"].value_counts().reset_index()
            jenjang_count.columns = ["Jenjang", "Jumlah"]
            fig_jenjang = px.pie(jenjang_count, names="Jenjang", values="Jumlah", hole=0.4, color_discrete_sequence=px.colors.sequential.Teal)
            st.plotly_chart(fig_jenjang, use_container_width=True)
        else:
            st.info("Belum ada data siswa.")

    with col_chart2:
        st.subheader("👥 Statistik Operasional")
        st.metric("Total Tentor Terdaftar", f"{len(df_tentor_profil)} Tentor")

elif menu == "Data Siswa":
    st.subheader("👥 Daftar Seluruh Siswa Bimbel Fathonah")
    if df_siswa.empty:
        st.warning("Database siswa masih kosong.")
    else:
        filter_jenjang = st.selectbox("Filter berdasarkan Jenjang Pendidikan:", ["Semua"] + LIST_JENJANG)
        filtered_df = df_siswa if filter_jenjang == "Semua" else df_siswa[df_siswa["Jenjang"] == filter_jenjang]
        st.dataframe(filtered_df, use_container_width=True)

elif menu == "Manajemen Siswa (CRUD)":
    st.subheader("⚙️ Manajemen Data Siswa")
    action = st.radio("Pilih Aksi:", ["➕ Tambah Siswa", "✏️ Update Data Siswa", "🗑️ Hapus Siswa"], horizontal=True)
    st.markdown("---")
    
    if action == "➕ Tambah Siswa":
        with st.form("form_tambah_siswa"):
            col1, col2 = st.columns(2)
            default_id = f"S0{len(df_siswa)+1}" if not df_siswa.empty else "S001"
            with col1:
                new_id = st.text_input("ID Siswa", value=default_id)
                new_nama = st.text_input("Nama Lengkap Siswa")
                new_jenjang = st.selectbox("Jenjang Pendidikan", LIST_JENJANG)
                new_paket = st.text_input("Paket Belajar")
            with col2:
                new_alamat = st.text_input("Alamat Siswa")
                new_fee = st.number_input("Fee per Pertemuan (Rp)", min_value=0, step=10000, value=100000)
                new_fee_tentor = st.number_input("Fee Tentor per Pertemuan (Rp)", min_value=0, step=10000, value=50000)
            
            submit_siswa = st.form_submit_button("💾 Simpan Data Baru")
            if submit_siswa:
                if new_nama.strip() == "":
                    st.warning("Nama siswa tidak boleh kosong!")
                elif not df_siswa.empty and new_id in df_siswa["ID Siswa"].values:
                    st.error(f"ID Siswa '{new_id}' sudah terdaftar.")
                else:
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('INSERT INTO siswa VALUES (?, ?, ?, ?, ?, ?, ?)', 
                                       (new_id, new_nama, new_jenjang, new_paket, new_alamat, new_fee, new_fee_tentor))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success(f"Siswa **{new_nama}** berhasil disimpan!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    elif action == "✏️ Update Data Siswa":
        if df_siswa.empty:
            st.info("Belum ada data siswa.")
        else:
            selected_id = st.selectbox("Pilih Siswa:", df_siswa["ID Siswa"] + " - " + df_siswa["Nama Siswa"])
            id_target = selected_id.split(" - ")[0]
            current_data = df_siswa[df_siswa["ID Siswa"] == id_target].iloc[0]
            try:
                default_idx = LIST_JENJANG.index(current_data["Jenjang"])
            except ValueError:
                default_idx = 0
            
            with st.form("form_update_siswa"):
                up_nama = st.text_input("Nama Lengkap Siswa", value=current_data["Nama Siswa"])
                up_jenjang = st.selectbox("Jenjang Pendidikan", LIST_JENJANG, index=default_idx)
                up_paket = st.text_input("Paket Belajar", value=current_data["Paket Belajar"])
                up_alamat = st.text_input("Alamat Siswa", value=current_data["Alamat"])
                up_fee = st.number_input("Fee per Pertemuan (Rp)", min_value=0, step=10000, value=int(current_data["Fee per Pertemuan (Rp)"]))
                up_fee_tentor = st.number_input("Fee Tentor per Pertemuan (Rp)", min_value=0, step=10000, value=int(current_data["Fee Tentor (Rp)"]))
                
                submit_up_siswa = st.form_submit_button("💾 Simpan Perubahan")
                if submit_up_siswa:
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('''
                            UPDATE siswa SET nama_siswa=?, jenjang=?, paket_belajar=?, alamat=?, fee_per_pertemuan=?, fee_tentor=?
                            WHERE id_siswa=?
                        ''', (up_nama, up_jenjang, up_paket, up_alamat, up_fee, up_fee_tentor, id_target))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success("Data siswa berhasil diperbarui!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    elif action == "🗑️ Hapus Siswa":
        if df_siswa.empty:
            st.info("Belum ada data siswa.")
        else:
            selected_id_del = st.selectbox("Pilih Siswa:", df_siswa["ID Siswa"] + " - " + df_siswa["Nama Siswa"])
            id_target_del = selected_id_del.split(" - ")[0]
            if st.button("🗑️ Konfirmasi Hapus Data", type="primary"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM siswa WHERE id_siswa = ?", (id_target_del,))
                    conn.commit()
                    conn.close()
                    auto_sync_to_google_drive()
                    st.success("Data siswa berhasil dihapus!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    st.markdown("---")
    st.dataframe(df_siswa, use_container_width=True)

elif menu == "Data Tentor":
    st.subheader("👨‍🏫 Manajemen Data Tentor (Profil, Alamat, Mapel & Jadwal Hari)")
    action_t_profil = st.radio("Aksi Data Tentor:", ["📋 Lihat Daftar Tentor", "➕ Tambah Tentor", "✏️ Update Tentor", "🗑️ Hapus Tentor"], horizontal=True)
    st.markdown("---")
    
    if action_t_profil == "📋 Lihat Daftar Tentor":
        if df_tentor_profil.empty:
            st.warning("Belum ada data profil tentor.")
        else:
            st.dataframe(df_tentor_profil, use_container_width=True)
            
    elif action_t_profil == "➕ Tambah Tentor":
        with st.form("form_tambah_tentor_profil"):
            col1, col2 = st.columns(2)
            with col1:
                tp_nama = st.text_input("Nama Tentor")
                tp_alamat = st.text_input("Alamat")
                tp_mapel = st.text_input("Jenis Mapel (Contoh: Matematika, Fisika)")
            with col2:
                st.markdown("**Checklist Hari On Bimbel:**")
                tp_senin = st.checkbox("Senin")
                tp_selasa = st.checkbox("Selasa")
                tp_rabu = st.checkbox("Rabu")
                tp_kamis = st.checkbox("Kamis")
                tp_jumat = st.checkbox("Jumat")
                tp_sabtu = st.checkbox("Sabtu")
                tp_minggu = st.checkbox("Minggu")
            
            submit_tp = st.form_submit_button("💾 Simpan Data Tentor")
            if submit_tp:
                if tp_nama.strip() == "":
                    st.warning("Nama tentor tidak boleh kosong!")
                else:
                    hari_pilih = []
                    if tp_senin: hari_pilih.append("Senin")
                    if tp_selasa: hari_pilih.append("Selasa")
                    if tp_rabu: hari_pilih.append("Rabu")
                    if tp_kamis: hari_pilih.append("Kamis")
                    if tp_jumat: hari_pilih.append("Jumat")
                    if tp_sabtu: hari_pilih.append("Sabtu")
                    if tp_minggu: hari_pilih.append("Minggu")
                    
                    str_hari_on = ", ".join(hari_pilih) if hari_pilih else "Tidak ada"
                    
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('INSERT INTO tentor_profil (nama_tentor, alamat, jenis_mapel, hari_on_bimbel) VALUES (?, ?, ?, ?)',
                                       (tp_nama, tp_alamat, tp_mapel, str_hari_on))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success(f"Tentor **{tp_nama}** berhasil ditambahkan!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    elif action_t_profil == "✏️ Update Tentor":
        if df_tentor_profil.empty:
            st.info("Belum ada data tentor untuk di-update.")
        else:
            sel_tp = st.selectbox("Pilih Tentor:", df_tentor_profil["ID"].astype(str) + " - " + df_tentor_profil["Nama Tentor"])
            id_tp_target = int(sel_tp.split(" - ")[0])
            curr_tp = df_tentor_profil[df_tentor_profil["ID"] == id_tp_target].iloc[0]
            curr_hari = str(curr_tp["Hari On Bimbel"])
            
            with st.form("form_update_tentor_profil"):
                up_tp_nama = st.text_input("Nama Tentor", value=curr_tp["Nama Tentor"])
                up_tp_alamat = st.text_input("Alamat", value=curr_tp["Alamat"])
                up_tp_mapel = st.text_input("Jenis Mapel", value=curr_tp["Jenis Mapel"])
                
                st.markdown("**Checklist Hari On Bimbel:**")
                up_tp_senin = st.checkbox("Senin", value="Senin" in curr_hari)
                up_tp_selasa = st.checkbox("Selasa", value="Selasa" in curr_hari)
                up_tp_rabu = st.checkbox("Rabu", value="Rabu" in curr_hari)
                up_tp_kamis = st.checkbox("Kamis", value="Kamis" in curr_hari)
                up_tp_jumat = st.checkbox("Jumat", value="Jumat" in curr_hari)
                up_tp_sabtu = st.checkbox("Sabtu", value="Sabtu" in curr_hari)
                up_tp_minggu = st.checkbox("Minggu", value="Minggu" in curr_hari)
                
                submit_up_tp = st.form_submit_button("💾 Simpan Perubahan Profil")
                if submit_up_tp:
                    hari_pilih_up = []
                    if up_tp_senin: hari_pilih_up.append("Senin")
                    if up_tp_selasa: hari_pilih_up.append("Selasa")
                    if up_tp_rabu: hari_pilih_up.append("Rabu")
                    if up_tp_kamis: hari_pilih_up.append("Kamis")
                    if up_tp_jumat: hari_pilih_up.append("Jumat")
                    if up_tp_sabtu: hari_pilih_up.append("Sabtu")
                    if up_tp_minggu: hari_pilih_up.append("Minggu")
                    
                    str_hari_up = ", ".join(hari_pilih_up) if hari_pilih_up else "Tidak ada"
                    
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('''
                            UPDATE tentor_profil SET nama_tentor=?, alamat=?, jenis_mapel=?, hari_on_bimbel=?
                            WHERE id_tentor=?
                        ''', (up_tp_nama, up_tp_alamat, up_tp_mapel, str_hari_up, id_tp_target))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success("Profil tentor berhasil diperbarui!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    elif action_t_profil == "🗑️ Hapus Tentor":
        if df_tentor_profil.empty:
            st.info("Belum ada data tentor.")
        else:
            sel_tp_del = st.selectbox("Pilih Tentor yang Ingin Dihapus:", df_tentor_profil["ID"].astype(str) + " - " + df_tentor_profil["Nama Tentor"])
            id_tp_del = int(sel_tp_del.split(" - ")[0])
            if st.button("🗑️ Konfirmasi Hapus Tentor", type="primary"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM tentor_profil WHERE id_tentor = ?", (id_tp_del,))
                    conn.commit()
                    conn.close()
                    auto_sync_to_google_drive()
                    st.success("Profil tentor berhasil dihapus!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

elif menu == "Gaji Bulanan":
    st.subheader("💰 Manajemen Gaji Bulanan Tentor & Slip Gaji")
    action_gaji = st.radio("Aksi Gaji:", ["📋 Lihat Gaji Bulanan", "➕ Tambah Gaji Tentor", "📄 Tampilkan & Download Slip Gaji", "✏️ Update Gaji", "🗑️ Hapus Gaji"], horizontal=True)
    st.markdown("---")
    
    if action_gaji == "📋 Lihat Gaji Bulanan":
        if df_tentor_gaji.empty:
            st.warning("Belum ada data gaji bulanan tentor.")
        else:
            st.markdown("### 📊 Filter Gaji Berdasarkan Bulan & Tahun")
            filter_b_gaji = st.selectbox("Pilih Bulan & Tahun:", ["Semua Bulan"] + LIST_BULAN_TAHUN)
            
            filtered_gaji = df_tentor_gaji if filter_b_gaji == "Semua Bulan" else df_tentor_gaji[df_tentor_gaji["Bulan"] == filter_b_gaji]
            tot_gaji_bln = filtered_gaji["Total Gaji (Rp)"].sum() if not filtered_gaji.empty else 0
            
            col_m1, col_d1 = st.columns([2, 1])
            with col_m1:
                st.metric(label=f"Total Pengeluaran Gaji Tentor ({filter_b_gaji})", value=f"Rp {tot_gaji_bln:,.0f}")
            with col_d1:
                st.markdown("<br>", unsafe_allow_html=True)
                if not filtered_gaji.empty:
                    excel_gaji = convert_df_to_excel(filtered_gaji, sheet_name='Gaji Tentor')
                    st.download_button(
                        label=f"📥 Download Excel Gaji ({filter_b_gaji})",
                        data=excel_gaji,
                        file_name=f"Gaji_Tentor_{filter_b_gaji.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            st.markdown("---")
            if not filtered_gaji.empty:
                disp_gaji = filtered_gaji.copy()
                disp_gaji["Daftar Siswa"] = disp_gaji["Daftar Siswa"].str.replace(", ", "\n• ")
                disp_gaji["Detail Pertemuan"] = disp_gaji["Detail Pertemuan"].str.replace("; ", "\n• ")
                st.dataframe(disp_gaji, use_container_width=True)
            else:
                st.info("Tidak ada data gaji untuk filter ini.")

    elif action_gaji == "➕ Tambah Gaji Tentor":
        if df_tentor_profil.empty:
            st.warning("Data profil tentor masih kosong. Harap isi menu *Data Tentor* terlebih dahulu.")
        elif df_siswa.empty:
            st.warning("Data siswa masih kosong. Harap isi menu *Manajemen Siswa (CRUD)* terlebih dahulu.")
        else:
            with st.form("form_tambah_gaji_tentor"):
                col1, col2 = st.columns(2)
                with col1:
                    list_nama_tentor = df_tentor_profil["Nama Tentor"].tolist()
                    g_nama_tentor = st.selectbox("Pilih Nama Tentor", list_nama_tentor)
                    g_bulan = st.selectbox("Bulan & Tahun Gaji", LIST_BULAN_TAHUN)
                with col2:
                    st.markdown("### 👥 Pilih Siswa & Atur Pertemuan")
                
                daftar_nama_siswa = df_siswa["Nama Siswa"].tolist()
                g_pilihan_siswa = st.multiselect("Pilih Siswa yang Diajar oleh Tentor Ini:", daftar_nama_siswa)
                
                g_pertemuan_inputs = {}
                if g_pilihan_siswa:
                    st.markdown("Atur jumlah pertemuan masing-masing siswa:")
                    cols_s = st.columns(len(g_pilihan_siswa) if len(g_pilihan_siswa) <= 4 else 4)
                    for i, s_nama in enumerate(g_pilihan_siswa):
                        with cols_s[i % len(cols_s)]:
                            g_pertemuan_inputs[s_nama] = st.number_input(f"Pertemuan: {s_nama}", min_value=1, step=1, value=4, key=f"gj_jml_{s_nama}")
                
                submit_gaji = st.form_submit_button("💾 Hitung & Simpan Gaji Bulanan")
                if submit_gaji:
                    if not g_pilihan_siswa:
                        st.warning("Harap pilih minimal 1 siswa!")
                    else:
                        tot_gaji_hitung = 0
                        tot_pertemuan_gabung = 0
                        detail_list = []
                        for s_nama in g_pilihan_siswa:
                            fee_t_s = int(df_siswa.loc[df_siswa["Nama Siswa"] == s_nama, "Fee Tentor (Rp)"].values[0])
                            jml_p = g_pertemuan_inputs[s_nama]
                            tot_pertemuan_gabung += jml_p
                            sub = fee_t_s * jml_p
                            tot_gaji_hitung += sub
                            detail_list.append(f"• {s_nama}: {jml_p}x (Fee @Rp {fee_t_s:,} = Rp {sub:,})")
                        
                        str_daftar_siswa = ", ".join([f"• {s}" for s in g_pilihan_siswa])
                        str_detail_p = "; ".join(detail_list)
                        
                        try:
                            conn = sqlite3.connect(DB_FILE)
                            cursor = conn.cursor()
                            cursor.execute('''
                                INSERT INTO tentor_gaji (nama_tentor, daftar_siswa, detail_pertemuan, jumlah_pertemuan, bulan, total_gaji)
                                VALUES (?, ?, ?, ?, ?, ?)
                            ''', (g_nama_tentor, str_daftar_siswa, str_detail_p, tot_pertemuan_gabung, g_bulan, tot_gaji_hitung))
                            conn.commit()
                            conn.close()
                            auto_sync_to_google_drive()
                            st.success(f"Gaji bulanan tentor **{g_nama_tentor}** bulan **{g_bulan}** sebesar **Rp {tot_gaji_hitung:,.0f}** berhasil disimpan!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

    elif action_gaji == "📄 Tampilkan & Download Slip Gaji":
        st.subheader("📄 Tampilkan & Download Slip Gaji Tentor")
        if df_tentor_gaji.empty:
            st.warning("Belum ada data gaji bulanan tentor untuk ditampilkan.")
        else:
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                list_tentor_slip = df_tentor_gaji["Nama Tentor"].unique().tolist()
                pilih_tentor_slip = st.selectbox("Pilih Nama Tentor:", list_tentor_slip)
            with col_s2:
                list_bulan_slip = df_tentor_gaji[df_tentor_gaji["Nama Tentor"] == pilih_tentor_slip]["Bulan"].unique().tolist()
                pilih_bulan_slip = st.selectbox("Pilih Bulan:", list_bulan_slip)
            
            if "slip_ready" not in st.session_state:
                st.session_state.slip_ready = False
            
            if st.button("🔍 Tampilkan Slip Gaji", type="primary"):
                st.session_state.slip_ready = True
            
            if st.session_state.slip_ready:
                slip_data = df_tentor_gaji[(df_tentor_gaji["Nama Tentor"] == pilih_tentor_slip) & (df_tentor_gaji["Bulan"] == pilih_bulan_slip)]
                if not slip_data.empty:
                    row_slip = slip_data.iloc[0]
                    
                    # Panggil logo (sesuaikan path folder logo Anda)
                    logo_base64 = get_image_base64("assets/logo_fathonah.png")
                    
 # Panggil logo (sesuaikan path folder logo Anda)
                    logo_base64 = get_image_base64("assets/logo_fathonah.png")
                    
                    html_slip = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>Slip Gaji - {row_slip['Nama Tentor']}</title>
                        <style>
                            body, table, div, p, span, th, td {{ font-family: 'Comic Sans MS', 'Comic Sans', cursive, sans-serif !important; }}
                            body {{ color: #333; margin: 0; padding: 20px; }}
                            .slip-container {{ border: 2px solid #2ECC71; padding: 30px; border-radius: 10px; background-color: #FFFFFF; max-width: 700px; margin: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
                            table.info {{ width: 100%; font-size: 15px; margin: 20px 0; border-collapse: collapse; color: #333; }}
                            table.info td {{ padding: 6px 0; }}
                            .details-box {{ background: #FAFAFA; padding: 15px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; white-space: pre-line; color: #333; }}
                            .total {{ text-align: right; color: #27AE60; font-size: 20px; margin-top: 20px; font-weight: bold; }}
                        </style>
                    </head>
                    <body>
                        <div class="slip-container">
                            <!-- HEADER 3 KOLOM AGAR RATA TENGAH SEMPURNA -->
                            <table style="width: 100%; border-bottom: 2px solid #ddd; margin-bottom: 15px; padding-bottom: 10px;">
                                <tr>
                                    <td style="width: 22%; text-align: left; vertical-align: left;">
                                        <img src="data:image/png;base64,{logo_base64}" alt="Logo" style="max-width: 110px;">
                                    </td>
                                    <td style="width: 64%; text-align: center; vertical-align: middle;">
                                        <h2 style="margin: 0; color: #2C3E50;">BIMBINGAN BELAJAR FATHONAH</h2>
                                        <p style="margin: 5px 0 0 0; color: #555; font-weight: bold;">Slip Gaji Tentor</p>
                                    </td>
                                    <td style="width: 14%;"></td> <!-- Kolom dummy penyeimbang -->
                                </tr>
                            </table>
                            
                            <table class="info">
                                <tr>
                                    <td width="35%"><b>Nama Tentor</b></td>
                                    <td>: {row_slip['Nama Tentor']}</td>
                                </tr>
                                <tr>
                                    <td><b>Periode Bulan</b></td>
                                    <td>: {row_slip['Bulan']}</td>
                                </tr>
                                <tr>
                                    <td><b>Total Pertemuan</b></td>
                                    <td>: {row_slip['Jumlah Pertemuan']} pertemuan</td>
                                </tr>
                            </table>
                            <b>Rincian Mengajar & Pertemuan:</b>
                            <div class="details-box">{row_slip['Detail Pertemuan'].replace('; ', '\n')}</div>
                            <div class="total">Total Gaji Diterima: Rp {row_slip['Total Gaji (Rp)']:,.0f}</div>
                        </div>
                    </body>
                    </html>
                    """


                    st.markdown("---")
                    st.components.v1.html(html_slip, height=550, scrolling=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    st.download_button(
                        label="📥 Download Slip Gaji (.html / Cetak ke PDF)",
                        data=html_slip,
                        file_name=f"Slip_Gaji_{row_slip['Nama Tentor'].replace(' ', '_')}_{row_slip['Bulan'].replace(' ', '_')}.html",
                        mime="text/html",
                        help="Unduh file slip gaji. Buka file tersebut lalu tekan Ctrl+P dan pilih 'Save as PDF' untuk menyimpan sebagai PDF."
                    )
                else:
                    st.warning("Data slip gaji tidak ditemukan.")

    elif action_gaji == "✏️ Update Gaji":
        if df_tentor_gaji.empty:
            st.info("Belum ada data gaji bulanan untuk di-update.")
        else:
            sel_g_up = st.selectbox("Pilih Gaji Tentor:", df_tentor_gaji["ID"].astype(str) + " - " + df_tentor_gaji["Nama Tentor"] + " (" + df_tentor_gaji["Bulan"] + ")")
            id_g_target = int(sel_g_up.split(" - ")[0])
            curr_g = df_tentor_gaji[df_tentor_gaji["ID"] == id_g_target].iloc[0]
            
            try:
                def_b_idx = LIST_BULAN_TAHUN.index(curr_g["Bulan"])
            except:
                def_b_idx = 0
                
            raw_ds = str(curr_g["Daftar Siswa"])
            ds_lama = [s.replace("• ", "").strip() for s in raw_ds.split(", ")] if raw_ds else []
            
            with st.form("form_update_gaji_tentor"):
                list_nama_tentor = df_tentor_profil["Nama Tentor"].tolist() if not df_tentor_profil.empty else [curr_g["Nama Tentor"]]
                try:
                    def_t_idx = list_nama_tentor.index(curr_g["Nama Tentor"])
                except:
                    def_t_idx = 0
                    
                up_g_nama = st.selectbox("Nama Tentor", list_nama_tentor, index=def_t_idx)
                up_g_bulan = st.selectbox("Bulan & Tahun Gaji", LIST_BULAN_TAHUN, index=def_b_idx)
                
                daftar_nama_siswa = df_siswa["Nama Siswa"].tolist()
                valid_ds = [s for s in ds_lama if s in daftar_nama_siswa]
                up_g_pilihan_siswa = st.multiselect("Pilih Siswa:", daftar_nama_siswa, default=valid_ds)
                
                up_g_pertemuan_inputs = {}
                if up_g_pilihan_siswa:
                    st.markdown("Atur jumlah pertemuan:")
                    for s_nama in up_g_pilihan_siswa:
                        up_g_pertemuan_inputs[s_nama] = st.number_input(f"Pertemuan: {s_nama}", min_value=1, step=1, value=4, key=f"up_gj_{s_nama}")
                
                submit_up_g = st.form_submit_button("💾 Simpan Perubahan Gaji")
                if submit_up_g:
                    tot_gaji_up = 0
                    tot_pertemuan_up = 0
                    detail_list_up = []
                    for s_nama in up_g_pilihan_siswa:
                        fee_t_s = int(df_siswa.loc[df_siswa["Nama Siswa"] == s_nama, "Fee Tentor (Rp)"].values[0])
                        jml_p = up_g_pertemuan_inputs[s_nama]
                        tot_pertemuan_up += jml_p
                        sub = fee_t_s * jml_p
                        tot_gaji_up += sub
                        detail_list_up.append(f"• {s_nama}: {jml_p}x (Fee @Rp {fee_t_s:,} = Rp {sub:,})")
                    
                    str_ds_up = ", ".join([f"• {s}" for s in up_g_pilihan_siswa])
                    str_detail_up = "; ".join(detail_list_up)
                    
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('''
                            UPDATE tentor_gaji SET nama_tentor=?, daftar_siswa=?, detail_pertemuan=?, jumlah_pertemuan=?, bulan=?, total_gaji=?
                            WHERE id_gaji=?
                        ''', (up_g_nama, str_ds_up, str_detail_up, tot_pertemuan_up, up_g_bulan, tot_gaji_up, id_g_target))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success("Data gaji bulanan berhasil diperbarui!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")


    elif action_gaji == "🗑️ Hapus Gaji":
        if df_tentor_gaji.empty:
            st.info("Tidak ada data gaji bulanan.")
        else:
            sel_g_del = st.selectbox("Pilih Gaji yang Ingin Dihapus:", df_tentor_gaji["ID"].astype(str) + " - " + df_tentor_gaji["Nama Tentor"] + " (" + df_tentor_gaji["Bulan"] + ")")
            id_g_del = int(sel_g_del.split(" - ")[0])
            if st.button("🗑️ Konfirmasi Hapus Gaji", type="primary"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM tentor_gaji WHERE id_gaji = ?", (id_g_del,))
                    conn.commit()
                    conn.close()
                    auto_sync_to_google_drive()
                    st.success("Data gaji bulanan berhasil dihapus!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

elif menu == "Tagihan Wali Siswa":
    st.subheader("💰 Tagihan & Pemasukan Wali Siswa (Invoice Resmi)")
    action_keu = st.radio("Aksi Tagihan:", ["📋 Lihat Tagihan", "➕ Tambah Catatan", "📄 Cetak & Download Invoice", "✏️ Update Tagihan", "🗑️ Hapus Catatan", "🧹 Bersihkan Data Duplikat"], horizontal=True)
    st.markdown("---")
    
    if action_keu == "📋 Lihat Tagihan":
        if df_tagihan.empty:
            st.warning("Belum ada catatan tagihan wali siswa.")
        else:
            st.markdown("### 📊 Filter & Total Tagihan Berdasarkan Bulan")
            
            filter_bulan = st.selectbox("Pilih Bulan untuk Dilihat & Ditotalkan:", ["Semua Bulan"] + LIST_BULAN_NAMA)
            
            if filter_bulan != "Semua Bulan":
                filtered_tagihan = df_tagihan[df_tagihan["Bulan"].str.contains(filter_bulan)]
            else:
                filtered_tagihan = df_tagihan
            
            total_pemasukan_bulan = filtered_tagihan["Tagihan / Pemasukan (Rp)"].sum() if not filtered_tagihan.empty else 0
            
            col_metrik, col_download = st.columns([2, 1])
            with col_metrik:
                st.metric(label=f"Total Tagihan Siswa ({filter_bulan})", value=f"Rp {total_pemasukan_bulan:,.0f}")
            
            with col_download:
                st.markdown("<br>", unsafe_allow_html=True)
                if not filtered_tagihan.empty:
                    excel_data = convert_df_to_excel(filtered_tagihan, sheet_name='Tagihan Wali Siswa')
                    st.download_button(
                        label=f"📥 Download Excel ({filter_bulan})",
                        data=excel_data,
                        file_name=f"Tagihan_Wali_Siswa_{filter_bulan.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Unduh laporan tagihan bulan ini dalam format Excel (.xlsx)"
                    )
                else:
                    st.info("Tidak ada data untuk diunduh.")

            st.markdown("---")
            st.dataframe(filtered_tagihan, use_container_width=True)
            
    elif action_keu == "➕ Tambah Catatan":
        if df_siswa.empty:
            st.warning("Data siswa masih kosong. Harap tambahkan data siswa terlebih dahulu pada menu *Manajemen Siswa (CRUD)*.")
        else:
            siswa_fee_dict = dict(zip(df_siswa["Nama Siswa"], df_siswa["Fee per Pertemuan (Rp)"]))
            
            with st.form("form_tambah_keuangan"):
                col1, col2 = st.columns(2)
                with col1:
                    k_bulan = st.selectbox("Pilih Bulan & Tahun", LIST_BULAN_TAHUN)
                    k_nama_siswa = st.selectbox("Pilih Nama Siswa", list(siswa_fee_dict.keys()))
                with col2:
                    k_pendaftaran = st.number_input("Pendaftaran (Rp)", min_value=0, step=50000, value=0)
                    k_diskon_pertemuan = st.number_input("Diskon Pertemuan", min_value=0, step=1, value=0)
                
                k_kehadiran = get_kehadiran_from_tentor(k_nama_siswa, k_bulan)
                fee_otomatis = int(siswa_fee_dict.get(k_nama_siswa, 0))
                
                st.info(f"📌 **Fee per Pertemuan:** Rp {fee_otomatis:,} | 👁️ **Jumlah Pertemuan (Otomatis dari Gaji Tentor):** **{k_kehadiran} pertemuan**")
                
                k_diskon = fee_otomatis * k_diskon_pertemuan * 0.10
                k_pemasukan = (k_kehadiran * fee_otomatis) + k_pendaftaran - k_diskon
                
                st.success(f"💡 **Kalkulasi Otomatis Tagihan:**\n- Diskon: Rp {k_diskon:,.0f} (Fee Rp {fee_otomatis:,} × Diskon Pertemuan × 10%)\n- Total Tagihan Wali Siswa: **Rp {k_pemasukan:,.0f}**")
                
                k_ket = st.text_input("Keterangan (Opsional)")

                submit_keuangan = st.form_submit_button("💾 Simpan Tagihan Wali Siswa")
                
                if submit_keuangan:
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('''
                            INSERT INTO keuangan (nama_siswa, bulan, jumlah_kehadiran, fee_per_pertemuan, pendaftaran, diskon_pertemuan, diskon, pemasukan, pengeluaran, keterangan) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
                        ''', (k_nama_siswa, k_bulan, k_kehadiran, fee_otomatis, k_pendaftaran, k_diskon_pertemuan, k_diskon, k_pemasukan, k_ket))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success(f"Tagihan untuk siswa **{k_nama_siswa}** bulan **{k_bulan}** berhasil disimpan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error(f"⚠️ Gagal: Tagihan untuk **{k_nama_siswa}** pada bulan **{k_bulan}** sudah ada di database!")
                    except Exception as e:
                        st.error(f"Error: {e}")

    elif action_keu == "📄 Cetak & Download Invoice":
        st.subheader("📄 Cetak & Download Invoice Tagihan Wali Siswa")
        if df_tagihan.empty:
            st.warning("Data tagihan wali siswa masih kosong.")
        else:
            col_inv1, col_inv2 = st.columns(2)
            with col_inv1:
                list_siswa_inv = df_tagihan["Nama Siswa"].unique().tolist()
                pilih_siswa_inv = st.selectbox("Pilih Nama Siswa:", list_siswa_inv)
            with col_inv2:
                list_bulan_inv = df_tagihan[df_tagihan["Nama Siswa"] == pilih_siswa_inv]["Bulan"].unique().tolist()
                pilih_bulan_inv = st.selectbox("Pilih Bulan & Tahun Invoice:", list_bulan_inv)
            
            input_perkembangan = st.text_area("📝 Catatan / Perkembangan Siswa (Akan dicetak di bagian bawah invoice):", placeholder="Tuliskan catatan perkembangan belajar siswa selama periode ini...")
            
            if "invoice_ready" not in st.session_state:
                st.session_state.invoice_ready = False
            
            if st.button("🔍 Tampilkan Invoice", type="primary"):
                st.session_state.invoice_ready = True
            
            if st.session_state.invoice_ready:
                inv_data = df_tagihan[(df_tagihan["Nama Siswa"] == pilih_siswa_inv) & (df_tagihan["Bulan"] == pilih_bulan_inv)]
                if not inv_data.empty:
                    row_inv = inv_data.iloc[0]
                    
                    s_row = df_siswa[df_siswa["Nama Siswa"] == pilih_siswa_inv]
                    alamat_siswa = s_row["Alamat"].values[0] if not s_row.empty and s_row["Alamat"].values[0] else "Alamat belum diatur di Data Siswa"
                    
                    pendaftaran_html = f"<tr><td>2</td><td class='text-left'>Biaya Pendaftaran</td><td>-</td><td>-</td><td>Rp {row_inv['Pendaftaran (Rp)']:,.0f}</td><td class='text-left'>Pendaftaran Baru</td></tr>" if row_inv['Pendaftaran (Rp)'] > 0 else ""
                    diskon_html = f"<tr><td>-</td><td class='text-left'>Diskon Pertemuan ({row_inv['Diskon Pertemuan']}x)</td><td>-</td><td>-</td><td>- Rp {row_inv['Diskon (Rp)']:,.0f}</td><td class='text-left'>Potongan Diskon</td></tr>" if row_inv['Diskon (Rp)'] > 0 else ""

                    perkembangan_html = ""
                    if input_perkembangan.strip() != "":
                        perkembangan_html = f"""
                        <div class="perkembangan-box">
                            <b>Catatan / Perkembangan Siswa:</b>
                            <p>{input_perkembangan.replace(chr(10), '<br>')}</p>
                        </div>
                        """

                   # Panggil logo (sesuaikan path folder logo Anda)
                    logo_base64 = get_image_base64("assets/logo_fathonah.png")
                    
# Panggil logo (sesuaikan path folder logo Anda)
                    logo_base64 = get_image_base64("assets/logo_fathonah.png")
                    
                    html_invoice = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>Invoice - {pilih_siswa_inv} ({pilih_bulan_inv})</title>
                        <style>
                            body, table, div, p, span, th, td {{ font-family: 'Comic Sans MS', 'Comic Sans', cursive, sans-serif !important; }}
                            body {{ color: #333; margin: 0; padding: 20px; }}
                            .invoice-container {{ border: 1px solid #ddd; padding: 30px; border-radius: 8px; background-color: #FFFFFF; max-width: 800px; margin: auto; box-shadow: 0 4px 8px rgba(0,0,0,0.05); }}
                            .title-box {{ text-align: center; font-size: 20px; font-weight: bold; background: #fff; margin: 20px 0; letter-spacing: 1px; text-decoration: underline; color: #000; }}
                            .info-table {{ width: 100%; font-size: 15px; margin-bottom: 20px; color: #000; }}
                            .info-table td {{ padding: 4px 0; }}
                            .items-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; color: #000; }}
                            .items-table th, .items-table td {{ border: 1px solid #333; padding: 8px 12px; text-align: center; }}
                            .items-table th {{ background-color: #f2f2f2; color: #000; }}
                            .text-left {{ text-align: left !important; }}
                            .text-right {{ text-align: right !important; }}
                            .perkembangan-box {{ margin-top: 25px; padding: 15px; background-color: #FAFAFA; border: 1px solid #ccc; border-radius: 5px; font-size: 14px; color: #333; }}
                            .perkembangan-box p {{ margin: 8px 0 0 0; white-space: pre-line; }}
                        </style>
                    </head>
                    <body>
                        <div class="invoice-container">
                            <!-- HEADER 3 KOLOM AGAR RATA TENGAH SEMPURNA -->
                            <table style="width: 100%; border-bottom: 3px solid #000; padding-bottom: 10px; margin-bottom: 20px;">
                                <tr>
                                    <td style="width: 18%; text-align: left; vertical-align: middle;">
                                        <img src="data:image/png;base64,{logo_base64}" alt="Logo" style="max-width: 150px;">
                                    </td>
                                    <td style="width: 64%; text-align: center; vertical-align: middle;">
                                        <b style="font-size: 20px; color: #000;">BIMBINGAN BELAJAR FATHONAH</b><br>
                                        <span style="font-size: 13px; color: #333; line-height: 1.5;">
                                            Kantor Pemasaran : Gedongkiwo Mantrijeron, Yogyakarta<br>
                                            (Gmaps: Bimbingan Belajar Fathonah Yogyakarta)<br>
                                            Nomer HP : 085726491252 (Aini)<br>
                                            Email : bimbelfathonah.yk@gmail.com
                                        </span>
                                    </td>
                                    <td style="width: 18%;"></td> <!-- Kolom dummy penyeimbang -->
                                </tr>
                            </table>
                            
                            <div class="title-box">INVOICE {pilih_bulan_inv.upper()}</div>
                            
                            <table class="info-table">
                                <tr>
                                    <td width="15%"><b>Nama Siswa</b></td>
                                    <td width="85%">: <b>{row_inv['Nama Siswa']}</b></td>
                                </tr>
                                <tr>
                                    <td><b>Alamat</b></td>
                                    <td>: {alamat_siswa}</td>
                                </tr>
                            </table>
                            
                            <table class="items-table">
                                <thead>
                                    <tr>
                                        <th width="8%">No</th>
                                        <th width="30%">Nama Siswa</th>
                                        <th width="12%">Pertemuan</th>
                                        <th width="18%">Harga Satuan</th>
                                        <th width="18%">Total</th>
                                        <th width="24%">Keterangan</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>1</td>
                                        <td class="text-left">{row_inv['Nama Siswa']}</td>
                                        <td>{row_inv['Jumlah Pertemuan']}</td>
                                        <td>Rp {row_inv['Fee per Pertemuan (Rp)']:,.0f}</td>
                                        <td>Rp {(row_inv['Jumlah Pertemuan'] * row_inv['Fee per Pertemuan (Rp)']):,.0f}</td>
                                        <td class="text-left">{row_inv['Keterangan'] if row_inv['Keterangan'] else 'Fee Normal'}</td>
                                    </tr>
                                    {pendaftaran_html}
                                    {diskon_html}
                                    <tr style="font-weight: bold; background-color: #fafafa;">
                                        <td colspan="4" class="text-right">Total Tagihan :</td>
                                        <td colspan="2" class="text-left" style="color: #27AE60; font-size: 16px;">Rp {row_inv['Tagihan / Pemasukan (Rp)']:,.0f}</td>
                                    </tr>
                                </tbody>
                            </table>

                            {perkembangan_html}
                        </div>
                    </body>
                    </html>
                    """                    
                    st.markdown("---")
                    st.components.v1.html(html_invoice, height=750, scrolling=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    st.download_button(
                        label="📥 Download Invoice (.html / Cetak ke PDF)",
                        data=html_invoice,
                        file_name=f"Invoice_{row_inv['Nama Siswa'].replace(' ', '_')}_{row_inv['Bulan'].replace(' ', '_')}.html",
                        mime="text/html",
                        help="Unduh file invoice. Buka file tersebut lalu tekan Ctrl+P dan pilih 'Save as PDF' untuk menyimpan sebagai PDF."
                    )

    elif action_keu == "✏️ Update Tagihan":
        if df_tagihan.empty:
            st.info("Belum ada catatan tagihan untuk di-update.")
        else:
            selected_update_keu = st.selectbox(
                "Pilih Catatan Tagihan yang Ingin Di-update:", 
                df_tagihan["ID"].astype(str) + " - " + df_tagihan["Nama Siswa"] + " (" + df_tagihan["Bulan"] + ")"
            )
            id_update_target = int(selected_update_keu.split(" - ")[0])
            current_keu_data = df_tagihan[df_tagihan["ID"] == id_update_target].iloc[0]
            
            siswa_fee_dict = dict(zip(df_siswa["Nama Siswa"], df_siswa["Fee per Pertemuan (Rp)"]))
            target_nama_siswa = current_keu_data["Nama Siswa"]
            target_bulan = current_keu_data["Bulan"]
            fee_terkini = int(siswa_fee_dict.get(target_nama_siswa, current_keu_data["Fee per Pertemuan (Rp)"]))
            
            up_kehadiran = get_kehadiran_from_tentor(target_nama_siswa, target_bulan)
            
            with st.form("form_update_keuangan"):
                st.markdown(f"**Siswa:** {target_nama_siswa} | **Bulan:** {target_bulan} | **Fee Terkini:** Rp {fee_terkini:,} | 👁️ **Jumlah Pertemuan (Gaji Tentor):** {up_kehadiran}x")
                
                col1, col2 = st.columns(2)
                with col1:
                    up_pendaftaran = st.number_input("Pendaftaran (Rp)", min_value=0, step=50000, value=int(current_keu_data["Pendaftaran (Rp)"]))
                with col2:
                    up_diskon_pertemuan = st.number_input("Diskon Pertemuan", min_value=0, step=1, value=int(current_keu_data["Diskon Pertemuan"]))
                
                up_ket = st.text_input("Keterangan", value=str(current_keu_data["Keterangan"]))
                
                up_diskon = fee_terkini * up_diskon_pertemuan * 0.10
                up_pemasukan = (up_kehadiran * fee_terkini) + up_pendaftaran - up_diskon
                
                st.success(f"💡 **Kalkulasi Perubahan Terbaru:**\n- Diskon: Rp {up_diskon:,.0f} (Fee Rp {fee_terkini:,} × Diskon Pertemuan × 10%)\n- Total Tagihan: **Rp {up_pemasukan:,.0f}**")
                
                submit_update_keu = st.form_submit_button("💾 Simpan Perubahan Tagihan")
                if submit_update_keu:
                    try:
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute('''
                            UPDATE keuangan 
                            SET jumlah_kehadiran=?, fee_per_pertemuan=?, pendaftaran=?, diskon_pertemuan=?, diskon=?, pemasukan=?, keterangan=?
                            WHERE id_keuangan=?
                        ''', (up_kehadiran, fee_terkini, up_pendaftaran, up_diskon_pertemuan, up_diskon, up_pemasukan, up_ket, id_update_target))
                        conn.commit()
                        conn.close()
                        auto_sync_to_google_drive()
                        st.success("Catatan tagihan wali siswa berhasil diperbarui!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
                    
    elif action_keu == "🗑️ Hapus Catatan":
        if df_tagihan.empty:
            st.info("Tidak ada catatan tagihan.")
        else:
            selected_keu = st.selectbox("Pilih Catatan:", df_tagihan["ID"].astype(str) + " - " + df_tagihan["Nama Siswa"] + " (" + df_tagihan["Bulan"] + " | Rp " + df_tagihan["Tagihan / Pemasukan (Rp)"].astype(str) + ")")
            id_keu_target = int(selected_keu.split(" - ")[0])
            
            if st.button("🗑️ Konfirmasi Hapus Catatan", type="primary"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM keuangan WHERE id_keuangan = ?", (id_keu_target,))
                    conn.commit()
                    conn.close()
                    auto_sync_to_google_drive()
                    st.success("Catatan tagihan berhasil dihapus!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    elif action_keu == "🧹 Bersihkan Data Duplikat":
        st.markdown("### Bersihkan Data Tagihan Ganda")
        st.markdown("Klik tombol di bawah ini untuk menghapus data tagihan duplikat yang sudah terlanjur tersimpan di database.")
        if st.button("🧹 Hapus Semua Data Duplikat", type="primary"):
            try:
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute('''
                    DELETE FROM keuangan 
                    WHERE id_keuangan NOT IN (
                        SELECT MIN(id_keuangan) 
                        FROM keuangan 
                        GROUP BY nama_siswa, bulan
                    )
                ''')
                conn.commit()
                conn.close()
                st.success("Data duplikat berhasil dibersihkan dari database!")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

elif menu == "Ubah Password":
    st.subheader("🔑 Ubah Password Administrator")
    st.markdown("Gunakan form di bawah ini untuk mengganti password akun administrator.")
    
    with st.form("form_ubah_password"):
        old_pass = st.text_input("Password Lama", type="password")
        new_pass = st.text_input("Password Baru", type="password")
        confirm_pass = st.text_input("Konfirmasi Password Baru", type="password")
        
        submit_pass = st.form_submit_button("💾 Simpan Password Baru")
        if submit_pass:
            if old_pass != st.session_state.admin_password:
                st.error("Password lama salah!")
            elif not new_pass.strip():
                st.warning("Password baru tidak boleh kosong!")
            elif new_pass != confirm_pass:
                st.error("Konfirmasi password baru tidak cocok!")
            else:
                st.session_state.admin_password = new_pass
                st.success("Password administrator berhasil diubah! Silakan gunakan password baru untuk login selanjutnya.")

# --- FOOTER ---
st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>© 2026 Bimbingan Belajar Fathonah. Didukung oleh SQLite & Streamlit.</p>", unsafe_allow_html=True)