import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="Executive Admin Kas RT", page_icon="👑", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 2. INJEKSI CSS KUSTOM
# ==========================================
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        transition: transform 0.2s ease-in-out;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-5px);
        border: 1px solid #4CAF50;
    }
    hr {
        border: 0;
        height: 1px;
        background-image: linear-gradient(to right, rgba(255,255,255,0), rgba(255,255,255,0.5), rgba(255,255,255,0));
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. DATABASE & FUNGSI PENDUKUNG
# ==========================================
FILE_DATA = 'database_kas_rt.csv'
FILE_WARGA = 'data_warga.xlsx'

def muat_data():
    if not os.path.exists(FILE_DATA):
        df = pd.DataFrame(columns=['ID', 'Tanggal', 'Jenis', 'Kategori', 'Keterangan', 'Blok_Warga', 'Nominal'])
        df.to_csv(FILE_DATA, index=False)
        return df
    return pd.read_csv(FILE_DATA)

def simpan_data(df):
    df.to_csv(FILE_DATA, index=False)

# Fungsi baru untuk memuat daftar warga dari Excel khusus
def muat_warga():
    if not os.path.exists(FILE_WARGA):
        # Membuat file template jika belum ada
        df_warga = pd.DataFrame({
            'Nama_dan_Blok': ['Blok A1 - Pak Budi', 'Blok A2 - Pak Andi', 'Blok B1 - Bu Siti', 'Fasum', 'Lainnya']
        })
        df_warga.to_excel(FILE_WARGA, index=False, engine='openpyxl')
        return df_warga['Nama_dan_Blok'].tolist()
    else:
        try:
            df_warga = pd.read_excel(FILE_WARGA, engine='openpyxl')
            return df_warga.iloc[:, 0].astype(str).tolist() # Mengambil kolom pertama apapun namanya
        except:
            return ["Error membaca file warga", "Lainnya"]

def export_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laporan_Kas')
    return output.getvalue()

df = muat_data()
DAFTAR_BLOK = muat_warga()

# ==========================================
# 4. SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='text-align: center;'>👑 Admin Panel</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Sistem Administrasi RT Babelan</p>", unsafe_allow_html=True)
    st.divider()
    
    menu = st.radio(
        "📂 PILIH MENU:",
        ["📊 Dashboard Utama", "📝 Input Transaksi", "📁 Rekap & Laporan"]
    )
    
    st.divider()
    st.caption("© 2026 | Dikembangkan untuk manajemen lingkungan yang transparan.")

# ==========================================
# 5. ROUTING HALAMAN BERDASARKAN MENU
# ==========================================

# ------------------------------------------
# HALAMAN: DASHBOARD UTAMA
# ------------------------------------------
if menu == "📊 Dashboard Utama":
    st.title("Ringkasan Keuangan Lingkungan")
    st.markdown("Pantau arus kas, pemasukan iuran, dan pengeluaran operasional secara *real-time*.")
    st.divider()
    
    if not df.empty:
        total_masuk = df[df['Jenis'] == 'Pemasukan']['Nominal'].sum()
        total_keluar = df[df['Jenis'] == 'Pengeluaran']['Nominal'].sum()
        saldo = total_masuk - total_keluar
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Pemasukan Kas (In)", f"Rp {total_masuk:,.0f}")
        col2.metric("Pengeluaran (Out)", f"Rp {total_keluar:,.0f}")
        col3.metric("Saldo Tersedia", f"Rp {saldo:,.0f}")
        
        st.write("")
        st.write("")
        st.subheader("📈 Analisis Arus Kas")
        
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        df_grafik = df.dropna(subset=['Tanggal']).copy()
        
        if not df_grafik.empty:
            df_group = df_grafik.groupby(['Tanggal', 'Jenis'])['Nominal'].sum().unstack().fillna(0)
            st.area_chart(df_group)
        else:
            st.info("Belum ada data tanggal yang valid untuk ditampilkan di grafik.")
    else:
        st.info("💡 Belum ada data transaksi yang tercatat. Silakan masuk ke menu Input Transaksi.")

# ------------------------------------------
# HALAMAN: INPUT TRANSAKSI
# ------------------------------------------
elif menu == "📝 Input Transaksi":
    st.title("Formulir Pencatatan Baru")
    st.markdown("Masukkan detail transaksi. Logika formulir akan menyesuaikan secara otomatis.")
    st.write("")
    
    # Menghilangkan 'st.form' agar UI bisa merespons seketika saat 'Pemasukan/Pengeluaran' diklik
    st.subheader("1. Informasi Dasar")
    col_a, col_b = st.columns(2)
    
    with col_a:
        jenis_trx = st.radio("Jenis Transaksi", ["Pemasukan", "Pengeluaran"], horizontal=True)
        tanggal = st.date_input("Tanggal Transaksi", datetime.today())
        
    with col_b:
        # LOGIKA NOMINAL PINTAR
        if jenis_trx == "Pemasukan":
            opsi_nominal = {
                "35 Ribu (Rp 35.000)": 35000,
                "50 Ribu (Rp 50.000)": 50000,
                "60 Ribu (Rp 60.000)": 60000,
                "70 Ribu (Rp 70.000)": 70000,
                "80 Ribu (Rp 80.000)": 80000,
                "90 Ribu (Rp 90.000)": 90000,
                "100 Ribu (Rp 100.000)": 100000,
                "Input Manual (Lainnya)": "manual"
            }
            pilihan = st.selectbox("Besaran Nominal", list(opsi_nominal.keys()))
            
            if opsi_nominal[pilihan] == "manual":
                nominal = st.number_input("Masukkan Nominal Bebas (Rp)", min_value=0, step=5000)
            else:
                nominal = opsi_nominal[pilihan]
        else:
            # Jika pengeluaran, langsung tampilkan input angka karena pengeluaran jarang pas
            nominal = st.number_input("Besaran Pengeluaran (Rp)", min_value=0, step=5000)
            
    st.divider()
    st.subheader("2. Detail & Klasifikasi")
    
    col_c, col_d = st.columns(2)
    with col_c:
        kategori = st.selectbox("Kategori Alokasi", ["Iuran Bulanan", "Sumbangan Warga", "Biaya Kebersihan", "Gaji Keamanan", "Perbaikan Fasilitas", "Lain-lain"])
    with col_d:
        blok = st.selectbox("Blok / Identitas Warga (Otomatis dari Excel)", DAFTAR_BLOK)
        
    keterangan = st.text_input("Catatan Tambahan (Opsional)", placeholder="Contoh: Iuran bulan September")
    
    st.write("")
    submit_btn = st.button("💾 Simpan Data Transaksi", use_container_width=True, type="primary")
    
    if submit_btn:
        if nominal <= 0:
            st.error("⚠️ Transaksi ditolak: Nominal tidak boleh Rp 0!")
        else:
            id_baru = f"TRX-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            data_baru = pd.DataFrame([{
                'ID': id_baru,
                'Tanggal': tanggal,
                'Jenis': jenis_trx,
                'Kategori': kategori,
                'Keterangan': keterangan,
                'Blok_Warga': blok if jenis_trx == 'Pemasukan' else '-',
                'Nominal': nominal
            }])
            
            df = pd.concat([df, data_baru], ignore_index=True)
            simpan_data(df)
            
            # Notifikasi pop-up elegan di pojok kanan bawah
            st.toast("✅ Berhasil! Data masuk ke buku besar.", icon='🎉')

# ------------------------------------------
# HALAMAN: REKAP & LAPORAN
# ------------------------------------------
elif menu == "📁 Rekap & Laporan":
    st.title("Laporan Keuangan & Database")
    st.markdown("Filter, periksa, dan unduh rekapan kas lingkungan dalam format Excel.")
    
    col_filter1, col_filter2 = st.columns([1, 2])
    with col_filter1:
        filter_jenis = st.selectbox("Sortir Berdasarkan:", ["Tampilkan Semua", "Pemasukan Saja", "Pengeluaran Saja"])
    
    if filter_jenis == "Pemasukan Saja":
        df_tampil = df[df['Jenis'] == 'Pemasukan'].copy()
    elif filter_jenis == "Pengeluaran Saja":
        df_tampil = df[df['Jenis'] == 'Pengeluaran'].copy()
    else:
        df_tampil = df.copy()
        
    if not df_tampil.empty:
        st.write("")
        st.dataframe(
            df_tampil.style.format({'Nominal': 'Rp {:,.0f}'}),
            use_container_width=True,
            hide_index=True,
            height=400
        )
        
        st.write("")
        file_excel = export_excel(df_tampil)
        st.download_button(
            label="📥 Unduh Laporan Resmi (.xlsx)",
            data=file_excel,
            file_name=f"Laporan_Kas_RT_{datetime.now().strftime('%d_%b_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.warning("⚠️ Tidak ada riwayat transaksi yang ditemukan untuk filter tersebut.")
