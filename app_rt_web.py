import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

# ==========================================
# KONFIGURASI & FUNGSI PENDUKUNG
# ==========================================
st.set_page_config(page_title="Web Admin Kas RT", page_icon="🏘️", layout="wide")

FILE_DATA = 'database_kas_rt.csv'

# Daftar Blok/Warga untuk dropdown (Bisa disesuaikan dengan lingkungan)
DAFTAR_BLOK = ["Blok A1", "Blok A2", "Blok B1", "Blok B2", "Fasum", "Lainnya"]

def muat_data():
    if not os.path.exists(FILE_DATA):
        df = pd.DataFrame(columns=['ID', 'Tanggal', 'Jenis', 'Kategori', 'Keterangan', 'Blok_Warga', 'Nominal'])
        df.to_csv(FILE_DATA, index=False)
        return df
    return pd.read_csv(FILE_DATA)

def simpan_data(df):
    df.to_csv(FILE_DATA, index=False)

def export_excel(df):
    output = io.BytesIO()
    # Menggunakan openpyxl sebagai engine untuk export Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laporan_Kas')
    return output.getvalue()

# Load Data
df = muat_data()

st.title("🏘️ Sistem Informasi Kas & Iuran RT")
st.markdown("Aplikasi web untuk mencatat administrasi keuangan dan iuran warga kawasan Babelan dan sekitarnya.")

# ==========================================
# NAVIGASI TAB
# ==========================================
tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📝 Input Transaksi", "📁 Laporan & Export"])

# ------------------------------------------
# TAB 1: DASHBOARD
# ------------------------------------------
with tab1:
    st.header("Ringkasan Keuangan")
    
    if not df.empty:
        total_masuk = df[df['Jenis'] == 'Pemasukan']['Nominal'].sum()
        total_keluar = df[df['Jenis'] == 'Pengeluaran']['Nominal'].sum()
        saldo = total_masuk - total_keluar
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Pemasukan", f"Rp {total_masuk:,.0f}")
        col2.metric("Total Pengeluaran", f"Rp {total_keluar:,.0f}")
        col3.metric("Saldo Kas Tersedia", f"Rp {saldo:,.0f}")
        
        st.divider()
        
        # Grafik sederhana menggunakan chart bawaan Streamlit
        st.subheader("Grafik Arus Kas")
        
        # PERBAIKAN: Menangani error format tanggal & menghapus baris yang gagal dikonversi
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        df_grafik = df.dropna(subset=['Tanggal']).copy()
        
        if not df_grafik.empty:
            df_group = df_grafik.groupby(['Tanggal', 'Jenis'])['Nominal'].sum().unstack().fillna(0)
            st.bar_chart(df_group)
        else:
            st.info("Belum ada data tanggal yang valid untuk ditampilkan di grafik.")
    else:
        st.info("Belum ada data transaksi yang tercatat. Silakan input data terlebih dahulu.")

# ------------------------------------------
# TAB 2: INPUT TRANSAKSI
# ------------------------------------------
with tab2:
    st.header("Catat Transaksi Baru")
    
    with st.form("form_transaksi", clear_on_submit=True):
        col_a, col_b = st.columns(2)
        
        with col_a:
            jenis_trx = st.radio("Jenis Transaksi", ["Pemasukan", "Pengeluaran"], horizontal=True)
            tanggal = st.date_input("Tanggal", datetime.today())
            
        with col_b:
            nominal = st.number_input("Nominal (Rp)", min_value=0, step=10000)
            
        st.markdown("**Detail Keterangan**")
        kategori = st.selectbox("Kategori", ["Iuran Bulanan", "Sumbangan", "Kebersihan", "Keamanan", "Perbaikan Fasilitas", "Lain-lain"])
        blok = st.selectbox("Blok / Warga Terkait", DAFTAR_BLOK)
        keterangan = st.text_input("Keterangan Tambahan", placeholder="Contoh: Iuran bulan September")
        
        submit_btn = st.form_submit_button("Simpan Transaksi")
        
        if submit_btn:
            if nominal <= 0:
                st.error("Nominal tidak boleh kosong (Rp 0)!")
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
                st.success("✅ Data berhasil disimpan!")
                st.rerun()

# ------------------------------------------
# TAB 3: LAPORAN & EXPORT
# ------------------------------------------
with tab3:
    st.header("Rekapitulasi Data")
    
    # Filter Data
    filter_jenis = st.selectbox("Filter berdasarkan Jenis:", ["Semua", "Pemasukan", "Pengeluaran"])
    
    if filter_jenis != "Semua":
        df_tampil = df[df['Jenis'] == filter_jenis].copy()
    else:
        df_tampil = df.copy()
        
    if not df_tampil.empty:
        # Menampilkan tabel interaktif
        st.dataframe(
            df_tampil.style.format({'Nominal': 'Rp {:,.0f}'}),
            use_container_width=True,
            hide_index=True
        )
        
        # Tombol Download Excel menggunakan Openpyxl
        file_excel = export_excel(df_tampil)
        st.download_button(
            label="📥 Download Laporan (Excel)",
            data=file_excel,
            file_name=f"Laporan_Kas_RT_{datetime.now().strftime('%Y%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Tidak ada data yang sesuai dengan filter pencarian.")
