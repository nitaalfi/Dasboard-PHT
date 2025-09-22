import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt

# --- Konfigurasi Halaman ---
st.set_page_config(
    page_title="Dashboard Monitoring Aset Perhutani",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Tambahkan Logo + Judul (Perbaikan: Optional logo) ---
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image("logo.webp", width=100)  # Jika file tidak ada, skip tanpa error
    except FileNotFoundError:
        st.write("📊")  # Placeholder jika logo hilang
with col2:
    st.markdown(
        "<h1 style='color:#1b5e20; margin-bottom:0;'>Dashboard Monitoring Aset Perhutani</h1>",
        unsafe_allow_html=True
    )
    st.caption("Visualisasi dan Analisis Data Aset")

# --- Sidebar ---
st.sidebar.header("📂 Menu Utama")
uploaded_file = st.sidebar.file_uploader("Upload file Excel data aset", type=["xlsx"])

if uploaded_file:
    try:
        # Membaca data Excel
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Membersihkan nama kolom (strip spasi dan ubah ke huruf kapital untuk konsistensi)
        df.columns = df.columns.str.strip().str.title()

        # Hapus semua baris yang seluruh kolomnya kosong
        df = df.dropna(how='all')

        # Cari baris header yang benar (Perbaikan: Lebih aman)
        temp_df = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
        header_candidates = temp_df[temp_df.apply(lambda r: r.astype(str).str.contains("No. Urut", case=False).any(), axis=1)]
        if not header_candidates.empty:
            header_row = header_candidates.index[0]
        else:
            st.sidebar.warning("Header 'No. Urut' tidak ditemukan. Menggunakan baris pertama sebagai header.")
            header_row = 0  # Default ke baris pertama

        # Baca ulang dengan header yang tepat
        df = pd.read_excel(uploaded_file, engine="openpyxl", skiprows=header_row)

        # Bersihkan kolom
        df.columns = df.columns.str.strip().str.title()

        # Menampilkan kolom yang tersedia (untuk debugging)
        st.sidebar.info(f"Kolom yang terdeteksi: {', '.join(df.columns.tolist())}")
        
        # Sidebar - Monitoring
        st.sidebar.subheader("Monitoring Filter")
        
        # Filter KPH - handle berbagai kemungkinan nama kolom
        kph_col = None
        for col in ['Nama Satker', 'Nama Satker*', 'Kph', 'KPH', 'kph', 'Kesatuan Pengelolaan Hutan']:
            if col in df.columns:
                kph_col = col
                break

        if kph_col:
            kph_list = sorted(df[kph_col].dropna().astype(str).unique())
            selected_kph = st.sidebar.multiselect("Pilih KPH", options=kph_list, default=kph_list)
        else:
            st.sidebar.error("Kolom KPH tidak ditemukan di data.")
            selected_kph = []
        
        # Filter Kondisi - handle berbagai kemungkinan nama kolom
        kondisi_col = None
        for col in ['Kondisi', 'Kondisi Aset', 'Kondisi Aset*', 'Keadaan', 'Condition']:
            if col in df.columns:
                kondisi_col = col
                break
        
        if kondisi_col:
            kondisi_list = sorted(df[kondisi_col].dropna().astype(str).unique())
            selected_kondisi = st.sidebar.multiselect(f"Pilih {kondisi_col}", options=kondisi_list, default=kondisi_list)
        else:
            selected_kondisi = []

        # Filter Jenis Aset
        jenis_col = None
        for col in ['Jenis Aset', 'Jenis', 'Kategori', 'Tipe', 'Type', 'Klasifikasi']:
            if col in df.columns:
                jenis_col = col
                break
        
        if jenis_col:
            jenis_list = sorted(df[jenis_col].dropna().astype(str).unique())
            selected_jenis = st.sidebar.multiselect(f"Pilih {jenis_col}", options=jenis_list, default=jenis_list)
        else:
            selected_jenis = []
        
        # Filter Tanggal Perolehan (Perbaikan: Try-except untuk parsing tanggal)
        tgl_col = None
        for col in ['Tanggal Perolehan', 'Tanggal', 'Tanggal*', 'Date', 'Tanggal Pembelian']:
            if col in df.columns:
                tgl_col = col
                break

        selected_tahun = []
        if tgl_col:
            try:
                # Ekstrak tahun dari kolom tanggal
                df['Tahun'] = pd.to_datetime(df[tgl_col], errors='coerce').dt.year
                tahun_non_null = df['Tahun'].dropna().astype(int).unique()
                tahun_list = sorted(tahun_non_null)
                if tahun_list:
                    selected_tahun = st.sidebar.multiselect("Pilih Tahun Perolehan", options=tahun_list, default=tahun_list)
                else:
                    st.sidebar.warning("Tidak ada tahun valid di kolom tanggal (semua tanggal invalid).")
            except Exception as e:
                st.sidebar.error(f"Error parsing tanggal: {str(e)}. Filter tahun diabaikan.")
                df['Tahun'] = np.nan  # Fallback
        else:
            st.sidebar.info("Kolom tanggal tidak ditemukan.")
        
        tahun_col = 'Tahun'
        
        # Filter data berdasarkan pilihan
        filtered_df = df.copy()
        
        if selected_kph and kph_col:
            filtered_df = filtered_df[filtered_df[kph_col].astype(str).isin(selected_kph)]
        
        if kondisi_col and selected_kondisi:
            filtered_df = filtered_df[filtered_df[kondisi_col].astype(str).isin(selected_kondisi)]
        
        if tgl_col and selected_tahun and not filtered_df[tahun_col].isna().all():  # Perbaikan: Cek tidak semua NaN
            valid_tahun = filtered_df[tahun_col].dropna().astype(int)
            filtered_df = filtered_df[valid_tahun.isin(selected_tahun)]
        
        if jenis_col and selected_jenis:
            filtered_df = filtered_df[filtered_df[jenis_col].astype(str).isin(selected_jenis)]
        
        # Menangani kolom nilai aset dengan berbagai nama
        nilai_col = None
        for col in ['Nilai Aset', 'Nilai Aset*', 'Nilai', 'Harga', 'Value', 'Nilai Perolehan*', 'Harga Perolehan']:
            if col in df.columns:
                nilai_col = col
                break
        
        # Konversi kolom nilai ke numeric jika ada
        if nilai_col:
            filtered_df[nilai_col] = pd.to_numeric(filtered_df[nilai_col], errors='coerce')
        
        # Sidebar - Export data filtered (Perbaikan: Gunakan openpyxl)
        def to_excel(df_export):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:  # Ganti ke openpyxl
                df_export.to_excel(writer, index=False, sheet_name='Data Filtered')
            processed_data = output.getvalue()
            return processed_data
        
        st.sidebar.subheader("Export Data")
        if not filtered_df.empty:
            excel_data = to_excel(filtered_df)
            st.sidebar.download_button(
                label="Download Data Filtered ke Excel",
                data=excel_data,
                file_name='data_aset_filtered.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.sidebar.info("Data filter kosong, tidak bisa diexport.")
        
        # --- Main Page (Sama seperti sebelumnya, tapi dengan perbaikan minor di metric dan grafik) ---
        st.header("Monitoring Data Aset")
        
        # Ringkasan Statistik
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Aset", len(filtered_df))
        
        with col2:
            if nilai_col:
                total_nilai = filtered_df[nilai_col].sum()
                st.metric("Total Nilai Aset", f"Rp {total_nilai:,.0f}" if not pd.isna(total_nilai) else "Rp 0")
            else:
                st.metric("Total Nilai Aset", "Kolom tidak ditemukan")
                
        with col3:
            if kondisi_col:
                # Hitung aset dengan kondisi baik (case insensitive, abaikan NaN)
                kondisi_clean = filtered_df[kondisi_col].astype(str).str.lower()
                mask_baik = ~kondisi_clean.isna() & kondisi_clean.str.contains('baik|bagus|good|excellent|perfect', na=False)
                kondisi_baik = mask_baik.sum()
                st.metric("Aset Kondisi Baik", kondisi_baik)
            else:
                st.metric("Aset Kondisi Baik", "Kolom tidak ditemukan")
                
        with col4:
            # Hitung data tidak lengkap (minimal 2 kolom penting kosong)
            kolom_penting = [kph_col, nilai_col, jenis_col, kondisi_col]
            kolom_penting = [col for col in kolom_penting if col is not None and col in filtered_df.columns]
            if kolom_penting:
                jml_tidak_lengkap = (filtered_df[kolom_penting].isnull().sum(axis=1) >= 2).sum()
                st.metric("Aset Data Tidak Lengkap", jml_tidak_lengkap)
            else:
                st.metric("Aset Data Tidak Lengkap", "Data tidak cukup")
        
        st.subheader("Data Aset (Setelah Filter)")
        st.dataframe(filtered_df, height=400)
        
        # Statistik jumlah aset per jenis aset
        if jenis_col and not filtered_df.empty:
            st.subheader("Jumlah Aset per Jenis Aset")
            jumlah_per_jenis = filtered_df[jenis_col].value_counts()
            
            fig, ax = plt.subplots(figsize=(10, 6))
            jumlah_per_jenis.plot(kind='bar', ax=ax, color='skyblue')
            ax.set_title('Distribusi Aset per Jenis')
            ax.set_ylabel('Jumlah Aset')
            ax.tick_params(axis='x', rotation=45)
            st.pyplot(fig)
            plt.close(fig)  # Per fig untuk aman
        else:
            st.info("Kolom jenis
