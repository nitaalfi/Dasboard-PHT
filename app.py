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
        # Membaca data Excel (Opsional: Cache untuk performa)
        # @st.cache_data
        # def load_data(file):
        #     return pd.read_excel(file, engine='openpyxl')
        # df = load_data(uploaded_file)
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

        # Hapus baris kosong lagi setelah skiprows
        df = df.dropna(how='all')

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
                tahun_non_null = df['Tahun'].dropna()
                if not tahun_non_null.empty:
                    tahun_list = sorted(tahun_non_null.astype(int).unique())
                    selected_tahun = st.sidebar.multiselect("Pilih Tahun Perolehan", options=tahun_list, default=tahun_list)
                else:
                    st.sidebar.warning("Tidak ada tahun valid di kolom tanggal (semua tanggal invalid).")
            except Exception as e:
                st.sidebar.error(f"Error parsing tanggal: {str(e)}. Filter tahun diabaikan.")
                df['Tahun'] = np.nan  # Fallback
        else:
            st.sidebar.info("Kolom tanggal tidak ditemukan.")
        
        tahun_col = 'Tahun'
        
        # Filter data berdasarkan pilihan (Perbaikan: Handle NaN di tahun)
        filtered_df = df.copy()
        
        if selected_kph and kph_col:
            filtered_df = filtered_df[filtered_df[kph_col].astype(str).isin(selected_kph)]
        
        if kondisi_col and selected_kondisi:
            filtered_df = filtered_df[filtered_df[kondisi_col].astype(str).isin(selected_kondisi)]
        
        if tgl_col and selected_tahun:
            # Handle NaN: Fill dengan -1 sementara untuk filter, lalu drop kolom
            filtered_df[tahun_col] = filtered_df[tahun_col].fillna(-1)
            valid_mask = filtered_df[tahun_col].astype(int).isin(selected_tahun)
            filtered_df = filtered_df[valid_mask]
            # Drop kolom Tahun sementara jika tidak diperlukan
            if tahun_col not in df.columns:
                filtered_df = filtered_df.drop(columns=[tahun_col])
        
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
        
        # Sidebar - Export data filtered (Perbaikan: Try-except dan openpyxl)
        def to_excel(df_export):
            try:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Data Filtered')
                processed_data = output.getvalue()
                return processed_data
            except Exception as e:
                st.sidebar.error(f"Error export: {str(e)}")
                return None
        
        st.sidebar.subheader("Export Data")
        if not filtered_df.empty:
            excel_data = to_excel(filtered_df)
            if excel_data:
                st.sidebar.download_button(
                    label="Download Data Filtered ke Excel",
                    data=excel_data,
                    file_name='data_aset_filtered.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.sidebar.info("Data filter kosong, tidak bisa diexport.")
        
        # --- Main Page ---
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
            plt.close(fig)  # Bersihkan memori per figure
        else:
            st.info("Kolom jenis aset tidak ditemukan atau data kosong.")
        
        # Jumlah aset bermasalah
        st.subheader("Analisis Kondisi Aset")
        
        if kondisi_col and not filtered_df.empty:
            col1, col2 = st.columns(2)
            
            with col1:
                # Hitung aset bermasalah (tidak baik, termasuk NaN)
                kondisi_clean = filtered_df[kondisi_col].astype(str).str.lower()
                mask_bermasalah = ~kondisi_clean.str.contains('baik|bagus|good|excellent|perfect', na=False)
                jml_bermasalah = mask_bermasalah.sum()
                st.metric("Jumlah Aset Bermasalah", jml_bermasalah)
            
            with col2:
                # Distribusi kondisi
                distribusi_kondisi = filtered_df[kondisi_col].value_counts()
                fig, ax = plt.subplots(figsize=(8, 6))
                distribusi_kondisi.plot(kind='pie', autopct='%1.1f%%', ax=ax)
                ax.set_title('Distribusi Kondisi Aset')
                ax.set_ylabel('')
                st.pyplot(fig)
                plt.close(fig)
        else:
            st.info("Kolom kondisi tidak ditemukan atau data kosong.")
        
        # Analisis Nilai Aset
        st.subheader("Analisis Nilai Aset")
        
        if nilai_col and not filtered_df.empty:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                total_nilai = filtered_df[nilai_col].sum()
                st.metric("Total Nilai Perolehan", f"Rp {total_nilai:,.0f}" if not pd.isna(total_nilai) else "Rp 0")
            
            with col2:
                rata_nilai = filtered_df[nilai_col].mean()
                st.metric("Rata-rata Nilai", f"Rp {rata_nilai:,.0f}" if not pd.isna(rata_nilai) else "Rp 0")
            
            with col3:
                median_nilai = filtered_df[nilai_col].median()
                st.metric("Median Nilai", f"Rp {median_nilai:,.0f}" if not pd.isna(median_nilai) else "Rp 0")
            
            # Nilai tertinggi dan terendah (Perbaikan: Cek data tidak kosong)
            nilai_clean = filtered_df[nilai_col].dropna()
            if not nilai_clean.empty:
                max_nilai = nilai_clean.max()
                min_nilai = nilai_clean.min()
                
                aset_max = filtered_df.loc[filtered_df[nilai_col].idxmax()]  # Lebih aman daripada iloc[0]
                aset_min = filtered_df.loc[filtered_df[nilai_col].idxmin()]
                
                st.markdown("**Aset dengan Nilai Tertinggi:**")
                st.write(f"Nilai: Rp {max_nilai:,.0f}")
                if kph_col and kph_col in aset_max:
                    st.write(f"KPH: {aset_max[kph_col]}")
                if jenis_col and jenis_col in aset_max:
                    st.write(f"Jenis: {aset_max[jenis_col]}")
                
                st.markdown("**Aset dengan Nilai Terendah:**")
                st.write(f"Nilai: Rp {min_nilai:,.0f}")
                if kph_col and kph_col in aset_min:
                    st.write(f"KPH: {aset_min[kph_col]}")
                if jenis_col and jenis_col in aset_min:
                    st.write(f"Jenis: {aset_min[jenis_col]}")
            else:
                st.warning("Tidak ada data nilai aset yang valid.")
        else:
            st.info("Kolom nilai aset tidak ditemukan atau data kosong.")
        
        # Grafik dan Statistik Lanjutan
        st.header("Statistik dan Grafik Lanjutan")
        
        if nilai_col and not filtered_df.empty:
            nilai_clean = filtered_df[nilai_col].dropna()
            if not nilai_clean
