import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta
import io

# ==========================================
# 1. SETUP HALAMAN & KONFIGURASI DASHBOARD
# ==========================================
st.set_page_config(page_title="BCP & SPS Visit Planner", layout="wide", page_icon="⚙️")

st.markdown("""
<style>
    .main-title { font-size: 32px; font-weight: 800; color: #1E3A8A; margin-bottom: -5px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 20px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Project Tracker: BCP & SPS Visit Auto-Planner</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Otomatisasi Jadwal 3 Bulan & Tracking Biaya Aktual</div>', unsafe_allow_html=True)

# ==========================================
# 2. ENGINE PENJADWALAN OTOMATIS (90 HARI)
# ==========================================
def apply_3_month_schedule(df):
    """Mendistribusikan site kosong ke rentang 90 hari"""
    start_date = pd.to_datetime('today').normalize()
    
    # Deteksi mana yang Plan Date-nya kosong ('NY' atau NaN)
    mask_needs_plan = df['Plan Date'].isna() | (df['Plan Date'] == 'NY')
    total_sites_to_plan = mask_needs_plan.sum()
    
    if total_sites_to_plan > 0:
        step_days = 90 / total_sites_to_plan
        starts = []
        for i in range(total_sites_to_plan):
            # Hitung hari
            current_start = start_date + timedelta(days=int(i * step_days))
            starts.append(current_start)
            
        # Isi ke dataframe (Start dan Asumsi End + 2 Hari)
        df.loc[mask_needs_plan, 'Plan Date'] = starts
        df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2) 
        
    return df

# Helper Function: Download Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Jadwal_Terupdate')
    return output.getvalue()

# ==========================================
# 3. UPLOAD FILE 
# ==========================================
st.sidebar.header("📁 Konfigurasi & Upload")
uploaded_file = st.sidebar.file_uploader("Unggah File (BCP & SPS visit.xlsx)", type=["xlsx", "csv"])
auto_schedule = st.sidebar.checkbox("Aktifkan Penjadwalan Otomatis (90 Hari)", value=True)

if uploaded_file:
    # Membaca data jika file diupload
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    # Standardisasi Nama Kolom (Karena dari snippet ada spasi ' Biaya Onsite ')
    df.rename(columns=lambda x: x.strip(), inplace=True)
    
    # Pastikan Kolom Biaya adalah Numerik (Bersihkan Rp/titik jika ada)
    if 'Biaya Onsite' in df.columns:
        df['Biaya Onsite'] = df['Biaya Onsite'].astype(str).str.replace(r'\D', '', regex=True)
        df['Biaya Onsite'] = pd.to_numeric(df['Biaya Onsite'], errors='coerce').fillna(0)
    else:
        df['Biaya Onsite'] = 0

    # 1. Bersihkan Kolom Tanggal Aktual
    df['Date Actual'] = pd.to_datetime(df['Date Actual'], errors='coerce')

    # 2. Set Status "Done" Jika Date Actual Ada Isinya
    df['Status Tracking'] = df['Date Actual'].apply(lambda x: 'Done' if pd.notnull(x) else 'Plan / Pending')
    
    # 3. Konversi format Plan Date sebelum Auto Plan
    df['Plan Date'] = pd.to_datetime(df['Plan Date'], errors='coerce')
    
    # 4. Buat Plan End awal dari Plan Date + 2 hari
    df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2) 

    # 5. Jalankan Logic Auto-Plan
    if auto_schedule:
        df = apply_3_month_schedule(df)
        
    # ==========================================
    # 4. HEADER KPI DASHBOARD
    # ==========================================
    total_sites = len(df)
    total_done = len(df[df['Status Tracking'] == 'Done'])
    progress = (total_done/total_sites)*100 if total_sites > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Site Target", total_sites)
    col2.metric("Site Selesai (Done)", total_done, f"{progress:.1f}% Progress")
    col3.metric("Estimasi Biaya Total", f"Rp {df['Biaya Onsite'].sum():,.0f}")
    
    # Hitung Realisasi Berdasarkan Site yang DONE saja
    realisasi = df.loc[df['Status Tracking'] == 'Done', 'Biaya Onsite'].sum()
    col4.metric("Biaya Terserap (Done)", f"Rp {realisasi:,.0f}")

    st.markdown("---")

    # ==========================================
    # 5. PANEL VISUALISASI
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📅 Gantt Chart Timeline", "🗺️ Sebaran Area", "📋 Database & Export"])

    with tab1:
        st.subheader("Distribusi SLA Eksekusi (3 Bulan)")
        # Plotly tidak bisa plot NaT, drop sementara untuk visual
        df_plot = df.dropna(subset=['Plan Date']).copy()
        
        if not df_plot.empty:
            # Gunakan kolom 'Final' sebagai tipe pekerjaan (misal: BCP/SPS)
            fig_timeline = px.timeline(
                df_plot, x_start="Plan Date", x_end="Plan End", y="Site ID", color="Status Tracking",
                hover_data=["City", "TE NAME", "Final"], 
                color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"},
            )
            fig_timeline.update_yaxes(autorange="reversed")
            fig_timeline.update_layout(height=500)
            st.plotly_chart(fig_timeline, use_container_width=True)
        else:
            st.warning("Tidak ada data jadwal valid untuk divisualisasikan.")

    with tab2:
        st.subheader("Sebaran Pekerjaan Berdasarkan Kabupaten/Kota")
        if 'City' in df.columns:
            city_count = df.groupby(['City', 'Status Tracking']).size().reset_index(name='Jumlah Site')
            fig_city = px.bar(
                city_count, x="City", y="Jumlah Site", color="Status Tracking", 
                barmode="group", color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"}
            )
            st.plotly_chart(fig_city, use_container_width=True)
        else:
            st.info("Kolom 'City' tidak ditemukan dalam tabel.")

    with tab3:
        st.subheader("Database Tabel")
        
        # Display copy agar format datetime enak dibaca
        df_display = df.copy()
        df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Selesai')
        df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        
        st.dataframe(df_display, use_container_width=True, height=350)
        
        # Tombol Unduh Otomatis Ke Excel
        excel_file = to_excel(df)
        st.download_button(
            label="📥 Download Jadwal 3 Bulan ke Excel",
            data=excel_file,
            file_name="Jadwal_BCP_SPS_Terupdate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Silakan unggah file 'BCP & SPS visit.xlsx' Anda pada menu samping kiri untuk memulai.")
