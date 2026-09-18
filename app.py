import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import numpy as np

# --- 1. Konfigurasi Halaman & Styling ---
st.set_page_config(page_title="BCP & SPS Project Auto-Planner", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .main-header { font-size: 28px; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-header { font-size: 15px; color: #4B5563; margin-bottom: 25px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Project Workspace: BCP & SPS Visit Auto-Planner</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Otomatisasi Distribusi Timeplan 3 Bulan & Tracking Biaya</div>', unsafe_allow_html=True)

# --- 2. Logika Auto-Planner 3 Bulan ---
def generate_3_months_plan(df):
    """
    Fungsi mendistribusikan jadwal pekerjaan site secara merata selama 90 Hari (3 Bulan)
    """
    total_sites = len(df)
    if total_sites == 0: return df
    
    start_date = pd.to_datetime('today').normalize()
    # Menghitung jeda hari antar pekerjaan agar merata selama 90 hari
    days_spacing = 90 / total_sites 
    
    plan_starts = []
    plan_ends = []
    
    for i in range(total_sites):
        current_start = start_date + timedelta(days=(i * days_spacing))
        # Asumsi durasi standar pengerjaan 1 site BCP/SPS adalah 2-3 Hari
        current_end = current_start + timedelta(days=2) 
        
        plan_starts.append(current_start)
        plan_ends.append(current_end)
        
    df['Plan Start'] = plan_starts
    df['Plan End'] = plan_ends
    return df

# --- 3. Data Loader & Generator ---
st.sidebar.header("📁 Unggah File Tracker")
uploaded_file = st.sidebar.file_uploader("Upload Sheet Data (Excel/CSV)", type=["xlsx", "csv"])

# Trigger Auto-Plan
auto_plan_toggle = st.sidebar.checkbox("Aktifkan Auto-Plan 3 Bulan", value=True, help="Sistem akan otomatis mengatur jadwal jika kolom Plan Start tidak ada")

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error membaca file: {e}")
        st.stop()
else:
    # Dummy Data Menggunakan format real jika file tidak ada
    st.sidebar.info("Gunakan format ini pada Excel Anda (menampilkan template data):")
    df = pd.DataFrame({
        'Site ID': ['PLK002', 'BNT001', 'KKN005', 'MTW012', 'PPS023', 'TML009', 'NTH080', 'KSN007', 'PRC008', 'PLK046'],
        'Tipe Pekerjaan': ['BCP', 'SPS Visit', 'BCP', 'SPS Visit', 'BCP', 'SPS Visit', 'BCP', 'BCP', 'SPS Visit', 'SPS Visit'],
        'Biaya Plan': [10000000, 12000000, 10000000, 15000000, 11000000, 12500000, 10000000, 10500000, 12000000, 15000000],
        'Tanggal Actual': [pd.NaT, '2026-09-15', pd.NaT, pd.NaT, pd.NaT, '2026-09-17', pd.NaT, pd.NaT, pd.NaT, pd.NaT],
        'Biaya Actual': [0, 11500000, 0, 0, 0, 12500000, 0, 0, 0, 0]
    })

# --- 4. Proses Eksekusi Data (Core Engine) ---
# Jalankan Auto Planner jika diaktifkan atau kolom Plan Start tidak ditemukan di excel
if auto_plan_toggle or 'Plan Start' not in df.columns:
    df = generate_3_months_plan(df)

# Standardisasi Tipe Data Waktu
df['Tanggal Actual'] = pd.to_datetime(df['Tanggal Actual'], errors='coerce')
df['Plan Start'] = pd.to_datetime(df['Plan Start'], errors='coerce')
df['Plan End'] = pd.to_datetime(df['Plan End'], errors='coerce')

# Otomasi Logika Status Pekerjaan
df['Status'] = df['Tanggal Actual'].apply(lambda x: 'Done' if pd.notnull(x) else 'Plan / On Progress')

# --- 5. Dashboard Panel KPI ---
col1, col2, col3, col4 = st.columns(4)
total_sites = len(df)
total_done = len(df[df['Status'] == 'Done'])
progress_pct = (total_done/total_sites)*100 if total_sites > 0 else 0

col1.metric("Total Site", total_sites)
col2.metric("Site Selesai (Done)", total_done, f"{progress_pct:.1f}%")
col3.metric("Estimasi Total Biaya (Plan)", f"Rp {df['Biaya Plan'].sum():,.0f}")
col4.metric("Realisasi Biaya (Actual)", f"Rp {df['Biaya Actual'].sum():,.0f}")
st.markdown("---")

# --- 6. Gantt Chart Timeline & Komparasi ---
tab1, tab2, tab3 = st.tabs(["📅 Gantt Chart Planner", "💰 Komparasi Biaya", "📋 Database Tracker"])

with tab1:
    fig_timeline = px.timeline(
        df, x_start="Plan Start", x_end="Plan End", y="Site ID", color="Status",
        hover_data=["Tipe Pekerjaan", "Tanggal Actual"],
        color_discrete_map={"Done": "#10B981", "Plan / On Progress": "#F59E0B"}
    )
    fig_timeline.update_yaxes(autorange="reversed") 
    fig_timeline.update_layout(height=500, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig_timeline, use_container_width=True)

with tab2:
    df_cost = df.melt(id_vars=['Site ID', 'Status'], value_vars=['Biaya Plan', 'Biaya Actual'], 
                      var_name='Kategori', value_name='Nilai (Rp)')
    fig_cost = px.bar(
        df_cost, x="Site ID", y="Nilai (Rp)", color="Kategori", barmode="group",
        color_discrete_map={"Biaya Plan": "#3B82F6", "Biaya Actual": "#6366F1"}
    )
    fig_cost.update_layout(height=500, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig_cost, use_container_width=True)

with tab3:
    df_display = df.copy()
    df_display['Plan Start'] = df_display['Plan Start'].dt.strftime('%d-%b-%Y')
    df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y')
    df_display['Tanggal Actual'] = df_display['Tanggal Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Dieksekusi')
    st.dataframe(df_display, use_container_width=True)
