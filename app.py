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
            current_start = start_date + timedelta(days=int(i * step_days))
            starts.append(current_start)
            
        df.loc[mask_needs_plan, 'Plan Date'] = starts
        df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2) 
        
    return df

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Jadwal_Terupdate')
    return output.getvalue()

# ==========================================
# 3. UPLOAD FILE & DETEKSI KOLOM PINTAR
# ==========================================
st.sidebar.header("📁 Konfigurasi & Upload")
uploaded_file = st.sidebar.file_uploader("Unggah File (Excel/CSV)", type=["xlsx", "csv"])
auto_schedule = st.sidebar.checkbox("Aktifkan Penjadwalan Otomatis (90 Hari)", value=True)

if uploaded_file:
    # 1. Membaca Data
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, sep=None, engine='python')
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    # 2. Algoritma Pencarian Kolom (Anti-KeyError)
    cols_lower = df.columns.str.lower().str.strip() # Normalisasi teks sementara
    
    col_actual = next((c for c, lower in zip(df.columns, cols_lower) if 'actual' in lower), None)
    col_plan = next((c for c, lower in zip(df.columns, cols_lower) if 'plan' in lower), None)
    col_biaya = next((c for c, lower in zip(df.columns, cols_lower) if 'biaya' in lower), None)
    col_site = next((c for c, lower in zip(df.columns, cols_lower) if 'site' in lower), None)
    col_city = next((c for c, lower in zip(df.columns, cols_lower) if 'city' in lower or 'kota' in lower), None)

    # Cek jika kolom vital benar-benar tidak ada di Excel
    if not col_actual or not col_plan:
        st.error(f"Sistem tidak dapat menemukan kolom Tanggal/Plan di excel Anda.")
        st.info(f"Kolom yang terdeteksi di file Anda: {', '.join(df.columns.tolist())}")
        st.stop()

    # 3. Rename ke nama standar untuk di-proses program
    df.rename(columns={
        col_actual: 'Date Actual',
        col_plan: 'Plan Date',
        col_biaya: 'Biaya Onsite',
        col_site: 'Site ID',
        col_city: 'City'
    }, inplace=True)

    # 4. Standardisasi Kolom Biaya
    if 'Biaya Onsite' in df.columns:
        df['Biaya Onsite'] = df['Biaya Onsite'].astype(str).str.replace(r'\D', '', regex=True)
        df['Biaya Onsite'] = pd.to_numeric(df['Biaya Onsite'], errors='coerce').fillna(0)
    else:
        df['Biaya Onsite'] = 0

    # 5. Konversi Waktu dan Labeling Status
    df['Date Actual'] = pd.to_datetime(df['Date Actual'], errors='coerce')
    df['Status Tracking'] = df['Date Actual'].apply(lambda x: 'Done' if pd.notnull(x) else 'Plan / Pending')
    df['Plan Date'] = pd.to_datetime(df['Plan Date'], errors='coerce')
    df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2) 

    # 6. Jalankan Logika Auto-Plan 3 Bulan
    if auto_schedule:
        df = apply_3_month_schedule(df)

    # ==========================================
    # 4. TAMPILAN DASHBOARD UTAMA
    # ==========================================
    total_sites = len(df)
    total_done = len(df[df['Status Tracking'] == 'Done'])
    progress = (total_done/total_sites)*100 if total_sites > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Site Target", total_sites)
    col2.metric("Site Selesai (Done)", total_done, f"{progress:.1f}% Progress")
    col3.metric("Estimasi Biaya Total", f"Rp {df['Biaya Onsite'].sum():,.0f}")
    
    realisasi = df.loc[df['Status Tracking'] == 'Done', 'Biaya Onsite'].sum()
    col4.metric("Biaya Terserap (Done)", f"Rp {realisasi:,.0f}")

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs(["📅 Gantt Chart Timeline", "🗺️ Sebaran Area", "📋 Database & Export"])

    with tab1:
        st.subheader("Distribusi SLA Eksekusi (3 Bulan)")
        df_plot = df.dropna(subset=['Plan Date']).copy()
        if not df_plot.empty:
            fig_timeline = px.timeline(
                df_plot, x_start="Plan Date", x_end="Plan End", y="Site ID", color="Status Tracking",
                color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"}
            )
            fig_timeline.update_yaxes(autorange="reversed")
            fig_timeline.update_layout(height=500)
            st.plotly_chart(fig_timeline, use_container_width=True)
        else:
            st.warning("Data timeline belum tersedia.")

    with tab2:
        st.subheader("Sebaran Pekerjaan Berdasarkan Area")
        if 'City' in df.columns:
            city_count = df.groupby(['City', 'Status Tracking']).size().reset_index(name='Jumlah Site')
            fig_city = px.bar(
                city_count, x="City", y="Jumlah Site", color="Status Tracking", 
                barmode="group", color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"}
            )
            st.plotly_chart(fig_city, use_container_width=True)
        else:
            st.info("Kolom Area/City tidak ditemukan pada file.")

    with tab3:
        st.subheader("Database Tabel")
        df_display = df.copy()
        df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Selesai')
        df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        
        st.dataframe(df_display, use_container_width=True, height=350)
        
        st.download_button(
            label="📥 Download Jadwal 3 Bulan ke Excel",
            data=to_excel(df),
            file_name="Jadwal_BCP_SPS_Terupdate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Silakan unggah file Excel Anda pada menu samping kiri untuk memulai.")
