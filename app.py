import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# ==========================================
# 1. SETUP HALAMAN & CUSTOM CSS (ENTERPRISE UI)
# ==========================================
st.set_page_config(page_title="BCP & SPS Executive Dashboard", layout="wide", page_icon="🏢")

st.markdown("""
<style>
    /* Global Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    /* Hero/Banner Header */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 35px 40px;
        border-radius: 12px;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .hero-title { font-size: 36px; font-weight: 800; margin-bottom: 5px; letter-spacing: -0.5px; }
    .hero-subtitle { font-size: 16px; font-weight: 400; color: #94a3b8; }
    
    /* Metric Cards Hover Effect */
    div[data-testid="metric-container"] {
        background-color: #ffffff; border: 1px solid #e2e8f0;
        padding: 20px; border-radius: 12px;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
        border-color: #3b82f6;
    }
    div[data-testid="metric-container"] label { font-weight: 600; color: #64748b; font-size: 14px; }
    div[data-testid="metric-container"] div { color: #0f172a; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

# Top Banner
st.markdown("""
<div class="hero-container">
    <div class="hero-title">Executive Dashboard: BCP & SPS Operation</div>
    <div class="hero-subtitle">Integrated Analytics, Cost Tracking & Automated 90-Day Scheduling</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. CORE ENGINE: MULTI-SHEET PROCESSOR
# ==========================================
@st.cache_data
def process_excel(file, auto_plan=True):
    try:
        xls = pd.ExcelFile(file)
        
        # --- PROSES SHEET 1: BCP ---
        df_bcp = pd.DataFrame()
        if 'BCP' in xls.sheet_names:
            df_bcp = pd.read_excel(xls, sheet_name='BCP', header=1) # Header BCP di baris ke-2
            df_bcp.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            # Mapping Kolom Standar
            bcp_map = {'Site ID': 'Site ID', 'City': 'Area/City', 'TE NAME': 'Engineer', 'Final': 'Kategori', 'Biaya Onsite': 'Biaya', 'Plan Date': 'Plan Date', 'Date Actual': 'Date Actual'}
            df_bcp = df_bcp.rename(columns={k: v for k, v in bcp_map.items() if k in df_bcp.columns})
            df_bcp['Source Sheet'] = 'BCP'
            
        # --- PROSES SHEET 2: SPS VISIT ---
        df_sps = pd.DataFrame()
        if 'SPS Visit' in xls.sheet_names:
            df_sps = pd.read_excel(xls, sheet_name='SPS Visit', header=0) # Header SPS di baris ke-1
            df_sps.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            # Mapping Kolom Standar
            sps_map = {'Site ID': 'Site ID', 'Kabupaten': 'Area/City', 'TE Name': 'Engineer', 'TOTAL': 'Biaya', 'Plan Visited': 'Plan Date'}
            df_sps = df_sps.rename(columns={k: v for k, v in sps_map.items() if k in df_sps.columns})
            df_sps['Source Sheet'] = 'SPS Visit'
            df_sps['Kategori'] = 'SPS Visit'
            df_sps['Date Actual'] = pd.NaT # Set kosong jika SPS belum punya kolom aktual

        # --- GABUNGKAN (MERGE) KEDUA SHEET ---
        df_master = pd.concat([df_bcp, df_sps], ignore_index=True)
        
        # Standarisasi Data
        if 'Area/City' in df_master.columns: df_master['Area/City'] = df_master['Area/City'].astype(str).str.upper()
        if 'Biaya' in df_master.columns:
            df_master['Biaya'] = df_master['Biaya'].astype(str).str.replace(r'\D', '', regex=True)
            df_master['Biaya'] = pd.to_numeric(df_master['Biaya'], errors='coerce').fillna(0)
            
        # Logika Status
        if 'Date Actual' in df_master.columns:
            df_master['Date Actual'] = pd.to_datetime(df_master['Date Actual'], errors='coerce')
        else:
            df_master['Date Actual'] = pd.NaT
        df_master['Status'] = df_master['Date Actual'].apply(lambda x: 'Selesai (Done)' if pd.notnull(x) else 'Pending / On-Plan')
        
        # Logika Penjadwalan Otomatis (90 Hari)
        df_master['Plan Date'] = pd.to_datetime(df_master['Plan Date'], errors='coerce')
        if auto_plan:
            mask_empty_plan = df_master['Plan Date'].isna()
            if mask_empty_plan.sum() > 0:
                start_date = pd.to_datetime('today').normalize()
                step = 90 / mask_empty_plan.sum()
                starts = [start_date + timedelta(days=int(i * step)) for i in range(mask_empty_plan.sum())]
                df_master.loc[mask_empty_plan, 'Plan Date'] = starts
                
        df_master['Plan End'] = df_master['Plan Date'] + pd.Timedelta(days=2)
        return df_master
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
        return None

def convert_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Master_Schedule')
    return output.getvalue()

# ==========================================
# 3. SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/7991/7991055.png", width=60) # Icon Dekorasi
    st.markdown("### Control Panel")
    uploaded_file = st.file_uploader("📂 Unggah File Tracker (Excel)", type=["xlsx"])
    auto_schedule = st.toggle("Aktifkan Auto-Planner 90 Hari", value=True)
    st.markdown("---")
    st.caption("Sistem ini otomatis mendeteksi, membersihkan, dan menggabungkan sheet 'BCP' dan 'SPS Visit'.")

# ==========================================
# 4. DASHBOARD RENDER
# ==========================================
if uploaded_file:
    df = process_excel(uploaded_file, auto_plan=auto_schedule)
    
    if df is not None:
        # --- KPI SECTION ---
        total_sites = len(df)
        total_done = len(df[df['Status'] == 'Selesai (Done)'])
        progress_pct = (total_done / total_sites) * 100 if total_sites > 0 else 0
        
        budget_total = df['Biaya'].sum()
        budget_spent = df.loc[df['Status'] == 'Selesai (Done)', 'Biaya'].sum()
        
        st.markdown(f"**Overall Project Progress:** `{progress_pct:.1f}%`")
        st.progress(int(progress_pct))
        st.write("")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📌 Total Site (BCP + SPS)", f"{total_sites} Sites")
        col2.metric("✅ Selesai Dikerjakan", f"{total_done} Sites", f"Sisa {total_sites - total_done} Pending")
        col3.metric("💰 Total Anggaran (RAB)", f"Rp {budget_total:,.0f}")
        col4.metric("📈 Biaya Terserap", f"Rp {budget_spent:,.0f}", f"Sisa Rp {budget_total - budget_spent:,.0f}", delta_color="off")
        
        st.write("<br>", unsafe_allow_html=True)
        
        # --- TAB NAVIGATION ---
        tab1, tab2, tab3 = st.tabs(["📊 Executive Timeline", "🗺️ Geo-Spatial & Budget", "📑 Master Database"])
        
        with tab1:
            st.markdown("#### Jadwal Distribusi Paralel (Gantt View)")
            df_plot = df.dropna(subset=['Plan Date']).sort_values('Plan Date')
            if not df_plot.empty:
                fig_gantt = px.timeline(
                    df_plot, x_start="Plan Date", x_end="Plan End", y="Site ID", color="Status",
                    hover_data=["Area/City", "Source Sheet", "Biaya"],
                    color_discrete_map={"Selesai (Done)": "#10b981", "Pending / On-Plan": "#f59e0b"},
                    template="plotly_white"
                )
                fig_gantt.update_yaxes(autorange="reversed", showgrid=True)
                fig_gantt.update_layout(height=650, margin=dict(t=10, b=20, l=10, r=10), font=dict(family="Inter"))
                st.plotly_chart(fig_gantt, use_container_width=True)
        
        with tab2:
            st.markdown("#### Analisis Sebaran Area & Alokasi Anggaran")
            c1, c2 = st.columns(2)
            
            with c1:
                # Stacked Bar Chart per Kota
                city_status = df.groupby(['Area/City', 'Status']).size().reset_index(name='Jumlah')
                fig_bar = px.bar(city_status, x="Area/City", y="Jumlah", color="Status", 
                                 title="Volume Pekerjaan Berdasarkan Area",
                                 color_discrete_map={"Selesai (Done)": "#10b981", "Pending / On-Plan": "#f59e0b"},
                                 template="plotly_white")
                fig_bar.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with c2:
                # Donut Chart Biaya
                city_budget = df.groupby('Area/City')['Biaya'].sum().reset_index()
                fig_pie = px.pie(city_budget, values='Biaya', names='Area/City', hole=0.5,
                                 title="Distribusi Anggaran (RAB) per Area",
                                 template="plotly_white")
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)
                
        with tab3:
            st.markdown("#### Master Tracker Database")
            # Styling Datagrid
            df_display = df.copy()
            df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d %b %Y').fillna('-')
            df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d %b %Y').fillna('-')
            
            cols_to_show = ['Site ID', 'Source Sheet', 'Area/City', 'Engineer', 'Status', 'Plan Date', 'Date Actual', 'Biaya']
            existing_cols = [c for c in cols_to_show if c in df_display.columns]
            
            st.dataframe(df_display[existing_cols], use_container_width=True, height=450)
            
            # Export Action
            excel_data = convert_to_excel(df)
            st.download_button(
                label="📥 Unduh Master Data ke Excel",
                data=excel_data,
                file_name=f"BCP_SPS_Master_Plan_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
else:
    # State Awal (Kosong)
    st.info("💡 **Petunjuk:** Silakan unggah file Excel Anda pada panel di sebelah kiri. Sistem akan langsung membaca sheet BCP beserta SPS dan menyajikannya dalam satu laporan manajerial.")
