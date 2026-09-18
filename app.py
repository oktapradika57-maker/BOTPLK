import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# ==========================================
# 1. KONFIGURASI HALAMAN & DESIGN SYSTEM (UI)
# ==========================================
st.set_page_config(
    page_title="Executive Operations Dashboard | BCP & SPS",
    layout="wide",
    page_icon="📊"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Header Utama */
    .main-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 30px 35px;
        border-radius: 14px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
    }
    .header-title { font-size: 28px; font-weight: 700; margin: 0; letter-spacing: -0.5px; }
    .header-subtitle { font-size: 14px; color: #94a3b8; margin-top: 5px; }
    
    /* Styling Kartu Metrik */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 18px 20px;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    div[data-testid="metric-container"] label { font-size: 13px; font-weight: 600; color: #64748b; }
    div[data-testid="metric-container"] div { font-size: 22px; font-weight: 700; color: #0f172a; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. PROSES & PEMBERSIHAN DATA (MULTI-SHEET)
# ==========================================
@st.cache_data
def load_and_clean_data(file, auto_plan=True):
    try:
        xls = pd.ExcelFile(file)
        
        # --- Sheet BCP ---
        df_bcp = pd.DataFrame()
        if 'BCP' in xls.sheet_names:
            df_bcp = pd.read_excel(xls, sheet_name='BCP', header=1)
            df_bcp.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            # Normalisasi kolom BCP
            bcp_rename = {
                'Site ID': 'Site ID', 'City': 'Area/City', 'TE NAME': 'PIC Engineer', 
                'Final': 'Kategori', 'Biaya Onsite': 'Biaya', 'Plan Date': 'Plan Date', 'Date Actual': 'Date Actual'
            }
            df_bcp = df_bcp.rename(columns={k: v for k, v in bcp_rename.items() if k in df_bcp.columns})
            df_bcp['Source Program'] = 'BCP Visit'
            
        # --- Sheet SPS Visit ---
        df_sps = pd.DataFrame()
        if 'SPS Visit' in xls.sheet_names:
            df_sps = pd.read_excel(xls, sheet_name='SPS Visit', header=0)
            df_sps.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            # Normalisasi kolom SPS
            sps_rename = {
                'Site ID': 'Site ID', 'Kabupaten': 'Area/City', 'TE Name': 'PIC Engineer', 
                'TOTAL': 'Biaya', 'Plan Visited': 'Plan Date'
            }
            df_sps = df_sps.rename(columns={k: v for k, v in sps_rename.items() if k in df_sps.columns})
            df_sps['Source Program'] = 'SPS Visit'
            df_sps['Kategori'] = 'SPS Visit'
            df_sps['Date Actual'] = pd.NaT

        # Gabungkan Data
        df_master = pd.concat([df_bcp, df_sps], ignore_index=True)
        
        # Bersihkan & Standarisasi Biaya (Pastikan murni angka jutaan/ratusan ribu, hindari lonjakan desimal/puluhan juta)
        if 'Biaya' in df_master.columns:
            df_master['Biaya'] = df_master['Biaya'].astype(str).str.replace(r'\D', '', regex=True)
            df_master['Biaya'] = pd.to_numeric(df_master['Biaya'], errors='coerce').fillna(0)
            
        # Standarisasi Teks Area & PIC
        if 'Area/City' in df_master.columns:
            df_master['Area/City'] = df_master['Area/City'].astype(str).str.upper().str.strip()
        if 'PIC Engineer' in df_master.columns:
            df_master['PIC Engineer'] = df_master['PIC Engineer'].astype(str).str.title().str.strip()
            df_master['PIC Engineer'] = df_master['PIC Engineer'].replace(['Nan', ''], 'Unassigned')

        # Status Pekerjaan Berdasarkan Tanggal Actual
        if 'Date Actual' in df_master.columns:
            df_master['Date Actual'] = pd.to_datetime(df_master['Date Actual'], errors='coerce')
        else:
            df_master['Date Actual'] = pd.NaT
            
        df_master['Status Progress'] = df_master['Date Actual'].apply(lambda x: 'Done (Selesai)' if pd.notnull(x) else 'Pending (On-Plan)')
        
        # Penjadwalan Otomatis (90 Hari / 3 Bulan) untuk Site yang kosong
        df_master['Plan Date'] = pd.to_datetime(df_master['Plan Date'], errors='coerce')
        if auto_plan:
            mask_empty = df_master['Plan Date'].isna()
            if mask_empty.sum() > 0:
                start_date = pd.to_datetime('today').normalize()
                step = 90 / mask_empty.sum()
                starts = [start_date + timedelta(days=int(i * step)) for i in range(mask_empty.sum())]
                df_master.loc[mask_empty, 'Plan Date'] = starts
                
        df_master['Plan End'] = df_master['Plan Date'] + pd.Timedelta(days=2)
        return df_master
        
    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")
        return None

def export_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Master_Rekap_Project')
    return output.getvalue()

# ==========================================
# 3. SIDEBAR NAVIGATION & UPLOAD
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=50)
    st.markdown("### **Panel Kontrol Data**")
    uploaded_file = st.file_uploader("Unggah File Tracker (.xlsx)", type=["xlsx"])
    auto_schedule = st.toggle("Aktifkan Auto-Scheduler (3 Bulan)", value=True)
    st.markdown("---")
    st.info("Sistem membaca lembar kerja **BCP** dan **SPS Visit**, merapikan nominal anggaran, serta memetakan beban kerja per PIC secara otomatis.")

# ==========================================
# 4. DASHBOARD UTAMA
# ==========================================
st.markdown("""
<div class="main-header">
    <div class="header-title">Executive Project Monitoring & Analytics</div>
    <div class="header-subtitle">Integrasi BCP & SPS Visit Workplan • Analisis Beban Kerja Personil & Realisasi Anggaran</div>
</div>
""", unsafe_allow_html=True)

if uploaded_file:
    df = load_and_clean_data(uploaded_file, auto_plan=auto_schedule)
    
    if df is not None:
        # Hitung Metrik Utama
        total_sites = len(df)
        done_sites = len(df[df['Status Progress'] == 'Done (Selesai)'])
        progress_rate = (done_sites / total_sites) * 100 if total_sites > 0 else 0
        
        total_rab = df['Biaya'].sum()
        actual_cost = df.loc[df['Status Progress'] == 'Done (Selesai)', 'Biaya'].sum()
        
        # Render Kartu Metrik
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Target Site", f"{total_sites} Sites")
        col2.metric("Site Selesai (Done)", f"{done_sites} Sites", f"{progress_rate:.1f}% Progress")
        col3.metric("Total RAB Project", f"Rp {total_rab:,.0f}")
        col4.metric("Realisasi Biaya Terserap", f"Rp {actual_cost:,.0f}")
        
        st.write("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # 5. TAB ANALISIS PROFESIONAL
        # ==========================================
        tab1, tab2, tab3, tab4 = st.tabs([
            "👥 Analisis PIC / Engineer", 
            "📅 Timeline & Gantt Chart", 
            "🗺️ Sebaran Area & Biaya", 
            "📑 Database & Download Excel"
        ])
        
        # --- TAB 1: ANALISIS PER PIC ---
        with tab1:
            st.markdown("#### **Beban Kerja & Produktivitas per Personil (PIC / Engineer)**")
            st.caption("Menampilkan jumlah site total, site selesai, serta akumulasi anggaran yang dipegang oleh masing-masing engineer.")
            
            # Agregasi data per PIC
            pic_summary = df.groupby('PIC Engineer').agg(
                Total_Sites=('Site ID', 'count'),
                Done_Sites=('Status Progress', lambda x: (x == 'Done (Selesai)').sum()),
                Pending_Sites=('Status Progress', lambda x: (x == 'Pending (On-Plan)').sum()),
                Total_Budget=('Biaya', 'sum')
            ).reset_index().sort_values(by='Total_Sites', ascending=False)
            
            pic_summary['Completion Rate (%)'] = ((pic_summary['Done_Sites'] / pic_summary['Total_Sites']) * 100).round(1)
            
            # Grafik Bar PIC
            fig_pic = px.bar(
                pic_summary, x='PIC Engineer', y=['Done_Sites', 'Pending_Sites'],
                title="Jumlah Site Ditangani per Personil (PIC)",
                labels={'value': 'Jumlah Site', 'PIC Engineer': 'Nama Engineer', 'variable': 'Status'},
                color_discrete_map={'Done_Sites': '#10b981', 'Pending_Sites': '#f59e0b'},
                template='plotly_white'
            )
            fig_pic.update_layout(xaxis_tickangle=-45, height=450)
            st.plotly_chart(fig_pic, use_container_width=True)
            
            # Tabel Detail PIC
            st.markdown("##### Tabel Rincian Per Personil")
            st.dataframe(
                pic_summary.rename(columns={
                    'PIC Engineer': 'Nama Personil (PIC)',
                    'Total_Sites': 'Total Site',
                    'Done_Sites': 'Site Selesai',
                    'Pending_Sites': 'Site Pending',
                    'Total_Budget': 'Total RAB (Rp)'
                }), 
                use_container_width=True, 
                hide_index=True
            )

        # --- TAB 2: TIMELINE / GANTT ---
        with tab2:
            st.markdown("#### **Jadwal Eksekusi Kerja (3 Bulan / 90 Hari Kedepan)**")
            df_plot = df.dropna(subset=['Plan Date']).sort_values('Plan Date')
            
            if not df_plot.empty:
                fig_gantt = px.timeline(
                    df_plot, x_start="Plan Date", x_end="Plan End", y="Site ID", color="Status Progress",
                    hover_data=["Area/City", "PIC Engineer", "Source Program", "Biaya"],
                    color_discrete_map={"Done (Selesai)": "#10b981", "Pending (On-Plan)": "#3b82f6"},
                    template="plotly_white"
                )
                fig_gantt.update_yaxes(autorange="reversed")
                fig_gantt.update_layout(height=600, margin=dict(t=20, b=20))
                st.plotly_chart(fig_gantt, use_container_width=True)

        # --- TAB 3: SEBARAN AREA & BIAYA ---
        with tab3:
            st.markdown("#### **Analisis Wilayah & Distribusi Anggaran**")
            col_a, col_b = st.columns(2)
            
            with col_a:
                area_count = df.groupby(['Area/City', 'Status Progress']).size().reset_index(name='Jumlah')
                fig_area = px.bar(
                    area_count, x='Area/City', y='Jumlah', color='Status Progress',
                    title="Volume Site per Kabupaten / Area",
                    color_discrete_map={"Done (Selesai)": "#10b981", "Pending (On-Plan)": "#f59e0b"},
                    template="plotly_white"
                )
                fig_area.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_area, use_container_width=True)
                
            with col_b:
                area_budget = df.groupby('Area/City')['Biaya'].sum().reset_index()
                fig_pie = px.pie(
                    area_budget, values='Biaya', names='Area/City', hole=0.4,
                    title="Proporsi Anggaran Berdasarkan Area",
                    template="plotly_white"
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

        # --- TAB 4: DATABASE & DOWNLOAD EXCEL ---
        with tab4:
            st.markdown("#### **Master Data Rekapitulasi & Unduh Excel**")
            
            df_display = df.copy()
            df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Selesai')
            df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d-%b-%Y').fillna('-')
            df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y').fillna('-')
            
            cols = ['Site ID', 'Source Program', 'Area/City', 'PIC Engineer', 'Kategori', 'Status Progress', 'Plan Date', 'Date Actual', 'Biaya']
            available_cols = [c for c in cols if c in df_display.columns]
            
            st.dataframe(df_display[available_cols], use_container_width=True, height=450)
            
            # Tombol Download Excel Profesional
            excel_bytes = export_excel(df)
            st.download_button(
                label="📥 Download Master Rekap Project ke Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"Master_Plan_BCP_SPS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
else:
    st.info("👈 **Silakan unggah file Excel Anda melalui panel di sebelah kiri** untuk memulai analisis data secara otomatis.")
