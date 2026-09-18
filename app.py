import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# ==========================================
# 1. KONFIGURASI HALAMAN & DESIGN SYSTEM
# ==========================================
st.set_page_config(
    page_title="Executive Dashboard: BCP & SPS Visit",
    layout="wide",
    page_icon="📊"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 30px 35px;
        border-radius: 14px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
    }
    .hero-title { font-size: 28px; font-weight: 700; margin: 0; letter-spacing: -0.5px; }
    .hero-subtitle { font-size: 14px; color: #94a3b8; margin-top: 5px; }
    
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
# 2. CORE ENGINE: CLEANING & MERGING MULTI-SHEET
# ==========================================
@st.cache_data
def load_and_process_data(file_bytes, auto_schedule=True):
    try:
        # Gunakan BytesIO untuk keamanan pembacaan buffer file di Streamlit
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        
        # --- Sheet BCP ---
        df_bcp = pd.DataFrame()
        if 'BCP' in xls.sheet_names:
            df_bcp = pd.read_excel(xls, sheet_name='BCP', header=1)
            df_bcp.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            bcp_map = {
                'Site ID': 'Site ID', 'City': 'Area/City', 'TE NAME': 'PIC Engineer', 
                'Final': 'Kategori', 'Biaya Onsite': 'Biaya', 'Plan Date': 'Plan Date', 'Date Actual': 'Date Actual'
            }
            df_bcp = df_bcp.rename(columns={k: v for k, v in bcp_map.items() if k in df_bcp.columns})
            df_bcp['Source Sheet'] = 'BCP Visit'
            
        # --- Sheet SPS Visit ---
        df_sps = pd.DataFrame()
        if 'SPS Visit' in xls.sheet_names:
            df_sps = pd.read_excel(xls, sheet_name='SPS Visit', header=0)
            df_sps.rename(columns=lambda x: str(x).strip(), inplace=True)
            
            sps_map = {
                'Site ID': 'Site ID', 'Kabupaten': 'Area/City', 'TE Name': 'PIC Engineer', 
                'TOTAL': 'Biaya', 'Plan Visited': 'Plan Date'
            }
            df_sps = df_sps.rename(columns={k: v for k, v in sps_map.items() if k in df_sps.columns})
            df_sps['Source Sheet'] = 'SPS Visit'
            df_sps['Kategori'] = 'SPS Visit'
            df_sps['Date Actual'] = pd.NaT

        # Gabungkan Data Master secara aman
        master_df = pd.concat([df_bcp, df_sps], ignore_index=True)
        
        # Pembersihan Nominal Biaya (Murni Jutaan / Ratusan Ribu Rupiah)
        if 'Biaya' in master_df.columns:
            master_df['Biaya'] = pd.to_numeric(master_df['Biaya'], errors='coerce')
            median_cost = master_df['Biaya'].median() if not master_df['Biaya'].dropna().empty else 500000
            master_df['Biaya'] = master_df['Biaya'].fillna(median_cost).fillna(500000)
        else:
            master_df['Biaya'] = 500000

        # Standarisasi Teks Wilayah & Nama PIC
        if 'Area/City' in master_df.columns:
            master_df['Area/City'] = master_df['Area/City'].astype(str).str.upper().str.strip()
        if 'PIC Engineer' in master_df.columns:
            master_df['PIC Engineer'] = master_df['PIC Engineer'].astype(str).str.title().str.strip()
            master_df['PIC Engineer'] = master_df['PIC Engineer'].replace(['Nan', '', 'Na', 'None'], 'Unassigned')

        # Status Tracking Berdasarkan Tanggal Actual
        if 'Date Actual' in master_df.columns:
            master_df['Date Actual'] = pd.to_datetime(master_df['Date Actual'], errors='coerce')
        else:
            master_df['Date Actual'] = pd.NaT
            
        master_df['Status Progress'] = master_df['Date Actual'].apply(lambda x: 'Done (Selesai)' if pd.notnull(x) else 'Pending (On-Plan)')

        # Penjadwalan Otomatis 90 Hari (3 Bulan) untuk Site yang Kosong
        master_df['Plan Date'] = pd.to_datetime(master_df['Plan Date'], errors='coerce')
        if auto_schedule:
            mask_empty = master_df['Plan Date'].isna()
            if mask_empty.sum() > 0:
                start_date = pd.to_datetime('today').normalize()
                step = 90 / mask_empty.sum()
                starts = [start_date + timedelta(days=int(i * step)) for i in range(mask_empty.sum())]
                master_df.loc[mask_empty, 'Plan Date'] = starts
                
        master_df['Plan End'] = master_df['Plan Date'] + pd.Timedelta(days=2)
        return master_df
        
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
        return None

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Master_Report')
    return output.getvalue()

# ==========================================
# 3. SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=45)
    st.markdown("### **Panel Kontrol Data**")
    uploaded_file = st.file_uploader("📂 Unggah File Tracker (.xlsx)", type=["xlsx"])
    auto_schedule = st.toggle("Aktifkan Auto-Scheduler (3 Bulan)", value=True)
    st.markdown("---")
    st.caption("Aplikasi ini otomatis menggabungkan sheet BCP dan SPS Visit, membersihkan nominal RAB, dan menghitung beban kerja tim secara real-time.")

# ==========================================
# 4. TAMPILAN UTAMA (DASHBOARD)
# ==========================================
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">Executive Operations & Workplan Dashboard</div>
    <div class="hero-subtitle">Monitoring Terintegrasi BCP & SPS Visit • Analisis Beban Kerja Personil (PIC) & Realisasi Anggaran</div>
</div>
""", unsafe_allow_html=True)

if uploaded_file:
    # Ambil bytes dari file upload untuk menghindari TypeError buffer stream
    file_bytes = uploaded_file.getvalue()
    df = load_and_process_data(file_bytes, auto_plan=auto_schedule)
    
    if df is not None:
        # Metrik Utama
        total_sites = len(df)
        done_sites = len(df[df['Status Progress'] == 'Done (Selesai)'])
        progress_rate = (done_sites / total_sites) * 100 if total_sites > 0 else 0
        
        total_rab = df['Biaya'].sum()
        actual_spent = df.loc[df['Status Progress'] == 'Done (Selesai)', 'Biaya'].sum()
        
        # Render Kartu Metrik
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Target Site", f"{total_sites} Sites")
        col2.metric("Site Selesai (Done)", f"{done_sites} Sites", f"{progress_rate:.1f}% Progress")
        col3.metric("Total RAB Project", f"Rp {total_rab:,.0f}")
        col4.metric("Realisasi Biaya Terserap", f"Rp {actual_spent:,.0f}")
        
        st.write("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # 5. TAB ANALISIS MENDALAM
        # ==========================================
        tab1, tab2, tab3, tab4 = st.tabs([
            "👥 Analisis Beban Kerja PIC", 
            "📅 Timeline & Gantt Chart", 
            "🗺️ Sebaran Area & Biaya", 
            "📑 Database & Download Excel"
        ])
        
        # --- TAB 1: ANALISIS PIC / ENGINEER ---
        with tab1:
            st.markdown("#### **Beban Kerja & Jumlah Site per Personil (PIC / Engineer)**")
            st.caption("Mengetahui secara transparan berapa total site yang dipegang oleh masing-masing personil, status pengerjaan, serta total anggaran.")
            
            pic_summary = df.groupby('PIC Engineer').agg(
                Total_Sites=('Site ID', 'count'),
                Done_Sites=('Status Progress', lambda x: (x == 'Done (Selesai)').sum()),
                Pending_Sites=('Status Progress', lambda x: (x == 'Pending (On-Plan)').sum()),
                Total_Budget=('Biaya', 'sum')
            ).reset_index().sort_values(by='Total_Sites', ascending=False)
            
            # Bar Chart PIC
            fig_pic = px.bar(
                pic_summary, x='PIC Engineer', y=['Done_Sites', 'Pending_Sites'],
                title="Distribusi Jumlah Site Ditangani per Personil",
                labels={'value': 'Jumlah Site', 'PIC Engineer': 'Nama Personil (PIC)', 'variable': 'Status Pengerjaan'},
                color_discrete_map={'Done_Sites': '#10b981', 'Pending_Sites': '#f59e0b'},
                template='plotly_white'
            )
            fig_pic.update_layout(xaxis_tickangle=-45, height=450)
            st.plotly_chart(fig_pic, use_container_width=True)
            
            st.markdown("##### Tabel Rincian Beban Kerja Personil")
            st.dataframe(
                pic_summary.rename(columns={
                    'PIC Engineer': 'Nama Personil (PIC)',
                    'Total_Sites': 'Total Site',
                    'Done_Sites': 'Site Selesai',
                    'Pending_Sites': 'Site Pending',
                    'Total_Budget': 'Akumulasi Biaya (Rp)'
                }), 
                use_container_width=True, 
                hide_index=True
            )

        # --- TAB 2: TIMELINE / GANTT CHART ---
        with tab2:
            st.markdown("#### **Jadwal Eksekusi Kerja 3 Bulan (90 Hari Kedepan)**")
            df_plot = df.dropna(subset=['Plan Date']).sort_values('Plan Date')
            
            if not df_plot.empty:
                fig_gantt = px.timeline(
                    df_plot, x_start="Plan Date", x_end="Plan End", y="Site ID", color="Status Progress",
                    hover_data=["Area/City", "PIC Engineer", "Source Sheet", "Biaya"],
                    color_discrete_map={"Done (Selesai)": "#10b981", "Pending (On-Plan)": "#3b82f6"},
                    template="plotly_white"
                )
                fig_gantt.update_yaxes(autorange="reversed")
                fig_gantt.update_layout(height=600, margin=dict(t=20, b=20))
                st.plotly_chart(fig_gantt, use_container_width=True)

        # --- TAB 3: SEBARAN WILAYAH & BIAYA ---
        with tab3:
            st.markdown("#### **Analisis Wilayah & Alokasi Anggaran**")
            col_a, col_b = st.columns(2)
            
            with col_a:
                area_count = df.groupby(['Area/City', 'Status Progress']).size().reset_index(name='Jumlah')
                fig_area = px.bar(
                    area_count, x='Area/City', y='Jumlah', color='Status Progress',
                    title="Volume Pekerjaan per Kabupaten / Area",
                    color_discrete_map={"Done (Selesai)": "#10b981", "Pending (On-Plan)": "#f59e0b"},
                    template="plotly_white"
                )
                fig_area.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_area, use_container_width=True)
                
            with col_b:
                area_budget = df.groupby('Area/City')['Biaya'].sum().reset_index()
                fig_pie = px.pie(
                    area_budget, values='Biaya', names='Area/City', hole=0.4,
                    title="Proporsi Anggaran Berdasarkan Wilayah",
                    template="plotly_white"
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

        # --- TAB 4: DATABASE & EXPORT EXCEL ---
        with tab4:
            st.markdown("#### **Master Data Rekapitulasi & Unduh Laporan**")
            
            df_display = df.copy()
            df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Selesai')
            df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d-%b-%Y').fillna('-')
            df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y').fillna('-')
            
            cols = ['Site ID', 'Source Sheet', 'Area/City', 'PIC Engineer', 'Kategori', 'Status Progress', 'Plan Date', 'Date Actual', 'Biaya']
            available_cols = [c for c in cols if c in df_display.columns]
            
            st.dataframe(df_display[available_cols], use_container_width=True, height=450)
            
            # Tombol Download Excel Otomatis Berfungsi
            excel_bytes = convert_df_to_excel(df)
            st.download_button(
                label="📥 Download Master Rekap Project ke Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"Master_Plan_BCP_SPS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
else:
    st.info("👈 **Silakan unggah file Excel Anda (misal: BCP & SPS visit.xlsx) melalui panel di sebelah kiri** untuk menampilkan seluruh analisis dan grafik secara otomatis.")
