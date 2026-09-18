import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import io

# ==========================================
# 1. SETUP HALAMAN & STYLING
# ==========================================
st.set_page_config(page_title="BCP & SPS Visit Planner", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .main-title { font-size: 30px; font-weight: bold; color: #0f172a; margin-bottom: 0px; }
    .sub-title { font-size: 16px; color: #64748b; margin-bottom: 25px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Dashboard Operasional: BCP & SPS Visit</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Sistem Scheduling Otomatis (90 Hari) & Analisa Biaya Realisasi</div>', unsafe_allow_html=True)

# ==========================================
# 2. FUNGSI LOGIKA (AUTO-PLAN & EXPORT)
# ==========================================
def apply_3_month_schedule(df):
    """Mendistribusikan site yang belum memiliki jadwal ke rentang 90 hari"""
    start_date = pd.to_datetime('today').normalize()
    
    # Cari site yang Plan Date-nya kosong atau NaT
    mask_needs_plan = df['Plan Date'].isna()
    total_sites_to_plan = mask_needs_plan.sum()
    
    if total_sites_to_plan > 0:
        step_days = 90 / total_sites_to_plan
        starts = [start_date + timedelta(days=int(i * step_days)) for i in range(total_sites_to_plan)]
        
        df.loc[mask_needs_plan, 'Plan Date'] = starts
        
    # Set Plan End = Plan Date + 2 Hari (SLA Pengerjaan)
    df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2)
    return df

def convert_df_to_excel(df):
    """Fungsi export dataframe pandas ke format Excel murni"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Jadwal_Kerja_Terupdate')
    return output.getvalue()

# ==========================================
# 3. UPLOAD FILE & DATA PROCESSING
# ==========================================
st.sidebar.header("📁 Konfigurasi & Upload")
uploaded_file = st.sidebar.file_uploader("Unggah File (BCP & SPS visit.xlsx)", type=["xlsx"])
auto_schedule = st.sidebar.checkbox("Aktifkan Penjadwalan Otomatis (90 Hari)", value=True)

if uploaded_file:
    # 1. Membaca file dengan deteksi Header Pintar (Karena header asli di baris ke-2 / index 1)
    df_raw = pd.read_excel(uploaded_file)
    if 'Site ID' not in df_raw.columns:
        # Jika 'Site ID' tidak ada di column, kemungkinan itu judul merged cell, baca ulang dari baris ke-2
        df = pd.read_excel(uploaded_file, header=1)
    else:
        df = df_raw.copy()
        
    # Bersihkan nama kolom dari spasi berlebih
    df.rename(columns=lambda x: str(x).strip(), inplace=True)
    
    # 2. Pembersihan Data Biaya
    if 'Biaya Onsite' in df.columns:
        df['Biaya Onsite'] = df['Biaya Onsite'].astype(str).str.replace(r'\D', '', regex=True)
        df['Biaya Onsite'] = pd.to_numeric(df['Biaya Onsite'], errors='coerce').fillna(0)
    else:
        df['Biaya Onsite'] = 0

    # 3. Status Tracking & Konversi Waktu Aktual
    if 'Date Actual' in df.columns:
        df['Date Actual'] = pd.to_datetime(df['Date Actual'], errors='coerce')
        df['Status Tracking'] = df['Date Actual'].apply(lambda x: 'Done' if pd.notnull(x) else 'Plan / Pending')
    else:
        df['Status Tracking'] = 'Plan / Pending'
        df['Date Actual'] = pd.NaT

    # 4. Handle Plan Date (Ubah text seperti 'NY' jadi kosong/NaT)
    if 'Plan Date' in df.columns:
        df['Plan Date'] = pd.to_datetime(df['Plan Date'], errors='coerce')
    else:
        df['Plan Date'] = pd.NaT

    # 5. Eksekusi Penjadwalan Otomatis
    if auto_schedule:
        df = apply_3_month_schedule(df)
    else:
        df['Plan End'] = df['Plan Date'] + pd.Timedelta(days=2)

    # Gabungkan typo penulisan City (Misal: GUNUNG MAS vs Gunung Mas)
    if 'City' in df.columns:
        df['City'] = df['City'].astype(str).str.upper()

    # ==========================================
    # 4. DASHBOARD KPI (METRICS)
    # ==========================================
    total_sites = len(df)
    total_done = len(df[df['Status Tracking'] == 'Done'])
    progress = (total_done / total_sites) * 100 if total_sites > 0 else 0
    total_budget = df['Biaya Onsite'].sum()
    actual_spent = df.loc[df['Status Tracking'] == 'Done', 'Biaya Onsite'].sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Site Target", f"{total_sites} Site")
    col2.metric("Pekerjaan Selesai (Done)", f"{total_done} Site", f"{progress:.1f}% Progress")
    col3.metric("Estimasi Total Biaya (RAB)", f"Rp {total_budget:,.0f}")
    col4.metric("Biaya Terserap Aktual", f"Rp {actual_spent:,.0f}")
    st.markdown("---")

    # ==========================================
    # 5. VISUALISASI & TAB KONTROL
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📅 Timeline Project (Gantt)", "🗺️ Distribusi Area", "📋 Database & Export Excel"])

    with tab1:
        st.subheader("Jadwal Eksekusi Harian (90 Hari Kedepan)")
        df_plot = df.dropna(subset=['Plan Date']).copy()
        if not df_plot.empty:
            fig_gantt = px.timeline(
                df_plot, 
                x_start="Plan Date", 
                x_end="Plan End", 
                y="Site ID", 
                color="Status Tracking",
                hover_data=["City", "TE NAME", "Final"], 
                color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"}
            )
            fig_gantt.update_yaxes(autorange="reversed")
            fig_gantt.update_layout(height=600)
            st.plotly_chart(fig_gantt, use_container_width=True)
            st.info("💡 Tip: Arahkan kursor ke pojok kanan atas grafik ini dan klik ikon kamera (Download plot as a png) untuk menyimpan grafik ke gambar.")

    with tab2:
        st.subheader("Sebaran Pekerjaan dan Serapan Biaya per Kabupaten/Kota")
        if 'City' in df.columns:
            # Chart 1: Jumlah Site per Kota
            city_count = df.groupby(['City', 'Status Tracking']).size().reset_index(name='Jumlah Site')
            fig_city = px.bar(
                city_count, x="City", y="Jumlah Site", color="Status Tracking", 
                barmode="group", color_discrete_map={"Done": "#10B981", "Plan / Pending": "#F59E0B"},
                title="Beban Kerja Per NOP"
            )
            st.plotly_chart(fig_city, use_container_width=True)
            
            # Chart 2: Total Biaya Plan per Kota
            city_budget = df.groupby('City')['Biaya Onsite'].sum().reset_index()
            fig_budget = px.pie(
                city_budget, values='Biaya Onsite', names='City', hole=0.4, 
                title="Distribusi RAB (Plan Budget) per Kota"
            )
            st.plotly_chart(fig_budget, use_container_width=True)

    with tab3:
        st.subheader("Database Rekapitulasi (Siap Unduh)")
        
        # Display copy agar format tanggal bersih di layar (tanpa mengubah data asli di memori)
        df_display = df.copy()
        df_display['Date Actual'] = df_display['Date Actual'].dt.strftime('%d-%b-%Y').fillna('Belum Dieksekusi')
        df_display['Plan Date'] = df_display['Plan Date'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        df_display['Plan End'] = df_display['Plan End'].dt.strftime('%d-%b-%Y').fillna('No Plan')
        
        # Munculkan tabel untuk di review
        st.dataframe(df_display[['Site ID', 'City', 'TE NAME', 'Final', 'Status Tracking', 'Plan Date', 'Date Actual', 'Biaya Onsite']], use_container_width=True, height=400)
        
        # Tombol eksekusi Export
        excel_data = convert_df_to_excel(df)
        st.download_button(
            label="📥 Download Hasil Perencanaan ke File Excel (.xlsx)",
            data=excel_data,
            file_name=f"BCP_SPS_Planner_Update_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Silakan unggah file Excel ('BCP & SPS visit_2.xlsx' atau sejenisnya) di menu samping (sidebar) kiri untuk memunculkan Dashboard.")
