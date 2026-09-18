import io
from datetime import datetime, timedelta
import pandas as pd
import plotly.express as px
import streamlit as st

import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. KONFIGURASI HALAMAN & DESIGN SYSTEM
# ==========================================
st.set_page_config(
    page_title="Executive Dashboard: BCP & SPS Visit",
    layout="wide",
    page_icon="📊",
)

st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. CORE ENGINE: DATA PROCESSOR
# ==========================================
def load_and_process_data(file_bytes, auto_schedule=True):
  try:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    # --- Sheet BCP ---
    df_bcp = pd.DataFrame()
    if "BCP" in xls.sheet_names:
      df_bcp = pd.read_excel(xls, sheet_name="BCP", header=1)
      df_bcp.rename(columns=lambda x: str(x).strip(), inplace=True)

      bcp_map = {
          "Site ID": "Site ID",
          "City": "Area/City",
          "TE NAME": "PIC Engineer",
          "Final": "Kategori",
          "Biaya Onsite": "Biaya",
          "Plan Date": "Plan Date",
          "Date Actual": "Date Actual",
      }
      df_bcp = df_bcp.rename(
          columns={k: v for k, v in bcp_map.items() if k in df_bcp.columns}
      )
      df_bcp["Source Sheet"] = "BCP Visit"

      df_bcp["Biaya"] = pd.to_numeric(df_bcp["Biaya"], errors="coerce")
      df_bcp["Plan Date"] = pd.to_datetime(df_bcp["Plan Date"], errors="coerce")
      df_bcp["Date Actual"] = pd.to_datetime(
          df_bcp["Date Actual"], errors="coerce"
      )

    # --- Sheet SPS Visit ---
    df_sps = pd.DataFrame()
    if "SPS Visit" in xls.sheet_names:
      df_sps = pd.read_excel(xls, sheet_name="SPS Visit", header=0)
      df_sps.rename(columns=lambda x: str(x).strip(), inplace=True)

      sps_map = {
          "Site ID": "Site ID",
          "Kabupaten": "Area/City",
          "TE Name": "PIC Engineer",
          "TOTAL": "Biaya",
          "Plan Visited": "Plan Date",
      }
      df_sps = df_sps.rename(
          columns={k: v for k, v in sps_map.items() if k in df_sps.columns}
      )
      df_sps["Source Sheet"] = "SPS Visit"
      df_sps["Kategori"] = "SPS Visit"
      df_sps["Date Actual"] = pd.NaT

      df_sps["Biaya"] = pd.to_numeric(df_sps["Biaya"], errors="coerce")
      df_sps["Plan Date"] = pd.to_datetime(df_sps["Plan Date"], errors="coerce")

    # Gabungkan Data Master
    master_df = pd.concat([df_bcp, df_sps], ignore_index=True)
    master_df = master_df.dropna(subset=["Site ID"]).copy()

    # Imputasi biaya kosong dengan median agar total RAB akurat
    median_cost = (
        master_df["Biaya"].median()
        if not master_df["Biaya"].dropna().empty
        else 500000
    )
    master_df["Biaya"] = master_df["Biaya"].fillna(median_cost).fillna(500000)

    # Standarisasi Teks Wilayah & Nama PIC
    if "Area/City" in master_df.columns:
      master_df["Area/City"] = (
          master_df["Area/City"].astype(str).str.upper().str.strip()
      )
    if "PIC Engineer" in master_df.columns:
      master_df["PIC Engineer"] = (
          master_df["PIC Engineer"].astype(str).str.title().str.strip()
      )
      master_df["PIC Engineer"] = master_df["PIC Engineer"].replace(
          ["Nan", "", "Na", "None"], "Unassigned"
      )

    # Status Tracking Berdasarkan Tanggal Actual
    master_df["Status Progress"] = master_df["Date Actual"].apply(
        lambda x: "Done (Selesai)" if pd.notnull(x) else "Pending (On-Plan)"
    )

    # Penjadwalan Otomatis 90 Hari (3 Bulan) untuk Site yang Kosong
    if auto_schedule:
      mask_empty = master_df["Plan Date"].isna()
      if mask_empty.sum() > 0:
        start_date = pd.to_datetime("today").normalize()
        step = 90 / mask_empty.sum()
        starts = [
            start_date + timedelta(days=int(i * step))
            for i in range(mask_empty.sum())
        ]
        master_df.loc[mask_empty, "Plan Date"] = starts

    master_df["Plan End"] = master_df["Plan Date"] + pd.Timedelta(days=2)
    return master_df

  except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses file: {e}")
    return None


# ==========================================
# 3. GENERATOR EXCEL EXECUTIVE REPORT
# ==========================================
def convert_df_to_excel(df):
  output = io.BytesIO()
  wb = openpyxl.Workbook()

  # Definisi Palette Warna & Font
  NAVY_HEADER_FILL = PatternFill(
      start_color="0F172A", end_color="0F172A", fill_type="solid"
  )
  KPI_TITLE_FILL = PatternFill(
      start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
  )
  KPI_VAL_FILL = PatternFill(
      start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
  )
  ZEBRA_FILL = PatternFill(
      start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
  )
  WHITE_FILL = PatternFill(
      start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
  )
  TOTAL_FILL = PatternFill(
      start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"
  )

  FONT_TITLE = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
  FONT_SUBTITLE = Font(name="Segoe UI", size=9, italic=True, color="94A3B8")
  FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
  FONT_SEC_HEADER = Font(name="Segoe UI", size=12, bold=True, color="0F172A")
  FONT_KPI_NUM = Font(name="Segoe UI", size=15, bold=True, color="0F172A")
  FONT_KPI_LABEL = Font(name="Segoe UI", size=8, bold=True, color="475569")
  FONT_BODY = Font(name="Segoe UI", size=9, color="0F172A")
  FONT_BOLD = Font(name="Segoe UI", size=9, bold=True, color="0F172A")

  THIN_SIDE = Side(border_style="thin", color="CBD5E1")
  THIN_BORDER = Border(
      left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE
  )
  DOUBLE_BOTTOM = Border(
      top=THIN_SIDE,
      bottom=Side(border_style="double", color="0F172A"),
      left=THIN_SIDE,
      right=THIN_SIDE,
  )

  ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
  ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
  ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
  ALIGN_HEADER = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )

  # -------------------------------------------------------------
  # TAB 1: EXECUTIVE DASHBOARD
  # -------------------------------------------------------------
  ws_dash = wb.active
  ws_dash.title = "📌 Executive Dashboard"
  ws_dash.views.sheetView[0].showGridLines = True

  # Header Banner Block
  for r in range(1, 3):
    for c in range(1, 11):
      cell = ws_dash.cell(row=r, column=c)
      cell.fill = NAVY_HEADER_FILL

  ws_dash.merge_cells("A1:J1")
  ws_dash.merge_cells("A2:J2")

  ws_dash["A1"] = "EXECUTIVE OPERATIONS & WORKPLAN REPORT"
  ws_dash["A1"].font = FONT_TITLE
  ws_dash["A1"].alignment = ALIGN_LEFT

  ws_dash["A2"] = (
      "Monitoring BCP & SPS Visit • Auto-Generated Analytics Summary"
  )
  ws_dash["A2"].font = FONT_SUBTITLE
  ws_dash["A2"].alignment = ALIGN_LEFT

  ws_dash.row_dimensions[1].height = 24
  ws_dash.row_dimensions[2].height = 18

  # Hitung Metrik KPI
  total_sites = len(df)
  done_sites = (
      len(df[df["Status Progress"] == "Done (Selesai)"])
      if "Status Progress" in df.columns
      else 0
  )
  progress_pct = (done_sites / total_sites) if total_sites > 0 else 0
  total_rab = df["Biaya"].sum() if "Biaya" in df.columns else 0
  actual_spent = (
      df.loc[df["Status Progress"] == "Done (Selesai)", "Biaya"].sum()
      if "Biaya" in df.columns and "Status Progress" in df.columns
      else 0
  )

  kpis = [
      ("TOTAL TARGET SITE", total_sites, "0", "A", "B"),
      ("SITE DONE (SELESAI)", done_sites, "0", "C", "D"),
      ("PROGRESS RATE", progress_pct, "0.0%", "E", "F"),
      ("TOTAL RAB PROJECT", total_rab, '"Rp "#,##0', "G", "H"),
      ("REALISASI BIAYA", actual_spent, '"Rp "#,##0', "I", "J"),
  ]

  for title, val, num_fmt, col_start, col_end in kpis:
    c1, c2 = f"{col_start}4", f"{col_end}4"
    v1, v2 = f"{col_start}5", f"{col_end}5"

    ws_dash.merge_cells(f"{c1}:{c2}")
    ws_dash.merge_cells(f"{v1}:{v2}")

    ws_dash[c1] = title
    ws_dash[c1].font = FONT_KPI_LABEL
    ws_dash[c1].fill = KPI_TITLE_FILL
    ws_dash[c1].alignment = ALIGN_CENTER

    ws_dash[v1] = val
    ws_dash[v1].font = FONT_KPI_NUM
    ws_dash[v1].fill = KPI_VAL_FILL
    ws_dash[v1].alignment = ALIGN_CENTER
    ws_dash[v1].number_format = num_fmt

    for r in range(4, 6):
      for col_letter in [col_start, col_end]:
        c_idx = openpyxl.utils.column_index_from_string(col_letter)
        ws_dash.cell(row=r, column=c_idx).border = THIN_BORDER

  ws_dash.row_dimensions[4].height = 18
  ws_dash.row_dimensions[5].height = 26

  # Tabel Summary 1: Beban Kerja PIC
  ws_dash["A7"] = "📊 Summary Beban Kerja PIC / Engineer"
  ws_dash["A7"].font = FONT_SEC_HEADER

  pic_summary = (
      df.groupby("PIC Engineer")
      .agg(
          Total_Sites=("Site ID", "count"),
          Done_Sites=(
              "Status Progress",
              lambda x: (x == "Done (Selesai)").sum(),
          ),
          Pending_Sites=(
              "Status Progress",
              lambda x: (x == "Pending (On-Plan)").sum(),
          ),
          Total_Budget=("Biaya", "sum"),
      )
      .reset_index()
      .sort_values(by="Total_Sites", ascending=False)
  )

  pic_headers = [
      "Nama Personil (PIC)",
      "Total Site",
      "Site Selesai",
      "Site Pending",
      "Total Budget (Rp)",
  ]
  for c_idx, h in enumerate(pic_headers, start=1):
    cell = ws_dash.cell(row=8, column=c_idx, value=h)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER_FILL
    cell.alignment = ALIGN_HEADER
    cell.border = THIN_BORDER

  ws_dash.row_dimensions[8].height = 22

  row_idx = 9
  start_pic_row = row_idx
  for _, row in pic_summary.iterrows():
    fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
    c1 = ws_dash.cell(row=row_idx, column=1, value=row["PIC Engineer"])
    c2 = ws_dash.cell(row=row_idx, column=2, value=row["Total_Sites"])
    c3 = ws_dash.cell(row=row_idx, column=3, value=row["Done_Sites"])
    c4 = ws_dash.cell(row=row_idx, column=4, value=row["Pending_Sites"])
    c5 = ws_dash.cell(row=row_idx, column=5, value=row["Total_Budget"])

    for c in [c1, c2, c3, c4, c5]:
      c.font = FONT_BODY
      c.fill = fill
      c.border = THIN_BORDER

    c1.alignment = ALIGN_LEFT
    c2.alignment = ALIGN_CENTER
    c2.number_format = "#,##0"
    c3.alignment = ALIGN_CENTER
    c3.number_format = "#,##0"
    c4.alignment = ALIGN_CENTER
    c4.number_format = "#,##0"
    c5.alignment = ALIGN_RIGHT
    c5.number_format = '"Rp "#,##0'
    row_idx += 1

  end_pic_row = row_idx - 1
  tot_row = row_idx
  ws_dash.cell(row=tot_row, column=1, value="TOTAL").alignment = ALIGN_LEFT
  ws_dash.cell(
      row=tot_row,
      column=2,
      value=f"=SUM(B{start_pic_row}:B{end_pic_row})",
  ).number_format = "#,##0"
  ws_dash.cell(
      row=tot_row,
      column=3,
      value=f"=SUM(C{start_pic_row}:C{end_pic_row})",
  ).number_format = "#,##0"
  ws_dash.cell(
      row=tot_row,
      column=4,
      value=f"=SUM(D{start_pic_row}:D{end_pic_row})",
  ).number_format = "#,##0"
  ws_dash.cell(
      row=tot_row,
      column=5,
      value=f"=SUM(E{start_pic_row}:E{end_pic_row})",
  ).number_format = '"Rp "#,##0'

  for c_idx in range(1, 6):
    cell = ws_dash.cell(row=tot_row, column=c_idx)
    cell.font = FONT_BOLD
    cell.fill = TOTAL_FILL
    cell.border = DOUBLE_BOTTOM
    if c_idx in [2, 3, 4]:
      cell.alignment = ALIGN_CENTER
    elif c_idx == 5:
      cell.alignment = ALIGN_RIGHT

  # Tabel Summary 2: Sebaran Area
  area_start_row = tot_row + 3
  ws_dash.cell(
      row=area_start_row - 1,
      column=1,
      value="🗺️ Summary Sebaran Area & Anggaran",
  ).font = FONT_SEC_HEADER

  area_summary = (
      df.groupby("Area/City")
      .agg(Total_Sites=("Site ID", "count"), Total_Budget=("Biaya", "sum"))
      .reset_index()
      .sort_values(by="Total_Sites", ascending=False)
  )

  area_headers = ["Kabupaten / Area", "Jumlah Site", "Total Budget (Rp)"]
  for c_idx, h in enumerate(area_headers, start=1):
    cell = ws_dash.cell(row=area_start_row, column=c_idx, value=h)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER_FILL
    cell.alignment = ALIGN_HEADER
    cell.border = THIN_BORDER

  ws_dash.row_dimensions[area_start_row].height = 22

  r_idx = area_start_row + 1
  start_area_row = r_idx
  for _, row in area_summary.iterrows():
    fill = ZEBRA_FILL if r_idx % 2 == 0 else WHITE_FILL
    c1 = ws_dash.cell(row=r_idx, column=1, value=row["Area/City"])
    c2 = ws_dash.cell(row=r_idx, column=2, value=row["Total_Sites"])
    c3 = ws_dash.cell(row=r_idx, column=3, value=row["Total_Budget"])

    for c in [c1, c2, c3]:
      c.font = FONT_BODY
      c.fill = fill
      c.border = THIN_BORDER

    c1.alignment = ALIGN_LEFT
    c2.alignment = ALIGN_CENTER
    c2.number_format = "#,##0"
    c3.alignment = ALIGN_RIGHT
    c3.number_format = '"Rp "#,##0'
    r_idx += 1

  end_area_row = r_idx - 1
  tot_area_row = r_idx
  ws_dash.cell(row=tot_area_row, column=1, value="TOTAL").alignment = ALIGN_LEFT
  ws_dash.cell(
      row=tot_area_row,
      column=2,
      value=f"=SUM(B{start_area_row}:B{end_area_row})",
  ).number_format = "#,##0"
  ws_dash.cell(
      row=tot_area_row,
      column=3,
      value=f"=SUM(C{start_area_row}:C{end_area_row})",
  ).number_format = '"Rp "#,##0'

  for c_idx in range(1, 4):
    cell = ws_dash.cell(row=tot_area_row, column=c_idx)
    cell.font = FONT_BOLD
    cell.fill = TOTAL_FILL
    cell.border = DOUBLE_BOTTOM
    if c_idx == 2:
      cell.alignment = ALIGN_CENTER
    elif c_idx == 3:
      cell.alignment = ALIGN_RIGHT

  # Grafik 1: Bar Chart PIC Workload
  chart1 = BarChart()
  chart1.type = "col"
  chart1.style = 10
  chart1.title = "Distribusi Site per PIC Engineer"
  chart1.y_axis.title = "Jumlah Site"
  chart1.x_axis.title = "PIC Engineer"

  data1 = Reference(
      ws_dash, min_col=3, min_row=8, max_col=4, max_row=end_pic_row
  )
  cats1 = Reference(ws_dash, min_col=1, min_row=9, max_row=end_pic_row)
  chart1.add_data(data1, titles_from_data=True)
  chart1.set_categories(cats1)
  chart1.width, chart1.height = 16, 10
  ws_dash.add_chart(chart1, "G7")

  # Grafik 2: Pie Chart Regional Budget
  chart2 = PieChart()
  chart2.title = "Proporsi Anggaran per Wilayah"
  data2 = Reference(
      ws_dash, min_col=3, min_row=area_start_row, max_row=end_area_row
  )
  labels2 = Reference(
      ws_dash, min_col=1, min_row=area_start_row + 1, max_row=end_area_row
  )
  chart2.add_data(data2, titles_from_data=True)
  chart2.set_categories(labels2)
  chart2.width, chart2.height = 14, 10
  ws_dash.add_chart(chart2, "G21")

  # Atur Lebar Kolom Dashboard
  for col in ws_dash.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_dash.column_dimensions[col_letter].width = max(max_len + 4, 15)

  ws_dash.column_dimensions["F"].width = 4

  # -------------------------------------------------------------
  # TAB 2: MASTER DATA REKAP
  # -------------------------------------------------------------
  ws_master = wb.create_sheet(title="📑 Master Data Rekap")
  ws_master.views.sheetView[0].showGridLines = True

  cols_to_export = [
      "Site ID",
      "Source Sheet",
      "Area/City",
      "PIC Engineer",
      "Kategori",
      "Status Progress",
      "Plan Date",
      "Date Actual",
      "Biaya",
  ]
  available_cols = [c for c in cols_to_export if c in df.columns]

  # Header Row
  for c_idx, col_name in enumerate(available_cols, start=1):
    cell = ws_master.cell(row=1, column=c_idx, value=col_name)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER_FILL
    cell.alignment = ALIGN_HEADER
    cell.border = THIN_BORDER

  ws_master.row_dimensions[1].height = 25

  # Data Rows
  for r_idx, (_, row) in enumerate(df[available_cols].iterrows(), start=2):
    fill = ZEBRA_FILL if r_idx % 2 == 0 else WHITE_FILL
    for c_idx, col_name in enumerate(available_cols, start=1):
      val = row[col_name]
      cell = ws_master.cell(row=r_idx, column=c_idx)
      cell.font = FONT_BODY
      cell.fill = fill
      cell.border = THIN_BORDER

      if "Date" in col_name or "Actual" in col_name:
        if pd.notnull(val) and isinstance(
            val, (pd.Timestamp, datetime, pd.DatetimeIndex)
        ):
          cell.value = pd.to_datetime(val).strftime("%d-%b-%Y")
          cell.alignment = ALIGN_CENTER
        else:
          cell.value = "-"
          cell.alignment = ALIGN_CENTER
      elif col_name == "Biaya":
        cell.value = float(val) if pd.notnull(val) else 0.0
        cell.number_format = '"Rp "#,##0'
        cell.alignment = ALIGN_RIGHT
      elif col_name in ["Site ID", "Status Progress", "Source Sheet", "Kategori"]:
        cell.value = str(val) if pd.notnull(val) else "-"
        cell.alignment = ALIGN_CENTER
      else:
        cell.value = str(val) if pd.notnull(val) else "-"
        cell.alignment = ALIGN_LEFT

  ws_master.freeze_panes = "A2"

  for col in ws_master.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_master.column_dimensions[col_letter].width = max(max_len + 4, 15)

  wb.save(output)
  return output.getvalue()


# ==========================================
# 4. SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
  st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=45)
  st.markdown("### **Panel Kontrol Data**")
  uploaded_file = st.file_uploader(
      "📂 Unggah File Tracker (.xlsx)", type=["xlsx"]
  )
  auto_schedule = st.toggle("Aktifkan Auto-Scheduler (3 Bulan)", value=True)
  st.markdown("---")
  st.caption(
      "Aplikasi ini otomatis menggabungkan sheet BCP dan SPS Visit, membersihkan"
      " nominal RAB, dan menghitung beban kerja tim secara real-time."
  )

# ==========================================
# 5. TAMPILAN UTAMA (DASHBOARD WEB)
# ==========================================
st.markdown(
    """
<div class="hero-banner">
    <div class="hero-title">Executive Operations & Workplan Dashboard</div>
    <div class="hero-subtitle">Monitoring Terintegrasi BCP & SPS Visit • Analisis Beban Kerja Personil (PIC) & Realisasi Anggaran</div>
</div>
""",
    unsafe_allow_html=True,
)

if uploaded_file is not None:
  file_bytes = uploaded_file.read()
  df = load_and_process_data(file_bytes, auto_schedule=auto_schedule)

  if df is not None and not df.empty:
    total_sites = len(df)
    done_sites = len(df[df["Status Progress"] == "Done (Selesai)"])
    progress_rate = (done_sites / total_sites) * 100 if total_sites > 0 else 0

    total_rab = df["Biaya"].sum()
    actual_spent = df.loc[
        df["Status Progress"] == "Done (Selesai)", "Biaya"
    ].sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Target Site", f"{total_sites} Sites")
    col2.metric(
        "Site Selesai (Done)",
        f"{done_sites} Sites",
        f"{progress_rate:.1f}% Progress",
    )
    col3.metric("Total RAB Project", f"Rp {total_rab:,.0f}")
    col4.metric("Realisasi Biaya Terserap", f"Rp {actual_spent:,.0f}")

    st.write("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs([
        "👥 Analisis Beban Kerja PIC",
        "📅 Timeline & Gantt Chart",
        "🗺️ Sebaran Area & Biaya",
        "📑 Database & Download Excel",
    ])

    with tab1:
      st.markdown(
          "#### **Beban Kerja & Jumlah Site per Personil (PIC / Engineer)**"
      )
      st.caption(
          "Mengetahui secara transparan berapa total site yang dipegang oleh"
          " masing-masing personil, status pengerjaan, serta total anggaran."
      )

      pic_summary = (
          df.groupby("PIC Engineer")
          .agg(
              Total_Sites=("Site ID", "count"),
              Done_Sites=(
                  "Status Progress",
                  lambda x: (x == "Done (Selesai)").sum(),
              ),
              Pending_Sites=(
                  "Status Progress",
                  lambda x: (x == "Pending (On-Plan)").sum(),
              ),
              Total_Budget=("Biaya", "sum"),
          )
          .reset_index()
          .sort_values(by="Total_Sites", ascending=False)
      )

      fig_pic = px.bar(
          pic_summary,
          x="PIC Engineer",
          y=["Done_Sites", "Pending_Sites"],
          title="Distribusi Jumlah Site Ditangani per Personil",
          labels={
              "value": "Jumlah Site",
              "PIC Engineer": "Nama Personil (PIC)",
              "variable": "Status Pengerjaan",
          },
          color_discrete_map={
              "Done_Sites": "#10b981",
              "Pending_Sites": "#f59e0b",
          },
          template="plotly_white",
      )
      fig_pic.update_layout(xaxis_tickangle=-45, height=450)
      st.plotly_chart(fig_pic, use_container_width=True)

      st.markdown("##### Tabel Rincian Beban Kerja Personil")
      st.dataframe(
          pic_summary.rename(columns={
              "PIC Engineer": "Nama Personil (PIC)",
              "Total_Sites": "Total Site",
              "Done_Sites": "Site Selesai",
              "Pending_Sites": "Site Pending",
              "Total_Budget": "Akumulasi Biaya (Rp)",
          }),
          use_container_width=True,
          hide_index=True,
      )

    with tab2:
      st.markdown(
          "#### **Jadwal Eksekusi Kerja 3 Bulan (90 Hari Kedepan)**"
      )
      df_plot = df.dropna(subset=["Plan Date"]).sort_values("Plan Date")

      if not df_plot.empty:
        fig_gantt = px.timeline(
            df_plot,
            x_start="Plan Date",
            x_end="Plan End",
            y="Site ID",
            color="Status Progress",
            hover_data=["Area/City", "PIC Engineer", "Source Sheet", "Biaya"],
            color_discrete_map={
                "Done (Selesai)": "#10b981",
                "Pending (On-Plan)": "#3b82f6",
            },
            template="plotly_white",
        )
        fig_gantt.update_yaxes(autorange="reversed")
        fig_gantt.update_layout(height=600, margin=dict(t=20, b=20))
        st.plotly_chart(fig_gantt, use_container_width=True)

    with tab3:
      st.markdown("#### **Analisis Wilayah & Alokasi Anggaran**")
      col_a, col_b = st.columns(2)

      with col_a:
        area_count = (
            df.groupby(["Area/City", "Status Progress"])
            .size()
            .reset_index(name="Jumlah")
        )
        fig_area = px.bar(
            area_count,
            x="Area/City",
            y="Jumlah",
            color="Status Progress",
            title="Volume Pekerjaan per Kabupaten / Area",
            color_discrete_map={
                "Done (Selesai)": "#10b981",
                "Pending (On-Plan)": "#f59e0b",
            },
            template="plotly_white",
        )
        fig_area.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_area, use_container_width=True)

      with col_b:
        area_budget = df.groupby("Area/City")["Biaya"].sum().reset_index()
        fig_pie = px.pie(
            area_budget,
            values="Biaya",
            names="Area/City",
            hole=0.4,
            title="Proporsi Anggaran Berdasarkan Wilayah",
            template="plotly_white",
        )
        fig_pie.update_traces(
            textposition="inside", textinfo="percent+label"
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with tab4:
      st.markdown("#### **Master Data Rekapitulasi & Unduh Laporan**")

      df_display = df.copy()
      df_display["Date Actual"] = (
          df_display["Date Actual"]
          .dt.strftime("%d-%b-%Y")
          .fillna("Belum Selesai")
      )
      df_display["Plan Date"] = (
          df_display["Plan Date"].dt.strftime("%d-%b-%Y").fillna("-")
      )
      df_display["Plan End"] = (
          df_display["Plan End"].dt.strftime("%d-%b-%Y").fillna("-")
      )

      cols = [
          "Site ID",
          "Source Sheet",
          "Area/City",
          "PIC Engineer",
          "Kategori",
          "Status Progress",
          "Plan Date",
          "Date Actual",
          "Biaya",
      ]
      available_cols = [c for c in cols if c in df_display.columns]

      st.dataframe(
          df_display[available_cols], use_container_width=True, height=450
      )

      # Panggil fungsi generator Excel eksekutif
      excel_bytes = convert_df_to_excel(df)
      st.download_button(
          label="📥 Download Executive Report Excel (.xlsx)",
          data=excel_bytes,
          file_name=(
              f"Executive_Report_BCP_SPS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
          ),
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )
else:
  st.info(
      "👈 **Silakan unggah file Excel Anda melalui panel di sebelah kiri**"
      " untuk menampilkan seluruh analisis dan grafik secara otomatis."
  )
