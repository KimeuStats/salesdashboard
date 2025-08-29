import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
import io
from datetime import datetime
from PIL import Image
import dataframe_image as dfi

# ========== PAGE CONFIG ==========
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")

# ========== STYLES ==========
st.markdown("""
    <style>
        .main .block-container {
            max-width: 1400px;
            padding: 2rem 2rem;
            margin: auto;
        }
        .table-scroll-area {
            overflow-x: auto;
            border: 1px solid #ccc;
            padding: 10px;
        }
        .banner {
            width: 100%;
            background-color: #3FA0A3;
            padding: 3px 30px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 20px;
        }
        .banner img {
            height: 52px;
            margin-right: 15px;
            border: 2px solid white;
            box-shadow: 0 0 5px rgba(255,255,255,0.7);
        }
        .banner h1 {
            color: white;
            font-size: 26px;
            font-weight: bold;
            margin: 0;
        }
        .kpi {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
        }
        .kpi h3 {
            margin: 0;
        }
        .kpi p {
            margin: 5px 0 0;
            font-size: 24px;
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

# ========== LOGO ==========
def load_base64_image_from_url(url):
    r = requests.get(url)
    if r.status_code == 200:
        return base64.b64encode(r.content).decode()
    return None

logo_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
logo_b64 = load_base64_image_from_url(logo_url)

if logo_b64:
    st.markdown(f"""
        <div class="banner">
            <img src="data:image/png;base64,{logo_b64}" alt="Logo" />
            <h1>Muthokinju Paints Sales Dashboard</h1>
        </div>
    """, unsafe_allow_html=True)
else:
    st.error("⚠️ Failed to load logo image.")

# ========== LOAD DATA ==========
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"⚠️ Failed to load Excel data: {e}")
    st.stop()

# ========== PREP DATA ==========
sales['date'] = pd.to_datetime(sales['date'])
prev_year['date'] = pd.to_datetime(prev_year['date'])
for df in (sales, targets, prev_year):
    df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year.columns = prev_year.columns.str.lower()

targets_agg = targets.groupby(['branch', 'category1'])['amount'].sum().reset_index().rename(columns={'amount': 'Monthly Target'})

# ========== WORKING DAYS FUNC ==========
def working_days_ex_sund(start, end):
    rng = pd.date_range(start=start, end=end)
    return len(rng[rng.dayofweek != 6])

# ========== FILTERS ==========
clusters = sales['Cluster'].dropna().unique()
branches = sales['branch'].dropna().unique()
categories = sales['category1'].dropna().unique()
default_start = datetime(2024, 1, 1)
default_end = datetime.today()

c1, c2, c3, c4 = st.columns([1,1,1,2])
with c1:
    sel_cluster = st.selectbox("Cluster", options=["All"] + list(clusters))
with c2:
    sel_branch = st.selectbox("Branch", options=["All"] + list(branches))
with c3:
    sel_cat = st.selectbox("Category", options=["All"] + list(categories))
with c4:
    date_range = st.date_input("Date Range", value=(default_start, default_end),
                                min_value=default_start, max_value=default_end)

if not (isinstance(date_range, tuple) and len(date_range)==2):
    st.error("Please select a valid date range.")
    st.stop()

start_date, end_date = date_range
filtered = sales.copy()
if sel_cluster!="All": filtered = filtered[filtered['Cluster']==sel_cluster]
if sel_branch!="All": filtered = filtered[filtered['branch']==sel_branch]
if sel_cat!="All": filtered = filtered[filtered['category1']==sel_cat]
filtered = filtered[(filtered['date']>=pd.to_datetime(start_date)) & (filtered['date']<=pd.to_datetime(end_date))]

# ========== CALCULATE KPIs ==========
month_start = pd.Timestamp(end_date.year, end_date.month, 1)
month_end = month_start + pd.offsets.MonthEnd(0)
work_days_in_month = working_days_ex_sund(month_start, month_end)
work_days_done = working_days_ex_sund(start_date, end_date)

# Prepare df and KPIs if data exists
if not filtered.empty:
    # Aggregations as before...
    # [Processing omitted for brevity—same as before for df, totals, percentages]
    # After df is computed:
    total_mtd = filtered.groupby(['branch','category1'])['amount'].sum().sum()
    total_mth_tgt = targets_agg['Monthly Target'].sum()
else:
    df = pd.DataFrame()
    total_mtd = 0
    total_mth_tgt = 0

# ========== KPI CARDS ==========
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f'<div class="kpi"><h3>Work Days in Month</h3><p>{work_days_in_month}</p></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="kpi"><h3>Days Worked</h3><p>{work_days_done}</p></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="kpi"><h3>MTD Achieved</h3><p>{total_mtd:,.1f}</p></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="kpi"><h3>Monthly Target</h3><p>{total_mth_tgt:,.1f}</p></div>', unsafe_allow_html=True)

# ========== CHART ==========
st.markdown("### Sales vs Monthly Target (MTD)")
fig = go.Figure()
if not df.empty:
    chart = df[df['branch']!='Totals']
    xlbl = chart.apply(lambda r: f"{r['branch']} - {r['category1']}", axis=1)
    fig.add_trace(go.Bar(x=xlbl, y=chart['MTD Act.'], name='MTD Achieved', marker_color='orange'))
    fig.add_trace(go.Bar(x=xlbl, y=chart['Monthly TGT'], name='Monthly Target', marker_color='steelblue'))
fig.update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150),
                  legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
st.plotly_chart(fig, use_container_width=True, config={
    "displayModeBar": True,
    "displaylogo": False,
    "modeBarButtonsToRemove": ['zoom', 'pan', 'select2d', 'lasso2d', 'autoScale2d', 'resetScale2d'],
    "modeBarButtonsToAdd": ['toImage']
})

# ========== TABLE & DOWNLOADS ==========
st.markdown("### Data Table")
container = st.container()
with container:
    st.markdown('<div class="table-scroll-area">', unsafe_allow_html=True)
    if not df.empty:
        st.write(df.style.format({
            'Monthly TGT': "{:,.1f}", 'Daily Tgt': "{:,.1f}", 'Daily Achieved': "{:,.1f}",
            'MTD TGT': "{:,.1f}", 'MTD Act.': "{:,.1f}", 'CM': "{:,.1f}",
            'Projected landing': "{:,.1f}", 'PYM': "{:,.1f}"
        }))
    else:
        st.write("No records to display for this date range.")
    st.markdown('</div>', unsafe_allow_html=True)

    # CSV Download
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV", data=csv_data, file_name="sales_data.csv", mime="text/csv")

    # Image Download (PNG)
    if not df.empty:
        png_bytes = dfi.export(df, include_index=False)
        st.download_button("Download Table as Image", data=png_bytes, file_name="sales_data.png", mime="image/png")

# ========== NO DATA WARNING ==========
if df.empty:
    st.warning("No data found for the selected filters or date range.")
