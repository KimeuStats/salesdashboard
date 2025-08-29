import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
import io
from datetime import datetime

# Optional for exporting table as image
try:
    import dataframe_image as dfi
    HAS_DFI = True
except ModuleNotFoundError:
    HAS_DFI = False

# ========== PAGE CONFIG ==========
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")

# ========== STYLES ==========
st.markdown("""
<style>
    .main .block-container { max-width: 1400px; padding: 2rem; margin: auto; }
    .table-scroll-area { overflow-x: auto; border: 1px solid #ccc; padding: 10px; }
    .kpi { background-color: #f0f0f0; padding: 15px; border-radius: 8px; text-align: center; }
    .kpi h3 { margin: 0; }
    .kpi p { margin: 5px 0 0; font-size: 24px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ========== LOGO + HEADER ==========
def load_image_b64(url):
    resp = requests.get(url)
    return base64.b64encode(resp.content).decode() if resp.status_code == 200 else None

logo_b64 = load_image_b64("https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png")
if logo_b64:
    st.markdown(f"""
    <div style="display:flex; align-items:center; margin-bottom:20px;">
      <img src="data:image/png;base64,{logo_b64}" style="height:60px; margin-right:15px;" />
      <h1 style="margin:0;">Muthokinju Paints Sales Dashboard</h1>
    </div>""", unsafe_allow_html=True)
else:
    st.error("⚠️ Failed to load logo image.")

# ========== LOAD DATA ==========
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

# Clean amounts
for df in (sales, targets, prev_year):
    if 'amount' in df.columns:
        df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

# Normalize columns
sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year.columns = prev_year.columns.str.lower()

sales['date'] = pd.to_datetime(sales['date'])
prev_year['date'] = pd.to_datetime(prev_year['date'])

# Aggregate targets
targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'Monthly Target'})

# ========== HELPER FUNCTION ==========
def working_days_excl_sundays(start_date, end_date):
    dates = pd.date_range(start=start_date, end=end_date)
    return len([d for d in dates if d.weekday() != 6])

# ========== FILTER UI ==========
clusters = sales['Cluster'].dropna().unique().tolist()
branches = sales['branch'].dropna().unique().tolist()
categories = sales['category1'].dropna().unique().tolist()

c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
with c1:
    sel_cluster = st.selectbox("Cluster", ["All"] + clusters)
with c2:
    sel_branch = st.selectbox("Branch", ["All"] + branches)
with c3:
    sel_category = st.selectbox("Category", ["All"] + categories)
with c4:
    date_range = st.date_input("Date Range", value=(datetime(2024, 1, 1), datetime.today()),
                               min_value=datetime(2024, 1, 1), max_value=datetime.today())

if not (isinstance(date_range, tuple) and len(date_range) == 2):
    st.error("Please select a valid date range.")
    st.stop()

start_date, end_date = date_range

# Filter data
filtered = sales.copy()
if sel_cluster != "All":
    filtered = filtered[filtered["Cluster"] == sel_cluster]
if sel_branch != "All":
    filtered = filtered[filtered["branch"] == sel_branch]
if sel_category != "All":
    filtered = filtered[filtered["category1"] == sel_category]
filtered = filtered[(filtered["date"] >= start_date) & (filtered["date"] <= end_date)]

# Calculate working days
end_ts = pd.Timestamp(end_date)
days_in_month = pd.Timestamp(end_ts).days_in_month
month_start = pd.Timestamp(end_ts.year, end_ts.month, 1)
month_end = pd.Timestamp(end_ts.year, end_ts.month, days_in_month)

work_days_in_month = working_days_excl_sundays(month_start, month_end)
work_days_done = working_days_excl_sundays(start_date, end_date)

# ========== BUILD METRIC DATAFRAME ==========
if not filtered.empty:
    days_passed = work_days_done
    # Aggregates...
    mtd = filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'MTD Act.'})
    daily = filtered[filtered['date'] == end_ts].groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'Daily Achieved'})
    prev_filtered = prev_year[(prev_year['date'] >= pd.Timestamp(end_ts.year - 1, end_ts.month, 1)) &
                              (prev_year['date'] <= pd.Timestamp(end_ts.year - 1, end_ts.month, days_in_month))]
    py = prev_filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'PYM'})

    df = mtd.merge(daily, on=['branch', 'category1'], how='left').merge(targets_agg, on=['branch', 'category1'], how='left').merge(py, on=['branch', 'category1'], how='left')
    df.fillna(0, inplace=True)

    # Calculations
    df['Daily Tgt'] = df['Monthly Target'] / work_days_in_month
    df['MTD TGT'] = df['Daily Tgt'] * days_passed
    df['MTD Var'] = np.where(df['MTD TGT'] == 0, 0, (df['MTD Act.'] - df['MTD TGT']) / df['MTD TGT'])
    df['Achieved VS Monthly tgt'] = np.where(df['Monthly Target'] == 0, 0, (df['MTD Act.'] - df['Monthly Target']) / df['Monthly Target'])
    df['Projected landing'] = np.where(days_passed == 0, 0, (df['MTD Act.'] / days_passed) * work_days_in_month)
    df['CM VS PYM'] = np.where(df['PYM'] == 0, 0, (df['MTD Act.'] - df['PYM']) / df['PYM'])

    # Format columns
    pct_cols = ['MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']
    for col in pct_cols:
        df[col] = (df[col] * 100).round(1).astype(str) + '%'

    # Totals row
    total_vals = df[df['branch'] != 'Totals']
    total_row = {
        'branch': 'Totals', 'category1':'', 'Monthly Target': total_vals['Monthly Target'].sum(),
        'MTD Act.': total_vals['MTD Act.'].sum(), 'Daily Tgt': total_vals['Daily Tgt'].sum(),
        'MTD TGT': total_vals['MTD TGT'].sum(), 'MTD Var': '',
        'Achieved VS Monthly tgt': '', 'Projected landing': total_vals['Projected landing'].sum(),
        'PYM': total_vals['PYM'].sum(), 'CM VS PYM': ''
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
else:
    df = pd.DataFrame()

# ========== DISPLAY KPIs ==========
k1, k2, k3, k4 = st.columns(4)
k1.markdown(f'<div class="kpi"><h3>Work Days in Month</h3><p>{work_days_in_month}</p></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi"><h3>Days Worked</h3><p>{work_days_done}</p></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi"><h3>MTD Achieved</h3><p>{filtered["amount"].sum() if not filtered.empty else 0:,.1f}</p></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi"><h3>Monthly Target</h3><p>{df["Monthly Target"].sum() if not df.empty else 0:,.1f}</p></div>', unsafe_allow_html=True)

# ========== CHART ==========
st.markdown("### Sales vs Monthly Target (MTD)")
fig = go.Figure()
if not df.empty:
    chart_df = df[df['branch'] != 'Totals']
    x_labels = chart_df.apply(lambda r: f"{r['branch']} - {r['category1']}", axis=1)
    fig.add_trace(go.Bar(x=x_labels, y=chart_df['MTD Act.'], name='MTD Achieved', marker_color='orange'))
    fig.add_trace(go.Bar(x=x_labels, y=chart_df['Monthly Target'], name='Monthly Target', marker_color='steelblue'))
fig.update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150),
                  legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))

st.plotly_chart(fig, use_container_width=True, config={
    "displaylogo": False,
    "displayModeBar": True,
    "modeBarButtonsToRemove": ["zoom", "pan", "select2d", "lasso2d", "autoScale2d", "resetScale2d"],
    "modeBarButtonsToAdd": ["toImage"]
})

# ========== TABLE & DOWNLOAD OPTIONS ==========
st.markdown("### Data Table")
with st.container():
    st.markdown('<div class="table-scroll-area">', unsafe_allow_html=True)
    if not df.empty:
        st.write(df.style.format({col: "{:,.1f}" for col in df.columns if col not in pct_cols}))
    else:
        st.write("No data available for the selected range.")
    st.markdown('</div>', unsafe_allow_html=True)

    # CSV Download
    st.download_button("Download CSV", data=df.to_csv(index=False).encode('utf-8'), file_name="data.csv", mime="text/csv")

    # Image Download if supported
    if HAS_DFI and not df.empty:
        buf = io.BytesIO()
        dfi.export(df, buf, table_conversion="matplotlib")
        buf.seek(0)
        st.download_button("Download Table as PNG", data=buf.read(), file_name="data_table.png", mime="image/png")
    elif not HAS_DFI:
        st.info("Install `dataframe_image` for table image downloads.")

if df.empty:
    st.warning("No data found for selected filters or date range.")
