import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
import io
from datetime import datetime

# Optional import for exporting table as image
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

# ========== LOGO & HEADER ==========
def load_base64_image_from_url(url):
    resp = requests.get(url)
    return base64.b64encode(resp.content).decode() if resp.status_code == 200 else None

logo_b64 = load_base64_image_from_url(
    "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
)
if logo_b64:
    st.markdown(f"""
    <div style="display:flex; align-items:center; margin-bottom:20px;">
        <img src="data:image/png;base64,{logo_b64}" style="height:60px; margin-right:15px;" />
        <h1 style="margin:0;">Muthokinju Paints Sales Dashboard</h1>
    </div>
    """, unsafe_allow_html=True)
else:
    st.error("⚠️ Failed to load logo.")

# ========== LOAD DATA ==========
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year_sales = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"⚠️ Failed to load Excel data: {e}")
    st.stop()

# ========== CLEAN 'amount' COLUMNS IF PRESENT ==========
for df in (sales, targets, prev_year_sales):
    if 'amount' in df.columns:
        df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

# ========== STANDARDIZE COLUMNS & PARSE DATES ==========
sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year_sales.columns = prev_year_sales.columns.str.lower()

sales['date'] = pd.to_datetime(sales['date'])
prev_year_sales['date'] = pd.to_datetime(prev_year_sales['date'])

# ========== AGGREGATE TARGETS ==========
targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'monthly_target'})

# ========== HELPER – WORKING DAYS EXCLUDING SUNDAYS ==========
def working_days_excl_sundays(start_date, end_date):
    dates = pd.date_range(start=start_date, end=end_date)
    return len([d for d in dates if d.weekday() != 6])

# ========== FILTER UI ==========
clusters = sales['Cluster'].dropna().unique().tolist()
branches = sales['branch'].dropna().unique().tolist()
categories = sales['category1'].dropna().unique().tolist()

default_start = datetime(2024, 1, 1)
default_end = datetime.today()

c1, c2, c3, c4 = st.columns([1,1,1,2])
with c1:
    sel_cluster = st.selectbox("Cluster", ["All"] + clusters)
with c2:
    sel_branch = st.selectbox("Branch", ["All"] + branches)
with c3:
    sel_category = st.selectbox("Category", ["All"] + categories)
with c4:
    date_range = st.date_input("Date Range", value=(default_start, default_end),
                               min_value=default_start, max_value=default_end)

if not (isinstance(date_range, tuple) and len(date_range) == 2):
    st.error("Please select a valid start and end date.")
    st.stop()

start_date, end_date = date_range

# ========== FILTER DATA ==========
filtered = sales.copy()
if sel_cluster != "All":
    filtered = filtered[filtered["Cluster"] == sel_cluster]
if sel_branch != "All":
    filtered = filtered[filtered["branch"] == sel_branch]
if sel_category != "All":
    filtered = filtered[filtered["category1"] == sel_category]
filtered = filtered[(filtered["date"] >= pd.to_datetime(start_date)) & (filtered["date"] <= pd.to_datetime(end_date))]

# ========== CALCULATE WORKING DAYS ==========
end_ts = pd.Timestamp(end_date)
month_start = pd.Timestamp(end_ts.year, end_ts.month, 1)
# Fix: compute days in month correctly via pandas Timestamp
month_end = pd.Timestamp(end_ts.year, end_ts.month, pd.Timestamp(end_ts).days_in_month)

work_days_in_month = working_days_excl_sundays(month_start, month_end)
work_days_done = working_days_excl_sundays(start_date, end_date)

# ========== BUILD DATAFRAME WITH METRICS ==========
if not filtered.empty:
    days_passed = working_days_excl_sundays(start_date, end_date)

    mtd_agg = filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'mtd_achieved'})
    daily_achieved = filtered[filtered['date'] == end_ts].groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'daily_achieved'})

    prev_year_filtered = prev_year_sales[
        (prev_year_sales['date'] >= pd.Timestamp(end_ts.year - 1, end_ts.month, 1)) &
        (prev_year_sales['date'] <= pd.Timestamp(end_ts.year - 1, end_ts.month, pd.Timestamp(end_ts).days_in_month))
    ]
    pym_agg = prev_year_filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})

    df = (mtd_agg.merge(daily_achieved, on=['branch', 'category1'], how='left')
           .merge(targets_agg, on=['branch', 'category1'], how='left')
           .merge(pym_agg, on=['branch', 'category1'], how='left'))
    df.fillna(0, inplace=True)

    df['daily_tgt'] = df['monthly_target'] / work_days_in_month
    df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt'] == 0, 0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'])
    df['mtd_tgt'] = df['daily_tgt'] * days_passed
    df['mtd_var'] = np.where(df['mtd_tgt'] == 0, 0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'])
    df['cm'] = df['mtd_achieved']
    df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target'] == 0, 0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'])
    df['projected_landing'] = np.where(days_passed == 0, 0, (df['mtd_achieved'] / days_passed) * work_days_in_month)
    df['cm_vs_pym'] = np.where(df['pym'] == 0, 0, (df['cm'] - df['pym']) / df['pym'])

    df.rename(columns={
        'monthly_target': 'Monthly TGT', 'daily_tgt': 'Daily Tgt', 'daily_achieved': 'Daily Achieved',
        'mtd_tgt': 'MTD TGT', 'mtd_achieved': 'MTD Act.', 'mtd_var': 'MTD Var', 'cm': 'CM',
        'achieved_vs_monthly_tgt': 'Achieved VS Monthly tgt', 'projected_landing': 'Projected landing',
        'pym': 'PYM', 'cm_vs_pym': 'CM VS PYM'
    }, inplace=True)

    total_vals = df[df['branch'] != 'Totals']
    tot = {
        'branch': 'Totals', 'category1': '',
        'Monthly TGT': total_vals['Monthly TGT'].sum(),
        'Daily Tgt': total_vals['Daily Tgt'].sum(),
        'Daily Achieved': total_vals['Daily Achieved'].sum(),
        'Achieved vs Daily Tgt': ((total_vals['Daily Achieved'].sum() - total_vals['Daily Tgt'].sum()) / total_vals['Daily Tgt'].sum()) if total_vals['Daily Tgt'].sum() else 0,
        'MTD TGT': total_vals['MTD TGT'].sum(),
        'MTD Act.': total_vals['MTD Act.'].sum(),
        'MTD Var': ((total_vals['MTD Act.'].sum() - total_vals['MTD TGT'].sum()) / total_vals['MTD TGT'].sum()) if total_vals['MTD TGT'].sum() else 0,
        'CM': total_vals['CM'].sum(),
        'Achieved VS Monthly tgt': (total_vals['MTD Act.'].sum() / total_vals['Monthly TGT'].sum()) if total_vals['Monthly TGT'].sum() else 0,
        'Projected landing': total_vals['Projected landing'].sum(),
        'PYM': total_vals['PYM'].sum(),
        'CM VS PYM': ((total_vals['CM'].sum() - total_vals['PYM'].sum()) / total_vals['PYM'].sum()) if total_vals['PYM'].sum() else 0
    }
    df = pd.concat([df[df['branch'] != 'Totals'], pd.DataFrame([tot])], ignore_index=True)
    for col in ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']:
        df[col] = (df[col].astype(float) * 100).round(1).astype(str) + '%'

    desired_order = [
        "branch", "category1", "Monthly TGT", "Daily Tgt", "Daily Achieved", "Achieved vs Daily Tgt",
        "MTD TGT", "MTD Act.", "MTD Var", "CM", "Achieved VS Monthly tgt", "Projected landing", "PYM", "CM VS PYM"
    ]
    df = df[desired_order]
else:
    df = pd.DataFrame()

# ========== KPI CARDS ==========
k1, k2, k3, k4 = st.columns(4)
k1.markdown(f'<div class="kpi"><h3>Work Days in Month</h3><p>{work_days_in_month}</p></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi"><h3>Days Worked</h3><p>{work_days_done}</p></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi"><h3>MTD Achieved</h3><p>{filtered["amount"].sum() if not filtered.empty else 0:,.1f}</p></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi"><h3>Monthly Target</h3><p>{targets_agg["monthly_target"].sum():,.1f}</p></div>', unsafe_allow_html=True)

# ========== CHART ==========
st.markdown("### Sales vs Monthly Target (MTD)")
fig = go.Figure()
if not df.empty:
    chart_df = df[df['branch'] != 'Totals']
    xlabels = chart_df.apply(lambda r: f"{r['branch']} - {r['category1']}", axis=1)
    fig.add_trace(go.Bar(x=xlabels, y=chart_df['MTD Act.'], name='MTD Achieved', marker_color='orange'))
    fig.add_trace(go.Bar(x=xlabels, y=chart_df['Monthly TGT'], name='Monthly Target', marker_color='steelblue'))
fig.update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150),
                  legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))

st.plotly_chart(fig, use_container_width=True, config={
    "displaylogo": False,
    "displayModeBar": True,
    "modeBarButtonsToRemove": ["zoom", "pan", "select2d", "lasso2d", "autoScale2d", "resetScale2d"],
    "modeBarButtonsToAdd": ["toImage"]
})

# ========== TABLE & DOWNLOADS ==========
st.markdown("### Data Table")
with st.container():
    st.markdown('<div class="table-scroll-area">', unsafe_allow_html=True)
    if not df.empty:
        st.write(df.style.format({col: "{:,.1f}" for col in [
            'Monthly TGT', 'Daily Tgt', 'Daily Achieved', 'MTD TGT', 'MTD Act.', 'CM', 'Projected landing', 'PYM']}))
    else:
        st.write("No records to display for the selected filters.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.download_button("Download CSV", data=df.to_csv(index=False).encode('utf-8'), file_name="sales_data.csv", mime="text/csv")

    if HAS_DFI and not df.empty:
        buf = io.BytesIO()
        dfi.export(df, buf, table_conversion="matplotlib")
        buf.seek(0)
        st.download_button("Download Table as Image (PNG)", data=buf.read(), file_name="sales_data.png", mime="image/png")
    elif not HAS_DFI:
        st.info("Install `dataframe_image` to enable table image download.")

if df.empty:
    st.warning("No data found for the selected filters or date range.")
