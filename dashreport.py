import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
import io

# CONFIG
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")

# --- LOGO LOADING ---
def load_base64_image_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        return base64.b64encode(response.content).decode()
    return None

logo_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
logo_base64 = load_base64_image_from_url(logo_url)
if logo_base64:
    st.markdown(f"""
        <div style="width: 100%; background-color: #3FA0A3; padding: 10px 30px; display: flex; align-items: center; justify-content: center; margin-bottom: 20px;">
            <img src="data:image/png;base64,{logo_base64}" style="height: 52px; margin-right: 15px; border: 2px solid white;">
            <h1 style="color: white; font-size: 26px;">Muthokinju Paints Sales Dashboard</h1>
        </div>
    """, unsafe_allow_html=True)

# --- LOAD DATA ---
url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
sales = pd.read_excel(url, sheet_name="CY", engine="openpyxl")
targets = pd.read_excel(url, sheet_name="TARGETS", engine="openpyxl")
prev_year_sales = pd.read_excel(url, sheet_name="PY", engine="openpyxl")

sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year_sales.columns = prev_year_sales.columns.str.lower()

sales['date'] = pd.to_datetime(sales['date'])
prev_year_sales['date'] = pd.to_datetime(prev_year_sales['date'])

for df in [sales, targets, prev_year_sales]:
    df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum()
targets_agg.rename(columns={'amount': 'monthly_target'}, inplace=True)

# --- WORKING DAYS FUNCTION ---
def working_days_excl_sundays(start_date, end_date):
    return len([d for d in pd.date_range(start=start_date, end=end_date) if d.weekday() != 6])

# --- FILTERS ---
clusters = sales["Cluster"].dropna().unique()
branches = sales["branch"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min = sales["date"].min()
date_max = sales["date"].max()

col1, col2, col3 = st.columns(3)
with col1:
    selected_cluster = st.selectbox("Cluster", ["All"] + list(clusters))
with col2:
    selected_branch = st.selectbox("Branch", ["All"] + list(branches))
with col3:
    selected_category = st.selectbox("Category", ["All"] + list(categories))

col_from, col_to = st.columns(2)
with col_from:
    start_date = st.date_input("From", value=date_min, min_value=date_min, max_value=date_max)
with col_to:
    end_date = st.date_input("To", value=date_max, min_value=date_min, max_value=date_max)

# --- FILTER DATA ---
filtered = sales.copy()
if selected_cluster != "All":
    filtered = filtered[filtered["Cluster"] == selected_cluster]
if selected_branch != "All":
    filtered = filtered[filtered["branch"] == selected_branch]
if selected_category != "All":
    filtered = filtered[filtered["category1"] == selected_category]

filtered = filtered[(filtered["date"] >= pd.to_datetime(start_date)) & (filtered["date"] <= pd.to_datetime(end_date))]

# --- AGGREGATIONS ---
end_dt = pd.to_datetime(end_date)
days_passed = working_days_excl_sundays(start_date, end_date)
month_start = pd.Timestamp(end_dt.year, end_dt.month, 1)
month_end = pd.Timestamp(end_dt.year, end_dt.month, end_dt.days_in_month)
total_working_days = working_days_excl_sundays(month_start, month_end)

mtd_agg = filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'mtd_achieved'})
daily_achieved = filtered[filtered['date'] == end_dt].groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'daily_achieved'})

prev_year_filtered = prev_year_sales[
    (prev_year_sales['date'] >= pd.Timestamp(end_dt.year - 1, end_dt.month, 1)) &
    (prev_year_sales['date'] <= pd.Timestamp(end_dt.year - 1, end_dt.month, end_dt.days_in_month))
]
pym_agg = prev_year_filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})

df = mtd_agg.merge(daily_achieved, on=['branch', 'category1'], how='left') \
            .merge(targets_agg, on=['branch', 'category1'], how='left') \
            .merge(pym_agg, on=['branch', 'category1'], how='left')
df.fillna(0, inplace=True)

# --- CALCULATIONS ---
df['daily_tgt'] = df['monthly_target'] / total_working_days
df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt'] == 0, 0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'])
df['mtd_tgt'] = df['daily_tgt'] * days_passed
df['mtd_var'] = np.where(df['mtd_tgt'] == 0, 0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'])
df['cm'] = df['mtd_achieved']
df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target'] == 0, 0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'])
df['projected_landing'] = np.where(days_passed == 0, 0, (df['mtd_achieved'] / days_passed) * total_working_days)
df['cm_vs_pym'] = np.where(df['pym'] == 0, 0, (df['cm'] - df['pym']) / df['pym'])

# --- RENAME FOR DISPLAY ---
df.rename(columns={
    'monthly_target': 'Monthly TGT',
    'daily_tgt': 'Daily Tgt',
    'daily_achieved': 'Daily Achieved',
    'achieved_vs_daily_tgt': 'Achieved vs Daily Tgt',
    'mtd_tgt': 'MTD TGT',
    'mtd_achieved': 'MTD Act.',
    'mtd_var': 'MTD Var',
    'cm': 'CM',
    'achieved_vs_monthly_tgt': 'Achieved VS Monthly tgt',
    'projected_landing': 'Projected landing',
    'pym': 'PYM',
    'cm_vs_pym': 'CM VS PYM'
}, inplace=True)

# --- TOTALS ROW ---
def safe_div(n, d): return (n - d) / d if d != 0 else 0
totals_row = {
    'branch': 'Totals',
    'category1': '',
    'Monthly TGT': df['Monthly TGT'].sum(),
    'Daily Tgt': df['Daily Tgt'].sum(),
    'Daily Achieved': df['Daily Achieved'].sum(),
    'Achieved vs Daily Tgt': safe_div(df['Daily Achieved'].sum(), df['Daily Tgt'].sum()),
    'MTD TGT': df['MTD TGT'].sum(),
    'MTD Act.': df['MTD Act.'].sum(),
    'MTD Var': safe_div(df['MTD Act.'].sum(), df['MTD TGT'].sum()),
    'CM': df['CM'].sum(),
    'Achieved VS Monthly tgt': safe_div(df['MTD Act.'].sum(), df['Monthly TGT'].sum()),
    'Projected landing': df['Projected landing'].sum(),
    'PYM': df['PYM'].sum(),
    'CM VS PYM': safe_div(df['CM'].sum(), df['PYM'].sum())
}
df = pd.concat([df, pd.DataFrame([totals_row])], ignore_index=True)

# --- FORMAT % COLUMNS ---
percent_cols = ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']
for col in percent_cols:
    df[col] = (df[col].astype(float) * 100).round(1).astype(str) + '%'

# --- COLUMN ORDER ---
order = ["branch", "category1", "Monthly TGT", "Daily Tgt", "Daily Achieved", "Achieved vs Daily Tgt",
         "MTD TGT", "MTD Act.", "MTD Var", "CM", "Achieved VS Monthly tgt", "Projected landing", "PYM", "CM VS PYM"]
df = df[order]

# --- CHART ---
st.markdown("### 📊 Sales vs Monthly Target (MTD)")
chart_df = df[df['branch'] != 'Totals']
x = chart_df.apply(lambda r: f"{r['branch']} - {r['category1']}", axis=1)

fig = go.Figure()
fig.add_trace(go.Bar(x=x, y=chart_df['MTD Act.'], name='MTD Achieved', marker_color='orange'))
fig.add_trace(go.Bar(x=x, y=chart_df['Monthly TGT'], name='Monthly Target', marker_color='steelblue'))
fig.update_layout(barmode='group', height=500, margin=dict(b=150), xaxis_tickangle=-45)
st.plotly_chart(fig, use_container_width=True)

# --- STYLED TABLE ---
format_dict = {
    'Monthly TGT': "{:,.1f}", 'Daily Tgt': "{:,.1f}", 'Daily Achieved': "{:,.1f}",
    'MTD TGT': "{:,.1f}", 'MTD Act.': "{:,.1f}", 'CM': "{:,.1f}",
    'Projected landing': "{:,.1f}", 'PYM': "{:,.1f}"
}

def highlight_totals(row): return ['background-color: #b2dfdb; font-weight: bold;' if row['branch'] == 'Totals' else '' for _ in row]
def highlight_percent(val):
    try:
        v = float(val.strip('%'))
        if v > 0: return 'background-color: #d0f0c0'
        elif v < 0: return 'background-color: #ffc0cb'
    except: pass
    return ''

styled = df.style.format(format_dict)\
    .apply(highlight_totals, axis=1)\
    .applymap(highlight_percent, subset=percent_cols)

st.markdown("### 📋 Summary Table")
st.dataframe(styled, use_container_width=True)
