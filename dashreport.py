import streamlit as st
import pandas as pd
import numpy as np
import requests
import base64
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import openpyxl

# === PAGE CONFIG ===
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")

# === LOAD DATA ===
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year_sales = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"Failed to load Excel data: {e}")
    st.stop()

# Normalize column names
sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year_sales.columns = prev_year_sales.columns.str.lower()

sales['date'] = pd.to_datetime(sales['date'])
prev_year_sales['date'] = pd.to_datetime(prev_year_sales['date'])

for df in [sales, targets, prev_year_sales]:
    df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'monthly_target'})

# === HELPER FUNCTION ===
def working_days_excl_sundays(start_date, end_date):
    return len([d for d in pd.date_range(start=start_date, end=end_date) if d.weekday() != 6])

# === FILTERS ===
clusters = sales["Cluster"].dropna().unique()
branches = sales["branch"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min, date_max = sales["date"].min(), sales["date"].max()

col1, col2, col3 = st.columns(3)
with col1:
    selected_cluster = st.selectbox("Cluster", options=["All"] + list(clusters))
with col2:
    selected_branch = st.selectbox("Branch", options=["All"] + list(branches))
with col3:
    selected_category = st.selectbox("Category", options=["All"] + list(categories))

col_from, col_to = st.columns(2)
with col_from:
    st.markdown("<div style='background-color:#7b38d8; color:white; padding:8px; font-weight:bold;'>From</div>", unsafe_allow_html=True)
    start_date = st.date_input("", value=date_min, min_value=date_min, max_value=date_max, key="from_date")
with col_to:
    st.markdown("<div style='background-color:#7b38d8; color:white; padding:8px; font-weight:bold;'>To</div>", unsafe_allow_html=True)
    end_date = st.date_input("", value=date_max, min_value=date_min, max_value=date_max, key="to_date")

# === APPLY FILTERS ===
filtered = sales.copy()
if selected_cluster != "All":
    filtered = filtered[filtered["Cluster"] == selected_cluster]
if selected_branch != "All":
    filtered = filtered[filtered["branch"] == selected_branch]
if selected_category != "All":
    filtered = filtered[filtered["category1"] == selected_category]
filtered = filtered[(filtered["date"] >= pd.to_datetime(start_date)) & (filtered["date"] <= pd.to_datetime(end_date))]

if filtered.empty:
    st.warning("No sales data found for the selected filters or date range.")
    st.stop()

# === CALCULATIONS ===
end_dt = pd.to_datetime(end_date)
month_start = pd.Timestamp(end_dt.year, end_dt.month, 1)
month_end = pd.Timestamp(end_dt.year, end_dt.month, end_dt.days_in_month)

days_worked = working_days_excl_sundays(month_start, end_dt)
total_working_days = working_days_excl_sundays(month_start, month_end)

mtd_agg = filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'mtd_achieved'})
daily_achieved = filtered[filtered['date'] == end_dt].groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'daily_achieved'})

prev_year_filtered = prev_year_sales[
    (prev_year_sales['date'] >= pd.Timestamp(end_dt.year - 1, end_dt.month, 1)) &
    (prev_year_sales['date'] <= pd.Timestamp(end_dt.year - 1, end_dt.month, end_dt.days_in_month))
]
pym_agg = prev_year_filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})

df = (mtd_agg.merge(daily_achieved, on=['branch', 'category1'], how='left')
         .merge(targets_agg, on=['branch', 'category1'], how='left')
         .merge(pym_agg, on=['branch', 'category1'], how='left'))
df.fillna(0, inplace=True)

df['daily_tgt'] = np.where(total_working_days > 0, df['monthly_target'] / total_working_days, 0)
df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt'] > 0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'], 0)
df['mtd_tgt'] = df['daily_tgt'] * days_worked
df['mtd_var'] = np.where(df['mtd_tgt'] > 0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'], 0)
df['cm'] = df['mtd_achieved']
df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target'] > 0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'], 0)
df['projected_landing'] = np.where(days_worked > 0, (df['mtd_achieved'] / days_worked) * total_working_days, 0)
df['cm_vs_pym'] = np.where(df['pym'] > 0, (df['cm'] - df['pym']) / df['pym'], 0)

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

# === TOTALS ROW with paints monthly target sum only ===
paints_mask = df['category1'].str.lower() == 'paints'
paints_monthly_tgt_sum = df.loc[paints_mask, 'Monthly TGT'].sum()

totals = df.select_dtypes(include=np.number).sum()
totals['Monthly TGT'] = paints_monthly_tgt_sum

totals['Achieved vs Daily Tgt'] = (totals['Daily Achieved'] - totals['Daily Tgt']) / totals['Daily Tgt'] if totals['Daily Tgt'] else 0
totals['MTD Var'] = (totals['MTD Act.'] - totals['MTD TGT']) / totals['MTD TGT'] if totals['MTD TGT'] else 0
totals['Achieved VS Monthly tgt'] = (totals['MTD Act.'] - totals['Monthly TGT']) / totals['Monthly TGT'] if totals['Monthly TGT'] else 0
totals['CM VS PYM'] = (totals['CM'] - totals['PYM']) / totals['PYM'] if totals['PYM'] else 0

totals_row = {'branch': 'Total', 'category1': ''}
for col in df.columns:
    if col in totals.index:
        totals_row[col] = totals[col]
    else:
        totals_row[col] = ''

df_display = pd.concat([df, pd.DataFrame([totals_row])], ignore_index=True)

# === FORMATTING FOR AGGRID ===
percent_cols = [
    'Achieved vs Daily Tgt',
    'MTD Var',
    'Achieved VS Monthly tgt',
    'CM VS PYM'
]

format_number_js = JsCode("""
function(params) {
    if (params.value === null || params.value === undefined) return '';
    if (typeof params.value === 'number') {
        return params.value.toLocaleString();
    }
    return params.value;
}
""")

format_percent_js = JsCode("""
function(params) {
    if (params.value === null || params.value === undefined) return '';
    if (typeof params.value === 'number') {
        return (params.value * 100).toFixed(1) + '%';
    }
    return params.value;
}
""")

gb = GridOptionsBuilder.from_dataframe(df_display)

for col in df_display.columns:
    if col in percent_cols:
        gb.configure_column(col, cellRenderer=format_percent_js, type=['numericColumn','rightAligned'], sortable=True)
    elif np.issubdtype(df_display[col].dtype, np.number):
        gb.configure_column(col, cellRenderer=format_number_js, type=['numericColumn','rightAligned'], sortable=True)
    else:
        gb.configure_column(col, type=['textColumn','leftAligned'], sortable=True)

gb.configure_grid_options(domLayout='normal')
grid_options = gb.build()

# === DISPLAY TABLE ===
st.markdown("### Sales Data")
AgGrid(
    df_display,
    gridOptions=grid_options,
    theme='material',
    enable_enterprise_modules=True,
    height=400,
    fit_columns_on_grid_load=True
)

# === EXCEL DOWNLOAD ===
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sales Data')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

excel_data = to_excel(df_display)

st.download_button(
    label="📥 Download data as Excel",
    data=excel_data,
    file_name='sales_dashboard_data.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
