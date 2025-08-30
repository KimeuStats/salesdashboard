import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import io
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# === PAGE CONFIG ===
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")

# === STYLES ===
st.markdown("""
    <style>
        .main .block-container {
            max-width: 1400px;
            padding: 2rem 2rem;
            margin: auto;
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
        .ag-theme-material .ag-header {
            background-color: #7b38d8 !important;
            color: white !important;
            font-weight: bold !important;
        }
    </style>
""", unsafe_allow_html=True)

# === LOGO ===
def load_base64_image_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        return base64.b64encode(response.content).decode()
    return None

logo_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
logo_base64 = load_base64_image_from_url(logo_url)

if logo_base64:
    st.markdown(f"""
        <div class="banner">
            <img src="data:image/png;base64,{logo_base64}" alt="Logo" />
            <h1>Muthokinju Paints Sales Dashboard</h1>
        </div>
    """, unsafe_allow_html=True)
else:
    st.error("⚠️ Failed to load logo image.")

# === LOAD DATA ===
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year_sales = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"⚠️ Failed to load Excel data: {e}")
    st.stop()

# === CLEAN DATA ===
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
    st.warning("⚠️ No sales data found for the selected filters or date range.")
    st.stop()

# === CORRECT WORKING DAYS LOGIC ===
end_dt = pd.to_datetime(end_date)
month_start = pd.Timestamp(end_dt.year, end_dt.month, 1)
month_end = pd.Timestamp(end_dt.year, end_dt.month, end_dt.days_in_month)

days_worked = working_days_excl_sundays(month_start, end_dt)
total_working_days = working_days_excl_sundays(month_start, month_end)

# === AGGREGATIONS ===
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

df['daily_tgt'] = np.where(total_working_days>0, df['monthly_target']/total_working_days, 0)
df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt']>0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'], 0)
df['mtd_tgt'] = df['daily_tgt'] * days_worked
df['mtd_var'] = np.where(df['mtd_tgt']>0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'], 0)
df['cm'] = df['mtd_achieved']
df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target']>0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'], 0)
df['projected_landing'] = np.where(days_worked>0, (df['mtd_achieved'] / days_worked) * total_working_days, 0)
df['cm_vs_pym'] = np.where(df['pym']>0, (df['cm'] - df['pym']) / df['pym'], 0)

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

# === KPI CALCULATIONS ===
kpi1 = df['MTD Act.'].sum()
kpi2 = df['Monthly TGT'].sum()
kpi3 = df['Daily Achieved'].sum()
kpi4 = df['Projected landing'].sum()

# === STYLES ===
st.markdown("""
<style>
.kpi-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 16px;
    margin-top: 10px;
    justify-content: space-between;
}
.kpi-box {
    flex: 1 1 calc(20% - 16px);
    background-color: #f7f7fb;
    border-left: 6px solid #7b38d8;
    border-radius: 10px;
    padding: 16px;
    min-width: 150px;
    box-shadow: 1px 1px 4px rgba(0,0,0,0.05);
}
.kpi-box h4 {
    margin: 0;
    font-size: 14px;
    color: #555;
    font-weight: 600;
}
.kpi-box p {
    margin: 5px 0 0 0;
    font-size: 22px;
    font-weight: bold;
    color: #222;
}
@media only screen and (max-width: 768px) {
    .kpi-box {
        flex: 1 1 calc(48% - 16px);
    }
}
</style>
""", unsafe_allow_html=True)

# === KPI DISPLAY ===
st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-box">
        <h4>🏅 MTD Achieved</h4>
        <p>{kpi1:,.0f}</p>
    </div>
    <div class="kpi-box">
        <h4>🎯 Monthly Target</h4>
        <p>{kpi2:,.0f}</p>
    </div>
    <div class="kpi-box">
        <h4>📅 Daily Achieved</h4>
        <p>{kpi3:,.0f}</p>
    </div>
    <div class="kpi-box">
        <h4>📈 Projected Landing</h4>
        <p>{kpi4:,.0f}</p>
    </div>
    <div class="kpi-box">
        <h4>💼 Days Worked</h4>
        <p>{days_worked} / {total_working_days}</p>
    </div>
</div>
""", unsafe_allow_html=True)

# === SALES VS TARGET CHART ===
st.markdown("### 📊 Sales vs Monthly Target (MTD)")
df_chart = df.copy()
x_labels = df_chart.apply(lambda row: f"{row['branch']} - {row['category1']}", axis=1)

fig = go.Figure([
    go.Bar(x=x_labels, y=df_chart['MTD Act.'], name='MTD Achieved', marker_color='orange'),
    go.Bar(x=x_labels, y=df_chart['Monthly TGT'], name='Monthly Target', marker_color='steelblue')
])
fig.update_layout(barmode='group', xaxis_tickangle=-45,
                  height=500, margin=dict(b=150),
                  legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
st.plotly_chart(fig, use_container_width=True)

# === AGGRID DISPLAY with Paints replaced by Totals ===
df_display = df.copy()
percent_cols = ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']

def safe_div(n, d): return (n - d) / d if d else 0

# Calculate totals sums for numeric columns
actual_sums = df_display[['Daily Achieved', 'MTD Act.', 'Projected landing', 'CM']].sum()
totals_values = {
    'Monthly TGT': df['Monthly TGT'].sum(),
    'Daily Tgt': df['Daily Tgt'].sum(),
    'MTD TGT': df['MTD TGT'].sum(),
    'PYM': df['PYM'].sum()
}

# Paints row = duplicate of totals (sums)
paints_row = {
    'branch': 'Paints',
    'category1': '',
    'Monthly TGT': totals_values['Monthly TGT'],
    'Daily Tgt': totals_values['Daily Tgt'],
    'MTD TGT': totals_values['MTD TGT'],
    'PYM': totals_values['PYM'],
    'Daily Achieved': actual_sums['Daily Achieved'],
    'MTD Act.': actual_sums['MTD Act.'],
    'Projected landing': actual_sums['Projected landing'],
    'CM': actual_sums['CM'],
    'Achieved vs Daily Tgt': safe_div(actual_sums['Daily Achieved'], totals_values['Daily Tgt']),
    'MTD Var': safe_div(actual_sums['MTD Act.'], totals_values['MTD TGT']),
    'Achieved VS Monthly tgt': safe_div(actual_sums['MTD Act.'], totals_values['Monthly TGT']),
    'CM VS PYM': safe_div(actual_sums['CM'], totals_values['PYM']),
    'is_totals': False
}
paints_row = pd.DataFrame([paints_row])

totals = paints_row.copy()
totals['branch'] = 'Totals'
totals['is_totals'] = True

# Remove existing Paints row(s) if any
df_display = df_display[df_display['category1'].str.lower() != 'paints']

# Append Paints and Totals rows
df_display = pd.concat([df_display, paints_row], ignore_index=True)
df_display = pd.concat([df_display, totals], ignore_index=True)

# AGGRID OPTIONS
gb = GridOptionsBuilder.from_dataframe(df_display)

for col in df_display.columns:
    if col in percent_cols:
        gb.configure_column(col, type=["numericColumn", "customPercentageFormat"],
                            cellStyle={'color': 'green'},
                            valueFormatter="(x*100).toFixed(2) + '%'")
    elif col in ['Monthly TGT', 'Daily Tgt', 'MTD TGT', 'PYM', 'Daily Achieved', 'MTD Act.', 'Projected landing', 'CM']:
        gb.configure_column(col, type=["numericColumn", "customNumberFormat"],
                            valueFormatter="x.toLocaleString()")

# Style totals row
gb.configure_grid_options(getRowStyle=JsCode("""
    function(params) {
        if (params.data.is_totals) {
            return {'background-color': '#7b38d8', 'color': 'white', 'font-weight': 'bold'};
        }
    }
"""))

gridOptions = gb.build()

grid_response = AgGrid(
    df_display,
    gridOptions=gridOptions,
    enable_enterprise_modules=False,
    theme='material',
    height=500,
    fit_columns_on_grid_load=True,
    allow_unsafe_jscode=True,
    reload_data=True,
)

# === DOWNLOAD BUTTON ===
def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Report')
    workbook = writer.book
    worksheet = writer.sheets['Report']

    fill = PatternFill(start_color='7b38d8', end_color='7b38d8', fill_type='solid')

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        if row[0].value == 'Totals':
            for cell in row:
                cell.fill = fill
                cell.font = openpyxl.styles.Font(color='FFFFFF', bold=True)

    writer.save()
    processed_data = output.getvalue()
    return processed_data

# Prepare download data without is_totals column
df_download = df_display.drop(columns=['is_totals'])
excel_data = to_excel(df_download)

st.download_button(
    label='⬇️ Download Report as Excel',
    data=excel_data,
    file_name='Muthokinju_Paints_Sales_Report.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
