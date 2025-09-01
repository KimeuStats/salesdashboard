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

# === VIEW TOGGLE ===
view_option = st.radio("Select View:", options=["Detailed View", "General View"])

# === APPLY FILTERS ===
filtered = sales.copy()

if selected_category != "All":
    filtered = filtered[filtered["category1"] == selected_category]

if view_option == "Detailed View":
    if selected_cluster != "All":
        filtered = filtered[filtered["Cluster"] == selected_cluster]
    if selected_branch != "All":
        filtered = filtered[filtered["branch"] == selected_branch]
else:
    # General View: only cluster filter, no branch filter
    if selected_cluster != "All":
        filtered = filtered[filtered["Cluster"] == selected_cluster]

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
if view_option == "Detailed View":
    group_cols = ['branch', 'category1']
else:
    group_cols = ['Cluster', 'category1']

mtd_agg = filtered.groupby(group_cols, as_index=False)['amount'].sum().rename(columns={'amount': 'mtd_achieved'})
daily_achieved = filtered[filtered['date'] == end_dt].groupby(group_cols, as_index=False)['amount'].sum().rename(columns={'amount': 'daily_achieved'})

prev_year_filtered = prev_year_sales[
    (prev_year_sales['date'] >= pd.Timestamp(end_dt.year - 1, end_dt.month, 1)) &
    (prev_year_sales['date'] <= pd.Timestamp(end_dt.year - 1, end_dt.month, end_dt.days_in_month))
]

if view_option == "Detailed View":
    targets_agg_view = targets_agg
    pym_agg = prev_year_filtered.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})
else:
    targets_agg_view = targets.groupby(['cluster', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'monthly_target'})
    pym_agg = prev_year_filtered.groupby(['Cluster', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})

df = (mtd_agg
      .merge(daily_achieved, on=group_cols, how='left')
      .merge(targets_agg_view, on=[group_cols[0], 'category1'], how='left')
      .merge(pym_agg, on=group_cols, how='left'))

df.fillna(0, inplace=True)

df['daily_tgt'] = np.where(total_working_days > 0, df['monthly_target'] / total_working_days, 0)
df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt'] > 0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'], 0)
df['mtd_tgt'] = df['daily_tgt'] * days_worked
df['mtd_var'] = np.where(df['mtd_tgt'] > 0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'], 0)
df['cm'] = df['mtd_achieved']
df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target'] > 0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'], 0)
df['projected_landing'] = np.where(days_worked > 0, (df['mtd_achieved'] / days_worked) * total_working_days, 0)
df['proj_variance'] = np.where(df['monthly_target'] > 0, (df['projected_landing'] - df['monthly_target']) / df['monthly_target'], 0)
df['growth'] = np.where(df['pym'] > 0, (df['mtd_achieved'] - df['pym']) / df['pym'], 0)

# === RENAME COLUMNS FOR GENERAL VIEW FOR DISPLAY CONSISTENCY ===
if view_option == "General View":
    df.rename(columns={"Cluster": "cluster"}, inplace=True)
    display_group_col = "cluster"
else:
    display_group_col = "branch"

# === KPI METRICS ===
kpi1 = df['mtd_achieved'].sum()
kpi2 = df['monthly_target'].sum()
kpi3 = df['cm'].sum()
kpi4 = df['growth'].mean()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Sales (MTD)", f"{kpi1:,.0f}")
col2.metric("Target", f"{kpi2:,.0f}")
col3.metric("Contribution Margin (CM)", f"{kpi3:,.0f}")
col4.metric("Growth", f"{kpi4:.2%}")

# === PLOTLY BAR CHART ===
df['label'] = df[display_group_col].astype(str) + " - " + df['category1'].astype(str)
bar1 = go.Bar(x=df['label'], y=df['mtd_achieved'], name='Sales MTD')
bar2 = go.Bar(x=df['label'], y=df['monthly_target'], name='Target')
bar3 = go.Bar(x=df['label'], y=df['pym'], name='Prev Year Sales')
bar4 = go.Bar(x=df['label'], y=df['projected_landing'], name='Projected Landing')

layout = go.Layout(barmode='group',
                   title="Sales Performance",
                   xaxis_title="Branch - Category" if view_option == "Detailed View" else "Cluster - Category",
                   yaxis_title="Amount")

fig = go.Figure(data=[bar1, bar2, bar3, bar4], layout=layout)
st.plotly_chart(fig, use_container_width=True)

# === AGGRID TABLE ===
df_display = df[[display_group_col, 'category1', 'mtd_achieved', 'monthly_target', 'daily_achieved', 'daily_tgt',
                 'achieved_vs_daily_tgt', 'mtd_tgt', 'mtd_var', 'cm', 'achieved_vs_monthly_tgt',
                 'projected_landing', 'proj_variance', 'growth']]

df_display.columns = [display_group_col.capitalize(), 'Category', 'Sales MTD', 'Target', 'Daily Sales',
                      'Daily Target', 'Achieved vs Daily Target', 'MTD Target', 'MTD Variance', 'Contribution Margin',
                      'Achieved vs Monthly Target', 'Projected Landing', 'Projected Variance', 'Growth']

gb = GridOptionsBuilder.from_dataframe(df_display)
gb.configure_pagination(paginationAutoPageSize=True)
gb.configure_default_column(editable=False, groupable=False, filter=True, resizable=True)
gb.configure_selection(selection_mode="single", use_checkbox=True)

# Conditional styling for positive/negative growth:
cellsytle_jscode = JsCode("""
function(params) {
    if (params.value > 0) {
        return {'color': 'green', 'fontWeight': 'bold'};
    } else if (params.value < 0) {
        return {'color': 'red', 'fontWeight': 'bold'};
    }
    return null;
}
""")
gb.configure_column("Growth", cellStyle=cellsytle_jscode)
gb.configure_column("Achieved vs Daily Target", cellStyle=cellsytle_jscode)
gb.configure_column("MTD Variance", cellStyle=cellsytle_jscode)
gb.configure_column("Achieved vs Monthly Target", cellStyle=cellsytle_jscode)
gb.configure_column("Projected Variance", cellStyle=cellsytle_jscode)

grid_options = gb.build()
AgGrid(df_display, gridOptions=grid_options, theme='material', height=400, fit_columns_on_grid_load=True)

# === EXCEL EXPORT FUNCTION ===
def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sales Report')
    workbook = writer.book
    worksheet = writer.sheets['Sales Report']

    # Apply conditional formatting
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # Growth column index (0-based)
    growth_col_idx = df.columns.get_loc('Growth') + 1  # +1 for Excel indexing

    # Apply coloring row by row
    for row in range(2, len(df) + 2):  # Excel rows start at 1 + header row
        cell = worksheet.cell(row=row, column=growth_col_idx)
        if cell.value is not None:
            if cell.value > 0:
                cell.fill = green_fill
            elif cell.value < 0:
                cell.fill = red_fill

    writer.save()
    processed_data = output.getvalue()
    return processed_data

df_download = df_display.copy()
df_download['Growth'] = df_download['Growth'].apply(lambda x: f"{x:.2%}")

excel_data = to_excel(df_download)
st.download_button(label='📥 Download Excel', data=excel_data, file_name='Sales_Report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
