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
        .view-button {
            background-color: #7b38d8;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-right: 10px;
            font-weight: bold;
        }
        .view-button.active {
            background-color: #5a2aa3;
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

# === VIEW SELECTION ===
view_option = st.radio(
    "Select View:",
    ["General View (All Clusters)", "Detailed View (Filter by Cluster/Branch)"],
    horizontal=True,
    index=1
)

# === FILTERS ===
clusters = sales["Cluster"].dropna().unique()
branches = sales["branch"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min, date_max = sales["date"].min(), sales["date"].max()

if view_option == "General View (All Clusters)":
    # For general view, don't show cluster and branch filters
    selected_cluster = "All"
    selected_branch = "All"
    
    col1, col2 = st.columns(2)
    with col1:
        selected_category = st.selectbox("Category", options=["All"] + list(categories))
    with col2:
        pass  # Empty column for layout
else:
    # For detailed view, show all filters
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
# For general view, group by category only
if view_option == "General View (All Clusters)":
    mtd_agg = filtered.groupby(['category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'mtd_achieved'})
    daily_achieved = filtered[filtered['date'] == end_dt].groupby(['category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'daily_achieved'})

    prev_year_filtered = prev_year_sales[
        (prev_year_sales['date'] >= pd.Timestamp(end_dt.year - 1, end_dt.month, 1)) &
        (prev_year_sales['date'] <= pd.Timestamp(end_dt.year - 1, end_dt.month, end_dt.days_in_month))
    ]
    pym_agg = prev_year_filtered.groupby(['category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'pym'})
    
    # For general view, sum targets across all branches
    targets_agg_general = targets.groupby(['category1'], as_index=False)['monthly_target'].sum()
    
    df = (mtd_agg.merge(daily_achieved, on=['category1'], how='left')
             .merge(targets_agg_general, on=['category1'], how='left')
             .merge(pym_agg, on=['category1'], how='left'))
    df.fillna(0, inplace=True)
    
    # Add empty branch column for consistency
    df['branch'] = 'All Branches'
    
else:
    # For detailed view, use the original grouping
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
total_daily_achieved = df['Daily Achieved'].sum()
total_mtd_achieved = df['MTD Act.'].sum()
total_mtd_target = df['MTD TGT'].sum()
total_monthly_target = df['Monthly TGT'].sum()
total_projected_landing = df['Projected landing'].sum()
total_prev_year = df['PYM'].sum()

# === KPI DISPLAY ===
def kpi_card(title, value, delta=None, delta_color='green', width=1):
    delta_html = ""
    if delta is not None:
        sign = "▲" if delta > 0 else "▼" if delta < 0 else ""
        color = delta_color if delta > 0 else 'red' if delta < 0 else 'gray'
        delta_html = f'<div style="font-size:14px; color:{color}; font-weight:bold;">{sign} {delta:.1%}</div>'
    return f"""
        <div style="background-color:#f0f2f6; padding:15px; margin:5px; border-radius:10px; text-align:center; flex: {width};">
            <div style="font-size:18px; color:#555; font-weight:bold; margin-bottom:5px;">{title}</div>
            <div style="font-size:24px; color:#7b38d8; font-weight:bold;">{value:,.0f}</div>
            {delta_html}
        </div>
    """

col1, col2, col3, col4, col5, col6 = st.columns([1,1,1,1,1,1])
col1.markdown(kpi_card("Daily Achieved", total_daily_achieved), unsafe_allow_html=True)
col2.markdown(kpi_card("MTD Achieved", total_mtd_achieved), unsafe_allow_html=True)
col3.markdown(kpi_card("MTD Target", total_mtd_target), unsafe_allow_html=True)
col4.markdown(kpi_card("Monthly Target", total_monthly_target), unsafe_allow_html=True)
col5.markdown(kpi_card("Projected Landing", total_projected_landing), unsafe_allow_html=True)
col6.markdown(kpi_card("PYM", total_prev_year), unsafe_allow_html=True)

# === PLOTLY BAR CHART ===
if view_option == "General View (All Clusters)":
    x = df['category1']
else:
    x = df.apply(lambda row: f"{row['branch']} - {row['category1']}", axis=1)

trace_mtd = go.Bar(
    x=x,
    y=df['MTD Act.'],
    name='MTD Achieved',
    marker_color='purple'
)

trace_target = go.Bar(
    x=x,
    y=df['Monthly TGT'],
    name='Monthly Target',
    marker_color='pink'
)

layout = go.Layout(
    title="Monthly Sales Achieved vs Target",
    xaxis=dict(title='Category' if view_option=="General View (All Clusters)" else "Branch - Category", tickangle=45),
    yaxis=dict(title='Amount'),
    barmode='group',
    height=400,
    margin=dict(b=120)
)

fig = go.Figure(data=[trace_mtd, trace_target], layout=layout)
st.plotly_chart(fig, use_container_width=True)

# === AGGRID TABLE ===
aggrid_columns = [
    {"field": "branch", "headerName": "Branch", "hide": view_option == "General View (All Clusters)"},
    {"field": "category1", "headerName": "Category"},
    {"field": "Daily Tgt", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "Daily Achieved", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "Achieved vs Daily Tgt", "type": "numericColumn", "valueFormatter": {'function': "formatPercent(x)"}},
    {"field": "MTD TGT", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "MTD Act.", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "MTD Var", "type": "numericColumn", "valueFormatter": {'function': "formatPercent(x)"}},
    {"field": "Monthly TGT", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "Achieved VS Monthly tgt", "type": "numericColumn", "valueFormatter": {'function': "formatPercent(x)"}},
    {"field": "Projected landing", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "PYM", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "CM", "type": "numericColumn", "valueFormatter": {'function': "x.toLocaleString()"}},
    {"field": "CM VS PYM", "type": "numericColumn", "valueFormatter": {'function': "formatPercent(x)"}},
]

js_code = JsCode("""
    function formatPercent(params) {
        if(params == null) return "";
        return (params * 100).toFixed(1) + "%";
    }
""")

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(filterable=True, sortable=True, resizable=True)
gb.configure_columns(
    ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM'],
    type=["numericColumn"], 
    cellRenderer=JsCode("""
        function(params) {
            if(params.value < 0){
                return '<span style="color:red;font-weight:bold;">' + (params.value*100).toFixed(1) + '%</span>';
            }
            else{
                return '<span style="color:green;font-weight:bold;">' + (params.value*100).toFixed(1) + '%</span>';
            }
        }
    """)
)
gb.configure_selection(selection_mode="single")
gb.configure_pagination(paginationAutoPageSize=True)
gridOptions = gb.build()

st.subheader("Detailed Sales Table")
grid_response = AgGrid(df, gridOptions=gridOptions, enable_enterprise_modules=False, theme='material', fit_columns_on_grid_load=True)

# === EXCEL EXPORT ===
def to_excel(df_export):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df_export.to_excel(writer, index=False, sheet_name='Sales Data')
    workbook = writer.book
    worksheet = writer.sheets['Sales Data']

    # Conditional formatting
    red_fill = PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid')
    green_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')

    for col_letter in ['E', 'H', 'J', 'N']:  # Columns with percentage
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        worksheet.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df_export) + 1}',
                                             CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))
        worksheet.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df_export) + 1}',
                                             CellIsRule(operator='greaterThanOrEqual', formula=['0'], fill=green_fill))

    writer.save()
    processed_data = output.getvalue()
    return processed_data

excel_data = to_excel(df)

st.download_button(
    label="📥 Download Excel",
    data=excel_data,
    file_name='sales_data.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
