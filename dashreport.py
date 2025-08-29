import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import streamlit.components.v1 as components
import io


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
        .scrollable-table-container {
            max-height: 550px;
            overflow-y: auto;
            border: 1px solid #ccc;
            padding-right: 10px;
            margin: auto;
            width: 100%;
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
        /* Color column headers background */
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

targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum()
targets_agg.rename(columns={'amount': 'monthly_target'}, inplace=True)

# === HELPER FUNCTION ===
def working_days_excl_sundays(start_date, end_date):
    return len([d for d in pd.date_range(start=start_date, end=end_date) if d.weekday() != 6])

# === FILTERS ===
clusters = sales["Cluster"].dropna().unique()
branches = sales["branch"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min = sales["date"].min()
date_max = sales["date"].max()

col1, col2, col3 = st.columns(3)
with col1:
    selected_cluster = st.selectbox("Cluster", options=["All"] + list(clusters))
with col2:
    selected_branch = st.selectbox("Branch", options=["All"] + list(branches))
with col3:
    selected_category = st.selectbox("Category", options=["All"] + list(categories))

# === CUSTOM DATE PICKERS ===
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

# === HANDLE NO DATA ===
if filtered.empty:
    st.warning("⚠️ No sales data found for the selected filters or date range.")

    fig = go.Figure()
    fig.update_layout(title="No Data Available",
                      xaxis_title="Branch - Category",
                      yaxis_title="Amount",
                      height=400,
                      modebar_remove=["zoom", "pan", "select", "zoomIn", "zoomOut", "resetScale2d",
                                      "autoScale2d", "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"])
    st.plotly_chart(fig, use_container_width=True)

    empty_df = pd.DataFrame(columns=[
        "branch", "category1", "Monthly TGT", "Daily Tgt", "Daily Achieved", "Achieved vs Daily Tgt",
        "MTD TGT", "MTD Act.", "MTD Var", "CM", "Achieved VS Monthly tgt", "Projected landing", "PYM", "CM VS PYM"
    ])
    st.dataframe(empty_df)
    st.stop()

# === AGGREGATIONS ===
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

df['daily_tgt'] = df['monthly_target'] / total_working_days
df['achieved_vs_daily_tgt'] = np.where(df['daily_tgt'] == 0, 0, (df['daily_achieved'] - df['daily_tgt']) / df['daily_tgt'])
df['mtd_tgt'] = df['daily_tgt'] * days_passed
df['mtd_var'] = np.where(df['mtd_tgt'] == 0, 0, (df['mtd_achieved'] - df['mtd_tgt']) / df['mtd_tgt'])
df['cm'] = df['mtd_achieved']
df['achieved_vs_monthly_tgt'] = np.where(df['monthly_target'] == 0, 0, (df['mtd_achieved'] - df['monthly_target']) / df['monthly_target'])
df['projected_landing'] = np.where(days_passed == 0, 0, (df['mtd_achieved'] / days_passed) * total_working_days)
df['cm_vs_pym'] = np.where(df['pym'] == 0, 0, (df['cm'] - df['pym']) / df['pym'])

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

# === KPI CARDS ===
kpi1 = df['MTD Act.'].sum()
kpi2 = df['Monthly TGT'].sum()
kpi3 = df['Daily Achieved'].sum()
kpi4 = df['Projected landing'].sum()

colA, colB, colC, colD = st.columns(4)
colA.metric("💰 MTD Achieved", f"{kpi1:,.0f}")
colB.metric("🎯 Monthly Target", f"{kpi2:,.0f}")
colC.metric("📅 Daily Achieved", f"{kpi3:,.0f}")
colD.metric("📈 Projected Landing", f"{kpi4:,.0f}")

# === CHART ===
st.markdown("### 📊 Sales vs Monthly Target (MTD)")
df_chart = df.copy()
x_labels = df_chart.apply(lambda row: f"{row['branch']} - {row['category1']}", axis=1)

fig = go.Figure()
fig.add_trace(go.Bar(x=x_labels, y=df_chart['MTD Act.'], name='MTD Achieved', marker_color='orange'))
fig.add_trace(go.Bar(x=x_labels, y=df_chart['Monthly TGT'], name='Monthly Target', marker_color='steelblue'))

fig.update_layout(
    barmode='group',
    xaxis_tickangle=-45,
    height=500,
    margin=dict(b=150),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    modebar_remove=["zoom", "pan", "select", "zoomIn", "zoomOut", "resetScale2d", "autoScale2d", 
                    "lasso2d", "hoverClosestCartesian", "hoverCompareCartesian"]
)
st.plotly_chart(fig, use_container_width=True)



# === Prepare dataframe for AgGrid ===
df_display = df.copy()

# Columns used in % calculations
percent_cols = ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']
numeric_cols_to_sum = [
    'Monthly TGT', 'Daily Tgt', 'Daily Achieved',
    'MTD TGT', 'MTD Act.', 'CM', 'Projected landing', 'PYM'
]

# --- Create totals row ---
totals_dict = {
    col: df[col].sum() if col in numeric_cols_to_sum else '' for col in df.columns
}
totals_dict['branch'] = 'Totals'
totals_dict['category1'] = ''

# Recalculate percentages in totals row using base totals
def safe_div(n, d):
    return (n - d) / d if d != 0 else 0

totals_dict['Achieved vs Daily Tgt'] = safe_div(totals_dict['Daily Achieved'], totals_dict['Daily Tgt'])
totals_dict['MTD Var'] = safe_div(totals_dict['MTD Act.'], totals_dict['MTD TGT'])
totals_dict['Achieved VS Monthly tgt'] = safe_div(totals_dict['MTD Act.'], totals_dict['Monthly TGT'])
totals_dict['CM VS PYM'] = safe_div(totals_dict['CM'], totals_dict['PYM'])

# Append totals row
df_display = pd.concat([df_display, pd.DataFrame([totals_dict])], ignore_index=True)
df_display['is_totals'] = df_display['branch'] == 'Totals'

# --- Round values for display in the app ---
for col in df_display.columns:
    if col in percent_cols:
        df_display[col] = (df_display[col].astype(float) * 100).round(1)  # percentages shown as 0-100 scale in app
    elif pd.api.types.is_numeric_dtype(df_display[col]):
        df_display[col] = df_display[col].round(1)

# === Configure AgGrid ===
gb = GridOptionsBuilder.from_dataframe(df_display)
gb.configure_default_column(filter=True, sortable=True, resizable=True, autoHeight=True)
gb.configure_column("is_totals", hide=True)

# Conditional formatting for percentage columns
cell_style_jscode = JsCode("""
function(params) {
    if (params.value == null) return {};
    if (params.value < 0) {
        return {color: 'black', backgroundColor: '#ffc0cb', fontWeight: 'bold', textAlign: 'center'};
    } else if (params.value > 0) {
        return {color: 'black', backgroundColor: '#d0f0c0', textAlign: 'center'};
    }
    return {textAlign: 'center'};
}
""")

for col in percent_cols:
    gb.configure_column(
        col,
        cellStyle=cell_style_jscode,
        type=["numericColumn", "numberColumnFilter", "customNumericFormat"],
        valueFormatter="x.toFixed(1) + '%'",
        headerClass='header-center'
    )

# Totals row styling
totals_row_style = JsCode("""
function(params) {
    if (params.data.is_totals) {
        return {
            'backgroundColor': '#b2dfdb',
            'fontWeight': 'bold',
            'fontSize': '14px',
            'textAlign': 'center'
        }
    }
    return {};
}
""")

gb.configure_grid_options(getRowStyle=totals_row_style)

# === Custom CSS for Grid ===
custom_css = """
.ag-theme-material .ag-header-cell-label {
    justify-content: center !important;
    font-weight: bold !important;
    background-color: #d3d3d3 !important;
    color: #222222 !important;
}
.ag-theme-material .ag-cell {
    border: 1px solid #ccc !important;
    text-align: center !important;
}
.ag-theme-material .ag-row {
    border-bottom: 1px solid #ccc !important;
}
"""
st.markdown(f"<style>{custom_css}</style>", unsafe_allow_html=True)

# === Render Table ===
st.markdown("### <center>📋 <span style='font-size:22px; font-weight:bold; color:#7b38d8;'>PERFORMANCE TABLE</span></center>", unsafe_allow_html=True)

AgGrid(
    df_display,
    gridOptions=gb.build(),
    enable_enterprise_modules=False,
    allow_unsafe_jscode=True,
    theme="material",
    height=500,
    fit_columns_on_grid_load=False,
    reload_data=True
)

# === Excel Download ===
import openpyxl

df_excel = df_display.copy()

# Drop 'is_totals' and '::auto_unique_id::' if present
df_excel = df_excel.drop(columns=['is_totals', '::auto_unique_id::'], errors='ignore')

# Convert percentage columns back to decimal for Excel formatting
for col in percent_cols:
    df_excel[col] = df_excel[col] / 100

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
    df_excel.to_excel(writer, index=False, sheet_name='Performance')
    workbook = writer.book
    worksheet = writer.sheets['Performance']
    
    header = list(df_excel.columns)
    for col_name in percent_cols:
        if col_name in header:
            col_idx = header.index(col_name) + 1
            for row in range(2, len(df_excel) + 2):
                worksheet.cell(row=row, column=col_idx).number_format = '0.0%'

excel_buffer.seek(0)

st.download_button(
    label="📥 Download Table as Excel",
    data=excel_buffer,
    file_name="sales_dashboard_with_totals.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
