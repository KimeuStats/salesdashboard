# === IMPORTS ===
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

# === PAGE CONFIG & STYLES ===
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")
st.markdown("""
<style>
    .main .block-container { max-width: 1400px; padding: 2rem; margin: auto; }
    .banner { background-color:#3FA0A3; padding:10px 30px; display:flex; align-items:center; justify-content:center; margin-bottom:20px; }
    .banner img { height:52px; margin-right:15px; border:2px solid white; box-shadow:0 0 5px rgba(255,255,255,0.7); }
    .banner h1 { color:white; font-size:26px; font-weight:bold; margin:0; }
    .ag-theme-material .ag-header { background-color:#7b38d8!important; color:white!important; font-weight:bold!important; }
    .kpi-grid { display:flex; flex-wrap:wrap; gap:16px; margin-top:10px; justify-content:space-between; }
    .kpi-box {
        flex:1 1 calc(20% - 16px);
        background-color:#f7f7fb;
        border-left:6px solid #7b38d8;
        border-radius:10px;
        padding:16px;
        min-width:150px;
        box-shadow:1px 1px 4px rgba(0,0,0,0.05);
    }
    .kpi-box h4 { margin:0; font-size:14px; color:#555; font-weight:600; }
    .kpi-box p { margin:5px 0 0; font-size:22px; font-weight:bold; color:#222; }
    @media (max-width:768px) { .kpi-box { flex:1 1 calc(48% - 16px); } }
</style>
""", unsafe_allow_html=True)

# === LOGO ===
def load_base64_image(url):
    resp = requests.get(url)
    return base64.b64encode(resp.content).decode() if resp.status_code == 200 else None

logo_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
logo_base64 = load_base64_image(logo_url)
if logo_base64:
    st.markdown(f'<div class="banner"><img src="data:image/png;base64,{logo_base64}" /><h1>Muthokinju Paints Sales Dashboard</h1></div>', unsafe_allow_html=True)
else:
    st.error("⚠️ Failed to load logo.")

# === DATA LOADING & CLEANING ===
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year_sales = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"⚠️ Data load failed: {e}")
    st.stop()

sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year_sales.columns = prev_year_sales.columns.str.lower()
sales['date'] = pd.to_datetime(sales['date'])
prev_year_sales['date'] = pd.to_datetime(prev_year_sales['date'])

for df in [sales, targets, prev_year_sales]:
    df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

targets_agg = targets.groupby(['branch', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'monthly_target'})

# === HELPER ===
def working_days_excl_sundays(start, end):
    return sum(1 for d in pd.date_range(start, end) if d.weekday() != 6)

# === VIEW TOGGLE ===
view = st.radio("Select View:", ("Branch-Level View", "General View (Cluster + Category)"), horizontal=True)

# === COMMON FILTERS & RANGE ===
clusters = sales["Cluster"].dropna().tolist()
branches = sales["branch"].dropna().tolist()
categories = sales["category1"].dropna().tolist()
min_date, max_date = sales['date'].min(), sales['date'].max()

# === DASHBOARD LOGIC ===
def run_view(branch_level=True):
    # Filters
    col1, col2, col3 = st.columns(3)
    selected_cluster = col1.selectbox("Cluster", ["All"] + clusters, key=f"cluster_{view}")
    selected_branch = None
    if branch_level:
        selected_branch = col2.selectbox("Branch", ["All"] + branches, key=f"branch_{view}")
    selected_category = col3.selectbox("Category", ["All"] + categories, key=f"category_{view}")

    from_col, to_col = st.columns(2)
    with from_col:
        st.markdown("<div style='background:#7b38d8;color:white;padding:8px;font-weight:bold;'>From</div>", unsafe_allow_html=True)
        start = st.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date, key=f"from_{view}")
    with to_col:
        st.markdown("<div style='background:#7b38d8;color:white;padding:8px;font-weight:bold;'>To</div>", unsafe_allow_html=True)
        end = st.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date, key=f"to_{view}")

    df = sales.copy()
    if selected_cluster != "All":
        df = df[df["Cluster"] == selected_cluster]
    if branch_level and selected_branch != "All":
        df = df[df["branch"] == selected_branch]
    if selected_category != "All":
        df = df[df["category1"] == selected_category]
    df = df[(df['date'] >= pd.to_datetime(start)) & (df['date'] <= pd.to_datetime(end))]

    if df.empty:
        st.warning("⚠️ No data for these filters.")
        return

    # Aggregation columns
    group_cols = ['branch', 'category1'] if branch_level else ['Cluster', 'category1']
    mtd = df.groupby(group_cols, as_index=False)['amount'].sum().rename(columns={'amount': 'MTD Act.'})
    daily = df[df['date'] == pd.to_datetime(end)].groupby(group_cols, as_index=False)['amount'].sum().rename(columns={'amount': 'Daily Achieved'})
    prev_filt = prev_year_sales[(prev_year_sales['date'] >= pd.Timestamp(end.year - 1, end.month, 1)) & (prev_year_sales['date'] <= pd.Timestamp(end.year - 1, end.month, end.days_in_month))]
    pym = prev_filt.groupby(group_cols, as_index=False)['amount'].sum().rename(columns={'amount': 'PYM'})

    df_main = mtd.merge(daily, on=group_cols, how='left').merge(targets_agg if branch_level else pym, on=group_cols, how='left').merge(pym, on=group_cols, how='left')
    df_main.fillna(0, inplace=True)

    # Dates & projections
    days_done = working_days_excl_sundays(pd.Timestamp(end.year, end.month, 1), end)
    total_days = working_days_excl_sundays(pd.Timestamp(end.year, end.month, 1), pd.Timestamp(end.year, end.month, end.days_in_month))

    if branch_level:
        df_main['Daily Tgt'] = np.where(total_days>0, df_main['monthly_target'] / total_days, 0)
        df_main['Projected landing'] = np.where(days_done>0, (df_main['MTD Act.'] / days_done) * total_days, 0)
        df_main['CM VS PYM'] = np.where(df_main['PYM'] > 0, (df_main['MTD Act.'] - df_main['PYM']) / df_main['PYM'], 0)
    else:
        df_main['Projected landing'] = np.where(days_done>0, (df_main['MTD Act.'] / days_done) * total_days, 0)
        df_main['CM VS PYM'] = np.where(df_main['PYM'] > 0, (df_main['MTD Act.'] - df_main['PYM']) / df_main['PYM'], 0)

    # KPIs
    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-box"><h4>🏅 MTD Achieved</h4><p>{df_main['MTD Act.'].sum():,.0f}</p></div>
        <div class="kpi-box"><h4>🎯 Monthly Target</h4><p>{df_main['monthly_target'].sum():,.0f}</p></div>
        <div class="kpi-box"><h4>📅 Daily Achieved</h4><p>{df_main['Daily Achieved'].sum():,.0f}</p></div>
        <div class="kpi-box"><h4>📈 Projected Landing</h4><p>{df_main['Projected landing'].sum():,.0f}</p></div>
        <div class="kpi-box"><h4>💼 Days Worked</h4><p>{days_done} / {total_days}</p></div>
    </div>
    """, unsafe_allow_html=True)

    # Bar chart
    st.markdown("###  Sales vs Targets")
    labels = df_main.apply(lambda r: f"{' - '.join(map(str, [r[c] for c in group_cols]))}", axis=1)
    bars = [go.Bar(x=labels, y=df_main['MTD Act.'], name='MTD Act.', marker_color='orange'),
            go.Bar(x=labels, y=df_main['monthly_target'] if branch_level else df_main['PYM'], name='Monthly Tgt' if branch_level else 'PYM', marker_color='steelblue')]
    fig = go.Figure(bars).update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig, use_container_width=True)

    # AgGrid table
    df_main['CM VS PYM'] = (df_main['CM VS PYM'] * 100).round(1)
    gb = GridOptionsBuilder.from_dataframe(df_main)
    gb.configure_default_column(filter=True, sortable=True, resizable=True, autoHeight=True)
    gb.configure_column("CM VS PYM", type=["numericColumn"], cellStyle=JsCode("""
        function(params){
            if(params.value<0) {return {backgroundColor:'#ffc0cb', color:'black', fontWeight:'bold', textAlign:'center'};}
            else if(params.value>0) {return {backgroundColor:'#d0f0c0', color:'black', textAlign:'center'};}
            return {textAlign:'center'};
        }
    """), valueFormatter="x.toFixed(1) + '%'")
    st.markdown("###  Performance Table")
    AgGrid(df_main, gridOptions=gb.build(), enable_enterprise_modules=False, allow_unsafe_jscode=True, theme="material", height=500, fit_columns_on_grid_load=True)

    # Excel download
    df_excel = df_main.copy()
    df_excel['CM VS PYM'] = df_excel['CM VS PYM'] / 100
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='View')
        ws = writer.sheets['View']
        neg = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')
        pos = PatternFill(start_color='D0F0C0', end_color='D0F0C0', fill_type='solid')
        col_idx = df_excel.columns.get_loc('CM VS PYM') + 1
        for row in range(2, len(df_excel) + 2):
            ws.cell(row, col_idx).number_format = '0.0%'
        ws.conditional_formatting.add(f"{openpyxl.utils.get_column_letter(col_idx)}2:{openpyxl.utils.get_column_letter(col_idx)}{len(df_excel)+1}", CellIsRule(operator='lessThan', formula=['0'], fill=neg))
        ws.conditional_formatting.add(f"{openpyxl.utils.get_column_letter(col_idx)}2:{openpyxl.utils.get_column_letter(col_idx)}{len(df_excel)+1}", CellIsRule(operator='greaterThan', formula=['0'], fill=pos))
    buffer.seek(0)

    label = "📥 Download Branch View as Excel" if branch_level else "📥 Download General View as Excel"
    fname = "branch_level_view.xlsx" if branch_level else "general_view.xlsx"
    st.download_button(label=label, data=buffer, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Run the selected view
run_view(branch_level=(view == "Branch-Level View"))
