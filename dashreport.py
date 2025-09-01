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

# === VIEW TOGGLE ===
view_option = st.radio(
    "Select View:",
    ["Branch-Level View", "General View (Cluster + Category Only)"],
    horizontal=True
)

# === FILTER SETUP ===
clusters = sales["Cluster"].dropna().unique()
branches = sales["branch"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min, date_max = sales["date"].min(), sales["date"].max()

# === GENERAL VIEW ===
if view_option == "General View (Cluster + Category Only)":
    st.subheader("🧾 General View (Cluster + Category Only)")

    # Filters
    col1, col2 = st.columns(2)
    with col1:
        selected_cluster_g = st.selectbox("Cluster", options=["All"] + list(clusters), key="cluster_general")
    with col2:
        selected_category_g = st.selectbox("Category", options=["All"] + list(categories), key="category_general")

    col_from, col_to = st.columns(2)
    with col_from:
        st.markdown("<div style='background-color:#7b38d8; color:white; padding:8px; font-weight:bold;'>From</div>", unsafe_allow_html=True)
        start_date_g = st.date_input("", value=date_min, min_value=date_min, max_value=date_max, key="from_date_g")
    with col_to:
        st.markdown("<div style='background-color:#7b38d8; color:white; padding:8px; font-weight:bold;'>To</div>", unsafe_allow_html=True)
        end_date_g = st.date_input("", value=date_max, min_value=date_min, max_value=date_max, key="to_date_g")

    # Filter data
    general_df = sales.copy()
    if selected_cluster_g != "All":
        general_df = general_df[general_df["Cluster"] == selected_cluster_g]
    if selected_category_g != "All":
        general_df = general_df[general_df["category1"] == selected_category_g]

    general_df = general_df[(general_df["date"] >= pd.to_datetime(start_date_g)) & (general_df["date"] <= pd.to_datetime(end_date_g))]

    if general_df.empty:
        st.warning("⚠️ No data for selected filters.")
        st.stop()

    # Aggregations
    mtd_agg = general_df.groupby(['Cluster', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'MTD Achieved'})
    daily_agg = general_df[general_df['date'] == pd.to_datetime(end_date_g)].groupby(['Cluster', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'Daily Achieved'})

    prev_year_filtered = prev_year_sales[
        (prev_year_sales['date'] >= pd.Timestamp(end_date_g.year - 1, end_date_g.month, 1)) &
        (prev_year_sales['date'] <= pd.Timestamp(end_date_g.year - 1, end_date_g.month, end_date_g.days_in_month))
    ]
    pym_agg = prev_year_filtered.groupby(['Cluster', 'category1'], as_index=False)['amount'].sum().rename(columns={'amount': 'PYM'})

    df_g = (mtd_agg
        .merge(daily_agg, on=['Cluster', 'category1'], how='left')
        .merge(pym_agg, on=['Cluster', 'category1'], how='left'))

    df_g.fillna(0, inplace=True)

    # Working days
    days_worked_g = working_days_excl_sundays(pd.Timestamp(end_date_g.year, end_date_g.month, 1), end_date_g)
    total_days_g = working_days_excl_sundays(pd.Timestamp(end_date_g.year, end_date_g.month, 1), pd.Timestamp(end_date_g.year, end_date_g.month, end_date_g.days_in_month))

    # Projections
    df_g['Projected Landing'] = np.where(days_worked_g > 0, (df_g['MTD Achieved'] / days_worked_g) * total_days_g, 0)
    df_g['CM VS PYM'] = np.where(df_g['PYM'] > 0, (df_g['MTD Achieved'] - df_g['PYM']) / df_g['PYM'], 0)

    # === KPIs ===
    kpi_mtd = df_g['MTD Achieved'].sum()
    kpi_daily = df_g['Daily Achieved'].sum()
    kpi_proj = df_g['Projected Landing'].sum()

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-box"><h4>🏅 MTD Achieved</h4><p>{kpi_mtd:,.0f}</p></div>
        <div class="kpi-box"><h4>📅 Daily Achieved</h4><p>{kpi_daily:,.0f}</p></div>
        <div class="kpi-box"><h4>📈 Projected Landing</h4><p>{kpi_proj:,.0f}</p></div>
        <div class="kpi-box"><h4>💼 Days Worked</h4><p>{days_worked_g} / {total_days_g}</p></div>
    </div>
    """, unsafe_allow_html=True)

    # === BAR CHART ===
    st.markdown("### 📊 MTD Achieved per Cluster & Category")
    x_labels_g = df_g.apply(lambda row: f"{row['Cluster']} - {row['category1']}", axis=1)

    fig_g = go.Figure([
        go.Bar(x=x_labels_g, y=df_g['MTD Achieved'], name='MTD Achieved', marker_color='orange'),
        go.Bar(x=x_labels_g, y=df_g['PYM'], name='PYM', marker_color='steelblue')
    ])
    fig_g.update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig_g, use_container_width=True)

    # === AGGRID TABLE ===
    percent_cols = ['CM VS PYM']
    df_g_display = df_g.copy()
    df_g_display['CM VS PYM'] = (df_g_display['CM VS PYM'] * 100).round(1)

    gb_g = GridOptionsBuilder.from_dataframe(df_g_display)
    for col in percent_cols:
        gb_g.configure_column(
            col,
            type=["numericColumn"],
            cellStyle=JsCode("""
                function(params) {
                    if (params.value < 0) {
                        return {color: 'black', backgroundColor: '#ffc0cb', fontWeight: 'bold', textAlign: 'center'};
                    } else if (params.value > 0) {
                        return {color: 'black', backgroundColor: '#d0f0c0', textAlign: 'center'};
                    }
                    return {textAlign: 'center'};
                }
            """),
            valueFormatter="x.toFixed(1) + '%'"
        )
    st.markdown("### 🧾 Performance Table")
    AgGrid(df_g_display, gridOptions=gb_g.build(), enable_enterprise_modules=False,
           allow_unsafe_jscode=True, theme="material", height=500, fit_columns_on_grid_load=True)

    # === EXCEL DOWNLOAD ===
    df_excel_g = df_g_display.copy()
    df_excel_g['CM VS PYM'] = df_excel_g['CM VS PYM'] / 100  # revert for Excel %

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_excel_g.to_excel(writer, index=False, sheet_name='General_View')
        ws = writer.sheets['General_View']
        fill_neg = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')
        fill_pos = PatternFill(start_color='D0F0C0', end_color='D0F0C0', fill_type='solid')
        for row in range(2, len(df_excel_g) + 2):
            cell = ws.cell(row=row, column=df_excel_g.columns.get_loc('CM VS PYM') + 1)
            cell.number_format = '0.0%'
        ws.conditional_formatting.add(
            f"{openpyxl.utils.get_column_letter(df_excel_g.columns.get_loc('CM VS PYM') + 1)}2:" +
            f"{openpyxl.utils.get_column_letter(df_excel_g.columns.get_loc('CM VS PYM') + 1)}{len(df_excel_g)+1}",
            CellIsRule(operator='lessThan', formula=['0'], fill=fill_neg))
        ws.conditional_formatting.add(
            f"{openpyxl.utils.get_column_letter(df_excel_g.columns.get_loc('CM VS PYM') + 1)}2:" +
            f"{openpyxl.utils.get_column_letter(df_excel_g.columns.get_loc('CM VS PYM') + 1)}{len(df_excel_g)+1}",
            CellIsRule(operator='greaterThan', formula=['0'], fill=fill_pos))

    excel_buffer.seek(0)
    st.download_button(
        label="📥 Download General View as Excel",
        data=excel_buffer,
        file_name="general_view_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
