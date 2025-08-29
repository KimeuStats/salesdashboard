import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests

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

# === FORMATTING LOGIC ===

# Format percentage columns
percent_cols = ['Achieved vs Daily Tgt', 'MTD Var', 'Achieved VS Monthly tgt', 'CM VS PYM']
for col in percent_cols:
    df[col] = df[col].astype(str).str.replace('%', '')  # clean any pre-formatting
    df[col] = df[col].astype(float)

# Save a version for interactive table (as you asked)
df_display = df.copy()
for col in percent_cols:
    df_display[col] = df_display[col].round(1).astype(str) + '%'

# === FORMATTING HELPERS ===
def highlight_comparisons(val):
    if isinstance(val, float):
        if val < 0:
            return 'background-color: #ffc0cb; color: black; font-weight: bold;'
        elif val > 0:
            return 'background-color: #d0f0c0; color: black;'
    return ''

def highlight_totals(row):
    if row['branch'] == 'Totals':
        return ['background-color: #b2dfdb; font-weight: bold; font-size:16px; border: 2px solid #00796b'] * len(row)
    return [''] * len(row)

# === APPLY STYLES (non-sortable, pretty preview) ===
styled_df = df.style\
    .applymap(highlight_comparisons, subset=percent_cols)\
    .apply(highlight_totals, axis=1)\
    .set_table_styles([
        {'selector': 'thead th', 'props': [('background-color', '#b2dfdb'), ('color', 'black'),
                                           ('font-weight', 'bold'), ('text-align', 'center'),
                                           ('font-size', '13px'), ('border', '1px solid #999'),
                                           ('white-space', 'nowrap'), ('padding', '5px')]},
        {'selector': 'td', 'props': [('text-align', 'center'), ('font-size', '13px'),
                                     ('white-space', 'nowrap'), ('padding', '5px')]}
    ])\
    .format({
        'Monthly TGT': "{:,.1f}",
        'Daily Tgt': "{:,.1f}",
        'Daily Achieved': "{:,.1f}",
        'MTD TGT': "{:,.1f}",
        'MTD Act.': "{:,.1f}",
        'CM': "{:,.1f}",
        'Projected landing': "{:,.1f}",
        'PYM': "{:,.1f}",
        'Achieved vs Daily Tgt': "{:.1f}%",
        'MTD Var': "{:.1f}%",
        'Achieved VS Monthly tgt': "{:.1f}%",
        'CM VS PYM': "{:.1f}%"
    })

st.markdown("#### 📋 Styled Summary Table")
st.markdown("<div class='scrollable-table-container'>", unsafe_allow_html=True)
st.markdown(styled_df.to_html(), unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# === INTERACTIVE VERSION ===
st.markdown("#### 🧮 Interactive Table (Sortable & Filterable)")
st.dataframe(df_display, use_container_width=True)

# === DOWNLOAD ===
csv_data = df.to_csv(index=False).encode('utf-8')
st.download_button("Download Table as CSV", data=csv_data, file_name='sales_dashboard.csv', mime='text/csv')
