import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import base64
import requests
import io
from datetime import datetime

# Optional: export table as image
try:
    import dataframe_image as dfi
    HAS_DFI = True
except ModuleNotFoundError:
    HAS_DFI = False

# === PAGE CONFIG & STYLES ===
st.set_page_config(layout="wide", page_title="Muthokinju Paints Sales Dashboard")
st.markdown("""
<style>
.main .block-container { max-width: 1400px; padding: 2rem; margin: auto; }
.table-scroll-area { overflow-x: auto; border: 1px solid #ccc; padding: 10px; }
</style>
""", unsafe_allow_html=True)

# === LOGO HEADER ===
def load_image_b64(url):
    resp = requests.get(url)
    return base64.b64encode(resp.content).decode() if resp.status_code == 200 else None

logo_b64 = load_image_b64("https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png")
if logo_b64:
    st.markdown(f"""
    <div style="display:flex; align-items:center; margin-bottom:20px;">
      <img src="data:image/png;base64,{logo_b64}" style="height:60px; margin-right:15px;" />
      <h1 style="margin:0;">Muthokinju Paints Sales Dashboard</h1>
    </div>""", unsafe_allow_html=True)
else:
    st.error("⚠️ Logo failed to load.")

# === DATA LOADING ===
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
    targets = pd.read_excel(file_url, sheet_name="TARGETS", engine="openpyxl")
    prev_year = pd.read_excel(file_url, sheet_name="PY", engine="openpyxl")
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.stop()

# Clean and standardize
for df in (sales, targets, prev_year):
    if 'amount' in df.columns:
        df['amount'] = df['amount'].astype(str).str.replace(',', '').astype(float)

sales.columns = [col if col=='Cluster' else col.lower() for col in sales.columns]
targets.columns = targets.columns.str.lower()
prev_year.columns = prev_year.columns.str.lower()
sales['date'] = pd.to_datetime(sales['date'])
prev_year['date'] = pd.to_datetime(prev_year['date'])

# Targets aggregation
targets_agg = targets.groupby(['branch','category1'], as_index=False)['amount'].sum().rename(columns={'amount':'Monthly Target'})

# === HELPER FUNCTION ===
def working_days_excl_sundays(start_date, end_date):
    return len([d for d in pd.date_range(start=start_date, end=end_date) if d.weekday() != 6])

# === FILTER UI ===
clusters = sales['Cluster'].dropna().unique().tolist()
branches = sales['branch'].dropna().unique().tolist()
categories = sales['category1'].dropna().unique().tolist()

c1, c2, c3, c4 = st.columns([1,1,1,2])
with c1:
    sel_cluster = st.selectbox("Cluster", ["All"] + clusters)
with c2:
    sel_branch = st.selectbox("Branch", ["All"] + branches)
with c3:
    sel_category = st.selectbox("Category", ["All"] + categories)
with c4:
    dr = st.date_input("Date Range", value=(sales['date'].min(), sales['date'].max()),
                       min_value=sales['date'].min(), max_value=sales['date'].max())

if not isinstance(dr, tuple) or len(dr)!=2:
    st.error("Select a valid date range.")
    st.stop()
start_date, end_date = dr

# Filter data
filtered = sales.copy()
if sel_cluster!="All":
    filtered = filtered[filtered['Cluster']==sel_cluster]
if sel_branch!="All":
    filtered = filtered[filtered['branch']==sel_branch]
if sel_category!="All":
    filtered = filtered[filtered['category1']==sel_category]
filtered = filtered[(filtered['date']>=start_date)&(filtered['date']<=end_date)]

# Working days
end_ts = pd.Timestamp(end_date)
dim = pd.Timestamp(end_ts).days_in_month
wd_month = working_days_excl_sundays(pd.Timestamp(end_ts.year, end_ts.month,1), pd.Timestamp(end_ts.year,end_ts.month,dim))
wd_done = working_days_excl_sundays(start_date, end_date)

# === AGGREGATION ===
if not filtered.empty:
    days_passed = wd_done
    mtd = filtered.groupby(['branch','category1'], as_index=False)['amount'].sum().rename(columns={'amount':'MTD Act.'})
    daily = filtered[filtered['date']==end_ts].groupby(['branch','category1'], as_index=False)['amount'].sum().rename(columns={'amount':'Daily Achieved'})
    prev_f = prev_year[(prev_year['date']>=pd.Timestamp(end_ts.year-1,end_ts.month,1))&(prev_year['date']<=pd.Timestamp(end_ts.year-1,end_ts.month,dim))]
    py = prev_f.groupby(['branch','category1'], as_index=False)['amount'].sum().rename(columns={'amount':'PYM'})
    df = (mtd.merge(daily,on=['branch','category1'],how='left')
               .merge(targets_agg,on=['branch','category1'],how='left')
               .merge(py,on=['branch','category1'],how='left')).fillna(0)
    df['Daily Tgt'] = df['Monthly Target']/wd_month
    df['MTD TGT'] = df['Daily Tgt']*days_passed
    df['MTD Var'] = np.where(df['MTD TGT']==0,0,(df['MTD Act.']-df['MTD TGT'])/df['MTD TGT'])
    df['Achieved VS Monthly tgt'] = np.where(df['Monthly Target']==0,0,(df['MTD Act.']-df['Monthly Target'])/df['Monthly Target'])
    df['Projected landing'] = np.where(days_passed==0,0,(df['MTD Act.']/days_passed)*wd_month)
    df['CM VS PYM'] = np.where(df['PYM']==0,0,(df['MTD Act.']-df['PYM'])/df['PYM'])
    for col in ['MTD Var','Achieved VS Monthly tgt','CM VS PYM']:
        df[col] = (df[col]*100).round(1).astype(str)+'%'
    total = df[df['branch']!='Totals']
    tot_row = {'branch':'Totals','category1':'','Monthly Target':total['Monthly Target'].sum(),
               'MTD Act.':total['MTD Act.'].sum(),'Daily Tgt':total['Daily Tgt'].sum(),
               'MTD TGT':total['MTD TGT'].sum(),'Projected landing':total['Projected landing'].sum(),
               'PYM':total['PYM'].sum()}
    df = pd.concat([df,pd.DataFrame([tot_row])],ignore_index=True)
else:
    df = pd.DataFrame()

# === KPI CARDS ===
k1,k2,k3,k4 = st.columns(4)
k1.metric("Work Days in Month", wd_month)
k2.metric("Days Worked", wd_done)
k3.metric("MTD Achieved", f"{filtered['amount'].sum():,.1f}" if not filtered.empty else "0")
k4.metric("Monthly Target", f"{df['Monthly Target'].sum():,.1f}" if not df.empty else "0")

# === CHART ===
st.markdown("### Sales vs Monthly Target (MTD)")
fig = go.Figure()
if not df.empty:
    chart_df = df[df['branch']!='Totals']
    xlabels = chart_df.apply(lambda r: f"{r['branch']} - {r['category1']}", axis=1)
    fig.add_trace(go.Bar(x=xlabels, y=chart_df['MTD Act.'], name='MTD Achieved', marker_color='orange'))
    fig.add_trace(go.Bar(x=xlabels, y=chart_df['Monthly Target'], name='Monthly Target', marker_color='steelblue'))
fig.update_layout(barmode='group', xaxis_tickangle=-45, height=500, margin=dict(b=150),
                  legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
st.plotly_chart(fig, use_container_width=True, config={
    "displaylogo": False, "displayModeBar": True,
    "modeBarButtonsToRemove": ["zoom", "pan", "select2d", "lasso2d", "autoScale2d", "resetScale2d"],
    "modeBarButtonsToAdd": ["toImage"]
})

# === TABLE & DOWNLOAD ===
st.markdown("### Data Table")
st.markdown('<div class="table-scroll-area">', unsafe_allow_html=True)
if not df.empty:
    st.write(df)
else:
    st.write("No data to display.")
st.markdown('</div>', unsafe_allow_html=True)
st.download_button("Download CSV", df.to_csv(index=False).encode('utf-8'), file_name="data.csv", mime="text/csv")
if HAS_DFI and not df.empty:
    buf = io.BytesIO()
    dfi.export(df, buf, table_conversion="matplotlib")
    buf.seek(0)
    st.download_button("Download Table as PNG", buf.read(), "table.png", "image/png")
elif not HAS_DFI:
    st.info("Install dataframe_image to enable image download.")
