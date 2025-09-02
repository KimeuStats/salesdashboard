# general_view.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import requests
import base64

# === PAGE CONFIG ===
st.set_page_config(page_title="General View", layout="wide")

# === LOGO + HEADER ===
def load_base64_image_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        return base64.b64encode(response.content).decode()
    return None

logo_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/nhmllogo.png"
logo_base64 = load_base64_image_from_url(logo_url)

if logo_base64:
    st.markdown(f"""
        <div style='display: flex; align-items: center; background-color: #3FA0A3; padding: 10px 30px;'>
            <img src="data:image/png;base64,{logo_base64}" height="52" style="margin-right: 15px;" />
            <h1 style='color:white;'>General Sales View (Cluster Level)</h1>
        </div>
    """, unsafe_allow_html=True)

# === LOAD DATA ===
file_url = "https://raw.githubusercontent.com/kimeustats/salesdashboard/main/data1.xlsx"
try:
    sales = pd.read_excel(file_url, sheet_name="CY", engine="openpyxl")
except Exception as e:
    st.error(f"⚠️ Failed to load data: {e}")
    st.stop()

# === CLEAN DATA ===
sales.columns = [col if col == 'Cluster' else col.lower() for col in sales.columns]
sales['date'] = pd.to_datetime(sales['date'])
sales['amount'] = sales['amount'].astype(str).str.replace(',', '').astype(float)

# === FILTERS ===
clusters = sales["Cluster"].dropna().unique()
categories = sales["category1"].dropna().unique()
date_min, date_max = sales["date"].min(), sales["date"].max()

col1, col2 = st.columns(2)
with col1:
    selected_cluster = st.selectbox("Select Cluster", ["All"] + list(clusters))
with col2:
    selected_category = st.selectbox("Select Category", ["All"] + list(categories))

col3, col4 = st.columns(2)
with col3:
    start_date = st.date_input("From Date", date_min, min_value=date_min, max_value=date_max)
with col4:
    end_date = st.date_input("To Date", date_max, min_value=date_min, max_value=date_max)

# === APPLY FILTERS ===
filtered = sales.copy()
if selected_cluster != "All":
    filtered = filtered[filtered["Cluster"] == selected_cluster]
if selected_category != "All":
    filtered = filtered[filtered["category1"] == selected_category]
filtered = filtered[(filtered["date"] >= pd.to_datetime(start_date)) & (filtered["date"] <= pd.to_datetime(end_date))]

if filtered.empty:
    st.warning("⚠️ No data available for selected filters.")
    st.stop()

# === AGGREGATE ===
df_grouped = filtered.groupby(['Cluster', 'category1'], as_index=False)['amount'].sum()
df_grouped.rename(columns={'amount': 'Total Sales'}, inplace=True)

# === KPI ===
st.metric("💰 Total Sales", f"{df_grouped['Total Sales'].sum():,.0f}")

# === CHART ===
fig = go.Figure()
fig.add_trace(go.Bar(
    x=df_grouped.apply(lambda row: f"{row['Cluster']} - {row['category1']}", axis=1),
    y=df_grouped["Total Sales"],
    marker_color='teal'
))
fig.update_layout(title="Total Sales by Cluster & Category", xaxis_tickangle=-45)
st.plotly_chart(fig, use_container_width=True)

# === TABLE ===
st.markdown("### 🔍 Aggregated Sales Table")
st.dataframe(df_grouped.style.format({"Total Sales": "{:,.0f}"}), use_container_width=True)
