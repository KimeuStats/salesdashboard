import streamlit as st
import requests

owner = "kimeustats"            # your GitHub username/org
repo = "salesdashboard"         # your repo name
path = "nhmllogo.png"           # path to file inside repo
branch = "main"

token = st.secrets["github"]["token"]

url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}?ref={branch}"
headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github.v3.raw"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    st.success("✅ Token and API call worked!")
    st.write(f"File size: {len(response.content)} bytes")
else:
    st.error(f"Failed to load file: HTTP {response.status_code}")
    st.write(response.json())  # show error details
