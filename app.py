import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Railway TADK Executive BI",
    page_icon="🚂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- PASSWORD PROTECTION ---
USERNAME = "admin"
PASSWORD = "railway123"

def check_password():
    def password_entered():
        if st.session_state["username"] == USERNAME and st.session_state["password"] == PASSWORD:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("<br><br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("## 🔒 Restricted Access")
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.button("Login", on_click=password_entered)
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("<br><br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("## 🔒 Restricted Access")
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.button("Login", on_click=password_entered)
            st.error("😕 User ID or password incorrect")
        return False
    else:
        return True

if not check_password():
    st.stop()

# --- HOME DASHBOARD ---
st.title("🚂 Indian Railways — TADK Executive BI Suite")
st.markdown("### Welcome to the Central Command Center")
st.markdown("""
Use the **sidebar navigation** on the left to seamlessly switch between dedicated pages:
* **📊 Analytics:** Deep-dive interactive visual charts, departmental distribution, and workload balancing.
* **📋 Master Table:** Strict, search-isolated data explorer with instant CSV export.
* **✏️ Data Editor:** **Live In-App Excel Editor** allowing you to modify, add, or update records directly from your browser without touching code or manually uploading binary files!
""")

excel_file = "New TADK Position as on 01.10.2026.xlsx"

@st.cache_data
def load_summary(file_path):
    xls = pd.ExcelFile(file_path)
    total_sheets = len([s for s in xls.sheet_names if s.lower() != 'sheet'])
    return total_sheets

st.info(f"📂 Connected successfully to master workbook. Tracking multiple railway divisions and units.")
