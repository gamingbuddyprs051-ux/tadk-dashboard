import pandas as pd
import streamlit as st
import plotly.express as px

# Page configuration
st.set_page_config(
    page_title="Indian Railways — TADK Executive BI Suite",
    page_icon="🚂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ADVANCED STYLING & UI POLISH ---
st.markdown("""
    <style>
    .main {
        background-color: #0b0f19;
    }
    .stMetric {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(56, 189, 248, 0.2);
        padding: 20px;
        border-radius: 14px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .stMetric label {
        color: #94a3b8 !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    .stMetric .stMetricValue {
        color: #f8fafc !important;
        font-weight: 800 !important;
        font-size: 1.8rem !important;
    }
    h1, h2, h3 {
        font-family: 'Inter', sans-serif;
        color: #f1f5f9;
        letter-spacing: -0.5px;
    }
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    </style>
""", unsafe_allow_html=True)

# --- PASSWORD PROTECTION SETUP ---
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
            st.markdown("Please sign in to view the Indian Railways TADK Executive Dashboard.")
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

# --- HEADER SECTION ---
st.title("🚂 Indian Railways — TADK Position Executive BI Suite")
st.markdown("### Real-time Resource Intelligence & Vacancy Monitoring")
st.markdown("---")

excel_file = "New TADK Position as on 01.10.2026.xlsx"

@st.cache_data
def load_and_clean_data(file_path):
    xls = pd.ExcelFile(file_path)
    all_data = []
    for sheet in xls.sheet_names:
        if sheet.lower() == 'sheet':
            continue
        df = pd.read_excel(file_path, sheet_name=sheet)
        if df.shape[0] > 1:
            raw_headers = df.iloc[0].values
            headers = []
            for i, h in enumerate(raw_headers):
                if pd.isna(h):
                    headers.append(f"Extra_{i}")
                else:
                    headers.append(str(h).strip().replace('\n', ' '))
            
            sub_df = df.iloc[1:].copy()
            sub_df = sub_df.iloc[:, :len(headers)]
            sub_df.columns = headers[:sub_df.shape[1]]
            sub_df['Division/Unit'] = sheet
            
            if len(sub_df.columns) > 1:
                dept_col = sub_df.columns[1]
                sub_df[dept_col] = sub_df[dept_col].ffill()
                
            all_data.append(sub_df)
    
    if all_data:
        standardized_dfs = []
        for d in all_data:
            cols = list(d.columns)
            temp_df = pd.DataFrame()
            temp_df['Sl. No'] = d.iloc[:, 0] if len(cols) > 0 else None
            temp_df['Department'] = d.iloc[:, 1] if len(cols) > 1 else None
            temp_df['Total Posts'] = d.iloc[:, 2] if len(cols) > 2 else None
            temp_df['Attached Post'] = d.iloc[:, 3] if len(cols) > 3 else None
            temp_df['Post Type'] = d.iloc[:, 4] if len(cols) > 4 else None
            temp_df['Currency Upto'] = d.iloc[:, 5] if len(cols) > 5 else None
            temp_df['Incumbent Officer'] = d.iloc[:, 6] if len(cols) > 6 else None
            temp_df['Engagement Type'] = d.iloc[:, 7] if len(cols) > 7 else None
            temp_df['Division'] = d['Division/Unit']
            standardized_dfs.append(temp_df)
            
        combined = pd.concat(standardized_dfs, ignore_index=True)
        combined = combined[~combined['Department'].astype(str).str.lower().str.contains('total', na=False)]
        combined = combined[~combined['Sl. No'].astype(str).str.lower().str.contains('total', na=False)]
        
        dept_mapping = {
            'Engg': 'Engineering', 'Engg.': 'Engineering', ' Engineering': 'Engineering',
            'S&T': 'Signal & Telecom', 'Signal': 'Signal & Telecom',
            'Accounts': 'Account', 'Accounts ': 'Account',
            'Elect.': 'Electrical', 'Elect': 'Electrical',
            'Electrical(G)': 'Electrical', 'Electrical(OP)': 'Electrical',
            'Electrical(TRD)': 'Electrical', 'Electrical(TRS)': 'Electrical',
            'Elect./TRD': 'Electrical', 'Elect./G': 'Electrical', 'Elect./OP': 'Electrical',
            'Elect./TRS': 'Electrical', 'Elect (Gen.)': 'Electrical', 'Elec/TRD': 'Electrical', 'Elect/TRO': 'Electrical',
            'Mech.': 'Mechanical', 'Mech': 'Mechanical', 'Mech(C&W)': 'Mechanical',
            'Mech./C&W': 'Mechanical', 'Mech./ Dsl.': 'Mechanical',
            'Stores': 'Store', 'Store Depot/SPJ': 'Store',
            'Admin': 'General Admin', 'Gen (Admin)': 'General Admin',
            'General Admin.': 'General Admin', 'Admn General ': 'General Admin', 'Admn General': 'General Admin',
            'Pers.': 'Personnel', 'Personnel ': 'Personnel',
            'Optg.': 'Operating', 'Operating': 'Operating', 'OPTG': 'Operating',
            'Med.': 'Medical', 'Medilcal': 'Medical', 'Medical ': 'Medical',
            'Comml.': 'Commercial', 'Commercilal': 'Commercial', 'Commercial ': 'Commercial', 'COMMERCIAL': 'Commercial',
        }
        
        def standardize_dept(val):
            s = str(val).strip()
            return dept_mapping.get(s, s.title())
            
        combined['Department'] = combined['Department'].fillna('Unspecified').apply(standardize_dept)
        combined['Post Type'] = combined['Post Type'].fillna('Unspecified').astype(str).str.strip()
        combined['Engagement Type'] = combined['Engagement Type'].fillna('Unspecified').astype(str).str.strip()
        
        def clean_engagement(val):
            s = str(val).strip().lower()
            if any(k in s for k in ['vacant', 'vaccant', 'nil', '...', '---', '----', ' --', 'unspecified', 'nan']):
                return 'Vacant / Unspecified'
            elif 'contract' in s or 'attendant' in s:
                return 'Contractual'
            elif 'railway' in s or 'tadk' in s:
                return 'Railway Regular'
            else:
                return 'Other / Processing'
                
        combined['Status'] = combined['Engagement Type'].apply(clean_engagement)
        
        def clean_post_type(val):
            s = str(val).strip().lower()
            if 'work' in s:
                return 'Work Charge'
            elif 'revenue' in s:
                return 'Revenue (Permanent/Temp)'
            else:
                return 'Other'
                
        combined['Structure'] = combined['Post Type'].apply(clean_post_type)
        return combined
    return pd.DataFrame()

df = load_and_clean_data(excel_file)

if df.empty:
    st.error("No valid data found in the Excel workbook.")
else:
    # --- SIDEBAR FILTERS ---
    st.sidebar.header("🎛️ Executive Filters & Search")
    st.sidebar.markdown("Filter all metrics, charts, and tables dynamically:")
    
    # Global Search Box
    search_query = st.sidebar.text_input("🔍 Search Officer, Post, or Dept:", "").strip().lower()
    
    all_divs = sorted(df['Division'].unique().tolist())
    selected_divs = st.sidebar.multiselect("Select Division(s):", options=all_divs, default=all_divs)
    
    filtered_by_div = df[df['Division'].isin(selected_divs)]
    all_depts = sorted(filtered_by_div['Department'].unique().tolist())
    selected_depts = st.sidebar.multiselect("Select Department(s):", options=all_depts, default=all_depts)
    
    all_status = sorted(filtered_by_div['Status'].unique().tolist())
    selected_status = st.sidebar.multiselect("Select Engagement Status:", options=all_status, default=all_status)
    
    # Apply All Filters
    filtered_df = filtered_by_div[
        filtered_by_div['Department'].isin(selected_depts) &
        filtered_by_div['Status'].isin(selected_status)
    ]
    
    # Apply Search Filter (strict search isolation)
    if search_query:
        mask = filtered_df.astype(str).apply(lambda x: x.str.lower().str.contains(search_query)).any(axis=1)
        filtered_df = filtered_df[mask]
    
    # --- TOP EXECUTIVE METRICS ---
    m1, m2, m3, m4 = st.columns(4)
    total_pos = len(filtered_df)
    contractual_pos = len(filtered_df[filtered_df['Status'] == 'Contractual'])
    vacant_pos = len(filtered_df[filtered_df['Status'] == 'Vacant / Unspecified'])
    railway_pos = len(filtered_df[filtered_df['Status'] == 'Railway Regular'])
    
    m1.metric("Filtered Posts", total_pos, delta="Active View")
    m2.metric("Contractual", contractual_pos, delta=f"{round(contractual_pos/total_pos*100, 1) if total_pos else 0}% share")
    m3.metric("Vacant / Unspecified", vacant_pos, delta="Requires Action", delta_color="inverse")
    m4.metric("Railway Regular", railway_pos, delta="Permanent Staff")
    
    st.markdown("---")
    
    # --- INTERACTIVE CHARTS ---
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("📊 Division Breakdown (Filtered)")
        div_counts = filtered_df['Division'].value_counts().reset_index()
        div_counts.columns = ['Division', 'Count']
        fig_div = px.bar(
            div_counts, x='Division', y='Count', color='Division', text='Count',
            color_discrete_sequence=px.colors.sequential.Tealgrn,
            template='plotly_dark'
        )
        fig_div.update_traces(textposition="outside")
        fig_div.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=30, b=20, l=20, r=20))
        st.plotly_chart(fig_div, use_container_width=True)
        
    with c2:
        st.subheader("🍩 Engagement Status Distribution")
        status_counts = filtered_df['Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        fig_status = px.pie(
            status_counts, names='Status', values='Count', hole=0.6,
            color_discrete_sequence=px.colors.qualitative.Pastel,
            template='plotly_dark'
        )
        fig_status.update_traces(textposition='inside', textinfo='percent+label')
        fig_status.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=30, b=20, l=20, r=20))
        st.plotly_chart(fig_status, use_container_width=True)
        
    c3, c4 = st.columns(2)
    
    with c3:
        st.subheader("🏢 Departmental Strength")
        dept_counts = filtered_df['Department'].value_counts().head(10).reset_index()
        dept_counts.columns = ['Department', 'Count']
        fig_dept = px.bar(
            dept_counts, x='Count', y='Department', orientation='h',
            color='Count', color_continuous_scale='Sunsetdark',
            template='plotly_dark', text='Count'
        )
        fig_dept.update_layout(yaxis={'categoryorder':'total ascending'}, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=30, b=20, l=20, r=20))
        st.plotly_chart(fig_dept, use_container_width=True)
        
    with c4:
        st.subheader("⚙️ Post Structure (Revenue vs Work Charge)")
        struct_counts = filtered_df['Structure'].value_counts().reset_index()
        struct_counts.columns = ['Structure', 'Count']
        fig_struct = px.bar(
            struct_counts, x='Structure', y='Count', color='Structure', text='Count',
            color_discrete_sequence=px.colors.qualitative.Bold,
            template='plotly_dark'
        )
        fig_struct.update_traces(textposition="outside")
        fig_struct.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=30, b=20, l=20, r=20))
        st.plotly_chart(fig_struct, use_container_width=True)
        
    st.markdown("---")
    
    # --- STRICTLY FILTERED DATA TABLE & CSV EXPORT ---
    st.subheader("📋 Filtered Active Records Explorer")
    
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl1:
        st.markdown(f"Showing **{len(filtered_df)}** records matching your active filters and search criteria.")
    with col_dl2:
        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Export Filtered CSV",
            data=csv_data,
            file_name="filtered_tadk_executive_report.csv",
            mime="text/csv",
        )
        
    st.dataframe(filtered_df, use_container_width=True, height=450)
