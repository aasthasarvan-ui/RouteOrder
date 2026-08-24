import streamlit as st
import pandas as pd
import datetime
import pytz
import sqlite3

# Page Configuration & Styling
st.set_page_config(
    page_title="Enterprise Plan & Dispatch Hub", 
    page_icon="🚚", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 8 Professional Enterprise Themes
THEMES = {
    "💼 Classic Enterprise Navy": {
        "icon": "💼", "bg": "#f4f6f9", "text": "#1f2937", "card_bg": "#ffffff", "border": "#cbd5e1",
        "btn_bg": "#1e3a8a", "btn_hover": "#1d4ed8", "primary": "#2563eb", "input_bg": "#ffffff", "input_text": "#1f2937"
    },
    "🌙 Modern Dark ERP": {
        "icon": "🌙", "bg": "#0b0f19", "text": "#f3f4f6", "card_bg": "#1f2937", "border": "#374151",
        "btn_bg": "#374151", "btn_hover": "#4b5563", "primary": "#3b82f6", "input_bg": "#111827", "input_text": "#f3f4f6"
    },
    "📊 Corporate Slate": {
        "icon": "📊", "bg": "#eef2f5", "text": "#0f172a", "card_bg": "#ffffff", "border": "#94a3b8",
        "btn_bg": "#475569", "btn_hover": "#334155", "primary": "#0284c7", "input_bg": "#ffffff", "input_text": "#0f172a"
    },
    "☀️ Clean Light Minimal": {
        "icon": "☀️", "bg": "#ffffff", "text": "#111827", "card_bg": "#f9fafb", "border": "#d1d5db",
        "btn_bg": "#0f172a", "btn_hover": "#1e293b", "primary": "#10b981", "input_bg": "#ffffff", "input_text": "#111827"
    },
    "⚡ Cyber Blue": {
        "icon": "⚡", "bg": "#f0fdfa", "text": "#042f2e", "card_bg": "#ccfbf1", "border": "#5eead4",
        "btn_bg": "#0d9488", "btn_hover": "#0f766e", "primary": "#14b8a6", "input_bg": "#ffffff", "input_text": "#042f2e"
    },
    "🌲 Emerald Corporate": {
        "icon": "🌲", "bg": "#f0fdf4", "text": "#14532d", "card_bg": "#dcfce7", "border": "#86efac",
        "btn_bg": "#16a34a", "btn_hover": "#15803d", "primary": "#22c55e", "input_bg": "#ffffff", "input_text": "#14532d"
    },
    "🍇 Executive Burgundy": {
        "icon": "🍇", "bg": "#fdf2f8", "text": "#500724", "card_bg": "#fce7f3", "border": "#f472b6",
        "btn_bg": "#db2777", "btn_hover": "#be185d", "primary": "#ec4899", "input_bg": "#ffffff", "input_text": "#500724"
    },
    "🪙 Titanium Charcoal": {
        "icon": "🪙", "bg": "#18181b", "text": "#fafafa", "card_bg": "#27272a", "border": "#52525b",
        "btn_bg": "#52525b", "btn_hover": "#71717a", "primary": "#e4e4e7", "input_bg": "#09090b", "input_text": "#fafafa"
    }
}

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

# --- Ensure All Tables Exist to Prevent Errors ---
def init_dispatch_db():
    conn = sqlite3.connect("sales_history.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS history_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            files_count INTEGER,
            total_qty REAL,
            status TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS output_files_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            file_type TEXT,
            file_data BLOB,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS dispatch_manifests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_date TEXT,
            file_name TEXT,
            transporter_name TEXT,
            vehicle_no TEXT,
            driver_name TEXT,
            driver_phone TEXT,
            dispatch_status TEXT,
            remarks TEXT,
            updated_at TEXT
        )
    """)
    conn.commit()
    conn.close()

init_dispatch_db()

if "dispatch_theme" not in st.session_state:
    st.session_state.dispatch_theme = "💼 Classic Enterprise Navy"

t = THEMES[st.session_state.dispatch_theme]

st.markdown(f"""
    <style>
        .stApp {{ background-color: {t['bg']}; color: {t['text']}; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
        h1, h2, h3, h4, h5, h6, p, span, label {{ color: {t['text']} !important; }}
        input, textarea, select {{ background-color: {t['input_bg']} !important; color: {t['input_text']} !important; border: 1px solid {t['border']} !important; }}
        .stButton>button {{ background-color: {t['btn_bg']} !important; color: #ffffff !important; font-weight: 600; border-radius: 4px; border: 1px solid {t['border']}; }}
        .stButton>button:hover {{ background-color: {t['btn_hover']} !important; }}
        div[data-testid="stExpander"] {{ background-color: {t['card_bg']}; border: 1px solid {t['border']}; }}
    </style>
""", unsafe_allow_html=True)

with st.expander("⚙️ Dispatch Control Panel & Theme Selection", expanded=False):
    st.selectbox("Select Theme", list(THEMES.keys()), key="dispatch_theme", index=list(THEMES.keys()).index(st.session_state.dispatch_theme))

st.title(f"🚚 Enterprise Plan, Demand & Dispatch Hub ({st.session_state.dispatch_theme})")
st.markdown("Demand app ke generated outputs aur database ke sath linked.")
st.markdown("---")

try:
    conn = sqlite3.connect("sales_history.db")
    df_outputs = pd.read_sql("SELECT * FROM output_files_ledger ORDER BY id DESC", conn)
    df_history = pd.read_sql("SELECT * FROM history_logs ORDER BY id DESC", conn)
    df_manifests = pd.read_sql("SELECT * FROM dispatch_manifests ORDER BY id DESC", conn)
    conn.close()
except Exception as e:
    st.error(f"Database connection error: {str(e)}")
    df_outputs = pd.DataFrame()
    df_history = pd.DataFrame()
    df_manifests = pd.DataFrame()

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Demand & Plan Overview", 
    "📋 Dispatch Manifest & Vehicle Allocation", 
    "🚚 Live Dispatch Tracking", 
    "🗄️ Dispatch History & Audit"
])

with tab1:
    st.subheader("📈 Demand Summary")
    if not df_history.empty:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Batches", len(df_history))
        col2.metric("Cumulative Qty", f"{df_history['total_qty'].sum():,.0f} Units")
        col3.metric("System Status", "🟢 Connected")
        st.dataframe(df_history, use_container_width=True)
    else:
        st.warning("⚠️ Abhi tak koi demand process nahi ki gayi hai.")

with tab2:
    st.subheader("🚚 Create Dispatch Manifest")
    if not df_outputs.empty:
        file_options = df_outputs['file_name'].tolist()
        with st.form("dispatch_form"):
            selected_file = st.selectbox("Select Output File", file_options)
            c1, c2 = st.columns(2)
            with c1:
                transporter = st.text_input("Transporter Name")
                vehicle_no = st.text_input("Vehicle Number")
            with c2:
                driver_name = st.text_input("Driver Name")
                driver_phone = st.text_input("Driver Phone")
            status = st.selectbox("Status", ["Scheduled", "Loading", "Dispatched", "In Transit", "Delivered"])
            remarks = st.text_area("Remarks")
            if st.form_submit_button("💾 Save Manifest"):
                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO dispatch_manifests (batch_date, file_name, transporter_name, vehicle_no, driver_name, driver_phone, dispatch_status, remarks, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (get_ist_now().strftime("%Y-%m-%d"), selected_file, transporter, vehicle_no, driver_name, driver_phone, status, remarks, get_ist_now().strftime("%Y-%m-%d %H:%M:%S")))
                conn.commit()
                conn.close()
                st.success("✅ Manifest saved successfully!")
                st.rerun()
    else:
        st.info("No generated output files available yet.")

with tab3:
    st.subheader("📍 Active Dispatches")
    if not df_manifests.empty:
        st.dataframe(df_manifests, use_category_width=True if 'use_category_width' in dir(st) else True)
    else:
        st.info("No manifests created yet.")

with tab4:
    st.subheader("🗄️ Dispatch Audit Logs")
    if not df_manifests.empty:
        st.dataframe(df_manifests, use_container_width=True)
    else:
        st.info("No logs available.")
