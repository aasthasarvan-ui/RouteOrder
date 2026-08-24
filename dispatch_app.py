import streamlit as st
import pandas as pd
import datetime
import pytz
import io
import sqlite3
import urllib.parse
from email.message import EmailMessage
import smtplib

# Page Configuration & Styling
st.set_page_config(
    page_title="Enterprise Plan & Dispatch Hub", 
    page_icon="🚚", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 8 Professional Enterprise Themes (Matching app.py)
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

# --- Initialize Dispatch Specific Tables in Shared Database ---
def init_dispatch_db():
    conn = sqlite3.connect("sales_history.db")
    cursor = conn.cursor()
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

# Session State Defaults for Dispatch App
if "dispatch_theme" not in st.session_state:
    st.session_state.dispatch_theme = "💼 Classic Enterprise Navy"

t = THEMES[st.session_state.dispatch_theme]

# Professional CSS Injection
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

# Top Control Panel for Theme & Settings
with st.expander("⚙️ Dispatch Control Panel & Theme Selection", expanded=False):
    st.selectbox("Select Theme", list(THEMES.keys()), key="dispatch_theme", index=list(THEMES.keys()).index(st.session_state.dispatch_theme))

st.title(f"🚚 Enterprise Plan, Demand & Dispatch Hub ({st.session_state.dispatch_theme})")
st.markdown("Demand app (`app.py`) ke generated outputs aur database (`sales_history.db`) ke sath seamlessly linked.")
st.markdown("---")

# Fetch Data from Shared Database
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

# Tabs for Plan, Demand & Dispatch Operations
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Demand & Plan Overview", 
    "📋 Dispatch Manifest & Vehicle Allocation", 
    "🚚 Live Dispatch Tracking", 
    "🗄️ Dispatch History & Audit"
])

# --- TAB 1: DEMAND & PLAN OVERVIEW ---
with tab1:
    st.subheader("📈 Demand Summary from app.py")
    if not df_history.empty:
        total_batches = len(df_history)
        total_qty_sum = df_history['total_qty'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Batches Processed", total_batches)
        col2.metric("Cumulative Demand Qty", f"{total_qty_sum:,.0f} Units")
        col3.metric("System Status", "🟢 Connected to app.py DB")
        
        st.markdown("---")
        st.markdown("##### Recent Demand Execution Logs")
        st.dataframe(df_history, use_container_width=True)
    else:
        st.warning("⚠️ Abhi tak `app.py` se koi demand process nahi ki gayi hai. Kripya pehle `app.py` chala kar files process karein.")

# --- TAB 2: DISPATCH MANIFEST & VEHICLE ALLOCATION ---
with tab2:
    st.subheader("🚚 Create Dispatch Manifest & Assign Vehicle")
    if not df_outputs.empty:
        file_options = df_outputs['file_name'].tolist()
        
        with st.form("dispatch_form"):
            selected_file = st.selectbox("Select Generated Output File", file_options)
            
            c1, c2 = st.columns(2)
            with c1:
                transporter = st.text_input("Transporter Name (e.g., VRL Logistics, Blue Dart)")
                vehicle_no = st.text_input("Vehicle Number (e.g., PB08AB1234)")
            with c2:
                driver_name = st.text_input("Driver Name")
                driver_phone = st.text_input("Driver Phone Number")
            
            status = st.selectbox("Initial Dispatch Status", ["Scheduled", "Loading", "Dispatched", "In Transit", "Delivered"])
            remarks = st.text_area("Dispatch Remarks / Notes")
            
            submitted = st.form_submit_button("💾 Save Dispatch Manifest")
            if submitted:
                if selected_file and vehicle_no:
                    try:
                        conn = sqlite3.connect("sales_history.db")
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO dispatch_manifests (batch_date, file_name, transporter_name, vehicle_no, driver_name, driver_phone, dispatch_status, remarks, updated_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (get_ist_now().strftime("%Y-%m-%d"), selected_file, transporter, vehicle_no, driver_name, driver_phone, status, remarks, get_ist_now().strftime("%Y-%m-%d %H:%M:%S")))
                        conn.commit()
                        conn.close()
                        st.success(f"✅ Manifest successfully created for file: {selected_file}!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error saving manifest: {str(ex)}")
                else:
                    st.warning("⚠️ Kripya File aur Vehicle Number zaroor bharein.")
    else:
        st.info("No generated files found in `output_files_ledger`. Process files in `app.py` first.")

# --- TAB 3: LIVE DISPATCH TRACKING & DOWNLOAD ---
with tab3:
    st.subheader("📍 Active Dispatches & Document Downloads")
    if not df_manifests.empty:
        st.dataframe(df_manifests, use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### Quick Download Output File for Specific Manifest")
        manifest_id = st.number_input("Enter Manifest ID to Download File", min_value=1, step=1, key="man_dl_id")
        if st.button("📥 Fetch & Download Linked Excel"):
            try:
                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.execute("SELECT file_name FROM dispatch_manifests WHERE id = ?", (manifest_id,))
                m_row = cursor.fetchone()
                if m_row:
                    f_name = m_row[0]
                    cursor.execute("SELECT file_data FROM output_files_ledger WHERE file_name = ?", (f_name,))
                    f_data_row = cursor.fetchone()
                    if f_data_row:
                        st.download_button(
                            label=f"💾 Save {f_name}",
                            data=f_data_row[0],
                            file_name=f_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Linked file binary data not found in ledger.")
                else:
                    st.warning("Invalid Manifest ID.")
                conn.close()
            except Exception as e:
                st.error(f"Error fetching file: {str(e)}")
    else:
        st.info("No dispatch manifests created yet.")

# --- TAB 4: DISPATCH HISTORY & AUDIT ---
with tab4:
    st.subheader("🗄️ Full Dispatch Audit Logs")
    if not df_manifests.empty:
        st.dataframe(df_manifests, use_container_width=True)
        
        csv_bytes = df_manifests.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Export Dispatch Audit to CSV",
            data=csv_bytes,
            file_name=f"Dispatch_Audit_Report_{get_ist_now().strftime('%Y-%m-%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No logs to display.")