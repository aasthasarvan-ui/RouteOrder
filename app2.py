import streamlit as st
import pandas as pd
import openpyxl
import datetime
import pytz
import io
import re
import zipfile
import sqlite3
import smtplib
import json
import urllib.parse
from email.message import EmailMessage
from fpdf import FPDF
import streamlit.components.v1 as components

# ==========================================
# PAGE CONFIGURATION & THEME ENGINE
# ==========================================
st.set_page_config(
    page_title="Enterprise Dispatch Planning Hub", 
    page_icon="🚚", 
    layout="wide",
    initial_sidebar_state="expanded"
)

THEMES = {
    "💼 Enterprise Navy": {
        "bg": "#f4f6f9", "text": "#1f2937", "card_bg": "#ffffff", "border": "#cbd5e1",
        "btn_bg": "#1e3a8a", "btn_hover": "#1d4ed8", "primary": "#2563eb", "input_bg": "#ffffff", "input_text": "#1f2937"
    },
    "🌙 Modern Dark Logistics": {
        "bg": "#0b0f19", "text": "#f3f4f6", "card_bg": "#1f2937", "border": "#374151",
        "btn_bg": "#374151", "btn_hover": "#4b5563", "primary": "#3b82f6", "input_bg": "#111827", "input_text": "#f3f4f6"
    },
    "🌲 Supply Chain Emerald": {
        "bg": "#f0fdf4", "text": "#14532d", "card_bg": "#dcfce7", "border": "#86efac",
        "btn_bg": "#16a34a", "btn_hover": "#15803d", "primary": "#22c55e", "input_bg": "#ffffff", "input_text": "#14532d"
    }
}

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

# Default Session State
if "selected_theme" not in st.session_state:
    st.session_state.selected_theme = "💼 Enterprise Navy"

t = THEMES[st.session_state.selected_theme]

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: {t['bg']}; color: {t['text']}; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{ color: {t['text']} !important; }}
        input, textarea, select {{ background-color: {t['input_bg']} !important; color: {t['input_text']} !important; border: 1px solid {t['border']} !important; }}
        .stButton>button {{ width: 100%; height: 38px; background-color: {t['btn_bg']} !important; color: #ffffff !important; font-size: 13px !important; font-weight: 600 !important; border-radius: 4px; border: 1px solid {t['border']}; }}
        button[kind="primary"] {{ background-color: {t['primary']} !important; color: #ffffff !important; }}
        div[data-testid="stExpander"] {{ background-color: {t['card_bg']}; border: 1px solid {t['border']}; border-radius: 4px; }}
        div[data-testid="stDataFrame"] {{ border: 1px solid {t['border']}; border-radius: 4px; background-color: {t['card_bg']}; }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==========================================
# DATABASE INITIALIZATION
# ==========================================
def init_dispatch_db():
    conn = sqlite3.connect("dispatch_logistics.db")
    cursor = conn.cursor()
    
    # 1. Fleet Master
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fleet_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_no TEXT UNIQUE,
            vehicle_type TEXT,
            capacity_bags INTEGER,
            capacity_mt REAL,
            transporter_name TEXT,
            driver_name TEXT,
            driver_phone TEXT,
            status TEXT DEFAULT 'Available'
        )
    """)
    
    # 2. Loading Bays Master
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS loading_bays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bay_no TEXT UNIQUE,
            bay_name TEXT,
            status TEXT DEFAULT 'Open'
        )
    """)
    
    # 3. Dispatch Trips Master
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS dispatch_trips (
            trip_id TEXT PRIMARY KEY,
            trip_date TEXT,
            route_no TEXT,
            vehicle_no TEXT,
            transporter_name TEXT,
            driver_name TEXT,
            driver_phone TEXT,
            loading_bay TEXT,
            total_bags REAL,
            total_weight_mt REAL,
            capacity_utilization_pct REAL,
            status TEXT,
            created_at TEXT
        )
    """)
    
    # 4. Trip Line Items (Agency & SKU Allocation)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS trip_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id TEXT,
            file_name TEXT,
            order_no TEXT,
            agency_no TEXT,
            route_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            allocated_bags REAL,
            allocated_weight_mt REAL,
            delivery_seq INTEGER,
            status TEXT,
            FOREIGN KEY (trip_id) REFERENCES dispatch_trips(trip_id)
        )
    """)
    
    # 5. Seed default fleet if empty
    cursor.execute("SELECT COUNT(*) FROM fleet_master")
    if cursor.fetchone()[0] == 0:
        default_fleet = [
            ('PB-10-AZ-1122', '10 Wheeler Truck', 400, 20.0, 'National Logistics', 'Gurpreet Singh', '9876543210', 'Available'),
            ('PB-08-BX-4455', '12 Wheeler Multi-Axle', 600, 30.0, 'Speedway Cargo', 'Baljit Sharma', '9812345678', 'Available'),
            ('PB-29-CD-9900', 'Canter / Eicher', 200, 10.0, 'Punjab Roadlines', 'Ramesh Kumar', '9823456789', 'Available'),
            ('PB-11-GH-3321', '14 Wheeler Heavy', 800, 40.0, 'Apex Transporters', 'Jarnail Singh', '9834567890', 'Available')
        ]
        cursor.executemany("""
            INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, default_fleet)

    # 6. Seed loading bays if empty
    cursor.execute("SELECT COUNT(*) FROM loading_bays")
    if cursor.fetchone()[0] == 0:
        default_bays = [
            ('BAY-01', 'North Plant Main Gate', 'Open'),
            ('BAY-02', 'Storage Silo Bay 2', 'Open'),
            ('BAY-03', 'Express Bulk Bay 3', 'Open')
        ]
        cursor.executemany("INSERT INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, ?)", default_bays)
        
    conn.commit()
    conn.close()

init_dispatch_db()

# ==========================================
# HELPER FUNCTIONS: DEMAND EXTRACTION
# ==========================================
def extract_pending_demand_from_sales_db():
    """Extracts processed orders from sales_history.db output_files_ledger"""
    pending_rows = []
    try:
        conn = sqlite3.connect("sales_history.db")
        cursor = conn.cursor()
        cursor.execute("SELECT file_name, file_type, file_data, created_at FROM output_files_ledger")
        rows = cursor.fetchall()
        conn.close()

        for fname, ftype, fdata, fdate in rows:
            try:
                wb = openpyxl.load_workbook(io.BytesIO(fdata), data_only=True)
                ws = wb["Order Data"] if "Order Data" in wb.sheetnames else wb.active
                for r in range(6, ws.max_row + 1):
                    order_no = ws.cell(row=r, column=2).value
                    dr_code = ws.cell(row=r, column=7).value
                    ref_no = ws.cell(row=r, column=9).value
                    fg_code = ws.cell(row=r, column=16).value
                    qty = ws.cell(row=r, column=19).value
                    route = ws.cell(row=r, column=26).value
                    agency = ws.cell(row=r, column=27).value
                    
                    if qty is not None and str(qty).strip() != "":
                        try:
                            f_qty = float(qty)
                            if f_qty > 0:
                                pending_rows.append({
                                    "Source File": fname,
                                    "File Type": ftype,
                                    "Order No": str(order_no),
                                    "Route No": str(route),
                                    "Agency No": str(agency),
                                    "DR Code": str(dr_code),
                                    "FG Code": str(fg_code),
                                    "Bags Qty": f_qty,
                                    "Weight (MT)": f_qty * 0.05, # Assuming 50 kg bag standard = 0.05 MT
                                    "Order Ref": str(ref_no)
                                })
                        except ValueError:
                            pass
            except Exception:
                pass
    except Exception as e:
        pass
    return pd.DataFrame(pending_rows)

def extract_demand_from_uploaded_excel(files):
    """Directly extracts demand from uploaded output or raw excel files"""
    rows = []
    for f in files:
        fbytes = f.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(fbytes), data_only=True)
        ws = wb["Order Data"] if "Order Data" in wb.sheetnames else wb.active
        for r in range(6, ws.max_row + 1):
            order_no = ws.cell(row=r, column=2).value
            dr_code = ws.cell(row=r, column=7).value
            ref_no = ws.cell(row=r, column=9).value
            fg_code = ws.cell(row=r, column=16).value
            qty = ws.cell(row=r, column=19).value
            route = ws.cell(row=r, column=26).value
            agency = ws.cell(row=r, column=27).value
            
            if qty is not None and str(qty).strip() != "":
                try:
                    f_qty = float(qty)
                    if f_qty > 0:
                        rows.append({
                            "Source File": f.name,
                            "File Type": "Direct Upload",
                            "Order No": str(order_no),
                            "Route No": str(route),
                            "Agency No": str(agency),
                            "DR Code": str(dr_code),
                            "FG Code": str(fg_code),
                            "Bags Qty": f_qty,
                            "Weight (MT)": f_qty * 0.05,
                            "Order Ref": str(ref_no)
                        })
                except ValueError:
                    pass
    return pd.DataFrame(rows)

# ==========================================
# SIDEBAR NAVIGATION & SETTINGS
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/delivery-truck.png", width=60)
    st.title("Logistics Command Hub")
    
    app_mode = st.radio(
        "Navigation Menu",
        ["🚚 Dispatch Trip Planner", "📋 Active & Completed Trips", "🚛 Fleet & Bay Master", "📊 Logistics KPI Analytics"]
    )
    
    st.markdown("---")
    st.subheader("⚙️ Theme")
    st.session_state.selected_theme = st.selectbox("Interface Theme", list(THEMES.keys()), index=0)
    
    st.markdown("---")
    st.subheader("📬 Notifications Config")
    transporter_email = st.text_input("Transporter Email", value="dispatch.logistics@enterprise.com")
    logistics_whatsapp = st.text_input("Logistics Contact (WhatsApp)", value="919876543210")

# ==========================================
# MODULE 1: DISPATCH TRIP PLANNER
# ==========================================
if app_mode == "🚚 Dispatch Trip Planner":
    st.title("🚚 Real-Time Dispatch Planning & Vehicle Allocation")
    st.markdown("Sales orders ke processed data ko select karke Route-wise truck capacity optimize kijiye aur Loading Slips generate karein.")
    
    # 1. Choose Data Source
    src_choice = st.radio("Select Inbound Demand Source:", ["📂 Load from Sales Order Automation Database (sales_history.db)", "📤 Upload Output/Demand Excel Files Manually"], horizontal=True)
    
    df_demand = pd.DataFrame()
    if "Load from Sales" in src_choice:
        df_demand = extract_pending_demand_from_sales_db()
        if df_demand.empty:
            st.info("ℹ️ `sales_history.db` mein koi output files nahi mili. Kripya pehle Sales Order Hub run karein ya manual Excel upload karein.")
    else:
        uploaded_outputs = st.file_uploader("Upload Generated Output Excel Files (*_Valid.xlsx / Output.xlsx)", type=["xlsx"], accept_multiple_files=True)
        if uploaded_outputs:
            df_demand = extract_demand_from_uploaded_excel(uploaded_outputs)

    if not df_demand.empty:
        # Check already allocated items
        conn = sqlite3.connect("dispatch_logistics.db")
        df_allocated = pd.read_sql("SELECT DISTINCT order_no, agency_no, route_no, fg_code FROM trip_order_items", conn)
        conn.close()

        st.markdown("---")
        st.subheader("📦 Inbound Demand Overview & Route Clustering")
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        total_bags_demand = df_demand["Bags Qty"].sum()
        total_tonnage = df_demand["Weight (MT)"].sum()
        unique_routes = df_demand["Route No"].nunique()
        unique_agencies = df_demand["Agency No"].nunique()
        
        col_m1.metric("Total Order Bags", f"{total_bags_demand:,.0f} Bags")
        col_m2.metric("Total Tonnage", f"{total_tonnage:,.2f} MT")
        col_m3.metric("Routes in Demand", unique_routes)
        col_m4.metric("Total Agencies", unique_agencies)
        
        # Route-wise group table
        route_summary = df_demand.groupby("Route No").agg({
            "Agency No": "nunique",
            "Bags Qty": "sum",
            "Weight (MT)": "sum"
        }).reset_index().rename(columns={"Agency No": "Total Agencies"})
        
        st.dataframe(route_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🚛 Intelligent Vehicle Allocation & Trip Formulation")
        
        col_p1, col_p2 = st.columns([1, 1])
        with col_p1:
            selected_route = st.selectbox("1. Select Route for Dispatch Trip", route_summary["Route No"].tolist())
            
            # Filter demand for selected route
            route_df = df_demand[df_demand["Route No"] == str(selected_route)].copy()
            st.markdown(f"**Total Bags on Route {selected_route}:** `{route_df['Bags Qty'].sum():,.0f}` | **Weight:** `{route_df['Weight (MT)'].sum():,.2f} MT`")
            
            # Fetch Available Vehicles
            conn_fl = sqlite3.connect("dispatch_logistics.db")
            available_vehicles = pd.read_sql("SELECT vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone FROM fleet_master WHERE status='Available'", conn_fl)
            available_bays = pd.read_sql("SELECT bay_no, bay_name FROM loading_bays WHERE status='Open'", conn_fl)
            conn_fl.close()
            
            veh_options = [f"{r['vehicle_no']} | {r['vehicle_type']} (Cap: {r['capacity_bags']} Bags / {r['capacity_mt']} MT)" for _, r in available_vehicles.iterrows()]
            selected_veh_str = st.selectbox("2. Assign Available Vehicle", veh_options if veh_options else ["No Vehicles Available"])
            selected_bay = st.selectbox("3. Assign Loading Bay", [f"{r['bay_no']} - {r['bay_name']}" for _, r in available_bays.iterrows()])

        with col_p2:
            st.markdown("#### 📋 Order Selection & Sequence")
            selected_agencies = st.multiselect("Filter Agencies to Load in this Trip:", route_df["Agency No"].unique().tolist(), default=route_df["Agency No"].unique().tolist())
            
            filtered_trip_demand = route_df[route_df["Agency No"].isin(selected_agencies)]
            trip_bags = filtered_trip_demand["Bags Qty"].sum()
            trip_weight = filtered_trip_demand["Weight (MT)"].sum()
            
            if selected_veh_str != "No Vehicles Available":
                veh_no = selected_veh_str.split(" | ")[0]
                veh_row = available_vehicles[available_vehicles["vehicle_no"] == veh_no].iloc[0]
                cap_bags = veh_row["capacity_bags"]
                utilization_pct = (trip_bags / cap_bags * 100) if cap_bags > 0 else 0
                
                st.metric("Allocated Trip Bags", f"{trip_bags:,.0f} / {cap_bags} Bags", f"{utilization_pct:.1f}% Utilization")
                if utilization_pct > 100:
                    st.error(f"🚨 **Overload Alert:** Capacity exceeded by {trip_bags - cap_bags:,.0f} bags! Please split the trip or choose a larger vehicle.")
                elif utilization_pct < 70:
                    st.warning("⚠️ **Low Utilization Warning:** Under 70% capacity.")
                else:
                    st.success("🟢 **Optimal Load Allocation!**")

        st.markdown("---")
        if st.button("🚀 Confirm & Generate Dispatch Trip (Gate Pass & Loading Slip)", type="primary"):
            if selected_veh_str == "No Vehicles Available":
                st.error("❌ Kripya pehle fleet master mein vehicle add ya free karein.")
            elif filtered_trip_demand.empty:
                st.error("❌ Is trip ke liye koi orders select nahi kiye gaye.")
            else:
                veh_no = selected_veh_str.split(" | ")[0]
                veh_row = available_vehicles[available_vehicles["vehicle_no"] == veh_no].iloc[0]
                bay_code = selected_bay.split(" - ")[0]
                
                ist_now = get_ist_now()
                trip_id = f"TRIP-{selected_route}-{ist_now.strftime('%Y%m%d%H%M%S')}"
                
                conn = sqlite3.connect("dispatch_logistics.db")
                cursor = conn.cursor()
                
                # Insert trip master
                cursor.execute("""
                    INSERT INTO dispatch_trips (trip_id, trip_date, route_no, vehicle_no, transporter_name, driver_name, driver_phone, loading_bay, total_bags, total_weight_mt, capacity_utilization_pct, status, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    trip_id,
                    ist_now.strftime("%Y-%m-%d"),
                    str(selected_route),
                    veh_no,
                    veh_row["transporter_name"],
                    veh_row["driver_name"],
                    veh_row["driver_phone"],
                    bay_code,
                    trip_bags,
                    trip_weight,
                    round((trip_bags / veh_row["capacity_bags"] * 100), 2),
                    "Planned",
                    ist_now.strftime("%Y-%m-%d %H:%M:%S")
                ))
                
                # Insert trip items
                items_to_insert = []
                for seq, (_, row) in enumerate(filtered_trip_demand.iterrows(), 1):
                    items_to_insert.append((
                        trip_id,
                        row["Source File"],
                        row["Order No"],
                        str(row["Agency No"]),
                        str(row["Route No"]),
                        str(row["DR Code"]),
                        str(row["FG Code"]),
                        row["Bags Qty"],
                        row["Weight (MT)"],
                        seq,
                        "Assigned"
                    ))
                    
                cursor.executemany("""
                    INSERT INTO trip_order_items (trip_id, file_name, order_no, agency_no, route_no, dr_code, fg_code, allocated_bags, allocated_weight_mt, delivery_seq, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, items_to_insert)
                
                # Update vehicle status
                cursor.execute("UPDATE fleet_master SET status='Assigned to Trip' WHERE vehicle_no=?", (veh_no,))
                
                conn.commit()
                conn.close()
                
                st.success(f"🎉 **Dispatch Trip '{trip_id}' Created Successfully!** Gate Pass & Loading Slip generated.")
                st.rerun()

# ==========================================
# MODULE 2: ACTIVE & COMPLETED TRIPS
# ==========================================
elif app_mode == "📋 Active & Completed Trips":
    st.title("📋 Active Trips, Loading Slips & Gate Pass Management")
    
    conn = sqlite3.connect("dispatch_logistics.db")
    df_trips = pd.read_sql("SELECT * FROM dispatch_trips ORDER BY created_at DESC", conn)
    conn.close()
    
    if not df_trips.empty:
        st.dataframe(df_trips, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🔍 Trip Details & Document Generation")
        
        selected_trip_id = st.selectbox("Select Trip ID to Inspect/Print:", df_trips["trip_id"].tolist())
        trip_row = df_trips[df_trips["trip_id"] == selected_trip_id].iloc[0]
        
        conn = sqlite3.connect("dispatch_logistics.db")
        df_trip_items = pd.read_sql("SELECT agency_no, dr_code, fg_code, allocated_bags, allocated_weight_mt, delivery_seq, status FROM trip_order_items WHERE trip_id=? ORDER BY delivery_seq ASC", conn, params=(selected_trip_id,))
        conn.close()
        
        c_t1, c_t2, c_t3, c_t4 = st.columns(4)
        c_t1.metric("Trip ID", trip_row["trip_id"])
        c_t2.metric("Vehicle No", trip_row["vehicle_no"])
        c_t3.metric("Driver", f"{trip_row['driver_name']} ({trip_row['driver_phone']})")
        c_t4.metric("Status", trip_row["status"])
        
        st.markdown("##### 📦 Material Loading Manifest (Agency-wise Sequence):")
        st.dataframe(df_trip_items, use_container_width=True)
        
        # Action Buttons
        col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
        
        with col_btn1:
            # Generate Loading Slip PDF
            try:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(190, 10, "ENTERPRISE DISPATCH & LOADING SLIP", ln=True, align="C")
                pdf.set_font("Arial", "", 10)
                pdf.cell(190, 6, f"Generated On (IST): {get_ist_now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
                pdf.ln(5)
                
                pdf.set_font("Arial", "B", 11)
                pdf.cell(95, 7, f"Trip ID: {trip_row['trip_id']}", border=1)
                pdf.cell(95, 7, f"Date: {trip_row['trip_date']}", border=1, ln=True)
                pdf.cell(95, 7, f"Vehicle No: {trip_row['vehicle_no']}", border=1)
                pdf.cell(95, 7, f"Route No: {trip_row['route_no']}", border=1, ln=True)
                pdf.cell(95, 7, f"Transporter: {trip_row['transporter_name']}", border=1)
                pdf.cell(95, 7, f"Loading Bay: {trip_row['loading_bay']}", border=1, ln=True)
                pdf.cell(95, 7, f"Driver Name: {trip_row['driver_name']}", border=1)
                pdf.cell(95, 7, f"Driver Phone: {trip_row['driver_phone']}", border=1, ln=True)
                pdf.ln(5)
                
                pdf.set_font("Arial", "B", 10)
                pdf.cell(20, 8, "Seq", border=1)
                pdf.cell(30, 8, "Agency", border=1)
                pdf.cell(40, 8, "DR Code", border=1)
                pdf.cell(45, 8, "FG Code", border=1)
                pdf.cell(30, 8, "Bags Qty", border=1)
                pdf.cell(25, 8, "Weight (MT)", border=1, ln=True)
                
                pdf.set_font("Arial", "", 10)
                for _, itm in df_trip_items.iterrows():
                    pdf.cell(20, 7, str(itm["delivery_seq"]), border=1)
                    pdf.cell(30, 7, str(itm["agency_no"]), border=1)
                    pdf.cell(40, 7, str(itm["dr_code"]), border=1)
                    pdf.cell(45, 7, str(itm["fg_code"]), border=1)
                    pdf.cell(30, 7, f"{itm['allocated_bags']:,.0f}", border=1)
                    pdf.cell(25, 7, f"{itm['allocated_weight_mt']:,.2f}", border=1, ln=True)
                    
                pdf.set_font("Arial", "B", 10)
                pdf.cell(135, 8, "TOTALS", border=1, align="R")
                pdf.cell(30, 8, f"{trip_row['total_bags']:,.0f}", border=1)
                pdf.cell(25, 8, f"{trip_row['total_weight_mt']:,.2f}", border=1, ln=True)
                
                pdf.ln(15)
                pdf.cell(60, 6, "Driver Signature: _________", ln=False)
                pdf.cell(65, 6, "Security Gate In: _________", ln=False)
                pdf.cell(65, 6, "Dispatch Supervisor: _________", ln=True)
                
                pdf_bytes = bytes(pdf.output())
                st.download_button(
                    label="📄 Download Loading Slip (PDF)",
                    data=pdf_bytes,
                    file_name=f"Loading_Slip_{selected_trip_id}.pdf",
                    mime="application/pdf",
                    key="btn_pdf_trip"
                )
            except Exception as e:
                st.error(f"PDF error: {str(e)}")

        with col_btn2:
            # WhatsApp dispatch alert
            wa_text = f"🚛 Enterprise Dispatch Alert!\nTrip ID: {trip_row['trip_id']}\nVehicle: {trip_row['vehicle_no']}\nDriver: {trip_row['driver_name']} ({trip_row['driver_phone']})\nRoute: {trip_row['route_no']}\nTotal Bags: {trip_row['total_bags']} ({trip_row['total_weight_mt']} MT)\nBay: {trip_row['loading_bay']}"
            wa_url = f"https://wa.me/{logistics_whatsapp}?text={urllib.parse.quote(wa_text)}"
            st.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer;">📱 WhatsApp Driver & Gate</button></a>', unsafe_allow_html=True)

        with col_btn3:
            # Status update dropdown
            new_status = st.selectbox("Update Trip Status:", ["Planned", "Loading in Progress", "Gate Out / Dispatched", "Delivered"], index=["Planned", "Loading in Progress", "Gate Out / Dispatched", "Delivered"].index(trip_row["status"]))
            if st.button("🔄 Update Status"):
                conn = sqlite3.connect("dispatch_logistics.db")
                cur = conn.cursor()
                cur.execute("UPDATE dispatch_trips SET status=? WHERE trip_id=?", (new_status, selected_trip_id))
                if new_status == "Delivered":
                    cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_row["vehicle_no"],))
                conn.commit()
                conn.close()
                st.success("✅ Status updated!")
                st.rerun()

        with col_btn4:
            if st.button("🗑️ Cancel Trip & Release Truck"):
                conn = sqlite3.connect("dispatch_logistics.db")
                cur = conn.cursor()
                cur.execute("DELETE FROM trip_order_items WHERE trip_id=?", (selected_trip_id,))
                cur.execute("DELETE FROM dispatch_trips WHERE trip_id=?", (selected_trip_id,))
                cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_row["vehicle_no"],))
                conn.commit()
                conn.close()
                st.warning(f"Trip {selected_trip_id} cancelled & vehicle released.")
                st.rerun()

    else:
        st.info("No dispatch trips created yet. Use 'Dispatch Trip Planner' tab to create trips.")

# ==========================================
# MODULE 3: FLEET & BAY MASTER MANAGEMENT
# ==========================================
elif app_mode == "🚛 Fleet & Bay Master":
    st.title("🚛 Transporter Fleet & Loading Bay Master")
    
    conn = sqlite3.connect("dispatch_logistics.db")
    df_fleet = pd.read_sql("SELECT * FROM fleet_master", conn)
    df_bays = pd.read_sql("SELECT * FROM loading_bays", conn)
    conn.close()
    
    tab_f1, tab_f2 = st.tabs(["🚛 Fleet Management", "🏭 Loading Bays"])
    
    with tab_f1:
        st.subheader("Active Transporter Fleet")
        st.dataframe(df_fleet, use_container_width=True)
        
        with st.expander("➕ Add New Vehicle to Fleet"):
            c_v1, c_v2, c_v3 = st.columns(3)
            with c_v1:
                new_v_no = st.text_input("Vehicle No (e.g., PB-10-XY-9988)")
                new_v_type = st.selectbox("Vehicle Type", ["10 Wheeler Truck", "12 Wheeler Multi-Axle", "14 Wheeler Heavy", "Canter / Eicher", "Mini Truck"])
            with c_v2:
                new_v_cap_bags = st.number_input("Capacity (Bags)", min_value=50, max_value=2000, value=500, step=50)
                new_v_cap_mt = st.number_input("Capacity (Metric Tons)", min_value=2.0, max_value=100.0, value=25.0, step=1.0)
            with c_v3:
                new_v_trans = st.text_input("Transporter Name", "Apex Logistics")
                new_v_driver = st.text_input("Driver Name", "Sukhdev Singh")
                new_v_phone = st.text_input("Driver Phone", "9876501234")
                
            if st.button("➕ Save Vehicle"):
                if new_v_no:
                    try:
                        conn = sqlite3.connect("dispatch_logistics.db")
                        cur = conn.cursor()
                        cur.execute("""
                            INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, 'Available')
                        """, (new_v_no, new_v_type, new_v_cap_bags, new_v_cap_mt, new_v_trans, new_v_driver, new_v_phone))
                        conn.commit()
                        conn.close()
                        st.success(f"Vehicle {new_v_no} added successfully!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error: {str(ex)}")

    with tab_f2:
        st.subheader("Plant Loading Bay Setup")
        st.dataframe(df_bays, use_container_width=True)
        
        with st.expander("➕ Add Loading Bay"):
            b1, b2 = st.columns(2)
            with b1:
                new_bay_no = st.text_input("Bay Code (e.g., BAY-04)")
            with b2:
                new_bay_name = st.text_input("Bay Location Name", "South Silo Discharge 4")
            if st.button("➕ Add Bay"):
                if new_bay_no:
                    conn = sqlite3.connect("dispatch_logistics.db")
                    cur = conn.cursor()
                    cur.execute("INSERT OR REPLACE INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, 'Open')", (new_bay_no, new_bay_name))
                    conn.commit()
                    conn.close()
                    st.success("Bay added!")
                    st.rerun()

# ==========================================
# MODULE 4: LOGISTICS KPI ANALYTICS
# ==========================================
elif app_mode == "📊 Logistics KPI Analytics":
    st.title("📊 Supply Chain & Dispatch Analytics Dashboard")
    
    conn = sqlite3.connect("dispatch_logistics.db")
    df_trips = pd.read_sql("SELECT * FROM dispatch_trips", conn)
    df_items = pd.read_sql("SELECT * FROM trip_order_items", conn)
    conn.close()
    
    if not df_trips.empty:
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        col_k1.metric("Total Trips Planned", len(df_trips))
        col_k2.metric("Total Dispatched Bags", f"{df_trips['total_bags'].sum():,.0f}")
        col_k3.metric("Average Fleet Utilization", f"{df_trips['capacity_utilization_pct'].mean():.1f}%")
        col_k4.metric("Active Loading Trips", len(df_trips[df_trips['status'] != 'Delivered']))
        
        st.markdown("---")
        t_col1, t_col2 = st.columns(2)
        with t_col1:
            st.markdown("##### 🚛 Route-wise Dispatched Tonnage (MT)")
            st.bar_chart(df_trips.groupby("route_no")["total_weight_mt"].sum())
        with t_col2:
            st.markdown("##### 📦 Transporter-wise Bag Share")
            st.bar_chart(df_trips.groupby("transporter_name")["total_bags"].sum())
    else:
        st.info("No dispatch data available to build analytics.")
