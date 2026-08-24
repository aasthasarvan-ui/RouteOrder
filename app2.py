import streamlit as st
import pandas as pd
import openpyxl
import datetime
import pytz
import io
import urllib.parse
import sqlite3
from fpdf import FPDF

# ==========================================
# PAGE CONFIGURATION & THEME
# ==========================================
st.set_page_config(
    page_title="Enterprise Dispatch & Logistics Hub", 
    page_icon="🚚", 
    layout="wide",
    initial_sidebar_state="expanded"
)

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

# Theme CSS
st.markdown(
    """
    <style>
        .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Segoe UI', Tahoma, sans-serif; }
        .stButton>button { width: 100%; border-radius: 6px; font-weight: 600; }
        div[data-testid="stDataFrame"] { border: 1px solid #cbd5e1; border-radius: 6px; }
        div[data-testid="stExpander"] { border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff; }
    </style>
    """,
    unsafe_allow_html=True
)

DB_NAME = "enterprise_logistics_hub.db"

# ==========================================
# DATABASE INITIALIZATION (MODULAR ARCHITECTURE)
# ==========================================
def init_all_databases():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    
    # 1. Archive Upload Database
    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploaded_files_archive (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            upload_timestamp TEXT,
            total_records INTEGER,
            file_size_kb REAL
        )
    """)

    # 2. Pending Orders Database
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pending_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT,
            order_no TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            bags_qty REAL,
            weight_mt REAL,
            order_ref TEXT,
            status TEXT DEFAULT 'Pending',
            uploaded_at TEXT
        )
    """)

    # 3. Fleet & Loading Bays
    cur.execute("""
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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS loading_bays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bay_no TEXT UNIQUE,
            bay_name TEXT,
            status TEXT DEFAULT 'Open'
        )
    """)

    # 4. Dispatch Trips & Loading Slips Database
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trip_loading_slips (
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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trip_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id TEXT,
            order_no TEXT,
            agency_no TEXT,
            route_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            allocated_bags REAL,
            allocated_weight_mt REAL,
            delivery_seq INTEGER,
            FOREIGN KEY (trip_id) REFERENCES trip_loading_slips(trip_id)
        )
    """)

    # 5. Daily Dispatch Sale Register Database
    cur.execute("""
        CREATE TABLE IF NOT EXISTS daily_dispatch_register (
            register_id INTEGER PRIMARY KEY AUTOINCREMENT,
            dispatch_date TEXT,
            trip_id TEXT,
            vehicle_no TEXT,
            transporter_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            order_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            dispatched_bags REAL,
            dispatched_weight_mt REAL,
            bay_no TEXT,
            dispatched_at TEXT
        )
    """)

    # Seed Default Fleet & Bays if empty
    cur.execute("SELECT COUNT(*) FROM fleet_master")
    if cur.fetchone()[0] == 0:
        cur.executemany("""
            INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, [
            ('PB-10-AZ-1122', '10 Wheeler Truck', 400, 20.0, 'National Logistics', 'Gurpreet Singh', '9876543210', 'Available'),
            ('PB-08-BX-4455', '12 Wheeler Multi-Axle', 600, 30.0, 'Speedway Cargo', 'Baljit Sharma', '9812345678', 'Available'),
            ('PB-29-CD-9900', 'Canter / Eicher', 200, 10.0, 'Punjab Roadlines', 'Ramesh Kumar', '9823456789', 'Available')
        ])

    cur.execute("SELECT COUNT(*) FROM loading_bays")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, ?)", [
            ('BAY-01', 'North Plant Main Gate', 'Open'),
            ('BAY-02', 'Storage Silo Bay 2', 'Open')
        ])

    conn.commit()
    conn.close()

init_all_databases()

# ==========================================
# EXCEL HELPER UTILITY
# ==========================================
def to_excel_download_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# ==========================================
# SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/delivery-truck.png", width=55)
    st.title("Logistics Master Hub")
    
    menu = st.radio(
        "Navigation",
        [
            "📥 Upload & Inbound Demand",
            "🚚 Route Dispatch & Trip Planner",
            "📋 Loading Slips & Active Trips",
            "📖 Daily Dispatch Sale Register",
            "⏳ Pending Orders Ledger",
            "🗄️ File Upload Archive",
            "🚛 Fleet & Bay Master"
        ]
    )
    st.markdown("---")
    whatsapp_no = st.text_input("WhatsApp Alert Contact", value="919876543210")

# ==========================================
# MODULE 1: UPLOAD & INBOUND DEMAND
# ==========================================
if menu == "📥 Upload & Inbound Demand":
    st.title("📥 Inbound Demand Excel Upload")
    st.markdown("Nayi sales output excel files upload karein. **Duplicate route records ko automatically filter kiya jayega.**")

    uploaded_files = st.file_uploader(
        "Upload Sales Output Excel Files (*.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    
    if uploaded_files and st.button("🚀 Process & Save into Pending Demand", type="primary"):
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Existing pending routes ko fetch karna to prevent duplication
        cur.execute("SELECT DISTINCT route_no FROM pending_orders WHERE status='Pending'")
        existing_routes = set(r[0] for r in cur.fetchall())
        
        total_added = 0
        total_skipped_dup = 0
        
        for f in uploaded_files:
            file_bytes = f.getvalue()
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb["Order Data"] if "Order Data" in wb.sheetnames else wb.active
            
            file_added_count = 0
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
                        route_str = str(route).strip() if route else "Unassigned"
                        
                        # Duplicate Route Check
                        if route_str in existing_routes:
                            total_skipped_dup += 1
                            continue
                            
                        if f_qty > 0:
                            cur.execute("""
                                INSERT INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?)
                            """, (
                                f.name,
                                str(order_no) if order_no else "N/A",
                                route_str,
                                str(agency).strip() if agency else "N/A",
                                str(dr_code) if dr_code else "N/A",
                                str(fg_code) if fg_code else "N/A",
                                f_qty,
                                round(f_qty * 0.05, 2),
                                str(ref_no) if ref_no else "N/A",
                                get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
                            ))
                            file_added_count += 1
                            total_added += 1
                    except ValueError:
                        pass
            
            # Archive File Record
            cur.execute("""
                INSERT OR REPLACE INTO uploaded_files_archive (file_name, upload_timestamp, total_records, file_size_kb)
                VALUES (?, ?, ?, ?)
            """, (f.name, get_ist_now().strftime("%Y-%m-%d %H:%M:%S"), file_added_count, round(len(file_bytes)/1024, 2)))
            
        conn.commit()
        conn.close()
        
        st.success(f"✅ Processing Complete: {total_added} naye orders add hue.")
        if total_skipped_dup > 0:
            st.warning(f"⚠️ {total_skipped_dup} orders skip kar diye gaye kyunki unka Route pehle se Pending Database me majood hai.")

# ==========================================
# MODULE 2: ROUTE DISPATCH & TRIP PLANNER
# ==========================================
elif menu == "🚚 Route Dispatch & Trip Planner":
    st.title("🚚 Route Dispatch Planning & Truck Allocation")
    
    conn = sqlite3.connect(DB_NAME)
    df_pending = pd.read_sql("SELECT * FROM pending_orders WHERE status='Pending'", conn)
    
    if df_pending.empty:
        st.info("ℹ️ Koi pending orders nahi hain. Kripya pehle Excel file upload karein.")
        conn.close()
    else:
        # Route Overview
        route_summary = df_pending.groupby("route_no").agg({
            "agency_no": "nunique",
            "bags_qty": "sum",
            "weight_mt": "sum"
        }).reset_index().rename(columns={"agency_no": "Agencies", "bags_qty": "Total Bags", "weight_mt": "Total MT"})
        
        st.subheader("📊 Pending Routes Summary")
        st.dataframe(route_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🚛 Assign Vehicle & Create Trip")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_route = st.selectbox("Select Route for Trip", route_summary["route_no"].tolist())
            route_orders = df_pending[df_pending["route_no"] == selected_route]
            
            avail_veh = pd.read_sql("SELECT * FROM fleet_master WHERE status='Available'", conn)
            avail_bays = pd.read_sql("SELECT * FROM loading_bays WHERE status='Open'", conn)
            
            veh_choices = [f"{r['vehicle_no']} | {r['vehicle_type']} (Cap: {r['capacity_bags']} Bags)" for _, r in avail_veh.iterrows()]
            selected_veh_str = st.selectbox("Select Vehicle", veh_choices if veh_choices else ["No Vehicles Available"])
            selected_bay = st.selectbox("Select Loading Bay", [f"{r['bay_no']} - {r['bay_name']}" for _, r in avail_bays.iterrows()])
            
        with col2:
            st.markdown("##### 📦 Select Agencies for this Trip")
            agencies = route_orders["agency_no"].unique().tolist()
            sel_agencies = st.multiselect("Filter Agencies:", agencies, default=agencies)
            
            filtered_orders = route_orders[route_orders["agency_no"].isin(sel_agencies)]
            trip_bags = filtered_orders["bags_qty"].sum()
            trip_mt = filtered_orders["weight_mt"].sum()
            
            if selected_veh_str != "No Vehicles Available":
                v_no = selected_veh_str.split(" | ")[0]
                v_row = avail_veh[avail_veh["vehicle_no"] == v_no].iloc[0]
                cap = v_row["capacity_bags"]
                util = (trip_bags / cap * 100) if cap > 0 else 0
                st.metric("Total Load", f"{trip_bags:,.0f} Bags ({trip_mt:.2f} MT)", f"{util:.1f}% Capacity")
                
        if st.button("🚀 Confirm Trip & Generate Loading Slip", type="primary"):
            if selected_veh_str == "No Vehicles Available" or filtered_orders.empty:
                st.error("❌ Valid Vehicle ya Orders select karein.")
            else:
                cur = conn.cursor()
                now = get_ist_now()
                trip_id = f"TRIP-{selected_route}-{now.strftime('%Y%m%d%H%M%S')}"
                v_no = selected_veh_str.split(" | ")[0]
                v_row = avail_veh[avail_veh["vehicle_no"] == v_no].iloc[0]
                bay_no = selected_bay.split(" - ")[0]
                
                # Insert to Trip Loading Slips Table
                cur.execute("""
                    INSERT INTO trip_loading_slips (trip_id, trip_date, route_no, vehicle_no, transporter_name, driver_name, driver_phone, loading_bay, total_bags, total_weight_mt, capacity_utilization_pct, status, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Planned', ?)
                """, (
                    trip_id, now.strftime("%Y-%m-%d"), selected_route, v_no,
                    v_row["transporter_name"], v_row["driver_name"], v_row["driver_phone"],
                    bay_no, trip_bags, trip_mt, round((trip_bags/v_row["capacity_bags"]*100), 2),
                    now.strftime("%Y-%m-%d %H:%M:%S")
                ))
                
                # Insert order items & Update Pending status
                for seq, (_, row) in enumerate(filtered_orders.iterrows(), 1):
                    cur.execute("""
                        INSERT INTO trip_order_items (trip_id, order_no, agency_no, route_no, dr_code, fg_code, allocated_bags, allocated_weight_mt, delivery_seq)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (trip_id, row["order_no"], row["agency_no"], row["route_no"], row["dr_code"], row["fg_code"], row["bags_qty"], row["weight_mt"], seq))
                    
                    cur.execute("UPDATE pending_orders SET status='Assigned' WHERE id=?", (row["id"],))
                
                cur.execute("UPDATE fleet_master SET status='Assigned' WHERE vehicle_no=?", (v_no,))
                conn.commit()
                st.success(f"🎉 Trip {trip_id} successfully created!")
                st.rerun()
        conn.close()

# ==========================================
# MODULE 3: LOADING SLIPS & ACTIVE TRIPS
# ==========================================
elif menu == "📋 Loading Slips & Active Trips":
    st.title("📋 Trip Slips & Active Dispatches")
    
    conn = sqlite3.connect(DB_NAME)
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips ORDER BY created_at DESC", conn)
    
    # Global Search
    search_q = st.text_input("🔍 Search Trip Slips (Trip ID, Vehicle, Route, Transporter):", "")
    if search_q:
        df_trips = df_trips[df_trips.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)]
    
    st.dataframe(df_trips, use_container_width=True)
    
    # Download & Multiple Delete
    col_d1, col_d2 = st.columns([1, 2])
    with col_d1:
        if not df_trips.empty:
            st.download_button("📥 Download Trips to Excel", to_excel_download_bytes(df_trips), "Dispatch_Trips.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
    with col_d2:
        with st.expander("🗑️ Multiple Delete Trips"):
            delete_trips = st.multiselect("Select Trip IDs to Delete:", df_trips["trip_id"].tolist() if not df_trips.empty else [])
            if st.button("Delete Selected Trips", type="secondary"):
                cur = conn.cursor()
                for tid in delete_trips:
                    # Vehicle release
                    v = cur.execute("SELECT vehicle_no FROM trip_loading_slips WHERE trip_id=?", (tid,)).fetchone()
                    if v:
                        cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (v[0],))
                    cur.execute("DELETE FROM trip_order_items WHERE trip_id=?", (tid,))
                    cur.execute("DELETE FROM trip_loading_slips WHERE trip_id=?", (tid,))
                conn.commit()
                st.success("Trips deleted & vehicles released.")
                st.rerun()

    if not df_trips.empty:
        st.markdown("---")
        st.subheader("📄 Print Slip & Dispatch Workflow")
        sel_trip = st.selectbox("Select Trip ID:", df_trips["trip_id"].tolist())
        trip_data = df_trips[df_trips["trip_id"] == sel_trip].iloc[0]
        items_df = pd.read_sql("SELECT * FROM trip_order_items WHERE trip_id=? ORDER BY delivery_seq", conn, params=(sel_trip,))
        
        st.write(items_df)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            # PDF Generation
            try:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 14)
                pdf.cell(190, 8, "ENTERPRISE DISPATCH LOADING SLIP", ln=True, align="C")
                pdf.set_font("Arial", "", 9)
                pdf.cell(190, 5, f"Trip ID: {trip_data['trip_id']} | Date: {trip_data['trip_date']}", ln=True, align="C")
                pdf.ln(4)
                
                pdf.set_font("Arial", "B", 9)
                pdf.cell(95, 6, f"Vehicle: {trip_data['vehicle_no']}", border=1)
                pdf.cell(95, 6, f"Route: {trip_data['route_no']}", border=1, ln=True)
                pdf.cell(95, 6, f"Transporter: {trip_data['transporter_name']}", border=1)
                pdf.cell(95, 6, f"Bay: {trip_data['loading_bay']}", border=1, ln=True)
                pdf.cell(95, 6, f"Driver: {trip_data['driver_name']} ({trip_data['driver_phone']})", border=1)
                pdf.cell(95, 6, f"Total Bags: {trip_data['total_bags']}", border=1, ln=True)
                pdf.ln(4)
                
                pdf.cell(15, 6, "Seq", 1)
                pdf.cell(35, 6, "Agency", 1)
                pdf.cell(40, 6, "DR Code", 1)
                pdf.cell(45, 6, "FG Code", 1)
                pdf.cell(30, 6, "Bags", 1)
                pdf.cell(25, 6, "MT", 1, ln=True)
                pdf.set_font("Arial", "", 8)
                for _, it in items_df.iterrows():
                    pdf.cell(15, 5, str(it["delivery_seq"]), 1)
                    pdf.cell(35, 5, str(it["agency_no"]), 1)
                    pdf.cell(40, 5, str(it["dr_code"]), 1)
                    pdf.cell(45, 5, str(it["fg_code"]), 1)
                    pdf.cell(30, 5, f"{it['allocated_bags']:,.0f}", 1)
                    pdf.cell(25, 5, f"{it['allocated_weight_mt']:,.2f}", 1, ln=True)
                    
                st.download_button("📄 Download PDF Slip", bytes(pdf.output()), f"Slip_{sel_trip}.pdf", "application/pdf")
            except Exception as e:
                st.error(f"PDF error: {e}")

        with c2:
            wa_text = f"Enterprise Dispatch: Trip {trip_data['trip_id']} | Vehicle {trip_data['vehicle_no']} | Route {trip_data['route_no']} | Bags {trip_data['total_bags']}"
            st.markdown(f'<a href="https://wa.me/{whatsapp_no}?text={urllib.parse.quote(wa_text)}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp Alert</button></a>', unsafe_allow_html=True)
            
        with c3:
            if trip_data["status"] != "Dispatched" and st.button("🏁 Mark Dispatched & Move to Daily Register", type="primary"):
                cur = conn.cursor()
                # Insert into Daily Dispatch Sale Register
                for _, it in items_df.iterrows():
                    cur.execute("""
                        INSERT INTO daily_dispatch_register (dispatch_date, trip_id, vehicle_no, transporter_name, route_no, agency_no, order_no, dr_code, fg_code, dispatched_bags, dispatched_weight_mt, bay_no, dispatched_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        trip_data["trip_date"], trip_data["trip_id"], trip_data["vehicle_no"],
                        trip_data["transporter_name"], trip_data["route_no"], it["agency_no"],
                        it["order_no"], it["dr_code"], it["fg_code"], it["allocated_bags"],
                        it["allocated_weight_mt"], trip_data["loading_bay"], get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
                    ))
                cur.execute("UPDATE trip_loading_slips SET status='Dispatched' WHERE trip_id=?", (sel_trip,))
                cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_data["vehicle_no"],))
                conn.commit()
                st.success("✅ Trip dispatched & recorded in Daily Dispatch Register!")
                st.rerun()
    conn.close()

# ==========================================
# MODULE 4: DAILY DISPATCH SALE REGISTER
# ==========================================
elif menu == "📖 Daily Dispatch Sale Register":
    st.title("📖 Daily Dispatch Sale Register Database")
    
    conn = sqlite3.connect(DB_NAME)
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register ORDER BY register_id DESC", conn)
    
    search_r = st.text_input("🔍 Search Register (Agency, Route, Order, Trip ID, Vehicle):", "")
    if search_r:
        df_reg = df_reg[df_reg.apply(lambda r: r.astype(str).str.contains(search_r, case=False).any(), axis=1)]
        
    st.dataframe(df_reg, use_container_width=True)
    
    c_r1, c_r2 = st.columns([1, 2])
    with c_r1:
        if not df_reg.empty:
            st.download_button("📥 Export Register to Excel", to_excel_download_bytes(df_reg), "Daily_Dispatch_Sale_Register.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c_r2:
        with st.expander("🗑️ Multiple Delete Records"):
            del_reg_ids = st.multiselect("Select Register IDs to Delete:", df_reg["register_id"].tolist() if not df_reg.empty else [])
            if st.button("Delete Selected Records"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM daily_dispatch_register WHERE register_id=?", [(i,) for i in del_reg_ids])
                conn.commit()
                st.success("Selected records deleted.")
                st.rerun()
    conn.close()

# ==========================================
# MODULE 5: PENDING ORDERS LEDGER
# ==========================================
elif menu == "⏳ Pending Orders Ledger":
    st.title("⏳ Pending Orders Database")
    
    conn = sqlite3.connect(DB_NAME)
    df_p = pd.read_sql("SELECT * FROM pending_orders ORDER BY id DESC", conn)
    
    search_p = st.text_input("🔍 Search Pending Orders (Order No, Route, Agency, FG Code):", "")
    if search_p:
        df_p = df_p[df_p.apply(lambda r: r.astype(str).str.contains(search_p, case=False).any(), axis=1)]
        
    st.dataframe(df_p, use_container_width=True)
    
    c_p1, c_p2 = st.columns([1, 2])
    with c_p1:
        if not df_p.empty:
            st.download_button("📥 Export Pending Orders to Excel", to_excel_download_bytes(df_p), "Pending_Orders.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c_p2:
        with st.expander("🗑️ Multiple Delete Pending Orders"):
            del_p_ids = st.multiselect("Select Order IDs to Delete:", df_p["id"].tolist() if not df_p.empty else [])
            if st.button("Delete Selected Orders"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM pending_orders WHERE id=?", [(i,) for i in del_p_ids])
                conn.commit()
                st.success("Selected pending orders removed.")
                st.rerun()
    conn.close()

# ==========================================
# MODULE 6: FILE UPLOAD ARCHIVE
# ==========================================
elif menu == "🗄️ File Upload Archive":
    st.title("🗄️ Uploaded Input File Archive")
    
    conn = sqlite3.connect(DB_NAME)
    df_arch = pd.read_sql("SELECT * FROM uploaded_files_archive ORDER BY id DESC", conn)
    
    search_a = st.text_input("🔍 Search Archive Files:", "")
    if search_a:
        df_arch = df_arch[df_arch.apply(lambda r: r.astype(str).str.contains(search_a, case=False).any(), axis=1)]
        
    st.dataframe(df_arch, use_container_width=True)
    
    c_a1, c_a2 = st.columns([1, 2])
    with c_a1:
        if not df_arch.empty:
            st.download_button("📥 Export Archive Log to Excel", to_excel_download_bytes(df_arch), "Uploaded_Archive_Logs.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c_a2:
        with st.expander("🗑️ Multiple Delete Archive Logs"):
            del_a_ids = st.multiselect("Select File IDs to Delete:", df_arch["id"].tolist() if not df_arch.empty else [])
            if st.button("Delete Selected Archive Logs"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM uploaded_files_archive WHERE id=?", [(i,) for i in del_a_ids])
                conn.commit()
                st.success("Selected archive records deleted.")
                st.rerun()
    conn.close()

# ==========================================
# MODULE 7: FLEET & BAY MASTER
# ==========================================
elif menu == "🚛 Fleet & Bay Master":
    st.title("🚛 Fleet & Loading Bay Masters")
    
    conn = sqlite3.connect(DB_NAME)
    df_fleet = pd.read_sql("SELECT * FROM fleet_master", conn)
    df_bays = pd.read_sql("SELECT * FROM loading_bays", conn)
    
    t1, t2 = st.tabs(["🚛 Fleet Master", "🏭 Loading Bays"])
    
    with t1:
        st.dataframe(df_fleet, use_container_width=True)
        c_f1, c_f2 = st.columns([1, 2])
        with c_f1:
            st.download_button("📥 Export Fleet to Excel", to_excel_download_bytes(df_fleet), "Fleet_Master.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c_f2:
            with st.expander("🗑️ Delete Vehicles"):
                del_v_ids = st.multiselect("Select Vehicle IDs:", df_fleet["id"].tolist())
                if st.button("Delete Vehicles"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM fleet_master WHERE id=?", [(i,) for i in del_v_ids])
                    conn.commit()
                    st.rerun()
                    
        with st.expander("➕ Add New Vehicle"):
            v1, v2, v3 = st.columns(3)
            with v1:
                v_no = st.text_input("Vehicle No (e.g. PB-10-XX-1234)")
                v_type = st.selectbox("Type", ["10 Wheeler Truck", "12 Wheeler Multi-Axle", "14 Wheeler Heavy", "Canter / Eicher"])
            with v2:
                v_bags = st.number_input("Bags Capacity", value=500, step=50)
                v_mt = st.number_input("MT Capacity", value=25.0, step=1.0)
            with v3:
                v_trans = st.text_input("Transporter Name")
                v_driver = st.text_input("Driver Name")
                v_phone = st.text_input("Driver Phone")
            if st.button("Save Vehicle"):
                if v_no:
                    cur = conn.cursor()
                    cur.execute("INSERT OR REPLACE INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status) VALUES (?, ?, ?, ?, ?, ?, ?, 'Available')", (v_no, v_type, v_bags, v_mt, v_trans, v_driver, v_phone))
                    conn.commit()
                    st.rerun()

    with t2:
        st.dataframe(df_bays, use_container_width=True)
        c_b1, c_b2 = st.columns([1, 2])
        with c_b1:
            st.download_button("📥 Export Bays to Excel", to_excel_download_bytes(df_bays), "Loading_Bays.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c_b2:
            with st.expander("🗑️ Delete Bays"):
                del_b_ids = st.multiselect("Select Bay IDs:", df_bays["id"].tolist())
                if st.button("Delete Bays"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM loading_bays WHERE id=?", [(i,) for i in del_b_ids])
                    conn.commit()
                    st.rerun()
                    
        with st.expander("➕ Add New Bay"):
            b1, b2 = st.columns(2)
            with b1:
                b_no = st.text_input("Bay No (e.g. BAY-03)")
            with b2:
                b_name = st.text_input("Bay Name / Location")
            if st.button("Save Bay"):
                if b_no:
                    cur = conn.cursor()
                    cur.execute("INSERT OR REPLACE INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, 'Open')", (b_no, b_name))
                    conn.commit()
                    st.rerun()
    conn.close()
