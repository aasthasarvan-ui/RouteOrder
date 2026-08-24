import sqlite3
from datetime import datetime


# ============================================================
# DATABASE CONNECTION
# ============================================================

DB_NAME = "dispatch_system.db"


def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row

    # Foreign key enable
    conn.execute("PRAGMA foreign_keys = ON")

    return conn


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    # ========================================================
    # ROUTE PLAN
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS route_plan (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            plan_date TEXT NOT NULL,

            route_no TEXT NOT NULL,

            sequence_no INTEGER DEFAULT 1,

            agency_no TEXT NOT NULL,

            dr_code TEXT,

            customer_name TEXT,

            address TEXT,

            planned_qty REAL DEFAULT 0,

            vehicle_no TEXT,

            driver_name TEXT,

            planned_time TEXT,

            status TEXT DEFAULT 'Planned',

            remarks TEXT,

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,

            CHECK(planned_qty >= 0),

            CHECK(status IN (
                'Planned',
                'Assigned',
                'Dispatched',
                'Completed',
                'Cancelled'
            ))
        )
    """)


    # ========================================================
    # DISPATCH
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS dispatch (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            dispatch_date TEXT NOT NULL,

            route_no TEXT NOT NULL,

            agency_no TEXT NOT NULL,

            dr_code TEXT,

            fg_code TEXT,

            quantity REAL DEFAULT 0,

            vehicle_no TEXT,

            driver_name TEXT,

            dispatch_time TEXT,

            status TEXT DEFAULT 'Pending',

            remarks TEXT,

            route_plan_id INTEGER,

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,

            CHECK(quantity >= 0),

            CHECK(status IN (
                'Pending',
                'Loading',
                'Dispatched',
                'Delivered',
                'Cancelled'
            )),

            FOREIGN KEY(route_plan_id)
                REFERENCES route_plan(id)
                ON DELETE SET NULL
                ON UPDATE CASCADE
        )
    """)


    # ========================================================
    # VEHICLE MASTER
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vehicle_master (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            vehicle_no TEXT UNIQUE NOT NULL,

            vehicle_type TEXT,

            capacity REAL DEFAULT 0,

            driver_name TEXT,

            driver_mobile TEXT,

            status TEXT DEFAULT 'Available',

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            CHECK(capacity >= 0),

            CHECK(status IN (
                'Available',
                'Assigned',
                'Loading',
                'On Route',
                'Maintenance',
                'Inactive'
            ))
        )
    """)


    # ========================================================
    # DRIVER MASTER
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS driver_master (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            driver_name TEXT UNIQUE NOT NULL,

            mobile_no TEXT,

            license_no TEXT,

            status TEXT DEFAULT 'Active',

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            CHECK(status IN (
                'Active',
                'Inactive',
                'On Leave'
            ))
        )
    """)


    # ========================================================
    # DISPATCH STATUS HISTORY
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS dispatch_status_history (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            dispatch_id INTEGER NOT NULL,

            old_status TEXT,

            new_status TEXT NOT NULL,

            changed_at TEXT DEFAULT CURRENT_TIMESTAMP,

            remarks TEXT,

            FOREIGN KEY(dispatch_id)
                REFERENCES dispatch(id)
                ON DELETE CASCADE
        )
    """)


    # ========================================================
    # ROUTE PLAN STATUS HISTORY
    # ========================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS route_plan_status_history (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            route_plan_id INTEGER NOT NULL,

            old_status TEXT,

            new_status TEXT NOT NULL,

            changed_at TEXT DEFAULT CURRENT_TIMESTAMP,

            remarks TEXT,

            FOREIGN KEY(route_plan_id)
                REFERENCES route_plan(id)
                ON DELETE CASCADE
        )
    """)


    # ========================================================
    # ROUTE PLAN INDEXES
    # ========================================================

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_date
        ON route_plan(plan_date)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_route
        ON route_plan(route_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_agency
        ON route_plan(agency_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_dr_code
        ON route_plan(dr_code)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_vehicle
        ON route_plan(vehicle_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_plan_status
        ON route_plan(status)
    """)


    # ========================================================
    # DISPATCH INDEXES
    # ========================================================

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_date
        ON dispatch(dispatch_date)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_route
        ON dispatch(route_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_agency
        ON dispatch(agency_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_dr_code
        ON dispatch(dr_code)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_vehicle
        ON dispatch(vehicle_no)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_status
        ON dispatch(status)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_route_plan
        ON dispatch(route_plan_id)
    """)


    # ========================================================
    # VEHICLE INDEXES
    # ========================================================

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_vehicle_status
        ON vehicle_master(status)
    """)


    # ========================================================
    # DRIVER INDEXES
    # ========================================================

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_driver_status
        ON driver_master(status)
    """)


    # ========================================================
    # HISTORY INDEXES
    # ========================================================

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_dispatch_history
        ON dispatch_status_history(dispatch_id)
    """)

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_route_history
        ON route_plan_status_history(route_plan_id)
    """)


    # ========================================================
    # ROUTE PLAN STATUS TRIGGER
    # ========================================================

    cursor.execute("""
        CREATE TRIGGER IF NOT EXISTS trg_route_plan_status_history

        AFTER UPDATE OF status ON route_plan

        WHEN OLD.status <> NEW.status

        BEGIN

            INSERT INTO route_plan_status_history
            (
                route_plan_id,
                old_status,
                new_status,
                changed_at
            )

            VALUES
            (
                NEW.id,
                OLD.status,
                NEW.status,
                CURRENT_TIMESTAMP
            );

        END;
    """)


    # ========================================================
    # DISPATCH STATUS TRIGGER
    # ========================================================

    cursor.execute("""
        CREATE TRIGGER IF NOT EXISTS trg_dispatch_status_history

        AFTER UPDATE OF status ON dispatch

        WHEN OLD.status <> NEW.status

        BEGIN

            INSERT INTO dispatch_status_history
            (
                dispatch_id,
                old_status,
                new_status,
                changed_at
            )

            VALUES
            (
                NEW.id,
                OLD.status,
                NEW.status,
                CURRENT_TIMESTAMP
            );

        END;
    """)


    # ========================================================
    # UPDATED_AT TRIGGER - ROUTE PLAN
    # ========================================================

    cursor.execute("""
        CREATE TRIGGER IF NOT EXISTS trg_route_plan_updated_at

        AFTER UPDATE ON route_plan

        FOR EACH ROW

        BEGIN

            UPDATE route_plan
            SET updated_at = CURRENT_TIMESTAMP
            WHERE id = NEW.id;

        END;
    """)


    # ========================================================
    # UPDATED_AT TRIGGER - DISPATCH
    # ========================================================

    cursor.execute("""
        CREATE TRIGGER IF NOT EXISTS trg_dispatch_updated_at

        AFTER UPDATE ON dispatch

        FOR EACH ROW

        BEGIN

            UPDATE dispatch
            SET updated_at = CURRENT_TIMESTAMP
            WHERE id = NEW.id;

        END;
    """)


    # ========================================================
    # INITIAL DATABASE COMMIT
    # ========================================================

    conn.commit()
    conn.close()

    print("Database initialized successfully.")
    print("Database:", DB_NAME)


# ============================================================
# ROUTE PLAN - INSERT
# ============================================================

def create_route_plan(
    plan_date,
    route_no,
    agency_no,
    sequence_no=1,
    dr_code=None,
    customer_name=None,
    address=None,
    planned_qty=0,
    vehicle_no=None,
    driver_name=None,
    planned_time=None,
    remarks=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO route_plan
        (
            plan_date,
            route_no,
            sequence_no,
            agency_no,
            dr_code,
            customer_name,
            address,
            planned_qty,
            vehicle_no,
            driver_name,
            planned_time,
            status,
            remarks
        )

        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            'Planned',
            ?
        )
    """, (
        plan_date,
        route_no,
        sequence_no,
        agency_no,
        dr_code,
        customer_name,
        address,
        planned_qty,
        vehicle_no,
        driver_name,
        planned_time,
        remarks
    ))

    route_plan_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return route_plan_id


# ============================================================
# DISPATCH - INSERT
# ============================================================

def create_dispatch(
    dispatch_date,
    route_no,
    agency_no,
    dr_code=None,
    fg_code=None,
    quantity=0,
    vehicle_no=None,
    driver_name=None,
    dispatch_time=None,
    route_plan_id=None,
    remarks=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO dispatch
        (
            dispatch_date,
            route_no,
            agency_no,
            dr_code,
            fg_code,
            quantity,
            vehicle_no,
            driver_name,
            dispatch_time,
            status,
            remarks,
            route_plan_id
        )

        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?, ?, ?,
            'Pending',
            ?, ?
        )
    """, (
        dispatch_date,
        route_no,
        agency_no,
        dr_code,
        fg_code,
        quantity,
        vehicle_no,
        driver_name,
        dispatch_time,
        remarks,
        route_plan_id
    ))

    dispatch_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return dispatch_id


# ============================================================
# UPDATE ROUTE PLAN STATUS
# ============================================================

def update_route_plan_status(
    route_plan_id,
    new_status,
    remarks=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE route_plan

        SET
            status = ?,
            remarks = COALESCE(?, remarks)

        WHERE id = ?
    """, (
        new_status,
        remarks,
        route_plan_id
    ))

    conn.commit()
    conn.close()


# ============================================================
# UPDATE DISPATCH STATUS
# ============================================================

def update_dispatch_status(
    dispatch_id,
    new_status,
    remarks=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE dispatch

        SET
            status = ?,
            remarks = COALESCE(?, remarks)

        WHERE id = ?
    """, (
        new_status,
        remarks,
        dispatch_id
    ))

    conn.commit()
    conn.close()


# ============================================================
# GET ROUTE PLAN
# ============================================================

def get_route_plan(
    plan_date=None,
    route_no=None,
    status=None
):

    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT
            rp.*,
            vm.vehicle_type,
            vm.capacity,
            dm.mobile_no AS driver_mobile

        FROM route_plan rp

        LEFT JOIN vehicle_master vm
            ON rp.vehicle_no = vm.vehicle_no

        LEFT JOIN driver_master dm
            ON rp.driver_name = dm.driver_name

        WHERE 1=1
    """

    params = []

    if plan_date:
        query += " AND rp.plan_date = ?"
        params.append(plan_date)

    if route_no:
        query += " AND rp.route_no = ?"
        params.append(route_no)

    if status:
        query += " AND rp.status = ?"
        params.append(status)

    query += """
        ORDER BY
            rp.plan_date,
            rp.route_no,
            rp.sequence_no
    """

    cursor.execute(query, params)

    rows = cursor.fetchall()

    conn.close()

    return rows


# ============================================================
# GET DISPATCH
# ============================================================

def get_dispatch(
    dispatch_date=None,
    route_no=None,
    status=None
):

    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT
            d.*,

            rp.customer_name,
            rp.address,
            rp.sequence_no,

            vm.vehicle_type,
            vm.capacity,

            dm.mobile_no AS driver_mobile

        FROM dispatch d

        LEFT JOIN route_plan rp
            ON d.route_plan_id = rp.id

        LEFT JOIN vehicle_master vm
            ON d.vehicle_no = vm.vehicle_no

        LEFT JOIN driver_master dm
            ON d.driver_name = dm.driver_name

        WHERE 1=1
    """

    params = []

    if dispatch_date:
        query += " AND d.dispatch_date = ?"
        params.append(dispatch_date)

    if route_no:
        query += " AND d.route_no = ?"
        params.append(route_no)

    if status:
        query += " AND d.status = ?"
        params.append(status)

    query += """
        ORDER BY
            d.dispatch_date,
            d.route_no,
            d.id
    """

    cursor.execute(query, params)

    rows = cursor.fetchall()

    conn.close()

    return rows


# ============================================================
# GET DISPATCH HISTORY
# ============================================================

def get_dispatch_history(dispatch_id):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *

        FROM dispatch_status_history

        WHERE dispatch_id = ?

        ORDER BY changed_at DESC
    """, (dispatch_id,))

    rows = cursor.fetchall()

    conn.close()

    return rows


# ============================================================
# GET ROUTE PLAN HISTORY
# ============================================================

def get_route_plan_history(route_plan_id):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *

        FROM route_plan_status_history

        WHERE route_plan_id = ?

        ORDER BY changed_at DESC
    """, (route_plan_id,))

    rows = cursor.fetchall()

    conn.close()

    return rows


# ============================================================
# ADD VEHICLE
# ============================================================

def add_vehicle(
    vehicle_no,
    vehicle_type=None,
    capacity=0,
    driver_name=None,
    driver_mobile=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO vehicle_master
        (
            vehicle_no,
            vehicle_type,
            capacity,
            driver_name,
            driver_mobile,
            status
        )

        VALUES
        (
            ?, ?, ?, ?, ?,
            'Available'
        )
    """, (
        vehicle_no,
        vehicle_type,
        capacity,
        driver_name,
        driver_mobile
    ))

    conn.commit()
    conn.close()


# ============================================================
# ADD DRIVER
# ============================================================

def add_driver(
    driver_name,
    mobile_no=None,
    license_no=None
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO driver_master
        (
            driver_name,
            mobile_no,
            license_no,
            status
        )

        VALUES
        (
            ?, ?, ?,
            'Active'
        )
    """, (
        driver_name,
        mobile_no,
        license_no
    ))

    conn.commit()
    conn.close()


# ============================================================
# ROUTE PLAN -> DISPATCH
# ============================================================

def create_dispatch_from_route_plan(
    route_plan_id,
    fg_code=None,
    quantity=None,
    dispatch_time=None,
    remarks=None
):

    conn = get_connection()
    cursor = conn.cursor()

    # Get route plan
    cursor.execute("""
        SELECT *

        FROM route_plan

        WHERE id = ?
    """, (route_plan_id,))

    route = cursor.fetchone()

    if not route:
        conn.close()
        raise ValueError("Route Plan not found.")


    # If quantity is not supplied,
    # use planned quantity
    if quantity is None:
        quantity = route["planned_qty"]


    # Create dispatch
    cursor.execute("""
        INSERT INTO dispatch
        (
            dispatch_date,
            route_no,
            agency_no,
            dr_code,
            fg_code,
            quantity,
            vehicle_no,
            driver_name,
            dispatch_time,
            status,
            remarks,
            route_plan_id
        )

        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?, ?, ?,
            'Pending',
            ?,
            ?
        )
    """, (
        route["plan_date"],
        route["route_no"],
        route["agency_no"],
        route["dr_code"],
        fg_code,
        quantity,
        route["vehicle_no"],
        route["driver_name"],
        dispatch_time,
        remarks,
        route_plan_id
    ))

    dispatch_id = cursor.lastrowid


    # Update route status
    cursor.execute("""
        UPDATE route_plan

        SET status = 'Assigned'

        WHERE id = ?
    """, (route_plan_id,))


    conn.commit()
    conn.close()

    return dispatch_id


# ============================================================
# DAILY DASHBOARD
# ============================================================

def get_daily_summary(date_value):

    conn = get_connection()
    cursor = conn.cursor()


    # Route count
    cursor.execute("""
        SELECT COUNT(*) AS total

        FROM route_plan

        WHERE plan_date = ?
    """, (date_value,))

    total_routes = cursor.fetchone()["total"]


    # Planned quantity
    cursor.execute("""
        SELECT COALESCE(SUM(planned_qty), 0) AS total

        FROM route_plan

        WHERE plan_date = ?
    """, (date_value,))

    planned_qty = cursor.fetchone()["total"]


    # Dispatch quantity
    cursor.execute("""
        SELECT COALESCE(SUM(quantity), 0) AS total

        FROM dispatch

        WHERE dispatch_date = ?
    """, (date_value,))

    dispatch_qty = cursor.fetchone()["total"]


    # Pending dispatch
    cursor.execute("""
        SELECT COUNT(*) AS total

        FROM dispatch

        WHERE dispatch_date = ?
        AND status = 'Pending'
    """, (date_value,))

    pending_dispatch = cursor.fetchone()["total"]


    # Completed dispatch
    cursor.execute("""
        SELECT COUNT(*) AS total

        FROM dispatch

        WHERE dispatch_date = ?
        AND status = 'Delivered'
    """, (date_value,))

    completed_dispatch = cursor.fetchone()["total"]


    conn.close()


    return {
        "date": date_value,
        "total_routes": total_routes,
        "planned_qty": planned_qty,
        "dispatch_qty": dispatch_qty,
        "pending_dispatch": pending_dispatch,
        "completed_dispatch": completed_dispatch
    }


# ============================================================
# TEST / START DATABASE
# ============================================================

if __name__ == "__main__":

    init_database()

    print()
    print("==========================================")
    print(" ROUTE & DISPATCH DATABASE READY")
    print("==========================================")