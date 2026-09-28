import os
import sqlite3
import datetime
import csv
import random
import re
from io import BytesIO, TextIOWrapper
from functools import wraps
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, session, redirect, url_for, Response
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import math

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'inventory-portal-secure-key-prod-2024')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = datetime.timedelta(days=7)
if os.environ.get('SECURE_COOKIE', '').lower() in ('true', '1') or os.environ.get('FLASK_ENV') == 'production':
    app.config['SESSION_COOKIE_SECURE'] = True

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'inventory.db')
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads', 'dispatch_docs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

PLATFORMS = [
    'Amazon', 'Blinkit', 'Zepto', 'Instamart', 'Flipkart', 
    'BigBasket', 'JioMart', 'Flipkart Minutes', 'Amazon Now', 'Dmart Ready'
]

CHANNELS = ['GT', 'MT']

DEFAULT_NAMES = {
    'admin@company.com': 'Admin Manager (Full Access)',
    'inventory@company.com': 'Inventory Lead (Master Stock & Incoming Stock)',
    'dispatch@company.com': 'Dispatch Lead (All Platform Dispatches)',
    'dispatch_online@company.com': 'Online Dispatch Lead (Online Channels)',
    'dispatch_gt@company.com': 'GT Dispatch Lead (General Trade)',
    'dispatch_mt@company.com': 'MT Dispatch Lead (Modern Trade)',
    'person1@company.com': 'Person 1: Priya (Zepto & Instamart Logistics)',
    'person2@company.com': 'Person 2: Rahul (Blinkit & BigBasket Logistics)',
    'person3@company.com': 'Person 3: Amit (Amazon & Marketplaces Logistics)',
    'b2b@company.com': 'Person 4: Suresh (GT & MT B2B Trade Logistics)',
    'joint_gt@company.com': 'Inventory & GT Dispatch Lead (Joint Role)',
    'viewer@company.com': 'Guest Viewer (Read-Only)'
}

def init_schema():
    """Ensures courier columns, doc upload columns, permissions, and sales_data table exist."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    
    # 1. Courier, logistics & document columns on dispatch_online & dispatch_gt_mt
    extra_cols = [
        ('courier_name', 'TEXT DEFAULT \'\''),
        ('tracking_id', 'TEXT DEFAULT \'\''),
        ('courier_charges', 'REAL DEFAULT 0.0'),
        ('dispatch_status', 'TEXT DEFAULT \'Dispatched\''),
        ('eway_bill_no', 'TEXT DEFAULT \'\''),
        ('shipping_notes', 'TEXT DEFAULT \'\''),
        ('po_doc_url', 'TEXT DEFAULT \'\''),
        ('po_doc_name', 'TEXT DEFAULT \'\''),
        ('invoice_doc_url', 'TEXT DEFAULT \'\''),
        ('invoice_doc_name', 'TEXT DEFAULT \'\''),
        ('po_date', 'TEXT DEFAULT \'\''),
        ('po_value', 'REAL DEFAULT 0.0'),
        ('pickup_date', 'TEXT DEFAULT \'\''),
        ('actual_weight', 'REAL DEFAULT 0.0'),
        ('charged_weight', 'REAL DEFAULT 0.0')
    ]
    for table_name in ['dispatch_online', 'dispatch_gt_mt']:
        cur.execute(f"PRAGMA table_info({table_name})")
        existing_cols = [c[1] for c in cur.fetchall()]
        for col_name, col_def in extra_cols:
            if col_name not in existing_cols:
                cur.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_def}")
                
    # 2. General & Channel Dispatch Roles (permissions table)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS permissions (
        role_key TEXT PRIMARY KEY,
        authorized_email TEXT NOT NULL,
        description TEXT
    )
    ''')
    
    initial_roles = [
        ('ADMIN', 'admin@company.com', 'Administrator (Full Access across all sheets and permissions)'),
        ('INVENTORY_MANAGER', 'inventory@company.com', 'Inventory Lead (Exclusively updates Master Sheet baseline stock & incoming inventory)'),
        ('DISPATCH_MANAGER', 'dispatch@company.com', 'Dispatch Lead (All Platforms Dispatcher)'),
        ('DISPATCH_ONLINE_EMAILS', 'dispatch_online@company.com, dispatch@company.com', 'Authorized email(s) for Online Dispatch (comma-separated)'),
        ('DISPATCH_GT_EMAILS', 'dispatch_gt@company.com, b2b@company.com, dispatch@company.com', 'Authorized email(s) for GT Dispatch (comma-separated)'),
        ('DISPATCH_MT_EMAILS', 'dispatch_mt@company.com, b2b@company.com, dispatch@company.com', 'Authorized email(s) for MT Dispatch (comma-separated)')
    ]
    for r_key, r_email, r_desc in initial_roles:
        cur.execute('''
        INSERT OR IGNORE INTO permissions (role_key, authorized_email, description)
        VALUES (?, ?, ?)
        ''', (r_key, r_email, r_desc))
        
    # 3. Platform & Channel specific assignments table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS platform_assignments (
        platform_key TEXT PRIMARY KEY,
        channel_type TEXT,
        assigned_email TEXT NOT NULL,
        assigned_name TEXT NOT NULL
    )
    ''')
    
    initial_platform_assignments = [
        ('Zepto', 'online', 'person1@company.com', 'Person 1: Priya (Zepto Lead)'),
        ('Instamart', 'online', 'person1@company.com', 'Person 1: Priya (Instamart Lead)'),
        ('Flipkart Minutes', 'online', 'person1@company.com', 'Person 1: Priya (FK Minutes Lead)'),
        ('Blinkit', 'online', 'person2@company.com', 'Person 2: Rahul (Blinkit Lead)'),
        ('BigBasket', 'online', 'person2@company.com', 'Person 2: Rahul (BigBasket Lead)'),
        ('Dmart Ready', 'online', 'person2@company.com', 'Person 2: Rahul (Dmart Ready Lead)'),
        ('Amazon', 'online', 'person3@company.com', 'Person 3: Amit (Amazon Lead)'),
        ('Flipkart', 'online', 'person3@company.com', 'Person 3: Amit (Flipkart Lead)'),
        ('JioMart', 'online', 'person3@company.com', 'Person 3: Amit (JioMart Lead)'),
        ('Amazon Now', 'online', 'person3@company.com', 'Person 3: Amit (Amazon Now Lead)'),
        ('GT', 'gtmt', 'b2b@company.com', 'Person 4: Suresh (GT Trade Lead)'),
        ('MT', 'gtmt', 'b2b@company.com', 'Person 4: Suresh (MT Trade Lead)')
    ]
    for p_key, c_type, a_email, a_name in initial_platform_assignments:
        cur.execute('''
        INSERT OR IGNORE INTO platform_assignments (platform_key, channel_type, assigned_email, assigned_name)
        VALUES (?, ?, ?, ?)
        ''', (p_key, c_type, a_email, a_name))
        
    # 4. Sales Data table (Consumer Offtake from Platforms)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS sales_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        platform TEXT NOT NULL,
        sku TEXT NOT NULL,
        product_name TEXT,
        brand TEXT,
        units_sold INTEGER NOT NULL DEFAULT 0,
        revenue REAL DEFAULT 0.0,
        store_location TEXT DEFAULT '',
        source_file TEXT DEFAULT '',
        uploaded_by TEXT DEFAULT '',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 5. Date-wise & Batch-wise Inventory Inwarding Table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS stock_inward_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inward_date TEXT NOT NULL,
        batch_no TEXT NOT NULL,
        sku TEXT NOT NULL,
        product_name TEXT NOT NULL,
        brand TEXT NOT NULL,
        quantity_added INTEGER NOT NULL DEFAULT 0,
        mfg_date TEXT DEFAULT '',
        expiry_date TEXT DEFAULT '',
        supplier_po_ref TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        added_by TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (sku) REFERENCES products(sku)
    )
    ''')

    # 6. Registered Brands Table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS brands (
        name TEXT PRIMARY KEY,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    cur.execute('INSERT OR IGNORE INTO brands (name) SELECT DISTINCT brand FROM products WHERE brand IS NOT NULL AND TRIM(brand) != ""')
        
    # Backfill missing courier costing data for realistic view
    for tbl in ['dispatch_online', 'dispatch_gt_mt']:
        cur.execute(f"SELECT id, dispatch_date, appointment_date, COALESCE(actual_sent, po_quantity, 100) as qty, po_value, courier_charges, po_date, pickup_date, actual_weight, charged_weight FROM {tbl}")
        rows = cur.fetchall()
        for r in rows:
            updates = []
            params = []
            qty = r['qty'] or 50
            if not r['po_date']:
                p_date = r['dispatch_date'] or datetime.date.today().strftime('%Y-%m-%d')
                updates.append("po_date = ?")
                params.append(p_date)
            if not r['pickup_date']:
                pk_date = r['dispatch_date'] or datetime.date.today().strftime('%Y-%m-%d')
                updates.append("pickup_date = ?")
                params.append(pk_date)
            if not r['po_value'] or r['po_value'] == 0.0:
                p_val = round(qty * random.uniform(320.0, 750.0), 2)
                updates.append("po_value = ?")
                params.append(p_val)
            else:
                p_val = r['po_value']
            if not r['courier_charges'] or r['courier_charges'] == 0.0:
                c_chg = round(p_val * random.uniform(0.045, 0.095), 2)
                updates.append("courier_charges = ?")
                params.append(c_chg)
            if not r['actual_weight'] or r['actual_weight'] == 0.0:
                act_wt = round(qty * 0.28, 2)
                chg_wt = round(act_wt * random.uniform(1.05, 1.22), 2)
                updates.append("actual_weight = ?")
                params.append(act_wt)
                updates.append("charged_weight = ?")
                params.append(chg_wt)
            if updates:
                params.append(r['id'])
                cur.execute(f"UPDATE {tbl} SET {', '.join(updates)} WHERE id = ?", params)

    # Seed sample batches if empty
    cur.execute("SELECT COUNT(*) FROM stock_inward_batches")
    if cur.fetchone()[0] == 0:
        sample_batches = [
            ('2024-05-02', 'BAT-2024-0502', 'B1-001', 'Body Wash 250ml', 'Brand A', 500, '2024-04-20', '2026-04-19', 'PO-SUP-8910', 'Primary factory delivery - Plant 1', 'admin@company.com'),
            ('2024-05-05', 'BAT-2024-0505', 'B1-002', 'Face Cleanser 100ml', 'Brand A', 450, '2024-04-25', '2026-04-24', 'PO-SUP-8912', 'Raw material batch cleared QC', 'inventory@company.com'),
            ('2024-05-08', 'BAT-2024-0508', 'B1-003', 'Hydrating Moisturizer 50g', 'Brand A', 600, '2024-04-28', '2026-04-27', 'PO-SUP-8915', 'Central warehouse restocking', 'inventory@company.com'),
            ('2024-05-12', 'BAT-2024-0512', 'B2-001', 'Vitamin C Serum 30ml', 'Brand B', 400, '2024-05-01', '2025-11-01', 'PO-SUP-8920', 'Imported active ingredients batch', 'admin@company.com'),
            ('2024-05-15', 'BAT-2024-0515', 'B2-002', 'Retinol Night Cream 50g', 'Brand B', 350, '2024-05-04', '2025-11-04', 'PO-SUP-8924', 'Factory replenishment lote B', 'inventory@company.com')
        ]
        cur.executemany('''
            INSERT INTO stock_inward_batches (inward_date, batch_no, sku, product_name, brand, quantity_added, mfg_date, expiry_date, supplier_po_ref, notes, added_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', sample_batches)
    
    # 7. Secure User Authentication Table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        is_active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        default_pwd_hash = generate_password_hash('Admin@123')
        for u_email, u_name in DEFAULT_NAMES.items():
            cur.execute('''
            INSERT OR IGNORE INTO users (email, password_hash, full_name, is_active)
            VALUES (?, ?, ?, 1)
            ''', (u_email, default_pwd_hash, u_name))

    # 8. Purchase Orders (Multi-Product Header Table)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS purchase_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_number TEXT UNIQUE NOT NULL,
        po_date TEXT NOT NULL,
        platform_or_channel TEXT NOT NULL,
        channel_type TEXT NOT NULL,
        location TEXT DEFAULT '',
        distributor_name TEXT DEFAULT '',
        appointment_date TEXT DEFAULT '',
        total_items INTEGER DEFAULT 0,
        total_quantity INTEGER DEFAULT 0,
        total_po_value REAL DEFAULT 0.0,
        status TEXT DEFAULT 'Pending Fulfillment',
        po_doc_url TEXT DEFAULT '',
        po_doc_name TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        created_by TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Migration for GRN tracking fields on purchase_orders
    cur.execute("PRAGMA table_info(purchase_orders)")
    po_cols = [col[1] for col in cur.fetchall()]
    if 'grn_number' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN grn_number TEXT DEFAULT ''")
    if 'grn_date' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN grn_date TEXT DEFAULT ''")
    if 'grn_notes' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN grn_notes TEXT DEFAULT ''")
    if 'grn_done_by' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN grn_done_by TEXT DEFAULT ''")
    if 'grn_done_at' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN grn_done_at TEXT DEFAULT ''")
    if 'is_packed' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN is_packed INTEGER DEFAULT 0")
    if 'is_pickup_ready' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN is_pickup_ready INTEGER DEFAULT 0")
    if 'is_dispatched' not in po_cols:
        cur.execute("ALTER TABLE purchase_orders ADD COLUMN is_dispatched INTEGER DEFAULT 0")

    cur.execute("UPDATE purchase_orders SET is_packed = 1, is_pickup_ready = 1, is_dispatched = 1 WHERE status IN ('Partially Dispatched', 'Fully Dispatched', 'GRN Done') AND (is_packed = 0 OR is_packed IS NULL)")

    # 9. Purchase Order Line Items (Multi-Products across any Brand)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS po_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_id INTEGER NOT NULL,
        po_number TEXT NOT NULL,
        sku TEXT NOT NULL,
        product_name TEXT NOT NULL,
        brand TEXT NOT NULL,
        category TEXT DEFAULT '',
        ordered_quantity INTEGER NOT NULL DEFAULT 1,
        unit_price REAL DEFAULT 0.0,
        total_value REAL DEFAULT 0.0,
        dispatched_quantity INTEGER DEFAULT 0,
        notes TEXT DEFAULT '',
        FOREIGN KEY (po_id) REFERENCES purchase_orders(id) ON DELETE CASCADE
    )
    ''')

    # Seed sample multi-product POs if empty
    cur.execute("SELECT COUNT(*) FROM purchase_orders")
    if cur.fetchone()[0] == 0:
        sample_pos = [
            ('PO-BLK-2024-8841', '2024-05-18', 'Blinkit', 'online', 'Blinkit Gurugram FC-3', '', '2024-05-21', 3, 430, 71500.0, 'Pending Fulfillment', '', '', 'Priority restock order across brands', 'admin@company.com'),
            ('PO-ZEP-2024-1092', '2024-05-19', 'Zepto', 'online', 'Zepto Whitefield DC-2, Bengaluru', '', '2024-05-22', 2, 280, 48200.0, 'Pending Fulfillment', '', '', 'Quick Commerce replenishment batch', 'person1@company.com'),
            ('PO-GT-MUM-4401', '2024-05-16', 'GT', 'gtmt', 'Western Distributors, Mumbai', 'Western Distributors', '2024-05-20', 3, 350, 68000.0, 'Partially Dispatched', '', '', 'B2B distributor monthly consignment', 'joint_gt@company.com')
        ]
        for spo in sample_pos:
            cur.execute('''
            INSERT INTO purchase_orders (po_number, po_date, platform_or_channel, channel_type, location, distributor_name, appointment_date, total_items, total_quantity, total_po_value, status, po_doc_url, po_doc_name, notes, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', spo)
            po_id = cur.lastrowid
            
            if spo[0] == 'PO-BLK-2024-8841':
                items = [
                    (po_id, spo[0], 'CS0001', 'Cleanser', 'California Skin+', 'Skincare', 150, 220.0, 33000.0, 0, ''),
                    (po_id, spo[0], 'NL001', 'NutraChips Rock Salt Banana Chips', 'NutraChips', 'Chips', 200, 65.0, 13000.0, 0, ''),
                    (po_id, spo[0], 'B1-001', 'Body Wash 250ml', 'Brand A', 'General', 80, 318.75, 25500.0, 0, '')
                ]
            elif spo[0] == 'PO-ZEP-2024-1092':
                items = [
                    (po_id, spo[0], 'CS0002', 'Serum', 'California Skin+', 'Skincare', 100, 350.0, 35000.0, 0, ''),
                    (po_id, spo[0], 'NL004', 'NutraCookies Sugar Free Oats Cookies', 'NutraCookies', 'Cookies', 180, 73.33, 13200.0, 0, '')
                ]
            else:
                items = [
                    (po_id, spo[0], 'CS0004', 'Moisturizer', 'California Skin+', 'Skincare', 100, 280.0, 28000.0, 50, ''),
                    (po_id, spo[0], 'NL003', 'NutraBites Baked Bhujia', 'NutraBites', 'Snacks', 150, 80.0, 12000.0, 150, ''),
                    (po_id, spo[0], 'B1-001', 'Body Wash 250ml', 'Brand A', 'General', 100, 280.0, 28000.0, 0, '')
                ]
            cur.executemany('''
            INSERT INTO po_items (po_id, po_number, sku, product_name, brand, category, ordered_quantity, unit_price, total_value, dispatched_quantity, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', items)

    # 10. Automatically sync all PO line items to relevant dispatch tables
    sync_all_pos_to_dispatches(cur)

    # 11. Blinkit Dedicated AI Replenishment & Storage Cost Optimization Tables
    cur.execute('''
    CREATE TABLE IF NOT EXISTS blinkit_inventory_current (
        item_id INTEGER,
        facility_id INTEGER,
        item_name TEXT,
        brand_name TEXT,
        upc TEXT,
        uom TEXT,
        facility_name TEXT,
        net_scheduled INTEGER DEFAULT 0,
        incoming_scheduled INTEGER DEFAULT 0,
        recalled_inventory INTEGER DEFAULT 0,
        total_sellable INTEGER DEFAULT 0,
        warehouse_stock INTEGER DEFAULT 0,
        in_between_stock INTEGER DEFAULT 0,
        darkstore_stock INTEGER DEFAULT 0,
        total_unsellable INTEGER DEFAULT 0,
        damaged INTEGER DEFAULT 0,
        lost INTEGER DEFAULT 0,
        expired INTEGER DEFAULT 0,
        near_expiry INTEGER DEFAULT 0,
        sales_7d INTEGER DEFAULT 0,
        sales_15d INTEGER DEFAULT 0,
        sales_30d INTEGER DEFAULT 0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (item_id, facility_id)
    )
    ''')

    cur.execute('''
    CREATE TABLE IF NOT EXISTS blinkit_sales_orders (
        order_id TEXT PRIMARY KEY,
        order_date TEXT,
        item_id INTEGER,
        product_name TEXT,
        brand_name TEXT,
        upc TEXT,
        supply_city TEXT,
        supply_state TEXT,
        customer_city TEXT,
        customer_state TEXT,
        order_status TEXT,
        quantity INTEGER DEFAULT 1,
        mrp REAL DEFAULT 0.0,
        selling_price REAL DEFAULT 0.0,
        total_gross_amount REAL DEFAULT 0.0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    cur.execute('''
    CREATE TABLE IF NOT EXISTS blinkit_facility_settings (
        facility_name TEXT PRIMARY KEY,
        facility_id INTEGER DEFAULT 0,
        lead_time_days INTEGER DEFAULT 8,
        safety_stock_days INTEGER DEFAULT 3,
        target_max_days INTEGER DEFAULT 21,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    default_blinkit_facilities = [
        ('Mumbai M12 - Feeder Warehouse', 5406, 8, 3, 21),
        ('Mumbai M10 - Feeder', 2123, 8, 3, 21),
        ('Bengaluru B5 - Feeder', 5397, 8, 3, 21),
        ('Bengaluru B3', 1873, 8, 3, 21),
        ('Pune P3 - Feeder Warehouse', 4572, 8, 3, 21),
        ('Kundli Feeder', 2010, 8, 3, 21),
        ('Faridabad - Feeder', 5096, 8, 3, 21),
        ('Ahmedabad A2 - Feeder', 2470, 8, 3, 21),
        ('Hyderabad H3 - Feeder', 3201, 8, 3, 21),
        ('Chennai C5 - Feeder', 3262, 8, 3, 21),
        ('Kolkata K6 - Feeder Warehouse', 4842, 9, 3, 21),
        ('Nagpur N1 - Feeder', 2468, 8, 3, 21)
    ]
    for fac in default_blinkit_facilities:
        cur.execute('''
        INSERT OR IGNORE INTO blinkit_facility_settings (facility_name, facility_id, lead_time_days, safety_stock_days, target_max_days)
        VALUES (?, ?, ?, ?, ?)
        ''', fac)

    # 12. Zepto Dedicated Sales Offtake Table (Anti-Overlap Idempotent Upsert)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS zepto_sales_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        sku_number TEXT NOT NULL,
        sku_name TEXT NOT NULL,
        ean TEXT,
        sku_category TEXT,
        sku_sub_category TEXT,
        brand_name TEXT,
        manufacturer_name TEXT,
        manufacturer_id TEXT,
        city TEXT NOT NULL,
        units_sold INTEGER DEFAULT 0,
        mrp REAL DEFAULT 0.0,
        gmv REAL DEFAULT 0.0,
        matched_sku TEXT,
        matched_product_name TEXT,
        source_file TEXT,
        uploaded_by TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(sale_date, sku_number, city)
    )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_zepto_date ON zepto_sales_orders(sale_date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_zepto_city ON zepto_sales_orders(city)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_zepto_sku ON zepto_sales_orders(sku_number)')

    # 13. Swiggy Instamart Dedicated Product Sales Offtake Table (Anti-Overlap Idempotent Upsert)
    cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='instamart_sales_orders'")
    tbl_row = cur.fetchone()
    if tbl_row and 'campaign_id' in tbl_row[0]:
        cur.execute("DROP TABLE instamart_sales_orders")

    cur.execute('''
    CREATE TABLE IF NOT EXISTS instamart_sales_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        item_code TEXT NOT NULL,
        product_name TEXT NOT NULL,
        variant TEXT DEFAULT '',
        brand_name TEXT DEFAULT '',
        city TEXT NOT NULL,
        area_name TEXT DEFAULT '',
        store_id TEXT DEFAULT '',
        units_sold INTEGER DEFAULT 0,
        mrp REAL DEFAULT 0.0,
        gmv REAL DEFAULT 0.0,
        matched_sku TEXT,
        matched_product_name TEXT,
        source_file TEXT,
        uploaded_by TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(sale_date, item_code, city, store_id)
    )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_instamart_date ON instamart_sales_orders(sale_date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_instamart_item ON instamart_sales_orders(item_code)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_instamart_city ON instamart_sales_orders(city)')

    # 14. BigBasket Dedicated Sales Offtake Table (Anti-Overlap Idempotent Upsert)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS bigbasket_sales_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        source_sku_id TEXT NOT NULL,
        sku_description TEXT NOT NULL,
        sku_weight TEXT DEFAULT '',
        brand_slug TEXT DEFAULT '',
        city TEXT NOT NULL,
        business_type TEXT DEFAULT 'b2c',
        top_slug TEXT DEFAULT '',
        mid_slug TEXT DEFAULT '',
        leaf_slug TEXT DEFAULT '',
        units_sold INTEGER DEFAULT 0,
        mrp REAL DEFAULT 0.0,
        sales_amount REAL DEFAULT 0.0,
        matched_sku TEXT,
        matched_product_name TEXT,
        source_file TEXT,
        uploaded_by TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(sale_date, source_sku_id, city, business_type)
    )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_bb_date ON bigbasket_sales_orders(sale_date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_bb_city ON bigbasket_sales_orders(city)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_bb_sku ON bigbasket_sales_orders(source_sku_id)')

    conn.commit()
    conn.close()

# ==============================================================================
# CANONICAL PRODUCT RESOLUTION & HARMONIZATION ENGINE
# ==============================================================================

def normalize_date_str(d_str):
    """Normalizes any date string (YYYYMMDD, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD, or ranges) into standard ISO YYYY-MM-DD."""
    if not d_str:
        return ''
    d_str = str(d_str).strip().strip('"').strip("'")
    if ' - ' in d_str:
        d_str = d_str.split(' - ')[0].strip()
    elif ' to ' in d_str.lower():
        d_str = re.split(r'\s+to\s+', d_str, flags=re.IGNORECASE)[0].strip()
        
    compact_match = re.match(r'^(\d{4})(\d{2})(\d{2})$', d_str)
    if compact_match:
        return f"{compact_match.group(1)}-{compact_match.group(2)}-{compact_match.group(3)}"

    for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d', '%d-%b-%Y', '%d %b %Y', '%Y%m%d'):
        try:
            return datetime.datetime.strptime(d_str, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return d_str

ZEPTO_EAN_MAP = {
    '8908028836101': ('CS0007', 'California Skin+ Triple Action Acne Scar Clear', 'California Skin+'),
    '8908028836040': ('CS0005', 'California Skin+ One Hour Acne Spot Relief', 'California Skin+'),
    '8908028836194': ('NL004', 'NutraCookies Sugar Free Oats Cookies', 'NutraCookies'),
    '8908028836156': ('NL003', 'NutraBites Baked Bhujia', 'NutraBites'),
    '8908028836125': ('CS0006', 'California Skin+ No-Cast, Hyaluronic Glow Sunscreen SPF50++++', 'California Skin+'),
    '8908028836095': ('CS0003', 'California Skin+ Triple Action Acne Relief Pimple Patches', 'California Skin+'),
    '8908028836026': ('CS0002', 'California Skin+ Acne Control Serum', 'California Skin+'),
    '8908028836002': ('CS0001', 'California Skin+ Acne Control Face Wash Cleanser', 'California Skin+'),
}

PRODUCT_SYNONYM_RULES = [
    # (SKU, Canonical Name, Brand, Category, [Synonym Patterns / Keywords])
    ('CS0004', 'Moisturizer', 'California Skin+', 'Skincare', [
        'moisturiz', 'moisturis', 'moistuir', 'barrier repair', 'ceramide', 'repair cream', 'hydrating gel', 'moisturising'
    ]),
    ('CS0006', 'Sunscreen', 'California Skin+', 'Skincare', [
        'sunscreen', 'sun screen', 'sunblock', 'sun block', 'spf50', 'spf 50', 'no-cast', 'cica sunscreen', 'glow sunscreen'
    ]),
    ('CS0001', 'Cleanser', 'California Skin+', 'Skincare', [
        'cleanser', 'face wash', 'facewash', 'cleansing gel', 'acne wash', 'cica cleanser', 'face cleanser'
    ]),
    ('CS0003', 'Patches', 'California Skin+', 'Skincare', [
        'patch', 'patches', 'pimple patch', 'pimple patches', 'acne patch', 'acne patches', 'hydrocolloid', 'spot patch'
    ]),
    ('CS0005', 'Spot Relief', 'California Skin+', 'Skincare', [
        'spot relief', 'spot treatment', 'acne spot', 'one hour acne', '1 hour acne', 'spot corrector'
    ]),
    ('CS0007', 'Scar Cream', 'California Skin+', 'Skincare', [
        'scar cream', 'scar clear', 'acne scar', 'scar gel', 'scar repair', 'scar'
    ]),
    ('CS0002', 'Serum', 'California Skin+', 'Skincare', [
        'serum', 'acne control face serum', 'face serum', 'acne serum', 'niacinamide serum'
    ]),
    ('NL001', 'NutraChips Rock Salt Banana Chips', 'NutraChips', 'Chips', [
        'banana chip', 'banana chips', 'rock salt banana', 'banana'
    ]),
    ('NL002', 'NutraChips Baked Ragi Chips', 'NutraChips', 'Chips', [
        'ragi chip', 'ragi chips', 'baked ragi', 'ragi'
    ]),
    ('NL003', 'NutraBites Baked Bhujia', 'NutraBites', 'Bites', [
        'bhujia', 'baked bhujia', 'nutrabites', 'protein bhujia'
    ]),
    ('NL008', 'NutraCookies Sugar Free Chocolate Cookies', 'NutraCookies', 'Cookies', [
        'chocolate cookie', 'chocolate cookies', 'choco cookie', 'choco cookies', 'chocolate'
    ]),
    ('NL009', 'NutraCookies Sugar Free Protein Cookies', 'NutraCookies', 'Cookies', [
        'protein cookie', 'protein cookies'
    ]),
    ('NL004', 'NutraCookies Sugar Free Oats Cookies', 'NutraCookies', 'Cookies', [
        'oats cookie', 'oats cookies', 'oat cookie', 'oat cookies', 'oats', 'oat', 'cookie', 'cookies'
    ]),
    ('NL005', 'Hair & Skin Supplements Gummies', 'NutraGummies', 'Gummies', [
        'hair skin', 'hair & skin', 'hair and skin', 'hair gummy', 'hair gummies', 'hair'
    ]),
    ('NL006', 'Sleep Supplements Gummies', 'NutraGummies', 'Gummies', [
        'sleep gummy', 'sleep gummies', 'deep sleep', 'melatonin', 'sleep supplement', 'sleep'
    ]),
    ('NL007', 'Handwash', 'Derma+', 'Personal Care', [
        'handwash', 'hand wash', 'hand soap', 'aqua fresh hand wash', 'aquafresh'
    ]),
    ('NUNOO2024', 'NUTRANOODLES ATTA NOODLES', 'NutraNoodles', 'FMCG', [
        'noodle', 'noodles', 'nutranoodles', 'atta noodle'
    ]),
    ('B1-001', 'Body Wash 250ml', 'Brand A', 'General', [
        'body wash', 'bodywash'
    ]),
    ('SKU-001', 'Herbal Face Wash', 'Brand A', 'General', [
        'herbal face wash', 'herbal wash'
    ])
]

def resolve_canonical_product(sku_raw, name_raw, ean=None, cur=None):
    """
    Comprehensive Canonical Product Resolver.
    Intelligently maps platform variations, supplier PO lines, and diverse spelling
    (e.g., 'Moistuirseer', 'Barrier Repair Moisturizer', 'CICA Sunscreen SPF50')
    to the single canonical master catalog product and SKU.
    """
    # 1. EAN barcode lookup
    if ean:
        clean_ean = str(ean).strip()
        if clean_ean in ZEPTO_EAN_MAP:
            s, n, b = ZEPTO_EAN_MAP[clean_ean]
            return s, n, b, 'Skincare'

    # 2. Direct SKU lookup
    clean_sku = str(sku_raw or '').strip().upper()
    if cur and clean_sku:
        try:
            prod = cur.execute('SELECT sku, name, brand, category FROM products WHERE UPPER(sku) = ?', (clean_sku,)).fetchone()
            if prod:
                return prod['sku'], prod['name'], prod['brand'], prod['category']
        except Exception:
            pass

    # 3. Direct product name lookup against database
    clean_name = str(name_raw or '').strip()
    if cur and clean_name:
        try:
            prod = cur.execute('SELECT sku, name, brand, category FROM products WHERE LOWER(name) = LOWER(?)', (clean_name,)).fetchone()
            if prod:
                return prod['sku'], prod['name'], prod['brand'], prod['category']
        except Exception:
            pass

    # 4. Synonym & Keyword Rules Engine (handles typos, varied phrasing across platforms)
    raw_text = (str(name_raw or '') + ' ' + str(sku_raw or '')).lower()
    norm = re.sub(r'[^a-z0-9\s]', ' ', raw_text)
    norm = re.sub(r'\s+', ' ', norm).strip()

    for c_sku, c_name, c_brand, c_cat, patterns in PRODUCT_SYNONYM_RULES:
        for pat in patterns:
            if pat in norm:
                return c_sku, c_name, c_brand, c_cat

    # 5. Fallback substring matching on DB product names
    if cur and norm:
        try:
            db_prods = cur.execute('SELECT sku, name, brand, category FROM products').fetchall()
            for p in db_prods:
                p_norm = re.sub(r'[^a-z0-9\s]', ' ', p['name'].lower()).strip()
                if p_norm and (p_norm in norm or norm in p_norm):
                    return p['sku'], p['name'], p['brand'], p['category']
        except Exception:
            pass

    return clean_sku or 'UNKNOWN', name_raw or 'Unmapped Product', 'Other', 'General'

def map_product_to_internal(sku_raw, name_raw, ean=None, cur=None):
    """Backwards-compatible wrapper returning (sku, name, brand)."""
    s, n, b, _ = resolve_canonical_product(sku_raw, name_raw, ean, cur)
    return s, n, b

def sync_all_pos_to_dispatches(cur):
    """Ensures every PO line item has a corresponding linked entry in dispatch_online or dispatch_gt_mt."""
    cur.execute('SELECT * FROM purchase_orders')
    pos = cur.fetchall()
    for po in pos:
        po_id = po['id']
        cur.execute('SELECT * FROM po_items WHERE po_id = ?', (po_id,))
        items = cur.fetchall()
        for it in items:
            c_sku, c_name, c_brand, c_cat = resolve_canonical_product(it['sku'], it['product_name'], None, cur)
            sku = c_sku
            po_num = po['po_number']
            is_gtmt = (po['platform_or_channel'] in ['GT', 'MT']) or (po['channel_type'] == 'gtmt')
            if is_gtmt:
                existing = cur.execute('SELECT id FROM dispatch_gt_mt WHERE po_number = ? AND sku = ?', (po_num, sku)).fetchone()
                if not existing:
                    channel = 'GT' if 'GT' in po['platform_or_channel'].upper() else 'MT'
                    buyer = po['distributor_name'] or po['location'] or 'Distributor'
                    disp_qty = it['dispatched_quantity'] or 0
                    ord_qty = it['ordered_quantity'] or 0
                    st = 'Dispatched' if (disp_qty >= ord_qty and ord_qty > 0) else ('Partially Dispatched' if disp_qty > 0 else 'Pending Booking')
                    cur.execute('''
                    INSERT INTO dispatch_gt_mt (
                        dispatch_date, brand, product_name, sku, category, channel, buyer_distributor, location,
                        po_number, invoice_no, po_quantity, actual_sent, appointment_date,
                        courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
                        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name, po_date, po_value
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, '', '', 0.0, ?, '', '', ?, ?, '', '', ?, ?)
                    ''', (
                        po['po_date'] or datetime.date.today().strftime('%Y-%m-%d'),
                        c_brand, c_name, sku, c_cat, channel, buyer, po['location'],
                        po_num, ord_qty, disp_qty, po['appointment_date'],
                        st, po['po_doc_url'], po['po_doc_name'], po['po_date'], it['total_value']
                    ))
            else:
                existing = cur.execute('SELECT id FROM dispatch_online WHERE po_number = ? AND sku = ?', (po_num, sku)).fetchone()
                if not existing:
                    platform = po['platform_or_channel']
                    disp_qty = it['dispatched_quantity'] or 0
                    ord_qty = it['ordered_quantity'] or 0
                    st = 'Dispatched' if (disp_qty >= ord_qty and ord_qty > 0) else ('Partially Dispatched' if disp_qty > 0 else 'Pending Booking')
                    cur.execute('''
                    INSERT INTO dispatch_online (
                        dispatch_date, brand, product_name, sku, category, platform, location,
                        po_number, po_quantity, actual_sent, appointment_date,
                        courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
                        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name, po_date, po_value
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '', 0.0, ?, '', '', ?, ?, '', '', ?, ?)
                    ''', (
                        po['po_date'] or datetime.date.today().strftime('%Y-%m-%d'),
                        c_brand, c_name, sku, c_cat, platform, po['location'],
                        po_num, ord_qty, disp_qty, po['appointment_date'],
                        st, po['po_doc_url'], po['po_doc_name'], po['po_date'], it['total_value']
                    ))

def resync_all_sales_and_pos_to_canonical(cur):
    """
    Harmonizes existing historical rows across po_items, dispatches, and sales_data
    so that varied supplier names and platform product keys (e.g., 'Moistuirseer',
    'California Skin+ Moisturizing Gel (with Ceramides)', 'Derma+ Aqua Fresh Hand Wash',
    '33049910', 'BLK-...') resolve cleanly to the master canonical product catalog.
    """
    try:
        # 1. Update po_items
        cur.execute("SELECT id, sku, product_name, brand, category FROM po_items")
        for it in cur.fetchall():
            c_sku, c_name, c_brand, c_cat = resolve_canonical_product(it['sku'], it['product_name'], None, cur)
            if (c_sku != 'UNKNOWN' and c_sku != it['sku']) or (c_name != 'Unmapped Product' and c_name != it['product_name']):
                cur.execute("""
                    UPDATE po_items
                    SET sku = ?, product_name = ?, brand = ?, category = ?
                    WHERE id = ?
                """, (c_sku, c_name, c_brand, c_cat, it['id']))

        # 2. Update dispatch_online
        cur.execute("SELECT id, sku, product_name, brand, category FROM dispatch_online")
        for it in cur.fetchall():
            c_sku, c_name, c_brand, c_cat = resolve_canonical_product(it['sku'], it['product_name'], None, cur)
            if (c_sku != 'UNKNOWN' and c_sku != it['sku']) or (c_name != 'Unmapped Product' and c_name != it['product_name']):
                cur.execute("""
                    UPDATE dispatch_online
                    SET sku = ?, product_name = ?, brand = ?, category = ?
                    WHERE id = ?
                """, (c_sku, c_name, c_brand, c_cat, it['id']))

        # 3. Update dispatch_gt_mt
        cur.execute("SELECT id, sku, product_name, brand, category FROM dispatch_gt_mt")
        for it in cur.fetchall():
            c_sku, c_name, c_brand, c_cat = resolve_canonical_product(it['sku'], it['product_name'], None, cur)
            if (c_sku != 'UNKNOWN' and c_sku != it['sku']) or (c_name != 'Unmapped Product' and c_name != it['product_name']):
                cur.execute("""
                    UPDATE dispatch_gt_mt
                    SET sku = ?, product_name = ?, brand = ?, category = ?
                    WHERE id = ?
                """, (c_sku, c_name, c_brand, c_cat, it['id']))

        # 4. Update sales_data (fix raw HSN/vendor codes like 33049910, 34013019, 21069099, BLK-..., etc.)
        cur.execute("SELECT DISTINCT sku, product_name FROM sales_data")
        distinct_sales_items = cur.fetchall()
        for row in distinct_sales_items:
            s_sku = row['sku']
            s_name = row['product_name']
            c_sku, c_name, c_brand, _ = resolve_canonical_product(s_sku, s_name, None, cur)
            if (c_sku != 'UNKNOWN' and c_sku != s_sku) or (c_name != 'Unmapped Product' and s_name and c_name != s_name):
                cur.execute("""
                    UPDATE sales_data
                    SET sku = ?, product_name = ?, brand = ?
                    WHERE sku = ? AND product_name = ?
                """, (c_sku, c_name, c_brand, s_sku, s_name))

        # 5. Update zepto_sales_orders & instamart_sales_orders matched columns
        cur.execute("SELECT DISTINCT sku_number, sku_name, ean, matched_sku FROM zepto_sales_orders")
        for row in cur.fetchall():
            c_sku, c_name, c_brand, _ = resolve_canonical_product(row['sku_number'], row['sku_name'], row['ean'], cur)
            if c_sku != 'UNKNOWN' and c_sku != row['matched_sku']:
                cur.execute("""
                    UPDATE zepto_sales_orders
                    SET matched_sku = ?, matched_product_name = ?
                    WHERE sku_number = ? AND (ean = ? OR (ean IS NULL AND ? IS NULL))
                """, (c_sku, c_name, row['sku_number'], row['ean'], row['ean']))

        cur.execute("SELECT DISTINCT item_code, product_name, matched_sku FROM instamart_sales_orders")
        for row in cur.fetchall():
            c_sku, c_name, c_brand, _ = resolve_canonical_product(row['item_code'], row['product_name'], None, cur)
            if c_sku != 'UNKNOWN' and c_sku != row['matched_sku']:
                cur.execute("""
                    UPDATE instamart_sales_orders
                    SET matched_sku = ?, matched_product_name = ?
                    WHERE item_code = ? AND product_name = ?
                """, (c_sku, c_name, row['item_code'], row['product_name']))

        # 6. Update bigbasket_sales_orders matched columns
        cur.execute("SELECT DISTINCT source_sku_id, sku_description, matched_sku FROM bigbasket_sales_orders")
        for row in cur.fetchall():
            c_sku, c_name, c_brand, _ = resolve_canonical_product(row['source_sku_id'], row['sku_description'], None, cur)
            if c_sku != 'UNKNOWN' and c_sku != row['matched_sku']:
                cur.execute("""
                    UPDATE bigbasket_sales_orders
                    SET matched_sku = ?, matched_product_name = ?
                    WHERE source_sku_id = ? AND sku_description = ?
                """, (c_sku, c_name, row['source_sku_id'], row['sku_description']))
    except Exception as e:
        print("Notice during canonical resync:", e)

init_schema()

# Harmonize existing database entries on launch
try:
    _conn = get_db()
    _cur = _conn.cursor()
    resync_all_sales_and_pos_to_canonical(_cur)
    _conn.commit()
    _conn.close()
except Exception as _e:
    print("Initial canonical resync notice:", _e)

def get_user_info():
    # Real session authentication: Read exclusively from encrypted server-side session cookie
    email = session.get('user_email', '').strip().lower()
    name = session.get('user_name', '').strip()
    if not email:
        email = 'viewer@company.com'
    if not name:
        name = DEFAULT_NAMES.get(email, email.split('@')[0].replace('.', ' ').title())
    return email, name

def get_user_profile(user_email=None):
    if not user_email:
        user_email, _ = get_user_info()
    user_email = user_email.lower().strip()
    
    conn = get_db()
    cur = conn.cursor()
    
    roles = {r['role_key']: r['authorized_email'].lower().strip() for r in cur.execute('SELECT role_key, authorized_email FROM permissions').fetchall()}
    assignments = cur.execute('SELECT * FROM platform_assignments').fetchall()
    conn.close()
    
    def parse_emails(email_str):
        if not email_str:
            return set()
        return {e.strip().lower() for e in email_str.split(',') if e.strip()}

    admin_emails = parse_emails(roles.get('ADMIN', ''))
    inventory_emails = parse_emails(roles.get('INVENTORY_MANAGER', ''))
    inventory_emails.add('joint_gt@company.com')

    online_dispatch_emails = parse_emails(roles.get('DISPATCH_ONLINE_EMAILS', ''))
    gt_dispatch_emails = parse_emails(roles.get('DISPATCH_GT_EMAILS', ''))
    gt_dispatch_emails.add('joint_gt@company.com')
    mt_dispatch_emails = parse_emails(roles.get('DISPATCH_MT_EMAILS', ''))
    
    general_dispatch_mgr = roles.get('DISPATCH_MANAGER', '').strip().lower()
    if general_dispatch_mgr:
        online_dispatch_emails.add(general_dispatch_mgr)
        gt_dispatch_emails.add(general_dispatch_mgr)
        mt_dispatch_emails.add(general_dispatch_mgr)

    is_admin = (user_email in admin_emails) or (user_email == roles.get('ADMIN', '').lower().strip())
    can_edit_inventory = is_admin or (user_email in inventory_emails)
    can_edit_dispatch_online = is_admin or (user_email in online_dispatch_emails)
    can_edit_dispatch_gt = is_admin or (user_email in gt_dispatch_emails)
    can_edit_dispatch_mt = is_admin or (user_email in mt_dispatch_emails)
    can_edit_dispatch = can_edit_dispatch_online or can_edit_dispatch_gt or can_edit_dispatch_mt
    
    if is_admin:
        allowed_courier_platforms = list(PLATFORMS)
        allowed_courier_channels = list(CHANNELS)
    else:
        allowed_courier_platforms = [r['platform_key'] for r in assignments if r['channel_type'] == 'online' and r['assigned_email'].lower().strip() == user_email]
        allowed_courier_channels = [r['platform_key'] for r in assignments if r['channel_type'] == 'gtmt' and r['assigned_email'].lower().strip() == user_email]
        
    if is_admin:
        role_title = 'Admin Manager (Full Access)'
    elif can_edit_inventory and can_edit_dispatch_gt and not can_edit_dispatch_online and not can_edit_dispatch_mt:
        role_title = 'Inventory & GT Dispatch Lead'
    elif can_edit_inventory and can_edit_dispatch:
        role_title = 'Inventory & All-Dispatch Lead'
    elif can_edit_inventory:
        role_title = 'Inventory Lead (Master Stock & Incoming Stock)'
    elif can_edit_dispatch_online and can_edit_dispatch_gt and can_edit_dispatch_mt:
        role_title = 'Dispatch Lead (All Platforms Dispatcher)'
    elif can_edit_dispatch_online:
        role_title = 'Online Dispatch Lead (Online Channels)'
    elif can_edit_dispatch_gt and can_edit_dispatch_mt:
        role_title = 'GT & MT Dispatch Lead (B2B Trade)'
    elif can_edit_dispatch_gt:
        role_title = 'GT Dispatch Lead (General Trade)'
    elif can_edit_dispatch_mt:
        role_title = 'MT Dispatch Lead (Modern Trade)'
    elif allowed_courier_platforms or allowed_courier_channels:
        names = []
        if allowed_courier_platforms:
            names.append(", ".join(allowed_courier_platforms[:2]) + ("..." if len(allowed_courier_platforms) > 2 else ""))
        if allowed_courier_channels:
            names.append(", ".join(allowed_courier_channels))
        role_title = f"Platform Lead ({' & '.join(names)})"
    else:
        role_title = 'Guest Viewer (Read-Only)'
            
    can_create_po = is_admin or bool(allowed_courier_platforms or allowed_courier_channels)
    is_platform_manager = is_admin or bool(allowed_courier_platforms or allowed_courier_channels)
    
    return {
        'user_email': user_email,
        'is_admin': is_admin,
        'can_edit_inventory': can_edit_inventory,
        'can_edit_dispatch': can_edit_dispatch,
        'can_edit_dispatch_online': can_edit_dispatch_online,
        'can_edit_dispatch_gt': can_edit_dispatch_gt,
        'can_edit_dispatch_mt': can_edit_dispatch_mt,
        'can_create_po': can_create_po,
        'is_platform_manager': is_platform_manager,
        'online_dispatch_emails': list(online_dispatch_emails),
        'gt_dispatch_emails': list(gt_dispatch_emails),
        'mt_dispatch_emails': list(mt_dispatch_emails),
        'allowed_courier_platforms': allowed_courier_platforms,
        'allowed_courier_channels': allowed_courier_channels,
        'allowed_platforms': allowed_courier_platforms,
        'allowed_channels': allowed_courier_channels,
        'role_title': role_title,
        'roles': roles
    }

def record_audit(conn, action_type, sheet_name, entity_ref, edit_summary, details=""):
    user_email, user_name = get_user_info()
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur = conn.cursor()
    cur.execute('''
    INSERT INTO audit_log (
        timestamp, action, details, user_name, user_email, action_type, sheet_name, entity_ref, edit_summary
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        now_str, action_type, details or edit_summary, user_name, user_email, action_type, sheet_name, entity_ref, edit_summary
    ))

@app.after_request
def add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    return response

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File too large. Maximum allowed upload size is 16 MB.'}), 413

PUBLIC_ENDPOINTS = {'login', 'logout', 'static', 'request_entity_too_large'}

@app.before_request
def require_login():
    if request.endpoint in PUBLIC_ENDPOINTS or (request.endpoint and request.endpoint.startswith('static')):
        return None
    if 'user_email' not in session:
        if request.path.startswith('/api/'):
            return jsonify({'error': 'Authentication required. Please log in.'}), 401
        return redirect(url_for('login', next=request.path))
    return None

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('user_email'):
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Authentication required. Please log in.'}), 401
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        if session.get('user_email'):
            return redirect(request.args.get('next') or url_for('index'))
        return render_template('login.html', next_url=request.args.get('next', ''))

    # POST
    data = request.form if request.form else (request.get_json(silent=True) or {})
    email = str(data.get('email', '')).strip().lower()
    password = str(data.get('password', ''))
    remember = bool(data.get('remember'))
    next_url = data.get('next') or request.args.get('next') or url_for('index')

    if not email or not password:
        err_msg = 'Please enter both email and password.'
        if request.is_json:
            return jsonify({'error': err_msg}), 400
        return render_template('login.html', error=err_msg, email=email, next_url=next_url), 400

    conn = get_db()
    cur = conn.cursor()
    user = cur.execute("SELECT * FROM users WHERE email = ? AND is_active = 1", (email,)).fetchone()
    
    if not user or not check_password_hash(user['password_hash'], password):
        conn.close()
        err_msg = 'Invalid email or password. Please verify your credentials.'
        if request.is_json:
            return jsonify({'error': err_msg}), 401
        return render_template('login.html', error=err_msg, email=email, next_url=next_url), 401

    session.permanent = remember
    session['user_email'] = user['email']
    session['user_name'] = user['full_name']
    
    try:
        record_audit(conn, 'LOGIN', 'Security & Access', user['email'], f"User {user['full_name']} logged in successfully")
        conn.commit()
    except Exception:
        pass
    conn.close()

    if request.is_json:
        return jsonify({'status': 'success', 'redirect': next_url})
    return redirect(next_url)

@app.route('/logout')
def logout():
    user_email = session.get('user_email')
    if user_email:
        try:
            conn = get_db()
            record_audit(conn, 'LOGOUT', 'Security & Access', user_email, f"User {user_email} logged out")
            conn.commit()
            conn.close()
        except Exception:
            pass
    session.clear()
    return redirect(url_for('login'))

@app.route('/api/change-password', methods=['POST'])
@login_required
def change_password():
    user_email = session.get('user_email')
    data = request.json or request.form or {}
    old_password = str(data.get('old_password', ''))
    new_password = str(data.get('new_password', ''))

    if not old_password or not new_password:
        return jsonify({'error': 'Current password and new password are required.'}), 400
    if len(new_password) < 6:
        return jsonify({'error': 'New password must be at least 6 characters.'}), 400

    conn = get_db()
    cur = conn.cursor()
    user = cur.execute("SELECT * FROM users WHERE email = ? AND is_active = 1", (user_email,)).fetchone()
    if not user or not check_password_hash(user['password_hash'], old_password):
        conn.close()
        return jsonify({'error': 'Incorrect current password.'}), 400

    new_hash = generate_password_hash(new_password)
    cur.execute("UPDATE users SET password_hash = ? WHERE email = ?", (new_hash, user_email))
    record_audit(conn, 'CHANGE_PASSWORD', 'Security & Access', user_email, f"User {user_email} changed their password")
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'message': 'Password updated successfully.'})

@app.route('/')
def index():
    return render_template('index.html', platforms=PLATFORMS)

@app.route('/api/user-profile')
def api_user_profile():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    profile['user_name'] = user_name
    return jsonify(profile)

@app.route('/api/permissions', methods=['GET'])
def get_permissions():
    conn = get_db()
    cur = conn.cursor()
    rows = cur.execute('SELECT * FROM permissions').fetchall()
    assignments = cur.execute('SELECT * FROM platform_assignments ORDER BY channel_type, platform_key').fetchall()
    conn.close()
    
    roles_map = {r['role_key']: {'email': r['authorized_email'], 'description': r['description']} for r in rows}
    assign_list = [dict(a) for a in assignments]
    return jsonify({
        'roles': roles_map,
        'assignments': assign_list
    })

@app.route('/api/permissions', methods=['POST'])
def update_permissions():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not profile['is_admin']:
        return jsonify({'error': 'Unauthorized: Only the Administrator can configure permissions and platform assignments.'}), 403
        
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    
    changes = []
    role_fields = [
        ('ADMIN', 'Admin'),
        ('INVENTORY_MANAGER', 'Inventory Lead'),
        ('DISPATCH_MANAGER', 'Dispatch Lead (All)'),
        ('DISPATCH_ONLINE_EMAILS', 'Online Dispatch Access'),
        ('DISPATCH_GT_EMAILS', 'GT Dispatch Access'),
        ('DISPATCH_MT_EMAILS', 'MT Dispatch Access')
    ]
    for r_key, r_label in role_fields:
        if r_key in data:
            val = str(data[r_key]).strip().lower()
            cur.execute('''
            INSERT INTO permissions (role_key, authorized_email, description)
            VALUES (?, ?, ?)
            ON CONFLICT(role_key) DO UPDATE SET authorized_email = excluded.authorized_email
            ''', (r_key, val, f'Authorized email(s) for {r_label}'))
            changes.append(f"{r_label} -> {val}")
            if val:
                default_pwd_hash = generate_password_hash('Admin@123')
                for single_email in [e.strip().lower() for e in val.split(',') if e.strip()]:
                    cur.execute('''
                    INSERT OR IGNORE INTO users (email, password_hash, full_name, is_active)
                    VALUES (?, ?, ?, 1)
                    ''', (single_email, default_pwd_hash, single_email.split('@')[0].title()))
        
    if 'assignments' in data and isinstance(data['assignments'], dict):
        for plat, info in data['assignments'].items():
            if isinstance(info, dict):
                p_email = info.get('email', '').strip().lower()
                p_name = info.get('name', '').strip()
            else:
                p_email = str(info).strip().lower()
                p_name = p_email.split('@')[0].title()
            if p_email:
                cur.execute('''
                UPDATE platform_assignments 
                SET assigned_email = ?, assigned_name = ?
                WHERE platform_key = ?
                ''', (p_email, p_name or plat, plat))
                changes.append(f"{plat} -> {p_email}")
                default_pwd_hash = generate_password_hash('Admin@123')
                cur.execute('''
                INSERT OR IGNORE INTO users (email, password_hash, full_name, is_active)
                VALUES (?, ?, ?, 1)
                ''', (p_email, default_pwd_hash, p_name or p_email.split('@')[0].title()))
                
    summary = ", ".join(changes) if changes else "No changes"
    record_audit(conn, 'UPDATE_PERMISSIONS', 'Security & Access', 'Assignments', f'Updated access settings: {summary}')
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Permissions successfully updated'})

@app.route('/api/stats')
def get_stats():
    conn = get_db()
    cur = conn.cursor()
    
    total_inv = cur.execute('SELECT SUM(total_inventory) FROM inventory_stock').fetchone()[0] or 0
    online_sent = cur.execute('SELECT SUM(actual_sent) FROM dispatch_online').fetchone()[0] or 0
    gtmt_sent = cur.execute('SELECT SUM(actual_sent) FROM dispatch_gt_mt').fetchone()[0] or 0
    total_adj = cur.execute('SELECT SUM(quantity) FROM adjustments').fetchone()[0] or 0
    total_sent = online_sent + gtmt_sent
    remaining_inv = total_inv - total_sent - total_adj
    
    online_courier_spent = cur.execute('SELECT SUM(courier_charges) FROM dispatch_online').fetchone()[0] or 0.0
    gtmt_courier_spent = cur.execute('SELECT SUM(courier_charges) FROM dispatch_gt_mt').fetchone()[0] or 0.0
    total_courier_charges = round(online_courier_spent + gtmt_courier_spent, 2)
    
    # Sales offtake total
    total_sales_units = cur.execute('SELECT SUM(units_sold) FROM sales_data').fetchone()[0] or 0
    total_sales_revenue = cur.execute('SELECT SUM(revenue) FROM sales_data').fetchone()[0] or 0.0
    pipeline_dark_store_stock = max(0, online_sent - total_sales_units)
    
    master_rows = get_master_data_internal(cur)
    low_stock_count = sum(1 for r in master_rows if r['status'] == 'Low Stock')
    total_skus = len(master_rows)
    
    po_online = cur.execute('''
        SELECT 
            SUM(po_quantity) as total_po_qty,
            SUM(actual_sent) as total_actual_sent
        FROM dispatch_online
    ''').fetchone()
    
    total_po_qty = (po_online['total_po_qty'] or 0)
    total_actual_po = (po_online['total_actual_sent'] or 0)
    fill_rate = round((total_actual_po / total_po_qty * 100), 1) if total_po_qty > 0 else 100.0
    
    conn.close()
    return jsonify({
        'total_inventory': total_inv,
        'online_sent': online_sent,
        'gtmt_sent': gtmt_sent,
        'total_sent': total_sent,
        'total_adjustments': total_adj,
        'remaining_inventory': remaining_inv,
        'total_skus': total_skus,
        'low_stock_count': low_stock_count,
        'total_po_qty': total_po_qty,
        'total_actual_po': total_actual_po,
        'fill_rate': fill_rate,
        'total_courier_charges': total_courier_charges,
        'total_sales_units': total_sales_units,
        'total_sales_revenue': total_sales_revenue,
        'pipeline_dark_store_stock': pipeline_dark_store_stock
    })

def get_master_data_internal(cur, brand=None):
    platform_cases = ', '.join([
        f"SUM(CASE WHEN o.platform = '{p}' THEN o.actual_sent ELSE 0 END) as sent_{p.lower().replace(' ', '_')}"
        for p in PLATFORMS
    ])
    
    where_clause = ""
    params = []
    if brand and brand != 'All':
        where_clause = "WHERE p.brand = ?"
        params.append(brand)
        
    query = f'''
    SELECT 
        p.sku,
        p.brand,
        p.name as product_name,
        p.category,
        COALESCE(s.total_inventory, 0) as total_inventory,
        COALESCE(s.reorder_threshold, 0) as reorder_threshold,
        {platform_cases},
        COALESCE(SUM(o.actual_sent), 0) as online_total,
        COALESCE(gt.gt_sent, 0) as gt_total,
        COALESCE(mt.mt_sent, 0) as mt_total,
        COALESCE(adj.adj_total, 0) as total_adjustments,
        COALESCE(sd.total_sold, 0) as total_sold
    FROM products p
    LEFT JOIN inventory_stock s ON p.sku = s.sku
    LEFT JOIN dispatch_online o ON p.sku = o.sku
    LEFT JOIN (
        SELECT sku, SUM(actual_sent) as gt_sent 
        FROM dispatch_gt_mt WHERE channel = 'GT' 
        GROUP BY sku
    ) gt ON p.sku = gt.sku
    LEFT JOIN (
        SELECT sku, SUM(actual_sent) as mt_sent 
        FROM dispatch_gt_mt WHERE channel = 'MT' 
        GROUP BY sku
    ) mt ON p.sku = mt.sku
    LEFT JOIN (
        SELECT sku, SUM(quantity) as adj_total 
        FROM adjustments 
        GROUP BY sku
    ) adj ON p.sku = adj.sku
    LEFT JOIN (
        SELECT sku, SUM(units_sold) as total_sold
        FROM sales_data
        GROUP BY sku
    ) sd ON p.sku = sd.sku
    {where_clause}
    GROUP BY p.sku
    ORDER BY p.brand, p.name
    '''
    
    rows = cur.execute(query, params).fetchall()
    results = []
    for r in rows:
        d = dict(r)
        grand_total = d['online_total'] + d['gt_total'] + d['mt_total']
        remaining = d['total_inventory'] - grand_total - d['total_adjustments']
        status = 'Low Stock' if remaining <= d['reorder_threshold'] else 'OK'
        d['grand_total'] = grand_total
        d['remaining_inventory'] = remaining
        d['status'] = status
        d['pipeline_stock'] = max(0, d['online_total'] - d.get('total_sold', 0))
        results.append(d)
    return results

@app.route('/api/master')
def get_master():
    brand = request.args.get('brand')
    conn = get_db()
    cur = conn.cursor()
    data = get_master_data_internal(cur, brand)
    conn.close()
    return jsonify(data)

# --- PLATFORM DEEP DIVE (ONLINE) ---
@app.route('/api/platform-detail')
def get_platform_detail():
    platform = request.args.get('platform', 'Blinkit')
    brand = request.args.get('brand', 'All')
    
    conn = get_db()
    cur = conn.cursor()
    
    where_prod = ""
    params_prod = [platform]
    if brand and brand != 'All':
        where_prod = "WHERE p.brand = ?"
        params_prod.append(brand)
        
    query_prod = f'''
    SELECT p.sku, p.brand, p.name as product_name, p.category,
           COALESCE(SUM(o.actual_sent), 0) as units_sent
    FROM products p
    LEFT JOIN dispatch_online o ON p.sku = o.sku AND o.platform = ?
    {where_prod}
    GROUP BY p.sku
    ORDER BY units_sent DESC, p.brand, p.name
    '''
    products_breakdown = [dict(r) for r in cur.execute(query_prod, params_prod).fetchall()]
    total_platform_sent = sum(r['units_sent'] for r in products_breakdown)
    for p in products_breakdown:
        p['share_percent'] = round((p['units_sent'] / total_platform_sent * 100), 1) if total_platform_sent > 0 else 0.0
        
    disp_query = 'SELECT * FROM dispatch_online WHERE platform = ?'
    disp_params = [platform]
    if brand and brand != 'All':
        disp_query += ' AND brand = ?'
        disp_params.append(brand)
    disp_query += ' ORDER BY id DESC'
    
    disp_rows = [dict(r) for r in cur.execute(disp_query, disp_params).fetchall()]
    for r in disp_rows:
        po_qty = r.get('po_quantity') or 0
        actual = r.get('actual_sent') or 0
        r['variance'] = po_qty - actual
        r['fulfillment_rate'] = round((actual / po_qty * 100), 1) if po_qty > 0 else 100.0
        
    locations = cur.execute('''
        SELECT location, SUM(actual_sent) as sent 
        FROM dispatch_online 
        WHERE platform = ? AND location != ''
        GROUP BY location 
        ORDER BY sent DESC 
        LIMIT 5
    ''', (platform,)).fetchall()
    
    # Platform sales offtake
    plat_sold = cur.execute('SELECT SUM(units_sold) FROM sales_data WHERE platform = ?', (platform,)).fetchone()[0] or 0
    plat_pipeline = max(0, total_platform_sent - plat_sold)
    
    total_po_qty = sum(r.get('po_quantity') or 0 for r in disp_rows)
    total_actual_sent = sum(r.get('actual_sent') or 0 for r in disp_rows)
    fill_rate = round((total_actual_sent / total_po_qty * 100), 1) if total_po_qty > 0 else 100.0
    
    conn.close()
    return jsonify({
        'platform': platform,
        'type': 'online',
        'total_units_sent': total_platform_sent,
        'total_units_sold': plat_sold,
        'pipeline_stock': plat_pipeline,
        'total_dispatches': len(disp_rows),
        'total_po_qty': total_po_qty,
        'fill_rate': fill_rate,
        'products': products_breakdown,
        'top_locations': [dict(l) for l in locations],
        'dispatches': disp_rows[:150]
    })

# --- GT / MT CHANNEL DEEP DIVE ---
@app.route('/api/gtmt-detail')
def get_gtmt_detail():
    channel = request.args.get('channel', 'GT')
    brand = request.args.get('brand', 'All')
    
    conn = get_db()
    cur = conn.cursor()
    
    where_prod = ""
    params_prod = [channel]
    if brand and brand != 'All':
        where_prod = "WHERE p.brand = ?"
        params_prod.append(brand)
        
    query_prod = f'''
    SELECT p.sku, p.brand, p.name as product_name, p.category,
           COALESCE(SUM(g.actual_sent), 0) as units_sent
    FROM products p
    LEFT JOIN dispatch_gt_mt g ON p.sku = g.sku AND g.channel = ?
    {where_prod}
    GROUP BY p.sku
    ORDER BY units_sent DESC, p.brand, p.name
    '''
    products_breakdown = [dict(r) for r in cur.execute(query_prod, params_prod).fetchall()]
    total_channel_sent = sum(r['units_sent'] for r in products_breakdown)
    for p in products_breakdown:
        p['share_percent'] = round((p['units_sent'] / total_channel_sent * 100), 1) if total_channel_sent > 0 else 0.0
        
    disp_query = 'SELECT * FROM dispatch_gt_mt WHERE channel = ?'
    disp_params = [channel]
    if brand and brand != 'All':
        disp_query += ' AND brand = ?'
        disp_params.append(brand)
    disp_query += ' ORDER BY id DESC'
    
    disp_rows = [dict(r) for r in cur.execute(disp_query, disp_params).fetchall()]
    for r in disp_rows:
        po_qty = r.get('po_quantity') or 0
        actual = r.get('actual_sent') or 0
        r['variance'] = po_qty - actual
        r['fulfillment_rate'] = round((actual / po_qty * 100), 1) if po_qty > 0 else 100.0
        
    top_buyers = cur.execute('''
        SELECT buyer_distributor, SUM(actual_sent) as sent 
        FROM dispatch_gt_mt 
        WHERE channel = ? AND buyer_distributor != ''
        GROUP BY buyer_distributor 
        ORDER BY sent DESC 
        LIMIT 5
    ''', (channel,)).fetchall()
    
    total_po_qty = sum(r.get('po_quantity') or 0 for r in disp_rows)
    total_actual_sent = sum(r.get('actual_sent') or 0 for r in disp_rows)
    fill_rate = round((total_actual_sent / total_po_qty * 100), 1) if total_po_qty > 0 else 100.0
    
    conn.close()
    return jsonify({
        'channel': channel,
        'type': 'gtmt',
        'channel_title': 'General Trade (GT)' if channel == 'GT' else 'Modern Trade (MT)',
        'total_units_sent': total_channel_sent,
        'total_dispatches': len(disp_rows),
        'total_po_qty': total_po_qty,
        'fill_rate': fill_rate,
        'products': products_breakdown,
        'top_buyers': [dict(b) for b in top_buyers],
        'dispatches': disp_rows
    })

@app.route('/api/products', methods=['GET', 'POST'])
def handle_products():
    if request.method == 'POST':
        user_email, user_name = get_user_info()
        profile = get_user_profile(user_email)
        if not (profile['is_admin'] or profile['can_edit_inventory']):
            return jsonify({'error': 'Permission Denied: Only Inventory Leads or Admins can add new products.'}), 403
            
        data = request.json or request.form
        sku = data.get('sku', '').strip().upper()
        name = data.get('name', '').strip()
        brand = data.get('brand', '').strip()
        category = data.get('category', 'General').strip() or 'General'
        try:
            initial_inventory = int(data.get('initial_inventory', 0))
        except:
            initial_inventory = 0
        try:
            reorder_threshold = int(data.get('reorder_threshold', 20))
        except:
            reorder_threshold = 20
        
        if not sku:
            return jsonify({'error': 'SKU code is required.'}), 400
        if not name:
            return jsonify({'error': 'Product name is required.'}), 400
        if not brand:
            return jsonify({'error': 'Brand name is required.'}), 400
            
        conn = get_db()
        cur = conn.cursor()
        cur.execute('''
            INSERT INTO products (sku, brand, name, category)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(sku) DO UPDATE SET
                brand = excluded.brand,
                name = excluded.name,
                category = excluded.category
        ''', (sku, brand, name, category))
        
        cur.execute('''
            INSERT INTO inventory_stock (sku, total_inventory, reorder_threshold)
            VALUES (?, ?, ?)
            ON CONFLICT(sku) DO UPDATE SET
                total_inventory = total_inventory + excluded.total_inventory,
                reorder_threshold = excluded.reorder_threshold
        ''', (sku, initial_inventory, reorder_threshold))
        
        record_audit(
            conn,
            'ADD_NEW_PRODUCT',
            'Inventory Stock & Batch Log',
            f'SKU {sku} | {brand}',
            f"Added new product '{name}' ({sku}) with initial stock {initial_inventory} and threshold {reorder_threshold}"
        )
        conn.commit()
        conn.close()
        
        return jsonify({'status': 'success', 'message': f"Product '{name}' ({sku}) successfully registered with {initial_inventory} units!"})

    conn = get_db()
    cur = conn.cursor()
    rows = cur.execute('SELECT sku, name, category, brand FROM products ORDER BY brand, name').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/products/<path:sku>', methods=['DELETE'])
def delete_product(sku):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied: Only Inventory Leads or Admins can remove products.'}), 403
        
    sku = sku.strip()
    conn = get_db()
    cur = conn.cursor()
    prod = cur.execute('SELECT sku, name, brand FROM products WHERE sku = ?', (sku,)).fetchone()
    if not prod:
        conn.close()
        return jsonify({'error': f"Product with SKU '{sku}' not found."}), 404
        
    name = prod['name']
    brand = prod['brand']
    
    # Delete from products and inventory_stock
    cur.execute('DELETE FROM inventory_stock WHERE sku = ?', (sku,))
    cur.execute('DELETE FROM products WHERE sku = ?', (sku,))
    
    cascade = request.args.get('cascade', '').lower() == 'true'
    if cascade:
        cur.execute('DELETE FROM stock_inward_batches WHERE sku = ?', (sku,))
        cur.execute('DELETE FROM dispatch_online WHERE sku = ?', (sku,))
        cur.execute('DELETE FROM dispatch_gt_mt WHERE sku = ?', (sku,))
        cur.execute('DELETE FROM adjustments WHERE sku = ?', (sku,))
        cur.execute('DELETE FROM platform_sales_data WHERE sku = ?', (sku,))
        
    record_audit(
        conn,
        'DELETE_PRODUCT',
        'Products & Stock Catalog',
        f"SKU {sku} | {brand}",
        f"Removed product '{name}' ({sku}) from portal."
    )
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'message': f"Product '{name}' ({sku}) has been successfully removed from the portal."
    })

@app.route('/api/brands', methods=['GET', 'POST'])
def handle_brands():
    email, name = get_user_info()
    profile = get_user_profile(email)
    
    if request.method == 'POST':
        if not (profile.get('is_admin') or profile.get('can_edit_inventory')):
            return jsonify({'error': 'Unauthorized: Only Admin Manager or Inventory Lead can add brands'}), 403
            
        data = request.get_json(silent=True) or request.form
        brand_name = (data.get('name') or data.get('brand') or '').strip()
        if not brand_name:
            return jsonify({'error': 'Brand name is required'}), 400
            
        conn = get_db()
        cur = conn.cursor()
        
        # Check if already exists (case-insensitive check)
        existing = cur.execute('SELECT name FROM brands WHERE LOWER(name) = LOWER(?)', (brand_name,)).fetchone()
        if not existing:
            # Also check products table in case it was used there
            existing = cur.execute('SELECT brand as name FROM products WHERE LOWER(brand) = LOWER(?)', (brand_name,)).fetchone()
            
        if existing:
            conn.close()
            return jsonify({'status': 'exists', 'brand': existing['name'], 'message': f"Brand '{existing['name']}' already exists."})
            
        cur.execute('INSERT INTO brands (name) VALUES (?)', (brand_name,))
        record_audit(conn, 'ADD_BRAND', 'Inventory & Brands', brand_name, f"Registered new brand '{brand_name}'")
        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'brand': brand_name, 'message': f"Brand '{brand_name}' added successfully."})

    # GET method
    conn = get_db()
    cur = conn.cursor()
    rows = cur.execute('''
        SELECT DISTINCT name FROM (
            SELECT name FROM brands WHERE name IS NOT NULL AND TRIM(name) != ""
            UNION
            SELECT brand AS name FROM products WHERE brand IS NOT NULL AND TRIM(brand) != ""
        ) ORDER BY name COLLATE NOCASE
    ''').fetchall()
    conn.close()
    return jsonify([r['name'] for r in rows if r['name']])

# --- DISPATCH ONLINE ---
@app.route('/api/dispatch-online', methods=['GET'])
def get_dispatch_online():
    brand = request.args.get('brand')
    platform = request.args.get('platform')
    search = request.args.get('search')
    
    conn = get_db()
    cur = conn.cursor()
    
    query = 'SELECT * FROM dispatch_online WHERE 1=1'
    params = []
    if brand and brand != 'All':
        query += ' AND brand = ?'
        params.append(brand)
    if platform and platform != 'All':
        query += ' AND platform = ?'
        params.append(platform)
    if search:
        query += ''' AND (
            product_name LIKE ? OR po_number LIKE ? OR location LIKE ? OR sku LIKE ? 
            OR courier_name LIKE ? OR tracking_id LIKE ? OR eway_bill_no LIKE ?
        )'''
        term = f'%{search}%'
        params.extend([term, term, term, term, term, term, term])
        
    query += ' ORDER BY id DESC'
    rows = cur.execute(query, params).fetchall()
    
    result = []
    for r in rows:
        d = dict(r)
        po_qty = d.get('po_quantity') or 0
        actual = d.get('actual_sent') or 0
        d['variance'] = po_qty - actual
        d['fulfillment_rate'] = round((actual / po_qty * 100), 1) if po_qty > 0 else 100.0
        result.append(d)
        
    conn.close()
    return jsonify(result)

@app.route('/api/dispatch-online', methods=['POST'])
def add_dispatch_online():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if not profile['can_edit_dispatch_online']:
        configured = ", ".join(profile.get('online_dispatch_emails', []))
        return jsonify({
            'error': f"Permission Denied: Your email ({user_email}) is not authorized to log Online Dispatches. Configured emails: {configured or 'None'}"
        }), 403
        
    data = request.json or {}
    platform = data.get('platform', '').strip()
    if not platform:
        return jsonify({'error': 'Platform is required'}), 400
        
    conn = get_db()
    cur = conn.cursor()
    
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    brand = data.get('brand', '').strip()
    category = data.get('category', '').strip()
    location = data.get('location', '').strip()
    po_number = data.get('po_number', '').strip()
    dispatch_date = data.get('dispatch_date') or datetime.date.today().strftime('%Y-%m-%d')
    appointment_date = data.get('appointment_date') or ''
    po_doc_url = data.get('po_doc_url', '').strip()
    po_doc_name = data.get('po_doc_name', '').strip()
    invoice_doc_url = data.get('invoice_doc_url', '').strip()
    invoice_doc_name = data.get('invoice_doc_name', '').strip()
    
    try:
        po_quantity = int(data.get('po_quantity', 0))
    except:
        po_quantity = 0
        
    try:
        actual_sent = int(data.get('actual_sent', 0))
    except:
        actual_sent = 0
        
    if sku:
        prod = cur.execute('SELECT name, brand, category FROM products WHERE sku = ?', (sku,)).fetchone()
        if prod:
            product_name = prod['name']
            brand = prod['brand']
            category = prod['category']
    elif product_name:
        prod = cur.execute('SELECT sku, brand, category FROM products WHERE LOWER(name) = LOWER(?)', (product_name,)).fetchone()
        if prod:
            sku = prod['sku']
            brand = prod['brand']
            category = prod['category']
            
    cur.execute('''
    INSERT INTO dispatch_online (
        dispatch_date, brand, product_name, sku, category, platform, location, 
        po_number, po_quantity, actual_sent, appointment_date, 
        courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '', 0.0, 'Dispatched', '', '', ?, ?, ?, ?)
    ''', (
        dispatch_date, brand, product_name, sku, category, platform, location, 
        po_number, po_quantity, actual_sent, appointment_date,
        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name
    ))
    new_id = cur.lastrowid
    
    edit_sum = f"Sent {actual_sent} units of {product_name} to {platform} (PO: {po_number}, PO Qty: {po_quantity}, Loc: {location})"
    record_audit(conn, 'ADD_ONLINE_DISPATCH', 'Dispatch Online', f'PO #{po_number} | SKU {sku}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'id': new_id})

@app.route('/api/dispatch-online/<int:item_id>/courier', methods=['POST'])
def update_dispatch_online_courier(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM dispatch_online WHERE id = ?', (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Dispatch record not found'}), 404
        
    platform = row['platform']
    can_edit = profile['is_admin'] or (platform in profile['allowed_courier_platforms'])
    if not can_edit:
        assigned_row = cur.execute('SELECT assigned_email, assigned_name FROM platform_assignments WHERE platform_key = ?', (platform,)).fetchone()
        owner_info = f"{assigned_row['assigned_name']} ({assigned_row['assigned_email']})" if assigned_row else "assigned platform lead"
        conn.close()
        return jsonify({
            'error': f"Permission Denied: Courier details for '{platform}' can only be updated by {owner_info}."
        }), 403
        
    data = request.json or {}
    courier_name = data.get('courier_name', '').strip()
    tracking_id = data.get('tracking_id', '').strip()
    try:
        courier_charges = float(data.get('courier_charges', 0.0))
    except:
        courier_charges = 0.0
    dispatch_status = data.get('dispatch_status', 'Dispatched').strip()
    eway_bill_no = data.get('eway_bill_no', '').strip()
    shipping_notes = data.get('shipping_notes', '').strip()
    try:
        actual_sent = int(data.get('actual_sent', row['actual_sent']))
    except:
        actual_sent = row['actual_sent']
    
    cur.execute('''
    UPDATE dispatch_online SET
        actual_sent = ?,
        courier_name = ?,
        tracking_id = ?,
        courier_charges = ?,
        dispatch_status = ?,
        eway_bill_no = ?,
        shipping_notes = ?
    WHERE id = ?
    ''', (actual_sent, courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes, item_id))
    
    # Sync dispatched quantity back to PO
    po_num = row['po_number']
    sku = row['sku']
    if po_num and sku:
        cur.execute('UPDATE po_items SET dispatched_quantity = ? WHERE po_number = ? AND sku = ?', (actual_sent, po_num, sku))
        po_row = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_num,)).fetchone()
        if po_row:
            po_id = po_row['id']
            stats = cur.execute('''
                SELECT SUM(ordered_quantity) as total_ordered, SUM(dispatched_quantity) as total_dispatched
                FROM po_items WHERE po_id = ?
            ''', (po_id,)).fetchone()
            tot_ord = stats['total_ordered'] or 0
            tot_disp = stats['total_dispatched'] or 0
            new_st = 'Fully Dispatched' if (tot_disp >= tot_ord and tot_ord > 0) else ('Partially Dispatched' if tot_disp > 0 else 'Pending Fulfillment')
            cur.execute('UPDATE purchase_orders SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_st, po_id))

    edit_sum = f"Fulfilled/Updated: {actual_sent} units sent via '{courier_name}', AWB: '{tracking_id}', Charges: ₹{courier_charges:,.2f}, Status: '{dispatch_status}' (Dispatch #{item_id}, {platform})"
    record_audit(conn, 'UPDATE_COURIER_ONLINE', 'Dispatch Online', f'Dispatch #{item_id} | {platform}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Courier & dispatch details updated successfully'})

@app.route('/api/dispatch-online/<int:item_id>', methods=['DELETE'])
def delete_dispatch_online(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if not profile['can_edit_dispatch_online']:
        return jsonify({'error': 'Permission Denied: Your email is not authorized to delete Online dispatches.'}), 403
        
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM dispatch_online WHERE id = ?', (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Not found'}), 404
        
    edit_sum = f"Deleted dispatch #{item_id}: {row['actual_sent']} units of {row['product_name']} to {row['platform']} (PO: {row['po_number']})"
    record_audit(conn, 'DELETE_ONLINE_DISPATCH', 'Dispatch Online', f'Dispatch #{item_id} | PO #{row["po_number"]}', edit_sum)
    cur.execute('DELETE FROM dispatch_online WHERE id = ?', (item_id,))

    # Reset PO item dispatched quantity if applicable
    po_num = row['po_number']
    sku = row['sku']
    if po_num and sku:
        cur.execute('UPDATE po_items SET dispatched_quantity = 0 WHERE po_number = ? AND sku = ?', (po_num, sku))
        po_row = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_num,)).fetchone()
        if po_row:
            po_id = po_row['id']
            stats = cur.execute('''
                SELECT SUM(ordered_quantity) as total_ordered, SUM(dispatched_quantity) as total_dispatched
                FROM po_items WHERE po_id = ?
            ''', (po_id,)).fetchone()
            tot_ord = stats['total_ordered'] or 0
            tot_disp = stats['total_dispatched'] or 0
            new_st = 'Fully Dispatched' if (tot_disp >= tot_ord and tot_ord > 0) else ('Partially Dispatched' if tot_disp > 0 else 'Pending Fulfillment')
            cur.execute('UPDATE purchase_orders SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_st, po_id))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# --- DISPATCH GT MT ---
@app.route('/api/dispatch-gt-mt', methods=['GET'])
def get_dispatch_gt_mt():
    brand = request.args.get('brand')
    channel = request.args.get('channel')
    search = request.args.get('search')
    
    conn = get_db()
    cur = conn.cursor()
    
    query = 'SELECT * FROM dispatch_gt_mt WHERE 1=1'
    params = []
    if brand and brand != 'All':
        query += ' AND brand = ?'
        params.append(brand)
    if channel and channel != 'All':
        query += ' AND channel = ?'
        params.append(channel)
    if search:
        query += ''' AND (
            product_name LIKE ? OR buyer_distributor LIKE ? OR location LIKE ? 
            OR invoice_no LIKE ? OR po_number LIKE ? OR courier_name LIKE ? OR tracking_id LIKE ?
        )'''
        term = f'%{search}%'
        params.extend([term, term, term, term, term, term, term])
        
    query += ' ORDER BY id DESC'
    rows = cur.execute(query, params).fetchall()
    
    result = []
    for r in rows:
        d = dict(r)
        po_qty = d.get('po_quantity') or 0
        actual = d.get('actual_sent') or 0
        d['variance'] = po_qty - actual
        d['fulfillment_rate'] = round((actual / po_qty * 100), 1) if po_qty > 0 else 100.0
        result.append(d)
        
    conn.close()
    return jsonify(result)

@app.route('/api/dispatch-gt-mt', methods=['POST'])
def add_dispatch_gt_mt():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    data = request.json or {}
    channel = data.get('channel', 'GT').strip().upper()
    if channel not in ['GT', 'MT']:
        channel = 'GT'
        
    if channel == 'GT' and not profile['can_edit_dispatch_gt']:
        configured = ", ".join(profile.get('gt_dispatch_emails', []))
        return jsonify({
            'error': f"Permission Denied: Your email ({user_email}) is not authorized to log GT Dispatches. Configured emails: {configured or 'None'}"
        }), 403
    elif channel == 'MT' and not profile['can_edit_dispatch_mt']:
        configured = ", ".join(profile.get('mt_dispatch_emails', []))
        return jsonify({
            'error': f"Permission Denied: Your email ({user_email}) is not authorized to log MT Dispatches. Configured emails: {configured or 'None'}"
        }), 403
        
    conn = get_db()
    cur = conn.cursor()
    
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    brand = data.get('brand', '').strip()
    category = data.get('category', '').strip()
    buyer = data.get('buyer_distributor', '').strip()
    location = data.get('location', '').strip()
    po_number = data.get('po_number', '').strip()
    invoice_no = data.get('invoice_no', '').strip()
    dispatch_date = data.get('dispatch_date') or datetime.date.today().strftime('%Y-%m-%d')
    appointment_date = data.get('appointment_date') or ''
    po_doc_url = data.get('po_doc_url', '').strip()
    po_doc_name = data.get('po_doc_name', '').strip()
    invoice_doc_url = data.get('invoice_doc_url', '').strip()
    invoice_doc_name = data.get('invoice_doc_name', '').strip()
    
    try:
        po_quantity = int(data.get('po_quantity', 0))
    except:
        po_quantity = 0
        
    try:
        actual_sent = int(data.get('actual_sent', 0))
    except:
        actual_sent = 0
        
    if sku:
        prod = cur.execute('SELECT name, brand, category FROM products WHERE sku = ?', (sku,)).fetchone()
        if prod:
            product_name = prod['name']
            brand = prod['brand']
            category = prod['category']
    elif product_name:
        prod = cur.execute('SELECT sku, brand, category FROM products WHERE LOWER(name) = LOWER(?)', (product_name,)).fetchone()
        if prod:
            sku = prod['sku']
            brand = prod['brand']
            category = prod['category']
            
    cur.execute('''
    INSERT INTO dispatch_gt_mt (
        dispatch_date, brand, product_name, sku, category, channel, buyer_distributor, location, 
        po_number, invoice_no, po_quantity, actual_sent, appointment_date,
        courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '', 0.0, 'Dispatched', '', '', ?, ?, ?, ?)
    ''', (
        dispatch_date, brand, product_name, sku, category, channel, buyer, location, 
        po_number, invoice_no, po_quantity, actual_sent, appointment_date,
        po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name
    ))
    new_id = cur.lastrowid
    
    edit_sum = f"Sent {actual_sent} units of {product_name} via {channel} to {buyer} (Inv: {invoice_no}, PO: {po_number})"
    record_audit(conn, 'ADD_GTMT_DISPATCH', 'Dispatch GT MT', f'{channel} | {buyer}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'id': new_id})

@app.route('/api/dispatch-gt-mt/<int:item_id>/courier', methods=['POST'])
def update_dispatch_gtmt_courier(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM dispatch_gt_mt WHERE id = ?', (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Dispatch record not found'}), 404
        
    channel = row['channel']
    can_edit = profile['is_admin'] or (channel in profile['allowed_courier_channels'])
    if not can_edit:
        assigned_row = cur.execute('SELECT assigned_email, assigned_name FROM platform_assignments WHERE platform_key = ?', (channel,)).fetchone()
        owner_info = f"{assigned_row['assigned_name']} ({assigned_row['assigned_email']})" if assigned_row else "assigned trade lead"
        conn.close()
        return jsonify({
            'error': f"Permission Denied: Transporter details for '{channel}' can only be updated by {owner_info}."
        }), 403
        
    data = request.json or {}
    courier_name = data.get('courier_name', '').strip()
    tracking_id = data.get('tracking_id', '').strip()
    try:
        courier_charges = float(data.get('courier_charges', 0.0))
    except:
        courier_charges = 0.0
    dispatch_status = data.get('dispatch_status', 'Dispatched').strip()
    eway_bill_no = data.get('eway_bill_no', '').strip()
    shipping_notes = data.get('shipping_notes', '').strip()
    try:
        actual_sent = int(data.get('actual_sent', row['actual_sent']))
    except:
        actual_sent = row['actual_sent']
    
    cur.execute('''
    UPDATE dispatch_gt_mt SET
        actual_sent = ?,
        courier_name = ?,
        tracking_id = ?,
        courier_charges = ?,
        dispatch_status = ?,
        eway_bill_no = ?,
        shipping_notes = ?
    WHERE id = ?
    ''', (actual_sent, courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes, item_id))
    
    # Sync dispatched quantity back to PO
    po_num = row['po_number']
    sku = row['sku']
    if po_num and sku:
        cur.execute('UPDATE po_items SET dispatched_quantity = ? WHERE po_number = ? AND sku = ?', (actual_sent, po_num, sku))
        po_row = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_num,)).fetchone()
        if po_row:
            po_id = po_row['id']
            stats = cur.execute('''
                SELECT SUM(ordered_quantity) as total_ordered, SUM(dispatched_quantity) as total_dispatched
                FROM po_items WHERE po_id = ?
            ''', (po_id,)).fetchone()
            tot_ord = stats['total_ordered'] or 0
            tot_disp = stats['total_dispatched'] or 0
            new_st = 'Fully Dispatched' if (tot_disp >= tot_ord and tot_ord > 0) else ('Partially Dispatched' if tot_disp > 0 else 'Pending Fulfillment')
            cur.execute('UPDATE purchase_orders SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_st, po_id))

    edit_sum = f"Fulfilled/Updated: {actual_sent} units sent via Transporter: '{courier_name}', LR/AWB: '{tracking_id}', Charges: ₹{courier_charges:,.2f}, Status: '{dispatch_status}' (GT/MT #{item_id}, {row['channel']} - {row['buyer_distributor']})"
    record_audit(conn, 'UPDATE_COURIER_GTMT', 'Dispatch GT MT', f'Dispatch #{item_id} | {row["channel"]}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Transporter & dispatch details updated successfully'})

@app.route('/api/dispatch-gt-mt/<int:item_id>', methods=['DELETE'])
def delete_dispatch_gt_mt(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM dispatch_gt_mt WHERE id = ?', (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Not found'}), 404
        
    channel = row['channel']
    if channel == 'GT' and not profile['can_edit_dispatch_gt']:
        conn.close()
        return jsonify({'error': 'Permission Denied: Your email is not authorized to delete GT dispatches.'}), 403
    elif channel == 'MT' and not profile['can_edit_dispatch_mt']:
        conn.close()
        return jsonify({'error': 'Permission Denied: Your email is not authorized to delete MT dispatches.'}), 403
        
    edit_sum = f"Deleted GT/MT dispatch #{item_id}: {row['actual_sent']} units of {row['product_name']} via {row['channel']} to {row['buyer_distributor']}"
    record_audit(conn, 'DELETE_GTMT_DISPATCH', 'Dispatch GT MT', f'Dispatch #{item_id} | {row["channel"]}', edit_sum)
    cur.execute('DELETE FROM dispatch_gt_mt WHERE id = ?', (item_id,))

    # Reset PO item dispatched quantity if applicable
    po_num = row['po_number']
    sku = row['sku']
    if po_num and sku:
        cur.execute('UPDATE po_items SET dispatched_quantity = 0 WHERE po_number = ? AND sku = ?', (po_num, sku))
        po_row = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_num,)).fetchone()
        if po_row:
            po_id = po_row['id']
            stats = cur.execute('''
                SELECT SUM(ordered_quantity) as total_ordered, SUM(dispatched_quantity) as total_dispatched
                FROM po_items WHERE po_id = ?
            ''', (po_id,)).fetchone()
            tot_ord = stats['total_ordered'] or 0
            tot_disp = stats['total_dispatched'] or 0
            new_st = 'Fully Dispatched' if (tot_disp >= tot_ord and tot_ord > 0) else ('Partially Dispatched' if tot_disp > 0 else 'Pending Fulfillment')
            cur.execute('UPDATE purchase_orders SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_st, po_id))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# --- DISPATCH DOCUMENT UPLOADS (PO & INVOICE COPIES - AUTO-SYNCED ACROSS PO) ---
@app.route('/api/dispatch/upload-docs', methods=['POST'])
def upload_dispatch_docs():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file attached'}), 400
        
    file = request.files['file']
    if not file or not file.filename:
        return jsonify({'error': 'Empty file or filename'}), 400
        
    dispatch_type = request.form.get('dispatch_type', 'online').strip().lower()
    doc_type = request.form.get('doc_type', 'po').strip().lower() # 'po' or 'invoice'
    item_id = request.form.get('dispatch_id', '').strip()
    po_number = request.form.get('po_number', '').strip()
    
    if not item_id and not po_number:
        return jsonify({'error': 'Dispatch ID or PO Number is required'}), 400
        
    conn = get_db()
    cur = conn.cursor()
    
    table_name = 'dispatch_online' if dispatch_type == 'online' else 'dispatch_gt_mt'
    if item_id:
        row = cur.execute(f"SELECT * FROM {table_name} WHERE id = ?", (item_id,)).fetchone()
    else:
        row = cur.execute(f"SELECT * FROM {table_name} WHERE po_number = ? LIMIT 1", (po_number,)).fetchone()
        
    if not row:
        conn.close()
        return jsonify({'error': 'Dispatch record not found'}), 404
        
    if dispatch_type == 'online':
        platform = row['platform']
        can_upload = profile['is_admin'] or profile['can_edit_dispatch_online'] or (platform in profile['allowed_courier_platforms'])
    else:
        channel = row['channel']
        can_upload = profile['is_admin'] or (channel == 'GT' and profile['can_edit_dispatch_gt']) or (channel == 'MT' and profile['can_edit_dispatch_mt']) or (channel in profile['allowed_courier_channels'])
        
    if not can_upload:
        conn.close()
        return jsonify({'error': 'Permission Denied: You are not authorized to upload documents for this dispatch.'}), 403
        
    ext = os.path.splitext(file.filename)[1].lower()
    allowed_exts = ['.pdf', '.png', '.jpg', '.jpeg', '.webp', '.xlsx', '.xls', '.csv', '.doc', '.docx']
    if ext not in allowed_exts:
        conn.close()
        return jsonify({'error': f'Unsupported file format {ext}. Allowed: PDF, Images, Excel, Word'}), 400
        
    clean_base = "".join(c for c in os.path.splitext(file.filename)[0] if c.isalnum() or c in ('-', '_')).strip() or 'doc'
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    po_ref_clean = (row['po_number'] or str(row['id'])).replace('/', '_').replace('-', '_')
    filename = f"{dispatch_type}_{po_ref_clean}_{doc_type}_{timestamp}_{clean_base}{ext}"
    dest_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(dest_path)
    
    doc_url = f"/uploads/dispatch_docs/{filename}"
    doc_name = file.filename
    
    po_num = row['po_number'] or po_number
    if po_num:
        # Grouped PO sync: One upload attaches document across ALL products under this PO!
        if doc_type == 'po':
            cur.execute(f"UPDATE {table_name} SET po_doc_url = ?, po_doc_name = ? WHERE po_number = ?", (doc_url, doc_name, po_num))
            cur.execute("UPDATE purchase_orders SET po_doc_url = ?, po_doc_name = ? WHERE po_number = ?", (doc_url, doc_name, po_num))
        else:
            cur.execute(f"UPDATE {table_name} SET invoice_doc_url = ?, invoice_doc_name = ? WHERE po_number = ?", (doc_url, doc_name, po_num))
    else:
        if doc_type == 'po':
            cur.execute(f"UPDATE {table_name} SET po_doc_url = ?, po_doc_name = ? WHERE id = ?", (doc_url, doc_name, row['id']))
        else:
            cur.execute(f"UPDATE {table_name} SET invoice_doc_url = ?, invoice_doc_name = ? WHERE id = ?", (doc_url, doc_name, row['id']))
        
    audit_label = 'PO Copy' if doc_type == 'po' else 'Invoice Copy'
    channel_ref = row['platform'] if dispatch_type == 'online' else row['channel']
    record_audit(conn, 'UPLOAD_DISPATCH_DOC', f"Dispatch {dispatch_type.upper()}", f"PO #{po_num} | {channel_ref}", f"Attached {audit_label}: '{doc_name}' for entire PO consignment")
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'doc_type': doc_type,
        'doc_url': doc_url,
        'doc_name': doc_name,
        'po_number': po_num,
        'message': f"{audit_label} attached to all products under PO '{po_num}'"
    })

# --- BATCH / PO CONSIGNMENT DISPATCH FULFILLMENT ---
@app.route('/api/dispatch/fulfill-po', methods=['POST'])
def fulfill_po_consignment():
    """Fulfill or update logistics for an entire PO consignment at once (courier, tracking, charges, invoice, actual units)."""
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    data = request.json or {}
    dispatch_type = str(data.get('dispatch_type', 'online')).strip().lower() # 'online' or 'gtmt'
    po_number = str(data.get('po_number', '')).strip()
    
    if not po_number:
        return jsonify({'error': 'PO Number is required'}), 400
        
    table_name = 'dispatch_online' if dispatch_type == 'online' else 'dispatch_gt_mt'
    conn = get_db()
    cur = conn.cursor()
    
    rows = cur.execute(f"SELECT * FROM {table_name} WHERE po_number = ?", (po_number,)).fetchall()
    if not rows:
        conn.close()
        return jsonify({'error': f"No dispatch records found for PO '{po_number}'"}), 404
        
    # Check permissions
    first_row = rows[0]
    if dispatch_type == 'online':
        platform = first_row['platform']
        can_edit = profile['is_admin'] or profile['can_edit_dispatch_online'] or (platform in profile['allowed_courier_platforms'])
    else:
        channel = first_row['channel']
        can_edit = profile['is_admin'] or (channel == 'GT' and profile['can_edit_dispatch_gt']) or (channel == 'MT' and profile['can_edit_dispatch_mt']) or (channel in profile['allowed_courier_channels'])
        
    if not can_edit:
        conn.close()
        return jsonify({'error': 'Permission Denied: Unauthorized to update this consignment dispatch.'}), 403
        
    courier_name = str(data.get('courier_name', '')).strip()
    tracking_id = str(data.get('tracking_id', '')).strip()
    try:
        courier_charges = float(data.get('courier_charges', 0.0))
    except:
        courier_charges = 0.0
    dispatch_status = str(data.get('dispatch_status', 'Dispatched')).strip()
    eway_bill_no = str(data.get('eway_bill_no', '')).strip()
    shipping_notes = str(data.get('shipping_notes', '')).strip()
    
    # Items: array of {id, sku, actual_sent}
    items_input = data.get('items', [])
    items_by_id = {it.get('id'): it for it in items_input if it.get('id') is not None}
    items_by_sku = {it.get('sku'): it for it in items_input if it.get('sku')}
    
    # Proportionate allocation of total courier charges across items
    total_ord_qty = sum(r['po_quantity'] or 1 for r in rows) or 1
    
    for idx, r in enumerate(rows):
        r_id = r['id']
        r_sku = r['sku']
        
        # Determine actual_sent for this item
        it_info = items_by_id.get(r_id) or items_by_sku.get(r_sku)
        if it_info and 'actual_sent' in it_info:
            try:
                item_actual_sent = int(it_info['actual_sent'])
            except:
                item_actual_sent = r['actual_sent']
        else:
            item_actual_sent = r['actual_sent'] if r['actual_sent'] > 0 else (r['po_quantity'] or 0)
            
        if courier_charges > 0:
            item_charge = round(courier_charges * ((r['po_quantity'] or 1) / total_ord_qty), 2)
        else:
            item_charge = 0.0
            
        cur.execute(f'''
        UPDATE {table_name} SET
            actual_sent = ?,
            courier_name = ?,
            tracking_id = ?,
            courier_charges = ?,
            dispatch_status = ?,
            eway_bill_no = ?,
            shipping_notes = ?
        WHERE id = ?
        ''', (item_actual_sent, courier_name, tracking_id, item_charge, dispatch_status, eway_bill_no, shipping_notes, r_id))
        
        # Update po_items
        cur.execute('''
        UPDATE po_items SET dispatched_quantity = ? WHERE po_number = ? AND sku = ?
        ''', (item_actual_sent, po_number, r_sku))
        
    # Recalculate parent PO overall status
    po_row = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_number,)).fetchone()
    if po_row:
        po_id = po_row['id']
        stats = cur.execute('''
            SELECT SUM(ordered_quantity) as total_ordered, SUM(dispatched_quantity) as total_dispatched
            FROM po_items WHERE po_id = ?
        ''', (po_id,)).fetchone()
        tot_ord = stats['total_ordered'] or 0
        tot_disp = stats['total_dispatched'] or 0
        new_st = 'Fully Dispatched' if (tot_disp >= tot_ord and tot_ord > 0) else ('Partially Dispatched' if tot_disp > 0 else 'Pending Fulfillment')
        cur.execute('UPDATE purchase_orders SET status = ?, is_packed = 1, is_pickup_ready = 1, is_dispatched = 1, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_st, po_id))
        
    audit_target = first_row['platform'] if dispatch_type == 'online' else first_row['channel']
    edit_sum = f"Consignment PO {po_number} fulfilled: Courier '{courier_name}', AWB '{tracking_id}', Total Charges ₹{courier_charges:,.2f}, Status '{dispatch_status}' across {len(rows)} products"
    record_audit(conn, 'FULFILL_PO_CONSIGNMENT', f"Dispatch {dispatch_type.upper()}", f"PO #{po_number} | {audit_target}", edit_sum)
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'message': f"Consignment for PO '{po_number}' updated successfully across all {len(rows)} products.",
        'po_number': po_number
    })

@app.route('/uploads/dispatch_docs/<path:filename>')
def serve_dispatch_doc(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

# --- PURCHASE ORDER CREATION & MANAGEMENT ---
@app.route('/api/purchase-orders', methods=['GET'])
def get_purchase_orders():
    conn = get_db()
    cur = conn.cursor()
    
    platform = request.args.get('platform', 'All')
    status = request.args.get('status', 'All')
    search = request.args.get('search', '').strip().lower()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    packed = request.args.get('packed', 'All')
    dispatch_status = request.args.get('dispatch_status', 'All')
    
    query = 'SELECT * FROM purchase_orders WHERE 1=1'
    params = []
    
    if platform and platform != 'All':
        query += ' AND platform_or_channel = ?'
        params.append(platform)
    if status and status != 'All':
        query += ' AND status = ?'
        params.append(status)
    if packed == 'Yes':
        query += ' AND is_packed = 1'
    elif packed == 'No':
        query += ' AND (is_packed = 0 OR is_packed IS NULL)'
    if dispatch_status == 'Yes':
        query += " AND (is_dispatched = 1 OR status IN ('Partially Dispatched', 'Fully Dispatched', 'GRN Done'))"
    elif dispatch_status == 'No':
        query += " AND (is_dispatched = 0 OR is_dispatched IS NULL) AND (status NOT IN ('Partially Dispatched', 'Fully Dispatched', 'GRN Done'))"
    if start_date:
        query += ' AND po_date >= ?'
        params.append(start_date)
    if end_date:
        query += ' AND po_date <= ?'
        params.append(end_date)
    if search:
        query += ' AND (LOWER(po_number) LIKE ? OR LOWER(location) LIKE ? OR LOWER(distributor_name) LIKE ?)'
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
        
    query += ' ORDER BY po_date DESC, id DESC'
    pos = [dict(r) for r in cur.execute(query, params).fetchall()]
    
    # Fetch line items for each PO
    for po in pos:
        items = cur.execute('SELECT * FROM po_items WHERE po_id = ? ORDER BY brand, product_name', (po['id'],)).fetchall()
        item_list = []
        for it in items:
            d = dict(it)
            d['quantity'] = d['ordered_quantity']
            d['line_total'] = d['total_value']
            item_list.append(d)
        po['items'] = item_list
        po['total_units'] = po.get('total_quantity', 0)
        po['total_value'] = po.get('total_po_value', 0.0)
        po['item_count'] = po.get('total_items', len(item_list))
        
        # Computed status fields for Yes/No columns
        is_disp = bool(po.get('is_dispatched')) or (po.get('status') in ['Partially Dispatched', 'Fully Dispatched', 'GRN Done'])
        po['is_dispatched'] = 1 if is_disp else 0
        po['is_packed'] = 1 if po.get('is_packed') else 0
        po['is_pickup_ready'] = 1 if po.get('is_pickup_ready') else 0
        po['packed_yes_no'] = 'Yes' if po['is_packed'] else 'No'
        po['pickup_ready_yes_no'] = 'Yes' if po['is_pickup_ready'] else 'No'
        po['dispatch_status_yes_no'] = 'Yes' if is_disp else 'No'
        
    all_pos = cur.execute('SELECT COUNT(*), SUM(total_quantity), SUM(total_po_value) FROM purchase_orders').fetchone()
    pending_cnt = cur.execute("SELECT COUNT(*) FROM purchase_orders WHERE status = 'Pending Fulfillment'").fetchone()[0]
    
    stats = {
        'total_pos': all_pos[0] or 0,
        'pending_pos': pending_cnt or 0,
        'total_units': all_pos[1] or 0,
        'total_value': round(all_pos[2] or 0.0, 2)
    }
    
    conn.close()
    return jsonify({
        'purchase_orders': pos,
        'stats': stats
    })

@app.route('/api/purchase-orders/<int:po_id>', methods=['GET'])
def get_purchase_order_detail(po_id):
    conn = get_db()
    cur = conn.cursor()
    po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    if not po:
        conn.close()
        return jsonify({'error': 'Purchase order not found'}), 404
        
    po_data = dict(po)
    items = cur.execute('SELECT * FROM po_items WHERE po_id = ? ORDER BY brand, product_name', (po_id,)).fetchall()
    item_list = []
    for it in items:
        d = dict(it)
        d['quantity'] = d['ordered_quantity']
        d['line_total'] = d['total_value']
        item_list.append(d)
    po_data['items'] = item_list
    po_data['total_units'] = po_data.get('total_quantity', 0)
    po_data['total_value'] = po_data.get('total_po_value', 0.0)
    po_data['item_count'] = po_data.get('total_items', len(item_list))
    
    is_disp = bool(po_data.get('is_dispatched')) or (po_data.get('status') in ['Partially Dispatched', 'Fully Dispatched', 'GRN Done'])
    po_data['is_dispatched'] = 1 if is_disp else 0
    po_data['is_packed'] = 1 if po_data.get('is_packed') else 0
    po_data['is_pickup_ready'] = 1 if po_data.get('is_pickup_ready') else 0
    po_data['packed_yes_no'] = 'Yes' if po_data['is_packed'] else 'No'
    po_data['pickup_ready_yes_no'] = 'Yes' if po_data['is_pickup_ready'] else 'No'
    po_data['dispatch_status_yes_no'] = 'Yes' if is_disp else 'No'
    
    conn.close()
    return jsonify({
        'status': 'success',
        'purchase_order': po_data,
        **po_data
    })

@app.route('/api/purchase-orders/<int:po_id>/toggle', methods=['POST'])
def toggle_po_status(po_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    if not po:
        conn.close()
        return jsonify({'error': 'Purchase order not found'}), 404
        
    plat = po['platform_or_channel']
    is_plat_mgr = profile.get('is_admin') or (plat in profile.get('allowed_courier_platforms', [])) or (plat in profile.get('allowed_courier_channels', []))
    can_toggle = profile.get('is_admin') or profile.get('can_edit_dispatch') or is_plat_mgr
    
    if not can_toggle:
        conn.close()
        return jsonify({'error': 'You do not have permission to update packing or dispatch status for this PO.'}), 403
        
    data = request.json or {}
    field = data.get('field')  # 'packed', 'pickup_ready', 'dispatched'
    val = data.get('value')    # 1 or 0, or None to invert
    
    if field == 'packed':
        new_val = (1 - (po['is_packed'] or 0)) if val is None else (1 if val else 0)
        if new_val == 0:
            cur.execute('UPDATE purchase_orders SET is_packed = 0, is_pickup_ready = 0, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (po_id,))
        else:
            cur.execute('UPDATE purchase_orders SET is_packed = 1, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (po_id,))
        msg = f"PO {po['po_number']} Packed status set to {'Yes' if new_val else 'No'}."
        record_audit(conn, 'UPDATE_PO_PACKED', 'Purchase Orders', po['po_number'], f"{msg} by {user_name}")
    elif field == 'pickup_ready':
        new_val = (1 - (po['is_pickup_ready'] or 0)) if val is None else (1 if val else 0)
        if new_val == 1:
            cur.execute('UPDATE purchase_orders SET is_pickup_ready = 1, is_packed = 1, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (po_id,))
        else:
            cur.execute('UPDATE purchase_orders SET is_pickup_ready = 0, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (po_id,))
        msg = f"PO {po['po_number']} Pickup (Packed & Ready) set to {'Yes' if new_val else 'No'}."
        record_audit(conn, 'UPDATE_PO_PICKUP_READY', 'Purchase Orders', po['po_number'], f"{msg} by {user_name}")
    elif field == 'dispatched':
        new_val = (1 - (po['is_dispatched'] or 0)) if val is None else (1 if val else 0)
        cur.execute('UPDATE purchase_orders SET is_dispatched = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_val, po_id))
        msg = f"PO {po['po_number']} Dispatch Status set to {'Yes' if new_val else 'No'}."
        record_audit(conn, 'UPDATE_PO_DISPATCHED', 'Purchase Orders', po['po_number'], f"{msg} by {user_name}")
    else:
        conn.close()
        return jsonify({'error': 'Invalid toggle field. Must be packed, pickup_ready, or dispatched.'}), 400
        
    conn.commit()
    updated_po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    conn.close()
    
    up_dict = dict(updated_po)
    up_dict['packed_yes_no'] = 'Yes' if up_dict.get('is_packed') else 'No'
    up_dict['pickup_ready_yes_no'] = 'Yes' if up_dict.get('is_pickup_ready') else 'No'
    up_dict['dispatch_status_yes_no'] = 'Yes' if (up_dict.get('is_dispatched') or up_dict.get('status') in ['Partially Dispatched', 'Fully Dispatched', 'GRN Done']) else 'No'

    return jsonify({
        'status': 'success',
        'message': msg,
        'purchase_order': up_dict,
        **up_dict
    })

@app.route('/api/purchase-orders', methods=['POST'])
def create_purchase_order():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    data = request.json or {}
    platform_or_channel = str(data.get('platform_or_channel', 'Blinkit')).strip()

    # Permission check: PO creation is strictly for dedicated Platform Managers (or Admin)
    is_platform_mgr = profile['is_admin'] or \
        (platform_or_channel in profile['allowed_courier_platforms']) or \
        (platform_or_channel in profile['allowed_courier_channels'])
        
    if not is_platform_mgr:
        return jsonify({
            'error': f"Permission Denied: Only the dedicated Platform Manager for '{platform_or_channel}' (or Admin) can create Purchase Orders for this channel."
        }), 403
        
    po_number = str(data.get('po_number', '')).strip()
    po_date = str(data.get('po_date', '')).strip() or datetime.date.today().strftime('%Y-%m-%d')
    channel_type = 'gtmt' if platform_or_channel in CHANNELS else 'online'
    location = str(data.get('location', '')).strip()
    distributor_name = str(data.get('distributor_name', '')).strip()
    appointment_date = str(data.get('appointment_date', '')).strip()
    notes = str(data.get('notes', '')).strip()
    po_doc_url = str(data.get('po_doc_url', '')).strip()
    po_doc_name = str(data.get('po_doc_name', '')).strip()
    status = str(data.get('status', 'Pending Fulfillment')).strip()
    
    items = data.get('items', [])
    if not po_number:
        return jsonify({'error': 'PO Number is required.'}), 400
    if not items or len(items) == 0:
        return jsonify({'error': 'A Purchase Order must contain at least 1 product line item.'}), 400
        
    conn = get_db()
    cur = conn.cursor()
    
    existing = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_number,)).fetchone()
    if existing:
        conn.close()
        return jsonify({'error': f"PO Number '{po_number}' already exists in the system."}), 400
        
    total_items = len(items)
    total_quantity = sum(int(it.get('ordered_quantity') or it.get('quantity') or 1) for it in items)
    total_po_value = sum(float(it.get('total_value') or (int(it.get('ordered_quantity') or it.get('quantity') or 1) * float(it.get('unit_price', 0.0)))) for it in items)
    
    is_packed = 1 if data.get('is_packed') else 0
    is_pickup_ready = 1 if data.get('is_pickup_ready') else 0
    is_dispatched = 1 if data.get('is_dispatched') or status in ['Partially Dispatched', 'Fully Dispatched', 'GRN Done'] else 0

    cur.execute('''
    INSERT INTO purchase_orders (
        po_number, po_date, platform_or_channel, channel_type, location, distributor_name,
        appointment_date, total_items, total_quantity, total_po_value, status,
        po_doc_url, po_doc_name, notes, created_by, is_packed, is_pickup_ready, is_dispatched
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        po_number, po_date, platform_or_channel, channel_type, location, distributor_name,
        appointment_date, total_items, total_quantity, round(total_po_value, 2), status,
        po_doc_url, po_doc_name, notes, user_email, is_packed, is_pickup_ready, is_dispatched
    ))
    po_id = cur.lastrowid
    
    is_gtmt = (platform_or_channel in ['GT', 'MT']) or (channel_type == 'gtmt')
    for it in items:
        sku = str(it.get('sku', '')).strip().upper()
        p_name = str(it.get('product_name', '')).strip()
        brand = str(it.get('brand', '')).strip()
        category = str(it.get('category', 'General')).strip()
        
        # Canonical product resolution: combines varied names (e.g. 'Moistuirseer', 'CICA Sunscreen') to master catalog
        c_sku, c_name, c_brand, c_cat = resolve_canonical_product(sku, p_name, None, cur)
        if c_sku != 'UNKNOWN':
            sku = c_sku
            p_name = c_name
            brand = c_brand
            category = c_cat

        qty = int(it.get('ordered_quantity') or it.get('quantity') or 1)
        tot_val = float(it.get('total_value') or 0.0)
        unit_price = float(it.get('unit_price') or (round(tot_val / qty, 2) if (qty and tot_val) else 0.0))
        if tot_val == 0.0 and unit_price > 0.0:
            tot_val = round(qty * unit_price, 2)
        item_notes = str(it.get('notes', '')).strip()
        
        cur.execute('''
        INSERT INTO po_items (
            po_id, po_number, sku, product_name, brand, category,
            ordered_quantity, unit_price, total_value, dispatched_quantity, notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
        ''', (po_id, po_number, sku, p_name, brand, category, qty, unit_price, round(tot_val, 2), item_notes))
        
        # Auto-route linked row into relevant dispatch log with actual_sent = 0 (Pending Booking)
        if is_gtmt:
            channel = 'GT' if 'GT' in platform_or_channel.upper() else 'MT'
            buyer = distributor_name or location or 'Distributor'
            cur.execute('''
            INSERT INTO dispatch_gt_mt (
                dispatch_date, brand, product_name, sku, category, channel, buyer_distributor, location,
                po_number, invoice_no, po_quantity, actual_sent, appointment_date,
                courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
                po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name, po_date, po_value
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, 0, ?, '', '', 0.0, 'Pending Booking', '', '', ?, ?, '', '', ?, ?)
            ''', (
                po_date or datetime.date.today().strftime('%Y-%m-%d'),
                brand, p_name, sku, category, channel, buyer, location,
                po_number, qty, appointment_date,
                po_doc_url, po_doc_name, po_date, round(tot_val, 2)
            ))
        else:
            cur.execute('''
            INSERT INTO dispatch_online (
                dispatch_date, brand, product_name, sku, category, platform, location,
                po_number, po_quantity, actual_sent, appointment_date,
                courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes,
                po_doc_url, po_doc_name, invoice_doc_url, invoice_doc_name, po_date, po_value
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, '', '', 0.0, 'Pending Booking', '', '', ?, ?, '', '', ?, ?)
            ''', (
                po_date or datetime.date.today().strftime('%Y-%m-%d'),
                brand, p_name, sku, category, platform_or_channel, location,
                po_number, qty, appointment_date,
                po_doc_url, po_doc_name, po_date, round(tot_val, 2)
            ))
        
    brands_in_po = list(set(it.get('brand', '') for it in items if it.get('brand')))
    brands_summary = ", ".join(brands_in_po[:3]) + ("..." if len(brands_in_po) > 3 else "")
    record_audit(
        conn, 'CREATE_PURCHASE_ORDER', 'Purchase Orders', po_number,
        f"Created PO {po_number} with {total_items} products ({total_quantity} units across {brands_summary}) for {platform_or_channel}"
    )
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'message': f"Purchase Order '{po_number}' created successfully.",
        'po_id': po_id,
        'po_number': po_number
    }), 201

@app.route('/api/purchase-orders/<int:po_id>', methods=['PUT'])
def update_purchase_order(po_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not (profile['is_admin'] or profile['can_edit_inventory'] or profile['can_edit_dispatch']):
        return jsonify({'error': 'Permission Denied: Unauthorized to edit Purchase Orders.'}), 403
        
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    if not po:
        conn.close()
        return jsonify({'error': 'Purchase order not found'}), 404
        
    po_number = str(data.get('po_number', po['po_number'])).strip()
    po_date = str(data.get('po_date', po['po_date'])).strip()
    platform_or_channel = str(data.get('platform_or_channel', po['platform_or_channel'])).strip()
    location = str(data.get('location', po['location'])).strip()
    appointment_date = str(data.get('appointment_date', po['appointment_date'])).strip()
    status = str(data.get('status', po['status'])).strip()
    notes = str(data.get('notes', po['notes'])).strip()
    
    items = data.get('items')
    if items is not None:
        if len(items) == 0:
            conn.close()
            return jsonify({'error': 'PO must have at least 1 item.'}), 400
        cur.execute('DELETE FROM po_items WHERE po_id = ?', (po_id,))
        total_items = len(items)
        total_quantity = sum(int(it.get('ordered_quantity', 1)) for it in items)
        total_po_value = sum(float(it.get('total_value') or (int(it.get('ordered_quantity', 1)) * float(it.get('unit_price', 0.0)))) for it in items)
        for it in items:
            c_sku, c_name, c_brand, c_cat = resolve_canonical_product(it.get('sku'), it.get('product_name'), None, cur)
            i_sku = c_sku if c_sku != 'UNKNOWN' else str(it.get('sku', '')).strip().upper()
            i_name = c_name if c_name != 'Unmapped Product' else it.get('product_name', '')
            i_brand = c_brand if c_brand != 'Other' else (it.get('brand', '') or '')
            i_cat = c_cat if c_cat != 'General' else (it.get('category', 'General') or 'General')

            cur.execute('''
            INSERT INTO po_items (
                po_id, po_number, sku, product_name, brand, category,
                ordered_quantity, unit_price, total_value, dispatched_quantity, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                po_id, po_number, i_sku, i_name,
                i_brand, i_cat, int(it.get('ordered_quantity', 1)),
                float(it.get('unit_price', 0.0)), round(float(it.get('total_value', 0.0)), 2),
                int(it.get('dispatched_quantity', 0)), it.get('notes', '')
            ))
    else:
        total_items = po['total_items']
        total_quantity = po['total_quantity']
        total_po_value = po['total_po_value']
        
    cur.execute('''
    UPDATE purchase_orders
    SET po_number = ?, po_date = ?, platform_or_channel = ?, location = ?,
        appointment_date = ?, status = ?, notes = ?, total_items = ?,
        total_quantity = ?, total_po_value = ?, updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    ''', (po_number, po_date, platform_or_channel, location, appointment_date, status, notes, total_items, total_quantity, total_po_value, po_id))
    
    record_audit(conn, 'UPDATE_PURCHASE_ORDER', 'Purchase Orders', po_number, f"Updated PO {po_number}")
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f"Purchase Order '{po_number}' updated successfully."})

@app.route('/api/purchase-orders/<int:po_id>', methods=['DELETE'])
def delete_purchase_order(po_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied: Only Admins and Inventory Leads can delete Purchase Orders.'}), 403
        
    conn = get_db()
    cur = conn.cursor()
    po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    if not po:
        conn.close()
        return jsonify({'error': 'Purchase order not found'}), 404
        
    po_num = po['po_number']
    # Clean up pending unfulfilled dispatch entries linked to this PO
    cur.execute("DELETE FROM dispatch_online WHERE po_number = ? AND (actual_sent = 0 OR dispatch_status = 'Pending Booking')", (po_num,))
    cur.execute("DELETE FROM dispatch_gt_mt WHERE po_number = ? AND (actual_sent = 0 OR dispatch_status = 'Pending Booking')", (po_num,))
    cur.execute('DELETE FROM po_items WHERE po_id = ?', (po_id,))
    cur.execute('DELETE FROM purchase_orders WHERE id = ?', (po_id,))
    record_audit(conn, 'DELETE_PURCHASE_ORDER', 'Purchase Orders', po_num, f"Deleted Purchase Order {po_num}")
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f"Purchase Order '{po_num}' deleted successfully."})

@app.route('/api/purchase-orders/<int:po_id>/grn', methods=['POST'])
def mark_purchase_order_grn(po_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    po = cur.execute('SELECT * FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    if not po:
        conn.close()
        return jsonify({'error': 'Purchase Order not found'}), 404
        
    po_platform = po['platform_or_channel']
    is_authorized = profile['is_admin'] or \
        (po_platform in profile['allowed_courier_platforms']) or \
        (po_platform in profile['allowed_courier_channels'])
        
    if not is_authorized:
        conn.close()
        return jsonify({
            'error': f"Permission Denied: Only the dedicated Platform Manager for '{po_platform}' (or Admin) can mark this PO as GRN Done."
        }), 403
        
    data = request.json or {}
    grn_number = str(data.get('grn_number', '')).strip()
    grn_date = str(data.get('grn_date', '')).strip() or datetime.date.today().strftime('%Y-%m-%d')
    grn_notes = str(data.get('grn_notes', '')).strip()
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    cur.execute('''
    UPDATE purchase_orders
    SET status = 'GRN Done',
        is_packed = 1,
        is_pickup_ready = 1,
        is_dispatched = 1,
        grn_number = ?,
        grn_date = ?,
        grn_notes = ?,
        grn_done_by = ?,
        grn_done_at = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    ''', (grn_number, grn_date, grn_notes, user_email, now_str, po_id))
    
    po_number = po['po_number']
    grn_tag = f" [GRN: {grn_number}]" if grn_number else " [GRN Done]"
    
    is_gtmt = (po['platform_or_channel'] in ['GT', 'MT']) or (po['channel_type'] == 'gtmt')
    if is_gtmt:
        cur.execute('''
        UPDATE dispatch_gt_mt
        SET dispatch_status = 'Delivered',
            shipping_notes = CASE 
                WHEN shipping_notes LIKE '%GRN%' THEN shipping_notes 
                WHEN shipping_notes = '' THEN ? 
                ELSE shipping_notes || ? 
            END
        WHERE po_number = ?
        ''', (f"GRN Done on {grn_date}{grn_tag}", f" | GRN Done on {grn_date}{grn_tag}", po_number))
    else:
        cur.execute('''
        UPDATE dispatch_online
        SET dispatch_status = 'Delivered',
            shipping_notes = CASE 
                WHEN shipping_notes LIKE '%GRN%' THEN shipping_notes 
                WHEN shipping_notes = '' THEN ? 
                ELSE shipping_notes || ? 
            END
        WHERE po_number = ?
        ''', (f"GRN Done on {grn_date}{grn_tag}", f" | GRN Done on {grn_date}{grn_tag}", po_number))
        
    record_audit(
        conn, 'MARK_GRN_DONE', 'Purchase Orders', po_number,
        f"Platform Manager {user_name} marked PO {po_number} as GRN Done (Ref: {grn_number or 'Confirmed'}, Date: {grn_date})"
    )
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'message': f"Purchase Order '{po_number}' successfully marked as GRN Done.",
        'po_id': po_id,
        'po_number': po_number,
        'grn_number': grn_number,
        'grn_date': grn_date
    })

@app.route('/api/purchase-orders/mark-grn', methods=['POST'])
def mark_purchase_order_grn_by_number():
    data = request.json or {}
    po_id = data.get('po_id')
    po_number = str(data.get('po_number', '')).strip()
    
    conn = get_db()
    cur = conn.cursor()
    if po_id:
        po = cur.execute('SELECT id FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
    elif po_number:
        po = cur.execute('SELECT id FROM purchase_orders WHERE po_number = ?', (po_number,)).fetchone()
    else:
        po = None
    conn.close()
    
    if not po:
        return jsonify({'error': 'Purchase Order not found'}), 404
        
    return mark_purchase_order_grn(po['id'])


# --- SALES DATA & SMART DIRECT PLATFORM UPLOAD ---
def auto_find_column(headers, candidate_keywords):
    """Finds matching column header from keywords with case-insensitive substring matching."""
    for idx, h in enumerate(headers):
        clean_h = str(h).lower().strip().replace('_', ' ').replace('-', ' ')
        for cand in candidate_keywords:
            if cand in clean_h:
                return idx
    return None

@app.route('/api/sales/upload', methods=['POST'])
def upload_sales_report():
    """Directly uploads and parses raw Excel or CSV downloaded from Blinkit, Zepto, Amazon, etc."""
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    platform = request.form.get('platform', '').strip()
    custom_date = request.form.get('sale_date', '').strip() or datetime.date.today().strftime('%Y-%m-%d')
    
    if not platform:
        return jsonify({'error': 'Platform selection is required'}), 400
        
    # Permission check: User must be Admin or assigned to this platform
    can_upload = profile['is_admin'] or (platform in profile['allowed_courier_platforms'])
    if not can_upload:
        return jsonify({'error': f"Permission Denied: You are not authorized to upload sales data for '{platform}'."}), 403
        
    filename = file.filename or 'sales_report.xlsx'
    ext = os.path.splitext(filename)[1].lower()
    
    conn = get_db()
    cur = conn.cursor()
    
    # Load all system products into memory for matching
    all_prods = cur.execute('SELECT sku, name, brand, category FROM products').fetchall()
    sku_lookup = {p['sku'].upper(): p for p in all_prods}
    name_lookup = {p['name'].lower().strip(): p for p in all_prods}
    
    rows_data = []
    if ext in ['.xlsx', '.xls']:
        try:
            wb = openpyxl.load_workbook(file.stream, data_only=True)
            ws = wb.active
            for r in ws.iter_rows(values_only=True):
                if any(r):
                    rows_data.append([str(c).strip() if c is not None else '' for c in r])
        except Exception as e:
            conn.close()
            return jsonify({'error': f'Failed to parse Excel file: {str(e)}'}), 400
    elif ext == '.csv':
        try:
            stream = TextIOWrapper(file.stream, encoding='utf-8-sig', errors='replace')
            reader = csv.reader(stream)
            for r in reader:
                if any(r):
                    rows_data.append([c.strip() for c in r])
        except Exception as e:
            conn.close()
            return jsonify({'error': f'Failed to parse CSV file: {str(e)}'}), 400
    else:
        conn.close()
        return jsonify({'error': 'Unsupported file format. Please upload .xlsx or .csv files.'}), 400
        
    if len(rows_data) < 2:
        conn.close()
        return jsonify({'error': 'The uploaded file is empty or contains only headers.'}), 400
        
    headers = rows_data[0]
    
    # Auto-detect column indexes
    sku_idx = auto_find_column(headers, ['sku', 'item code', 'itemcode', 'product code', 'seller sku', 'asin', 'fsn', 'code', 'barcode'])
    units_idx = auto_find_column(headers, ['units sold', 'units_sold', 'quantity', 'qty', 'units', 'sale qty', 'sold qty', 'volume', 'units ordered', 'qty sold', 'sales units'])
    date_idx = auto_find_column(headers, ['date', 'sale date', 'order date', 'day', 'transaction date'])
    name_idx = auto_find_column(headers, ['product name', 'item name', 'title', 'description', 'product', 'item title'])
    rev_idx = auto_find_column(headers, ['revenue', 'amount', 'gross sales', 'gmv', 'total value', 'net sales', 'total', 'sales value'])
    store_idx = auto_find_column(headers, ['store', 'dark store', 'city', 'hub', 'location', 'facility', 'warehouse'])
    
    if sku_idx is None and name_idx is None:
        conn.close()
        return jsonify({
            'error': f'Could not auto-detect SKU or Product Name column in headers: {headers[:10]}. Please ensure your file has an SKU or Product column.'
        }), 400
        
    if units_idx is None:
        conn.close()
        return jsonify({
            'error': f'Could not auto-detect Units Sold column in headers: {headers[:10]}.'
        }), 400
        
    imported_count = 0
    total_units = 0
    total_rev = 0.0
    unmatched_rows = []
    
    for row_num, row in enumerate(rows_data[1:], 2):
        if not any(row):
            continue
            
        raw_sku = row[sku_idx].strip() if (sku_idx is not None and sku_idx < len(row)) else ''
        raw_name = row[name_idx].strip() if (name_idx is not None and name_idx < len(row)) else ''
        raw_units = row[units_idx].strip() if (units_idx < len(row)) else '0'
        
        raw_date = custom_date
        if date_idx is not None and date_idx < len(row) and row[date_idx].strip():
            d_val = row[date_idx].strip()
            # Normalize dates if possible
            if len(d_val) >= 10 and '-' in d_val[:10]:
                raw_date = d_val[:10]
            elif '/' in d_val:
                parts = d_val.split(' ')[0].split('/')
                if len(parts) == 3:
                    if len(parts[2]) == 4:
                        raw_date = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
                    elif len(parts[0]) == 4:
                        raw_date = f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
                        
        raw_rev = 0.0
        if rev_idx is not None and rev_idx < len(row) and row[rev_idx].strip():
            clean_rev = row[rev_idx].replace('₹', '').replace(',', '').replace('Rs', '').strip()
            try:
                raw_rev = float(clean_rev)
            except:
                raw_rev = 0.0
                
        raw_store = row[store_idx].strip() if (store_idx is not None and store_idx < len(row)) else ''
        
        try:
            clean_units = int(float(raw_units.replace(',', '').strip()))
        except:
            clean_units = 0
            
        if clean_units <= 0:
            continue
            
        # Resolve against our canonical catalog and synonym engine
        c_sku, c_name, c_brand, _ = resolve_canonical_product(raw_sku, raw_name, None, cur)
        if c_sku != 'UNKNOWN':
            final_sku = c_sku
            final_name = c_name
            final_brand = c_brand
        else:
            final_sku = raw_sku or 'UNKNOWN'
            final_name = raw_name or 'Unmapped Product'
            final_brand = 'Other'
            unmatched_rows.append(f"Row {row_num}: {raw_sku} / {raw_name}")
            
        cur.execute('''
        INSERT INTO sales_data (
            sale_date, platform, sku, product_name, brand, units_sold, revenue, store_location, source_file, uploaded_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            raw_date, platform, final_sku, final_name, final_brand, clean_units, raw_rev, raw_store, filename, f"{user_name} ({user_email})"
        ))
        imported_count += 1
        total_units += clean_units
        total_rev += raw_rev
        
    edit_sum = f"Uploaded {filename} for {platform}: {imported_count} sales entries, {total_units:,} units sold (Rev: ₹{total_rev:,.2f})"
    record_audit(conn, 'UPLOAD_SALES_DATA', 'Sales & Pipeline', f'{platform} | {filename}', edit_sum)
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'platform': platform,
        'filename': filename,
        'imported_records': imported_count,
        'total_units_sold': total_units,
        'total_revenue': total_rev,
        'unmatched_count': len(unmatched_rows),
        'unmatched_samples': unmatched_rows[:5]
    })

@app.route('/api/sales', methods=['GET'])
def get_sales_records():
    brand = request.args.get('brand')
    platform = request.args.get('platform')
    search = request.args.get('search')
    limit = int(request.args.get('limit', 150))
    
    conn = get_db()
    cur = conn.cursor()
    
    query = 'SELECT * FROM sales_data WHERE 1=1'
    params = []
    if brand and brand != 'All':
        query += ' AND brand = ?'
        params.append(brand)
    if platform and platform != 'All':
        query += ' AND platform = ?'
        params.append(platform)
    if search:
        query += ' AND (product_name LIKE ? OR sku LIKE ? OR store_location LIKE ? OR source_file LIKE ?)'
        term = f'%{search}%'
        params.extend([term, term, term, term])
        
    query += ' ORDER BY sale_date DESC, id DESC LIMIT ?'
    params.append(limit)
    rows = cur.execute(query, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/sales/<int:item_id>', methods=['DELETE'])
def delete_sales_record(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM sales_data WHERE id = ?', (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Record not found'}), 404
        
    platform = row['platform']
    can_delete = profile['is_admin'] or (platform in profile['allowed_courier_platforms'])
    if not can_delete:
        conn.close()
        return jsonify({'error': 'Permission Denied: You cannot delete sales records for this platform.'}), 403
        
    edit_sum = f"Deleted sales entry #{item_id}: {row['units_sold']} units of {row['sku']} on {platform}"
    record_audit(conn, 'DELETE_SALES_RECORD', 'Sales & Pipeline', f'Sale #{item_id}', edit_sum)
    cur.execute('DELETE FROM sales_data WHERE id = ?', (item_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/sales/dashboard-analytics', methods=['GET'])
def get_sales_dashboard_analytics():
    """Returns aggregated metrics, platform breakdown, top products, top locations, and filtered sales records."""
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    platform = request.args.get('platform', '').strip()
    sku = request.args.get('sku', '').strip()
    location = request.args.get('location', '').strip()
    brand = request.args.get('brand', '').strip()
    search = request.args.get('search', '').strip()
    try:
        limit = int(request.args.get('limit', 200))
    except (ValueError, TypeError):
        limit = 200
    
    conn = get_db()
    cur = conn.cursor()
    
    # 1. Fetch available filter options from current dataset
    platforms_rows = cur.execute('''
        SELECT DISTINCT platform FROM sales_data 
        WHERE platform IS NOT NULL AND platform != '' 
        ORDER BY platform
    ''').fetchall()
    available_platforms = [r['platform'] for r in platforms_rows]
    
    products_rows = cur.execute('''
        SELECT DISTINCT sku, product_name, brand FROM sales_data 
        WHERE sku IS NOT NULL AND sku != '' 
        ORDER BY product_name
    ''').fetchall()
    available_products = [{'sku': r['sku'], 'name': r['product_name'] or r['sku'], 'brand': r['brand'] or ''} for r in products_rows]
    
    locations_rows = cur.execute('''
        SELECT DISTINCT store_location FROM sales_data 
        WHERE store_location IS NOT NULL AND store_location != '' 
        ORDER BY store_location
    ''').fetchall()
    available_locations = [r['store_location'] for r in locations_rows]
    
    date_bounds = cur.execute('''
        SELECT MIN(sale_date) as min_date, MAX(sale_date) as max_date 
        FROM sales_data WHERE sale_date IS NOT NULL AND sale_date != ''
    ''').fetchone()
    
    # 2. Build filtered WHERE clause
    where_clauses = ["1=1"]
    params = []
    
    if start_date:
        where_clauses.append("sale_date >= ?")
        params.append(start_date)
    if end_date:
        where_clauses.append("sale_date <= ?")
        params.append(end_date)
    if platform and platform != 'All':
        where_clauses.append("platform = ?")
        params.append(platform)
    if sku and sku != 'All':
        where_clauses.append("sku = ?")
        params.append(sku)
    if location and location != 'All':
        where_clauses.append("store_location = ?")
        params.append(location)
    if brand and brand != 'All':
        where_clauses.append("brand = ?")
        params.append(brand)
    if search:
        term = f"%{search}%"
        where_clauses.append("(sku LIKE ? OR product_name LIKE ? OR store_location LIKE ? OR platform LIKE ?)")
        params.extend([term, term, term, term])
        
    where_sql = " AND ".join(where_clauses)
    
    # 3. Overall KPIs
    kpi_row = cur.execute(f'''
        SELECT 
            COUNT(*) as total_records,
            COALESCE(SUM(units_sold), 0) as total_units,
            COALESCE(SUM(revenue), 0) as total_revenue,
            COUNT(DISTINCT platform) as active_platforms,
            COUNT(DISTINCT store_location) as active_locations,
            COUNT(DISTINCT sku) as active_products
        FROM sales_data
        WHERE {where_sql}
    ''', params).fetchone()
    
    total_units = kpi_row['total_units'] or 0
    total_revenue = round(kpi_row['total_revenue'] or 0.0, 2)
    avg_order_val = round(total_revenue / total_units, 2) if total_units > 0 else 0.0
    
    kpis = {
        'total_records': kpi_row['total_records'] or 0,
        'total_units': total_units,
        'total_revenue': total_revenue,
        'avg_unit_value': avg_order_val,
        'active_platforms': kpi_row['active_platforms'] or 0,
        'active_locations': kpi_row['active_locations'] or 0,
        'active_products': kpi_row['active_products'] or 0
    }
    
    # 4. Platform Breakdown
    plat_rows = cur.execute(f'''
        SELECT 
            platform,
            COALESCE(SUM(units_sold), 0) as units,
            COALESCE(SUM(revenue), 0) as revenue,
            COUNT(DISTINCT store_location) as locations_count,
            COUNT(DISTINCT sku) as products_count
        FROM sales_data
        WHERE {where_sql}
        GROUP BY platform
        ORDER BY units DESC, revenue DESC
    ''', params).fetchall()
    
    platforms_breakdown = []
    for r in plat_rows:
        u = r['units'] or 0
        rev = round(r['revenue'] or 0.0, 2)
        share_units = round((u / total_units * 100), 1) if total_units > 0 else 0.0
        share_rev = round((rev / total_revenue * 100), 1) if total_revenue > 0 else 0.0
        platforms_breakdown.append({
            'platform': r['platform'],
            'units': u,
            'revenue': rev,
            'locations_count': r['locations_count'],
            'products_count': r['products_count'],
            'unit_share_pct': share_units,
            'revenue_share_pct': share_rev
        })
        
    # 5. Top Products
    top_prod_rows = cur.execute(f'''
        SELECT 
            sku,
            COALESCE(product_name, sku) as product_name,
            COALESCE(brand, '') as brand,
            COALESCE(SUM(units_sold), 0) as units,
            COALESCE(SUM(revenue), 0) as revenue
        FROM sales_data
        WHERE {where_sql}
        GROUP BY sku
        ORDER BY units DESC, revenue DESC
        LIMIT 10
    ''', params).fetchall()
    
    top_products = [{
        'sku': r['sku'],
        'product_name': r['product_name'],
        'brand': r['brand'],
        'units': r['units'],
        'revenue': round(r['revenue'] or 0.0, 2),
        'share_pct': round((r['units'] / total_units * 100), 1) if total_units > 0 else 0.0
    } for r in top_prod_rows]
    
    # 6. Top Locations
    top_loc_rows = cur.execute(f'''
        SELECT 
            store_location as location,
            COALESCE(SUM(units_sold), 0) as units,
            COALESCE(SUM(revenue), 0) as revenue,
            COUNT(DISTINCT platform) as platforms_count
        FROM sales_data
        WHERE {where_sql} AND store_location IS NOT NULL AND store_location != ''
        GROUP BY store_location
        ORDER BY units DESC, revenue DESC
        LIMIT 10
    ''', params).fetchall()
    
    top_locations = [{
        'location': r['location'],
        'units': r['units'],
        'revenue': round(r['revenue'] or 0.0, 2),
        'platforms_count': r['platforms_count'],
        'share_pct': round((r['units'] / total_units * 100), 1) if total_units > 0 else 0.0
    } for r in top_loc_rows]
    
    # 7. Timeline (daily trend)
    timeline_rows = cur.execute(f'''
        SELECT 
            sale_date,
            COALESCE(SUM(units_sold), 0) as units,
            COALESCE(SUM(revenue), 0) as revenue
        FROM sales_data
        WHERE {where_sql} AND sale_date IS NOT NULL AND sale_date != ''
        GROUP BY sale_date
        ORDER BY sale_date ASC
        LIMIT 60
    ''', params).fetchall()
    
    timeline = [{
        'date': r['sale_date'],
        'units': r['units'],
        'revenue': round(r['revenue'] or 0.0, 2)
    } for r in timeline_rows]
    
    # 8. Filtered Sales Records
    records_params = list(params)
    records_params.append(limit)
    record_rows = cur.execute(f'''
        SELECT 
            id, sale_date, platform, sku, product_name, brand, 
            units_sold, revenue, store_location, source_file, created_at
        FROM sales_data
        WHERE {where_sql}
        ORDER BY sale_date DESC, id DESC
        LIMIT ?
    ''', records_params).fetchall()
    
    records = [dict(r) for r in record_rows]
    conn.close()
    
    return jsonify({
        'status': 'success',
        'kpis': kpis,
        'platforms': platforms_breakdown,
        'top_products': top_products,
        'top_locations': top_locations,
        'timeline': timeline,
        'records': records,
        'filter_options': {
            'platforms': available_platforms,
            'products': available_products,
            'locations': available_locations,
            'min_date': date_bounds['min_date'] if date_bounds else '',
            'max_date': date_bounds['max_date'] if date_bounds else ''
        }
    })

@app.route('/api/sales/export-csv', methods=['GET'])
def export_sales_csv():
    """Generates and streams a CSV of filtered sales data."""
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    platform = request.args.get('platform', '').strip()
    sku = request.args.get('sku', '').strip()
    location = request.args.get('location', '').strip()
    brand = request.args.get('brand', '').strip()
    search = request.args.get('search', '').strip()
    
    conn = get_db()
    cur = conn.cursor()
    
    where_clauses = ["1=1"]
    params = []
    
    if start_date:
        where_clauses.append("sale_date >= ?")
        params.append(start_date)
    if end_date:
        where_clauses.append("sale_date <= ?")
        params.append(end_date)
    if platform and platform != 'All':
        where_clauses.append("platform = ?")
        params.append(platform)
    if sku and sku != 'All':
        where_clauses.append("sku = ?")
        params.append(sku)
    if location and location != 'All':
        where_clauses.append("store_location = ?")
        params.append(location)
    if brand and brand != 'All':
        where_clauses.append("brand = ?")
        params.append(brand)
    if search:
        term = f"%{search}%"
        where_clauses.append("(sku LIKE ? OR product_name LIKE ? OR store_location LIKE ? OR platform LIKE ?)")
        params.extend([term, term, term, term])
        
    where_sql = " AND ".join(where_clauses)
    
    rows = cur.execute(f'''
        SELECT sale_date, platform, sku, product_name, brand, store_location, units_sold, revenue, source_file
        FROM sales_data
        WHERE {where_sql}
        ORDER BY sale_date DESC, id DESC
    ''', params).fetchall()
    conn.close()
    
    import io, csv
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Sale Date', 'Platform', 'SKU', 'Product Name', 'Brand', 'Location / City', 'Units Sold', 'Revenue (INR)', 'Source File'])
    
    for r in rows:
        rev_val = f"{r['revenue']:.2f}" if r['revenue'] is not None else '0.00'
        writer.writerow([
            r['sale_date'] or '',
            r['platform'] or '',
            r['sku'] or '',
            r['product_name'] or '',
            r['brand'] or '',
            r['store_location'] or '',
            r['units_sold'] or 0,
            rev_val,
            r['source_file'] or ''
        ])
        
    output.seek(0)
    today_str = datetime.date.today().strftime('%Y%m%d')
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename=Sales_Dashboard_Export_{today_str}.csv"}
    )

@app.route('/api/sales/pipeline-summary', methods=['GET'])
def get_pipeline_summary():
    """Computes SKU-level Dark Store Pipeline Stock, Daily Run-Rate, and Days of Inventory (DOI)."""
    brand = request.args.get('brand')
    conn = get_db()
    cur = conn.cursor()
    
    master_rows = get_master_data_internal(cur, brand)
    
    # Calculate daily run rate based on recent 30-day sales
    thirty_days_ago = (datetime.date.today() - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    sales_30d = cur.execute('''
        SELECT sku, SUM(units_sold) as sold_30d 
        FROM sales_data 
        WHERE sale_date >= ?
        GROUP BY sku
    ''', (thirty_days_ago,)).fetchall()
    sales_30d_map = {r['sku']: r['sold_30d'] for r in sales_30d}
    
    results = []
    total_pipeline_units = 0
    total_sales_units = 0
    critical_stockout_count = 0
    
    for r in master_rows:
        sku = r['sku']
        dispatched = r['online_total']
        sold = r.get('total_sold', 0)
        pipeline_stock = max(0, dispatched - sold)
        total_pipeline_units += pipeline_stock
        total_sales_units += sold
        
        sold_recent = sales_30d_map.get(sku, 0)
        daily_velocity = round(sold_recent / 30.0, 1) if sold_recent > 0 else 0.0
        
        rem_stock = r['remaining_inventory']
        if daily_velocity > 0:
            days_of_inventory = round(rem_stock / daily_velocity, 0)
        else:
            days_of_inventory = 999
            
        if days_of_inventory <= 7 and rem_stock > 0:
            doi_status = 'Critical Alert (<7 Days)'
            critical_stockout_count += 1
        elif days_of_inventory <= 15:
            doi_status = 'Reorder Soon (7-15 Days)'
        elif rem_stock <= 0:
            doi_status = 'Out of Stock'
            critical_stockout_count += 1
        else:
            doi_status = 'Healthy (>15 Days)'
            
        results.append({
            'sku': sku,
            'name': r['product_name'],
            'brand': r['brand'],
            'category': r['category'],
            'total_inventory': r['total_inventory'],
            'remaining_inventory': rem_stock,
            'online_dispatched': dispatched,
            'units_sold': sold,
            'pipeline_stock': pipeline_stock,
            'daily_velocity': daily_velocity,
            'days_of_inventory': days_of_inventory if days_of_inventory < 999 else '999+',
            'doi_status': doi_status
        })
        
    conn.close()
    return jsonify({
        'summary': {
            'total_pipeline_units': total_pipeline_units,
            'total_sales_units': total_sales_units,
            'critical_stockout_count': critical_stockout_count,
            'total_skus': len(results)
        },
        'skus': results
    })

# --- ADJUSTMENTS ---
@app.route('/api/adjustments', methods=['GET'])
def get_adjustments():
    brand = request.args.get('brand')
    search = request.args.get('search')
    
    conn = get_db()
    cur = conn.cursor()
    
    query = 'SELECT * FROM adjustments WHERE 1=1'
    params = []
    if brand and brand != 'All':
        query += ' AND brand = ?'
        params.append(brand)
    if search:
        query += ' AND (product_name LIKE ? OR reason LIKE ? OR adjustment_type LIKE ? OR sku LIKE ?)'
        term = f'%{search}%'
        params.extend([term, term, term, term])
        
    query += ' ORDER BY id DESC'
    rows = cur.execute(query, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/adjustments', methods=['POST'])
def add_adjustment():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied: Only Inventory Lead or Admin can record adjustments.'}), 403
        
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    brand = data.get('brand', '').strip()
    category = data.get('category', '').strip()
    channel = data.get('channel', '').strip()
    adj_type = data.get('adjustment_type', 'Sample').strip()
    reason = data.get('reason', '').strip()
    adj_date = data.get('adjustment_date') or datetime.date.today().strftime('%Y-%m-%d')
    
    try:
        qty = int(data.get('quantity', 0))
    except:
        qty = 0
        
    if sku:
        prod = cur.execute('SELECT name, brand, category FROM products WHERE sku = ?', (sku,)).fetchone()
        if prod:
            product_name = prod['name']
            brand = prod['brand']
            category = prod['category']
    elif product_name:
        prod = cur.execute('SELECT sku, brand, category FROM products WHERE LOWER(name) = LOWER(?)', (product_name,)).fetchone()
        if prod:
            sku = prod['sku']
            brand = prod['brand']
            category = prod['category']
            
    cur.execute('''
    INSERT INTO adjustments (
        adjustment_date, brand, channel, product_name, sku, category, adjustment_type, quantity, reason
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        adj_date, brand, channel, product_name, sku, category, adj_type, qty, reason
    ))
    new_id = cur.lastrowid
    
    edit_sum = f"Deducted {qty} units of {product_name} (Type: {adj_type}, Reason: {reason})"
    record_audit(conn, 'ADD_ADJUSTMENT', 'Adjustments', f'{sku} | {adj_type}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'id': new_id})

@app.route('/api/adjustments/<int:item_id>', methods=['DELETE'])
def delete_adjustment(item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied'}), 403
        
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute('SELECT * FROM adjustments WHERE id = ?', (item_id,)).fetchone()
    if row:
        edit_sum = f"Deleted adjustment #{item_id}: Restored {row['quantity']} units of {row['product_name']}"
        record_audit(conn, 'DELETE_ADJUSTMENT', 'Adjustments', f'Adjustment #{item_id}', edit_sum)
        cur.execute('DELETE FROM adjustments WHERE id = ?', (item_id,))
        conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# --- INVENTORY STOCK & THRESHOLDS ---
@app.route('/api/inventory-stock', methods=['GET'])
def get_inventory_stock():
    brand = request.args.get('brand', 'All').strip()
    sku = request.args.get('sku', 'All').strip()
    search = request.args.get('search', '').strip()
    
    conn = get_db()
    cur = conn.cursor()
    query = '''
    SELECT p.sku, p.name as product_name, p.brand, p.category, 
           COALESCE(s.total_inventory, 0) as total_inventory,
           COALESCE(s.reorder_threshold, 0) as reorder_threshold
    FROM products p
    LEFT JOIN inventory_stock s ON p.sku = s.sku
    WHERE 1=1
    '''
    params = []
    if brand and brand != 'All':
        query += ' AND p.brand = ?'
        params.append(brand)
    if sku and sku != 'All':
        query += ' AND p.sku = ?'
        params.append(sku)
    if search:
        term = f'%{search}%'
        query += ' AND (p.sku LIKE ? OR p.name LIKE ? OR p.category LIKE ?)'
        params.extend([term, term, term])
        
    query += ' ORDER BY p.brand, p.name'
    rows = cur.execute(query, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/inventory-stock', methods=['POST'])
def update_inventory_stock():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if not profile['can_edit_inventory']:
        return jsonify({
            'error': f"Permission Denied: Master Sheet Baseline Inventory and Reorder Thresholds are editable exclusively by the Inventory Lead ({profile['roles'].get('INVENTORY_MANAGER')})."
        }), 403
        
    data = request.json or {}
    sku = data.get('sku')
    total_inv = int(data.get('total_inventory', 0))
    threshold = int(data.get('reorder_threshold', 0))
    
    conn = get_db()
    cur = conn.cursor()
    old_row = cur.execute('SELECT total_inventory, reorder_threshold FROM inventory_stock WHERE sku = ?', (sku,)).fetchone()
    old_inv = old_row['total_inventory'] if old_row else 0
    old_thresh = old_row['reorder_threshold'] if old_row else 0
    
    cur.execute('''
    INSERT INTO inventory_stock (sku, total_inventory, reorder_threshold)
    VALUES (?, ?, ?)
    ON CONFLICT(sku) DO UPDATE SET
        total_inventory = excluded.total_inventory,
        reorder_threshold = excluded.reorder_threshold
    ''', (sku, total_inv, threshold))
    
    edit_sum = f"Updated Stock: {old_inv} -> {total_inv} | Threshold: {old_thresh} -> {threshold}"
    record_audit(conn, 'EDIT_STOCK', 'Inventory Stock', f'SKU {sku}', edit_sum)
    
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# --- AUDIT LOG ---
@app.route('/api/audit-log')
def get_audit_log():
    search = request.args.get('search', '')
    sheet = request.args.get('sheet', 'All')
    
    conn = get_db()
    cur = conn.cursor()
    
    query = 'SELECT * FROM audit_log WHERE 1=1'
    params = []
    if sheet and sheet != 'All':
        query += ' AND sheet_name = ?'
        params.append(sheet)
    if search:
        query += ' AND (user_name LIKE ? OR user_email LIKE ? OR edit_summary LIKE ? OR entity_ref LIKE ?)'
        term = f'%{search}%'
        params.extend([term, term, term, term])
        
    query += ' ORDER BY id DESC LIMIT 150'
    rows = cur.execute(query, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

# --- COURIER COSTING & LOGISTICS EXPENSES ---
def get_courier_costing_rows_internal(channel=None, search=None, start_date=None, end_date=None):
    conn = get_db()
    cur = conn.cursor()
    
    online_rows = cur.execute('''
        SELECT id, 'online' as dispatch_type, platform as channel, location,
               po_date, appointment_date, po_number, tracking_id, courier_name,
               COALESCE(actual_sent, po_quantity, 0) as sum_of_qty,
               COALESCE(po_value, 0.0) as po_value,
               COALESCE(courier_charges, 0.0) as courier_charges,
               pickup_date,
               COALESCE(actual_weight, 0.0) as actual_weight,
               COALESCE(charged_weight, 0.0) as charged_weight,
               dispatch_status, sku, product_name, brand, dispatch_date
        FROM dispatch_online
    ''').fetchall()
    
    gtmt_rows = cur.execute('''
        SELECT id, 'gt_mt' as dispatch_type, channel, location,
               po_date, appointment_date, po_number, tracking_id, courier_name,
               COALESCE(actual_sent, po_quantity, 0) as sum_of_qty,
               COALESCE(po_value, 0.0) as po_value,
               COALESCE(courier_charges, 0.0) as courier_charges,
               pickup_date,
               COALESCE(actual_weight, 0.0) as actual_weight,
               COALESCE(charged_weight, 0.0) as charged_weight,
               dispatch_status, sku, product_name, brand, dispatch_date
        FROM dispatch_gt_mt
    ''').fetchall()
    conn.close()
    
    all_records = []
    for r in list(online_rows) + list(gtmt_rows):
        d = dict(r)
        d['po_value'] = float(d['po_value'] or 0.0)
        d['courier_charges'] = float(d['courier_charges'] or 0.0)
        d['actual_weight'] = float(d['actual_weight'] or 0.0)
        d['charged_weight'] = float(d['charged_weight'] or 0.0)
        d['sum_of_qty'] = int(d['sum_of_qty'] or 0)
        
        # Calculate percentage: (courier_charges / po_value) * 100
        if d['po_value'] > 0:
            d['percentage'] = round((d['courier_charges'] / d['po_value']) * 100, 2)
        else:
            d['percentage'] = 0.0
            
        d['weight_variance'] = round(d['charged_weight'] - d['actual_weight'], 2)
        
        # Filter Channel
        if channel and channel != 'All':
            if d['channel'].lower().strip() != channel.lower().strip():
                continue
                
        # Filter Search
        if search:
            s_term = search.lower().strip()
            haystack = f"{d.get('po_number','')} {d.get('tracking_id','')} {d.get('location','')} {d.get('courier_name','')} {d.get('sku','')} {d.get('product_name','')} {d.get('channel','')}".lower()
            if s_term not in haystack:
                continue
                
        # Date filters
        rec_date = d.get('po_date') or d.get('dispatch_date') or d.get('appointment_date') or ''
        if start_date and rec_date and rec_date < start_date:
            continue
        if end_date and rec_date and rec_date > end_date:
            continue
            
        all_records.append(d)
        
    all_records.sort(key=lambda x: (x.get('po_date') or x.get('dispatch_date') or '', x.get('id', 0)), reverse=True)
    
    total_freight = sum(r['courier_charges'] for r in all_records)
    total_val = sum(r['po_value'] for r in all_records)
    total_act_wt = sum(r['actual_weight'] for r in all_records)
    total_chg_wt = sum(r['charged_weight'] for r in all_records)
    overall_pct = round((total_freight / total_val * 100), 2) if total_val > 0 else 0.0
    
    summary = {
        'total_freight_spent': round(total_freight, 2),
        'total_po_value': round(total_val, 2),
        'overall_freight_percentage': overall_pct,
        'total_actual_weight': round(total_act_wt, 2),
        'total_charged_weight': round(total_chg_wt, 2),
        'weight_discrepancy': round(total_chg_wt - total_act_wt, 2),
        'total_shipments': len(all_records)
    }
    
    return all_records, summary

@app.route('/api/courier-costing', methods=['GET'])
def get_courier_costing():
    channel = request.args.get('channel', 'All')
    search = request.args.get('search', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    rows, summary = get_courier_costing_rows_internal(channel, search, start_date, end_date)
    return jsonify({
        'rows': rows,
        'summary': summary
    })

@app.route('/api/courier-costing/<dispatch_type>/<int:item_id>', methods=['POST'])
def update_courier_costing_item(dispatch_type, item_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    table = 'dispatch_online' if dispatch_type == 'online' else 'dispatch_gt_mt'
    
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute(f"SELECT * FROM {table} WHERE id = ?", (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Dispatch entry not found'}), 404
        
    channel_name = row['platform'] if dispatch_type == 'online' else row['channel']
    
    can_edit = profile['is_admin'] or profile['can_edit_dispatch'] or (channel_name in profile['allowed_courier_platforms']) or (channel_name in profile['allowed_courier_channels'])
    if not can_edit:
        conn.close()
        return jsonify({'error': f"Permission Denied: You cannot edit courier costing for {channel_name}."}), 403
        
    data = request.json or request.form
    po_date = data.get('po_date', row['po_date'] or '').strip()
    appointment_date = data.get('appointment_date', row['appointment_date'] or '').strip()
    po_number = data.get('po_number', row['po_number'] or '').strip()
    tracking_id = data.get('tracking_id', row['tracking_id'] or '').strip()
    courier_name = data.get('courier_name', row['courier_name'] or '').strip()
    pickup_date = data.get('pickup_date', row['pickup_date'] or '').strip()
    
    try:
        po_value = float(data.get('po_value', row['po_value'] or 0.0))
    except:
        po_value = row['po_value'] or 0.0
        
    try:
        courier_charges = float(data.get('courier_charges', row['courier_charges'] or 0.0))
    except:
        courier_charges = row['courier_charges'] or 0.0
        
    try:
        actual_weight = float(data.get('actual_weight', row['actual_weight'] or 0.0))
    except:
        actual_weight = row['actual_weight'] or 0.0
        
    try:
        charged_weight = float(data.get('charged_weight', row['charged_weight'] or 0.0))
    except:
        charged_weight = row['charged_weight'] or 0.0
        
    dispatch_status = data.get('dispatch_status', row['dispatch_status'] or 'Dispatched').strip()
    
    cur.execute(f'''
        UPDATE {table}
        SET po_date = ?, appointment_date = ?, po_number = ?, tracking_id = ?,
            courier_name = ?, courier_charges = ?, po_value = ?, pickup_date = ?,
            actual_weight = ?, charged_weight = ?, dispatch_status = ?
        WHERE id = ?
    ''', (po_date, appointment_date, po_number, tracking_id, courier_name,
          courier_charges, po_value, pickup_date, actual_weight, charged_weight, dispatch_status, item_id))
          
    record_audit(
        conn,
        'UPDATE_COURIER_COSTING',
        'Courier Costing',
        f"Dispatch #{item_id} | {channel_name} | {po_number}",
        f"Updated Costing: PO Val ₹{po_value:,.2f}, Freight ₹{courier_charges:,.2f}, Act Wt {actual_weight}kg, Chg Wt {charged_weight}kg, LR {tracking_id}"
    )
    conn.commit()
    conn.close()
    
    return jsonify({'status': 'success', 'message': 'Courier costing details updated successfully'})

# --- DATE-WISE & BATCH-WISE INVENTORY INWARDING ---
@app.route('/api/inventory/batches', methods=['GET'])
def get_inventory_batches():
    sku = request.args.get('sku', '').strip()
    brand = request.args.get('brand', '').strip()
    search = request.args.get('search', '').strip()
    
    conn = get_db()
    cur = conn.cursor()
    query = 'SELECT * FROM stock_inward_batches WHERE 1=1'
    params = []
    
    if sku and sku != 'All':
        query += ' AND sku = ?'
        params.append(sku)
    if brand and brand != 'All':
        query += ' AND brand = ?'
        params.append(brand)
    if search:
        query += ' AND (batch_no LIKE ? OR sku LIKE ? OR product_name LIKE ? OR supplier_po_ref LIKE ? OR notes LIKE ?)'
        term = f'%{search}%'
        params.extend([term, term, term, term, term])
        
    query += ' ORDER BY inward_date DESC, id DESC'
    batches = [dict(r) for r in cur.execute(query, params).fetchall()]
    conn.close()
    return jsonify(batches)

@app.route('/api/inventory/batches', methods=['POST'])
def add_inventory_batch():
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied: Only Inventory Managers or Admins can inward stock batches.'}), 403
        
    data = request.json or request.form
    inward_date = data.get('inward_date', '').strip() or datetime.date.today().strftime('%Y-%m-%d')
    batch_no = data.get('batch_no', '').strip()
    sku = data.get('sku', '').strip().upper()
    product_name = data.get('product_name', '').strip()
    brand = data.get('brand', '').strip()
    
    try:
        quantity_added = int(data.get('quantity_added', 0))
    except:
        quantity_added = 0
        
    if not batch_no:
        return jsonify({'error': 'Batch Number / Lot # is required.'}), 400
    if not sku:
        return jsonify({'error': 'SKU is required.'}), 400
    if quantity_added <= 0:
        return jsonify({'error': 'Quantity added must be a positive integer.'}), 400
        
    mfg_date = data.get('mfg_date', '').strip()
    expiry_date = data.get('expiry_date', '').strip()
    supplier_po_ref = data.get('supplier_po_ref', '').strip()
    notes = data.get('notes', '').strip()
    
    conn = get_db()
    cur = conn.cursor()
    
    prod_row = cur.execute('SELECT name, brand FROM products WHERE sku = ?', (sku,)).fetchone()
    if prod_row:
        if not product_name:
            product_name = prod_row['name']
        if not brand:
            brand = prod_row['brand']
    else:
        if not product_name:
            product_name = f"Product {sku}"
        if not brand:
            brand = "Brand A"
        cur.execute('INSERT OR IGNORE INTO products (sku, brand, name, category) VALUES (?, ?, ?, ?)',
                    (sku, brand, product_name, 'General'))
                    
    cur.execute('''
        INSERT INTO stock_inward_batches 
        (inward_date, batch_no, sku, product_name, brand, quantity_added, mfg_date, expiry_date, supplier_po_ref, notes, added_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (inward_date, batch_no, sku, product_name, brand, quantity_added, mfg_date, expiry_date, supplier_po_ref, notes, user_email))
    new_batch_id = cur.lastrowid
    
    stk_row = cur.execute('SELECT total_inventory FROM inventory_stock WHERE sku = ?', (sku,)).fetchone()
    if stk_row:
        cur.execute('UPDATE inventory_stock SET total_inventory = total_inventory + ? WHERE sku = ?', (quantity_added, sku))
    else:
        cur.execute('INSERT INTO inventory_stock (sku, total_inventory, reorder_threshold) VALUES (?, ?, 20)', (sku, quantity_added))
        
    record_audit(
        conn,
        'INWARD_STOCK_BATCH',
        'Inventory Stock & Batch Log',
        f"Batch {batch_no} | {sku}",
        f"Inwarded {quantity_added} units on {inward_date} (Mfg: {mfg_date}, Exp: {expiry_date}, Supplier Ref: {supplier_po_ref})"
    )
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'batch_id': new_batch_id,
        'message': f"Successfully inwarded {quantity_added} units for SKU '{sku}' under batch '{batch_no}'."
    })

@app.route('/api/inventory/batches/<int:batch_id>', methods=['DELETE'])
def delete_inventory_batch(batch_id):
    user_email, user_name = get_user_info()
    profile = get_user_profile(user_email)
    
    if not (profile['is_admin'] or profile['can_edit_inventory']):
        return jsonify({'error': 'Permission Denied: Only Inventory Managers or Admins can delete stock batches.'}), 403
        
    conn = get_db()
    cur = conn.cursor()
    batch = cur.execute('SELECT * FROM stock_inward_batches WHERE id = ?', (batch_id,)).fetchone()
    if not batch:
        conn.close()
        return jsonify({'error': 'Batch not found'}), 404
        
    sku = batch['sku']
    qty = batch['quantity_added']
    batch_no = batch['batch_no']
    
    cur.execute('UPDATE inventory_stock SET total_inventory = MAX(0, total_inventory - ?) WHERE sku = ?', (qty, sku))
    cur.execute('DELETE FROM stock_inward_batches WHERE id = ?', (batch_id,))
    
    record_audit(
        conn,
        'DELETE_STOCK_BATCH',
        'Inventory Stock & Batch Log',
        f"Batch {batch_no} | {sku}",
        f"Reversed batch inwarding of {qty} units."
    )
    conn.commit()
    conn.close()
    
    return jsonify({'status': 'success', 'message': f"Batch '{batch_no}' deleted and {qty} units reversed from total inventory."})

# --- UNIVERSAL FILTERED EXCEL EXPORT ---
@app.route('/api/export-current-view')
def export_current_view():
    """Generates an Excel spreadsheet corresponding directly to the user's currently active tab and filters."""
    view = request.args.get('view', 'master').lower()
    platform = request.args.get('platform', 'All')
    brand = request.args.get('brand', 'All')
    sku = request.args.get('sku', 'All').strip()
    search = request.args.get('search', '').strip()
    channel = request.args.get('channel', 'All')
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    title_font = Font(color='1E3A8A', bold=True, size=14)
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    ws = wb.active
    
    if view in ['courier_costing', 'courier']:
        ws.title = "Courier Costing"
        ws.cell(1, 1, "💰 COURIER COSTING & FREIGHT EXPENSES").font = title_font
        ws.cell(2, 1, f"Channel: {channel} | Search: {search or 'None'} | Date: {start_date or 'Any'} to {end_date or 'Any'}")
        
        headers = [
            "Channel", "Location", "PO Date", "Appointment Date", "PO No", 
            "LR (Tracking/AWB)", "Courier Service", "Sum of Qty", "PO Value (₹)", 
            "Courier Charges (₹)", "Freight %", "Pickup Date", "Actual Weight (kg)", 
            "Charged Weight (kg)", "Weight Variance (kg)", "Dispatch Status", "SKU", "Product Name", "Brand"
        ]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='0F766E', end_color='0F766E', fill_type='solid')
            cell.font = header_font
            
        rows, summary = get_courier_costing_rows_internal(channel, search, start_date, end_date)
        for r_idx, r in enumerate(rows, 5):
            row_vals = [
                r['channel'], r['location'], r['po_date'], r['appointment_date'], r['po_number'],
                r['tracking_id'], r['courier_name'], r['sum_of_qty'], r['po_value'],
                r['courier_charges'], f"{r['percentage']}%", r['pickup_date'], r['actual_weight'],
                r['charged_weight'], r['weight_variance'], r['dispatch_status'], r['sku'], r['product_name'], r['brand']
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border
                
        last_row = 5 + len(rows)
        ws.cell(last_row, 1, "TOTALS").font = Font(bold=True)
        ws.cell(last_row, 8, sum(r['sum_of_qty'] for r in rows)).font = Font(bold=True)
        ws.cell(last_row, 9, summary['total_po_value']).font = Font(bold=True)
        ws.cell(last_row, 10, summary['total_freight_spent']).font = Font(bold=True)
        ws.cell(last_row, 11, f"{summary['overall_freight_percentage']}%").font = Font(bold=True)
        ws.cell(last_row, 13, summary['total_actual_weight']).font = Font(bold=True)
        ws.cell(last_row, 14, summary['total_charged_weight']).font = Font(bold=True)
        ws.cell(last_row, 15, summary['weight_discrepancy']).font = Font(bold=True)

    elif view in ['batches', 'inward_batches']:
        ws.title = "Stock Inward Batches"
        ws.cell(1, 1, "📦 INVENTORY INWARD LOG (DATE-WISE & BATCH-WISE)").font = title_font
        ws.cell(2, 1, f"Filtered by Brand: {brand} | SKU: {sku} | Search: {search or 'None'}")
        
        headers = ["ID", "Inward Date", "Batch No / Lot #", "SKU", "Product Name", "Brand", "Quantity Added", "Mfg Date", "Expiry Date", "Supplier / PO Ref", "Notes", "Added By", "Created At"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
            cell.font = header_font
            
        conn = get_db()
        cur = conn.cursor()
        b_query = 'SELECT * FROM stock_inward_batches WHERE 1=1'
        b_params = []
        if brand and brand != 'All':
            b_query += ' AND brand = ?'
            b_params.append(brand)
        if sku and sku != 'All':
            b_query += ' AND sku = ?'
            b_params.append(sku)
        if search:
            b_query += ' AND (batch_no LIKE ? OR sku LIKE ? OR product_name LIKE ? OR supplier_po_ref LIKE ?)'
            term = f'%{search}%'
            b_params.extend([term, term, term, term])
        b_query += ' ORDER BY inward_date DESC, id DESC'
        b_rows = cur.execute(b_query, b_params).fetchall()
        conn.close()
        
        for r_idx, r in enumerate(b_rows, 5):
            row_vals = [
                r['id'], r['inward_date'], r['batch_no'], r['sku'], r['product_name'], r['brand'],
                r['quantity_added'], r['mfg_date'], r['expiry_date'], r['supplier_po_ref'], r['notes'],
                r['added_by'], r['created_at']
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border

    elif view in ['stock', 'baselines', 'inventory_stock']:
        ws.title = "Current Stock & Thresholds"
        ws.cell(1, 1, "📦 CENTRAL INVENTORY STOCK & REORDER ALERTS").font = title_font
        ws.cell(2, 1, f"Filtered by Brand: {brand} | SKU: {sku} | Search: {search or 'None'}")
        
        headers = ["SKU", "Product Name", "Brand", "Category", "Total Central Inventory", "Reorder Alert Threshold", "Stock Status"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
            cell.font = header_font
            
        conn = get_db()
        cur = conn.cursor()
        s_query = '''
        SELECT p.sku, p.name as product_name, p.brand, p.category, 
               COALESCE(s.total_inventory, 0) as total_inventory,
               COALESCE(s.reorder_threshold, 0) as reorder_threshold
        FROM products p
        LEFT JOIN inventory_stock s ON p.sku = s.sku
        WHERE 1=1
        '''
        s_params = []
        if brand and brand != 'All':
            s_query += ' AND p.brand = ?'
            s_params.append(brand)
        if sku and sku != 'All':
            s_query += ' AND p.sku = ?'
            s_params.append(sku)
        if search:
            term = f'%{search}%'
            s_query += ' AND (p.sku LIKE ? OR p.name LIKE ? OR p.category LIKE ?)'
            s_params.extend([term, term, term])
        s_query += ' ORDER BY p.brand, p.name'
        s_rows = cur.execute(s_query, s_params).fetchall()
        conn.close()
        
        for r_idx, r in enumerate(s_rows, 5):
            status = 'Low Stock Alert' if r['total_inventory'] <= r['reorder_threshold'] else 'OK'
            row_vals = [
                r['sku'], r['product_name'], r['brand'], r['category'],
                r['total_inventory'], r['reorder_threshold'], status
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border

    elif view in ['online_platforms', 'platform_detail']:
        ws.title = f"Online - {platform if platform != 'All' else 'All Platforms'}"
        ws.cell(1, 1, f"⚡ ONLINE PLATFORM REPORT: {platform}").font = title_font
        ws.cell(2, 1, f"Brand: {brand} | Search: {search or 'None'}")
        
        headers = ["ID", "Date", "Brand", "Product Name", "SKU", "Platform", "Location", "PO Number", "PO Qty", "Actual Sent", "Shortage", "Fulfillment %", "Appointment Date", "Courier", "AWB / Tracking", "Freight (₹)", "Status"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='065F46', end_color='065F46', fill_type='solid')
            cell.font = header_font
            
        conn = get_db()
        cur = conn.cursor()
        q = 'SELECT * FROM dispatch_online WHERE 1=1'
        p = []
        if platform and platform != 'All':
            q += ' AND platform = ?'
            p.append(platform)
        if brand and brand != 'All':
            q += ' AND brand = ?'
            p.append(brand)
        if search:
            q += ' AND (sku LIKE ? OR product_name LIKE ? OR po_number LIKE ? OR tracking_id LIKE ?)'
            term = f'%{search}%'
            p.extend([term, term, term, term])
        q += ' ORDER BY id DESC'
        d_rows = cur.execute(q, p).fetchall()
        conn.close()
        
        for r_idx, r in enumerate(d_rows, 5):
            po_qty = r['po_quantity'] or 0
            actual = r['actual_sent'] or 0
            rate = f"{round(actual / po_qty * 100, 1)}%" if po_qty > 0 else "100%"
            row_vals = [
                r['id'], r['dispatch_date'], r['brand'], r['product_name'], r['sku'], r['platform'], r['location'],
                r['po_number'], po_qty, actual, po_qty - actual, rate, r['appointment_date'],
                r['courier_name'], r['tracking_id'], r['courier_charges'], r['dispatch_status']
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border

    elif view in ['gt', 'mt']:
        ch_title = view.upper()
        ws.title = ch_title
        ws.cell(1, 1, f"🏢 {ch_title} CHANNEL REPORT").font = title_font
        ws.cell(2, 1, f"Brand: {brand} | Search: {search or 'None'}")
        
        headers = ["ID", "Date", "Brand", "Product Name", "SKU", "Channel", "Location", "PO Number", "PO Qty", "Actual Sent", "Shortage", "Fulfillment %", "Appointment Date", "Courier", "AWB / Tracking", "Freight (₹)", "Status"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='1E3A8A' if view == 'mt' else '92400E', end_color='1E3A8A' if view == 'mt' else '92400E', fill_type='solid')
            cell.font = header_font
            
        conn = get_db()
        cur = conn.cursor()
        q = 'SELECT * FROM dispatch_gt_mt WHERE channel = ?'
        p = [ch_title]
        if brand and brand != 'All':
            q += ' AND brand = ?'
            p.append(brand)
        if search:
            q += ' AND (sku LIKE ? OR product_name LIKE ? OR po_number LIKE ? OR tracking_id LIKE ?)'
            term = f'%{search}%'
            p.extend([term, term, term, term])
        q += ' ORDER BY id DESC'
        g_rows = cur.execute(q, p).fetchall()
        conn.close()
        
        for r_idx, r in enumerate(g_rows, 5):
            po_qty = r['po_quantity'] or 0
            actual = r['actual_sent'] or 0
            rate = f"{round(actual / po_qty * 100, 1)}%" if po_qty > 0 else "100%"
            row_vals = [
                r['id'], r['dispatch_date'], r['brand'], r['product_name'], r['sku'], r['channel'], r['location'],
                r['po_number'], po_qty, actual, po_qty - actual, rate, r['appointment_date'],
                r['courier_name'], r['tracking_id'], r['courier_charges'], r['dispatch_status']
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border

    elif view in ['sales']:
        ws.title = "Sales & Offtake"
        ws.cell(1, 1, "📈 CONSUMER SALES & OFFTAKE DATA").font = title_font
        headers = ["ID", "Sale Date", "Platform", "SKU", "Product Name", "Brand", "Units Sold", "Revenue (₹)", "Store Location", "Source File", "Uploaded By", "Created At"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = PatternFill(start_color='065F46', end_color='065F46', fill_type='solid')
            cell.font = header_font
        conn = get_db()
        cur = conn.cursor()
        s_rows = cur.execute('SELECT * FROM sales_data ORDER BY sale_date DESC, id DESC').fetchall()
        conn.close()
        for r_idx, r in enumerate(s_rows, 5):
            row_vals = [
                r['id'], r['sale_date'], r['platform'], r['sku'], r['product_name'], r['brand'],
                r['units_sold'], r['revenue'], r['store_location'], r['source_file'], r['uploaded_by'], r['created_at']
            ]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(r_idx, c_idx, val).border = thin_border

    else:
        ws.title = "Master Summary"
        ws.cell(1, 1, "🌿 MASTER DISPATCH & STOCK SUMMARY").font = title_font
        ws.cell(2, 1, f"Filtered by Brand: {brand} | Search: {search or 'None'}")
        
        master_headers = ["SKU", "Brand", "Product Name", "Category"] + [f"{p} Sent" for p in PLATFORMS] + [
            "Online Total", "GT Total", "MT Total", "Grand Total", "Total Adjustments", "Total Inventory", "Remaining Inventory", "Stock Status"
        ]
        for c_idx, h in enumerate(master_headers, 1):
            cell = ws.cell(4, c_idx, h)
            cell.fill = header_fill
            cell.font = header_font
            
        conn = get_db()
        cur = conn.cursor()
        m_data = get_master_data_internal(cur, brand)
        conn.close()
        if search:
            term = search.lower().strip()
            m_data = [r for r in m_data if term in (r['sku'].lower() + " " + r['brand'].lower() + " " + r['product_name'].lower())]
            
        for r_idx, r in enumerate(m_data, 5):
            ws.cell(r_idx, 1, r['sku'])
            ws.cell(r_idx, 2, r['brand'])
            ws.cell(r_idx, 3, r['product_name'])
            ws.cell(r_idx, 4, r['category'])
            for p_idx, p in enumerate(PLATFORMS):
                ws.cell(r_idx, 5 + p_idx, r.get(f"sent_{p.lower().replace(' ', '_')}", 0))
            ws.cell(r_idx, 15, r['online_total'])
            ws.cell(r_idx, 16, r['gt_total'])
            ws.cell(r_idx, 17, r['mt_total'])
            ws.cell(r_idx, 18, r['grand_total'])
            ws.cell(r_idx, 19, r['total_adjustments'])
            ws.cell(r_idx, 20, r['total_inventory'])
            ws.cell(r_idx, 21, r['remaining_inventory'])
            ws.cell(r_idx, 22, r['status'])
            for c in range(1, 23):
                ws.cell(r_idx, c).border = thin_border

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    clean_view = view.replace('/', '_').replace('\\', '_')
    filename = f"Export_{clean_view}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    return send_file(
        stream,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

# --- EXCEL EXPORT (INCLUDES ALL 10 SHEETS + SALES & PIPELINE) ---
@app.route('/api/export-excel')
def export_excel():
    conn = get_db()
    cur = conn.cursor()
    
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    title_font = Font(color='1E3A8A', bold=True, size=14)
    alert_fill = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
    alert_font = Font(color='991B1B', bold=True)
    ok_fill = PatternFill(start_color='DCFCE7', end_color='DCFCE7', fill_type='solid')
    ok_font = Font(color='166534', bold=True)
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    # 1. Master Sheet
    ws_master = wb.active
    ws_master.title = "Master Sheet"
    ws_master.cell(1, 1, "🌿 MASTER DISPATCH SUMMARY — All Brands, Platforms, GT & MT").font = title_font
    ws_master.cell(2, 1, "Live totals across all Online platforms + GT + MT — Updated with PO Quantities and Actual Sent")
    
    master_headers = [
        "SKU", "Brand", "Product Name", "Category"
    ] + [f"{p} Units Sent" for p in PLATFORMS] + [
        "Online Total", "GT Total", "MT Total", "Grand Total", "Total Adjustments", 
        "Total Inventory", "Remaining Inventory", "Stock Status"
    ]
    
    for c_idx, h in enumerate(master_headers, 1):
        cell = ws_master.cell(4, c_idx, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
    master_data = get_master_data_internal(cur)
    for r_idx, r in enumerate(master_data, 5):
        ws_master.cell(r_idx, 1, r['sku'])
        ws_master.cell(r_idx, 2, r['brand'])
        ws_master.cell(r_idx, 3, r['product_name'])
        ws_master.cell(r_idx, 4, r['category'])
        for p_idx, p in enumerate(PLATFORMS):
            ws_master.cell(r_idx, 5 + p_idx, r.get(f"sent_{p.lower().replace(' ', '_')}", 0))
        ws_master.cell(r_idx, 15, r['online_total'])
        ws_master.cell(r_idx, 16, r['gt_total'])
        ws_master.cell(r_idx, 17, r['mt_total'])
        ws_master.cell(r_idx, 18, r['grand_total'])
        ws_master.cell(r_idx, 19, r['total_adjustments'])
        ws_master.cell(r_idx, 20, r['total_inventory'])
        ws_master.cell(r_idx, 21, r['remaining_inventory'])
        
        status_cell = ws_master.cell(r_idx, 22, r['status'])
        if r['status'] == 'Low Stock':
            status_cell.fill = alert_fill
            status_cell.font = alert_font
        else:
            status_cell.fill = ok_fill
            status_cell.font = ok_font
            
        for c in range(1, 23):
            ws_master.cell(r_idx, c).border = thin_border
            
    # 2. Dispatch Online
    ws_disp = wb.create_sheet("Dispatch Online")
    ws_disp.cell(1, 1, "📋 DISPATCH ONLINE — All Online Platforms with PO Tracking & Courier Details").font = title_font
    disp_headers = [
        "ID", "Date", "Brand", "Product Name", "SKU", "Category", "Platform", "Location", 
        "PO Number", "Quantity in PO", "Actual Sent in PO", "Shortage (Variance)", "Fulfillment %", "Appointment Date",
        "Courier Service", "Tracking ID / AWB", "Courier Charges (₹)", "Dispatch Status", "E-Way Bill No", "Shipping Notes"
    ]
    for c_idx, h in enumerate(disp_headers, 1):
        cell = ws_disp.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='065F46', end_color='065F46', fill_type='solid')
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    disp_rows = cur.execute('SELECT * FROM dispatch_online ORDER BY id DESC').fetchall()
    for r_idx, r in enumerate(disp_rows, 5):
        po_qty = r['po_quantity'] or 0
        actual = r['actual_sent'] or 0
        var = po_qty - actual
        rate = f"{round(actual / po_qty * 100, 1)}%" if po_qty > 0 else "100%"
        row_vals = [
            r['id'], r['dispatch_date'], r['brand'], r['product_name'], r['sku'], r['category'],
            r['platform'], r['location'], r['po_number'], po_qty, actual, var, rate, r['appointment_date'],
            r['courier_name'] or '', r['tracking_id'] or '', r['courier_charges'] or 0.0,
            r['dispatch_status'] or 'Dispatched', r['eway_bill_no'] or '', r['shipping_notes'] or ''
        ]
        for c_idx, val in enumerate(row_vals, 1):
            ws_disp.cell(r_idx, c_idx, val).border = thin_border
            
    # 3. Individual Platform Sheets
    for plat in PLATFORMS:
        ws_p = wb.create_sheet(plat[:31])
        ws_p.cell(1, 1, f"🏪  {plat} — Units Sent (All Brands)").font = title_font
        p_headers = ["SKU", "Brand", "Product Name", "Category", f"Total Units Sent to {plat}"]
        for c_idx, h in enumerate(p_headers, 1):
            c = ws_p.cell(4, c_idx, h)
            c.fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
            c.font = header_font
            
        p_data = cur.execute('''
            SELECT p.sku, p.brand, p.name, p.category, COALESCE(SUM(o.actual_sent), 0) as sent
            FROM products p
            LEFT JOIN dispatch_online o ON p.sku = o.sku AND o.platform = ?
            GROUP BY p.sku
            ORDER BY p.brand, p.name
        ''', (plat,)).fetchall()
        for r_idx, pr in enumerate(p_data, 5):
            ws_p.cell(r_idx, 1, pr['sku']).border = thin_border
            ws_p.cell(r_idx, 2, pr['brand']).border = thin_border
            ws_p.cell(r_idx, 3, pr['name']).border = thin_border
            ws_p.cell(r_idx, 4, pr['category']).border = thin_border
            ws_p.cell(r_idx, 5, pr['sent']).border = thin_border

    # 4. GT & MT Sheets
    for ch in CHANNELS:
        ws_ch = wb.create_sheet(f"Channel {ch}")
        ws_ch.cell(1, 1, f"🏪  {ch} (Trade) — Units Sent (All Brands)").font = title_font
        ch_headers = ["SKU", "Brand", "Product Name", "Category", f"Total Units Sent to {ch}"]
        for c_idx, h in enumerate(ch_headers, 1):
            c = ws_ch.cell(4, c_idx, h)
            c.fill = PatternFill(start_color='92400E', end_color='92400E', fill_type='solid')
            c.font = header_font
        ch_data = cur.execute('''
            SELECT p.sku, p.brand, p.name, p.category, COALESCE(SUM(g.actual_sent), 0) as sent
            FROM products p
            LEFT JOIN dispatch_gt_mt g ON p.sku = g.sku AND g.channel = ?
            GROUP BY p.sku
            ORDER BY p.brand, p.name
        ''', (ch,)).fetchall()
        for r_idx, chr in enumerate(ch_data, 5):
            ws_ch.cell(r_idx, 1, chr['sku']).border = thin_border
            ws_ch.cell(r_idx, 2, chr['brand']).border = thin_border
            ws_ch.cell(r_idx, 3, chr['name']).border = thin_border
            ws_ch.cell(r_idx, 4, chr['category']).border = thin_border
            ws_ch.cell(r_idx, 5, chr['sent']).border = thin_border

    # 5. Dispatch GT MT
    ws_gt = wb.create_sheet("Dispatch GT MT")
    ws_gt.cell(1, 1, "📋 DISPATCH GT MT — General Trade & Modern Trade with Courier & Charges").font = title_font
    gt_headers = [
        "ID", "Date", "Brand", "Product Name", "SKU", "Category", "Channel (GT/MT)", 
        "Buyer / Distributor", "Location", "PO Number", "Invoice No", "Quantity in PO", "Actual Sent in PO", "Shortage (Variance)", "Appointment Date",
        "Transporter / Courier", "LR / Tracking #", "Courier Charges (₹)", "Status", "E-Way Bill No", "Shipping Notes"
    ]
    for c_idx, h in enumerate(gt_headers, 1):
        cell = ws_gt.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='92400E', end_color='92400E', fill_type='solid')
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    gt_rows = cur.execute('SELECT * FROM dispatch_gt_mt ORDER BY id DESC').fetchall()
    for r_idx, r in enumerate(gt_rows, 5):
        po_qty = r['po_quantity'] or 0
        actual = r['actual_sent'] or 0
        var = po_qty - actual
        row_vals = [
            r['id'], r['dispatch_date'], r['brand'], r['product_name'], r['sku'], r['category'],
            r['channel'], r['buyer_distributor'], r['location'], r['po_number'], r['invoice_no'], po_qty, actual, var, r['appointment_date'],
            r['courier_name'] or '', r['tracking_id'] or '', r['courier_charges'] or 0.0,
            r['dispatch_status'] or 'Dispatched', r['eway_bill_no'] or '', r['shipping_notes'] or ''
        ]
        for c_idx, val in enumerate(row_vals, 1):
            ws_gt.cell(r_idx, c_idx, val).border = thin_border

    # 6. Sales Data & Pipeline Sheet (NEW!)
    ws_sales = wb.create_sheet("Sales & Pipeline")
    ws_sales.cell(1, 1, "📈 SALES OFFTAKE & DARK STORE PIPELINE STOCK").font = title_font
    sales_headers = ["ID", "Sale Date", "Platform", "SKU", "Product Name", "Brand", "Units Sold", "Revenue (₹)", "Dark Store / City", "Source File", "Uploaded By"]
    for c_idx, h in enumerate(sales_headers, 1):
        cell = ws_sales.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='7C3AED', end_color='7C3AED', fill_type='solid')
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    sales_rows = cur.execute('SELECT * FROM sales_data ORDER BY sale_date DESC, id DESC LIMIT 500').fetchall()
    for r_idx, sr in enumerate(sales_rows, 5):
        row_vals = [
            sr['id'], sr['sale_date'], sr['platform'], sr['sku'], sr['product_name'], sr['brand'],
            sr['units_sold'], sr['revenue'], sr['store_location'], sr['source_file'], sr['uploaded_by']
        ]
        for c_idx, val in enumerate(row_vals, 1):
            ws_sales.cell(r_idx, c_idx, val).border = thin_border
            
    # 7. Adjustments
    ws_adj = wb.create_sheet("Adjustments")
    ws_adj.cell(1, 1, "📉 NON-DISPATCH STOCK ADJUSTMENTS").font = title_font
    adj_headers = ["ID", "Date", "Brand", "Channel", "Product Name", "SKU", "Category", "Type", "Quantity", "Reason / Notes"]
    for c_idx, h in enumerate(adj_headers, 1):
        cell = ws_adj.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='4C1D95', end_color='4C1D95', fill_type='solid')
        cell.font = header_font
    adj_rows = cur.execute('SELECT * FROM adjustments ORDER BY id DESC').fetchall()
    for r_idx, r in enumerate(adj_rows, 5):
        row_vals = [
            r['id'], r['adjustment_date'], r['brand'], r['channel'], r['product_name'], r['sku'], r['category'],
            r['adjustment_type'], r['quantity'], r['reason']
        ]
        for c_idx, val in enumerate(row_vals, 1):
            ws_adj.cell(r_idx, c_idx, val).border = thin_border
            
    # 8. Inventory Stock & Thresholds
    ws_stk = wb.create_sheet("Inventory Stock")
    ws_stk.cell(1, 1, "📦 INVENTORY STOCK & THRESHOLDS").font = title_font
    stk_headers = ["SKU", "Brand", "Product Name", "Category", "Total Inventory", "Reorder Threshold"]
    for c_idx, h in enumerate(stk_headers, 1):
        cell = ws_stk.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
        cell.font = header_font
    stk_rows = cur.execute('''
        SELECT p.sku, p.brand, p.name, p.category, s.total_inventory, s.reorder_threshold
        FROM products p LEFT JOIN inventory_stock s ON p.sku = s.sku ORDER BY p.brand, p.name
    ''').fetchall()
    for r_idx, r in enumerate(stk_rows, 5):
        row_vals = [r['sku'], r['brand'], r['name'], r['category'], r['total_inventory'], r['reorder_threshold']]
        for c_idx, val in enumerate(row_vals, 1):
            ws_stk.cell(r_idx, c_idx, val).border = thin_border
            
    # 9. Audit Log Sheet
    ws_audit = wb.create_sheet("Audit Trail")
    ws_audit.cell(1, 1, "🕒 SYSTEM AUDIT TRAIL & CHANGE LOG").font = title_font
    audit_headers = ["Log ID", "Date & Time", "Person Name", "Email Address", "Action Type", "Sheet / Area", "Entity Reference", "Edit Summary"]
    for c_idx, h in enumerate(audit_headers, 1):
        cell = ws_audit.cell(4, c_idx, h)
        cell.fill = PatternFill(start_color='374151', end_color='374151', fill_type='solid')
        cell.font = header_font
    audit_rows = cur.execute('SELECT * FROM audit_log ORDER BY id DESC LIMIT 500').fetchall()
    for r_idx, ar in enumerate(audit_rows, 5):
        row_vals = [
            ar['id'], ar['timestamp'], ar['user_name'], ar['user_email'],
            ar['action_type'] or ar['action'], ar['sheet_name'], ar['entity_ref'], ar['edit_summary'] or ar['details']
        ]
        for c_idx, val in enumerate(row_vals, 1):
            ws_audit.cell(r_idx, c_idx, val).border = thin_border
            
    conn.close()
    
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    filename = f"Inventory_Master_Export_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    return send_file(
        stream,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

# ==============================================================================
# BLINKIT AI REPLENISHMENT & STORAGE COST OPTIMIZATION ENGINE
# ==============================================================================

def parse_blinkit_inventory_data(file_content, cur):
    """
    Parses raw Blinkit Inventory CSV/Excel lines and upserts into blinkit_inventory_current.
    """
    if isinstance(file_content, bytes):
        file_content = file_content.decode('utf-8', errors='ignore')
    
    lines = file_content.splitlines() if isinstance(file_content, str) else file_content
    header_idx = -1
    for i, l in enumerate(lines[:15]):
        if 'Item ID' in l and ('Warehouse Facility' in l or 'Facility ID' in l):
            header_idx = i
            break
            
    if header_idx == -1:
        return {'error': 'Could not detect valid Blinkit Inventory header row (expected Item ID, Warehouse Facility Name).'}
        
    reader = csv.DictReader(lines[header_idx:])
    rows_processed = 0
    facilities_seen = set()
    
    for r in reader:
        try:
            item_id = int(str(r.get('Item ID', '')).strip())
            fac_id = int(str(r.get('Warehouse Facility ID', '')).strip())
        except (ValueError, TypeError):
            continue
            
        item_name = str(r.get('Item Name', '')).strip()
        brand_name = str(r.get('Brand Name', '')).strip()
        upc = str(r.get('UPC', '')).strip()
        uom = str(r.get('UoM', '')).strip()
        fac_name = str(r.get('Warehouse Facility Name', '')).strip()
        facilities_seen.add(fac_name)
        
        def safe_int(key):
            try:
                v = r.get(key, 0)
                return int(float(str(v).replace(',', '').strip() or 0))
            except:
                return 0
                
        net_sched = safe_int('Net scheduled inventory')
        incom_sched = safe_int('Incoming scheduled inventory')
        recalled = safe_int('Recalled inventory')
        tot_sellable = safe_int('Total sellable')
        wh_stock = safe_int('Warehouse')
        in_between = safe_int('In-between')
        darkstore = safe_int('Darkstore')
        tot_unsellable = safe_int('Total unsellable')
        damaged = safe_int('Damaged')
        lost = safe_int('Lost')
        expired = safe_int('Expired')
        near_expiry = safe_int('Near Expiry')
        s7 = safe_int('Last 7 days')
        s15 = safe_int('Last 15 days')
        s30 = safe_int('Last 30 days')
        
        cur.execute('''
        INSERT INTO blinkit_inventory_current (
            item_id, facility_id, item_name, brand_name, upc, uom, facility_name,
            net_scheduled, incoming_scheduled, recalled_inventory, total_sellable,
            warehouse_stock, in_between_stock, darkstore_stock, total_unsellable,
            damaged, lost, expired, near_expiry, sales_7d, sales_15d, sales_30d, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(item_id, facility_id) DO UPDATE SET
            item_name=excluded.item_name,
            brand_name=excluded.brand_name,
            upc=excluded.upc,
            uom=excluded.uom,
            facility_name=excluded.facility_name,
            net_scheduled=excluded.net_scheduled,
            incoming_scheduled=excluded.incoming_scheduled,
            recalled_inventory=excluded.recalled_inventory,
            total_sellable=excluded.total_sellable,
            warehouse_stock=excluded.warehouse_stock,
            in_between_stock=excluded.in_between_stock,
            darkstore_stock=excluded.darkstore_stock,
            total_unsellable=excluded.total_unsellable,
            damaged=excluded.damaged,
            lost=excluded.lost,
            expired=excluded.expired,
            near_expiry=excluded.near_expiry,
            sales_7d=excluded.sales_7d,
            sales_15d=excluded.sales_15d,
            sales_30d=excluded.sales_30d,
            updated_at=CURRENT_TIMESTAMP
        ''', (
            item_id, fac_id, item_name, brand_name, upc, uom, fac_name,
            net_sched, incom_sched, recalled, tot_sellable,
            wh_stock, in_between, darkstore, tot_unsellable,
            damaged, lost, expired, near_expiry, s7, s15, s30
        ))
        rows_processed += 1
        
    return {
        'status': 'success',
        'rows_processed': rows_processed,
        'facilities_count': len(facilities_seen),
        'facilities': sorted(list(facilities_seen))
    }

def parse_blinkit_sales_data(file_content, cur):
    """
    Parses Blinkit Daily Sales Order CSV lines and deduplicates with INSERT OR IGNORE by order_id.
    """
    if isinstance(file_content, bytes):
        file_content = file_content.decode('utf-8', errors='ignore')
        
    lines = file_content.splitlines() if isinstance(file_content, str) else file_content
    reader = csv.DictReader(lines)
    
    total_in_file = 0
    new_inserted = 0
    dates_seen = set()
    
    for r in reader:
        order_id = str(r.get('Order Id', '')).strip()
        if not order_id:
            continue
        total_in_file += 1
        
        order_date = str(r.get('Order Date', '')).strip()
        if order_date:
            dates_seen.add(order_date)
            
        try:
            item_id = int(str(r.get('Item Id', '')).strip())
        except:
            item_id = 0
            
        p_name = str(r.get('Product Name', '')).strip()
        b_name = str(r.get('Brand Name', '')).strip()
        upc = str(r.get('UPC', '')).strip()
        supply_city = str(r.get('Supply City', '')).strip()
        supply_state = str(r.get('Supply State', '')).strip()
        cust_city = str(r.get('Customer City', '')).strip()
        cust_state = str(r.get('Customer State', '')).strip()
        order_status = str(r.get('Order Status', 'DELIVERED')).strip()
        
        try:
            qty = int(float(str(r.get('Quantity', 1)).strip() or 1))
        except:
            qty = 1
        try:
            mrp = float(str(r.get('MRP (Rs)', 0)).replace(',', '').strip() or 0.0)
        except:
            mrp = 0.0
        try:
            selling_price = float(str(r.get('Selling Price (Rs)', 0)).replace(',', '').strip() or 0.0)
        except:
            selling_price = 0.0
        try:
            tot_gross = float(str(r.get('Total Gross Bill Amount', 0)).replace(',', '').strip() or 0.0)
        except:
            tot_gross = 0.0
            
        cur.execute('''
        INSERT OR IGNORE INTO blinkit_sales_orders (
            order_id, order_date, item_id, product_name, brand_name, upc,
            supply_city, supply_state, customer_city, customer_state, order_status,
            quantity, mrp, selling_price, total_gross_amount
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            order_id, order_date, item_id, p_name, b_name, upc,
            supply_city, supply_state, cust_city, cust_state, order_status,
            qty, mrp, selling_price, tot_gross
        ))
        if cur.rowcount > 0:
            new_inserted += 1
            
    skipped = total_in_file - new_inserted
    min_date = min(dates_seen) if dates_seen else '-'
    max_date = max(dates_seen) if dates_seen else '-'

    # Anti-overlap sync to central sales_data with canonical product resolution:
    for d in dates_seen:
        cur.execute('DELETE FROM sales_data WHERE platform = ? AND sale_date = ?', ('Blinkit', d))
        blk_rows = cur.execute('''
        SELECT order_date, item_id, product_name, brand_name, upc, supply_city,
               SUM(quantity) as units, SUM(total_gross_amount) as revenue
        FROM blinkit_sales_orders
        WHERE order_date = ?
        GROUP BY order_date, item_id, supply_city
        ''', (d,)).fetchall()
        
        for r in blk_rows:
            c_sku, c_name, c_brand, _ = resolve_canonical_product(None, r['product_name'], r['upc'], cur)
            brand_to_use = c_brand if c_brand != 'Other' else (r['brand_name'] or 'Blinkit')
            name_to_use = c_name if c_name != 'Unmapped Product' else r['product_name']
            sku_to_use = c_sku if c_sku != 'UNKNOWN' else f"BLK-{r['item_id']}"
            
            cur.execute('''
            INSERT INTO sales_data (
                sale_date, platform, sku, product_name, brand, units_sold, revenue, store_location, source_file, uploaded_by
            ) VALUES (?, 'Blinkit', ?, ?, ?, ?, ?, ?, 'Blinkit Orders', 'Blinkit System')
            ''', (r['order_date'], sku_to_use, name_to_use, brand_to_use, r['units'], r['revenue'], r['supply_city']))
    
    return {
        'status': 'success',
        'total_orders_in_file': total_in_file,
        'new_orders_inserted': new_inserted,
        'existing_orders_skipped': skipped,
        'date_range': f"{min_date} to {max_date}" if min_date != '-' else '-'
    }

def seed_initial_blinkit_data_if_empty():
    """Auto-seeds initial Blinkit inventory and sales datasets if current tables are empty."""
    try:
        conn = get_db()
        cur = conn.cursor()
        count = cur.execute('SELECT COUNT(*) FROM blinkit_inventory_current').fetchone()[0]
        if count == 0:
            brain_uploads = os.path.join(os.path.expanduser('~'), '.gemini', 'antigravity', 'brain', 'c18c0bb2-1dda-4e5b-89b7-d36de3bd9d79', '.user_uploaded')
            p_inv = os.path.join(brain_uploads, 'media_1790511403279.csv')
            p_sales = os.path.join(brain_uploads, 'media_1790511445429.csv')
            if os.path.exists(p_inv):
                with open(p_inv, 'rb') as f:
                    parse_blinkit_inventory_data(f.read(), cur)
            if os.path.exists(p_sales):
                with open(p_sales, 'rb') as f:
                    parse_blinkit_sales_data(f.read(), cur)
            conn.commit()
            print("Auto-seeded Blinkit inventory & sales datasets successfully.")
        conn.close()
    except Exception as e:
        print("Blinkit seed notice:", e)

seed_initial_blinkit_data_if_empty()

# ==============================================================================
# QUICK-COMMERCE MULTI-PLATFORM PARSERS (ZEPTO & INSTAMART) WITH ANTI-OVERLAP
# ==============================================================================

# (Canonical resolver and normalization engine are defined at the top of the file)


def parse_zepto_sales_data(file_content, filename='Zepto_Sales.csv', user_info='Admin', cur=None):
    """
    Parses Zepto Daily Sales CSV/Excel lines.
    Deduplicates strictly using UNIQUE(sale_date, sku_number, city) with ON CONFLICT DO UPDATE.
    Guarantees zero overlap when same dates are re-uploaded.
    """
    if isinstance(file_content, bytes):
        if filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls'):
            wb = openpyxl.load_workbook(BytesIO(file_content), data_only=True)
            sheet = wb.active
            lines = []
            for row in sheet.iter_rows(values_only=True):
                lines.append(','.join(['' if v is None else str(v).replace(',', ' ') for v in row]))
            file_content = '\n'.join(lines)
        else:
            file_content = file_content.decode('utf-8', errors='ignore')
            
    lines = file_content.splitlines() if isinstance(file_content, str) else file_content
    header_idx = 0
    for idx, l in enumerate(lines[:10]):
        if 'SKU Number' in l or 'SKU Name' in l:
            header_idx = idx
            break
            
    reader = csv.DictReader(lines[header_idx:])
    total_in_file = 0
    upserted_count = 0
    dates_seen = set()
    total_units = 0
    total_gmv = 0.0
    cities_seen = set()
    
    for r in reader:
        raw_date = str(r.get('Date', '')).strip()
        sku_num = str(r.get('SKU Number', '')).strip()
        if not raw_date or not sku_num:
            continue
            
        norm_date = normalize_date_str(raw_date)
        dates_seen.add(norm_date)
        total_in_file += 1
        
        sku_name = str(r.get('SKU Name', '')).strip()
        ean = str(r.get('EAN', '')).strip()
        cat = str(r.get('SKU Category', '')).strip()
        subcat = str(r.get('SKU Sub Category', '')).strip()
        brand = str(r.get('Brand Name', '')).strip()
        mfg_name = str(r.get('Manufacturer Name', '')).strip()
        mfg_id = str(r.get('Manufacturer ID', '')).strip()
        city = str(r.get('City', 'Unknown')).strip()
        cities_seen.add(city)
        
        try:
            units = int(float(str(r.get('Sales (Qty) - Units', 0)).replace(',', '').strip() or 0))
        except:
            units = 0
        try:
            mrp = float(str(r.get('MRP', 0)).replace(',', '').strip() or 0.0)
        except:
            mrp = 0.0
        try:
            gmv = float(str(r.get('Gross Merchandise Value', 0)).replace(',', '').strip() or 0.0)
        except:
            gmv = 0.0
            
        total_units += units
        total_gmv += gmv
        
        matched_sku, matched_name, matched_brand = map_product_to_internal(sku_num, sku_name, ean, cur)
        if not brand and matched_brand:
            brand = matched_brand
            
        cur.execute('''
        INSERT INTO zepto_sales_orders (
            sale_date, sku_number, sku_name, ean, sku_category, sku_sub_category,
            brand_name, manufacturer_name, manufacturer_id, city,
            units_sold, mrp, gmv, matched_sku, matched_product_name, source_file, uploaded_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sale_date, sku_number, city) DO UPDATE SET
            units_sold = excluded.units_sold,
            mrp = excluded.mrp,
            gmv = excluded.gmv,
            sku_name = excluded.sku_name,
            ean = excluded.ean,
            sku_category = excluded.sku_category,
            sku_sub_category = excluded.sku_sub_category,
            brand_name = excluded.brand_name,
            manufacturer_name = excluded.manufacturer_name,
            matched_sku = excluded.matched_sku,
            matched_product_name = excluded.matched_product_name,
            source_file = excluded.source_file,
            uploaded_by = excluded.uploaded_by,
            created_at = CURRENT_TIMESTAMP
        ''', (
            norm_date, sku_num, sku_name, ean, cat, subcat,
            brand, mfg_name, mfg_id, city,
            units, mrp, gmv, matched_sku, matched_name, filename, user_info
        ))
        upserted_count += 1
        
    # Idempotent anti-overlap sync to central sales_data:
    for d in dates_seen:
        cur.execute('DELETE FROM sales_data WHERE platform = ? AND sale_date = ?', ('Zepto', d))
        cur.execute('''
        INSERT INTO sales_data (
            sale_date, platform, sku, product_name, brand, units_sold, revenue, store_location, source_file, uploaded_by
        )
        SELECT sale_date, 'Zepto', matched_sku, matched_product_name, brand_name, SUM(units_sold), SUM(gmv), city, source_file, uploaded_by
        FROM zepto_sales_orders
        WHERE sale_date = ?
        GROUP BY sale_date, matched_sku, city
        ''', (d,))
        
    return {
        'status': 'success',
        'platform': 'Zepto',
        'total_rows_in_file': total_in_file,
        'upserted_records': upserted_count,
        'total_units_sold': total_units,
        'total_gmv': round(total_gmv, 2),
        'unique_cities': len(cities_seen),
        'dates': sorted(list(dates_seen)),
        'date_range': f"{min(dates_seen)} to {max(dates_seen)}" if dates_seen else '-'
    }

def parse_instamart_sales_data(file_content, filename='Instamart_Sales.csv', user_info='Admin', cur=None):
    """
    Parses Swiggy Instamart Daily Product Sales Report (.xlsx / .csv).
    Also supports Campaign Ad Report as fallback.
    Deduplicates strictly using UNIQUE(sale_date, item_code, city, store_id) with ON CONFLICT DO UPDATE.
    Guarantees zero overlap when same dates are re-uploaded.
    """
    records_to_process = []
    is_excel = filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')
    
    if is_excel and isinstance(file_content, bytes):
        wb = openpyxl.load_workbook(BytesIO(file_content), data_only=True)
        sheet = wb['Sales Report'] if 'Sales Report' in wb.sheetnames else wb.active
        
        header_row = -1
        headers = {}
        for r in range(1, 25):
            row_vals = [sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)]
            str_vals = [str(v).strip().upper() for v in row_vals if v is not None]
            if any(h in str_vals for h in ['ORDERED_DATE', 'ORDER_DATE', 'UNITS_SOLD', 'ITEM_CODE', 'CAMPAIGN_ID']):
                header_row = r
                headers = {str(sheet.cell(row=r, column=c).value).strip().upper(): c for c in range(1, sheet.max_column + 1) if sheet.cell(row=r, column=c).value}
                break
                
        if header_row == -1:
            return {'status': 'error', 'error': 'Could not find Instamart table header (expected ORDERED_DATE, UNITS_SOLD, or ITEM_CODE)'}
            
        for r in range(header_row + 1, sheet.max_row + 1):
            def get_cell(patterns, default=''):
                for pat in patterns:
                    if pat in headers:
                        v = sheet.cell(row=r, column=headers[pat]).value
                        if v is not None:
                            return v
                for k, col in headers.items():
                    for pat in patterns:
                        if pat in k:
                            v = sheet.cell(row=r, column=col).value
                            if v is not None:
                                return v
                return default
                
            rec = {
                'date': normalize_date_str(get_cell(['ORDERED_DATE', 'ORDER_DATE', 'SALE_DATE', 'DATE', 'CAMPAIGN_START_DATE'])),
                'brand': str(get_cell(['BRAND_NAME', 'BRAND'], '')).strip(),
                'city': str(get_cell(['CITY'], 'Unknown')).strip(),
                'area': str(get_cell(['AREA_NAME', 'AREA', 'POD'], '')).strip(),
                'store': str(get_cell(['STORE_ID', 'POD_ID', 'STORE'], '')).strip(),
                'product_name': str(get_cell(['PRODUCT_NAME', 'ITEM_NAME', 'CAMPAIGN_NAME'], '')).strip(),
                'variant': str(get_cell(['VARIANT', 'SKU_NAME'], '')).strip(),
                'item_code': str(get_cell(['ITEM_CODE', 'SKU_NUMBER', 'SKU', 'CAMPAIGN_ID'], '')).strip(),
                'units': get_cell(['UNITS_SOLD', 'TOTAL_CONVERSIONS', 'QUANTITY', 'QTY'], 0),
                'mrp': get_cell(['BASE_MRP', 'MRP'], 0.0),
                'gmv': get_cell(['TOTAL_GMV', 'GMV', 'REVENUE', 'TOTAL_AMOUNT'], 0.0)
            }
            if rec['date'] and (rec['item_code'] or rec['product_name']):
                records_to_process.append(rec)
    else:
        if isinstance(file_content, bytes):
            file_content = file_content.decode('utf-8', errors='ignore')
        lines = file_content.splitlines() if isinstance(file_content, str) else file_content
        
        default_date = ''
        for l in lines[:15]:
            if 'From Date' in l:
                parts = [p.strip().strip('"') for p in l.split(',')]
                if len(parts) >= 2 and parts[1]:
                    default_date = normalize_date_str(parts[1])
                    
        header_idx = -1
        for i, l in enumerate(lines[:25]):
            up_l = l.upper()
            if any(h in up_l for h in ['ORDERED_DATE', 'ORDER_DATE', 'UNITS_SOLD', 'ITEM_CODE', 'CAMPAIGN_ID']):
                header_idx = i
                break
                
        if header_idx == -1:
            return {'status': 'error', 'error': 'Could not find Instamart table header (expected ORDERED_DATE, UNITS_SOLD, or ITEM_CODE)'}
            
        reader = csv.DictReader(lines[header_idx:])
        for r in reader:
            clean_r = {str(k).strip().upper(): v for k, v in r.items() if k}
            
            def get_r_val(patterns, default=''):
                for pat in patterns:
                    if pat in clean_r and clean_r[pat] is not None:
                        return clean_r[pat]
                for k, v in clean_r.items():
                    for pat in patterns:
                        if pat in k and v is not None:
                            return v
                return default
                
            d_val = normalize_date_str(get_r_val(['ORDERED_DATE', 'ORDER_DATE', 'SALE_DATE', 'DATE', 'CAMPAIGN_START_DATE'])) or default_date
            rec = {
                'date': d_val,
                'brand': str(get_r_val(['BRAND_NAME', 'BRAND'], '')).strip(),
                'city': str(get_r_val(['CITY'], 'Unknown')).strip(),
                'area': str(get_r_val(['AREA_NAME', 'AREA', 'POD'], '')).strip(),
                'store': str(get_r_val(['STORE_ID', 'POD_ID', 'STORE'], '')).strip(),
                'product_name': str(get_r_val(['PRODUCT_NAME', 'ITEM_NAME', 'CAMPAIGN_NAME'], '')).strip(),
                'variant': str(get_r_val(['VARIANT', 'SKU_NAME'], '')).strip(),
                'item_code': str(get_r_val(['ITEM_CODE', 'SKU_NUMBER', 'SKU', 'CAMPAIGN_ID'], '')).strip(),
                'units': get_r_val(['UNITS_SOLD', 'TOTAL_CONVERSIONS', 'QUANTITY', 'QTY'], 0),
                'mrp': get_r_val(['BASE_MRP', 'MRP'], 0.0),
                'gmv': get_r_val(['TOTAL_GMV', 'GMV', 'REVENUE', 'TOTAL_AMOUNT'], 0.0)
            }
            if rec['date'] and (rec['item_code'] or rec['product_name']):
                records_to_process.append(rec)

    total_in_file = len(records_to_process)
    upserted_count = 0
    dates_seen = set()
    total_units = 0
    total_gmv = 0.0
    cities_seen = set()
    
    for r in records_to_process:
        sale_date = r['date']
        item_code = r['item_code'] or 'INSTA-ITEM'
        p_name = r['product_name'] or 'Instamart Item'
        variant = r['variant']
        brand_name = r['brand']
        city = r['city'] or 'Unknown'
        area_name = r['area']
        store_id = r['store']
        
        try:
            units = int(float(str(r['units']).replace(',', '').strip() or 0))
        except:
            units = 0
        try:
            mrp = float(str(r['mrp']).replace(',', '').strip() or 0.0)
        except:
            mrp = 0.0
        try:
            gmv = float(str(r['gmv']).replace(',', '').strip() or 0.0)
        except:
            gmv = 0.0
            
        dates_seen.add(sale_date)
        cities_seen.add(city)
        total_units += units
        total_gmv += gmv
        
        matched_sku, matched_name, matched_brand = map_product_to_internal(item_code, p_name, None, cur)
        if not brand_name or brand_name.lower() == 'california nutralife wellness private limited':
            brand_name = matched_brand
            
        cur.execute('''
        INSERT INTO instamart_sales_orders (
            sale_date, item_code, product_name, variant, brand_name,
            city, area_name, store_id, units_sold, mrp, gmv,
            matched_sku, matched_product_name, source_file, uploaded_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sale_date, item_code, city, store_id) DO UPDATE SET
            units_sold = excluded.units_sold,
            mrp = excluded.mrp,
            gmv = excluded.gmv,
            product_name = excluded.product_name,
            variant = excluded.variant,
            brand_name = excluded.brand_name,
            area_name = excluded.area_name,
            matched_sku = excluded.matched_sku,
            matched_product_name = excluded.matched_product_name,
            source_file = excluded.source_file,
            uploaded_by = excluded.uploaded_by,
            created_at = CURRENT_TIMESTAMP
        ''', (
            sale_date, item_code, p_name, variant, brand_name,
            city, area_name, store_id, units, mrp, gmv,
            matched_sku, matched_name, filename, user_info
        ))
        upserted_count += 1
        
    # Idempotent anti-overlap sync to central sales_data:
    for d in dates_seen:
        cur.execute('DELETE FROM sales_data WHERE platform = ? AND sale_date = ?', ('Instamart', d))
        cur.execute('''
        INSERT INTO sales_data (
            sale_date, platform, sku, product_name, brand, units_sold, revenue, store_location, source_file, uploaded_by
        )
        SELECT sale_date, 'Instamart', matched_sku, matched_product_name, brand_name, SUM(units_sold), SUM(gmv), city, 'Instamart Sales Report', ?
        FROM instamart_sales_orders
        WHERE sale_date = ?
        GROUP BY sale_date, matched_sku, city
        ''', (user_info, d))
        
    return {
        'status': 'success',
        'platform': 'Instamart',
        'total_rows_in_file': total_in_file,
        'upserted_records': upserted_count,
        'total_units_sold': total_units,
        'total_gmv': round(total_gmv, 2),
        'unique_cities': len(cities_seen),
        'dates': sorted(list(dates_seen)),
        'date_range': f"{min(dates_seen)} to {max(dates_seen)}" if dates_seen else '-'
    }

BIGBASKET_INITIAL_SEED_CSV = '''date_range,source_city_name,business_type,brand_slug,top_slug,mid_slug,leaf_slug,source_sku_id,sku_description,sku_weight,Total_quantity,Total_mrp,Total_sales
20260927 - 20260927,Noida,b2c,nutracookies,snacks-branded-foods,biscuits-cookies,cookies,40371097,Oats Cookies - Sugar Free,100 g,3.0,357.0,296.31
20260927 - 20260927,Bangalore,b2c,nutracookies,snacks-branded-foods,biscuits-cookies,cookies,40371097,Oats Cookies - Sugar Free,100 g,11.0,1309.0,1086.47
20260927 - 20260927,Bangalore,b2c,nutrabites,snacks-branded-foods,snacks-namkeen,namkeen-savoury-snacks,40371096,High Protein Baked Bhujia,75 g,1.0,99.0,85.14
20260927 - 20260927,Bangalore,b2c,california-skin,beauty-hygiene,skin-care,face-care,40370930,Triple Action Acne Relief Pimple Patches,36 pcs,3.0,855.0,444.6
20260927 - 20260927,Gurgaon,b2c,california-skin,beauty-hygiene,skin-care,face-care,40370930,Triple Action Acne Relief Pimple Patches,36 pcs,5.0,1425.0,741.0
20260927 - 20260927,Gurgaon,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371094,Rock Salt Banana Chips,50 g,9.0,441.0,352.8
20260927 - 20260927,Gurgaon,b2c,nutracookies,snacks-branded-foods,biscuits-cookies,cookies,40371097,Oats Cookies - Sugar Free,100 g,7.0,833.0,691.39
20260927 - 20260927,Bangalore,bbdaily,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371095,Baked Ragi Chips,50 g,1.0,59.0,48.97
20260927 - 20260927,Gurgaon,b2c,nutrabites,snacks-branded-foods,snacks-namkeen,namkeen-savoury-snacks,40371096,High Protein Baked Bhujia,75 g,4.0,396.0,340.56
20260927 - 20260927,Bangalore,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371094,Rock Salt Banana Chips,50 g,4.0,196.0,156.8
20260927 - 20260927,Noida,b2c,california-skin,beauty-hygiene,skin-care,face-care,40370930,Triple Action Acne Relief Pimple Patches,36 pcs,3.0,855.0,444.6
20260927 - 20260927,Noida,b2c,nutrabites,snacks-branded-foods,snacks-namkeen,namkeen-savoury-snacks,40371096,High Protein Baked Bhujia,75 g,1.0,99.0,85.14
20260927 - 20260927,Mumbai,b2c,nutrabites,snacks-branded-foods,snacks-namkeen,namkeen-savoury-snacks,40371096,High Protein Baked Bhujia,75 g,2.0,198.0,170.28
20260927 - 20260927,Noida,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371095,Baked Ragi Chips,50 g,2.0,118.0,97.94
20260927 - 20260927,Noida,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371094,Rock Salt Banana Chips,50 g,2.0,98.0,78.4
20260927 - 20260927,Bangalore,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371095,Baked Ragi Chips,50 g,6.0,344.0,285.52
20260927 - 20260927,Gurgaon,b2c,nutrachips,snacks-branded-foods,snacks-namkeen,chips-corn-snacks,40371095,Baked Ragi Chips,50 g,5.0,295.0,244.85'''

def parse_bigbasket_sales_data(file_content, filename='BigBasket_Sales.csv', user_info='Admin', cur=None):
    """
    Parses BigBasket Daily Sales CSV/Excel reports.
    Deduplicates strictly using UNIQUE(sale_date, source_sku_id, city, business_type) with ON CONFLICT DO UPDATE.
    Guarantees zero overlap when same dates are re-uploaded.
    """
    if isinstance(file_content, bytes):
        if filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls'):
            wb = openpyxl.load_workbook(BytesIO(file_content), data_only=True)
            sheet = wb.active
            lines = []
            for row in sheet.iter_rows(values_only=True):
                lines.append(','.join(['' if v is None else str(v).replace(',', ' ') for v in row]))
            file_content = '\n'.join(lines)
        else:
            try:
                file_content = file_content.decode('utf-8-sig')
            except UnicodeDecodeError:
                file_content = file_content.decode('latin1', errors='replace')

    lines = [l.strip() for l in file_content.splitlines() if l.strip()]
    if not lines:
        return {'status': 'error', 'message': 'The uploaded file is empty.'}

    # Find header row
    header_idx = 0
    for idx, l in enumerate(lines[:10]):
        low = l.lower()
        if 'source_sku_id' in low or 'sku_description' in low or 'total_quantity' in low or 'total_sales' in low:
            header_idx = idx
            break

    reader = csv.DictReader(lines[header_idx:])
    total_in_file = 0
    upserted_count = 0
    dates_seen = set()
    total_units = 0
    total_sales_val = 0.0
    cities_seen = set()

    def get_val(r_dict, keys, default=''):
        for k, v in r_dict.items():
            if k and v is not None:
                clean_k = k.lower().strip().replace(' ', '_')
                for target in keys:
                    if clean_k == target or target in clean_k:
                        return v
        return default

    for r in reader:
        raw_date = str(get_val(r, ['date_range', 'sale_date', 'date'], '')).strip()
        sku_id = str(get_val(r, ['source_sku_id', 'sku_id', 'sku_number', 'sku'], '')).strip()
        sku_desc = str(get_val(r, ['sku_description', 'product_name', 'item_name', 'description'], '')).strip()

        if not raw_date or (not sku_id and not sku_desc):
            continue

        norm_date = normalize_date_str(raw_date)
        if not norm_date:
            continue

        dates_seen.add(norm_date)
        total_in_file += 1

        city = str(get_val(r, ['source_city_name', 'city_name', 'city'], 'Unknown')).strip()
        cities_seen.add(city)
        biz_type = str(get_val(r, ['business_type', 'channel', 'type'], 'b2c')).strip().lower()
        brand_slug = str(get_val(r, ['brand_slug', 'brand'], '')).strip()
        top_slug = str(get_val(r, ['top_slug', 'category'], '')).strip()
        mid_slug = str(get_val(r, ['mid_slug', 'subcategory'], '')).strip()
        leaf_slug = str(get_val(r, ['leaf_slug'], '')).strip()
        sku_weight = str(get_val(r, ['sku_weight', 'weight', 'pack_size'], '')).strip()

        try:
            units = int(float(str(get_val(r, ['total_quantity', 'quantity', 'units', 'qty'], 0)).replace(',', '').strip() or 0))
        except Exception:
            units = 0

        try:
            mrp = float(str(get_val(r, ['total_mrp', 'mrp'], 0.0)).replace(',', '').strip() or 0.0)
        except Exception:
            mrp = 0.0

        try:
            sales = float(str(get_val(r, ['total_sales', 'sales', 'revenue', 'gmv'], 0.0)).replace(',', '').strip() or 0.0)
        except Exception:
            sales = 0.0

        total_units += units
        total_sales_val += sales

        matched_sku, matched_name, matched_brand = map_product_to_internal(sku_id, sku_desc, None, cur)
        if not brand_slug:
            brand_slug = matched_brand

        cur.execute('''
        INSERT INTO bigbasket_sales_orders (
            sale_date, source_sku_id, sku_description, sku_weight, brand_slug,
            city, business_type, top_slug, mid_slug, leaf_slug,
            units_sold, mrp, sales_amount, matched_sku, matched_product_name, source_file, uploaded_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sale_date, source_sku_id, city, business_type) DO UPDATE SET
            units_sold = excluded.units_sold,
            mrp = excluded.mrp,
            sales_amount = excluded.sales_amount,
            sku_description = excluded.sku_description,
            sku_weight = excluded.sku_weight,
            brand_slug = excluded.brand_slug,
            top_slug = excluded.top_slug,
            mid_slug = excluded.mid_slug,
            leaf_slug = excluded.leaf_slug,
            matched_sku = excluded.matched_sku,
            matched_product_name = excluded.matched_product_name,
            source_file = excluded.source_file,
            uploaded_by = excluded.uploaded_by,
            created_at = CURRENT_TIMESTAMP
        ''', (
            norm_date, sku_id, sku_desc, sku_weight, brand_slug,
            city, biz_type, top_slug, mid_slug, leaf_slug,
            units, mrp, sales, matched_sku, matched_name, filename, user_info
        ))
        upserted_count += 1

    # Idempotent anti-overlap sync to central sales_data:
    for d in dates_seen:
        cur.execute('DELETE FROM sales_data WHERE platform = ? AND sale_date = ?', ('BigBasket', d))
        cur.execute('''
        INSERT INTO sales_data (
            sale_date, platform, sku, product_name, brand, units_sold, revenue, store_location, source_file, uploaded_by
        )
        SELECT sale_date, 'BigBasket', matched_sku, matched_product_name, 
               CASE 
                   WHEN brand_slug = 'california-skin' THEN 'California Skin+'
                   WHEN brand_slug = 'nutracookies' THEN 'NutraCookies'
                   WHEN brand_slug = 'nutrabites' THEN 'NutraBites'
                   WHEN brand_slug = 'nutrachips' THEN 'NutraChips'
                   ELSE brand_slug END,
               SUM(units_sold), SUM(sales_amount), city, ?, ?
        FROM bigbasket_sales_orders
        WHERE sale_date = ?
        GROUP BY sale_date, matched_sku, city
        ''', (filename, user_info, d))

    return {
        'status': 'success',
        'platform': 'BigBasket',
        'total_rows_in_file': total_in_file,
        'upserted_records': upserted_count,
        'total_units_sold': total_units,
        'total_sales_amount': round(total_sales_val, 2),
        'unique_cities': len(cities_seen),
        'dates': sorted(list(dates_seen)),
        'date_range': f"{min(dates_seen)} to {max(dates_seen)}" if dates_seen else '-'
    }

def seed_initial_quick_commerce_data_if_empty():
    """Auto-seeds initial Zepto, Instamart & BigBasket sales data if tables are empty."""
    try:
        conn = get_db()
        cur = conn.cursor()
        brain_uploads = os.path.join(os.path.expanduser('~'), '.gemini', 'antigravity', 'brain', 'c18c0bb2-1dda-4e5b-89b7-d36de3bd9d79', '.user_uploaded')
        
        # Zepto seed
        z_count = cur.execute('SELECT COUNT(*) FROM zepto_sales_orders').fetchone()[0]
        if z_count == 0:
            z_path = os.path.join(brain_uploads, 'media_1790524533163.csv')
            if os.path.exists(z_path):
                with open(z_path, 'rb') as f:
                    parse_zepto_sales_data(f.read(), 'Zepto_25-09-2026.csv', 'System Initial Seed', cur)
                print("Auto-seeded Zepto sales dataset successfully.")
                
        # Instamart seed
        i_count = cur.execute('SELECT COUNT(*) FROM instamart_sales_orders').fetchone()[0]
        if i_count == 0:
            i_path = os.path.join(brain_uploads, 'media_1790525376736.csv')
            if os.path.exists(i_path):
                with open(i_path, 'rb') as f:
                    parse_instamart_sales_data(f.read(), 'Instamart_25-09-2026.csv', 'System Initial Seed', cur)
                print("Auto-seeded Instamart sales dataset successfully.")

        # BigBasket seed
        bb_count = cur.execute('SELECT COUNT(*) FROM bigbasket_sales_orders').fetchone()[0]
        if bb_count == 0:
            parse_bigbasket_sales_data(BIGBASKET_INITIAL_SEED_CSV, 'BigBasket_Sales_27-09-2026.csv', 'System Initial Seed', cur)
            print("Auto-seeded BigBasket sales dataset successfully.")
                
        conn.commit()
        conn.close()
    except Exception as e:
        print("Quick-commerce seed notice:", e)

seed_initial_quick_commerce_data_if_empty()


@app.route('/api/blinkit/upload-inventory', methods=['POST'])
def upload_blinkit_inventory():
    user_email, user_name = get_user_info()
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    f = request.files['file']
    filename = (f.filename or '').lower()
    content = f.read()
    
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
        sheet = wb.active
        lines = []
        for row in sheet.iter_rows(values_only=True):
            lines.append(','.join(['' if v is None else str(v).replace(',', ' ') for v in row]))
        content = '\n'.join(lines)
        
    conn = get_db()
    cur = conn.cursor()
    res = parse_blinkit_inventory_data(content, cur)
    if 'error' in res:
        conn.close()
        return jsonify(res), 400
        
    conn.commit()
    record_audit(conn, 'UPLOAD_BLINKIT_INVENTORY', 'Blinkit Inventory', f.filename, f"Processed {res['rows_processed']} inventory records across {res['facilities_count']} facilities by {user_name}")
    conn.close()
    return jsonify(res)

@app.route('/api/blinkit/upload-sales', methods=['POST'])
def upload_blinkit_sales():
    user_email, user_name = get_user_info()
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    f = request.files['file']
    filename = (f.filename or '').lower()
    content = f.read()
    
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
        sheet = wb.active
        lines = []
        for row in sheet.iter_rows(values_only=True):
            lines.append(','.join(['' if v is None else str(v).replace(',', ' ') for v in row]))
        content = '\n'.join(lines)
        
    conn = get_db()
    cur = conn.cursor()
    res = parse_blinkit_sales_data(content, cur)
    conn.commit()
    record_audit(conn, 'UPLOAD_BLINKIT_SALES', 'Blinkit Sales Orders', f.filename, f"Ingested {res['new_orders_inserted']} new orders, skipped {res['existing_orders_skipped']} duplicates ({res['date_range']}) by {user_name}")
    conn.close()
    return jsonify(res)

# ==============================================================================
# QUICK-COMMERCE MULTI-PLATFORM API ENDPOINTS (ZEPTO, INSTAMART, BLINKIT & HUB)
# ==============================================================================

@app.route('/api/zepto/upload-sales', methods=['POST'])
def upload_zepto_sales():
    user_email, user_name = get_user_info()
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    f = request.files['file']
    filename = (f.filename or 'Zepto_Sales.csv')
    content = f.read()
    
    conn = get_db()
    cur = conn.cursor()
    res = parse_zepto_sales_data(content, filename, f"{user_name} ({user_email})", cur)
    conn.commit()
    record_audit(conn, 'UPLOAD_ZEPTO_SALES', 'Zepto Sales', filename, f"Ingested/updated {res['upserted_records']} Zepto sales rows ({res['total_units_sold']} units, ₹{res['total_gmv']:,.2f}) for dates: {res['date_range']} by {user_name}")
    conn.close()
    return jsonify(res)

@app.route('/api/zepto/sales-summary', methods=['GET'])
def get_zepto_sales_summary():
    conn = get_db()
    cur = conn.cursor()
    
    city_filter = request.args.get('city', 'All').strip()
    date_filter = request.args.get('date', 'All').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    where_clauses = []
    params = []
    if city_filter != 'All':
        where_clauses.append('city = ?')
        params.append(city_filter)
    if start_date:
        where_clauses.append('sale_date >= ?')
        params.append(start_date)
    if end_date:
        where_clauses.append('sale_date <= ?')
        params.append(end_date)
    elif date_filter and date_filter != 'All':
        where_clauses.append('sale_date = ?')
        params.append(date_filter)
        
    where_str = (' WHERE ' + ' AND '.join(where_clauses)) if where_clauses else ''
    
    kpi_row = cur.execute(f'''
    SELECT 
        COUNT(*) as total_rows,
        COALESCE(SUM(units_sold), 0) as total_units,
        COALESCE(SUM(gmv), 0.0) as total_gmv,
        COUNT(DISTINCT city) as total_cities,
        COUNT(DISTINCT sku_number) as total_skus,
        MIN(sale_date) as min_date,
        MAX(sale_date) as max_date
    FROM zepto_sales_orders {where_str}
    ''', params).fetchone()
    
    city_rows = cur.execute(f'''
    SELECT 
        city,
        SUM(units_sold) as units,
        SUM(gmv) as gmv,
        COUNT(DISTINCT sku_number) as skus_count
    FROM zepto_sales_orders {where_str}
    GROUP BY city
    ORDER BY units DESC
    LIMIT 30
    ''', params).fetchall()
    
    sku_rows = cur.execute(f'''
    SELECT 
        matched_sku,
        matched_product_name,
        brand_name,
        SUM(units_sold) as units,
        SUM(gmv) as gmv
    FROM zepto_sales_orders {where_str}
    GROUP BY matched_sku, matched_product_name, brand_name
    ORDER BY units DESC
    ''', params).fetchall()
    
    all_cities = [r[0] for r in cur.execute('SELECT DISTINCT city FROM zepto_sales_orders ORDER BY city').fetchall()]
    all_dates = [r[0] for r in cur.execute('SELECT DISTINCT sale_date FROM zepto_sales_orders ORDER BY sale_date DESC').fetchall()]
    
    records = cur.execute(f'''
    SELECT 
        id, sale_date, sku_number, sku_name, ean, city, units_sold, mrp, gmv,
        matched_sku, matched_product_name, brand_name, source_file, created_at
    FROM zepto_sales_orders {where_str}
    ORDER BY sale_date DESC, id DESC
    LIMIT 200
    ''', params).fetchall()
    
    conn.close()
    return jsonify({
        'status': 'success',
        'kpis': dict(kpi_row) if kpi_row else {},
        'cities': [dict(r) for r in city_rows],
        'skus': [dict(r) for r in sku_rows],
        'filter_options': {
            'cities': all_cities,
            'dates': all_dates
        },
        'records': [dict(r) for r in records]
    })

@app.route('/api/instamart/upload-sales', methods=['POST'])
def upload_instamart_sales():
    user_email, user_name = get_user_info()
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    f = request.files['file']
    filename = (f.filename or 'Instamart_Sales.csv')
    content = f.read()
    
    conn = get_db()
    cur = conn.cursor()
    res = parse_instamart_sales_data(content, filename, f"{user_name} ({user_email})", cur)
    if res.get('status') == 'error':
        conn.close()
        return jsonify(res), 400
        
    conn.commit()
    record_audit(conn, 'UPLOAD_INSTAMART_SALES', 'Instamart Sales', filename, f"Ingested/updated {res['upserted_records']} Instamart sales records ({res['total_units_sold']} units, ₹{res['total_gmv']:,.2f} GMV) for dates: {res['date_range']} by {user_name}")
    conn.close()
    return jsonify(res)

@app.route('/api/instamart/sales-summary', methods=['GET'])
def get_instamart_sales_summary():
    conn = get_db()
    cur = conn.cursor()
    
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    date_filter = request.args.get('date', 'All').strip()
    
    where_clauses = []
    params = []
    if start_date:
        where_clauses.append('sale_date >= ?')
        params.append(start_date)
    if end_date:
        where_clauses.append('sale_date <= ?')
        params.append(end_date)
    elif date_filter and date_filter != 'All':
        where_clauses.append('sale_date = ?')
        params.append(date_filter)
        
    where_str = (' WHERE ' + ' AND '.join(where_clauses)) if where_clauses else ''
    
    kpi_row = cur.execute(f'''
    SELECT 
        COUNT(*) as total_rows,
        COALESCE(SUM(units_sold), 0) as total_units,
        COALESCE(SUM(gmv), 0.0) as total_gmv,
        COUNT(DISTINCT matched_sku) as total_products,
        MIN(sale_date) as min_date,
        MAX(sale_date) as max_date
    FROM instamart_sales_orders {where_str}
    ''', params).fetchone()
    
    kpis = dict(kpi_row) if kpi_row else {}
    
    # Clean product sales ledger
    prod_rows = cur.execute(f'''
    SELECT 
        id,
        sale_date,
        matched_sku,
        COALESCE(matched_product_name, product_name) as product_name,
        variant,
        brand_name,
        city,
        area_name,
        units_sold,
        mrp,
        gmv,
        source_file,
        created_at
    FROM instamart_sales_orders {where_str}
    ORDER BY sale_date DESC, gmv DESC
    LIMIT 250
    ''', params).fetchall()
    
    conn.close()
    return jsonify({
        'status': 'success',
        'kpis': kpis,
        'records': [dict(r) for r in prod_rows]
    })

@app.route('/api/bigbasket/upload-sales', methods=['POST'])
def upload_bigbasket_sales():
    user_email, user_name = get_user_info()
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    f = request.files['file']
    filename = (f.filename or 'BigBasket_Sales.csv')
    content = f.read()
    
    conn = get_db()
    cur = conn.cursor()
    res = parse_bigbasket_sales_data(content, filename, f"{user_name} ({user_email})", cur)
    if res.get('status') == 'error':
        conn.close()
        return jsonify(res), 400
        
    conn.commit()
    record_audit(conn, 'UPLOAD_BIGBASKET_SALES', 'BigBasket Sales', filename, f"Ingested/updated {res['upserted_records']} BigBasket sales records ({res['total_units_sold']} units, ₹{res['total_sales_amount']:,.2f} Sales) for dates: {res['date_range']} by {user_name}")
    conn.close()
    return jsonify(res)

@app.route('/api/bigbasket/sales-summary', methods=['GET'])
def get_bigbasket_sales_summary():
    conn = get_db()
    cur = conn.cursor()
    
    city_filter = request.args.get('city', 'All').strip()
    date_filter = request.args.get('date', 'All').strip()
    biz_filter = request.args.get('business_type', 'All').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    where_clauses = []
    params = []
    if city_filter != 'All':
        where_clauses.append('city = ?')
        params.append(city_filter)
    if biz_filter != 'All':
        where_clauses.append('business_type = ?')
        params.append(biz_filter)
    if start_date:
        where_clauses.append('sale_date >= ?')
        params.append(start_date)
    if end_date:
        where_clauses.append('sale_date <= ?')
        params.append(end_date)
    elif date_filter and date_filter != 'All':
        where_clauses.append('sale_date = ?')
        params.append(date_filter)
        
    where_str = (' WHERE ' + ' AND '.join(where_clauses)) if where_clauses else ''
    
    kpi_row = cur.execute(f'''
    SELECT 
        COUNT(*) as total_rows,
        COALESCE(SUM(units_sold), 0) as total_units,
        COALESCE(SUM(sales_amount), 0.0) as total_sales,
        COUNT(DISTINCT city) as total_cities,
        COUNT(DISTINCT source_sku_id) as total_skus,
        MIN(sale_date) as min_date,
        MAX(sale_date) as max_date
    FROM bigbasket_sales_orders {where_str}
    ''', params).fetchone()
    
    city_rows = cur.execute(f'''
    SELECT 
        city,
        SUM(units_sold) as units,
        SUM(sales_amount) as sales,
        COUNT(DISTINCT source_sku_id) as skus_count
    FROM bigbasket_sales_orders {where_str}
    GROUP BY city
    ORDER BY sales DESC
    LIMIT 30
    ''', params).fetchall()
    
    sku_rows = cur.execute(f'''
    SELECT 
        matched_sku,
        matched_product_name,
        brand_slug,
        SUM(units_sold) as units,
        SUM(sales_amount) as sales
    FROM bigbasket_sales_orders {where_str}
    GROUP BY matched_sku, matched_product_name, brand_slug
    ORDER BY units DESC
    ''', params).fetchall()
    
    all_cities = [r[0] for r in cur.execute('SELECT DISTINCT city FROM bigbasket_sales_orders ORDER BY city').fetchall()]
    all_dates = [r[0] for r in cur.execute('SELECT DISTINCT sale_date FROM bigbasket_sales_orders ORDER BY sale_date DESC').fetchall()]
    all_biz = [r[0] for r in cur.execute('SELECT DISTINCT business_type FROM bigbasket_sales_orders ORDER BY business_type').fetchall()]
    
    records = cur.execute(f'''
    SELECT 
        id, sale_date, source_sku_id, sku_description, sku_weight, brand_slug,
        city, business_type, top_slug, mid_slug, leaf_slug, units_sold, mrp, sales_amount,
        matched_sku, matched_product_name, source_file, created_at
    FROM bigbasket_sales_orders {where_str}
    ORDER BY sale_date DESC, id DESC
    LIMIT 250
    ''', params).fetchall()
    
    conn.close()
    return jsonify({
        'status': 'success',
        'kpis': dict(kpi_row) if kpi_row else {},
        'cities': [dict(r) for r in city_rows],
        'skus': [dict(r) for r in sku_rows],
        'filter_options': {
            'cities': all_cities,
            'dates': all_dates,
            'business_types': all_biz
        },
        'records': [dict(r) for r in records]
    })

@app.route('/api/blinkit/sales-tab-summary', methods=['GET'])
def get_blinkit_sales_tab_summary():
    conn = get_db()
    cur = conn.cursor()
    
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    where_clauses = []
    params = []
    if start_date:
        where_clauses.append('order_date >= ?')
        params.append(start_date)
    if end_date:
        where_clauses.append('order_date <= ?')
        params.append(end_date)
        
    where_str = (' WHERE ' + ' AND '.join(where_clauses)) if where_clauses else ''
    
    kpi_row = cur.execute(f'''
    SELECT 
        COUNT(*) as total_orders,
        COALESCE(SUM(quantity), 0) as total_units,
        COALESCE(SUM(total_gross_amount), 0.0) as total_revenue,
        COUNT(DISTINCT supply_city) as unique_supply_cities,
        COUNT(DISTINCT customer_city) as unique_customer_cities,
        COUNT(DISTINCT item_id) as unique_items,
        MIN(order_date) as min_date,
        MAX(order_date) as max_date
    FROM blinkit_sales_orders {where_str}
    ''', params).fetchone()
    
    city_rows = cur.execute(f'''
    SELECT 
        supply_city as city,
        COUNT(*) as order_count,
        SUM(quantity) as units,
        SUM(total_gross_amount) as revenue
    FROM blinkit_sales_orders {where_str}
    GROUP BY supply_city
    ORDER BY revenue DESC
    LIMIT 20
    ''', params).fetchall()
    
    product_rows = cur.execute(f'''
    SELECT 
        product_name,
        brand_name,
        COUNT(*) as order_count,
        SUM(quantity) as units,
        SUM(total_gross_amount) as revenue
    FROM blinkit_sales_orders {where_str}
    GROUP BY product_name, brand_name
    ORDER BY units DESC
    ''', params).fetchall()
    
    recent_rows = cur.execute(f'''
    SELECT 
        order_id, order_date, product_name, brand_name, supply_city, customer_city,
        quantity, mrp, selling_price, total_gross_amount, order_status
    FROM blinkit_sales_orders {where_str}
    ORDER BY order_date DESC, created_at DESC
    LIMIT 200
    ''', params).fetchall()
    
    conn.close()
    return jsonify({
        'status': 'success',
        'kpis': dict(kpi_row) if kpi_row else {},
        'cities': [dict(r) for r in city_rows],
        'products': [dict(r) for r in product_rows],
        'recent_orders': [dict(r) for r in recent_rows]
    })

@app.route('/api/sales/hub-summary', methods=['GET'])
def get_sales_hub_summary():
    conn = get_db()
    cur = conn.cursor()
    
    blk = cur.execute('''
    SELECT COUNT(*) as orders, COALESCE(SUM(quantity), 0) as units, COALESCE(SUM(total_gross_amount), 0.0) as rev, MIN(order_date) as min_d, MAX(order_date) as max_d
    FROM blinkit_sales_orders
    ''').fetchone()
    
    zpt = cur.execute('''
    SELECT COUNT(*) as records, COALESCE(SUM(units_sold), 0) as units, COALESCE(SUM(gmv), 0.0) as rev, COUNT(DISTINCT city) as cities, MIN(sale_date) as min_d, MAX(sale_date) as max_d
    FROM zepto_sales_orders
    ''').fetchone()
    
    ins = cur.execute('''
    SELECT COUNT(*) as records, COALESCE(SUM(units_sold), 0) as units, COALESCE(SUM(gmv), 0.0) as rev, MIN(sale_date) as min_d, MAX(sale_date) as max_d
    FROM instamart_sales_orders
    ''').fetchone()

    bb = cur.execute('''
    SELECT COUNT(*) as records, COALESCE(SUM(units_sold), 0) as units, COALESCE(SUM(sales_amount), 0.0) as rev, COUNT(DISTINCT city) as cities, MIN(sale_date) as min_d, MAX(sale_date) as max_d
    FROM bigbasket_sales_orders
    ''').fetchone()
    
    cen = cur.execute('''
    SELECT COUNT(*) as rows, COALESCE(SUM(units_sold), 0) as units, COALESCE(SUM(revenue), 0.0) as rev
    FROM sales_data
    ''').fetchone()
    
    conn.close()
    return jsonify({
        'status': 'success',
        'blinkit': dict(blk) if blk else {},
        'zepto': dict(zpt) if zpt else {},
        'instamart': dict(ins) if ins else {},
        'bigbasket': dict(bb) if bb else {},
        'consolidated': dict(cen) if cen else {}
    })

@app.route('/api/blinkit/replenishment-model', methods=['GET'])
def get_blinkit_replenishment_model():
    conn = get_db()
    cur = conn.cursor()
    
    # Load settings
    settings_rows = cur.execute('SELECT * FROM blinkit_facility_settings').fetchall()
    facility_settings = {r['facility_name']: dict(r) for r in settings_rows}
    
    # Check filter queries
    filter_facility = request.args.get('facility', 'All').strip()
    filter_status = request.args.get('status', 'All').strip()
    filter_brand = request.args.get('brand', 'All').strip()
    
    query = 'SELECT * FROM blinkit_inventory_current'
    params = []
    clauses = []
    if filter_facility != 'All':
        clauses.append('facility_name = ?')
        params.append(filter_facility)
    if filter_brand != 'All':
        clauses.append('brand_name = ?')
        params.append(filter_brand)
    if clauses:
        query += ' WHERE ' + ' AND '.join(clauses)
    query += ' ORDER BY facility_name, item_name'
    
    inv_rows = cur.execute(query, tuple(params)).fetchall()
    
    # Global KPIs
    total_sellable_all = 0
    total_incoming_all = 0
    total_7d_sales_all = 0
    critical_count = 0
    reorder_due_count = 0
    healthy_count = 0
    overstocked_count = 0
    excess_capital_at_risk = 0.0
    
    model_rows = []
    
    for r in inv_rows:
        d = dict(r)
        fac_name = d['facility_name']
        fac_conf = facility_settings.get(fac_name, {
            'lead_time_days': 8,
            'safety_stock_days': 3,
            'target_max_days': 21
        })
        
        lead_time = int(fac_conf.get('lead_time_days', 8))
        safety_stock = int(fac_conf.get('safety_stock_days', 3))
        target_max_days = int(fac_conf.get('target_max_days', 21))
        
        stk = d['total_sellable'] or 0
        incom = d['incoming_scheduled'] or 0
        s7 = d['sales_7d'] or 0
        s30 = d['sales_30d'] or 0
        
        total_sellable_all += stk
        total_incoming_all += incom
        total_7d_sales_all += s7
        
        drr7 = s7 / 7.0
        drr30 = s30 / 30.0
        drr = round((0.7 * drr7 + 0.3 * drr30) if drr7 > 0 else drr30, 2)
        
        rop_units = round(drr * (lead_time + safety_stock), 1)
        target_max_units = round(drr * target_max_days, 1)
        eff_stk = stk + incom
        
        if drr > 0:
            current_doi = round(stk / drr, 1)
            pipeline_doi = round(eff_stk / drr, 1)
        else:
            current_doi = 999.0 if stk > 0 else 0.0
            pipeline_doi = 999.0 if eff_stk > 0 else 0.0
            
        if drr == 0:
            if stk == 0 and incom == 0:
                status = 'INACTIVE'
                status_label = 'No Sales / Inactive'
                rec_po = 0
            else:
                status = 'OVERSTOCKED'
                status_label = 'Dead Stock / High Storage Risk'
                rec_po = 0
                overstocked_count += 1
                excess_capital_at_risk += stk * 120.0
        elif pipeline_doi < lead_time:
            status = 'CRITICAL'
            status_label = f'🚨 Urgent Reorder (< {lead_time}d Delivery Window)'
            rec_po = int(max(0, math.ceil(target_max_units - eff_stk)))
            critical_count += 1
        elif pipeline_doi <= (lead_time + safety_stock):
            status = 'REORDER_DUE'
            status_label = f'⚠️ Reorder Due (At ROP {rop_units:.0f}u)'
            rec_po = int(max(0, math.ceil(target_max_units - eff_stk)))
            reorder_due_count += 1
        elif pipeline_doi <= target_max_days:
            status = 'HEALTHY'
            status_label = '🟢 Healthy Stock (Optimal)'
            rec_po = 0
            healthy_count += 1
        else:
            status = 'OVERSTOCKED'
            status_label = '🟣 Excess Stock (Storage Fee Risk)'
            rec_po = 0
            overstocked_count += 1
            excess_units = max(0, stk - target_max_units)
            excess_capital_at_risk += excess_units * 120.0
            
        d['drr'] = drr
        d['current_doi'] = current_doi
        d['pipeline_doi'] = pipeline_doi
        d['rop_units'] = rop_units
        d['target_max_units'] = target_max_units
        d['lead_time_days'] = lead_time
        d['safety_stock_days'] = safety_stock
        d['target_max_days'] = target_max_days
        d['status'] = status
        d['status_label'] = status_label
        d['recommended_po_qty'] = rec_po
        
        if filter_status == 'All' or filter_status == status:
            model_rows.append(d)
            
    all_facilities = [r[0] for r in cur.execute('SELECT DISTINCT facility_name FROM blinkit_inventory_current ORDER BY facility_name').fetchall()]
    all_brands = [r[0] for r in cur.execute('SELECT DISTINCT brand_name FROM blinkit_inventory_current ORDER BY brand_name').fetchall()]
    last_inv = cur.execute('SELECT MAX(updated_at) FROM blinkit_inventory_current').fetchone()
    last_inv_time = last_inv[0] if last_inv else None
    total_sales_records = cur.execute('SELECT COUNT(*), MIN(order_date), MAX(order_date) FROM blinkit_sales_orders').fetchone()
    
    conn.close()
    
    return jsonify({
        'status': 'success',
        'kpis': {
            'total_facilities': len(all_facilities),
            'total_items_monitored': len(inv_rows),
            'total_sellable_network': total_sellable_all,
            'total_incoming_scheduled': total_incoming_all,
            'total_weekly_sales': total_7d_sales_all,
            'critical_count': critical_count,
            'reorder_due_count': reorder_due_count,
            'healthy_count': healthy_count,
            'overstocked_count': overstocked_count,
            'excess_capital_at_risk': round(excess_capital_at_risk, 2),
            'last_inventory_sync': last_inv_time,
            'total_sales_orders': total_sales_records[0] or 0,
            'sales_date_range': f"{total_sales_records[1]} to {total_sales_records[2]}" if total_sales_records[1] else 'None'
        },
        'filters': {
            'facilities': all_facilities,
            'brands': all_brands
        },
        'rows': model_rows
    })

@app.route('/api/blinkit/settings', methods=['GET', 'POST'])
def manage_blinkit_settings():
    conn = get_db()
    cur = conn.cursor()
    
    if request.method == 'POST':
        user_email, user_name = get_user_info()
        profile = get_user_profile(user_email)
        if not (profile['is_admin'] or profile.get('can_create_po')):
            conn.close()
            return jsonify({'error': 'Unauthorized to change Blinkit settings'}), 403
            
        data = request.json or {}
        fac_name = str(data.get('facility_name', '')).strip()
        lead_time = int(data.get('lead_time_days', 8))
        safety_stock = int(data.get('safety_stock_days', 3))
        target_max = int(data.get('target_max_days', 21))
        
        if fac_name == 'ALL' or not fac_name:
            cur.execute('UPDATE blinkit_facility_settings SET lead_time_days = ?, safety_stock_days = ?, target_max_days = ?, updated_at = CURRENT_TIMESTAMP', (lead_time, safety_stock, target_max))
        else:
            cur.execute('''
            UPDATE blinkit_facility_settings 
            SET lead_time_days = ?, safety_stock_days = ?, target_max_days = ?, updated_at = CURRENT_TIMESTAMP
            WHERE facility_name = ?
            ''', (lead_time, safety_stock, target_max, fac_name))
            
        conn.commit()
        record_audit(conn, 'UPDATE_BLINKIT_SETTINGS', 'Blinkit Settings', fac_name or 'ALL', f"Set lead time: {lead_time}d, safety stock: {safety_stock}d, target max: {target_max}d by {user_name}")
        
    rows = cur.execute('SELECT * FROM blinkit_facility_settings ORDER BY facility_name').fetchall()
    settings_list = [dict(r) for r in rows]
    conn.close()
    return jsonify({'status': 'success', 'settings': settings_list})

if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8765))
    HOST = os.environ.get('HOST', '0.0.0.0' if os.environ.get('PORT') else '127.0.0.1')
    print(f"Starting Inventory Management System on http://{HOST}:{PORT}")
    app.run(host=HOST, port=PORT, debug=False)
