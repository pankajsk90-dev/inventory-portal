import os
import sqlite3
import datetime
import csv
import random
from io import BytesIO, TextIOWrapper
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'inventory-portal-secure-key-prod-2024')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit

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

    conn.commit()
    conn.close()

init_schema()

def get_user_info():
    email = request.headers.get('X-User-Email', '').strip().lower()
    if not email:
        email = request.args.get('user_email', '').strip().lower()
    if not email and request.form:
        email = request.form.get('user_email', '').strip().lower()
    # Defensive security: Unauthenticated or direct anonymous requests default to Guest Viewer (read-only)
    email = email or 'viewer@company.com'
    
    name = request.headers.get('X-User-Name', '').strip()
    if not name:
        name = request.args.get('user_name', '').strip()
    if not name and request.form:
        name = request.form.get('user_name', '').strip()
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
            
    return {
        'user_email': user_email,
        'is_admin': is_admin,
        'can_edit_inventory': can_edit_inventory,
        'can_edit_dispatch': can_edit_dispatch,
        'can_edit_dispatch_online': can_edit_dispatch_online,
        'can_edit_dispatch_gt': can_edit_dispatch_gt,
        'can_edit_dispatch_mt': can_edit_dispatch_mt,
        'online_dispatch_emails': list(online_dispatch_emails),
        'gt_dispatch_emails': list(gt_dispatch_emails),
        'mt_dispatch_emails': list(mt_dispatch_emails),
        'allowed_courier_platforms': allowed_courier_platforms,
        'allowed_courier_channels': allowed_courier_channels,
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
    
    cur.execute('''
    UPDATE dispatch_online SET
        courier_name = ?,
        tracking_id = ?,
        courier_charges = ?,
        dispatch_status = ?,
        eway_bill_no = ?,
        shipping_notes = ?
    WHERE id = ?
    ''', (courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes, item_id))
    
    edit_sum = f"Courier: '{courier_name}', AWB: '{tracking_id}', Charges: ₹{courier_charges:,.2f}, Status: '{dispatch_status}' (Dispatch #{item_id}, {platform})"
    record_audit(conn, 'UPDATE_COURIER_ONLINE', 'Dispatch Online', f'Dispatch #{item_id} | {platform}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Courier details updated successfully'})

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
    
    cur.execute('''
    UPDATE dispatch_gt_mt SET
        courier_name = ?,
        tracking_id = ?,
        courier_charges = ?,
        dispatch_status = ?,
        eway_bill_no = ?,
        shipping_notes = ?
    WHERE id = ?
    ''', (courier_name, tracking_id, courier_charges, dispatch_status, eway_bill_no, shipping_notes, item_id))
    
    edit_sum = f"Transporter: '{courier_name}', LR/AWB: '{tracking_id}', Charges: ₹{courier_charges:,.2f}, Status: '{dispatch_status}' (GT/MT #{item_id}, {row['channel']} - {row['buyer_distributor']})"
    record_audit(conn, 'UPDATE_COURIER_GTMT', 'Dispatch GT MT', f'Dispatch #{item_id} | {row["channel"]}', edit_sum)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Transporter details updated successfully'})

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
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# --- DISPATCH DOCUMENT UPLOADS (PO & INVOICE COPIES) ---
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
    
    if not item_id:
        return jsonify({'error': 'Dispatch ID is required'}), 400
        
    conn = get_db()
    cur = conn.cursor()
    
    table_name = 'dispatch_online' if dispatch_type == 'online' else 'dispatch_gt_mt'
    row = cur.execute(f"SELECT * FROM {table_name} WHERE id = ?", (item_id,)).fetchone()
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
    filename = f"{dispatch_type}_{item_id}_{doc_type}_{timestamp}_{clean_base}{ext}"
    dest_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(dest_path)
    
    doc_url = f"/uploads/dispatch_docs/{filename}"
    doc_name = file.filename
    
    if doc_type == 'po':
        cur.execute(f"UPDATE {table_name} SET po_doc_url = ?, po_doc_name = ? WHERE id = ?", (doc_url, doc_name, item_id))
    else:
        cur.execute(f"UPDATE {table_name} SET invoice_doc_url = ?, invoice_doc_name = ? WHERE id = ?", (doc_url, doc_name, item_id))
        
    audit_label = 'PO Copy' if doc_type == 'po' else 'Invoice Copy'
    channel_ref = row['platform'] if dispatch_type == 'online' else row['channel']
    record_audit(conn, 'UPLOAD_DISPATCH_DOC', f"Dispatch {dispatch_type.upper()}", f"Dispatch #{item_id} | {channel_ref}", f"Attached {audit_label}: '{doc_name}'")
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'success',
        'doc_type': doc_type,
        'doc_url': doc_url,
        'doc_name': doc_name,
        'message': f"{audit_label} attached successfully"
    })

@app.route('/uploads/dispatch_docs/<path:filename>')
def serve_dispatch_doc(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

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
            
        # Match against our products
        matched_prod = None
        if raw_sku and raw_sku.upper() in sku_lookup:
            matched_prod = sku_lookup[raw_sku.upper()]
        elif raw_name and raw_name.lower() in name_lookup:
            matched_prod = name_lookup[raw_name.lower()]
        else:
            # Substring matching fallback
            for p_name, prod in name_lookup.items():
                if raw_name and (p_name in raw_name.lower() or raw_name.lower() in p_name):
                    matched_prod = prod
                    break
                    
        if matched_prod:
            final_sku = matched_prod['sku']
            final_name = matched_prod['name']
            final_brand = matched_prod['brand']
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

if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8765))
    HOST = os.environ.get('HOST', '0.0.0.0' if os.environ.get('PORT') else '127.0.0.1')
    print(f"Starting Inventory Management System on http://{HOST}:{PORT}")
    app.run(host=HOST, port=PORT, debug=False)
