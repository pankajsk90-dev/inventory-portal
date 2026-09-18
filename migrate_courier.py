import sqlite3

def run_migration():
    conn = sqlite3.connect('inventory.db')
    cur = conn.cursor()
    
    tables = ['dispatch_online', 'dispatch_gt_mt']
    cols = [
        ('courier_name', 'TEXT DEFAULT \'\''),
        ('tracking_id', 'TEXT DEFAULT \'\''),
        ('courier_charges', 'REAL DEFAULT 0.0'),
        ('dispatch_status', 'TEXT DEFAULT \'Dispatched\''),
        ('eway_bill_no', 'TEXT DEFAULT \'\''),
        ('shipping_notes', 'TEXT DEFAULT \'\'')
    ]
    
    for t in tables:
        existing = [c[1] for c in cur.execute(f'PRAGMA table_info({t})').fetchall()]
        for col_name, col_type in cols:
            if col_name not in existing:
                cur.execute(f'ALTER TABLE {t} ADD COLUMN {col_name} {col_type}')
                print(f'Added {col_name} to {t}')
            else:
                print(f'{col_name} already exists in {t}')
                
    # Ensure LOGISTICS_MANAGER role exists in permissions table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS permissions (
        role_key TEXT PRIMARY KEY,
        assigned_email TEXT NOT NULL,
        description TEXT
    )
    ''')
    cur.execute('''
    INSERT OR IGNORE INTO permissions (role_key, assigned_email, description)
    VALUES ('LOGISTICS_MANAGER', 'logistics@company.com', 'Courier Service, Tracking ID, and Freight Charges Manager')
    ''')
    
    conn.commit()
    conn.close()
    print('Migration completed successfully!')

if __name__ == '__main__':
    run_migration()
