import sqlite3

def setup_permissions():
    conn = sqlite3.connect('c:/Users/Admin/Desktop/Inventory/inventory.db')
    cur = conn.cursor()
    cur.execute('''
    CREATE TABLE IF NOT EXISTS permissions (
        role_key TEXT PRIMARY KEY,
        authorized_email TEXT NOT NULL,
        description TEXT
    )
    ''')
    cur.execute("INSERT OR REPLACE INTO permissions VALUES ('INVENTORY_MANAGER', 'inventory@company.com', 'Can edit Total Inventory Base & Thresholds')")
    cur.execute("INSERT OR REPLACE INTO permissions VALUES ('DISPATCH_MANAGER', 'dispatch@company.com', 'Can log and edit Dispatch entries')")
    cur.execute("INSERT OR REPLACE INTO permissions VALUES ('ADMIN', 'admin@company.com', 'Super Admin with full access')")
    conn.commit()
    print("Permissions setup complete:")
    for row in cur.execute("SELECT * FROM permissions").fetchall():
        print(" ", row)
    conn.close()

if __name__ == '__main__':
    setup_permissions()
