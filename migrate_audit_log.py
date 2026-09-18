import sqlite3

def migrate():
    conn = sqlite3.connect('c:/Users/Admin/Desktop/Inventory/inventory.db')
    cur = conn.cursor()
    
    # Check existing columns
    columns = [row[1] for row in cur.execute("PRAGMA table_info(audit_log)").fetchall()]
    print("Existing audit_log columns:", columns)
    
    new_cols = [
        ('user_name', 'TEXT DEFAULT "System"'),
        ('user_email', 'TEXT DEFAULT "admin@company.com"'),
        ('action_type', 'TEXT DEFAULT "INFO"'),
        ('sheet_name', 'TEXT DEFAULT "General"'),
        ('entity_ref', 'TEXT DEFAULT ""'),
        ('edit_summary', 'TEXT DEFAULT ""')
    ]
    
    for col_name, col_type in new_cols:
        if col_name not in columns:
            cur.execute(f"ALTER TABLE audit_log ADD COLUMN {col_name} {col_type}")
            print(f"Added column {col_name}")
            
    conn.commit()
    conn.close()
    print("Migration finished successfully.")

if __name__ == '__main__':
    migrate()
