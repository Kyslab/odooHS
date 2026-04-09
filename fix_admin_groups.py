import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

# Check admin user groups
cur.execute("""
    SELECT g.id, g.name FROM res_groups g
    JOIN res_groups_users_rel r ON g.id = r.gid
    JOIN res_users u ON r.uid = u.id
    WHERE u.login = 'admin'
    ORDER BY g.id;
""")
print("=== Admin user groups ===")
for row in cur.fetchall():
    print(f"  [{row[0]}] {row[1]}")

# Check accounting groups
cur.execute("SELECT id, name FROM res_groups WHERE id IN (25,26,27,28) ORDER BY id;")
print("\n=== Accounting groups (25-28) ===")
for row in cur.fetchall():
    print(f"  [{row[0]}] {row[1]}")

# Add admin to Billing Administrator group (28)
admin_id_query = "SELECT id FROM res_users WHERE login='admin'"
cur.execute(f"SELECT 1 FROM res_groups_users_rel WHERE uid=({admin_id_query}) AND gid=28;")
if not cur.fetchone():
    cur.execute(f"INSERT INTO res_groups_users_rel (uid, gid) VALUES (({admin_id_query}), 28);")
    print("\nAdded admin to group 28 (Billing Administrator)")
else:
    print("\nAdmin already in group 28")

# Add to group 26 (Billing)
cur.execute(f"SELECT 1 FROM res_groups_users_rel WHERE uid=({admin_id_query}) AND gid=26;")
if not cur.fetchone():
    cur.execute(f"INSERT INTO res_groups_users_rel (uid, gid) VALUES (({admin_id_query}), 26);")
    print("Added admin to group 26 (Billing)")

# Clear cache
cur.execute("DELETE FROM ir_attachment WHERE url LIKE '/web/assets/%';")
print(f"Cleared {cur.rowcount} cached assets")

cur.close()
conn.close()
print("\nDone! Refresh browser with Ctrl+Shift+R")
