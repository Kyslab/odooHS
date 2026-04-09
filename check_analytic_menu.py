import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

# Tìm menu Analytic Accounts
cur.execute("""
    SELECT m.id, m.name, m.parent_id, p.name as parent_name
    FROM ir_ui_menu m
    LEFT JOIN ir_ui_menu p ON m.parent_id = p.id
    WHERE m.name::text ILIKE '%analytic%'
    ORDER BY m.id;
""")
print("=== Analytic menus ===")
for row in cur.fetchall():
    print(f"  [{row[0]}] {row[1]} | parent: [{row[2]}] {row[3]}")

# Kiểm tra group restriction
cur.execute("""
    SELECT r.menu_id, m.name, r.gid, g.name as gname
    FROM ir_ui_menu_group_rel r
    JOIN ir_ui_menu m ON r.menu_id = m.id
    JOIN res_groups g ON r.gid = g.id
    WHERE m.name::text ILIKE '%analytic%';
""")
print("\n=== Group restrictions on analytic menus ===")
rows = cur.fetchall()
if rows:
    for row in rows:
        print(f"  menu [{row[0]}] {row[1]} => group [{row[2]}] {row[3]}")
else:
    print("  (no restrictions)")

# Đếm analytic accounts đã tạo
cur.execute("SELECT COUNT(*) FROM account_analytic_account;")
print(f"\nTổng analytic accounts trong DB: {cur.fetchone()[0]}")

cur.close()
conn.close()
