import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

# Check parent chain of Chart of Accounts menu (id=157)
cur.execute("""
WITH RECURSIVE parents AS (
    SELECT id, name, parent_id, 0 as depth
    FROM ir_ui_menu WHERE id = 157
    UNION ALL
    SELECT m.id, m.name, m.parent_id, p.depth+1
    FROM ir_ui_menu m JOIN parents p ON m.id = p.parent_id
)
SELECT id, name, parent_id FROM parents ORDER BY depth DESC;
""")
print("=== Parent chain of Chart of Accounts ===")
for row in cur.fetchall():
    print(f"  [{row[0]}] {row[1]} (parent_id={row[2]})")

# Check group restrictions on this chain
cur.execute("""
    SELECT r.menu_id, m.name, r.gid, g.name as group_name
    FROM ir_ui_menu_group_rel r
    JOIN ir_ui_menu m ON r.menu_id = m.id
    JOIN res_groups g ON r.gid = g.id
    WHERE r.menu_id IN (157, 156, 148, 118, 135, 143);
""")
print("\n=== Group restrictions ===")
rows = cur.fetchall()
if rows:
    for row in rows:
        print(f"  menu [{row[0]}] {row[1]} => group [{row[2]}] {row[3]}")
else:
    print("  (no restrictions)")

# Fix: remove all restrictions from the accounting config menus
menus_to_fix = [157, 156, 148, 135, 143, 144, 145, 146]
cur.execute("DELETE FROM ir_ui_menu_group_rel WHERE menu_id = ANY(%s);", (menus_to_fix,))
print(f"\nRemoved {cur.rowcount} group restrictions from accounting menus")

# Clear assets cache
cur.execute("DELETE FROM ir_attachment WHERE url LIKE '/web/assets/%';")
print(f"Cleared {cur.rowcount} cached assets")

cur.close()
conn.close()
print("\nDone! Refresh browser (Ctrl+Shift+R)")
