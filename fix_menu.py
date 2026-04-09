"""
Fix OCA accounting report menus - remove Enterprise group restriction
"""
import psycopg2

conn = psycopg2.connect(
    host='localhost',
    port=5432,
    dbname='odoo_company',
    user='odoo17',
    password='odoo17pass'
)
conn.autocommit = True
cur = conn.cursor()

# Check table columns
cur.execute("""
    SELECT column_name FROM information_schema.columns
    WHERE table_name = 'ir_ui_menu_group_rel' ORDER BY ordinal_position;
""")
cols = cur.fetchall()
print(f"Columns: {cols}")

# Check current restrictions
cur.execute("SELECT * FROM ir_ui_menu_group_rel WHERE menu_id IN (326,327,328,329,331,332);")
rows = cur.fetchall()
print(f"Current restrictions: {rows}")

# Remove Enterprise group restriction from OCA accounting menus
menus_to_fix = [326, 327, 328, 329, 331, 332]
cur.execute("DELETE FROM ir_ui_menu_group_rel WHERE menu_id = ANY(%s);", (menus_to_fix,))
print(f"Deleted {cur.rowcount} group restriction(s)")

# Also clear web assets cache
cur.execute("DELETE FROM ir_attachment WHERE url LIKE '/web/assets/%';")
print(f"Cleared {cur.rowcount} cached asset(s)")

# Verify
cur.execute("SELECT * FROM ir_ui_menu_group_rel WHERE menu_id IN (326,327,328,329,331,332);")
rows = cur.fetchall()
print(f"Remaining restrictions: {rows}")

cur.close()
conn.close()
print("Done! Please restart Odoo.")
