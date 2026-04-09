import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
cur = conn.cursor()

# Find all top-level menus
cur.execute("""
WITH RECURSIVE menu_tree AS (
    SELECT id, name, parent_id, sequence, 0 as depth
    FROM ir_ui_menu
    WHERE parent_id IS NULL
    UNION ALL
    SELECT m.id, m.name, m.parent_id, m.sequence, t.depth+1
    FROM ir_ui_menu m
    JOIN menu_tree t ON m.parent_id = t.id
    WHERE t.depth < 3
)
SELECT depth, id, name FROM menu_tree
WHERE depth <= 3
ORDER BY depth, sequence;
""")
print("=== Full menu tree (depth 0-3) ===")
for row in cur.fetchall():
    indent = "  " * row[0]
    print(f"{indent}[{row[1]}] {row[2]}")

cur.close()
conn.close()
