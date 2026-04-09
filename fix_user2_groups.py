import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

uid = 2  # doanvanky36k21@gmail.com

# Check current groups of user 2
cur.execute(f"SELECT g.id, g.name FROM res_groups g JOIN res_groups_users_rel r ON g.id=r.gid WHERE r.uid={uid} ORDER BY g.id;")
print(f"=== Current groups of user id={uid} ===")
current_groups = [row[0] for row in cur.fetchall()]
for gid in current_groups:
    cur.execute(f"SELECT name FROM res_groups WHERE id={gid};")
    print(f"  [{gid}] {cur.fetchone()[0]}")

# Add accounting groups: 25 (readonly), 26 (billing), 28 (billing admin)
groups_to_add = [1, 25, 26, 28]  # Internal User + accounting groups
added = []
for gid in groups_to_add:
    if gid not in current_groups:
        try:
            cur.execute("INSERT INTO res_groups_users_rel (uid, gid) VALUES (%s, %s);", (uid, gid))
            added.append(gid)
        except Exception as e:
            print(f"  Skip group {gid}: {e}")

print(f"\nAdded groups: {added}")

# Clear assets cache
cur.execute("DELETE FROM ir_attachment WHERE url LIKE '/web/assets/%';")
print(f"Cleared {cur.rowcount} cached assets")

cur.close()
conn.close()
print("\nDone! Log out and log back in to Odoo, then press Ctrl+Shift+R")
