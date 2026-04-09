import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

# List all users
cur.execute("SELECT id, login, active FROM res_users ORDER BY id;")
print("=== All users ===")
users = cur.fetchall()
for row in users:
    print(f"  id={row[0]} login={row[1]} active={row[2]}")

# Check groups for user id=2 (usually admin in Odoo)
if users:
    uid = users[0][0]
    cur.execute(f"""
        SELECT g.id, g.name FROM res_groups g
        JOIN res_groups_users_rel r ON g.id = r.gid
        WHERE r.uid = {uid};
    """)
    print(f"\n=== Groups for user id={uid} ===")
    for row in cur.fetchall():
        print(f"  [{row[0]}] {row[1]}")

cur.close()
conn.close()
