import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
cur = conn.cursor()

# Check reconcile flag on VN accounts
cur.execute("""
    SELECT code, name, reconcile, account_type
    FROM account_account
    WHERE code ~ '^[1-9][0-9]{2,4}$'
    ORDER BY code;
""")
print("code | reconcile | type")
for row in cur.fetchall():
    code = row[0]
    name = str(row[1])[:40] if row[1] else ''
    rec = row[2]
    atype = row[3]
    flag = "✓" if rec else " "
    print(f"  {code:8} [{flag}] {atype:30} {name[:35]}")

cur.close()
conn.close()
