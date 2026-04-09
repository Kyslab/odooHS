import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
cur = conn.cursor()

# Count total accounts
cur.execute("SELECT COUNT(*) FROM account_account;")
print(f"Total accounts: {cur.fetchone()[0]}")

# Show first 20 accounts
cur.execute("SELECT code, name FROM account_account ORDER BY code LIMIT 20;")
print("\n=== First 20 accounts ===")
for row in cur.fetchall():
    print(f"  {row[0]}  {row[1]}")

# Check if VN accounts exist (3-digit codes like 111, 112)
cur.execute("SELECT COUNT(*) FROM account_account WHERE code ~ '^[0-9]{3}$';")
vn_count = cur.fetchone()[0]
print(f"\nVN-style 3-digit accounts: {vn_count}")

# Check company chart template
cur.execute("SELECT id, name, chart_template FROM res_company LIMIT 1;")
row = cur.fetchone()
print(f"\nCompany: {row[1]}, chart_template: {row[2]}")

cur.close()
conn.close()
