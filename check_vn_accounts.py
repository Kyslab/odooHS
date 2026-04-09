import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
cur = conn.cursor()

# Show VN-style accounts (1-4 digit codes)
cur.execute("""
    SELECT code, name FROM account_account
    WHERE code ~ '^[1-9][0-9]{2,4}$'
    ORDER BY code;
""")
print("=== VN accounts (TT200 style) ===")
for row in cur.fetchall():
    print(f"  {row[0]}  {row[1]}")

print()

# Show generic/system accounts
cur.execute("""
    SELECT code, name FROM account_account
    WHERE code ~ '^[0-9]{6,}$'
    ORDER BY code;
""")
print("=== System/generic accounts (6+ digits) ===")
for row in cur.fetchall():
    print(f"  {row[0]}  {row[1]}")

cur.close()
conn.close()
