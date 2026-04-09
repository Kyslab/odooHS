import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
cur = conn.cursor()
cur.execute("""
    SELECT name, state FROM ir_module_module
    WHERE name IN ('account','account_accountant','l10n_vn','account_financial_report','account_tax_balance','invoicing')
    ORDER BY name;
""")
for row in cur.fetchall():
    print(f"{row[0]} -> {row[1]}")
cur.close()
conn.close()
