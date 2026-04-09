import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

# Get company_id
cur.execute("SELECT id FROM res_company LIMIT 1;")
company_id = cur.fetchone()[0]

# Get currency_id (VND)
cur.execute("SELECT id FROM res_currency WHERE name='VND' LIMIT 1;")
row = cur.fetchone()
currency_id = row[0] if row else None
print(f"Company: {company_id}, Currency VND: {currency_id}")

# --- 1. Add TK 131 if missing ---
cur.execute("SELECT id FROM account_account WHERE code='131' AND company_id=%s;", (company_id,))
if not cur.fetchone():
    cur.execute("""
        INSERT INTO account_account
            (name, code, account_type, reconcile, company_id, deprecated, currency_id)
        VALUES
            ('{"en_US": "Trade receivables", "vi_VN": "Phải thu của khách hàng"}',
             '131', 'asset_receivable', TRUE, %s, FALSE, %s)
        RETURNING id;
    """, (company_id, currency_id))
    new_id = cur.fetchone()[0]
    print(f"Created TK 131 (id={new_id})")
else:
    print("TK 131 already exists")

# --- 2. Also add TK 111 and TK 112 parent accounts if missing ---
missing_parents = [
    ('111',  '{"en_US": "Cash", "vi_VN": "Tiền mặt"}', 'asset_cash', False),
    ('112',  '{"en_US": "Bank deposits", "vi_VN": "Tiền gửi ngân hàng"}', 'asset_cash', False),
    ('511',  '{"en_US": "Revenue from sales", "vi_VN": "Doanh thu bán hàng và cung cấp dịch vụ"}', 'income', False),
]
for code, name, atype, rec in missing_parents:
    cur.execute("SELECT id FROM account_account WHERE code=%s AND company_id=%s;", (code, company_id))
    if not cur.fetchone():
        cur.execute("""
            INSERT INTO account_account (name, code, account_type, reconcile, company_id, deprecated)
            VALUES (%s, %s, %s, %s, %s, FALSE) RETURNING id;
        """, (name, code, atype, rec, company_id))
        print(f"Created TK {code} (id={cur.fetchone()[0]})")
    else:
        print(f"TK {code} already exists")

# --- 3. Enable reconcile on key accounts that need it ---
reconcile_codes = [
    '131',   # Phải thu KH
    '1388',  # Phải thu khác
    '1381',  # Tài sản thiếu chờ xử lý
]
cur.execute("""
    UPDATE account_account SET reconcile = TRUE
    WHERE code = ANY(%s) AND company_id = %s AND reconcile = FALSE
    RETURNING code;
""", (reconcile_codes, company_id))
updated = [r[0] for r in cur.fetchall()]
if updated:
    print(f"Enabled reconcile on: {updated}")

# --- 4. Set TK 131 as default receivable for existing customers (optional) ---
cur.execute("SELECT id FROM account_account WHERE code='131' AND company_id=%s;", (company_id,))
row = cur.fetchone()
if row:
    tk131_id = row[0]
    cur.execute("UPDATE res_company SET account_default_pos_receivable_account_id = %s WHERE id = %s;",
                (tk131_id, company_id))
    print(f"Set TK 131 as default receivable account")

print("\nDone! Refresh browser.")
cur.close()
conn.close()
