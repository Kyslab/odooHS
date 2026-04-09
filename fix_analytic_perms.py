import psycopg2, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

uid = 2  # doanvanky36k21@gmail.com

# Kiểm tra group 15 đã có chưa
cur.execute("SELECT 1 FROM res_groups_users_rel WHERE uid=%s AND gid=15;", (uid,))
has_group = cur.fetchone()
print(f"Admin có group 15 (Analytic Accounting): {'CÓ' if has_group else 'CHƯA'}")

if not has_group:
    cur.execute("INSERT INTO res_groups_users_rel (uid, gid) VALUES (%s, 15);", (uid,))
    print("Đã thêm group 15!")

# Xóa session cache của Odoo để group mới có hiệu lực
cur.execute("DELETE FROM ir_sessions WHERE uid=%s;", (uid,))
print(f"Xóa {cur.rowcount} session cũ → bạn cần đăng nhập lại")

# Xóa assets cache
cur.execute("DELETE FROM ir_attachment WHERE url LIKE '/web/assets/%';")
print(f"Xóa {cur.rowcount} web assets cache")

cur.close()
conn.close()
print("\nXong! Hãy đăng xuất và đăng nhập lại Odoo.")
