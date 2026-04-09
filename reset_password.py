import psycopg2, sys
from passlib.context import CryptContext

sys.stdout.reconfigure(encoding='utf-8')

new_password = "admin"

ctx = CryptContext(schemes=['pbkdf2_sha512'])
hashed = ctx.hash(new_password)

conn = psycopg2.connect(host='localhost', port=5432, dbname='odoo_company', user='odoo17', password='odoo17pass')
conn.autocommit = True
cur = conn.cursor()

cur.execute("UPDATE res_users SET password = %s WHERE id = 2;", (hashed,))
print(f"Password reset thanh cong!")
print(f"Login: doanvanky36k21@gmail.com")
print(f"Password moi: {new_password}")

cur.close()
conn.close()
