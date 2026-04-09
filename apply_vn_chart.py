"""
Script áp dụng hệ thống tài khoản chuẩn VN (TT200) cho Odoo 17
Chạy: python apply_vn_chart.py
"""
import xmlrpc.client

# Kết nối Odoo
url = 'http://localhost:8017'
db = 'odoo_company'
username = 'admin'
password = input("Nhập password admin Odoo: ")

print(f"Đang kết nối {url}...")

common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
uid = common.authenticate(db, username, password, {})

if not uid:
    print("❌ Sai password hoặc chưa khởi động Odoo!")
    exit(1)

print(f"✅ Đăng nhập thành công (uid={uid})")

models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

# Kiểm tra chart template VN
print("\n📋 Kiểm tra template hệ thống tài khoản VN...")
company = models.execute_kw(db, uid, password, 'res.company', 'search_read',
    [[]], {'fields': ['id', 'name', 'chart_template'], 'limit': 1})
print(f"Công ty: {company[0]['name']}")
print(f"Chart hiện tại: {company[0].get('chart_template', 'chưa có')}")

company_id = company[0]['id']

# Áp dụng chart VN
print("\n🔄 Đang áp dụng hệ thống tài khoản VN (TT200)...")
try:
    models.execute_kw(db, uid, password, 'account.chart.template', 'try_loading',
        ['vn'], {'company': company_id, 'install_demo': False})
    print("✅ Áp dụng thành công!")
except Exception as e:
    print(f"❌ Lỗi: {e}")
    exit(1)

# Kiểm tra kết quả
count = models.execute_kw(db, uid, password, 'account.account', 'search_count', [[]])
print(f"\n📊 Tổng số tài khoản đã tạo: {count}")

# Kiểm tra các TK quan trọng
important = models.execute_kw(db, uid, password, 'account.account', 'search_read',
    [[['code', 'in', ['111', '112', '131', '331', '511', '632']]]],
    {'fields': ['code', 'name']})

print("\n✅ Các tài khoản quan trọng:")
for acc in sorted(important, key=lambda x: x['code']):
    print(f"  TK {acc['code']} - {acc['name']}")

print("\n🎉 Hoàn tất! Vào Odoo > Accounting > Chart of Accounts để xem.")
