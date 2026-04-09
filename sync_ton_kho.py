"""
sync_ton_kho.py
Reset ton kho thuc te tu file Tong Hop Ton Kho MISA -> Odoo stock.quant
- Doc file 6 cot: Ten kho | Ma hang | Ten hang | DVT | So luong | Gia tri
- Tim san pham theo ten vi_VN (vi default_code thuong NULL)
- UPDATE stock.quant neu da co, INSERT neu chua co

Usage:
    py -3 D:\odoo\sync_ton_kho.py "file_ton_kho.xlsx"
"""
import psycopg2, pandas as pd, sys, os
sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG  = dict(host='localhost', port=5432, dbname='odoo_company',
                  user='odoo17', password='odoo17pass')
COMPANY_ID = 1

KHO_MAP = {
    'Hang hoa':        42,
    'Vat tu':          39,
    'Cong cu dung cu': 40,
    'Thanh pham':      41,
}
KHO_MAP_VI = {
    'Hàng hóa':        42,
    'Vật tư':          39,
    'Công cụ dụng cụ': 40,
    'Thành phẩm':      41,
}
LOC_OVERRIDE = {
    'DO':    47,
    'VT003': 45,
    'VT004': 48,
    'VT052': 44,
}

def import_ton_kho(filepath):
    conn = psycopg2.connect(**DB_CONFIG)
    cur  = conn.cursor()

    df = pd.read_excel(filepath, sheet_name=0, header=None, dtype=str, skiprows=5)
    df.columns = ['ten_kho','ma','ten','dvt','sl','gia_tri'] + list(df.columns[6:])
    df = df[df['ma'].notna() & (df['ma'].str.strip().str.len() > 0)]
    df = df[~df['ten_kho'].fillna('').str.strip().str.lower().str.startswith('t\u1ed5ng')]

    print(f'\u0110\u1ecdc \u0111\u01b0\u1ee3c {len(df)} s\u1ea3n ph\u1ea9m\n')
    stats = dict(updated=0, inserted=0, notfound=0)

    for _, row in df.iterrows():
        ma      = str(row['ma']).strip()
        ten     = str(row['ten']).strip() if pd.notna(row['ten']) else ''
        ten_kho = str(row['ten_kho']).strip() if pd.notna(row['ten_kho']) else ''
        try:    sl = float(str(row['sl']).replace(',', ''))
        except: sl = 0.0

        cur.execute("SELECT pp.id FROM product_product pp WHERE pp.default_code=%s", (ma,))
        r = cur.fetchone()
        if not r:
            cur.execute("""SELECT pp.id FROM product_product pp
                JOIN product_template pt ON pt.id=pp.product_tmpl_id
                WHERE pt.name->>'vi_VN' ILIKE %s LIMIT 1""", (ten,))
            r = cur.fetchone()
        if not r:
            print(f'  \u2717 [{ma}] "{ten}" \u2014 kh\u00f4ng t\u00ecm th\u1ea5y')
            stats['notfound'] += 1
            continue

        product_id = r[0]
        loc_id = LOC_OVERRIDE.get(ma) or KHO_MAP_VI.get(ten_kho) or KHO_MAP.get(ten_kho, 39)

        cur.execute('SELECT id, quantity FROM stock_quant WHERE product_id=%s AND location_id=%s LIMIT 1',
                    (product_id, loc_id))
        q = cur.fetchone()
        if q:
            cur.execute('UPDATE stock_quant SET quantity=%s, write_date=NOW() WHERE id=%s', (sl, q[0]))
            print(f'  \u2713 [{ma:6s}] {ten[:32]:32s} {float(q[1]):>10,.3f} \u2192 {sl:>10,.3f}')
            stats['updated'] += 1
        else:
            cur.execute("""INSERT INTO stock_quant
                (product_id,location_id,quantity,reserved_quantity,
                 in_date,company_id,create_uid,write_uid,create_date,write_date)
                VALUES (%s,%s,%s,0,NOW(),%s,1,1,NOW(),NOW())""",
                (product_id, loc_id, sl, COMPANY_ID))
            print(f'  + [{ma:6s}] {ten[:32]:32s} INSERT {sl:>10,.3f}')
            stats['inserted'] += 1

        conn.commit()

    cur.close(); conn.close()
    print()
    print('=' * 60)
    print(f'K\u1ebe T QU\u1ea2: C\u1eadp nh\u1eadt={stats["updated"]}  T\u1ea1o m\u1edbi={stats["inserted"]}  Kh\u00f4ng t\u00ecm th\u1ea5y={stats["notfound"]}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: py -3 sync_ton_kho.py "file_ton_kho.xlsx"')
        sys.exit(1)
    fp = sys.argv[1]
    if not os.path.exists(fp):
        print(f'Kh\u00f4ng t\u00ecm th\u1ea5y file: {fp}'); sys.exit(1)
    import_ton_kho(fp)
