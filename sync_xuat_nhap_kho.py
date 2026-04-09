"""
sync_xuat_nhap_kho.py
Import Tổng Hợp Nhập Xuất Tồn MISA → Odoo 17 stock.move
- Mỗi sản phẩm có phát sinh → 1 stock.move NHẬP + 1 stock.move XUẤT
- Tự động cập nhật stock.quant
- Bỏ qua sản phẩm có SL = 0

Usage:
    py -3 D:\odoo\sync_xuat_nhap_kho.py "file_tong_hop.xlsx"
"""

import psycopg2, pandas as pd, sys, os
sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG  = dict(host='localhost', port=5432, dbname='odoo_company',
                  user='odoo17', password='odoo17pass')
COMPANY_ID = 1
MOVE_DATE  = '2026-04-30'   # Ngày cuối kỳ

# Virtual locations
LOC_VENDOR     = 4    # Partners/Vendors  (nguồn nhập)
LOC_PRODUCTION = 15   # Virtual/Production (đích xuất)
LOC_CUSTOMER   = 5    # Partners/Customers

# Internal location mặc định theo prefix mã hàng
# Nếu không có prefix phù hợp → dùng WH/Stock (id=8)
LOC_MAP = {
    'DO':   47,   # WH/Stock/Dầu DO tại bãi Hoàng Sơn
    'HH':   42,   # WH/Stock/Hàng hóa
    'VT':   39,   # WH/Stock/Vật tư (mặc định VT)
}
# Override riêng cho từng mã cụ thể
LOC_OVERRIDE = {
    'VT003': 45,  # Dầu máy   → WH/Stock/Dầu máy tại bãi Hoàng Sơn
    'VT004': 48,  # Dầu thủy lực → WH/Stock/Dầu thủy lực tại bãi Hoàng Sơn
    'VT052': 44,  # Dầu hộp số → WH/Stock/Dầu hộp số tại bãi Hoàng Sơn
}

def get_internal_loc(ma):
    ma = str(ma).strip().upper()
    if ma in LOC_OVERRIDE:
        return LOC_OVERRIDE[ma]
    for prefix, loc_id in LOC_MAP.items():
        if ma.startswith(prefix):
            return loc_id
    return 8   # WH/Stock fallback

def get_product(cur, ma, ten):
    """Tìm product theo default_code, fallback theo tên vi_VN"""
    cur.execute("SELECT pp.id, pt.uom_id FROM product_product pp "
                "JOIN product_template pt ON pt.id=pp.product_tmpl_id "
                "WHERE pp.default_code=%s", (ma,))
    r = cur.fetchone()
    if r: return r
    cur.execute("SELECT pp.id, pt.uom_id FROM product_product pp "
                "JOIN product_template pt ON pt.id=pp.product_tmpl_id "
                "WHERE pt.name->>'vi_VN' ILIKE %s LIMIT 1", (ten,))
    return cur.fetchone()

def insert_move(cur, product_id, uom_id, qty, loc_from, loc_to, name):
    """Tạo stock.move + stock.move.line ở trạng thái done"""
    cur.execute("""
        INSERT INTO stock_move (
            name, product_id, product_qty, product_uom_qty, product_uom,
            location_id, location_dest_id,
            state, date, company_id,
            procure_method, scrapped, priority,
            create_uid, write_uid, create_date, write_date
        ) VALUES (
            %s, %s, %s, %s, %s,
            %s, %s,
            'done', %s, %s,
            'make_to_stock', FALSE, '0',
            1, 1, NOW(), NOW()
        ) RETURNING id
    """, (name, product_id, qty, qty, uom_id,
          loc_from, loc_to, MOVE_DATE, COMPANY_ID))
    move_id = cur.fetchone()[0]

    cur.execute("""
        INSERT INTO stock_move_line (
            move_id, product_id, product_uom_id,
            quantity, quantity_product_uom,
            location_id, location_dest_id,
            state, picked, date, company_id,
            create_uid, write_uid, create_date, write_date
        ) VALUES (
            %s, %s, %s,
            %s, %s,
            %s, %s,
            'done', TRUE, %s, %s,
            1, 1, NOW(), NOW()
        )
    """, (move_id, product_id, uom_id,
          qty, qty,
          loc_from, loc_to, MOVE_DATE, COMPANY_ID))
    return move_id

def update_quant(cur, product_id, loc_id, delta):
    """Cộng/trừ delta vào stock.quant, tạo mới nếu chưa có"""
    cur.execute("""
        SELECT id, quantity FROM stock_quant
        WHERE product_id=%s AND location_id=%s
        LIMIT 1
    """, (product_id, loc_id))
    row = cur.fetchone()
    if row:
        new_qty = float(row[1]) + delta
        cur.execute("UPDATE stock_quant SET quantity=%s, write_date=NOW() WHERE id=%s",
                    (new_qty, row[0]))
        return new_qty
    else:
        cur.execute("""
            INSERT INTO stock_quant (product_id, location_id, quantity, reserved_quantity,
                in_date, company_id, create_uid, write_uid, create_date, write_date)
            VALUES (%s, %s, %s, 0, NOW(), %s, 1, 1, NOW(), NOW())
        """, (product_id, loc_id, delta, COMPANY_ID))
        return delta

# ── Main ─────────────────────────────────────────────────
def import_inventory(filepath):
    conn = psycopg2.connect(**DB_CONFIG)
    cur  = conn.cursor()

    df = pd.read_excel(filepath, sheet_name=0, header=None, dtype=str, skiprows=4)
    df.columns = ['ma', 'ten', 'dvt', 'nhap', 'xuat'] + list(df.columns[5:])
    df = df[df['ma'].notna() & (df['ma'].str.strip().str.len() > 0)]

    stats = dict(ok=0, skip=0, notfound=0)
    print(f'Đọc được {len(df)} sản phẩm\n')

    for _, row in df.iterrows():
        ma  = str(row['ma']).strip()
        ten = str(row['ten']).strip() if pd.notna(row['ten']) else ''
        nhap = float(str(row['nhap']).replace(',','')) if pd.notna(row['nhap']) and str(row['nhap']).strip() not in ('','nan','0') else 0.0
        xuat = float(str(row['xuat']).replace(',','')) if pd.notna(row['xuat']) and str(row['xuat']).strip() not in ('','nan','0') else 0.0

        if nhap == 0 and xuat == 0:
            stats['skip'] += 1
            continue

        prod = get_product(cur, ma, ten)
        if not prod:
            print(f'  ✗ [{ma}] "{ten}" — không tìm thấy sản phẩm')
            stats['notfound'] += 1
            continue

        product_id, uom_id = prod
        loc_internal = get_internal_loc(ma)

        print(f'  [{ma}] {ten[:35]:35s}', end='')

        if nhap > 0:
            mid = insert_move(cur, product_id, uom_id, nhap,
                              LOC_VENDOR, loc_internal, f'Nhập kho kỳ 02-04/2026 [{ma}]')
            new_q = update_quant(cur, product_id, loc_internal, nhap)
            print(f'  NK+{nhap:>10,.3f}', end='')

        if xuat > 0:
            mid = insert_move(cur, product_id, uom_id, xuat,
                              loc_internal, LOC_PRODUCTION, f'Xuất kho kỳ 02-04/2026 [{ma}]')
            new_q = update_quant(cur, product_id, loc_internal, -xuat)
            print(f'  XK-{xuat:>10,.3f}  →tồn={new_q:>10,.3f}', end='')

        print()
        conn.commit()
        stats['ok'] += 1

    cur.close(); conn.close()
    print()
    print('=' * 60)
    print(f'KẾT QUẢ: Thành công={stats["ok"]}  Bỏ qua(SL=0)={stats["skip"]}  Không tìm thấy={stats["notfound"]}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: py -3 sync_xuat_nhap_kho.py <file.xlsx>')
        sys.exit(1)
    fp = sys.argv[1]
    if not os.path.exists(fp):
        print(f'Không tìm thấy file: {fp}'); sys.exit(1)
    import_inventory(fp)
