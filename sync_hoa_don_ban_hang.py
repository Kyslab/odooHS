"""
sync_hoa_don_ban_hang.py
Import Sổ Nhật Ký Chung MISA → Hóa đơn bán hàng Odoo 17 (out_invoice)

Cấu trúc mỗi chứng từ:
  - Dòng TK credit (511xxx)  → Invoice product line  (display_type='product')
                                name = cột Diễn giải (E)
  - Dòng TK debit  (131xxx)  → Receivable line tổng  (display_type='payment_term')

Usage:
    py -3 sync_hoa_don_ban_hang.py <file.xlsx>
"""

import psycopg2, pandas as pd, sys, os, json
sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = dict(host='localhost', port=5432, dbname='odoo_company',
                 user='odoo17', password='odoo17pass')
COMPANY_ID        = 1
CURRENCY_ID       = 23    # VND
JOURNAL_ID        = 1     # INV - Hóa đơn bán hàng
UOM_UNIT_ID       = 1     # Đơn vị
ANALYTIC_PLAN_ID  = 4     # Plan "Mã thống kê"

COL = dict(
    ngay_hachtoan=0, ngay_ct=1, so_ct=2,
    dien_giai_chung=3, dien_giai=4,
    tk=5, ten_tk=6, tk_du=7, ten_tk_du=8,
    loai_tien=9, ty_gia=10,
    phat_sinh_no=11, phat_sinh_co=12,
    loai_ct=13, ma_doi_tuong=14, ten_doi_tuong=15,
    ma_thong_ke=16, ten_thong_ke=17,
    ma_hang_hoa=18,   # Cột S - Mã Hàng Hóa (mới thêm)
)

# ── Cache ────────────────────────────────────────────────
_account_cache  = {}
_partner_cache  = {}
_analytic_cache = {}
_product_cache  = {}   # key=default_code → (product_id, uom_id)

def get_account(cur, code):
    key = str(code).strip()
    if key not in _account_cache:
        cur.execute("SELECT id, account_type FROM account_account WHERE code=%s AND company_id=%s",
                    (key, COMPANY_ID))
        row = cur.fetchone()
        _account_cache[key] = (row[0], row[1]) if row else (None, None)
        if not row:
            print(f"  ⚠ TK '{key}' không tồn tại trong hệ thống!")
    return _account_cache[key]

def get_partner(cur, ma):
    key = str(ma).strip() if ma and str(ma).strip() not in ('', 'nan') else ''
    if not key:
        return None
    if key not in _partner_cache:
        cur.execute("SELECT id FROM res_partner WHERE ref=%s AND active=TRUE LIMIT 1", (key,))
        row = cur.fetchone()
        _partner_cache[key] = row[0] if row else None
        if not row:
            print(f"  ⚠ Đối tượng '{key}' không tìm thấy — bỏ trống partner")
    return _partner_cache[key]

def get_product(cur, ma_hang):
    """Tìm product_product theo default_code → trả về (product_id, uom_id) hoặc (None, None)"""
    key = str(ma_hang).strip() if ma_hang and str(ma_hang).strip() not in ('', 'nan') else ''
    if not key:
        return None, None
    if key not in _product_cache:
        cur.execute("""
            SELECT pp.id, pt.uom_id
            FROM product_product pp
            JOIN product_template pt ON pt.id = pp.product_tmpl_id
            WHERE pt.default_code = %s AND pp.active = TRUE
            LIMIT 1
        """, (key,))
        row = cur.fetchone()
        _product_cache[key] = (row[0], row[1]) if row else (None, None)
        if not row:
            print(f"  ⚠ Mã hàng '{key}' không tìm thấy trong Odoo — bỏ trống product")
    return _product_cache[key]

def get_or_create_analytic(cur, ma, ten):
    key = str(ma).strip() if ma and str(ma).strip() not in ('', 'nan') else ''
    if not key:
        return None
    if key not in _analytic_cache:
        cur.execute("SELECT id FROM account_analytic_account WHERE code=%s AND plan_id=%s",
                    (key, ANALYTIC_PLAN_ID))
        row = cur.fetchone()
        if row:
            _analytic_cache[key] = row[0]
        else:
            label = str(ten).strip() if ten and str(ten) != 'nan' else key
            cur.execute("""
                INSERT INTO account_analytic_account
                    (name, code, plan_id, company_id, active, create_uid, write_uid, create_date, write_date)
                VALUES (%s,%s,%s,%s,TRUE,1,1,NOW(),NOW()) RETURNING id
            """, (json.dumps({'vi_VN': label, 'en_US': label}), key, ANALYTIC_PLAN_ID, COMPANY_ID))
            _analytic_cache[key] = cur.fetchone()[0]
            print(f"  + Tạo analytic [{key}] {label}")
    return _analytic_cache[key]

def to_float(val):
    try:
        return float(str(val).replace(',', '').strip())
    except:
        return 0.0

def to_date(val):
    try:
        return pd.to_datetime(val).date()
    except:
        return __import__('datetime').date.today()

def safe_str(val):
    s = str(val).strip()
    return '' if s == 'nan' else s

def next_invoice_name(cur, prefix):
    """Sinh tên hóa đơn: BH0725.00001 → dùng thẳng làm name"""
    return prefix   # Dùng luôn Số CT của MISA làm số hóa đơn Odoo

# ── Đọc và nhóm Excel ───────────────────────────────────
def read_and_group(filepath):
    df = pd.read_excel(filepath, sheet_name=0, header=None, skiprows=4, dtype=str)
    df = df.dropna(how='all')
    df = df[df.iloc[:, COL['so_ct']].str.strip().str.len() > 0]
    vouchers = {}
    for _, row in df.iterrows():
        so_ct = safe_str(row.iloc[COL['so_ct']])
        if so_ct:
            vouchers.setdefault(so_ct, []).append(row)
    return vouchers

# ── Insert 1 hóa đơn ────────────────────────────────────
def insert_invoice(cur, so_ct, rows):
    move_date  = to_date(rows[0].iloc[COL['ngay_hachtoan']])
    ref        = safe_str(rows[0].iloc[COL['dien_giai_chung']])  # Số xe/container

    # Tách dòng: credit (511xxx) → product lines | debit (131) → receivable
    product_lines    = []
    receivable_total = 0.0
    receivable_acct  = None
    partner_id       = None

    for row in rows:
        tk_code = safe_str(row.iloc[COL['tk']])
        debit   = to_float(row.iloc[COL['phat_sinh_no']])
        credit  = to_float(row.iloc[COL['phat_sinh_co']])

        if debit == 0 and credit == 0:
            continue

        acct_id, acct_type = get_account(cur, tk_code)
        if not acct_id:
            return None, f"TK {tk_code} không tồn tại"

        ma_dt      = safe_str(row.iloc[COL['ma_doi_tuong']]) if pd.notna(row.iloc[COL['ma_doi_tuong']]) else ''
        ptnr       = get_partner(cur, ma_dt)

        if credit > 0:  # Dòng doanh thu → product line
            name      = safe_str(row.iloc[COL['dien_giai']]) or ref
            ma_tk     = safe_str(row.iloc[COL['ma_thong_ke']])  if len(row) > COL['ma_thong_ke']  else ''
            ten_tk    = safe_str(row.iloc[COL['ten_thong_ke']]) if len(row) > COL['ten_thong_ke'] else ''
            ma_hang   = safe_str(row.iloc[COL['ma_hang_hoa']]) if len(row) > COL['ma_hang_hoa'] else ''
            an_id     = get_or_create_analytic(cur, ma_tk, ten_tk)
            prod_id, uom_id = get_product(cur, ma_hang)

            product_lines.append(dict(
                account_id=acct_id, name=name,
                price_unit=credit, amount=credit,
                analytic=json.dumps({str(an_id): 100.0}) if an_id else None,
                product_id=prod_id,
                uom_id=uom_id or UOM_UNIT_ID,
                ma_hang=ma_hang,
            ))
            print(f"  Dịch vụ  TK {tk_code}: {credit:>14,.0f}  [{ma_hang}] {name[:40]}")

        elif debit > 0:  # Dòng công nợ → receivable (tổng hợp)
            receivable_total += debit
            receivable_acct   = acct_id
            if ptnr:
                partner_id = ptnr

    if not product_lines:
        return None, "Không có dòng doanh thu (credit)"
    if not receivable_acct:
        return None, "Không có dòng công nợ phải thu (debit)"

    total = round(sum(l['amount'] for l in product_lines), 2)
    if abs(total - round(receivable_total, 2)) > 0.01:
        return None, f"Mất cân bằng: DT={total:,.0f} ≠ PT={receivable_total:,.0f}"

    # ── INSERT account_move ──────────────────────────────
    cur.execute("""
        INSERT INTO account_move (
            name, ref, date, invoice_date, invoice_date_due,
            journal_id, move_type, state, auto_post, posted_before,
            currency_id, company_id,
            partner_id, commercial_partner_id,
            payment_state,
            amount_untaxed, amount_tax, amount_total,
            amount_residual,
            amount_untaxed_signed, amount_tax_signed,
            amount_total_signed, amount_total_in_currency_signed,
            amount_residual_signed,
            create_uid, write_uid, create_date, write_date
        ) VALUES (
            %s, %s, %s, %s, %s,
            %s, 'out_invoice', 'posted', 'no', TRUE,
            %s, %s,
            %s, %s,
            'not_paid',
            %s, 0, %s,
            %s,
            %s, 0,
            %s, %s,
            %s,
            1, 1, NOW(), NOW()
        ) RETURNING id
    """, (
        so_ct, ref, move_date, move_date, move_date,
        JOURNAL_ID, CURRENCY_ID, COMPANY_ID,
        partner_id, partner_id,
        total, total,
        total,
        total,
        total, total,
        total,
    ))
    move_id = cur.fetchone()[0]

    seq = 1
    # ── INSERT product lines (TK 511xxx, credit) ────────
    for line in product_lines:
        cur.execute("""
            INSERT INTO account_move_line (
                move_id, sequence, display_type,
                account_id, partner_id, name,
                product_id, product_uom_id,
                quantity, price_unit, discount,
                price_subtotal, price_total,
                debit, credit, balance,
                amount_currency, currency_id, company_currency_id,
                amount_residual, amount_residual_currency,
                analytic_distribution,
                journal_id, date, company_id,
                create_uid, write_uid, create_date, write_date
            ) VALUES (
                %s, %s, 'product',
                %s, %s, %s,
                %s, %s,
                1.0, %s, 0,
                %s, %s,
                0, %s, %s,
                %s, %s, %s,
                0, 0,
                %s,
                %s, %s, %s,
                1, 1, NOW(), NOW()
            )
        """, (
            move_id, seq, line['account_id'], partner_id, line['name'],
            line['product_id'], line['uom_id'],
            line['price_unit'],
            line['amount'], line['amount'],
            line['amount'], -line['amount'],
            -line['amount'], CURRENCY_ID, CURRENCY_ID,
            line['analytic'],
            JOURNAL_ID, move_date, COMPANY_ID,
        ))
        seq += 1

    # ── INSERT receivable line (TK 131, debit) ──────────
    cur.execute("""
        INSERT INTO account_move_line (
            move_id, sequence, display_type,
            account_id, partner_id, name,
            quantity, price_unit, discount,
            price_subtotal, price_total,
            debit, credit, balance,
            amount_currency, currency_id, company_currency_id,
            amount_residual, amount_residual_currency,
            date_maturity,
            journal_id, date, company_id,
            create_uid, write_uid, create_date, write_date
        ) VALUES (
            %s, %s, 'payment_term',
            %s, %s, %s,
            0, 0, 0,
            0, 0,
            %s, 0, %s,
            %s, %s, %s,
            %s, %s,
            %s,
            %s, %s, %s,
            1, 1, NOW(), NOW()
        )
    """, (
        move_id, seq, receivable_acct, partner_id, so_ct,
        total, total,
        total, CURRENCY_ID, CURRENCY_ID,
        total, total,
        move_date,
        JOURNAL_ID, move_date, COMPANY_ID,
    ))

    return move_id, None

# ── Main ─────────────────────────────────────────────────
def import_invoices(filepath):
    conn = psycopg2.connect(**DB_CONFIG)
    cur  = conn.cursor()

    vouchers = read_and_group(filepath)
    print(f"Đọc được {sum(len(v) for v in vouchers.values())} dòng → {len(vouchers)} hóa đơn\n")

    stats = dict(success=0, skip=0, error=0)

    for so_ct, rows in vouchers.items():
        print(f"─── {so_ct} ───")
        try:
            move_id, err = insert_invoice(cur, so_ct, rows)
            if err:
                print(f"  ✗ Bỏ qua: {err}\n")
                conn.rollback()
                stats['skip'] += 1
            else:
                conn.commit()
                # Lay tong de in
                cur.execute("SELECT amount_total FROM account_move WHERE id=%s", (move_id,))
                total = cur.fetchone()[0]
                print(f"  ✓ Hóa đơn {so_ct} (id={move_id}) | Tổng: {float(total):,.0f}đ\n")
                stats['success'] += 1
        except Exception as e:
            conn.rollback()
            import traceback; traceback.print_exc()
            print(f"  ✗ LỖI: {e}\n")
            stats['error'] += 1

    cur.close(); conn.close()
    print("=" * 55)
    print(f"KẾT QUẢ: Thành công={stats['success']}  Bỏ qua={stats['skip']}  Lỗi={stats['error']}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: py -3 sync_hoa_don_ban_hang.py <file.xlsx>")
        sys.exit(1)
    fp = sys.argv[1]
    if not os.path.exists(fp):
        print(f"Không tìm thấy file: {fp}"); sys.exit(1)
    import_invoices(fp)
