"""
sync_nvk.py
Import Sổ Nhật Ký Chung Nghiệp Vụ Khác MISA → Odoo 17
- Chứng từ nghiệp vụ khác (NVK*)
- Mỗi Số CT → 1 account.move (move_type='entry')
- Mỗi dòng Excel → 1 account.move.line
- Hỗ trợ 3 chiều analytic: Mã KMCP, Mã công trình, Mã thống kê

Usage:
    py -3 sync_nvk.py <file.xlsx>
"""

import psycopg2, pandas as pd, sys, os, json
sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = dict(host='localhost', port=5432, dbname='odoo_company',
                 user='odoo17', password='odoo17pass')
COMPANY_ID       = 1
CURRENCY_ID      = 23   # VND
JOURNAL_ID       = 3    # MISC - Hoạt động khác (Nghiệp vụ khác)

# Analytic plan IDs
PLAN_KMCP        = 3    # Khoản mục chi phí
PLAN_CONGTRINH   = 1    # Dự án / Công trình
PLAN_THONGKE     = 4    # Mã thống kê

# Vị trí cột file MISA Quỹ (22 cột, skip 4 dòng đầu)
COL = dict(
    ngay_hachtoan=0,  ngay_ct=1,      so_ct=2,
    dien_giai_chung=3, dien_giai=4,
    tk=5,             ten_tk=6,        tk_du=7,       ten_tk_du=8,
    loai_tien=9,      ty_gia=10,
    phat_sinh_no=11,  phat_sinh_co=12,
    loai_ct=13,       ma_doi_tuong=14, ten_doi_tuong=15,
    ma_kmcp=16,       ten_kmcp=17,
    ma_congtrinh=18,  ten_congtrinh=19,
    ma_thongke=20,    ten_thongke=21,
)

# ── Cache ────────────────────────────────────────────────
_account_cache  = {}
_partner_cache  = {}
_analytic_cache = {}   # key = (plan_id, code) → analytic_account.id

def get_account(cur, code):
    key = str(code).strip()
    if key not in _account_cache:
        cur.execute("SELECT id FROM account_account WHERE code=%s AND company_id=%s",
                    (key, COMPANY_ID))
        row = cur.fetchone()
        _account_cache[key] = row[0] if row else None
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

def get_or_create_analytic(cur, plan_id, ma, ten):
    """Tìm hoặc tạo analytic account theo plan_id và mã"""
    key_str = str(ma).strip() if ma and str(ma).strip() not in ('', 'nan') else ''
    if not key_str:
        return None
    cache_key = (plan_id, key_str)
    if cache_key not in _analytic_cache:
        cur.execute("SELECT id FROM account_analytic_account WHERE code=%s AND plan_id=%s",
                    (key_str, plan_id))
        row = cur.fetchone()
        if row:
            _analytic_cache[cache_key] = row[0]
        else:
            label = str(ten).strip() if ten and str(ten).strip() not in ('', 'nan') else key_str
            cur.execute("""
                INSERT INTO account_analytic_account
                    (name, code, plan_id, company_id, active, create_uid, write_uid, create_date, write_date)
                VALUES (%s, %s, %s, %s, TRUE, 1, 1, NOW(), NOW()) RETURNING id
            """, (json.dumps({'vi_VN': label, 'en_US': label}), key_str, plan_id, COMPANY_ID))
            new_id = cur.fetchone()[0]
            _analytic_cache[cache_key] = new_id
            print(f"  + Tạo analytic plan={plan_id} [{key_str}] {label} (id={new_id})")
    return _analytic_cache[cache_key]

def build_analytic_distribution(cur, row):
    """Hợp nhất 3 chiều analytic thành 1 dict JSON"""
    dist = {}

    ma_kmcp      = safe_str(row.iloc[COL['ma_kmcp']])      if len(row) > COL['ma_kmcp']      else ''
    ten_kmcp     = safe_str(row.iloc[COL['ten_kmcp']])     if len(row) > COL['ten_kmcp']     else ''
    ma_ct        = safe_str(row.iloc[COL['ma_congtrinh']]) if len(row) > COL['ma_congtrinh'] else ''
    ten_ct       = safe_str(row.iloc[COL['ten_congtrinh']])if len(row) > COL['ten_congtrinh']else ''
    ma_tk        = safe_str(row.iloc[COL['ma_thongke']])   if len(row) > COL['ma_thongke']   else ''
    ten_tk       = safe_str(row.iloc[COL['ten_thongke']])  if len(row) > COL['ten_thongke']  else ''

    an_kmcp = get_or_create_analytic(cur, PLAN_KMCP,      ma_kmcp, ten_kmcp)
    an_ct   = get_or_create_analytic(cur, PLAN_CONGTRINH, ma_ct,   ten_ct)
    an_tk   = get_or_create_analytic(cur, PLAN_THONGKE,   ma_tk,   ten_tk)

    if an_kmcp: dist[str(an_kmcp)] = 100.0
    if an_ct:   dist[str(an_ct)]   = 100.0
    if an_tk:   dist[str(an_tk)]   = 100.0

    return json.dumps(dist) if dist else None

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

# ── Check trùng ─────────────────────────────────────────
def is_duplicate(cur, so_ct):
    cur.execute("SELECT id FROM account_move WHERE name=%s AND journal_id=%s",
                (so_ct, JOURNAL_ID))
    return cur.fetchone() is not None

# ── Insert 1 chứng từ ────────────────────────────────────
def insert_voucher(cur, so_ct, rows):
    move_date = to_date(rows[0].iloc[COL['ngay_hachtoan']])
    ref       = safe_str(rows[0].iloc[COL['dien_giai_chung']])

    lines = []
    for row in rows:
        tk_code = safe_str(row.iloc[COL['tk']])
        debit   = to_float(row.iloc[COL['phat_sinh_no']])
        credit  = to_float(row.iloc[COL['phat_sinh_co']])

        if debit == 0 and credit == 0:
            continue

        acct_id = get_account(cur, tk_code)
        if not acct_id:
            return None, f"TK {tk_code} không tồn tại"

        ma_dt      = safe_str(row.iloc[COL['ma_doi_tuong']]) if len(row) > COL['ma_doi_tuong'] else ''
        partner_id = get_partner(cur, ma_dt)

        name      = safe_str(row.iloc[COL['dien_giai']]) or ref
        analytic  = build_analytic_distribution(cur, row)
        balance   = round(debit - credit, 2)

        lines.append(dict(
            account_id=acct_id, partner_id=partner_id, name=name,
            debit=debit, credit=credit, balance=balance,
            amount_currency=balance,
            analytic_distribution=analytic,
        ))

        lbl = 'Nợ' if debit else 'Có'
        amt = debit or credit
        print(f"  {lbl} TK {tk_code}: {amt:>14,.0f}  {name[:40]}")

    if not lines:
        return None, "Không có dòng hợp lệ"

    total_debit  = round(sum(l['debit']  for l in lines), 2)
    total_credit = round(sum(l['credit'] for l in lines), 2)
    if abs(total_debit - total_credit) > 0.01:
        return None, f"Mất cân bằng: Nợ={total_debit:,.0f} ≠ Có={total_credit:,.0f}"

    # ── INSERT account_move ──────────────────────────────
    cur.execute("""
        INSERT INTO account_move (
            name, ref, date,
            journal_id, move_type, state, auto_post, posted_before,
            currency_id, company_id,
            create_uid, write_uid, create_date, write_date
        ) VALUES (
            %s, %s, %s,
            %s, 'entry', 'posted', 'no', TRUE,
            %s, %s,
            1, 1, NOW(), NOW()
        ) RETURNING id
    """, (so_ct, ref, move_date, JOURNAL_ID, CURRENCY_ID, COMPANY_ID))
    move_id = cur.fetchone()[0]

    # ── INSERT account_move_line ─────────────────────────
    for seq, line in enumerate(lines, 1):
        cur.execute("""
            INSERT INTO account_move_line (
                move_id, sequence, display_type,
                account_id, partner_id, name,
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
                0, 0, 0,
                0, 0,
                %s, %s, %s,
                %s, %s, %s,
                %s, %s,
                %s,
                %s, %s, %s,
                1, 1, NOW(), NOW()
            )
        """, (
            move_id, seq,
            line['account_id'], line['partner_id'], line['name'],
            line['debit'], line['credit'], line['balance'],
            line['amount_currency'], CURRENCY_ID, CURRENCY_ID,
            line['amount_currency'], line['amount_currency'],
            line['analytic_distribution'],
            JOURNAL_ID, move_date, COMPANY_ID,
        ))

    return move_id, None

# ── Main ─────────────────────────────────────────────────
def import_vouchers(filepath):
    conn = psycopg2.connect(**DB_CONFIG)
    cur  = conn.cursor()

    vouchers = read_and_group(filepath)
    print(f"Đọc được {sum(len(v) for v in vouchers.values())} dòng → {len(vouchers)} chứng từ\n")

    stats = dict(success=0, skip=0, error=0, duplicate=0)

    for so_ct, rows in vouchers.items():
        print(f"─── {so_ct} ({len(rows)} dòng) ───")

        if is_duplicate(cur, so_ct):
            print(f"  ⚠ Đã tồn tại — bỏ qua\n")
            stats['duplicate'] += 1
            continue

        try:
            move_id, err = insert_voucher(cur, so_ct, rows)
            if err:
                print(f"  ✗ Bỏ qua: {err}\n")
                conn.rollback()
                stats['skip'] += 1
            else:
                conn.commit()
                total_debit = round(sum(to_float(r.iloc[COL['phat_sinh_no']]) for r in rows), 0)
                print(f"  ✓ {so_ct} (id={move_id}) | Tổng Nợ: {total_debit:,.0f}đ\n")
                stats['success'] += 1
        except Exception as e:
            conn.rollback()
            import traceback; traceback.print_exc()
            print(f"  ✗ LỖI: {e}\n")
            stats['error'] += 1

    cur.close(); conn.close()
    print("=" * 60)
    print(f"KẾT QUẢ: Thành công={stats['success']}  Bỏ qua={stats['skip']}  "
          f"Trùng={stats['duplicate']}  Lỗi={stats['error']}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: py -3 sync_nvk.py <file.xlsx>")
        sys.exit(1)

    fp = sys.argv[1]
    if not os.path.exists(fp):
        print(f"Không tìm thấy file: {fp}"); sys.exit(1)

    import_vouchers(fp)
