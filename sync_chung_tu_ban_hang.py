"""
sync_chung_tu_ban_hang.py
Import Sổ Nhật Ký Chung (xuất từ MISA) vào Odoo 17
- Mỗi Số CT → 1 account.move
- Mỗi dòng Excel → 1 account.move.line
- Tự động tạo analytic account từ mã thống kê (plan "Mã thống kê")

Usage:
    py -3 sync_chung_tu_ban_hang.py <file.xlsx> [journal_code]
    journal_code: INV (bán hàng), BILL (mua hàng), MISC (tổng hợp)
"""

import psycopg2
import pandas as pd
import sys, os, json

sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = dict(host='localhost', port=5432, dbname='odoo_company',
                 user='odoo17', password='odoo17pass')
COMPANY_ID  = 1
CURRENCY_ID = 23   # VND
ANALYTIC_PLAN_ID = 4  # Plan "Mã thống kê"

# Vị trí cột trong file MISA Sổ Nhật Ký Chung (0-based, bỏ 4 dòng đầu)
COL = dict(
    ngay_hachtoan=0, ngay_ct=1, so_ct=2,
    dien_giai_chung=3, dien_giai=4,
    tk=5, ten_tk=6, tk_du=7, ten_tk_du=8,
    loai_tien=9, ty_gia=10,
    phat_sinh_no=11, phat_sinh_co=12,
    loai_ct=13, ma_doi_tuong=14, ten_doi_tuong=15,
    ma_thong_ke=16, ten_thong_ke=17,
)

# ── Cache ────────────────────────────────────────────────
_account_cache   = {}
_partner_cache   = {}
_analytic_cache  = {}

def get_journal(cur, code):
    cur.execute("SELECT id, name->>'vi_VN' FROM account_journal WHERE code=%s", (code,))
    row = cur.fetchone()
    if not row:
        raise ValueError(f"Không tìm thấy journal code='{code}'!")
    return row[0], row[1]

def get_account(cur, code_str):
    key = str(code_str).strip()
    if key in _account_cache:
        return _account_cache[key]
    cur.execute("SELECT id FROM account_account WHERE code=%s AND company_id=%s",
                (key, COMPANY_ID))
    row = cur.fetchone()
    if not row:
        print(f"  ⚠ Không tìm thấy TK '{key}'!")
    _account_cache[key] = row[0] if row else None
    return _account_cache[key]

def get_partner(cur, ma):
    if not ma or str(ma).strip() in ('', 'nan'):
        return None
    key = str(ma).strip()
    if key in _partner_cache:
        return _partner_cache[key]
    cur.execute("SELECT id FROM res_partner WHERE ref=%s AND active=TRUE LIMIT 1", (key,))
    row = cur.fetchone()
    if not row:
        print(f"  ⚠ Không tìm thấy đối tượng ref='{key}' — bỏ trống partner")
    _partner_cache[key] = row[0] if row else None
    return _partner_cache[key]

def get_or_create_analytic(cur, ma, ten):
    """Tìm hoặc tạo analytic account theo mã thống kê"""
    if not ma or str(ma).strip() in ('', 'nan'):
        return None
    key = str(ma).strip()
    if key in _analytic_cache:
        return _analytic_cache[key]
    cur.execute("SELECT id FROM account_analytic_account WHERE code=%s AND plan_id=%s",
                (key, ANALYTIC_PLAN_ID))
    row = cur.fetchone()
    if row:
        _analytic_cache[key] = row[0]
    else:
        ten_str = str(ten).strip() if ten and str(ten) != 'nan' else key
        name_json = json.dumps({'vi_VN': ten_str, 'en_US': ten_str})
        cur.execute("""
            INSERT INTO account_analytic_account
                (name, code, plan_id, company_id, active, create_uid, write_uid, create_date, write_date)
            VALUES(%s, %s, %s, %s, TRUE, 1, 1, NOW(), NOW())
            RETURNING id
        """, (name_json, key, ANALYTIC_PLAN_ID, COMPANY_ID))
        new_id = cur.fetchone()[0]
        _analytic_cache[key] = new_id
        print(f"  + Tạo analytic: [{key}] {ten_str} (id={new_id})")
    return _analytic_cache[key]

def next_move_name(cur, journal_code, move_date):
    from datetime import date
    d = move_date if hasattr(move_date, 'year') else pd.to_datetime(move_date).date()
    prefix = f"{journal_code}/{d.year:04d}/{d.month:02d}/"
    cur.execute("SELECT COUNT(*) FROM account_move WHERE name LIKE %s", (prefix + '%',))
    cnt = cur.fetchone()[0]
    return f"{prefix}{cnt + 1:04d}"

def to_float(val):
    try:
        return float(str(val).replace(',', '').strip())
    except:
        return 0.0

def to_date(val):
    if pd.isna(val) if not isinstance(val, str) else val.strip() == '':
        return None
    try:
        return pd.to_datetime(val).date()
    except:
        return None

def safe_str(val):
    s = str(val).strip()
    return '' if s == 'nan' else s

# ── Đọc Excel ───────────────────────────────────────────
def read_excel(filepath):
    df = pd.read_excel(filepath, sheet_name=0, header=None, skiprows=4, dtype=str)
    df = df.dropna(how='all')
    df = df[df.iloc[:, COL['so_ct']].str.strip().str.len() > 0]
    return df

def group_vouchers(df):
    vouchers = {}
    for _, row in df.iterrows():
        so_ct = safe_str(row.iloc[COL['so_ct']])
        if so_ct:
            vouchers.setdefault(so_ct, []).append(row)
    return vouchers

# ── Import ───────────────────────────────────────────────
def import_vouchers(filepath, journal_code='INV'):
    conn = psycopg2.connect(**DB_CONFIG)
    cur  = conn.cursor()

    journal_id, journal_name = get_journal(cur, journal_code)
    print(f"Journal: [{journal_code}] {journal_name} (id={journal_id})")

    df       = read_excel(filepath)
    vouchers = group_vouchers(df)
    print(f"Đọc được {len(df)} dòng → {len(vouchers)} chứng từ\n")

    stats = dict(success=0, skip=0, error=0)

    for so_ct, rows in vouchers.items():
        print(f"─── {so_ct} ({len(rows)} dòng) ───")

        # Ngày và ref từ dòng đầu tiên
        move_date = to_date(rows[0].iloc[COL['ngay_hachtoan']]) or __import__('datetime').date.today()
        ref       = safe_str(rows[0].iloc[COL['dien_giai_chung']])

        lines = []
        ok    = True

        for row in rows:
            tk_code = safe_str(row.iloc[COL['tk']])
            debit   = to_float(row.iloc[COL['phat_sinh_no']])
            credit  = to_float(row.iloc[COL['phat_sinh_co']])

            if debit == 0 and credit == 0:
                continue

            account_id = get_account(cur, tk_code)
            if not account_id:
                print(f"  ✗ Bỏ qua CT {so_ct}: TK {tk_code} không tồn tại")
                ok = False
                break

            ma_dt      = safe_str(row.iloc[COL['ma_doi_tuong']])
            partner_id = get_partner(cur, ma_dt)

            # Diễn giải
            name   = safe_str(row.iloc[COL['dien_giai']]) or ref
            ma_tk  = safe_str(row.iloc[COL['ma_thong_ke']])  if len(row) > COL['ma_thong_ke']  else ''
            ten_tk = safe_str(row.iloc[COL['ten_thong_ke']]) if len(row) > COL['ten_thong_ke'] else ''
            if ten_tk and ten_tk not in name:
                name = f"{name} / {ten_tk}"

            # Analytic từ mã thống kê
            analytic_id   = get_or_create_analytic(cur, ma_tk, ten_tk)
            analytic_dist = json.dumps({str(analytic_id): 100.0}) if analytic_id else None

            balance         = round(debit - credit, 2)
            amount_currency = balance  # VND, ty_gia=1

            lines.append(dict(
                account_id=account_id, partner_id=partner_id, name=name,
                debit=debit, credit=credit, balance=balance,
                amount_currency=amount_currency,
                analytic_distribution=analytic_dist,
            ))
            lbl = 'Nợ' if debit else 'Có'
            amt = debit or credit
            atk = f"[{ma_tk}]" if ma_tk else ''
            print(f"  {lbl} TK {tk_code}: {amt:>14,.0f}  {atk} {ten_tk[:30]}")

        if not ok:
            stats['skip'] += 1
            print()
            continue

        # Kiểm tra cân bằng
        total_debit  = round(sum(l['debit']  for l in lines), 2)
        total_credit = round(sum(l['credit'] for l in lines), 2)
        if abs(total_debit - total_credit) > 0.01:
            print(f"  ✗ LỖI CÂN BẰNG: Nợ={total_debit:,.0f} ≠ Có={total_credit:,.0f}")
            stats['error'] += 1
            print()
            continue

        try:
            move_name = next_move_name(cur, journal_code, move_date)

            cur.execute("""
                INSERT INTO account_move
                    (name, date, journal_id, move_type, state, auto_post,
                     currency_id, company_id, ref, posted_before,
                     create_uid, write_uid, create_date, write_date)
                VALUES(%s,%s,%s,'entry','posted','no',%s,%s,%s,TRUE,1,1,NOW(),NOW())
                RETURNING id
            """, (move_name, move_date, journal_id, CURRENCY_ID, COMPANY_ID, ref))
            move_id = cur.fetchone()[0]

            for seq, line in enumerate(lines, 1):
                cur.execute("""
                    INSERT INTO account_move_line
                        (move_id, sequence, account_id, partner_id,
                         name, debit, credit, balance,
                         amount_currency, currency_id, company_currency_id,
                         amount_residual, amount_residual_currency,
                         journal_id, date, company_id, display_type,
                         analytic_distribution,
                         create_uid, write_uid, create_date, write_date)
                    VALUES
                        (%s,%s,%s,%s,
                         %s,%s,%s,%s,
                         %s,%s,%s,
                         %s,%s,
                         %s,%s,%s,'product',
                         %s,
                         1,1,NOW(),NOW())
                """, (
                    move_id, seq, line['account_id'], line['partner_id'],
                    line['name'], line['debit'], line['credit'], line['balance'],
                    line['amount_currency'], CURRENCY_ID, CURRENCY_ID,
                    line['amount_currency'], line['amount_currency'],
                    journal_id, move_date, COMPANY_ID,
                    line['analytic_distribution'],
                ))

            conn.commit()
            print(f"  ✓ {move_name} (id={move_id}) | Tổng: {total_debit:,.0f}đ\n")
            stats['success'] += 1

        except Exception as e:
            conn.rollback()
            print(f"  ✗ LỖI DB: {e}\n")
            stats['error'] += 1

    cur.close()
    conn.close()
    print("=" * 55)
    print(f"KẾT QUẢ: Thành công={stats['success']}  Bỏ qua={stats['skip']}  Lỗi={stats['error']}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: py -3 sync_chung_tu_ban_hang.py <file.xlsx> [journal_code]")
        sys.exit(1)
    filepath     = sys.argv[1]
    journal_code = sys.argv[2].upper() if len(sys.argv) > 2 else 'INV'
    if not os.path.exists(filepath):
        print(f"Không tìm thấy file: {filepath}")
        sys.exit(1)
    import_vouchers(filepath, journal_code)
