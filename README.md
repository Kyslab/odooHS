# Odoo 17 — Bộ Script Import MISA

> Cập nhật lần cuối: 2026-04-08
> Database: `odoo_company` | Host: `localhost:5432` | User: `odoo17`

---

## 2 Script chính cần nhớ

| Script | Dùng cho | Ghi chú |
|--------|----------|---------|
| `sync_hoa_don_ban_hang.py` | **Hóa đơn bán hàng** (BH*) | File 19 cột — tạo `out_invoice` |
| `sync_phieu_thu_chi.py` | **Tất cả loại còn lại** | File 22 cột — tạo journal entry, dùng `--journal` để chọn sổ |

> `sync_nvk.py` là bản copy của `sync_phieu_thu_chi.py` với journal mặc định = 3 (MISC), tạo ra để dễ nhớ.

### Cách dùng `sync_phieu_thu_chi.py` theo loại chứng từ

```bash
# Quỹ tiền mặt (PT*, PC*) — journal CSH1 id=7 (mặc định)
py -3 D:\odoo\sync_phieu_thu_chi.py "file_quy.xlsx"

# Nghiệp vụ khác (NVK*) — journal MISC id=3
py -3 D:\odoo\sync_phieu_thu_chi.py "file_nvk.xlsx" --journal 3

# Ngân hàng chung (NTTK*, UNC*) — journal BNK1 id=6 (tạm thời)
py -3 D:\odoo\sync_phieu_thu_chi.py "file_nganHang.xlsx" --journal 6

# Ngân hàng cụ thể (sau khi xác định đúng TK ngân hàng)
py -3 D:\odoo\sync_phieu_thu_chi.py "file_nganHang.xlsx" --journal <id>
# VCB1=9, VIB1=10, MBB1=12, BIDV1=24, BIDV2=25 ... (xem bảng Journals bên dưới)
```

---

## 1. `sync_hoa_don_ban_hang.py`

**Chức năng:** Import Sổ Nhật Ký Chung bán hàng từ MISA → Odoo **out_invoice**

**File Excel:** 19 cột (không có cột analytic KMCP/Công trình)

**Cấu trúc:**
- Mỗi Số CT → 1 `account.move` (move_type=`out_invoice`, journal INV id=1)
- Dòng có **Phát sinh Có > 0** → `display_type='product'` (dòng hàng hóa)
  - Tìm sản phẩm theo `ma_hang_hoa` (cột S) → `product.default_code`
  - Nếu không tìm thấy → tìm theo tên sản phẩm `pt.name->>'vi_VN'`
- Dòng có **Phát sinh Nợ > 0** → gộp thành 1 dòng `display_type='payment_term'` (công nợ phải thu)
- Analytic 1 chiều: Mã thống kê (plan_id=4)

**Cách dùng:**
```bash
py -3 D:\odoo\sync_hoa_don_ban_hang.py "C:\path\to\file.xlsx"
```

**Cột Excel (19 cột, skip 4 dòng đầu):**
| Index | Tên cột | Dùng cho |
|-------|---------|---------|
| 0 | Ngày hạch toán | move.date |
| 1 | Ngày CT | — |
| 2 | Số CT | move.name |
| 3 | Diễn giải chung | move.ref |
| 4 | Diễn giải | move_line.name / tên sản phẩm |
| 5 | Tài khoản | account_account.code |
| 6 | Tên TK | — |
| 7 | TK đối ứng | — |
| 8 | Tên TK đối ứng | — |
| 9 | Loại tiền | — |
| 10 | Tỷ giá | — |
| 11 | Phát sinh Nợ | debit |
| 12 | Phát sinh Có | credit |
| 13 | Loại CT | — |
| 14 | Mã đối tượng | res_partner.ref |
| 15 | Tên đối tượng | — |
| 16 | Mã thống kê | analytic plan_id=4 |
| 17 | Tên thống kê | — |
| 18 | Mã hàng hóa | product.default_code (cột S) |

**Kết quả đã import:** 5116 hóa đơn (BH0725 → BH1225), id move từ ~10000 trở đi

---

## 2. `sync_phieu_thu_chi.py`

**Chức năng:** Import Sổ Nhật Ký Chung Quỹ từ MISA → Odoo journal entry (CSH1)

**File Excel:** 22 cột (có 3 chiều analytic)

**Cấu trúc:**
- Mỗi Số CT → 1 `account.move` (move_type=`entry`, journal CSH1 id=7)
- Mỗi dòng Excel → 1 `account.move.line` (display_type=`product`)
- Kiểm tra cân bằng Nợ = Có trước khi insert
- Tự động tạo `analytic_account` nếu chưa có
- Analytic **3 chiều**:
  - Cột 16: Mã KMCP → plan_id=3 (Khoản mục chi phí)
  - Cột 18: Mã công trình → plan_id=1 (Dự án)
  - Cột 20: Mã thống kê → plan_id=4

**Cách dùng:**
```bash
# Mặc định journal CSH1 (id=7)
py -3 D:\odoo\sync_phieu_thu_chi.py "C:\path\to\file_quy.xlsx"

# Chỉ định journal khác nếu cần
py -3 D:\odoo\sync_phieu_thu_chi.py "C:\path\to\file.xlsx" --journal <journal_id>
```

**Cột Excel (22 cột, skip 4 dòng đầu):**
| Index | Tên cột | Dùng cho |
|-------|---------|---------|
| 0 | Ngày hạch toán | move.date |
| 1 | Ngày CT | — |
| 2 | Số CT | move.name |
| 3 | Diễn giải chung | move.ref |
| 4 | Diễn giải | move_line.name |
| 5 | Tài khoản | account_account.code |
| 6 | Tên TK | — |
| 7 | TK đối ứng | — |
| 8 | Tên TK đối ứng | — |
| 9 | Loại tiền | — |
| 10 | Tỷ giá | — |
| 11 | Phát sinh Nợ | debit |
| 12 | Phát sinh Có | credit |
| 13 | Loại CT | — |
| 14 | Mã đối tượng | res_partner.ref |
| 15 | Tên đối tượng | — |
| 16 | Mã KMCP | analytic plan_id=3 |
| 17 | Tên KMCP | — |
| 18 | Mã công trình | analytic plan_id=1 |
| 19 | Tên công trình | — |
| 20 | Mã thống kê | analytic plan_id=4 |
| 21 | Tên thống kê | — |

**Kết quả đã import:** 32 chứng từ PT/PC, id move 15223–15254

---

## 3. `sync_nvk.py`

**Chức năng:** Import Sổ Nhật Ký Chung Nghiệp Vụ Khác từ MISA → Odoo journal entry (MISC)

**File Excel:** 22 cột — **cấu trúc giống hệt file Quỹ**

**Cấu trúc:** Giống `sync_phieu_thu_chi.py`, chỉ khác journal mặc định là MISC (id=3)

**Cách dùng:**
```bash
py -3 D:\odoo\sync_nvk.py "C:\path\to\file_nvk.xlsx"
```

**Kết quả đã import:** 35 chứng từ NVK0725.00001–00035, id move 21392–21426

---

## Journals trong Odoo

| id | Code | Tên | Type | Dùng cho |
|----|------|-----|------|---------|
| 1 | INV | Hóa đơn bán hàng | sale | Bán hàng (BH) |
| 2 | BILL | Hoá đơn mua hàng | purchase | Mua hàng |
| 3 | MISC | Hoạt động khác | general | Nghiệp vụ khác (NVK) |
| 6 | BNK1 | Bank | bank | Ngân hàng chung |
| 7 | CSH1 | Cash / Tiền mặt | cash | Quỹ tiền mặt (PT/PC) |
| 9–40 | VCB1..BD15 | Các TK ngân hàng | bank | Chứng từ ngân hàng |

---

## Analytic Plans

| plan_id | Tên | Mã cột Excel |
|---------|-----|-------------|
| 1 | Dự án / Công trình | Mã công trình (cột 18) |
| 3 | Khoản mục chi phí | Mã KMCP (cột 16) |
| 4 | Mã thống kê | Mã thống kê (cột 20 hoặc 16) |

---

## Lưu ý quan trọng

- **Tất cả script đều kiểm tra trùng** theo `name` + `journal_id` trước khi insert
- **Cân bằng Nợ = Có** bắt buộc cho mỗi chứng từ — nếu lệch sẽ bị bỏ qua và báo lỗi
- **Analytic tự tạo** nếu mã chưa tồn tại trong Odoo (tự động tạo `account.analytic.account`)
- **Xem chứng từ** đã import: Kế toán → Kế toán → **Bút toán nhật ký** → filter theo Journal
- **Không tạo account.payment** — PT/PC/NVK được import dưới dạng bút toán nhật ký thuần túy

---

## Thêm loại chứng từ mới

Nếu có loại chứng từ mới (ví dụ ngân hàng):
1. Kiểm tra số cột file Excel: 19 cột → dùng cấu trúc `sync_hoa_don_ban_hang.py`, 22 cột → dùng `sync_phieu_thu_chi.py`
2. Tìm journal_id phù hợp trong bảng Journals ở trên
3. Copy script gần nhất, đổi `JOURNAL_ID` và tên file
4. Chạy thử với 1-2 chứng từ trước khi import toàn bộ

---

## 4. 

**Chức năng:** Import Tổng Hợp Nhập Xuất Tồn từ MISA → Odoo stock.move + stock.quant

**File Excel:** 5 cột (Mã hàng | Tên hàng | ĐVT | Số lượng Nhập kho | Số lượng Xuất kho), skip 4 dòng đầu

**Cấu trúc:**
- Bỏ qua sản phẩm có SL Nhập=0 và Xuất=0
- Tìm sản phẩm theo , fallback theo tên 
- Mỗi sản phẩm → 1 stock.move NHẬP (Vendors→Kho) + 1 stock.move XUẤT (Kho→Production)
- Tự động cập nhật  (đầu kỳ + nhập - xuất = cuối kỳ)
- Ngày move:  (cuối kỳ, chỉnh trong script nếu cần)

**Location map theo prefix mã hàng:**
| Prefix | Kho nội bộ |
|--------|-----------|
| DO | WH/Stock/Dầu DO tại bãi Hoàng Sơn (id=47) |
| HH* | WH/Stock/Hàng hóa (id=42) |
| VT* | WH/Stock/Vật tư (id=39) — mặc định |
| VT003 | WH/Stock/Dầu máy tại bãi Hoàng Sơn (id=45) |
| VT004 | WH/Stock/Dầu thủy lực tại bãi Hoàng Sơn (id=48) |
| VT052 | WH/Stock/Dầu hộp số tại bãi Hoàng Sơn (id=44) |

**Cách dùng:**

