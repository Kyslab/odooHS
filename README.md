# OdooHS — Hệ thống Odoo 17 Enterprise

Hệ thống quản lý kế toán doanh nghiệp trên Odoo 17, bao gồm module **Quản lý Tiền mặt & Ngân hàng** kiểu MISA.

> 📦 Database dùng chung trên **Neon Cloud** — liên hệ chủ repo để lấy thông tin kết nối.

---

## Yêu cầu hệ thống

| Phần mềm | Phiên bản | Link tải |
|---|---|---|
| **Python** | 3.11.x | https://www.python.org/downloads/release/python-3119/ |
| **Git** | Mới nhất | https://git-scm.com/download/win |
| **wkhtmltopdf** | 0.12.6.1 | https://github.com/wkhtmltopdf/packaging/releases/tag/0.12.6.1-2 |

---

## Cài đặt (chỉ làm 1 lần)

### Bước 1 — Tải repo này về máy

```bat
cd D:\
git clone https://github.com/Kyslab/odooHS.git odoo
cd D:\odoo
```

### Bước 2 — Tải Odoo 17 Community

```bat
git clone https://github.com/odoo/odoo.git --branch 17.0 --depth 1 community
```

> ⏳ Quá trình này mất 5–15 phút (~600MB).

### Bước 3 — Cài thư viện Python

```bat
pip install -r D:\odoo\community\requirements.txt
pip install psycopg2-binary
```

### Bước 4 — Cài wkhtmltopdf (để in PDF)

1. Tải file `wkhtmltox-0.12.6.1.1-...win64.exe` từ link ở bảng trên
2. Cài vào `C:\Program Files\wkhtmltopdf\`
3. Thêm `C:\Program Files\wkhtmltopdf\bin` vào **Windows PATH**

### Bước 5 — Tạo file cấu hình

```bat
copy D:\odoo\odoo_neon.conf.example D:\odoo\odoo_neon.conf
```

Mở file `D:\odoo\odoo_neon.conf` bằng Notepad, điền thông tin nhận từ chủ repo:

```ini
admin_passwd = ← mật khẩu master Odoo (hỏi chủ repo)
db_password  = ← mật khẩu Neon database (hỏi chủ repo)
```

### Bước 6 — Tạo thư mục data

```bat
mkdir D:\odoo\data
```

---

## Chạy Odoo hàng ngày

Bấm đúp vào file **`D:\odoo\start_neon.bat`**

Hoặc mở terminal:
```bat
cd D:\odoo
python community/odoo-bin -c odoo_neon.conf
```

Chờ khoảng **15–30 giây**, rồi mở trình duyệt vào:

**`http://localhost:8017`**

Để dừng: nhấn **Ctrl + C** trong cửa sổ terminal.

---

## Cấu trúc thư mục sau khi cài xong

```
D:\odoo\
├── community/               ← Odoo 17 source (clone từ GitHub Odoo)
├── custom/                  ← Module tùy chỉnh (repo này)
│   └── l10n_vn_cash_manager/
├── scripts/                 ← Script tiện ích
├── data/                    ← Session & file upload (tạo thủ công ở Bước 6)
├── odoo_neon.conf           ← Config kết nối Neon (tự tạo từ .example)
├── odoo_neon.conf.example   ← Template cấu hình (không chứa password)
├── start_neon.bat           ← Bấm đúp để chạy ✅
└── README.md
```

---

## Module l10n_vn_cash_manager

Module quản lý thu chi tiền mặt & ngân hàng:

| Mục menu | Chức năng |
|---|---|
| **Tien Mat → So Quy Tien Mat** | Danh sách phiếu thu/chi tiền mặt (TK 111x) |
| **Tien Mat → Thu Tien → Thu Tien Nhanh** | Tạo phiếu thu nhanh |
| **Tien Mat → Thu Tien → Thu Tien Theo Hoa Don** | Thu tiền + đối trừ TK 131 tự động |
| **Tien Mat → Chi Tien → Chi Tien Nhanh** | Tạo phiếu chi nhanh |
| **Tien Mat → Chi Tien → Tra Tien Theo Hoa Don** | Trả tiền + đối trừ TK 331 tự động |
| **Ngan Hang → ...** | Tương tự tiền mặt nhưng cho TK 112x |

---

## Lưu ý

- ⚠️ Database **dùng chung** — thay đổi của bạn ảnh hưởng đến mọi người
- 🔒 Không commit file `odoo_neon.conf` (chứa password) lên Git
- 🔁 Chỉ chạy **1 instance** tại một thời điểm (cùng port 8017)
