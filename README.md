# OdooHS — Hệ thống Odoo 17 Community

Hệ thống quản lý kế toán doanh nghiệp trên Odoo 17, bao gồm module **Quản lý Tiền mặt & Ngân hàng** kiểu MISA.

---

## 🖥️ Server VPS (DigitalOcean Singapore)

| Thông tin | Giá trị |
|-----------|---------|
| **IP** | `157.230.45.172` |
| **Truy cập Odoo** | http://157.230.45.172:8017 |
| **OS** | Ubuntu 24.04 LTS |
| **SSH user** | `root` |
| **Database** | `odoo_production` (PostgreSQL 15 local) |
| **DB user** | `odoo17` |
| **Odoo config** | `/etc/odoo.conf` |
| **Odoo log** | `/var/log/odoo/odoo.log` |
| **Custom module** | `/opt/odoo/custom_repo/custom/` |
| **Service** | `systemctl start\|stop\|restart odoo` |

> 🔒 Password SSH và DB lưu riêng — không commit lên Git.

---

## 🔄 Quy trình phát triển (Deploy Workflow)

```
[Máy Windows]  →  [GitHub]  →  [VPS DigitalOcean]
  Sửa code         git push      git pull + restart
```

### Khi muốn code thêm / sửa module:

**Bước 1** — Sửa code trên máy Windows như bình thường:
```
D:\odoo\custom\l10n_vn_cash_manager\
```

**Bước 2** — Bấm đúp vào file này để deploy lên VPS tự động:
```
D:\odoo\scripts\deploy_vps.bat
```

Script sẽ tự động:
1. Commit + push lên GitHub
2. VPS tự pull code mới
3. Restart Odoo

**Bước 3** — Mở trình duyệt kiểm tra:
```
http://157.230.45.172:8017
```

### SSH vào VPS thủ công:
```powershell
ssh root@157.230.45.172
```

### Xem log Odoo trên VPS:
```bash
tail -f /var/log/odoo/odoo.log
journalctl -u odoo -f
```

---

## ⚙️ Cấu hình `/etc/odoo.conf` trên VPS

```ini
[options]
admin_passwd = <master_password>
db_host = localhost
db_port = 5432
db_user = odoo17
db_name = odoo_production
db_password = <db_password>
addons_path = /opt/odoo/custom,
              /opt/odoo/oca/server-ux,
              /opt/odoo/oca/server-tools,
              /opt/odoo/oca/reporting-engine,
              /opt/odoo/oca/account-financial-tools,
              /opt/odoo/oca/account-financial-reporting,
              /opt/odoo/community/addons,
              /opt/odoo/community/odoo/addons
xmlrpc_port = 8017
logfile = /var/log/odoo/odoo.log
log_level = info
bin_path = /usr/local/bin
data_dir = /opt/odoo/data
workers = 0
max_cron_threads = 1
gevent_port = 8072
```

---

## 📦 Cấu trúc thư mục trên VPS

```
/opt/odoo/
├── community/          ← Odoo 17 source (clone từ GitHub Odoo)
├── custom_repo/        ← Repo này (clone từ GitHub Kyslab/odooHS)
│   └── custom/
│       └── l10n_vn_cash_manager/
├── custom -> custom_repo/custom   ← symlink
├── oca/
│   ├── server-ux/              ← date_range, date_range_account
│   ├── server-tools/
│   ├── reporting-engine/       ← report_xlsx
│   ├── account-financial-tools/
│   └── account-financial-reporting/  ← account_financial_report
└── data/               ← Session & filestore
```

---

## 🏠 Chạy Local (máy Windows)

### Kết nối Neon Cloud:
```bat
D:\odoo\start_neon.bat
```
→ Truy cập: http://localhost:8017

### Kết nối Local PostgreSQL:
```bat
D:\odoo\start_local.bat
```

---

## Cài đặt lần đầu (cho người mới clone repo)

### Yêu cầu hệ thống

| Phần mềm | Phiên bản |
|---|---|
| **Python** | 3.11.x |
| **Git** | Mới nhất |
| **wkhtmltopdf** | 0.12.6.1 |

### Bước 1 — Clone repo

```bat
cd D:\
git clone https://github.com/Kyslab/odooHS.git odoo
cd D:\odoo
```

### Bước 2 — Clone Odoo Community

```bat
git clone https://github.com/odoo/odoo.git --branch 17.0 --depth 1 community
```

### Bước 3 — Cài thư viện Python

```bat
pip install -r D:\odoo\community\requirements.txt
pip install psycopg2-binary
```

### Bước 4 — Tạo file cấu hình

```bat
copy D:\odoo\odoo_neon.conf.example D:\odoo\odoo_neon.conf
```

Mở `odoo_neon.conf`, điền thông tin nhận từ chủ repo:
```ini
admin_passwd = ← mật khẩu master Odoo (hỏi chủ repo)
db_password  = ← mật khẩu Neon database (hỏi chủ repo)
```

### Bước 5 — Tạo thư mục data

```bat
mkdir D:\odoo\data
```

### Bước 6 — Chạy Odoo

Bấm đúp vào **`D:\odoo\start_neon.bat`** → truy cập http://localhost:8017

---

## Module l10n_vn_cash_manager

| Mục menu | Chức năng |
|---|---|
| **Tien Mat → So Quy Tien Mat** | Danh sách phiếu thu/chi tiền mặt (TK 111x) |
| **Tien Mat → Thu Tien → Thu Tien Nhanh** | Tạo phiếu thu nhanh |
| **Tien Mat → Thu Tien → Thu Tien Theo Hoa Don** | Thu tiền + đối trừ TK 131 tự động |
| **Tien Mat → Chi Tien → Chi Tien Nhanh** | Tạo phiếu chi nhanh |
| **Tien Mat → Chi Tien → Tra Tien Theo Hoa Don** | Trả tiền + đối trừ TK 331 tự động |
| **Ngan Hang → So Ngan Hang** | Danh sách phiếu thu/chi ngân hàng (TK 112x) |
| **Ngan Hang → Thu Tien → Thu Tien Nhanh** | Thu tiền ngân hàng nhanh |
| **Ngan Hang → Thu Tien → Thu Tien Theo Hoa Don** | Thu NH + đối trừ TK 131 |
| **Ngan Hang → Chi Tien → Chi Tien Nhanh** | Chi tiền ngân hàng nhanh |
| **Ngan Hang → Chi Tien → Tra Tien Theo Hoa Don** | Trả NH + đối trừ TK 331 |

---

## Lưu ý

- ⚠️ VPS dùng chung — thay đổi ảnh hưởng tất cả mọi người
- 🔒 Không commit file chứa password lên Git (`odoo_neon.conf`, `odoo.conf`)
- 🚀 Dùng `deploy_vps.bat` để deploy code mới lên VPS
- 🔁 Chỉ chạy **1 instance** tại một thời điểm (cùng port 8017)
