#!/bin/bash
# ============================================================
# Odoo 17 VPS Install Script — Ubuntu 22.04
# Repo: https://github.com/Kyslab/odooHS
# Chạy với: bash install_vps.sh
# ============================================================
set -e

ODOO_DIR="/opt/odoo"
ODOO_USER="odoo"
ODOO_PORT="8017"
DB_USER="odoo17"
DB_PASS="Odoo17@VPS2024"
DB_NAME="odoo_production"
GITHUB_REPO="https://github.com/Kyslab/odooHS.git"
MASTER_PASS="admin_master_2024"

echo "============================================"
echo "  Odoo 17 VPS Auto-Install Script"
echo "  $(date)"
echo "============================================"

# ── 1. Cập nhật hệ thống ─────────────────────────────────────
echo ""
echo "[1/10] Cập nhật Ubuntu..."
apt-get update -qq
apt-get upgrade -y -qq
apt-get install -y -qq \
    git curl wget gnupg2 \
    python3.11 python3.11-dev python3.11-venv python3-pip \
    build-essential libssl-dev libffi-dev \
    libxml2-dev libxslt1-dev libjpeg-dev libpng-dev \
    libpq-dev libldap2-dev libsasl2-dev \
    node-less npm \
    xfonts-75dpi xfonts-base

# ── 2. Cài PostgreSQL 15 ──────────────────────────────────────
echo ""
echo "[2/10] Cài PostgreSQL 15..."
if ! command -v psql &> /dev/null; then
    curl -fsSL https://www.postgresql.org/media/keys/ACCC4CF8.asc | gpg --dearmor -o /etc/apt/trusted.gpg.d/postgresql.gpg
    echo "deb http://apt.postgresql.org/pub/repos/apt $(lsb_release -cs)-pgdg main" > /etc/apt/sources.list.d/pgdg.list
    apt-get update -qq
    apt-get install -y -qq postgresql-15
fi
systemctl start postgresql
systemctl enable postgresql

# Tạo PostgreSQL user cho Odoo
echo "[2/10] Tạo PostgreSQL user..."
su -c "psql -c \"SELECT 1 FROM pg_roles WHERE rolname='$DB_USER'\" | grep -q 1 || psql -c \"CREATE USER $DB_USER WITH PASSWORD '$DB_PASS' CREATEDB;\"" postgres

# ── 3. Cài wkhtmltopdf ───────────────────────────────────────
echo ""
echo "[3/10] Cài wkhtmltopdf..."
if ! command -v wkhtmltopdf &> /dev/null; then
    wget -q https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6.1-2/wkhtmltox_0.12.6.1-2.jammy_amd64.deb
    apt-get install -y -qq ./wkhtmltox_0.12.6.1-2.jammy_amd64.deb
    rm -f wkhtmltox_0.12.6.1-2.jammy_amd64.deb
fi

# ── 4. Tạo user hệ thống cho Odoo ────────────────────────────
echo ""
echo "[4/10] Tạo system user 'odoo'..."
if ! id "$ODOO_USER" &>/dev/null; then
    useradd -m -d "$ODOO_DIR" -U -r -s /bin/bash "$ODOO_USER"
fi
mkdir -p "$ODOO_DIR"
chown -R "$ODOO_USER:$ODOO_USER" "$ODOO_DIR"

# ── 5. Clone Odoo Community ───────────────────────────────────
echo ""
echo "[5/10] Clone Odoo 17 Community (có thể mất 5-10 phút)..."
if [ ! -d "$ODOO_DIR/community" ]; then
    sudo -u "$ODOO_USER" git clone https://github.com/odoo/odoo.git \
        --branch 17.0 --depth 1 \
        "$ODOO_DIR/community"
else
    echo "  → Đã có, bỏ qua."
fi

# ── 6. Clone custom module ────────────────────────────────────
echo ""
echo "[6/10] Clone custom module từ $GITHUB_REPO..."
if [ ! -d "$ODOO_DIR/custom" ]; then
    sudo -u "$ODOO_USER" git clone "$GITHUB_REPO" "$ODOO_DIR/custom_repo"
    # Custom modules nằm trong thư mục custom/ của repo
    if [ -d "$ODOO_DIR/custom_repo/custom" ]; then
        ln -sf "$ODOO_DIR/custom_repo/custom" "$ODOO_DIR/custom"
    else
        ln -sf "$ODOO_DIR/custom_repo" "$ODOO_DIR/custom"
    fi
else
    echo "  → Đã có, bỏ qua."
fi

# ── 7. Cài Python dependencies ────────────────────────────────
echo ""
echo "[7/10] Cài Python dependencies..."
sudo -u "$ODOO_USER" python3.11 -m pip install --quiet --upgrade pip
sudo -u "$ODOO_USER" python3.11 -m pip install --quiet -r "$ODOO_DIR/community/requirements.txt"
sudo -u "$ODOO_USER" python3.11 -m pip install --quiet psycopg2-binary

# ── 8. Tạo thư mục data & file cấu hình ─────────────────────
echo ""
echo "[8/10] Tạo cấu hình Odoo..."
mkdir -p "$ODOO_DIR/data"
chown -R "$ODOO_USER:$ODOO_USER" "$ODOO_DIR/data"

cat > /etc/odoo.conf << EOF
[options]
admin_passwd = $MASTER_PASS
db_host = localhost
db_port = 5432
db_user = $DB_USER
db_password = $DB_PASS
db_name = $DB_NAME
addons_path = $ODOO_DIR/custom,$ODOO_DIR/community/addons,$ODOO_DIR/community/odoo/addons
xmlrpc_port = $ODOO_PORT
logfile = /var/log/odoo/odoo.log
log_level = info
bin_path = /usr/local/bin
data_dir = $ODOO_DIR/data
workers = 2
max_cron_threads = 1
EOF

# Log directory
mkdir -p /var/log/odoo
chown -R "$ODOO_USER:$ODOO_USER" /var/log/odoo
chmod 640 /etc/odoo.conf
chown "$ODOO_USER:$ODOO_USER" /etc/odoo.conf

# ── 9. Tạo systemd service ────────────────────────────────────
echo ""
echo "[9/10] Tạo systemd service..."
cat > /etc/systemd/system/odoo.service << EOF
[Unit]
Description=Odoo 17 Server
After=network.target postgresql.service
Requires=postgresql.service

[Service]
Type=simple
SyslogIdentifier=odoo
PermissionsStartOnly=true
User=$ODOO_USER
Group=$ODOO_USER
ExecStart=/usr/bin/python3.11 $ODOO_DIR/community/odoo-bin -c /etc/odoo.conf
StandardOutput=journal+console
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable odoo

# ── 10. Tạo database & khởi động Odoo ───────────────────────
echo ""
echo "[10/10] Khởi động Odoo lần đầu..."

# Tạo database
su -c "psql -c \"SELECT 1 FROM pg_database WHERE datname='$DB_NAME'\" | grep -q 1 || psql -c \"CREATE DATABASE $DB_NAME OWNER $DB_USER;\"" postgres

# Khởi động service
systemctl start odoo

echo ""
echo "============================================"
echo "  ✅ Cài đặt HOÀN TẤT!"
echo "============================================"
echo ""
echo "  🌐 Truy cập Odoo tại:"
echo "     http://$(curl -s ifconfig.me):$ODOO_PORT"
echo ""
echo "  📋 Thông tin kết nối:"
echo "     Master Password : $MASTER_PASS"
echo "     Database        : $DB_NAME"
echo "     DB User         : $DB_USER"
echo "     DB Password     : $DB_PASS"
echo ""
echo "  📝 Xem log:"
echo "     journalctl -u odoo -f"
echo "     tail -f /var/log/odoo/odoo.log"
echo ""
echo "  🔄 Quản lý service:"
echo "     systemctl start|stop|restart odoo"
echo ""
echo "  ⚠️  Hãy đổi password trong /etc/odoo.conf sau khi cài xong!"
echo "============================================"
