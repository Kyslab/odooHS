@echo off
title Odoo 17 - Neon Cloud Database
cd /d D:\odoo
echo Dang khoi dong Odoo (Neon Cloud)...
echo Truy cap: http://localhost:8017
echo Nhan Ctrl+C de dung.
python community/odoo-bin -c odoo_neon.conf
pause
