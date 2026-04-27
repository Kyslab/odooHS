@echo off
title Odoo 17 - Local Database
cd /d D:\odoo
echo Dang khoi dong Odoo (Local PostgreSQL)...
echo Truy cap: http://localhost:8017
echo Nhan Ctrl+C de dung.
python community/odoo-bin -c odoo.conf
pause
