@echo off
title Odoo 17 Server
echo ========================================
echo    Odoo 17 Community - Starting...
echo ========================================
echo.
echo Server se chay tai: http://localhost:8017
echo De dung server: nhan Ctrl+C
echo.
cd /d D:\odoo\community
python odoo-bin -c ..\odoo.conf
pause
