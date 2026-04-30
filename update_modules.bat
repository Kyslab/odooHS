@echo off
title Cap nhat danh sach Module Odoo
echo ========================================
echo    Cap nhat danh sach Module Odoo 17
echo ========================================
echo.
echo [!] Hay TAT Odoo dang chay truoc khi chay file nay!
echo.
pause
echo.
echo Dang cap nhat module list...
cd /d D:\odoo\community
python odoo-bin -c ..\odoo.conf --stop-after-init -u base
echo.
echo ========================================
echo XONG! Hay khoi dong lai Odoo (start_odoo.bat)
echo ========================================
pause
