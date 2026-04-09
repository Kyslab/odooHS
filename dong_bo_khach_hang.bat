@echo off
chcp 65001 > nul
title Dong bo Khach hang MISA -> Odoo

echo.
echo ============================================================
echo   DONG BO KHACH HANG MISA --^> ODOO
echo ============================================================
echo   File Excel : C:\Users\DELL\Downloads\in xong xoa\Danh_sach_khach_hang.xlsx
echo   Odoo       : http://localhost:8017
echo   Log        : D:\odoo\sync_customers.log
echo ============================================================
echo.
echo [*] Dang kiem tra khach hang moi trong file Excel...
echo     (Chi nhung khach hang THEM MOI o cuoi file moi duoc up len)
echo     (Bao gom ca: Ten, Ma KH, SDT, Dia chi, MST, Nhom)
echo.

python D:\odoo\sync_customers.py

echo.
if %errorlevel% == 0 (
    echo [OK] Dong bo thanh cong!
) else (
    echo [LOI] Co loi xay ra, xem chi tiet tai: D:\odoo\sync_customers.log
)

echo.
echo ============================================================
echo Xem log day du tai: D:\odoo\sync_customers.log
echo ============================================================
echo.
pause
