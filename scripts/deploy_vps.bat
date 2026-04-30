@echo off
title Deploy len VPS
echo ============================================
echo   Deploy custom module len VPS
echo ============================================
echo.

REM -- Buoc 1: Push len GitHub
echo [1/3] Push code len GitHub...
git -C D:\odoo add custom/
git -C D:\odoo commit -m "Update: %date% %time%" 2>nul || echo (Khong co gi moi de commit)
git -C D:\odoo push origin main
if errorlevel 1 (
    echo LAI: Khong push duoc len GitHub!
    pause
    exit /b 1
)
echo OK - Da push len GitHub

REM -- Buoc 2: VPS pull va restart
echo.
echo [2/3] VPS dang pull code moi va restart Odoo...
"C:\Program Files\PuTTY\plink.exe" -ssh -pw "PGzU6JDFABfea@R" -batch root@157.230.45.172 "cd /opt/odoo/custom_repo && git pull && systemctl restart odoo && sleep 5 && systemctl is-active odoo"
if errorlevel 1 (
    echo LAI: VPS khong cap nhat duoc!
    pause
    exit /b 1
)

echo.
echo [3/3] Hoan tat!
echo ============================================
echo   Truy cap: http://157.230.45.172:8017
echo ============================================
echo.
pause
