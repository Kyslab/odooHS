@echo off
chcp 65001 > nul
python "D:\odoo\scripts\vaoKHGo.py" vaoKHGo
if errorlevel 1 (
    echo.
    echo Nhan phim bat ky de dong...
    pause > nul
)
