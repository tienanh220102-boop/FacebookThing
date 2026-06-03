@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===================================================
echo    FPT University - Lay toan bo lich su bai viet
echo    (Chi can chay 1 lan duy nhat)
echo    Uoc tinh: 3-5 phut
echo ===================================================
echo.
python lich_su.py
echo.
pause
