@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===================================================
echo    FPT University News Crawler
echo    Dang thu thap tin tuc...
echo ===================================================
echo.
python scripts\fb_crawler.py
echo.
echo ===================================================
echo    Hoan thanh! File Excel da luu vao:
echo    Trang web FPT\
echo ===================================================
echo.
pause
