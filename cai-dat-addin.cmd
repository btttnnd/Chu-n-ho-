@echo off
chcp 65001 >nul
title Cai dat Add-in Chuan the thuc ND 30 cho Word
echo ============================================
echo   CAI DAT ADD-IN "CHUAN THE THUC ND 30"
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo [LOI] Chua tim thay Node.js tren may.
    echo Vui long cai Node.js tai https://nodejs.org ^(ban LTS^),
    echo sau do chay lai file nay.
    echo.
    pause
    exit /b 1
)

echo Da tim thay Node.js. Dang dang ky add-in voi Word...
echo ^(lan dau chay se tu tai cong cu ho tro, mat khoang 30s-1 phut^)
echo.

npx --yes office-addin-dev-settings register "%~dp0manifest.xml"

if errorlevel 1 (
    echo.
    echo [LOI] Dang ky khong thanh cong. Xem thong bao loi o tren.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   DA DANG KY THANH CONG!
echo ============================================
echo.
echo Buoc tiep theo:
echo   1. Dong het cua so Word dang mo ^(neu co^).
echo   2. Mo lai Word.
echo   3. Vao tab Insert (Chen) - Add-ins - My Add-ins.
echo   4. Add-in "Chuan the thuc ND 30" se xuat hien san trong danh sach.
echo.
pause
