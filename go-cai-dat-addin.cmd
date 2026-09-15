@echo off
chcp 65001 >nul
title Go cai dat Add-in Chuan the thuc ND 30
echo Dang go dang ky add-in khoi Word...
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo [LOI] Chua tim thay Node.js tren may.
    pause
    exit /b 1
)

npx --yes office-addin-dev-settings unregister "%~dp0manifest.xml"

echo.
echo Da go xong. Khoi dong lai Word de ap dung.
pause
