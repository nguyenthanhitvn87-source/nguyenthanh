@echo off
chcp 65001 >nul
title Sap xep hoa don theo danh sach Excel
set "PS1=%~dp0sap-xep-hoa-don.ps1"
if not exist "%PS1%" (
  echo.
  echo KHONG TIM THAY FILE: %PS1%
  echo Hay de file sap-xep-hoa-don.ps1 va file .bat nay trong CUNG MOT thu muc,
  echo va giu nguyen ten file khi tai ve ^(khong de thanh "... ^(1^).ps1"^).
  echo.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath \"%PS1%\"" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1%"
echo.
echo === Cong cu da dong. Ma tra ve: %errorlevel% ===
echo Neu co loi, xem file loi-chay.txt trong thu muc nay.
pause
