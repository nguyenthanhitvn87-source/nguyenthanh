@echo off
rem ============================================================
rem  Chay chuong trinh voi quyen Administrator (Windows)
rem  Cach dung:  chay-admin.bat [cong] [trang]
rem  Vi du:      chay-admin.bat 8080 lich-bieu.html
rem ============================================================
setlocal EnableExtensions
chcp 65001 >nul 2>&1
title Chay chuong trinh - Administrator

set "PORT=%~1"
if "%PORT%"=="" set "PORT=8080"
set "TRANG=%~2"
if "%TRANG%"=="" set "TRANG=index.html"

rem --- 1. Chua co quyen admin thi tu xin qua UAC roi chay lai ---
net session >nul 2>&1
if errorlevel 1 (
  echo Dang xin quyen Administrator, bam "Yes" o hop thoai UAC...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList '%PORT%','%TRANG%' -Verb RunAs"
  exit /b
)

cd /d "%~dp0"
echo.
echo === Dang chay voi quyen Administrator ===
echo Thu muc : %CD%
echo Cong    : %PORT%
echo.

rem --- 2. Mo cong tren tuong lua cho mang noi bo (dien thoai vao duoc) ---
set "RULE=Chay chuong trinh %PORT%"
netsh advfirewall firewall show rule name="%RULE%" >nul 2>&1
if errorlevel 1 (
  netsh advfirewall firewall add rule name="%RULE%" dir=in action=allow protocol=TCP localport=%PORT% profile=private >nul 2>&1
  if errorlevel 1 (
    echo [!] Khong them duoc luat tuong lua, may khac co the khong vao duoc.
  ) else (
    echo [OK] Da mo cong %PORT% tren tuong lua ^(mang Private^).
  )
) else (
  echo [OK] Luat tuong lua cho cong %PORT% da co san.
)

rem --- 3. Tim mot may chu tinh co san tren may ---
set "SERVER="
where py >nul 2>&1 && set "SERVER=py -3 -m http.server %PORT% --bind 0.0.0.0"
if not defined SERVER (
  where python >nul 2>&1 && set "SERVER=python -m http.server %PORT% --bind 0.0.0.0"
)
if not defined SERVER (
  where npx >nul 2>&1 && set "SERVER=npx --yes http-server . -p %PORT% -a 0.0.0.0"
)

if not defined SERVER (
  echo [!] Khong thay Python lan Node tren may.
  echo     Mo thang file %TRANG% bang trinh duyet.
  start "" "%~dp0%TRANG%"
  echo.
  pause
  exit /b
)

rem --- 4. Bao dia chi de may khac trong nha go vao ---
echo.
echo Mo tren may nay      : http://localhost:%PORT%/%TRANG%
for /f "usebackq delims=" %%A in (`powershell -NoProfile -Command "(Get-NetIPAddress -AddressFamily IPv4 ^| Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } ^| Select-Object -First 1).IPAddress"`) do (
  echo Mo tren dien thoai   : http://%%A:%PORT%/%TRANG%
)
echo.
echo Bam Ctrl+C de dung may chu.
echo.

rem --- 5. Doi may chu len roi mo trinh duyet, sau do chay may chu ---
start "" /min powershell -NoProfile -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:%PORT%/%TRANG%'"
%SERVER%

echo.
echo May chu da dung.
pause
