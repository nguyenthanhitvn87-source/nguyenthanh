@echo off
chcp 65001 >nul
title In hoa don hang loat
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0in-hoa-don.ps1"
if errorlevel 1 pause
