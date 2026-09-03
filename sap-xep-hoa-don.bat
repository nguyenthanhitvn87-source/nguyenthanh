@echo off
chcp 65001 >nul
title Sap xep hoa don theo danh sach Excel
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0sap-xep-hoa-don.ps1"
if errorlevel 1 pause
