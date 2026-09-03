@echo off
chcp 65001 >nul
title Sap xep va in hoa don
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0sap-xep-in-hoa-don.ps1"
if errorlevel 1 pause
