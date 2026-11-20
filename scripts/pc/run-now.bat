@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0bmts-collect.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0bmts-send.ps1"
pause
