@echo off
:: 检查管理员权限，没有则自动提权
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
:: 以管理员身份执行 PowerShell 脚本
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0mac_changer.ps1"
pause
