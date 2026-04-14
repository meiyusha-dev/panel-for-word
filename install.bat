@echo off
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
if not exist "%~dp0install.ps1" (
    echo [ERROR] install.ps1 not found.
    pause
    exit /b 1
)
powershell -ExecutionPolicy Bypass -Command "Unblock-File -Path '%~dp0install.ps1' -ErrorAction SilentlyContinue"
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if %errorlevel% neq 0 (
    echo [ERROR] Installation failed.
    pause
)
