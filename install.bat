@echo off
:: Word Panel アドイン インストーラー
:: このファイルをダブルクリックして管理者として実行してください

:: 管理者権限チェック・昇格
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限で再起動しています...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: PowerShell スクリプトを実行
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"
