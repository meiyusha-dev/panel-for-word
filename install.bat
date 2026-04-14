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

:: install.ps1 の存在確認
if not exist "%~dp0install.ps1" (
    echo.
    echo [ERROR] install.ps1 が見つかりません。
    echo         install.bat と install.ps1 を同じフォルダに置いてください。
    echo.
    pause
    exit /b 1
)

:: ダウンロードブロックを解除してから実行
powershell -ExecutionPolicy Bypass -Command "Unblock-File -Path '%~dp0install.ps1' -ErrorAction SilentlyContinue"

:: PowerShell スクリプトを実行
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] インストールに失敗しました。
    pause
)
