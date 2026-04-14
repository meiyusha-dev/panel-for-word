@echo off
:: Word Panel アドイン インストーラー（単体版）
:: このファイルをダブルクリックして管理者として実行してください

:: 管理者権限チェック・昇格
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限で再起動しています...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: PowerShell スクリプトを一時ファイルにダウンロードして実行
set "PS1=%TEMP%\word-panel-install-%RANDOM%.ps1"
echo スクリプトを取得しています...
powershell -ExecutionPolicy Bypass -Command ^
  "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/otayoshino/panel-for-word/master/install.ps1' -OutFile '%PS1%' -UseBasicParsing"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] スクリプトのダウンロードに失敗しました。
    echo         インターネット接続を確認してください。
    pause
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%PS1%"
del "%PS1%" >nul 2>&1
