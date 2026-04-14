# ============================================================
# Word Panel アドイン インストーラー
# 管理者権限で実行してください（install.bat から起動してください）
# ============================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$manifest_url = "https://raw.githubusercontent.com/otayoshino/panel-for-word/master/manifest.xml"
$addinFolder  = "C:\OfficeAddins"
$shareName    = "OfficeAddins"

function Write-Step($msg) { Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }
function Write-Fail($msg) { Write-Host "    ERROR: $msg" -ForegroundColor Red }

Write-Host "======================================" -ForegroundColor Yellow
Write-Host "  Word Panel アドイン インストーラー  " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow

# 1. フォルダ作成
Write-Step "フォルダを作成しています..."
New-Item -ItemType Directory -Path $addinFolder -Force | Out-Null
Write-OK "$addinFolder を作成しました"

# 2. manifest.xml をダウンロード
Write-Step "manifest.xml をダウンロードしています..."
try {
    Invoke-WebRequest -Uri $manifest_url -OutFile "$addinFolder\manifest.xml" -UseBasicParsing
    Write-OK "manifest.xml を $addinFolder に保存しました"
} catch {
    Write-Fail "ダウンロードに失敗しました。インターネット接続を確認してください。"
    Write-Fail $_.Exception.Message
    pause; exit 1
}

# 3. フォルダを共有
Write-Step "フォルダを共有しています..."
if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
    Remove-SmbShare -Name $shareName -Force | Out-Null
}
New-SmbShare -Name $shareName -Path $addinFolder -FullAccess "Everyone" | Out-Null
$uncPath = "\\$env:COMPUTERNAME\$shareName"
Write-OK "共有パス: $uncPath"

# 4. Office 2019 向け WebView2 対応レジストリ
Write-Step "WebView2 対応レジストリを設定しています..."
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF" /v "Win32WebView2" /t REG_DWORD /d 1 /f | Out-Null
Write-OK "WebView2 レジストリを設定しました"

# 5. Word の信頼できるカタログに登録
Write-Step "Word のアドインカタログを登録しています..."
$catalogBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

# 既存の同じUNCパスのカタログがあればスキップ
$existing = Get-ChildItem -Path $catalogBase -ErrorAction SilentlyContinue |
    Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Url -eq $uncPath }

if ($existing) {
    Write-OK "カタログは既に登録済みです"
} else {
    $guid = [System.Guid]::NewGuid().ToString("B").ToUpper()
    $regPath = "$catalogBase\$guid"
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name "Id"    -Value $guid
    Set-ItemProperty -Path $regPath -Name "Url"   -Value $uncPath
    Set-ItemProperty -Path $regPath -Name "Flags" -Value 1
    Write-OK "カタログを登録しました: $uncPath"
}

# 完了
Write-Host "`n======================================" -ForegroundColor Yellow
Write-Host "  インストール完了！                  " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow
Write-Host @"

次の手順でアドインを追加してください：

  1. Word を完全に再起動する
  2. 「開発」タブ → 「アドイン」をクリック
     ※「開発」タブがない場合：
       「ファイル」→「オプション」→「リボンのユーザー設定」→「開発」にチェック
  3. 「共有フォルダ」タブ → 「Word Panel」→ 「追加」

"@ -ForegroundColor White

pause
