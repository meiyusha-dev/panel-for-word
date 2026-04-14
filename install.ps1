# ============================================================
# Word Panel アドイン インストーラー
# 管理者権限で実行してください（install.bat から起動してください）
# ============================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$manifest_url = "https://raw.githubusercontent.com/otayoshino/panel-for-word/master/manifest.xml"
$dictServer_url = "https://raw.githubusercontent.com/otayoshino/panel-for-word/master/dict-server.ps1"
$dictBase_url = "https://otayoshino.github.io/panel-for-word/dict"
$addinFolder  = "C:\OfficeAddins"
$dictFolder   = "$addinFolder\dict"
$shareName    = "OfficeAddins"
$taskName     = "WordPanel-DictServer"

$dictFiles = @(
    'base.dat.gz', 'cc.dat.gz', 'check.dat.gz',
    'tid.dat.gz', 'tid_map.dat.gz', 'tid_pos.dat.gz',
    'unk.dat.gz', 'unk_char.dat.gz', 'unk_compat.dat.gz',
    'unk_invoke.dat.gz', 'unk_map.dat.gz', 'unk_pos.dat.gz'
)

function Write-Step($msg) { Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }
function Write-Fail($msg) { Write-Host "    ERROR: $msg" -ForegroundColor Red }

Write-Host "======================================" -ForegroundColor Yellow
Write-Host "  Word Panel アドイン インストーラー  " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow

try {
    # 1. フォルダ作成
    Write-Step "フォルダを作成しています..."
    New-Item -ItemType Directory -Path $addinFolder -Force | Out-Null
    New-Item -ItemType Directory -Path $dictFolder  -Force | Out-Null
    Write-OK "$addinFolder を作成しました"

    # 2. manifest.xml をダウンロード
    Write-Step "manifest.xml をダウンロードしています..."
    Invoke-WebRequest -Uri $manifest_url -OutFile "$addinFolder\manifest.xml" -UseBasicParsing
    Write-OK "manifest.xml を保存しました"

    # 3. 辞書ファイルをダウンロード（12ファイル、計約17MB）
    Write-Step "辞書ファイルをダウンロードしています（約17MB）..."
    $i = 1
    foreach ($f in $dictFiles) {
        Write-Host "    ($i/$($dictFiles.Count)) $f ..." -NoNewline
        Invoke-WebRequest -Uri "$dictBase_url/$f" -OutFile "$dictFolder\$f" -UseBasicParsing
        Write-Host " OK" -ForegroundColor Green
        $i++
    }
    Write-OK "辞書ファイルを $dictFolder に保存しました"

    # 4. dict-server.ps1 をダウンロード
    Write-Step "辞書サーバースクリプトを配置しています..."
    Invoke-WebRequest -Uri $dictServer_url -OutFile "$addinFolder\dict-server.ps1" -UseBasicParsing
    Write-OK "dict-server.ps1 を保存しました"

    # 5. フォルダを共有
    Write-Step "フォルダを共有しています..."
    if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
        Remove-SmbShare -Name $shareName -Force | Out-Null
    }
    New-SmbShare -Name $shareName -Path $addinFolder -FullAccess "Everyone" | Out-Null
    $uncPath = "\\$env:COMPUTERNAME\$shareName"
    Write-OK "共有パス: $uncPath"

    # 6. Office 2019 向け WebView2 対応レジストリ
    Write-Step "WebView2 対応レジストリを設定しています..."
    reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF" /v "Win32WebView2" /t REG_DWORD /d 1 /f | Out-Null
    Write-OK "WebView2 レジストリを設定しました"

    # 7. Word の信頼できるカタログに登録
    Write-Step "Word のアドインカタログを登録しています..."
    $catalogBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
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

    # 8. 辞書サーバーをログオン時自動起動タスクに登録
    Write-Step "辞書サーバーをスタートアップタスクに登録しています..."
    $action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
                    -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$addinFolder\dict-server.ps1`""
    $trigger  = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
        -Settings $settings -RunLevel Highest -Force | Out-Null
    Write-OK "スケジュールタスク '$taskName' を登録しました"

    # 9. 辞書サーバーを今すぐ起動
    Write-Step "辞書サーバーを起動しています..."
    Start-ScheduledTask -TaskName $taskName
    Start-Sleep -Seconds 2
    $port = Get-NetTCPConnection -LocalPort 8642 -State Listen -ErrorAction SilentlyContinue
    if ($port) {
        Write-OK "辞書サーバーが localhost:8642 で起動しました"
    } else {
        Write-Host "    INFO: サーバーは次回ログオン時に自動起動します" -ForegroundColor Yellow
    }

    # 完了
    Write-Host "`n======================================" -ForegroundColor Yellow
    Write-Host "  インストール完了！                  " -ForegroundColor Yellow
    Write-Host "======================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "次の手順でアドインを追加してください：" -ForegroundColor White
    Write-Host ""
    Write-Host "  1. Word を完全に再起動する" -ForegroundColor White
    Write-Host "  2. 「開発」タブ → 「アドイン」をクリック" -ForegroundColor White
    Write-Host "     ※「開発」タブがない場合：" -ForegroundColor White
    Write-Host "       「ファイル」→「オプション」→「リボンのユーザー設定」→「開発」にチェック" -ForegroundColor White
    Write-Host "  3. 「共有フォルダ」タブ → 「Word Panel」→ 「追加」" -ForegroundColor White
    Write-Host ""

} catch {
    Write-Host ""
    Write-Fail "インストール中にエラーが発生しました："
    Write-Fail $_.Exception.Message
    Write-Host ""
    exit 1
} finally {
    pause
}
