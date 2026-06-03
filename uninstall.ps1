# ============================================================
# Word Panel Add-in Uninstaller
# Run as Administrator (launch via uninstall.bat)
# ============================================================

$ErrorActionPreference = "Continue"
$addinFolder  = "C:\OfficeAddins"
$shareName    = "OfficeAddins"
$taskName     = "WordPanel-DictServer"
$catalogBase  = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$uncPath      = "\\$env:COMPUTERNAME\$shareName"

function Write-Step($msg) { Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "    WARN: $msg" -ForegroundColor Yellow }
function Write-Fail($msg) { Write-Host "    ERROR: $msg" -ForegroundColor Red }

Write-Host "======================================" -ForegroundColor Yellow
Write-Host "  Word Panel Add-in Uninstaller       " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow

# Warn if Word is running
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host ""
    Write-Host "  [WARNING] Microsoft Word is currently running." -ForegroundColor Red
    Write-Host "  Please close Word before continuing." -ForegroundColor Red
    Write-Host ""
    $ans = Read-Host "  Have you closed Word? Continue? [Y/N]"
    if ($ans -notmatch '^[Yy]') {
        Write-Host "Uninstall cancelled." -ForegroundColor Yellow
        pause
        exit 0
    }
}

$results = @{
    task     = ""
    urlacl   = ""
    share    = ""
    catalog  = ""
    webview2 = ""
    folder   = ""
}

# 0. Kill any running dict-server / word-watcher processes
Write-Step "Stopping running dict-server processes..."
try {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match 'dict-server|word-watcher' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Write-OK "Processes stopped"
} catch {
    Write-Warn "Could not stop processes: $($_.Exception.Message)"
}

# 1. Stop and delete scheduled task
Write-Step "Removing scheduled task..."
try {
    schtasks /End /TN $taskName 2>&1 | Out-Null
    schtasks /Delete /TN $taskName /F 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-OK "Task '$taskName' deleted"
        $results.task = "OK"
    } else {
        Write-Warn "Task '$taskName' not found (skipped)"
        $results.task = "skip"
    }
} catch {
    Write-Fail "Failed to delete task: $($_.Exception.Message)"
    $results.task = "fail"
}

# 2. Remove URL ACL
Write-Step "Removing URL ACL..."
try {
    netsh http delete urlacl url=http://localhost:8642/ 2>&1 | Out-Null
    Write-OK "URL ACL (localhost:8642) removed"
    $results.urlacl = "OK"
} catch {
    Write-Fail "Failed to remove URL ACL: $($_.Exception.Message)"
    $results.urlacl = "fail"
}

# 3. Remove SMB share
Write-Step "Removing SMB share..."
try {
    if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
        Remove-SmbShare -Name $shareName -Force
        Write-OK "Share '$shareName' removed"
        $results.share = "OK"
    } else {
        Write-Warn "Share '$shareName' not found (skipped)"
        $results.share = "skip"
    }
} catch {
    Write-Fail "Failed to remove SMB share: $($_.Exception.Message)"
    $results.share = "fail"
}

# 4. Remove Word trusted catalog registry entry
Write-Step "Removing Word add-in catalog registry entry..."
try {
    $entries = Get-ChildItem -Path $catalogBase -ErrorAction SilentlyContinue |
        Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Url -eq $uncPath }
    if ($entries) {
        foreach ($entry in $entries) {
            Remove-Item $entry.PSPath -Recurse -Force
        }
        Write-OK "Catalog entry removed: $uncPath"
        $results.catalog = "OK"
    } else {
        Write-Warn "Catalog entry not found (skipped)"
        $results.catalog = "skip"
    }
} catch {
    Write-Fail "Failed to remove catalog entry: $($_.Exception.Message)"
    $results.catalog = "fail"
}

# 5. WebView2 registry - always skip (may be used by other Office add-ins)
Write-Step "WebView2 registry..."
Write-Warn "Win32WebView2 registry skipped (may be used by other Office add-ins)"
$results.webview2 = "skip"

# 6. Remove add-in folder
Write-Step "Removing add-in folder..."
try {
    if (Test-Path $addinFolder) {
        Remove-Item -Path $addinFolder -Recurse -Force
        Write-OK "$addinFolder removed"
        $results.folder = "OK"
    } else {
        Write-Warn "$addinFolder not found (skipped)"
        $results.folder = "skip"
    }
} catch {
    Write-Fail "Failed to remove folder: $($_.Exception.Message)"
    $results.folder = "fail"
}

# Summary
Write-Host ""
Write-Host "======================================" -ForegroundColor Yellow
Write-Host "  Uninstall Summary                  " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow
$hasFail = $false
$labels = @{
    task     = "Task       "
    urlacl   = "URL ACL    "
    share    = "SMB Share  "
    catalog  = "Catalog    "
    webview2 = "WebView2   "
    folder   = "Folder     "
}
foreach ($key in @("task", "urlacl", "share", "catalog", "webview2", "folder")) {
    $val   = $results[$key]
    $label = $labels[$key]
    if ($val -eq "OK") {
        Write-Host ("  {0}: OK" -f $label) -ForegroundColor Green
    } elseif ($val -eq "skip") {
        Write-Host ("  {0}: skipped" -f $label) -ForegroundColor Yellow
    } else {
        Write-Host ("  {0}: FAILED" -f $label) -ForegroundColor Red
        $hasFail = $true
    }
}
Write-Host ""
if ($hasFail) {
    Write-Host "  Some steps failed. Check the errors above." -ForegroundColor Red
} else {
    Write-Host "  Uninstall complete." -ForegroundColor Green
    Write-Host ""
    Write-Host "  Please restart Word to confirm the add-in has been removed." -ForegroundColor White
}
Write-Host ""
pause
if ($hasFail) { exit 1 }
