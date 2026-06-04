# ============================================================
# Word Panel Word Watcher
# Monitors for WINWORD.EXE launches and starts dict-server
# Started at logon via scheduled task (LogonTrigger)
# ============================================================

$vbs  = 'C:\OfficeAddins\launch-dict-server.vbs'
$port = 8642

function Start-DictServer {
    $inUse = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    if ($inUse) { return }
    if (Test-Path $vbs) {
        Start-Process 'wscript.exe' -ArgumentList "//B //NoLogo `"$vbs`"" -WindowStyle Hidden
    }
}

# If Word is already running at logon, start dict-server immediately
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Start-DictServer
}

# Poll for WINWORD instead of ManagementEventWatcher (WMI COM objects cause
# CLR teardown errors (0xc0000142) when Windows forcibly terminates the process)
$wasRunning = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
while ($true) {
    Start-Sleep 3
    $isRunning = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
    if ($isRunning -and -not $wasRunning) {
        Start-DictServer
    }
    $wasRunning = $isRunning
}
