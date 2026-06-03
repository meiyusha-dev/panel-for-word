# ============================================================
# Word Panel Word Watcher
# Monitors for WINWORD.EXE launches and starts dict-server
# Started at logon via scheduled task (LogonTrigger)
# ============================================================

$vbs     = 'C:\OfficeAddins\launch-dict-server.vbs'
$port    = 8642

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

$query   = "SELECT * FROM __InstanceCreationEvent WITHIN 3 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = 'WINWORD.EXE'"
$watcher = New-Object System.Management.ManagementEventWatcher($query)
$watcher.Start()

try {
    while ($true) {
        try {
            $null = $watcher.WaitForNextEvent()
            Start-DictServer
        } catch [System.Management.ManagementException] {
            Start-Sleep 10
            try { $watcher.Start() } catch {}
        } catch {
            Start-Sleep 5
        }
    }
} finally {
    try { $watcher.Stop(); $watcher.Dispose() } catch {}
}
