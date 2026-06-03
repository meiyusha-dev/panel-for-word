# ============================================================
# Word Panel - Scheduled Task Repair Script
# Registers LogonTrigger task that starts word-watcher.ps1
# Run as Administrator
# ============================================================

$taskName    = "WordPanel-DictServer"
$addinFolder = "C:\OfficeAddins"

$taskXml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Word Panel dict server watcher (starts with Word)</Description></RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>wscript.exe</Command>
      <Arguments>//B //NoLogo "C:\OfficeAddins\launch-word-watcher.vbs"</Arguments>
    </Exec>
  </Actions>
</Task>
'@

try {
    Write-Host "Repairing scheduled task..." -ForegroundColor Cyan

    # Stop any running watcher/dict-server processes
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match 'dict-server|word-watcher' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

    $tmpXml = "$env:TEMP\wordpanel-task.xml"
    $taskXml | Set-Content -Path $tmpXml -Encoding Unicode
    $result = schtasks /Create /TN $taskName /XML $tmpXml /F 2>&1
    Remove-Item $tmpXml -ErrorAction SilentlyContinue

    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task) {
        Write-Host "    OK: Task registered (LogonTrigger)" -ForegroundColor Green
    } else {
        Write-Host "    NG: Task registration failed: $result" -ForegroundColor Red
    }

    # Update launch-dict-server.vbs
    $vbs = 'CreateObject("WScript.Shell").Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\OfficeAddins\dict-server.ps1""", 0, False'
    Set-Content -Path "$addinFolder\launch-dict-server.vbs" -Value $vbs -Encoding ASCII
    Write-Host "    OK: launch-dict-server.vbs updated" -ForegroundColor Green

    # Update/create launch-word-watcher.vbs
    $watcherVbs = 'CreateObject("WScript.Shell").Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\OfficeAddins\word-watcher.ps1""", 0, False'
    Set-Content -Path "$addinFolder\launch-word-watcher.vbs" -Value $watcherVbs -Encoding ASCII
    Write-Host "    OK: launch-word-watcher.vbs updated" -ForegroundColor Green

    # Download latest word-watcher.ps1
    try {
        $watcherUrl = "https://raw.githubusercontent.com/meiyusha-dev/panel-for-word/master/word-watcher.ps1"
        Invoke-WebRequest -Uri $watcherUrl -OutFile "$addinFolder\word-watcher.ps1" -UseBasicParsing
        Write-Host "    OK: word-watcher.ps1 updated" -ForegroundColor Green
    } catch {
        # Use local copy if download fails
        $localWatcher = Join-Path (Split-Path $MyInvocation.MyCommand.Path) "word-watcher.ps1"
        if (Test-Path $localWatcher) {
            Copy-Item $localWatcher "$addinFolder\word-watcher.ps1" -Force
            Write-Host "    OK: word-watcher.ps1 copied from local" -ForegroundColor Green
        } else {
            Write-Host "    WARN: word-watcher.ps1 download failed and no local copy found" -ForegroundColor Yellow
        }
    }

    # Start the task immediately
    if ($task) {
        Start-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        Write-Host "    OK: Task started" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Done. Open Word to confirm dict-server starts." -ForegroundColor Yellow
} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
} finally {
    Read-Host "Press Enter to close"
}
