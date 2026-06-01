# ============================================================
# Word Panel - Scheduled Task Repair Script
# Registers WMI trigger (fires when WINWORD.EXE starts)
# Run as Administrator
# ============================================================

$taskName    = "WordPanel-DictServer"
$addinFolder = "C:\OfficeAddins"

$taskXml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Word Panel dict server (starts with Word only)</Description></RegistrationInfo>
  <Triggers>
    <WMIEventTrigger>
      <Enabled>true</Enabled>
      <Subscription>SELECT * FROM __InstanceCreationEvent WITHIN 3 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = 'WINWORD.EXE'</Subscription>
    </WMIEventTrigger>
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
    <ExecutionTimeLimit>PT12H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>wscript.exe</Command>
      <Arguments>//B //NoLogo "C:\OfficeAddins\launch-dict-server.vbs"</Arguments>
    </Exec>
  </Actions>
</Task>
'@

try {
    Write-Host "Repairing scheduled task..." -ForegroundColor Cyan

    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

    $tmpXml = "$env:TEMP\wordpanel-task.xml"
    $taskXml | Set-Content -Path $tmpXml -Encoding Unicode
    schtasks /Create /TN $taskName /XML $tmpXml /F 2>&1 | Out-Null
    Remove-Item $tmpXml -ErrorAction SilentlyContinue

    $trigger = (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue).Triggers.CimClass.CimClassName
    if ($trigger -eq 'MSFT_TaskEventTrigger') {
        Write-Host "    OK: WMI trigger registered successfully" -ForegroundColor Green
    } else {
        Write-Host "    NG: Trigger registration failed ($trigger)" -ForegroundColor Red
    }

    $vbs = 'CreateObject("WScript.Shell").Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\OfficeAddins\dict-server.ps1""", 0, False'
    Set-Content -Path "$addinFolder\launch-dict-server.vbs" -Value $vbs -Encoding ASCII
    Write-Host "    OK: launch script updated" -ForegroundColor Green

    Write-Host ""
    Write-Host "Done. Open Word and check the log." -ForegroundColor Yellow
} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
} finally {
    Read-Host "Press Enter to close"
}
