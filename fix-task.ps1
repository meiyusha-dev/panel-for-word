# ============================================================
# Word Panel - スケジュールタスク修復スクリプト
# WMIトリガー（Word起動時）を正しく登録します
# 管理者権限で実行してください
# ============================================================

$taskName    = "WordPanel-DictServer"
$addinFolder = "C:\OfficeAddins"

$taskXml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Word Panel 辞書サーバー (Word起動時のみ)</Description></RegistrationInfo>
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
    Write-Host "スケジュールタスクを修復しています..." -ForegroundColor Cyan

    # 既存タスクを完全削除してから再登録
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

    # Register-ScheduledTask -Xml は WMIEventTrigger を解釈できないため schtasks.exe を使う
    $tmpXml = "$env:TEMP\wordpanel-task.xml"
    $taskXml | Set-Content -Path $tmpXml -Encoding Unicode
    schtasks /Create /TN $taskName /XML $tmpXml /F 2>&1 | Out-Null
    Remove-Item $tmpXml -ErrorAction SilentlyContinue

    $trigger = (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue).Triggers.CimClass.CimClassName
    if ($trigger -eq 'MSFT_TaskEventTrigger') {
        Write-Host "    OK: WMIトリガーでタスクを登録しました" -ForegroundColor Green
    } else {
        Write-Host "    NG: トリガーの登録に失敗しました ($trigger)" -ForegroundColor Red
    }

    # VBSも最新版に上書き（古いファイルが残っている場合の対策）
    $vbs = 'CreateObject("WScript.Shell").Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\OfficeAddins\dict-server.ps1""", 0, False'
    Set-Content -Path "$addinFolder\launch-dict-server.vbs" -Value $vbs -Encoding ASCII
    Write-Host "    OK: 起動スクリプトを更新しました" -ForegroundColor Green

    Write-Host ""
    Write-Host "修復完了。Wordを起動してログを確認してください。" -ForegroundColor Yellow
} catch {
    Write-Host ""
    Write-Host "エラーが発生しました: $_" -ForegroundColor Red
} finally {
    Read-Host "Enterキーを押して閉じる"
}
