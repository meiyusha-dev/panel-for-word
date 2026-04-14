# ============================================================
# Word Panel 繧｢繝峨う繝ｳ 繧､繝ｳ繧ｹ繝医・繝ｩ繝ｼ
# 邂｡逅・・ｨｩ髯舌〒螳溯｡後＠縺ｦ縺上□縺輔＞・・nstall.bat 縺九ｉ襍ｷ蜍輔＠縺ｦ縺上□縺輔＞・・# ============================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$manifest_url    = "https://raw.githubusercontent.com/otayoshino/panel-for-word/master/manifest.xml"
$dictServer_url  = "https://raw.githubusercontent.com/otayoshino/panel-for-word/master/dict-server.ps1"
$dictBase_url    = "https://otayoshino.github.io/panel-for-word/dict"
$addinFolder     = "C:\OfficeAddins"
$dictFolder      = "$addinFolder\dict"
$shareName       = "OfficeAddins"
$taskName        = "WordPanel-DictServer"

$dictFiles = @(
    'base.dat.gz', 'cc.dat.gz', 'check.dat.gz',
    'tid.dat.gz', 'tid_map.dat.gz', 'tid_pos.dat.gz',
    'unk.dat.gz', 'unk_char.dat.gz', 'unk_compat.dat.gz',
    'unk_invoke.dat.gz', 'unk_map.dat.gz', 'unk_pos.dat.gz'
)

function Write-Step($msg) { Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "    WARN: $msg" -ForegroundColor Yellow }
function Write-Fail($msg) { Write-Host "    ERROR: $msg" -ForegroundColor Red }

Write-Host "======================================" -ForegroundColor Yellow
Write-Host "  Word Panel 繧｢繝峨う繝ｳ 繧､繝ｳ繧ｹ繝医・繝ｩ繝ｼ  " -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow

# 笏笏 蠢・医せ繝・ャ繝暦ｼ亥､ｱ謨励＠縺溘ｉ荳ｭ譁ｭ・・笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏
$coreOk = $false
try {
    # 1. 繝輔か繝ｫ繝菴懈・
    Write-Step "繝輔か繝ｫ繝繧剃ｽ懈・縺励※縺・∪縺・.."
    New-Item -ItemType Directory -Path $addinFolder -Force | Out-Null
    New-Item -ItemType Directory -Path $dictFolder  -Force | Out-Null
    Write-OK "$addinFolder 繧剃ｽ懈・縺励∪縺励◆"

    # 2. manifest.xml 繧偵ム繧ｦ繝ｳ繝ｭ繝ｼ繝・    Write-Step "manifest.xml 繧偵ム繧ｦ繝ｳ繝ｭ繝ｼ繝峨＠縺ｦ縺・∪縺・.."
    Invoke-WebRequest -Uri $manifest_url -OutFile "$addinFolder\manifest.xml" -UseBasicParsing
    Write-OK "manifest.xml 繧剃ｿ晏ｭ倥＠縺ｾ縺励◆"

    # 3. 繝輔か繝ｫ繝繧貞・譛・    Write-Step "繝輔か繝ｫ繝繧貞・譛峨＠縺ｦ縺・∪縺・.."
    if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
        Remove-SmbShare -Name $shareName -Force | Out-Null
    }
    New-SmbShare -Name $shareName -Path $addinFolder -FullAccess "Everyone" | Out-Null
    $uncPath = "\\$env:COMPUTERNAME\$shareName"
    Write-OK "蜈ｱ譛峨ヱ繧ｹ: $uncPath"

    # 4. Office 2019 蜷代￠ WebView2 蟇ｾ蠢懊Ξ繧ｸ繧ｹ繝医Μ
    Write-Step "WebView2 蟇ｾ蠢懊Ξ繧ｸ繧ｹ繝医Μ繧定ｨｭ螳壹＠縺ｦ縺・∪縺・.."
    reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF" /v "Win32WebView2" /t REG_DWORD /d 1 /f | Out-Null
    Write-OK "WebView2 繝ｬ繧ｸ繧ｹ繝医Μ繧定ｨｭ螳壹＠縺ｾ縺励◆"

    # 5. Word 縺ｮ菫｡鬆ｼ縺ｧ縺阪ｋ繧ｫ繧ｿ繝ｭ繧ｰ縺ｫ逋ｻ骭ｲ
    Write-Step "Word 縺ｮ繧｢繝峨う繝ｳ繧ｫ繧ｿ繝ｭ繧ｰ繧堤匳骭ｲ縺励※縺・∪縺・.."
    $catalogBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    $existing = Get-ChildItem -Path $catalogBase -ErrorAction SilentlyContinue |
        Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Url -eq $uncPath }
    if ($existing) {
        Write-OK "繧ｫ繧ｿ繝ｭ繧ｰ縺ｯ譌｢縺ｫ逋ｻ骭ｲ貂医∩縺ｧ縺・
    } else {
        $guid    = [System.Guid]::NewGuid().ToString("B").ToUpper()
        $regPath = "$catalogBase\$guid"
        New-Item -Path $regPath -Force | Out-Null
        Set-ItemProperty -Path $regPath -Name "Id"    -Value $guid
        Set-ItemProperty -Path $regPath -Name "Url"   -Value $uncPath
        Set-ItemProperty -Path $regPath -Name "Flags" -Value 1
        Write-OK "繧ｫ繧ｿ繝ｭ繧ｰ繧堤匳骭ｲ縺励∪縺励◆: $uncPath"
    }

    $coreOk = $true
} catch {
    Write-Fail "蠢・医せ繝・ャ繝励〒繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆・・
    Write-Fail $_.Exception.Message
}

# 笏笏 繧ｪ繝励す繝ｧ繝ｳ繧ｹ繝・ャ繝暦ｼ亥､ｱ謨励＠縺ｦ繧らｶ夊｡鯉ｼ・笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏
if ($coreOk) {
    # 6. 霎樊嶌繝輔ぃ繧､繝ｫ繧偵ム繧ｦ繝ｳ繝ｭ繝ｼ繝会ｼ育ｴ・7MB・・    Write-Step "霎樊嶌繝輔ぃ繧､繝ｫ繧偵ム繧ｦ繝ｳ繝ｭ繝ｼ繝峨＠縺ｦ縺・∪縺呻ｼ育ｴ・7MB・・.."
    $dictOk = $true
    try {
        $i = 1
        foreach ($f in $dictFiles) {
            Write-Host "    ($i/$($dictFiles.Count)) $f ..." -NoNewline
            Invoke-WebRequest -Uri "$dictBase_url/$f" -OutFile "$dictFolder\$f" -UseBasicParsing
            Write-Host " OK" -ForegroundColor Green
            $i++
        }
        Write-OK "霎樊嶌繝輔ぃ繧､繝ｫ繧・$dictFolder 縺ｫ菫晏ｭ倥＠縺ｾ縺励◆"
    } catch {
        $dictOk = $false
        Write-Warn "霎樊嶌繝繧ｦ繝ｳ繝ｭ繝ｼ繝峨ｒ繧ｹ繧ｭ繝・・縺励∪縺励◆・医ロ繝・ヨ繝ｯ繝ｼ繧ｯ蛻ｶ髯舌・蜿ｯ閭ｽ諤ｧ・・
        Write-Warn $_.Exception.Message
    }

    # 7. dict-server.ps1 繧帝・鄂ｮ縺励※繧ｹ繧ｿ繝ｼ繝医い繝・・繧ｿ繧ｹ繧ｯ縺ｫ逋ｻ骭ｲ
    if ($dictOk) {
        Write-Step "霎樊嶌繧ｵ繝ｼ繝舌・繧定ｨｭ螳壹＠縺ｦ縺・∪縺・.."
        try {
            Invoke-WebRequest -Uri $dictServer_url -OutFile "$addinFolder\dict-server.ps1" -UseBasicParsing
            $action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
                            -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$addinFolder\dict-server.ps1`""
            $trigger  = New-ScheduledTaskTrigger -AtLogOn
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
                -Settings $settings -RunLevel Highest -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
            Write-OK "霎樊嶌繧ｵ繝ｼ繝舌・繧ｿ繧ｹ繧ｯ '$taskName' 繧堤匳骭ｲ繝ｻ襍ｷ蜍輔＠縺ｾ縺励◆"
        } catch {
            Write-Warn "霎樊嶌繧ｵ繝ｼ繝舌・縺ｮ險ｭ螳壹ｒ繧ｹ繧ｭ繝・・縺励∪縺励◆"
            Write-Warn $_.Exception.Message
        }
    }
}

# 笏笏 螳御ｺ・Γ繝・そ繝ｼ繧ｸ 笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏笏
Write-Host ""
if ($coreOk) {
    Write-Host "======================================" -ForegroundColor Yellow
    Write-Host "  繧､繝ｳ繧ｹ繝医・繝ｫ螳御ｺ・ｼ・                 " -ForegroundColor Yellow
    Write-Host "======================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "谺｡縺ｮ謇矩・〒繧｢繝峨う繝ｳ繧定ｿｽ蜉縺励※縺上□縺輔＞・・ -ForegroundColor White
    Write-Host ""
    Write-Host "  1. Word 繧貞ｮ悟・縺ｫ蜀崎ｵｷ蜍輔☆繧・ -ForegroundColor White
    Write-Host "  2. 縲碁幕逋ｺ縲阪ち繝・竊・縲後い繝峨う繝ｳ縲阪ｒ繧ｯ繝ｪ繝・け" -ForegroundColor White
    Write-Host "     窶ｻ縲碁幕逋ｺ縲阪ち繝悶′縺ｪ縺・ｴ蜷茨ｼ・ -ForegroundColor White
    Write-Host "       縲後ヵ繧｡繧､繝ｫ縲坂・縲後が繝励す繝ｧ繝ｳ縲坂・縲後Μ繝懊Φ縺ｮ繝ｦ繝ｼ繧ｶ繝ｼ險ｭ螳壹坂・縲碁幕逋ｺ縲阪↓繝√ぉ繝・け" -ForegroundColor White
    Write-Host "  3. 縲悟・譛峨ヵ繧ｩ繝ｫ繝縲阪ち繝・竊・縲係ord Panel縲坂・ 縲瑚ｿｽ蜉縲・ -ForegroundColor White
    Write-Host ""
} else {
    Write-Host "======================================" -ForegroundColor Red
    Write-Host "  繧､繝ｳ繧ｹ繝医・繝ｫ縺ｫ螟ｱ謨励＠縺ｾ縺励◆          " -ForegroundColor Red
    Write-Host "======================================" -ForegroundColor Red
    Write-Host ""
}

# pause 縺ｯ蠢・★縺薙％縺ｧ螳溯｡鯉ｼ・xit 繧剃ｽｿ繧上↑縺・％縺ｨ縺ｧ遒ｺ螳溘↓蛻ｰ驕斐☆繧具ｼ・pause
if (-not $coreOk) { exit 1 }
