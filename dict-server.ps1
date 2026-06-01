# ============================================================
# Word Panel 辞書 HTTP サーバー
# localhost:8642 で C:\OfficeAddins\dict\ を配信する
# install.ps1 により Word 起動時に自動起動されます
# ============================================================

$port    = 8642
$dictDir = 'C:\OfficeAddins\dict'
$logFile = 'C:\OfficeAddins\dict-server.log'

function Write-Log($msg) {
    Add-Content $logFile "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg" -ErrorAction SilentlyContinue
}

# 既に起動中なら終了（ポート占有チェック）
$inUse = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($inUse) { Write-Log "Already running. Exit."; exit 0 }

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    $listener.Start()
    Write-Log "Started on port $port"
} catch {
    Write-Log "Failed to start: $_"
    exit 1
}

# シャットダウン時に listener を安全に停止するハンドラ
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    try { $listener.Stop() } catch {}
}

# GetContext() を非同期化し、30秒ごとに Word プロセス存在チェックを行う。
# Word が終了していればサーバーも自己停止することで、Windows シャットダウン時に
# .NET CLR が強制終了される 0xc0000142 エラーを防ぐ。
while ($listener.IsListening) {
    $async    = $listener.BeginGetContext($null, $null)
    $signaled = $async.AsyncWaitHandle.WaitOne(30000)

    if (-not $listener.IsListening) { break }

    if (-not $signaled) {
        if (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
            Write-Log "WINWORD.EXE not running. Stopping server."
            $listener.Stop()
        }
        continue
    }

    try {
        $ctx = $listener.EndGetContext($async)
        $req = $ctx.Request
        $res = $ctx.Response

        # CORS ヘッダー（GitHub Pages の HTTPS コンテキストから呼び出されるため必須）
        $res.Headers.Add('Access-Control-Allow-Origin',  '*')
        $res.Headers.Add('Access-Control-Allow-Methods', 'GET, HEAD, OPTIONS')

        if ($req.HttpMethod -eq 'OPTIONS') {
            $res.StatusCode = 204
            $res.Close()
            continue
        }

        $localPath = $req.Url.LocalPath.TrimStart('/')
        $file      = Join-Path $dictDir $localPath

        if (Test-Path $file -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($file)
            $res.ContentType      = 'application/octet-stream'
            $res.ContentLength64  = $bytes.Length
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
            $res.StatusCode = 200
        } else {
            $res.StatusCode = 404
        }
        $res.Close()
    } catch {
        try { $ctx.Response.StatusCode = 500; $ctx.Response.Close() } catch {}
    }
}
Write-Log "Server stopped."
