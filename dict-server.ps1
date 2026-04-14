# ============================================================
# Word Panel 辞書 HTTP サーバー
# localhost:8642 で C:\OfficeAddins\dict\ を配信する
# install.ps1 によりログオン時に自動起動されます
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

while ($listener.IsListening) {
    try {
        $ctx    = $listener.GetContext()
        $req    = $ctx.Request
        $res    = $ctx.Response

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
