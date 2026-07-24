#requires -Version 5.1
<#
.SYNOPSIS
    零安裝的 CORS 反向代理:把瀏覽器的請求加上 CORS 標頭後轉發到 Elasticsearch / Kibana。
    給「不能改 ES 設定」又想用 ip-graph.html 的 API 面板時使用。

.DESCRIPTION
    純 PowerShell + .NET 內建 HttpListener,Windows 5.1 直接跑,零依賴。
    - 幫每個回應補上 Access-Control-Allow-* 標頭,解掉瀏覽器 CORS 阻擋
    - 正確處理預檢 (OPTIONS)
    - 轉發 method / path / query / body 與 Authorization、Content-Type、kbn-xsrf 標頭
    - 後端錯誤(401/403/400…)會原樣帶回,瀏覽器才看得到真正的 ES 錯誤而不是 CORS 假象
    - -InjectAuth 可讓代理在伺服器端補認證,帳密就不必存在瀏覽器分頁

    用法:在 ip-graph.html 的「端點 URL」填代理位址 (例如 http://localhost:8080),
    模式與索引照舊。代理會把 /<index>/_search 轉到 -Backend。

.PARAMETER Backend
    真正要打的後端。ES 直連填 https://<host>:9200;Kibana proxy 模式填 https://<host>:5601。

.PARAMETER Listen
    本機監聽位址,預設 http://localhost:8080/。

.PARAMETER InjectAuth
    (選填) 由代理注入的 Authorization 值,例如 "ApiKey xx==" 或 "Basic <base64>"。
    設了這個,ip-graph 的「認證」欄就可以留空,帳密只留在這台跑代理的機器上。

.PARAMETER SkipCertCheck
    後端是自簽憑證時加上(SO 預設常是自簽)。只在你信任該後端時使用。

.EXAMPLE
    .\cors-proxy.ps1 -Backend https://securityonion:9200 -SkipCertCheck
    .\cors-proxy.ps1 -Backend https://so:9200 -InjectAuth "ApiKey AbCd==" -SkipCertCheck
    .\cors-proxy.ps1 -Backend https://kibana:5601 -Listen http://localhost:8090/ -SkipCertCheck

.NOTES
    若啟動時「存取被拒」:以系統管理員身分開 PowerShell,或先執行一次(換成你的 port):
        netsh http add urlacl url=http://localhost:8080/ user=Everyone
    本檔需存成 UTF-8 with BOM,否則 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Backend,
    [string]$Listen = "http://localhost:8080/",
    [string]$AllowOrigin = "*",
    [string]$InjectAuth,
    [switch]$SkipCertCheck
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
if (-not $Listen.EndsWith("/")) { $Listen += "/" }
$Backend = $Backend.TrimEnd("/")

# TLS 版本 + 自簽憑證處理(5.1 預設不一定含 TLS 1.2)
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}
if ($SkipCertCheck) {
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Listen)
try {
    $listener.Start()
} catch {
    Write-Error "無法在 $Listen 監聽:$($_.Exception.Message)"
    Write-Host "若是『存取被拒』:以系統管理員身分執行 PowerShell,或先執行一次:" -ForegroundColor Yellow
    Write-Host "  netsh http add urlacl url=$Listen user=Everyone" -ForegroundColor Yellow
    exit 1
}

Write-Host "CORS 反向代理已啟動" -ForegroundColor Green
Write-Host ("  監聽 : {0}" -f $Listen)
Write-Host ("  轉發 : {0}" -f $Backend)
if ($InjectAuth) { Write-Host "  認證 : 由代理注入(ip-graph 認證欄可留空)" -ForegroundColor Cyan }
if ($SkipCertCheck) { Write-Host "  憑證 : 略過驗證(僅限信任的後端)" -ForegroundColor DarkYellow }
Write-Host ("  在 ip-graph.html 的『端點 URL』填: {0}" -f $Listen.TrimEnd("/"))
Write-Host "按 Ctrl+C 停止。"
Write-Host ""

function Add-CorsHeaders($resp, $wantHeaders) {
    $resp.Headers["Access-Control-Allow-Origin"] = $AllowOrigin
    $resp.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    $allow = "Authorization,Content-Type,kbn-xsrf"
    if ($wantHeaders) { $allow = $wantHeaders }
    $resp.Headers["Access-Control-Allow-Headers"] = $allow
    $resp.Headers["Access-Control-Max-Age"] = "600"
}

while ($listener.IsListening) {
    $context = $null
    try { $context = $listener.GetContext() } catch { break }
    try {
        $req = $context.Request
        $resp = $context.Response

        # 預檢請求:只回 CORS 標頭,不轉發
        if ($req.HttpMethod -eq "OPTIONS") {
            Add-CorsHeaders $resp $req.Headers["Access-Control-Request-Headers"]
            $resp.StatusCode = 204
            $resp.Close()
            continue
        }

        $target = $Backend + $req.RawUrl
        $fwd = [System.Net.HttpWebRequest]::Create($target)
        $fwd.Method = $req.HttpMethod
        $fwd.AllowAutoRedirect = $false
        $fwd.AutomaticDecompression =
            [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

        # 認證:優先用 -InjectAuth,否則沿用瀏覽器帶來的
        if ($InjectAuth) { $fwd.Headers["Authorization"] = $InjectAuth }
        elseif ($req.Headers["Authorization"]) { $fwd.Headers["Authorization"] = $req.Headers["Authorization"] }
        if ($req.Headers["kbn-xsrf"]) { $fwd.Headers["kbn-xsrf"] = $req.Headers["kbn-xsrf"] }
        if ($req.ContentType) { $fwd.ContentType = $req.ContentType }

        # 轉發 body
        if ($req.HasEntityBody) {
            $ms = New-Object System.IO.MemoryStream
            $req.InputStream.CopyTo($ms)
            $bytes = $ms.ToArray()
            $fwd.ContentLength = $bytes.Length
            $os = $fwd.GetRequestStream()
            $os.Write($bytes, 0, $bytes.Length)
            $os.Close()
        }

        # 取後端回應(WebException 時把錯誤回應也帶回,才看得到真正錯誤)
        $backendResp = $null
        try {
            $backendResp = $fwd.GetResponse()
        } catch [System.Net.WebException] {
            if ($_.Exception.Response) { $backendResp = $_.Exception.Response }
            else { throw }
        }

        Add-CorsHeaders $resp $null
        $resp.StatusCode = [int]$backendResp.StatusCode
        if ($backendResp.ContentType) { $resp.ContentType = $backendResp.ContentType }

        $rs = $backendResp.GetResponseStream()
        $buf = New-Object byte[] 65536
        while (($n = $rs.Read($buf, 0, $buf.Length)) -gt 0) {
            $resp.OutputStream.Write($buf, 0, $n)
        }
        $rs.Close()
        $backendResp.Close()
        $resp.OutputStream.Close()
        $resp.Close()

        Write-Host ("{0}  {1} {2} -> {3}" -f (Get-Date -Format "HH:mm:ss"), $req.HttpMethod, $req.RawUrl, $resp.StatusCode)
    } catch {
        Write-Host ("錯誤: " + $_.Exception.Message) -ForegroundColor Red
        try {
            $resp = $context.Response
            Add-CorsHeaders $resp $null
            $resp.StatusCode = 502
            $resp.ContentType = "application/json"
            $emsg = '{"error":"proxy: ' + ($_.Exception.Message -replace '["\\]', ' ') + '"}'
            $eb = [System.Text.Encoding]::UTF8.GetBytes($emsg)
            $resp.OutputStream.Write($eb, 0, $eb.Length)
            $resp.Close()
        } catch {}
    }
}
$listener.Stop()
