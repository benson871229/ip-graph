#requires -Version 5.1
<#
.SYNOPSIS
    用帳密(或 API key)向 Security Onion 的 Elasticsearch 查 Suricata 告警,
    依 IP 彙總嚴重度與告警名稱,輸出成 ip-graph.html「告警疊圖」欄位可貼的清單。

.DESCRIPTION
    純 PowerShell + .NET,Windows 5.1 內建即可執行,零依賴,沒有 CORS 問題。
    專抓告警:預設過濾 event.dataset:alert(SO 的 Suricata 告警),不含一般 conn 流量。

    流程:認證 → 對 source.ip 與 destination.ip 各做一次 terms 聚合
    (子聚合:最嚴重度 min(event.severity) + 前幾大告警名稱 rule.name)→ 合併 → 輸出。

    輸出格式(工具會自動彙總同一 IP 的多行):
        ip,嚴重度,告警名稱          例如  10.20.0.199,1,ET MALWARE Cobalt Strike Beacon
    嚴重度沿用 Suricata:1 最嚴重、3 最輕。貼進工具後勾「告警風險模式」即依嚴重度上色。

.PARAMETER Server / Username / Password / Credential / ApiKey / Mode / SkipCertCheck
    連線與認證,與 get-so-graph.ps1 相同。省略 -Password 會隱藏提示輸入。

.PARAMETER Index
    索引 / index pattern。SO 常見為 *:so-*(跨叢集)。

.PARAMETER AlertFilter
    只抓告警的過濾條件(KQL/Lucene)。預設 event.dataset:alert。
    不同 SO 版本可能要改成 event.module:suricata 或 tags:alert。

.PARAMETER SevField / SigField / IpFields
    嚴重度欄位(需為數值,預設 event.severity)、告警名稱欄位(預設 rule.name)、
    要歸戶的 IP 欄位(預設 source.ip 與 destination.ip 都算)。

.PARAMETER Since / Query / Size / TopSigs / OutFile
    時間下界;額外過濾;每個 IP 欄位取前幾個 IP;每個 IP 取前幾個告警名稱;輸出檔。

.EXAMPLE
    .\get-so-alerts.ps1 -Server https://securityonion:9200 -Username analyst `
        -Since now-24h -OutFile alerts.csv -SkipCertCheck

.EXAMPLE
    # 舊版 SO 用 event.module 過濾、嚴重度欄位不同時
    .\get-so-alerts.ps1 -Server https://so:9200 -ApiKey "AbCd==" `
        -AlertFilter 'event.module:suricata' -SevField rule.severity -OutFile alerts.csv -SkipCertCheck

.NOTES
    產出的 alerts.csv 內容整份貼進 ip-graph.html 的「告警疊圖 · Suricata」欄位。
    本檔需存成 UTF-8 with BOM,否則 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Server,
    [string]$Username,
    [string]$Password,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$ApiKey,
    [ValidateSet('es', 'kibana')]
    [string]$Mode = 'es',
    [string]$Index = '*:so-*',
    [string]$AlertFilter = 'event.dataset:alert',
    [string]$SevField = 'event.severity',
    [string]$SigField = 'rule.name',
    [string[]]$IpFields = @('source.ip', 'destination.ip'),
    [string]$Since = 'now-24h',
    [string]$Query,
    [int]$Size = 500,
    [int]$TopSigs = 3,
    [string]$OutFile,
    [switch]$SkipCertCheck
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}
if ($SkipCertCheck) { [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }

# --------------------------------------------------------------------------- #
#  認證
# --------------------------------------------------------------------------- #
$authHeader = $null
if ($ApiKey) {
    $authHeader = "ApiKey $ApiKey"
}
elseif ($Credential) {
    $u = $Credential.UserName
    $p = $Credential.GetNetworkCredential().Password
    $authHeader = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${u}:${p}"))
}
elseif ($Username) {
    if (-not $Password) {
        $sec = Read-Host "密碼 ($Username)" -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try { $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    $authHeader = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${Username}:${Password}"))
}
else {
    Write-Error "請提供 -Username(可搭配 -Password)、-Credential 或 -ApiKey 其一。"
    exit 1
}

$base = $Server.TrimEnd('/')
$headers = @{ Authorization = $authHeader }
if ($Mode -eq 'kibana') {
    $headers['kbn-xsrf'] = 'true'
    $url = "$base/api/console/proxy?path=" + [Uri]::EscapeDataString("$Index/_search") + "&method=POST"
}
else {
    $url = "$base/$Index/_search"
}

function Invoke-ES($body) {
    try {
        return Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType 'application/json'
    }
    catch {
        $msg = $_.Exception.Message
        $r = $_.Exception.Response
        if ($r) {
            try {
                $sr = New-Object System.IO.StreamReader($r.GetResponseStream())
                $txt = $sr.ReadToEnd()
                if ($txt) { $msg += " | " + $txt.Substring(0, [Math]::Min(400, $txt.Length)) }
            } catch {}
        }
        throw $msg
    }
}

# --------------------------------------------------------------------------- #
#  查詢:對每個 IP 欄位做一次 terms 聚合,合併結果
#  agg = { ip -> @{ sev; sigs = HashSet } }  (sev 取最小=最嚴重)
# --------------------------------------------------------------------------- #
$agg = @{}
function Merge-Bucket([string]$ip, $sev, $sigNames) {
    if (-not $ip) { return }
    if (-not $agg.ContainsKey($ip)) {
        $agg[$ip] = @{ sev = 9; sigs = New-Object 'System.Collections.Generic.List[string]' }
    }
    $o = $agg[$ip]
    if ($null -ne $sev -and $sev -lt $o.sev) { $o.sev = [int]$sev }
    foreach ($s in $sigNames) {
        if ($s -and ($o.sigs -notcontains $s)) { $o.sigs.Add([string]$s) }
    }
}

try {
    foreach ($ipField in $IpFields) {
        $filter = @()
        if ($Since) { $filter += @{ range = @{ '@timestamp' = @{ gte = $Since } } } }
        $filter += @{ query_string = @{ query = $AlertFilter } }
        if ($Query) { $filter += @{ query_string = @{ query = $Query } } }
        $filter += @{ exists = @{ field = $ipField } }

        $ipAgg = @{ terms = @{ field = $ipField; size = $Size } }
        $ipAgg.aggs = @{ sigs = @{ terms = @{ field = $SigField; size = $TopSigs } } }
        if ($SevField) { $ipAgg.aggs.sev = @{ min = @{ field = $SevField } } }

        $body = @{ size = 0; query = @{ bool = @{ filter = $filter } }; aggs = @{ byip = $ipAgg } } |
            ConvertTo-Json -Depth 20 -Compress
        $resp = Invoke-ES $body
        $buckets = $null
        if ($resp.aggregations -and $resp.aggregations.byip) { $buckets = $resp.aggregations.byip.buckets }
        if ($null -eq $buckets) { throw "回應沒有 aggregations.byip — 檢查索引/欄位/權限或 -AlertFilter。" }

        foreach ($b in $buckets) {
            $sev = $null
            if ($SevField -and $b.sev -and $null -ne $b.sev.value) { $sev = [int][math]::Round($b.sev.value) }
            $sigNames = @()
            if ($b.sigs -and $b.sigs.buckets) { $sigNames = @($b.sigs.buckets | ForEach-Object { $_.key }) }
            Merge-Bucket ([string]$b.key) $sev $sigNames
        }
        Write-Host ("{0}:{1} 個 IP" -f $ipField, $buckets.Count)
    }
}
catch {
    Write-Error "查詢失敗:$($_.Exception.Message)"
    exit 2
}

# --------------------------------------------------------------------------- #
#  輸出:每個 (ip, 告警名稱) 一行;嚴重度夾在 1..3
# --------------------------------------------------------------------------- #
$rows = New-Object System.Collections.Generic.List[string]
foreach ($ip in $agg.Keys) {
    $o = $agg[$ip]
    $sev = if ($o.sev -ge 1 -and $o.sev -le 3) { $o.sev } elseif ($o.sev -lt 1) { 1 } else { 3 }
    if ($o.sigs.Count -gt 0) {
        foreach ($s in $o.sigs) {
            $clean = ($s -replace '[,\r\n]', ' ').Trim()
            $rows.Add("$ip,$sev,$clean")
        }
    }
    else {
        $rows.Add("$ip,$sev")
    }
}

if ($rows.Count -eq 0) {
    Write-Warning "沒有抓到任何告警。"
    Write-Host "可能:時間範圍內無告警,或 -AlertFilter / -SevField / -SigField 欄名與你的 SO 版本不符。" -ForegroundColor Cyan
    Write-Host ("目前設定 → 索引:{0}  過濾:{1}  嚴重度:{2}  名稱:{3}  時間:>= {4}" -f $Index, $AlertFilter, $SevField, $SigField, $Since)
    exit 2
}

# 依嚴重度(1 先)再依 IP 排序,方便閱讀
$sorted = $rows | Sort-Object @{ Expression = { [int]($_ -split ',')[1] } }, @{ Expression = { ($_ -split ',')[0] } }

if ($OutFile) {
    $sorted | Set-Content -LiteralPath $OutFile -Encoding UTF8
    Write-Host ("已輸出 {0} 行({1} 個 IP)到 {2}" -f $rows.Count, $agg.Count, $OutFile) -ForegroundColor Green
    Write-Host "把檔案內容整份貼進 ip-graph.html 的「告警疊圖 · Suricata」欄位,再勾「告警風險模式」。"
}
else {
    $sorted
    Write-Host ""
    Write-Host ("共 {0} 行、{1} 個 IP。加 -OutFile alerts.csv 可存檔。" -f $rows.Count, $agg.Count) -ForegroundColor Green
}
