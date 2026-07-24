#requires -Version 5.1
<#
.SYNOPSIS
    用帳密(或 API key)直接向 Security Onion 的 Elasticsearch / Kibana 查詢,
    在伺服器端彙總 IP 關係,輸出成 ip-graph.html 可直接拖入的 CSV。

.DESCRIPTION
    純 PowerShell + .NET,Windows 5.1 內建即可執行,零依賴。
    因為是 PowerShell 送請求(不是瀏覽器),所以「沒有 CORS 問題」,
    不必開 ES CORS、也不必跑 cors-proxy.ps1。

    流程:認證 → 送 composite 聚合 (來源 × 目的, sum(bytes)) → 分頁取回 → 寫 CSV。
    輸出格式為工具的「自訂 CSV」:  來源,目的,權重
    把產出的 .csv 拖進 ip-graph.html 即可生圖(hop / 資產表 / 威脅標記照常適用)。

.PARAMETER Server
    ES 或 Kibana 位址。ES 直連填 https://<host>:9200;Kibana proxy 填 https://<host>:5601。

.PARAMETER Username / Password
    SO/ES 帳密。省略 -Password 會以隱藏方式提示輸入(較安全,不會留在命令列歷史)。

.PARAMETER Credential
    改用 PSCredential(-Credential (Get-Credential))。與 -Username/-Password 擇一。

.PARAMETER ApiKey
    改用 API key(base64 形式,即 Kibana 建立時給的 "encoded" 值)。與帳密擇一。

.PARAMETER Mode
    es(預設,直連 :9200)或 kibana(走 :5601 的 console proxy)。

.PARAMETER Index / SrcField / DstField / BytesField
    索引與欄位。預設走 ECS。Zeek 原始欄位可改成 id.orig_h / id.resp_h / orig_ip_bytes。

.PARAMETER Since / Query
    時間下界(gte,如 now-24h、2026-07-01)與額外 KQL/Lucene 過濾。

.PARAMETER MaxPairs / OutFile / SkipCertCheck
    最多取幾對連線;輸出檔(不給則印到畫面);自簽憑證略過驗證。

.EXAMPLE
    # 帳密查最近 24 小時,輸出 CSV(密碼會提示輸入)
    .\get-so-graph.ps1 -Server https://securityonion:9200 -Username analyst `
        -Index "*:so-*" -Since now-24h -OutFile graph.csv -SkipCertCheck

.EXAMPLE
    # 聚焦某台主機、Zeek 原始欄位、用 API key
    .\get-so-graph.ps1 -Server https://so:9200 -ApiKey "AbCd12==" `
        -SrcField id.orig_h -DstField id.resp_h -BytesField orig_ip_bytes `
        -Query 'id.orig_h:10.20.0.30 OR id.resp_h:10.20.0.30' -OutFile ws.csv -SkipCertCheck

.NOTES
    安全:盡量用 -Credential 或省略 -Password(提示輸入),避免把明碼打在命令列。
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
    [string]$SrcField = 'source.ip',
    [string]$DstField = 'destination.ip',
    [string]$BytesField = 'network.bytes',
    [string]$Since = 'now-24h',
    [string]$Query,
    [int]$MaxPairs = 5000,
    [string]$OutFile,
    [switch]$SkipCertCheck
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# TLS + 自簽憑證(5.1 沒有 Invoke-RestMethod -SkipCertificateCheck,得自己設)
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}
# 注意:不能用 scriptblock 當回呼 ({$true}) — 5.1 由背景執行緒呼叫時會失敗,
# 出現「基礎連接已關閉: 無法為 SSL/TLS 安全通道建立信任關係」。必須用編譯過的 C# 委派。
if ($SkipCertCheck) {
    if (-not ("TrustAllCerts" -as [type])) {
        Add-Type -TypeDefinition @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class TrustAllCerts {
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback =
            delegate (object s, X509Certificate c, X509Chain ch, SslPolicyErrors e) { return true; };
    }
}
"@
    }
    [TrustAllCerts]::Enable()
}

# --------------------------------------------------------------------------- #
#  認證標頭
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

# --------------------------------------------------------------------------- #
#  查詢主體(composite 聚合,可帶 after 分頁)
# --------------------------------------------------------------------------- #
function Build-Body($after) {
    $filter = @()
    if ($Since) { $filter += @{ range = @{ '@timestamp' = @{ gte = $Since } } } }
    if ($Query) { $filter += @{ query_string = @{ query = $Query } } }
    $comp = @{
        size    = 1000
        sources = @(
            @{ s = @{ terms = @{ field = $SrcField } } },
            @{ d = @{ terms = @{ field = $DstField } } }
        )
    }
    if ($after) { $comp.after = $after }
    $pairs = @{ composite = $comp }
    if ($BytesField) { $pairs.aggs = @{ bytes = @{ sum = @{ field = $BytesField } } } }
    return @{ size = 0; query = @{ bool = @{ filter = $filter } }; aggs = @{ pairs = $pairs } }
}

function Invoke-ES($body) {
    $resp = $null
    try {
        $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType 'application/json'
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
    # SO 2.4 的 SSO 會對未帶登入 cookie 的請求回 200 + HTML 登入頁;
    # Invoke-RestMethod 會把它當字串傳回 → 明確報出來,不要讓後續誤判成權限問題
    if ($resp -is [string] -and $resp -match '<html|<!DOCTYPE|<form') {
        throw ("伺服器回的是 HTML 登入頁,不是 ES 回應 — 你的 SO 有 SSO 網頁登入,Basic auth 進不了 /kibana。" +
               "`n解法:請管理員在 SO 主機執行  sudo so-firewall includehost elasticsearch_rest 你的分析機IP" +
               "`n開通後改用 ES 直連:  -Server https://SO主機:9200  (去掉 -Mode kibana,帳密同 SO 登入)")
    }
    return $resp
}

# --------------------------------------------------------------------------- #
#  抓取
# --------------------------------------------------------------------------- #
$rows = New-Object System.Collections.Generic.List[string]
$after = $null; $pairs = 0; $page = 0; $truncated = $false

try {
    do {
        $body = Build-Body $after | ConvertTo-Json -Depth 20 -Compress
        $resp = Invoke-ES $body
        $agg = $null
        if ($resp.aggregations) { $agg = $resp.aggregations.pairs }
        if ($null -eq $agg) {
            # 最常見原因:索引樣式沒對到任何索引 (ES 回 200 但整個 aggregations 消失)。
            # 自動列出實際存在的索引,幫使用者找到正確的 -Index。
            $hint = "回應沒有 aggregations.pairs — 最常見原因是 -Index『$Index』沒對到任何索引。"
            try {
                $catPath = '_cat/indices?format=json&h=index'
                if ($Mode -eq 'kibana') {
                    $iu = "$base/api/console/proxy?path=" + [Uri]::EscapeDataString($catPath) + "&method=GET"
                    $ilist = Invoke-RestMethod -Uri $iu -Method Post -Headers $headers
                } else {
                    $ilist = Invoke-RestMethod -Uri "$base/$catPath" -Method Get -Headers $headers
                }
                $names = @($ilist | ForEach-Object { $_.index } | Where-Object { $_ -and $_[0] -ne '.' } | Sort-Object)
                $inter = @($names | Where-Object { $_ -match 'zeek|conn|suricata|so-|logs-' } | Select-Object -First 25)
                if ($inter.Count -gt 0) {
                    $hint += "`n這台 ES 實際存在的相關索引 (取前 25 個):`n  " + ($inter -join "`n  ")
                    $hint += "`n請改用符合的樣式重跑,例如 -Index `"logs-*`" (SO 2.4) 或 -Index `"so-*`" (SO 2.3)。"
                } elseif ($names.Count -gt 0) {
                    $hint += "`n這台 ES 的索引 (取前 25 個):`n  " + (@($names | Select-Object -First 25) -join "`n  ")
                } else {
                    $hint += "`n且列不出任何索引 — 可能是帳號權限不足。"
                }
            } catch { $hint += "`n(嘗試列出索引也失敗:$($_.Exception.Message))" }
            throw $hint
        }

        foreach ($b in $agg.buckets) {
            $s = $b.key.s; $d = $b.key.d
            if (-not $s -or -not $d) { continue }
            $w = if ($b.PSObject.Properties.Name -contains 'bytes' -and $null -ne $b.bytes.value) {
                [int64]$b.bytes.value
            } else { [int64]$b.doc_count }
            $rows.Add("$s,$d,$w")
            $pairs++
            if ($pairs -ge $MaxPairs) { $truncated = $true; break }
        }

        $after = $agg.after_key
        $page++
        Write-Host ("已取 {0} 對…" -f $pairs)
    } while ($after -and -not $truncated -and $page -lt 50)
}
catch {
    Write-Error "查詢失敗:$($_.Exception.Message)"
    exit 2
}

# --------------------------------------------------------------------------- #
#  輸出
# --------------------------------------------------------------------------- #
if ($rows.Count -eq 0) {
    Write-Warning "查詢成功但沒有連線資料(0 對)。"
    Write-Host "可能原因:索引/時間範圍/欄位名稱不符,或該範圍內確實沒有符合的資料。" -ForegroundColor Cyan
    Write-Host ("目前設定 → 索引:{0}  來源:{1}  目的:{2}  流量:{3}  時間:>= {4}" -f $Index, $SrcField, $DstField, $BytesField, $Since)
    exit 2
}

if ($OutFile) {
    $rows | Set-Content -LiteralPath $OutFile -Encoding UTF8
    Write-Host ("已輸出 {0} 條連線到 {1}{2}" -f $rows.Count, $OutFile, $(if ($truncated) { "(已達上限 $MaxPairs,已截斷)" } else { "" })) -ForegroundColor Green
    Write-Host "把這個檔案拖進 ip-graph.html 即可生圖。"
}
else {
    $rows
    Write-Host ""
    Write-Host ("共 {0} 條連線{1}。加 -OutFile graph.csv 可存檔後拖進工具。" -f $rows.Count, $(if ($truncated) { "(已截斷)" } else { "" })) -ForegroundColor Green
}
