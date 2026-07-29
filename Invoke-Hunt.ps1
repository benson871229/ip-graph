#requires -Version 5.1
<#
.SYNOPSIS
    排程化的 threat hunting:把 threat-hunting-kql.md 裡適合自動化的查詢
    一次跑完,輸出摘要與 CSV 報表。

.DESCRIPTION
    純 PowerShell + .NET,Windows 5.1 內建即可執行,零依賴,沒有 CORS 問題。
    Kibana 在 Basic 授權下沒有 Alerting/Watcher,無法在 Kibana 內排程,
    所以用這支腳本 + Windows 工作排程器達成同樣效果。

    只收錄「能寫成筆數門檻」的獵捕。需要看分布或人工判斷像不像的
    (掃描的 unique count、JA3 稀有度、DNS 隧道的域名長相)刻意不放進來,
    硬要自動化只會製造大量誤報。

    每條查詢會回報:命中筆數、涉及的來源 IP 數、前幾大來源。
    超過門檻的標成 ALERT,方便排程後只看有沒有 ALERT。

.PARAMETER Server / Username / Password / Credential / ApiKey / Mode / SkipCertCheck
    連線與認證,與 get-so-graph.ps1 相同。省略 -Password 會隱藏提示輸入。

.PARAMETER Index / Since
    索引樣式(SO 2.4 常見 logs-*;2.3 為 *:so-*)與時間下界。

.PARAMETER Category
    只跑某一類:告警 / 橫向移動 / 惡意連線 / 醫療 / 外流 / 偵察。省略則全跑。

.PARAMETER ListOnly
    只列出內建的獵捕清單,不連線。

.PARAMETER OutFile
    輸出 CSV 報表。

.EXAMPLE
    .\Invoke-Hunt.ps1 -ListOnly

.EXAMPLE
    .\Invoke-Hunt.ps1 -Server https://10.x.x.x:9200 -Username analyst `
        -Index "logs-*" -Since now-24h -OutFile hunt.csv -SkipCertCheck

.EXAMPLE
    # 只跑醫療專屬的獵捕
    .\Invoke-Hunt.ps1 -Server https://10.x.x.x:9200 -Username analyst -Category 醫療 -SkipCertCheck

.NOTES
    存取限制與 get-so-graph.ps1 相同:SO 2.4 的 SSO 會擋掉 443 上的 Basic auth,
    需要開放 9200 直連:sudo so-firewall includehost elasticsearch_rest <分析機IP>
    本檔需存成 UTF-8 with BOM,否則 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [string]$Server,
    [string]$Username,
    [string]$Password,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$ApiKey,
    [ValidateSet('es', 'kibana')]
    [string]$Mode = 'es',
    [string]$Index = 'logs-*',
    [string]$Since = 'now-24h',
    [string]$Category,
    [switch]$ListOnly,
    [string]$OutFile,
    [switch]$SkipCertCheck
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# --------------------------------------------------------------------------- #
#  獵捕清單:名稱 / 分類 / KQL / 門檻(命中數超過就標 ALERT)/ 說明
#  只放「筆數門檻就能判斷」的。要看分布的請人工在 Kibana 跑。
# --------------------------------------------------------------------------- #
$Hunts = @(
    @{ Name='高嚴重度告警';       Cat='告警';     Threshold=1;
       Kql='event.dataset:alert and event.severity:1';
       Desc='Suricata 嚴重度 1。每天第一個要看的。' }

    @{ Name='C2/木馬類告警';      Cat='告警';     Threshold=1;
       Kql='event.dataset:alert and rule.name:(*CNC* or *Trojan* or *Malware* or *Backdoor*)';
       Desc='ET 規則命中 C2 或惡意程式。' }

    @{ Name='已知後門用埠';       Cat='惡意連線'; Threshold=1;
       Kql='event.dataset:conn and destination.port:(4444 or 1337 or 31337 or 12345 or 5555 or 6666)';
       Desc='Metasploit 4444、BackOrifice 31337 等。誤報極低。' }

    @{ Name='非標準埠的 TLS';     Cat='惡意連線'; Threshold=20;
       Kql='event.dataset:ssl and not destination.port:(443 or 8443 or 993 or 995 or 465 or 636)';
       Desc='TLS 走在奇怪的埠上,C2 常見特徵。' }

    @{ Name='對外自簽憑證';       Cat='惡意連線'; Threshold=10;
       Kql='event.dataset:ssl and ssl.validation_status:*self*signed* and not destination.ip:10.0.0.0/8';
       Desc='已排除內部設備,只看對外的自簽憑證。' }

    @{ Name='WinRM 遠端執行';     Cat='橫向移動'; Threshold=1;
       Kql='event.dataset:conn and destination.port:(5985 or 5986)';
       Desc='醫療網路通常極少見,一出現就值得查。' }

    @{ Name='內網 SMB 橫向';      Cat='橫向移動'; Threshold=200;
       Kql='event.dataset:conn and destination.port:(445 or 139) and source.ip:10.0.0.0/8 and destination.ip:10.0.0.0/8';
       Desc='勒索軟體與 PsExec 主要途徑。門檻需依你環境基線調整。' }

    @{ Name='內網 RDP';           Cat='橫向移動'; Threshold=50;
       Kql='event.dataset:conn and destination.port:3389 and destination.ip:10.0.0.0/8';
       Desc='IT 維運會用,門檻請依基線調整。' }

    @{ Name='失敗連線暴增';       Cat='偵察';     Threshold=500;
       Kql='event.dataset:conn and connection.state:(S0 or REJ or RSTO)';
       Desc='掃描最可靠的訊號。門檻需依環境調整。' }

    @{ Name='大量 NXDOMAIN';      Cat='偵察';     Threshold=300;
       Kql='event.dataset:dns and dns.response_code:NXDOMAIN';
       Desc='DNS 隧道與 DGA 惡意程式的特徵。' }

    @{ Name='大量出站流量';       Cat='外流';     Threshold=1;
       Kql='event.dataset:conn and source.ip:10.0.0.0/8 and not destination.ip:10.0.0.0/8 and source.bytes > 104857600';
       Desc='單一連線送出超過 100 MB。' }

    @{ Name='罕見 TLD 查詢';      Cat='外流';     Threshold=20;
       Kql='event.dataset:dns and dns.question.name:(*.top or *.xyz or *.tk or *.ml or *.cf or *.gq or *.buzz)';
       Desc='惡意基礎設施常用的便宜 TLD。' }

    @{ Name='明文 DICOM';         Cat='醫療';     Threshold=1;
       Kql='event.dataset:conn and destination.port:(104 or 11112)';
       Desc='未加密的病患影像傳輸。合規盤點用,不是攻擊偵測。' }

    @{ Name='明文 HL7';           Cat='醫療';     Threshold=1;
       Kql='event.dataset:conn and destination.port:(2575 or 6661)';
       Desc='未加密的病歷資料傳輸。只看有無,不解析內容。' }

    @{ Name='醫療設備上的 IT 協定'; Cat='醫療';   Threshold=1;
       Kql='event.dataset:conn and destination.ip:(10.20.0.11 or 10.20.0.12) and destination.port:(22 or 23 or 445 or 3389 or 5985)';
       Desc='請把 IP 換成你的 modality。醫療設備行為固定,誤報很低。' }
)

if ($ListOnly) {
    Write-Host ("內建 {0} 條獵捕:" -f $Hunts.Count) -ForegroundColor Green
    foreach ($h in $Hunts) {
        Write-Host ""
        Write-Host ("  [{0}] {1}  (門檻 {2})" -f $h.Cat, $h.Name, $h.Threshold) -ForegroundColor Cyan
        Write-Host ("    {0}" -f $h.Desc)
        Write-Host ("    KQL: {0}" -f $h.Kql) -ForegroundColor DarkGray
    }
    exit 0
}

if (-not $Server) { Write-Error "請提供 -Server(或用 -ListOnly 只看清單)。"; exit 1 }

try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}
# 不能用 scriptblock 回呼 ({$true}) — 5.1 背景執行緒呼叫會失敗 (SSL/TLS 信任關係錯誤),用 C# 委派
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
#  認證
# --------------------------------------------------------------------------- #
$authHeader = $null
if ($ApiKey) { $authHeader = "ApiKey $ApiKey" }
elseif ($Credential) {
    $u = $Credential.UserName; $p = $Credential.GetNetworkCredential().Password
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
else { Write-Error "請提供 -Username(可搭配 -Password)、-Credential 或 -ApiKey 其一。"; exit 1 }

$base = $Server.TrimEnd('/')
$headers = @{ Authorization = $authHeader }
if ($Mode -eq 'kibana') {
    $headers['kbn-xsrf'] = 'true'
    $url = "$base/api/console/proxy?path=" + [Uri]::EscapeDataString("$Index/_search") + "&method=POST"
} else {
    $url = "$base/$Index/_search"
}

function Invoke-ES($body) {
    $resp = $null
    try {
        $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType 'application/json'
    } catch {
        $msg = $_.Exception.Message
        $r = $_.Exception.Response
        if ($r) {
            try {
                $sr = New-Object System.IO.StreamReader($r.GetResponseStream())
                $txt = $sr.ReadToEnd()
                if ($txt) { $msg += " | " + $txt.Substring(0, [Math]::Min(300, $txt.Length)) }
            } catch {}
        }
        throw $msg
    }
    if ($resp -is [string] -and $resp -match '<html|<!DOCTYPE|<form') {
        throw ("伺服器回的是 HTML 登入頁 — SO 的 SSO 擋住了 Basic auth。" +
               "`n請改用 ES 直連 9200:sudo so-firewall includehost elasticsearch_rest <你的分析機IP>")
    }
    return $resp
}

# --------------------------------------------------------------------------- #
#  執行
# --------------------------------------------------------------------------- #
$list = if ($Category) { $Hunts | Where-Object { $_.Cat -eq $Category } } else { $Hunts }
if (-not $list -or @($list).Count -eq 0) {
    Write-Error ("找不到分類『{0}』。可用:{1}" -f $Category, (($Hunts | ForEach-Object { $_.Cat } | Sort-Object -Unique) -join '、'))
    exit 1
}

Write-Host ("獵捕開始 — {0} 條 · 索引 {1} · 時間 >= {2}" -f @($list).Count, $Index, $Since) -ForegroundColor Green
Write-Host ""

$results = New-Object System.Collections.Generic.List[object]
$alerts = 0
foreach ($h in $list) {
    $q = @{
        size  = 0
        query = @{ bool = @{ filter = @(
            @{ range = @{ '@timestamp' = @{ gte = $Since } } },
            @{ query_string = @{ query = $h.Kql; analyze_wildcard = $true } }
        ) } }
        aggs  = @{
            srcs = @{ terms = @{ field = 'source.ip'; size = 5 } }
            nsrc = @{ cardinality = @{ field = 'source.ip' } }
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $hits = -1; $nsrc = 0; $top = ''
    try {
        $r = Invoke-ES $q
        if ($r.hits -and $null -ne $r.hits.total) {
            $hits = if ($r.hits.total.PSObject.Properties.Name -contains 'value') { [int]$r.hits.total.value } else { [int]$r.hits.total }
        }
        if ($r.aggregations) {
            if ($r.aggregations.nsrc) { $nsrc = [int]$r.aggregations.nsrc.value }
            if ($r.aggregations.srcs -and $r.aggregations.srcs.buckets) {
                $top = (@($r.aggregations.srcs.buckets | ForEach-Object { "$($_.key)($($_.doc_count))" }) -join ' ')
            }
        }
    } catch {
        Write-Host ("  [錯誤] {0}: {1}" -f $h.Name, $_.Exception.Message) -ForegroundColor Red
        $results.Add([pscustomobject]@{ 分類=$h.Cat; 名稱=$h.Name; 命中=''; 來源IP數=''; 狀態='ERROR'; 前幾大來源=$_.Exception.Message; KQL=$h.Kql })
        continue
    }

    $state = if ($hits -ge $h.Threshold) { 'ALERT' } elseif ($hits -gt 0) { 'info' } else { 'ok' }
    if ($state -eq 'ALERT') { $alerts++ }
    $color = switch ($state) { 'ALERT' { 'Red' } 'info' { 'Yellow' } default { 'DarkGray' } }
    Write-Host ("  {0,-6} {1,-22} 命中 {2,-8} 來源 {3,-5} {4}" -f $state, $h.Name, $hits, $nsrc, $top) -ForegroundColor $color

    $results.Add([pscustomobject]@{
        分類=$h.Cat; 名稱=$h.Name; 命中=$hits; 來源IP數=$nsrc; 狀態=$state; 前幾大來源=$top; KQL=$h.Kql })
}

Write-Host ""
if ($alerts -gt 0) { Write-Host ("⚠ 有 {0} 條超過門檻,請優先查看 ALERT 項目。" -f $alerts) -ForegroundColor Red }
else { Write-Host "沒有任何獵捕超過門檻。" -ForegroundColor Green }

if ($OutFile) {
    $results | Export-Csv -LiteralPath $OutFile -NoTypeInformation -Encoding UTF8
    Write-Host ("報表已輸出:{0}" -f $OutFile) -ForegroundColor Green
}
