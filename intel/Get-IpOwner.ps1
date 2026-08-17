#requires -Version 5.1
<#
.SYNOPSIS
    批次查出一堆 public IP 分別屬於誰(組織、國家、網段),並標出雲端/VPS 供應商。
    產出 CSV,可直接貼進 ip-graph 的資產表或用 Excel 看。

.DESCRIPTION
    純 PowerShell 5.1,不需安裝任何東西。用 RDAP(WHOIS 的現代版,RFC 9083):
      * HTTPS、免註冊、免 API key
      * 查的是各區域註冊機構(ARIN/RIPE/APNIC/LACNIC/AFRINIC)的權威資料
      * 走 rdap.org 自動轉到正確的註冊機構

    這支要在**有網路的機器**上跑。產出的 CSV 帶回離線環境即可。

    私有位址(10./172.16-31./192.168./127./169.254.)會自動跳過,不送出查詢。

.PARAMETER InFile
    IP 清單檔,每行一個。可含逗號分隔的其他欄位(只取第一個看起來像 IP 的字串),
    # 開頭的行會忽略 —— 所以可以直接餵 ip-graph 匯出的 CSV。

.PARAMETER Ip
    直接在命令列給 IP,可多個。與 -InFile 擇一或併用。

.PARAMETER OutFile
    輸出 CSV。省略則印到畫面。

.PARAMETER DelayMs
    每次查詢間隔毫秒,預設 300。RDAP 有速率限制,查很多時不要調太小。

.EXAMPLE
    .\Get-IpOwner.ps1 -InFile public-ips.txt -OutFile owners.csv

.EXAMPLE
    .\Get-IpOwner.ps1 -Ip 8.8.8.8,104.131.0.5 -OutFile owners.csv

.NOTES
    隱私提醒:查詢會把「目的地 IP」送到公開的註冊資料庫。查的是對方的登記資料,
    不會洩漏你的內網位址;但仍等於告訴外部「有人在查這個 IP」。
    這是 SOC 的標準做法,若你的規範不允許,請改用離線的 IP-to-ASN 資料集。

    本檔需存成 UTF-8 with BOM,否則 PowerShell 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [string]$InFile,
    [string[]]$Ip,
    [string]$OutFile,
    [int]$DelayMs = 300
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
} catch {}

# --------------------------------------------------------------------------- #
#  收集要查的 IP
# --------------------------------------------------------------------------- #
function Get-FirstIP([string]$s) {
    if (-not $s) { return $null }
    if ($s -match '(?<![\d.])(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?![\d.])') {
        if ([int]$Matches[1] -le 255 -and [int]$Matches[2] -le 255 -and
            [int]$Matches[3] -le 255 -and [int]$Matches[4] -le 255) { return $Matches[0] }
    }
    return $null
}
function Test-Private([string]$ip) {
    $p = $ip.Split('.')
    if ($p.Count -ne 4) { return $true }
    $a = [int]$p[0]; $b = [int]$p[1]
    if ($a -eq 10 -or $a -eq 127) { return $true }
    if ($a -eq 172 -and $b -ge 16 -and $b -le 31) { return $true }
    if ($a -eq 192 -and $b -eq 168) { return $true }
    if ($a -eq 169 -and $b -eq 254) { return $true }
    if ($a -eq 0 -or $a -ge 224) { return $true }        # 保留位址與多播
    return $false
}

$targets = New-Object System.Collections.Generic.List[string]
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$skipped = 0

function Add-Target([string]$raw) {
    $ip = Get-FirstIP $raw
    if (-not $ip) { return }
    if (Test-Private $ip) { $script:skipped++; return }
    if ($seen.Contains($ip)) { return }
    [void]$seen.Add($ip)
    $targets.Add($ip)
}

if ($InFile) {
    if (-not (Test-Path -LiteralPath $InFile)) { Write-Error "找不到檔案: $InFile"; exit 1 }
    foreach ($line in (Get-Content -LiteralPath $InFile)) {
        $l = $line.Trim()
        if (-not $l -or $l.StartsWith('#')) { continue }
        Add-Target $l
    }
}
if ($Ip) { foreach ($x in $Ip) { Add-Target $x } }

if ($targets.Count -eq 0) {
    Write-Error "沒有可查詢的 public IP。(私有位址會自動跳過,本次跳過 $skipped 個)"
    exit 1
}
Write-Host ("要查 {0} 個 public IP{1}" -f $targets.Count,
    $(if ($skipped) { "(已跳過 $skipped 個私有/保留位址)" } else { "" })) -ForegroundColor Green

# --------------------------------------------------------------------------- #
#  分類:雲端/VPS 是威脅獵捕最想先知道的一件事
# --------------------------------------------------------------------------- #
$CLOUD = 'amazon|aws|google|microsoft|azure|digitalocean|digital ocean|linode|vultr|ovh|' +
         'hetzner|contabo|scaleway|alibaba|aliyun|tencent|oracle|ibm cloud|leaseweb|' +
         'choopa|m247|hostwinds|namecheap|godaddy|dreamhost|rackspace|upcloud|kamatera'
$CDN   = 'cloudflare|akamai|fastly|incapsula|imperva|stackpath|limelight|edgecast|cdn77'
$TELCO = 'telecom|telkom|chunghwa|hinet|so-net|kddi|ntt|comcast|verizon|at&t|vodafone|' +
         'orange|deutsche telekom|telefonica|isp|broadband|communications'

function Get-Category([string]$text) {
    $t = $text.ToLower()
    if ($t -match $CLOUD) { return '雲端/VPS' }      # C2 與惡意基礎設施最常用
    if ($t -match $CDN)   { return 'CDN' }
    if ($t -match $TELCO) { return 'ISP/電信' }
    return ''
}

# --------------------------------------------------------------------------- #
#  RDAP 查詢
# --------------------------------------------------------------------------- #
# RDAP 回應的組織名在 entities[].vcardArray[1] 裡,格式是 [名稱, 參數, 型別, 值] 的陣列
function Get-VcardName($entity) {
    if (-not $entity) { return '' }
    $v = $entity.vcardArray
    if (-not $v -or $v.Count -lt 2) { return '' }
    foreach ($item in $v[1]) {
        if ($item -and $item.Count -ge 4 -and "$($item[0])" -eq 'fn') { return "$($item[3])" }
    }
    return ''
}
function Get-OrgName($resp) {
    if (-not $resp.entities) { return '' }
    # 優先取 registrant / administrative,沒有才退回第一個有名字的
    foreach ($want in @('registrant', 'administrative', 'technical')) {
        foreach ($e in $resp.entities) {
            if ($e.roles -and ($e.roles -contains $want)) {
                $n = Get-VcardName $e
                if ($n) { return $n }
            }
        }
    }
    foreach ($e in $resp.entities) {
        $n = Get-VcardName $e
        if ($n) { return $n }
    }
    return ''
}
function Get-Cidr($resp) {
    if ($resp.cidr0_cidrs -and $resp.cidr0_cidrs.Count -gt 0) {
        $c = $resp.cidr0_cidrs[0]
        if ($c.v4prefix) { return "$($c.v4prefix)/$($c.length)" }
        if ($c.v6prefix) { return "$($c.v6prefix)/$($c.length)" }
    }
    if ($resp.startAddress -and $resp.endAddress) { return "$($resp.startAddress) - $($resp.endAddress)" }
    return ''
}

$results = New-Object System.Collections.Generic.List[object]
$okCount = 0; $failCount = 0
$i = 0
foreach ($t in $targets) {
    $i++
    Write-Progress -Activity "RDAP 查詢中" -Status "$i / $($targets.Count)  $t" `
                   -PercentComplete ([int](100 * $i / $targets.Count))
    $row = [ordered]@{
        IP = $t; 組織 = ''; 網段名稱 = ''; 國家 = ''; 網段 = ''; 註冊機構 = ''; 分類 = ''; 狀態 = ''
    }
    try {
        $r = Invoke-RestMethod -Uri "https://rdap.org/ip/$t" -Method Get -TimeoutSec 20 `
                               -Headers @{ Accept = 'application/rdap+json' }
        $org = Get-OrgName $r
        $row.組織     = $org
        $row.網段名稱 = "$($r.name)"
        $row.國家     = "$($r.country)"
        $row.網段     = Get-Cidr $r
        if ($r.port43) { $row.註冊機構 = "$($r.port43)" }
        $row.分類     = Get-Category ("$org $($r.name) $($r.port43)")
        $row.狀態     = 'OK'
        $okCount++
    }
    catch {
        $row.狀態 = "失敗: " + $_.Exception.Message
        $failCount++
    }
    $results.Add([pscustomobject]$row)
    if ($DelayMs -gt 0 -and $i -lt $targets.Count) { Start-Sleep -Milliseconds $DelayMs }
}
Write-Progress -Activity "RDAP 查詢中" -Completed

# --------------------------------------------------------------------------- #
#  輸出。雲端/VPS 排前面 —— 那是最該先看的
# --------------------------------------------------------------------------- #
$order = @{ '雲端/VPS' = 0; 'ISP/電信' = 1; 'CDN' = 2; '' = 3 }
$sorted = $results | Sort-Object @{ Expression = { $order["$($_.分類)"] } }, 組織, IP

if ($OutFile) {
    $sorted | Export-Csv -LiteralPath $OutFile -NoTypeInformation -Encoding UTF8
    Write-Host ("已輸出 {0} 列到 {1}" -f $results.Count, $OutFile) -ForegroundColor Green
} else {
    $sorted | Format-Table -AutoSize
}

$cloud = @($results | Where-Object { $_.分類 -eq '雲端/VPS' }).Count
Write-Host ("成功 {0} · 失敗 {1}{2}" -f $okCount, $failCount,
    $(if ($cloud) { " · 其中 $cloud 個是雲端/VPS,建議優先查看" } else { "" })) `
    -ForegroundColor $(if ($failCount) { 'Yellow' } else { 'Green' })
if ($failCount -gt 0) {
    Write-Host "失敗多半是速率限制,可加大 -DelayMs 後只重跑失敗的那幾個。" -ForegroundColor Yellow
}
