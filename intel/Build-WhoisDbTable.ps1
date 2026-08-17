#requires -Version 5.1
<#
.SYNOPSIS
    把 WAFLogic/WHOIS-DB 這個 repo 裡所有的 WHOIS 原始記錄爬成一張
    「網段 -> 所屬機關」對照表(CSV),並可直接拿一堆 public IP 來比對。

.DESCRIPTION
    那個 repo 裡的資料散在好幾層:done/ 是 ARIN 原始輸出、registrars/<RIR>/ 是
    各區域註冊機構的記錄、大宗的 ARIN/RIPE/APNIC 還被壓在 zip 裡、
    blacklist_registrars/ 又整份複製一次。欄位名稱每家都不一樣
    (ARIN 用 NetRange/Organization、RIPE/APNIC 用 inetnum/descr、
    LACNIC 用 owner/inetrev、APNIC 系甚至把網段藏在 "% Information related to" 註解行)。

    這支把上面全部走一遍(含 zip,不落地解壓),統一成一張表:
        網段 / 起始IP / 結束IP / 組織 / 網段名稱 / 國家 / 註冊機構 / 分類

    同一個網段在多個檔重複出現時,留欄位最完整的那筆。
    ARIN 一次會回整條授權鏈(上層配置 + 下層指派),兩層都收 —— 查 IP 時取最精確的那層。

    產出的 CSV 完全離線可用,帶回內網就能查。

.PARAMETER RepoPath
    已下載好的 WHOIS-DB 目錄(git clone 或解壓 zip 後的資料夾)。

.PARAMETER Download
    自動從 GitHub 下載 repo 的 zip 再解析(需要網路)。與 -RepoPath 擇一。

.PARAMETER OutFile
    輸出的對照表 CSV,預設 .\whoisdb-ranges.csv。

.PARAMETER Table
    已經建好的對照表 CSV。給了這個就不重建,直接拿來查 —— 這是離線查詢的用法。

.PARAMETER Ip
    要查的 IP,可多個。會找**最精確**(前綴最長)的那個網段。

.PARAMETER InFile
    IP 清單檔,每行一個;每行只取第一個看起來像 IP 的字串,# 開頭忽略,
    所以可以直接餵 ip-graph 匯出的 CSV。

.PARAMETER MatchOut
    比對結果輸出 CSV。省略則印在畫面上。

.EXAMPLE
    # 有網路的機器上:建表
    .\Build-WhoisDbTable.ps1 -Download -OutFile whoisdb-ranges.csv

.EXAMPLE
    # 內網離線:拿建好的表查 IP
    .\Build-WhoisDbTable.ps1 -Table whoisdb-ranges.csv -InFile public-ips.txt -MatchOut who.csv

.NOTES
    這份資料的年份要看清楚:repo 最後更新是 2023-05,裡面的 WHOIS 檔案時間戳是 2015。
    IP 的所有權會換手,所以這張表是**線索**不是**權威**——
    命中了就當作一條可查的方向,查不到很正常(全表只有幾百個網段,
    而且偏 2015 年前後的惡意基礎設施)。要權威資料請用同資料夾的 Get-IpOwner.ps1 走 RDAP 即時查。

    repo 沒有附授權條款,所以這裡只放腳本、不轉存它的資料;要用請自己抓。

    本檔需存成 UTF-8 with BOM,否則 PowerShell 5.1 會用 Big5 解讀,繁中變亂碼。
#>
[CmdletBinding()]
param(
    [string]$RepoPath,
    [switch]$Download,
    [string]$OutFile = ".\whoisdb-ranges.csv",
    [string]$Table,
    [string[]]$Ip,
    [string]$InFile,
    [string]$MatchOut
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------- #
#  IP 換算
# --------------------------------------------------------------------------- #
# 註:全程用 double 做位址運算,不用 -shl/-shr/-band。
# PowerShell 5.1 的位元運算子對 [uint32] 會先提升型別,寫法容易踩到符號位與溢位;
# IPv4 只有 2^32,遠在 double 能精確表示的 2^53 之內,乘除法反而穩。
function ConvertTo-UInt32Ip([string]$ip) {
    if ($null -eq $ip) { return $null }
    $p = "$ip".Trim().Split('.')
    if ($p.Count -ne 4) { return $null }
    $n = [double]0
    foreach ($x in $p) {
        if ($x -notmatch '^\d{1,3}$' -or [int]$x -gt 255) { return $null }
        $n = $n * 256 + [int]$x
    }
    return [uint32]$n
}
function ConvertFrom-UInt32Ip($n) {
    $d = [double]$n
    '{0}.{1}.{2}.{3}' -f [int][math]::Floor($d / 16777216),
                         [int]([math]::Floor($d / 65536) % 256),
                         [int]([math]::Floor($d / 256) % 256),
                         [int]($d % 256)
}
# 網段大小 - 1(/24 -> 255)
function Get-BlockSpan([int]$len) {
    if ($len -le 0)  { return [double]4294967295 }
    if ($len -ge 32) { return [double]0 }
    return [double]([math]::Pow(2, 32 - $len) - 1)
}
function Get-MaskLen($start, $end) {
    # 起訖能不能剛好表示成一個 CIDR;不行回 -1
    $s = [double]$start; $e = [double]$end
    for ($len = 0; $len -le 32; $len++) {
        $span = Get-BlockSpan $len
        if (($s % ($span + 1)) -ne 0) { continue }      # 起點沒對齊這個大小
        if (($s + $span) -eq $e) { return $len }
    }
    return -1
}

# --------------------------------------------------------------------------- #
#  解析:各家 WHOIS 的欄位名稱不同,統一對照
# --------------------------------------------------------------------------- #
$RANGE_KEYS = @('netrange', 'inetnum', 'inetrev', 'iprange')
$CIDR_KEYS  = @('cidr', 'route')
$NAME_KEYS  = @('netname', 'netblockname')
# 依優先序取第一個有值的。org 排最後:RIPE/AfriNIC 的 org: 是代號(ORG-HOA1-RIPE),
# 真正看得懂的名字在 descr:(Hetzner Online AG),所以 descr 要排在 org 前面。
$ORG_KEYS   = @('organization', 'orgname', 'owner', 'customer', 'descr', 'org')
$RIRS       = @('ARIN', 'RIPE', 'APNIC', 'LACNIC', 'AfriNIC', 'TWNIC', 'KRNIC', 'Nic.br')

function ConvertTo-CleanValue([string]$v) {
    $v = ($v -replace '\s+', ' ').Trim()
    if ($v -eq '') { return '' }
    if (@('NULL', 'N/A', 'NA', '-') -contains $v.ToUpper()) { return '' }
    return $v
}

# 把 "1.2.3.0/24" / "1.2.3/24"(LACNIC 會這樣寫)/ "1.2.3.0 - 1.2.3.255" 都換成正規 CIDR
function ConvertTo-Cidr([string]$v) {
    $v = $v.Trim()
    if ($v -match '^(\d{1,3}(?:\.\d{1,3}){0,3})\s*/\s*(\d{1,2})$') {
        $a = $Matches[1]; $len = [int]$Matches[2]
        while ($a.Split('.').Count -lt 4) { $a += '.0' }     # 179.43.128/24 -> 179.43.128.0/24
        $n = ConvertTo-UInt32Ip $a
        if ($null -eq $n -or $len -gt 32) { return $null }
        $span = Get-BlockSpan $len
        $net  = [double]$n - ([double]$n % ($span + 1))       # 把低位歸零
        return (ConvertFrom-UInt32Ip $net) + "/$len"
    }
    return $null
}
function ConvertTo-CidrFromRange([string]$a, [string]$b) {
    $s = ConvertTo-UInt32Ip $a; $e = ConvertTo-UInt32Ip $b
    if ($null -eq $s -or $null -eq $e -or $e -lt $s) { return $null }
    $len = Get-MaskLen $s $e
    if ($len -ge 0) { return "$a/$len" }
    return "$a - $b"          # 不是整齊的 CIDR,原樣保留
}

function New-Rec { [pscustomobject]@{ Cidr=''; Range=''; Name=''; Org=@{}; Country=''; Source='' } }
function Test-RecEmpty($r) { return (-not $r.Cidr -and -not $r.Range) }
function Get-RecOrg($r) {
    foreach ($k in $ORG_KEYS) { if ($r.Org.ContainsKey($k) -and $r.Org[$k]) { return $r.Org[$k] } }
    return ''
}

function Read-WhoisText([string]$text, [string]$srcHint) {
    $out = New-Object System.Collections.Generic.List[object]
    $cur = New-Rec
    $fileSource = ''

    foreach ($raw in ($text -split "`r?`n")) {
        $line = $raw.TrimEnd()

        # APNIC / AfriNIC / TWNIC / KRNIC:網段只出現在這行註解裡,沒有 inetnum 欄位
        if ($line -match '^%\s*Information related to\s+([\d.]+)\s*-\s*([\d.]+)') {
            $a = $Matches[1]; $b = $Matches[2]
            if (-not (Test-RecEmpty $cur)) { $out.Add($cur); $cur = New-Rec; $cur.Source = $fileSource }
            $cur.Range = "$a - $b"
            $c = ConvertTo-CidrFromRange $a $b
            if ($c) { $cur.Cidr = $c }
            continue
        }
        if (-not $line -or $line[0] -eq '#' -or $line[0] -eq '%') { continue }
        if ($line -match '^([A-Za-z][A-Za-z0-9 _.-]*)\s*:\s*(.*)$') {
            $k = ($Matches[1].ToLower() -replace '[^a-z0-9]', '')
            $v = ConvertTo-CleanValue $Matches[2]
        } else { continue }
        if (-not $v) { continue }

        if ($k -eq 'source') {
            $v = ($v -split '[\s#%]')[0]                 # "RIPE # Filtered" -> "RIPE"
            $fileSource = $v
            if (-not $cur.Source) { $cur.Source = $v }
            continue
        }

        if ($RANGE_KEYS -contains $k -or $CIDR_KEYS -contains $k) {
            $c = ConvertTo-Cidr $v
            $isRange = $v -match '^(\d{1,3}(?:\.\d{1,3}){3})\s*-\s*(\d{1,3}(?:\.\d{1,3}){3})$'
            $rngText = ''
            if ($isRange) {
                $rngText = $v
                if (-not $c) { $c = ConvertTo-CidrFromRange $Matches[1] $Matches[2] }
            }
            if (-not $c -and -not $isRange) { continue }  # 不是網段寫法,忽略

            $keyNew = if ($c) { $c } else { $rngText }
            # 已經有網段而且不一樣 -> ARIN 回了下一筆記錄
            if (-not (Test-RecEmpty $cur) -and $keyNew -ne $cur.Cidr -and $keyNew -ne $cur.Range) {
                $out.Add($cur); $cur = New-Rec; $cur.Source = $fileSource
            }
            if ($c) { $cur.Cidr = $c }
            if ($rngText) { $cur.Range = $rngText }
            elseif ($c -and -not $cur.Range) { $cur.Range = $c }
            continue
        }

        if ($NAME_KEYS -contains $k) {
            if (-not $cur.Name) { $cur.Name = $v }
        }
        elseif ($k -eq 'country') {
            if (-not $cur.Country -and $v -match '^[A-Za-z]{2}$') { $cur.Country = $v.ToUpper() }
        }
        elseif ($ORG_KEYS -contains $k) {
            if (-not $cur.Org.ContainsKey($k)) { $cur.Org[$k] = $v }
        }
    }
    if (-not (Test-RecEmpty $cur)) { $out.Add($cur) }

    # done/ 底下的 ARIN 原始輸出沒有 source: 行,從內文的 whois 主機名推
    $guess = $srcHint
    if (-not $guess) {
        $low = $text.ToLower()
        foreach ($pair in @(@('whois.arin.net','ARIN'), @('ripe.net','RIPE'), @('apnic.net','APNIC'),
                            @('lacnic.net','LACNIC'), @('afrinic.net','AfriNIC'), @('twnic.net','TWNIC'),
                            @('kisa.or.kr','KRNIC'), @('registro.br','Nic.br'))) {
            if ($low.Contains($pair[0])) { $guess = $pair[1]; break }
        }
    }
    foreach ($r in $out) { if (-not $r.Source) { $r.Source = $guess } }
    return $out
}

# --------------------------------------------------------------------------- #
#  分類:雲端/VPS 是威脅獵捕最想先知道的一件事
# --------------------------------------------------------------------------- #
$CLOUD = 'amazon|aws|google|microsoft|azure|digitalocean|digital ocean|linode|vultr|ovh|' +
         'hetzner|contabo|scaleway|alibaba|aliyun|tencent|oracle|leaseweb|choopa|m247|' +
         'hostwinds|namecheap|godaddy|dreamhost|rackspace|upcloud|kamatera|hosting|host\b|' +
         'server|vps|colo|datacenter|data center'
$CDN   = 'cloudflare|akamai|fastly|incapsula|imperva|stackpath|limelight|edgecast|cdn'
$TELCO = 'telecom|telkom|chunghwa|hinet|so-net|kddi|\bntt\b|comcast|verizon|at&t|vodafone|' +
         'orange|deutsche telekom|telefonica|\bisp\b|broadband|communications|cable|wireless|mobile'
function Get-Category([string]$t) {
    if ($t -match $CLOUD) { return '雲端/VPS' }
    if ($t -match $CDN)   { return 'CDN' }
    if ($t -match $TELCO) { return 'ISP/電信' }
    return ''
}

# --------------------------------------------------------------------------- #
#  建表
# --------------------------------------------------------------------------- #
function Build-Table([string]$root) {
    $best = @{}          # 網段 -> @{ Score; Row }
    # 巢狀函式要改的計數器必須放在 script 範圍;放函式區域範圍的話,
    # 巢狀函式裡的 ++ 會另外開一個變數,外面永遠看到 0。
    $script:fileCount = 0
    $script:recCount  = 0

    $skipExt = @('.json', '.md', '.sh', '.png', '.jpg', '.gif', '.pdf')

    function Add-Rec($rec, $hint) {
        $key = if ($rec.Cidr) { $rec.Cidr } else { $rec.Range }
        if (-not $key) { return }
        $script:recCount++
        $org = Get-RecOrg $rec
        $score = 0
        foreach ($x in @($org, $rec.Name, $rec.Country)) { if ($x) { $score++ } }
        $old = $best[$key]
        if ($null -eq $old -or $score -gt $old.Score) {
            $best[$key] = @{
                Score = $score
                Row   = [pscustomobject]@{
                    Cidr = $key; Org = $org; Name = $rec.Name
                    Country = $rec.Country; Source = $rec.Source
                }
            }
        }
    }

    Write-Host "掃描 $root ..." -ForegroundColor Cyan
    $all = Get-ChildItem -LiteralPath $root -Recurse -File -Force |
           Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' }

    foreach ($f in $all) {
        $ext = $f.Extension.ToLower()
        if ($skipExt -contains $ext) { continue }

        # 路徑裡的 RIR 目錄名就是註冊機構
        $hint = ''
        $seg = $f.FullName -split '[\\/]'
        foreach ($r in $RIRS) { if ($seg -contains $r) { $hint = $r } }

        if ($ext -eq '.zip') {
            $zip = $null
            try { $zip = [IO.Compression.ZipFile]::OpenRead($f.FullName) }
            catch { Write-Warning "讀不了 zip: $($f.FullName)"; continue }
            try {
                foreach ($entry in $zip.Entries) {
                    if (-not $entry.Name) { continue }          # 目錄項
                    $sr = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
                    try { $text = $sr.ReadToEnd() } finally { $sr.Dispose() }
                    $script:fileCount++
                    foreach ($rec in (Read-WhoisText $text $hint)) { Add-Rec $rec $hint }
                }
            } finally { $zip.Dispose() }
        }
        else {
            $text = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
            $script:fileCount++
            foreach ($rec in (Read-WhoisText $text $hint)) { Add-Rec $rec $hint }
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($k in $best.Keys) {
        $r = $best[$k].Row
        $start = ''; $end = ''; $sortKey = [double]0; $endKey = [double]0
        if ($k -match '^(\d{1,3}(?:\.\d{1,3}){3})/(\d{1,2})$') {
            $n = ConvertTo-UInt32Ip $Matches[1]; $len = [int]$Matches[2]
            $span = Get-BlockSpan $len
            $start = ConvertFrom-UInt32Ip $n
            $end   = ConvertFrom-UInt32Ip ([double]$n + $span)
            $sortKey = [double]$n; $endKey = [double]$n + $span
        }
        elseif ($k -match '^(\d{1,3}(?:\.\d{1,3}){3}) - (\d{1,3}(?:\.\d{1,3}){3})$') {
            $start = $Matches[1]; $end = $Matches[2]
            $sortKey = [double](ConvertTo-UInt32Ip $start); $endKey = [double](ConvertTo-UInt32Ip $end)
        }
        $rows.Add([pscustomobject][ordered]@{
            網段 = $k; 起始IP = $start; 結束IP = $end
            組織 = $r.Org; 網段名稱 = $r.Name; 國家 = $r.Country
            註冊機構 = $r.Source; 分類 = (Get-Category "$($r.Org) $($r.Name)")
            _s = $sortKey; _e = $endKey
        })
    }
    Write-Host ("掃描 {0} 個檔 · 解析出 {1} 筆記錄 · 去重後 {2} 個網段" -f
        $script:fileCount, $script:recCount, $rows.Count) -ForegroundColor Green
    # 起點小的在前;起點相同時大的網段在前(上層配置排在它底下的指派之前)。
    # 一定要有第二排序鍵,否則同起點的先後取決於雜湊表列舉順序,每次跑結果會不一樣。
    return ($rows | Sort-Object @{Expression='_s'}, @{Expression='_e'; Descending=$true} |
            Select-Object * -ExcludeProperty _s, _e)
}

# --------------------------------------------------------------------------- #
#  輸出 CSV
# --------------------------------------------------------------------------- #
# Excel 要靠 BOM 才知道檔案是 UTF-8,沒有 BOM 中文欄名會變亂碼。
# 但 -Encoding UTF8 在 5.1 是「含 BOM」、在 7 以後變成「不含 BOM」,同一句話兩種結果,
# 所以這裡明講要哪一種,不靠版本預設。
function Export-CsvUtf8Bom($data, [string]$path) {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $data | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8BOM
    } else {
        $data | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    }
}

# --------------------------------------------------------------------------- #
#  查詢:最長前綴優先(ARIN 的上下層都在表裡,要取最精確的那筆)
# --------------------------------------------------------------------------- #
function Get-FirstIP([string]$s) {
    if ($s -match '(?<![\d.])(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?![\d.])') {
        if ([int]$Matches[1] -le 255 -and [int]$Matches[2] -le 255 -and
            [int]$Matches[3] -le 255 -and [int]$Matches[4] -le 255) { return $Matches[0] }
    }
    return $null
}
function Resolve-IpOwner($rows, [string]$ip) {
    $n = ConvertTo-UInt32Ip $ip
    if ($null -eq $n) { return $null }
    $hit = $null; $bestSize = [double]::PositiveInfinity
    foreach ($row in $rows) {
        $s = ConvertTo-UInt32Ip $row.起始IP
        $e = ConvertTo-UInt32Ip $row.結束IP
        if ($null -eq $s -or $null -eq $e) { continue }
        if ($n -ge $s -and $n -le $e) {
            $size = [double]$e - [double]$s
            if ($size -lt $bestSize) { $bestSize = $size; $hit = $row }   # 範圍越小越精確
        }
    }
    return $hit
}

# --------------------------------------------------------------------------- #
#  主流程
# --------------------------------------------------------------------------- #
# 注意:PowerShell 變數不分大小寫,承接資料的變數不能叫 $table —— 那就是 -Table 參數本身。
$db = $null

if ($Table) {
    if (-not (Test-Path -LiteralPath $Table)) { Write-Error "找不到對照表: $Table"; exit 1 }
    $db = @(Import-Csv -LiteralPath $Table -Encoding UTF8)
    Write-Host "載入對照表 $($db.Count) 個網段" -ForegroundColor Green
}
else {
    $root = $RepoPath
    $tmp = $null
    if ($Download) {
        try {
            [Net.ServicePointManager]::SecurityProtocol =
                [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
        } catch {}
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ("whoisdb_" + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tmp -Force | Out-Null
        $zipPath = Join-Path $tmp 'repo.zip'
        Write-Host "下載 WHOIS-DB ..." -ForegroundColor Cyan
        try {
            $iwr = @{
                Uri = 'https://github.com/WAFLogic/WHOIS-DB/archive/refs/heads/main.zip'
                OutFile = $zipPath; UseBasicParsing = $true
            }
            # 公司內部通常要走 Proxy;有設環境變數就沿用,順便帶上目前帳號的認證
            if ($env:HTTPS_PROXY) {
                $iwr.Proxy = $env:HTTPS_PROXY
                $iwr.ProxyUseDefaultCredentials = $true
            }
            Invoke-WebRequest @iwr
            if (-not (Test-Path -LiteralPath $zipPath) -or (Get-Item -LiteralPath $zipPath).Length -lt 100000) {
                throw "下載的檔案不完整"
            }
            Expand-Archive -LiteralPath $zipPath -DestinationPath $tmp -Force
            Remove-Item -LiteralPath $zipPath -Force
        }
        catch {
            # 這裡不能吞掉錯誤:失敗了還往下走,只會產出一張空表看起來卻像成功
            Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
            Write-Error ("下載失敗: {0}`n改用瀏覽器抓 https://github.com/WAFLogic/WHOIS-DB 的 ZIP,解壓後用 -RepoPath 指過去。" -f $_.Exception.Message)
            exit 1
        }
        $root = $tmp
    }
    if (-not $root) {
        Write-Error "請給 -RepoPath(已下載的 WHOIS-DB 目錄)或 -Download(自動下載),或 -Table(用現成的表查)。"
        exit 1
    }
    if (-not (Test-Path -LiteralPath $root)) { Write-Error "找不到目錄: $root"; exit 1 }

    $db = @(Build-Table $root)
    Export-CsvUtf8Bom $db $OutFile
    Write-Host "已輸出對照表: $OutFile" -ForegroundColor Green

    if ($tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

# --------------------------------------------------------------------------- #
#  有給 IP 才做比對
# --------------------------------------------------------------------------- #
$targets = New-Object System.Collections.Generic.List[string]
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$skipped = 0
# 內網位址查了也沒意義,這個庫只收 public IP —— 直接跳過,免得清單被自家 IP 灌滿
function Test-Private([string]$ip) {
    $p = $ip.Split('.')
    if ($p.Count -ne 4) { return $true }
    $a = [int]$p[0]; $b = [int]$p[1]
    if ($a -eq 10 -or $a -eq 127) { return $true }
    if ($a -eq 172 -and $b -ge 16 -and $b -le 31) { return $true }
    if ($a -eq 192 -and $b -eq 168) { return $true }
    if ($a -eq 169 -and $b -eq 254) { return $true }
    if ($a -eq 0 -or $a -ge 224) { return $true }
    return $false
}
function Add-Target([string]$raw) {
    $x = Get-FirstIP $raw
    if (-not $x) { return }
    if (Test-Private $x) { $script:skipped++; return }
    if ($seen.Add($x)) { $targets.Add($x) }
}
if ($InFile) {
    if (-not (Test-Path -LiteralPath $InFile)) { Write-Error "找不到檔案: $InFile"; exit 1 }
    foreach ($line in (Get-Content -LiteralPath $InFile)) {
        $l = $line.Trim()
        if (-not $l -or $l.StartsWith('#')) { continue }
        Add-Target $l
    }
}
# -Ip 可以是陣列,也可以是一整串 "a,b,c"(用 pwsh -File 呼叫時就會變成後者),兩種都收
if ($Ip) { foreach ($x in $Ip) { foreach ($piece in ($x -split '[,;\s]+')) { Add-Target $piece } } }

if ($targets.Count -eq 0) { return }

$hits = New-Object System.Collections.Generic.List[object]
$hitCount = 0
foreach ($t in $targets) {
    $row = Resolve-IpOwner $db $t
    if ($row) {
        $hitCount++
        $hits.Add([pscustomobject][ordered]@{
            IP = $t; 組織 = $row.組織; 網段名稱 = $row.網段名稱; 國家 = $row.國家
            網段 = $row.網段; 註冊機構 = $row.註冊機構; 分類 = $row.分類; 狀態 = '命中'
        })
    } else {
        $hits.Add([pscustomobject][ordered]@{
            IP = $t; 組織 = ''; 網段名稱 = ''; 國家 = ''; 網段 = ''
            註冊機構 = ''; 分類 = ''; 狀態 = '此庫查無'
        })
    }
}
# 雲端/VPS 排最前面 —— C2 與惡意基礎設施最常架在那裡
$order = @{ '雲端/VPS' = 0; 'ISP/電信' = 1; 'CDN' = 2; '' = 3 }
$sorted = $hits | Sort-Object @{ Expression = { if ($_.狀態 -eq '命中') { 0 } else { 1 } } },
                              @{ Expression = { $order["$($_.分類)"] } }, 組織, IP

if ($MatchOut) {
    Export-CsvUtf8Bom $sorted $MatchOut
    Write-Host "比對結果已輸出: $MatchOut" -ForegroundColor Green
} else {
    # 指定寬度:非互動主控台(排程、重導向)的預設寬度可能是 0,直接 Format-Table 會印出空白
    $sorted | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
}
Write-Host ("查 {0} 個 public IP,命中 {1},查無 {2}{3}" -f $targets.Count, $hitCount,
    ($targets.Count - $hitCount),
    $(if ($script:skipped) { "(另跳過 $($script:skipped) 個私有/保留位址)" } else { "" })
    ) -ForegroundColor $(if ($hitCount) { 'Green' } else { 'Yellow' })
if ($hitCount -lt $targets.Count) {
    Write-Host "查無的請用 Get-IpOwner.ps1 走 RDAP 即時查 —— 這個庫只有幾百個網段,而且是 2015 年前後的資料。" -ForegroundColor Yellow
}
