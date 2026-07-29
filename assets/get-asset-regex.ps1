#requires -Version 5.1
<#
.SYNOPSIS
    從 FortiGate 設定檔擷取 IP → 名稱,輸出資產清單 (ip,名稱,角色)。
    Regex 逐行版:先抓 edit 物件名,再從白名單 set 鍵用 regex 把 IP 抓出來做對應。

.DESCRIPTION
    純 PowerShell,Windows 5.1 內建即可執行,零依賴、離線可用。
    輸出可直接貼進 ip-graph.html 的「資產盤點表」欄位。

    與 stack 版 (get-asset.ps1) 的差別:
      * IP 一律用 regex 抓第一個 IPv4,天生避開「subnet 後面那段遮罩」的誤判。
      * 自動偵測編碼:UTF-8 (含/不含 BOM)、UTF-16 LE/BE、無 BOM 退回系統 ANSI (zh-TW=Big5)。
      * 解析不到時輸出診斷(讀入編碼、非空行數、config/edit/set 筆數、前 8 行樣本、下一步建議)。

    仍保留「分節堆疊」處理巢狀 config —— vip 的 realservers、interface 的 secondaryip、
    dhcp server 底下的 reserved-address 都是巢狀,靠堆疊才不會把子物件的 edit 名當成資產名。

    解析來源:
      config firewall address    位址物件 (subnet /32 主機、iprange)
      config firewall vip        虛擬 IP (extip 對外 / mappedip 內部)
      config system interface    介面自身 IP
      config ... reserved-address DHCP 保留 (IoMT 設備名常在這)

.EXAMPLE
    .\get-asset-regex.ps1 fgt.conf
    .\get-asset-regex.ps1 fgt.conf -OutFile assets.csv
    .\get-asset-regex.ps1 fgt.conf -ExpandRanges -IncludeSubnets -OutFile assets.csv

.NOTES
    本檔需存成「UTF-8 with BOM」,否則 PowerShell 5.1 會用 Big5 解讀,繁中註解變亂碼。
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,
    [string]$OutFile,
    [switch]$ExpandRanges,      # 展開 iprange / 小網段成逐一 IP
    [int]$MaxExpand = 256,      # 單一物件最多展開幾個 IP
    [switch]$IncludeSubnets     # 也輸出網段與 FQDN 物件 (供人工參考;精確比對用不到)
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
if (-not (Test-Path -LiteralPath $Path)) { Write-Error "找不到檔案: $Path"; exit 1 }

# --------------------------------------------------------------------------- #
#  編碼自動偵測:PowerShell 5.1 的 Get-Content -Encoding UTF8 不會認 UTF-16,
#  這裡自己讀 bytes 判斷 BOM;無 BOM 時用 null-byte 分布猜 UTF-16,再退回 UTF-8/ANSI。
# --------------------------------------------------------------------------- #
$script:UsedEnc = 'UTF-8'
function Read-TextAuto([string]$p) {
    $bytes = [System.IO.File]::ReadAllBytes($p)
    if ($bytes.Length -eq 0) { return '' }

    # 1) 有 BOM 直接認
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $script:UsedEnc = 'UTF-8 (BOM)'
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $script:UsedEnc = 'UTF-16 LE (BOM)'
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $script:UsedEnc = 'UTF-16 BE (BOM)'
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }

    # 2) 無 BOM:看 null byte 分布猜 UTF-16。ASCII 字在 LE 的高位(奇數位)、BE 的高位(偶數位)為 0。
    $lim = [Math]::Min($bytes.Length, 4000)
    $evenNull = 0; $oddNull = 0
    for ($i = 0; $i -lt $lim; $i++) {
        if ($bytes[$i] -eq 0) {
            if ($i % 2 -eq 0) { $evenNull++ } else { $oddNull++ }
        }
    }
    if (($evenNull + $oddNull) -gt ($lim / 4)) {
        if ($oddNull -gt $evenNull) {
            $script:UsedEnc = 'UTF-16 LE (無 BOM,推測)'
            return [System.Text.Encoding]::Unicode.GetString($bytes)
        }
        $script:UsedEnc = 'UTF-16 BE (無 BOM,推測)'
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
    }

    # 3) 無 BOM 且非 UTF-16:先試 UTF-8,出現替代字元 U+FFFD 就退回系統 ANSI (zh-TW = Big5)
    $utf8 = New-Object System.Text.UTF8Encoding($false, $false)
    $text = $utf8.GetString($bytes)
    if ($text.IndexOf([char]0xFFFD) -ge 0) {
        $script:UsedEnc = 'ANSI/Big5 (UTF-8 解碼失敗後退回)'
        return [System.Text.Encoding]::Default.GetString($bytes)
    }
    $script:UsedEnc = 'UTF-8 (無 BOM)'
    return $text
}

# --------------------------------------------------------------------------- #
#  工具函式
# --------------------------------------------------------------------------- #
function ConvertTo-UInt32IP([string]$ip) {
    if (-not $ip) { return $null }
    $p = $ip.Split('.')
    if ($p.Count -ne 4) { return $null }
    foreach ($x in $p) { if ($x -notmatch '^\d{1,3}$' -or [int]$x -gt 255) { return $null } }
    return ([uint32]$p[0] -shl 24) -bor ([uint32]$p[1] -shl 16) -bor ([uint32]$p[2] -shl 8) -bor [uint32]$p[3]
}
function ConvertFrom-UInt32IP([uint32]$n) {
    '{0}.{1}.{2}.{3}' -f (($n -shr 24) -band 255), (($n -shr 16) -band 255), (($n -shr 8) -band 255), ($n -band 255)
}
function Get-PrefixLen([string]$mask) {
    $n = ConvertTo-UInt32IP $mask
    if ($null -eq $n) { return $null }
    $bits = 0
    for ($i = 31; $i -ge 0; $i--) { if ($n -band ([uint32]1 -shl $i)) { $bits++ } else { break } }
    return $bits
}
function Unquote([string]$s) {
    if ($null -eq $s) { return "" }
    $s = $s.Trim()
    if ($s.Length -ge 2 -and $s.StartsWith('"') -and $s.EndsWith('"')) { $s = $s.Substring(1, $s.Length - 2) }
    return $s.Trim()
}
function Clean([string]$s) {
    if ($null -eq $s) { return "" }
    ($s -replace '[,\r\n]', ' ').Trim()
}
# 用 regex 抓字串裡「第一個」合法 IPv4;左右加負向斷言,避免咬到大數字的中段。
function Get-FirstIP([string]$s) {
    if (-not $s) { return $null }
    if ($s -match '(?<![\d.])(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?![\d.])') {
        if ([int]$Matches[1] -le 255 -and [int]$Matches[2] -le 255 -and
            [int]$Matches[3] -le 255 -and [int]$Matches[4] -le 255) {
            return $Matches[0]
        }
    }
    return $null
}

$results = New-Object System.Collections.Generic.List[object]
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
function Add-Entry([string]$ip, [string]$name, [string]$role) {
    if (-not $ip -or -not $name) { return }
    $key = "$ip|$name"
    if ($seen.Contains($key)) { return }
    [void]$seen.Add($key)
    $results.Add([pscustomobject]@{ ip = $ip; name = (Clean $name); role = (Clean $role) })
}

# --------------------------------------------------------------------------- #
#  把一個已收集完成的 edit 物件轉成資產項目(IP 一律經 Get-FirstIP)
# --------------------------------------------------------------------------- #
function Emit-Object([string]$section, [string]$name, [hashtable]$o) {
    if (-not $name -or $null -eq $o) { return }

    switch ($section) {

        'firewall address' {
            $comment = ''
            if ($o.ContainsKey('comment')) { $comment = $o['comment'] }
            elseif ($o.ContainsKey('associated-interface')) { $comment = 'iface:' + $o['associated-interface'] }
            $type = if ($o.ContainsKey('type')) { $o['type'] } else { '' }

            if ($o.ContainsKey('start-ip') -and $o.ContainsKey('end-ip')) {
                $sip = Get-FirstIP $o['start-ip']
                $eip = Get-FirstIP $o['end-ip']
                $s = ConvertTo-UInt32IP $sip
                $e = ConvertTo-UInt32IP $eip
                if ($null -ne $s -and $null -ne $e -and $e -ge $s) {
                    $count = [int64]$e - [int64]$s + 1
                    if ($ExpandRanges -and $count -le $MaxExpand) {
                        for ($i = [uint32]$s; $i -le [uint32]$e; $i++) { Add-Entry (ConvertFrom-UInt32IP $i) $name $comment }
                    } else {
                        Add-Entry $sip $name ("range $sip-$eip $comment")
                    }
                }
            }
            elseif ($o.ContainsKey('subnet')) {
                $parts = @($o['subnet'].Trim() -split '\s+')     # @() 保證是陣列,避免 5.1 純量陷阱
                $ip = Get-FirstIP $parts[0]
                $mask = if ($parts.Count -ge 2) { $parts[1] } else { '255.255.255.255' }
                if ($ip -and $ip -ne '0.0.0.0') {                # "all" 物件跳過
                    $plen = Get-PrefixLen $mask
                    if ($plen -eq 32 -or $null -eq $plen) {
                        Add-Entry $ip $name $comment             # 單一主機:精確比對最有用
                    }
                    elseif ($ExpandRanges) {
                        $net = ConvertTo-UInt32IP $ip
                        $size = [int64][math]::Pow(2, 32 - $plen)
                        if ($size -le $MaxExpand -and $null -ne $net) {
                            for ($i = [uint32]$net; $i -le [uint32]([int64]$net + $size - 1); $i++) {
                                Add-Entry (ConvertFrom-UInt32IP $i) $name $comment
                            }
                        }
                        elseif ($IncludeSubnets) {
                            Add-Entry "$ip/$plen" $name ("subnet $comment")
                        }
                    }
                    elseif ($IncludeSubnets) {
                        Add-Entry "$ip/$plen" $name ("subnet $comment")
                    }
                }
            }
            elseif ($type -eq 'fqdn' -and $o.ContainsKey('fqdn') -and $IncludeSubnets) {
                Add-Entry $o['fqdn'] $name ("fqdn " + $comment)
            }
        }

        'system interface' {
            if ($o.ContainsKey('ip')) {
                $ip = Get-FirstIP $o['ip']
                if ($ip -and $ip -ne '0.0.0.0') {
                    $desc = 'interface'
                    if ($o.ContainsKey('alias')) { $desc = 'iface:' + $o['alias'] }
                    elseif ($o.ContainsKey('description')) { $desc = $o['description'] }
                    Add-Entry $ip "FW-$name" $desc
                }
            }
        }

        'firewall vip' {
            if ($o.ContainsKey('extip')) {
                $ip = Get-FirstIP $o['extip']
                if ($ip) { Add-Entry $ip $name 'VIP-external' }
            }
            if ($o.ContainsKey('mappedip')) {
                $ip = Get-FirstIP $o['mappedip']
                if ($ip) { Add-Entry $ip $name 'VIP-internal' }
            }
        }

        'reserved-address' {
            if ($o.ContainsKey('ip')) {
                $ip = Get-FirstIP $o['ip']
                if ($ip) {
                    $desc = if ($o.ContainsKey('description')) { $o['description'] } else { '' }
                    $nm = if ($desc) { $desc } else { "DHCP-$name" }
                    $role = if ($o.ContainsKey('mac')) { "dhcp-reserved " + $o['mac'] } else { 'dhcp-reserved' }
                    Add-Entry $ip $nm $role
                }
            }
        }
    }
}

# --------------------------------------------------------------------------- #
#  逐行解析:regex 抓 config/edit/set,用分節堆疊正確處理巢狀
# --------------------------------------------------------------------------- #
$stat = @{ lines = 0; config = 0; edit = 0; set = 0 }
$text = Read-TextAuto $Path
$lines = @($text -split "\r?\n")

$stack = New-Object System.Collections.Generic.List[object]   # 每層: @{Section; EditName; Obj}
function Top { if ($stack.Count -gt 0) { $stack[$stack.Count - 1] } else { $null } }
function Flush-Top {
    $f = Top
    if ($null -ne $f -and $f.EditName) {
        Emit-Object $f.Section $f.EditName $f.Obj
        $f.EditName = $null
        $f.Obj = @{}
    }
}

foreach ($raw in $lines) {
    $line = $raw.Trim()
    if (-not $line -or $line.StartsWith('#')) { continue }
    $stat.lines++

    if ($line -match '^config\s+(.+)$') {
        $stat.config++
        $stack.Add([pscustomobject]@{ Section = (Unquote $Matches[1]).ToLower(); EditName = $null; Obj = @{} }) | Out-Null
        continue
    }
    if ($line -eq 'end') {
        Flush-Top
        if ($stack.Count -gt 0) { $stack.RemoveAt($stack.Count - 1) }
        continue
    }
    if ($line -eq 'next') {
        Flush-Top
        continue
    }
    if ($line -match '^edit\s+(.+)$') {
        $stat.edit++
        $f = Top
        if ($null -ne $f) {
            Flush-Top
            $f.EditName = Unquote $Matches[1]
            $f.Obj = @{}
        }
        continue
    }
    if ($line -match '^set\s+(\S+)\s+(.+)$') {
        $stat.set++
        $f = Top
        if ($null -ne $f -and $f.EditName) {
            $k = $Matches[1].ToLower()
            if (-not $f.Obj.ContainsKey($k)) { $f.Obj[$k] = Unquote $Matches[2] }
        }
        continue
    }
}
while ($stack.Count -gt 0) { Flush-Top; $stack.RemoveAt($stack.Count - 1) }

# --------------------------------------------------------------------------- #
#  輸出(或空結果時的診斷)
# --------------------------------------------------------------------------- #
if ($results.Count -eq 0) {
    Write-Warning "沒有解析到任何可用的資產項目。"
    Write-Host ""
    Write-Host "---- 診斷 --------------------------------" -ForegroundColor Yellow
    Write-Host ("檔案       : {0}" -f $Path)
    Write-Host ("讀入編碼   : {0}" -f $script:UsedEnc)
    Write-Host ("非空白行數 : {0}" -f $stat.lines)
    Write-Host ("config 區段: {0}" -f $stat.config)
    Write-Host ("edit 物件  : {0}" -f $stat.edit)
    Write-Host ("set 設定行 : {0}" -f $stat.set)
    Write-Host ""
    Write-Host "前 8 行(去註解/空白):"
    $sample = @($lines | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') } | Select-Object -First 8)
    foreach ($s in $sample) { Write-Host ("  " + $s.Trim()) }
    Write-Host ""
    if ($stat.edit -gt 0) {
        Write-Host "-> 有讀到 edit 物件卻沒輸出:多半是位址物件都是網段(非 /32 主機)。" -ForegroundColor Cyan
        Write-Host ("   試試:  .\get-asset-regex.ps1 `"{0}`" -IncludeSubnets    或加 -ExpandRanges" -f $Path) -ForegroundColor Cyan
    }
    elseif ($stat.config -eq 0) {
        Write-Host "-> 完全沒讀到 config 區段:可能是編碼判斷錯誤,或這不是 FortiGate 設定檔。" -ForegroundColor Cyan
        Write-Host "   請看上面『讀入編碼』與樣本行是否為亂碼。" -ForegroundColor Cyan
    }
    exit 2
}

$sorted = $results | Sort-Object @{ Expression = {
        $n = ConvertTo-UInt32IP $_.ip
        if ($null -eq $n) { [uint32]::MaxValue } else { $n }
    }
}, name
$out = foreach ($r in $sorted) {
    if ($r.role) { "$($r.ip),$($r.name),$($r.role)" } else { "$($r.ip),$($r.name)" }
}

if ($OutFile) {
    $out | Set-Content -LiteralPath $OutFile -Encoding UTF8
    Write-Host "已輸出 $($results.Count) 筆到 $OutFile  (編碼: $script:UsedEnc)" -ForegroundColor Green
    Write-Host "把檔案內容貼進 ip-graph.html 的「資產盤點表」欄位即可。"
} else {
    $out
    Write-Host ""
    Write-Host "共 $($results.Count) 筆 (編碼: $script:UsedEnc)。加 -OutFile assets.csv 可存檔;加 -ExpandRanges 展開範圍;加 -IncludeSubnets 一併列網段。" -ForegroundColor Green
}
