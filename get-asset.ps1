#requires -Version 5.1
<#
.SYNOPSIS
    從 FortiGate 設定檔擷取 IP 與對應名稱，輸出成資產清單 (ip,名稱,角色)。

.DESCRIPTION
    純 PowerShell，Windows 內建即可執行，不需安裝任何東西。
    解析來源：
      * config firewall address     — 位址物件 (subnet / iprange / fqdn)
      * config system interface     — 防火牆各介面自身 IP
      * config firewall vip         — 虛擬 IP (對外 IP 與內部主機對應)
      * config reserved-address     — DHCP 保留位址 (常含設備描述)
    輸出可直接貼進 ip-graph.html 的「資產盤點表」欄位。

.EXAMPLE
    .\Get-FortiAssets.ps1 fgt.conf
    .\Get-FortiAssets.ps1 fgt.conf -OutFile assets.csv
    .\Get-FortiAssets.ps1 fgt.conf -ExpandRanges -OutFile assets.csv
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,
    [string]$OutFile,
    [switch]$ExpandRanges,      # 展開 iprange / 小網段成逐一 IP
    [int]$MaxExpand = 256,      # 單一物件最多展開幾個 IP
    [switch]$IncludeSubnets     # 也輸出網段與 FQDN 物件 (供人工參考)
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
if (-not (Test-Path -LiteralPath $Path)) { Write-Error "找不到檔案: $Path"; exit 1 }

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
# 一律用 ,@() 回傳：逗號運算子可避免 PowerShell 把單一元素陣列展開成純量
# (否則 $t[0] 會取到字串的「第一個字元」)
function Tokens([string]$s) {
    if (-not $s) { return ,@() }
    return ,@($s -split '[\s]+' | Where-Object { $_ })
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
#  把一個已收集完成的物件轉成資產項目
# --------------------------------------------------------------------------- #
function Emit-Object([string]$section, [string]$name, [hashtable]$o) {
    if (-not $name -or $null -eq $o) { return }

    switch ($section) {

        'firewall address' {
            $type = if ($o.ContainsKey('type')) { $o['type'] } else { 'ipmask' }
            $comment = ''
            if ($o.ContainsKey('comment')) { $comment = $o['comment'] }
            elseif ($o.ContainsKey('associated-interface')) { $comment = 'iface:' + $o['associated-interface'] }

            if ($type -eq 'iprange' -and $o.ContainsKey('start-ip') -and $o.ContainsKey('end-ip')) {
                $s = ConvertTo-UInt32IP $o['start-ip']
                $e = ConvertTo-UInt32IP $o['end-ip']
                if ($null -ne $s -and $null -ne $e -and $e -ge $s) {
                    $count = [int64]$e - [int64]$s + 1
                    if ($ExpandRanges -and $count -le $MaxExpand) {
                        for ($i = [uint32]$s; $i -le [uint32]$e; $i++) { Add-Entry (ConvertFrom-UInt32IP $i) $name $comment }
                    } else {
                        Add-Entry $o['start-ip'] $name ("range $($o['start-ip'])-$($o['end-ip']) $comment")
                    }
                }
            }
            elseif ($o.ContainsKey('subnet')) {
                $t = Tokens $o['subnet']
                if ($t.Count -ge 1) {
                    $ip = $t[0]
                    $mask = if ($t.Count -ge 2) { $t[1] } else { '255.255.255.255' }
                    if ($ip -eq '0.0.0.0') { return }                 # "all" 物件跳過
                    $plen = Get-PrefixLen $mask
                    if ($plen -eq 32) {
                        Add-Entry $ip $name $comment                  # 單一主機：最有用
                    }
                    elseif ($null -ne $plen) {
                        $net = ConvertTo-UInt32IP $ip
                        $size = [int64][math]::Pow(2, 32 - $plen)
                        if ($ExpandRanges -and $size -le $MaxExpand -and $null -ne $net) {
                            for ($i = [uint32]$net; $i -le [uint32]([int64]$net + $size - 1); $i++) {
                                Add-Entry (ConvertFrom-UInt32IP $i) $name $comment
                            }
                        }
                        elseif ($IncludeSubnets) {
                            Add-Entry "$ip/$plen" $name ("subnet $comment")
                        }
                    }
                }
            }
            elseif ($type -eq 'fqdn' -and $o.ContainsKey('fqdn') -and $IncludeSubnets) {
                Add-Entry $o['fqdn'] $name ("fqdn " + $comment)
            }
        }

        'system interface' {
            if ($o.ContainsKey('ip')) {
                $t = Tokens $o['ip']
                if ($t.Count -ge 1 -and $t[0] -ne '0.0.0.0') {
                    $desc = 'interface'
                    if ($o.ContainsKey('alias')) { $desc = 'iface:' + $o['alias'] }
                    elseif ($o.ContainsKey('description')) { $desc = $o['description'] }
                    Add-Entry $t[0] "FW-$name" $desc
                }
            }
        }

        'firewall vip' {
            if ($o.ContainsKey('extip')) {
                $t = Tokens ($o['extip'] -replace '-', ' ')
                if ($t.Count -ge 1) { Add-Entry $t[0] $name 'VIP-external' }
            }
            if ($o.ContainsKey('mappedip')) {
                $t = Tokens ((Unquote $o['mappedip']) -replace '-', ' ')
                if ($t.Count -ge 1) { Add-Entry $t[0] $name 'VIP-internal' }
            }
        }

        'reserved-address' {
            if ($o.ContainsKey('ip')) {
                $t = Tokens $o['ip']
                if ($t.Count -ge 1) {
                    $desc = if ($o.ContainsKey('description')) { $o['description'] } else { '' }
                    $nm = if ($desc) { $desc } else { "DHCP-$name" }
                    $role = if ($o.ContainsKey('mac')) { "dhcp-reserved " + $o['mac'] } else { 'dhcp-reserved' }
                    Add-Entry $t[0] $nm $role
                }
            }
        }
    }
}

# --------------------------------------------------------------------------- #
#  逐行解析：用「區段堆疊」正確處理巢狀 config
# --------------------------------------------------------------------------- #
$stack = New-Object System.Collections.Generic.List[object]   # 每層: @{Section; EditName; Obj}

function Push-Frame([string]$sec) {
    $stack.Add([pscustomobject]@{ Section = $sec; EditName = $null; Obj = @{} }) | Out-Null
}
function Top { if ($stack.Count -gt 0) { $stack[$stack.Count - 1] } else { $null } }
function Flush-Top {
    $f = Top
    if ($null -ne $f -and $f.EditName) {
        Emit-Object $f.Section $f.EditName $f.Obj
        $f.EditName = $null
        $f.Obj = @{}
    }
}

foreach ($raw in (Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)) {
    $line = $raw.Trim()
    if (-not $line -or $line.StartsWith('#')) { continue }

    if ($line -match '^config\s+(.+)$') {
        Push-Frame ((Unquote $Matches[1]).ToLower())
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
        $f = Top
        if ($null -ne $f) {
            Flush-Top
            $f.EditName = Unquote $Matches[1]
            $f.Obj = @{}
        }
        continue
    }
    if ($line -match '^set\s+(\S+)\s+(.+)$') {
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
#  輸出
# --------------------------------------------------------------------------- #
if ($results.Count -eq 0) {
    Write-Warning "沒有解析到任何位址物件。請確認這是 FortiGate 的 .conf 設定檔。"
    exit 2
}

$sorted = $results | Sort-Object @{ Expression = {
        $n = ConvertTo-UInt32IP $_.ip
        if ($null -eq $n) { [uint32]::MaxValue } else { $n }
    }
}
$out = foreach ($r in $sorted) {
    if ($r.role) { "$($r.ip),$($r.name),$($r.role)" } else { "$($r.ip),$($r.name)" }
}

if ($OutFile) {
    $out | Set-Content -LiteralPath $OutFile -Encoding UTF8
    Write-Host "已輸出 $($results.Count) 筆到 $OutFile" -ForegroundColor Green
    Write-Host "把檔案內容貼進 ip-graph.html 的「資產盤點表」欄位即可。"
} else {
    $out
    Write-Host ""
    Write-Host "共 $($results.Count) 筆。加 -OutFile assets.csv 可存檔；加 -ExpandRanges 可展開 IP 範圍。" -ForegroundColor Green
}
