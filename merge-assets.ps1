#requires -Version 5.1
<#
.SYNOPSIS
    把防火牆擷取的「ip,名稱,角色」清單合併進 Excel 資產表:
    已有的 IP 補上設備名稱、沒有的在底下新增,所有更動一律標成紅字。

.DESCRIPTION
    純 PowerShell 5.1 + Excel COM(機器上有裝 Excel 即可,零安裝、離線可用)。

    合併規則:
      * Excel 裡已有該 IP、名稱欄是空的  -> 補上名稱(紅字)
      * Excel 裡已有該 IP、名稱相同      -> 不動
      * Excel 裡已有該 IP、名稱不同      -> 預設不覆蓋、列入報告;加 -Overwrite 才覆蓋(紅字)
      * Excel 裡沒有該 IP                -> 在資料最後新增一列(紅字)

    安全:預設「不動你的原檔」,另存成 <原檔名>-合併.xlsx;
    要直接改原檔請加 -InPlace(建議先備份)。執行前請先關閉該 Excel 檔。

.PARAMETER FirewallCsv
    防火牆擷取的清單(get-asset.ps1 / get-asset-regex.ps1 的輸出),
    每行「ip,名稱[,角色]」,# 開頭的行忽略。

.PARAMETER Excel
    你的資產表 .xlsx / .xlsm / .xls。

.PARAMETER Sheet
    工作表名稱或序號(預設 1 = 第一張)。

.PARAMETER IpColumn / NameColumn / RoleColumn
    欄位序號(A=1, B=2, ...)。不指定時自動偵測:
    先看第 1 列標題(ip / 名稱 / name / 設備 / 主機...),認不出 IP 欄再掃內容找最像 IP 的欄。
    RoleColumn 未指定時,角色資訊不寫入。

.PARAMETER HeaderRow
    標題列所在列號,預設 1。資料從下一列開始。

.PARAMETER Overwrite
    Excel 已有名稱但與防火牆不同時,覆蓋成防火牆的名稱(紅字)。預設不覆蓋。

.PARAMETER InPlace
    直接存回原檔。預設另存 <原檔名>-合併.xlsx。

.EXAMPLE
    .\merge-assets.ps1 -FirewallCsv assets.csv -Excel 資產清冊.xlsx

.EXAMPLE
    .\merge-assets.ps1 -FirewallCsv assets.csv -Excel 資產清冊.xlsx `
        -Sheet "伺服器" -IpColumn 2 -NameColumn 3 -Overwrite -InPlace

.NOTES
    需要本機安裝 Excel(用 COM 自動化;不裝任何套件)。
    本檔需存成 UTF-8 with BOM,否則 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$FirewallCsv,
    [Parameter(Mandatory = $true, Position = 1)]
    [string]$Excel,
    [object]$Sheet = 1,
    [int]$IpColumn = 0,
    [int]$NameColumn = 0,
    [int]$RoleColumn = 0,
    [int]$HeaderRow = 1,
    [switch]$Overwrite,
    [switch]$InPlace
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if (-not (Test-Path -LiteralPath $FirewallCsv)) { Write-Error "找不到防火牆清單: $FirewallCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $Excel))       { Write-Error "找不到 Excel 資產表: $Excel"; exit 1 }
$ExcelFull = (Resolve-Path -LiteralPath $Excel).Path

# --------------------------------------------------------------------------- #
#  讀防火牆清單 (ip,名稱[,角色])
# --------------------------------------------------------------------------- #
function Get-FirstIP([string]$s) {
    if (-not $s) { return $null }
    if ($s -match '(?<![\d.])(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?![\d.])') {
        if ([int]$Matches[1] -le 255 -and [int]$Matches[2] -le 255 -and
            [int]$Matches[3] -le 255 -and [int]$Matches[4] -le 255) { return $Matches[0] }
    }
    return $null
}
$fw = New-Object System.Collections.Generic.List[object]
$fwSeen = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($raw in (Get-Content -LiteralPath $FirewallCsv -Encoding UTF8)) {
    $line = $raw.Trim()
    if (-not $line -or $line.StartsWith('#')) { continue }
    $p = $line.Split(',')
    if ($p.Count -lt 2) { continue }
    $ip = Get-FirstIP $p[0].Trim()
    $name = $p[1].Trim()
    $role = ''
    if ($p.Count -ge 3) { $role = ($p[2..($p.Count - 1)] -join ',').Trim() }
    if (-not $ip -or -not $name) { continue }
    if ($fwSeen.Contains($ip)) { continue }          # 同 IP 多筆取第一筆
    [void]$fwSeen.Add($ip)
    $fw.Add([pscustomobject]@{ ip = $ip; name = $name; role = $role })
}
if ($fw.Count -eq 0) { Write-Error "防火牆清單沒有解析到任何「ip,名稱」資料。"; exit 1 }
Write-Host ("防火牆清單:{0} 筆不重複 IP" -f $fw.Count)

# --------------------------------------------------------------------------- #
#  開 Excel (COM)
# --------------------------------------------------------------------------- #
$xl = $null; $wb = $null; $ws = $null
try {
    $xl = New-Object -ComObject Excel.Application
} catch {
    Write-Error "無法啟動 Excel COM — 這台機器需要安裝 Excel 才能執行本工具。"
    exit 1
}
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($ExcelFull)
    try { $ws = $wb.Worksheets.Item($Sheet) }
    catch { throw "找不到工作表『$Sheet』。這本活頁簿的工作表:" + (@($wb.Worksheets | ForEach-Object { $_.Name }) -join '、') }

    $ur = $ws.UsedRange
    $urRow = $ur.Row; $urCol = $ur.Column
    $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
    $vals = $ur.Value2                                  # 一次讀進 2D 陣列,避免逐格 COM 往返
    if ($rows -le 0 -or $null -eq $vals) { throw "工作表是空的。" }
    # 單一儲存格時 Value2 不是陣列,包成陣列統一處理
    $single = $false
    if ($vals -isnot [System.Array]) { $single = $true }
    function CellText([int]$r, [int]$c) {               # r,c 為 1-based (UsedRange 內相對位置)
        if ($single) { if ($r -eq 1 -and $c -eq 1) { return [string]$vals } else { return '' } }
        $v = $vals[$r, $c]
        if ($null -eq $v) { return '' }
        return ([string]$v).Trim()
    }

    # ----------------------------------------------------------------------- #
    #  欄位偵測
    # ----------------------------------------------------------------------- #
    $hdrRel = $HeaderRow - $urRow + 1                   # 標題列在 UsedRange 內的相對列
    if ($IpColumn -le 0 -or $NameColumn -le 0) {
        if ($hdrRel -ge 1 -and $hdrRel -le $rows) {
            for ($c = 1; $c -le $cols; $c++) {
                $h = (CellText $hdrRel $c).ToLower()
                if (-not $h) { continue }
                if ($IpColumn -le 0 -and ($h -match '^ip$|ip ?address|ip位址|ip地址|^位址$')) { $IpColumn = $urCol + $c - 1 }
                elseif ($NameColumn -le 0 -and ($h -match '名稱|name|設備|裝置|主機|hostname')) { $NameColumn = $urCol + $c - 1 }
            }
        }
        # 標題認不出 IP 欄 -> 掃內容,IP 樣式最多的欄就是
        if ($IpColumn -le 0) {
            $best = 0; $bestCnt = 0
            for ($c = 1; $c -le $cols; $c++) {
                $cnt = 0
                for ($r = 1; $r -le [Math]::Min($rows, 200); $r++) {
                    if (Get-FirstIP (CellText $r $c)) { $cnt++ }
                }
                if ($cnt -gt $bestCnt) { $bestCnt = $cnt; $best = $urCol + $c - 1 }
            }
            if ($bestCnt -gt 0) { $IpColumn = $best }
        }
    }
    if ($IpColumn -le 0) { throw "認不出哪一欄是 IP。請用 -IpColumn 指定(A=1, B=2, ...)。" }
    if ($NameColumn -le 0) {
        $hdrs = @()
        if ($hdrRel -ge 1) { for ($c = 1; $c -le $cols; $c++) { $hdrs += ("{0}={1}" -f ($urCol + $c - 1), (CellText $hdrRel $c)) } }
        throw ("認不出哪一欄是設備名稱。請用 -NameColumn 指定(A=1, B=2, ...)。目前標題列:" + ($hdrs -join ' | '))
    }
    Write-Host ("欄位:IP=第 {0} 欄,名稱=第 {1} 欄{2}" -f $IpColumn, $NameColumn, $(if ($RoleColumn -gt 0) { ",角色=第 $RoleColumn 欄" } else { "" }))

    # ----------------------------------------------------------------------- #
    #  建 IP -> 實際列號 對照(資料列從 HeaderRow+1 開始)
    # ----------------------------------------------------------------------- #
    $ipRel = $IpColumn - $urCol + 1
    $map = @{}
    $lastDataRow = $HeaderRow
    for ($r = 1; $r -le $rows; $r++) {
        $actualRow = $urRow + $r - 1
        if ($actualRow -le $HeaderRow) { continue }
        $ipCell = Get-FirstIP (CellText $r $ipRel)
        $rowHasData = $false
        for ($c = 1; $c -le $cols; $c++) { if (CellText $r $c) { $rowHasData = $true; break } }
        if ($rowHasData -and $actualRow -gt $lastDataRow) { $lastDataRow = $actualRow }
        if ($ipCell -and -not $map.ContainsKey($ipCell)) { $map[$ipCell] = $actualRow }
    }
    Write-Host ("資產表:{0} 個既有 IP,資料至第 {1} 列" -f $map.Count, $lastDataRow)

    # ----------------------------------------------------------------------- #
    #  合併(所有寫入一律紅字 ColorIndex=3)
    # ----------------------------------------------------------------------- #
    $RED = 3
    $filled = 0; $added = 0; $same = 0; $overwritten = 0
    $conflicts = New-Object System.Collections.Generic.List[string]
    $nameRel = $NameColumn - $urCol + 1

    foreach ($e in $fw) {
        if ($map.ContainsKey($e.ip)) {
            $row = $map[$e.ip]
            $rel = $row - $urRow + 1
            $cur = ''
            if ($rel -ge 1 -and $rel -le $rows) { $cur = CellText $rel $nameRel }
            if (-not $cur) {
                $cell = $ws.Cells.Item($row, $NameColumn)
                $cell.Value2 = $e.name
                $cell.Font.ColorIndex = $RED
                if ($RoleColumn -gt 0 -and $e.role) {
                    $rc = $ws.Cells.Item($row, $RoleColumn)
                    if (-not ([string]$rc.Value2)) { $rc.Value2 = $e.role; $rc.Font.ColorIndex = $RED }
                }
                $filled++
            }
            elseif ($cur -eq $e.name) { $same++ }
            elseif ($Overwrite) {
                $cell = $ws.Cells.Item($row, $NameColumn)
                $cell.Value2 = $e.name
                $cell.Font.ColorIndex = $RED
                $overwritten++
            }
            else {
                $conflicts.Add(("第 {0} 列 {1}:資產表為『{2}』,防火牆為『{3}』" -f $row, $e.ip, $cur, $e.name))
            }
        }
        else {
            $lastDataRow++
            $c1 = $ws.Cells.Item($lastDataRow, $IpColumn)
            $c1.Value2 = $e.ip; $c1.Font.ColorIndex = $RED
            $c2 = $ws.Cells.Item($lastDataRow, $NameColumn)
            $c2.Value2 = $e.name; $c2.Font.ColorIndex = $RED
            if ($RoleColumn -gt 0 -and $e.role) {
                $c3 = $ws.Cells.Item($lastDataRow, $RoleColumn)
                $c3.Value2 = $e.role; $c3.Font.ColorIndex = $RED
            }
            $added++
        }
    }

    # ----------------------------------------------------------------------- #
    #  存檔:預設另存新檔,不動原檔
    # ----------------------------------------------------------------------- #
    if ($InPlace) {
        $wb.Save()
        $saved = $ExcelFull
    }
    else {
        $dir = [System.IO.Path]::GetDirectoryName($ExcelFull)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFull)
        $ext = [System.IO.Path]::GetExtension($ExcelFull).ToLower()
        $fmt = 51                                        # .xlsx
        if ($ext -eq '.xlsm') { $fmt = 52 }
        elseif ($ext -eq '.xls') { $fmt = 56 }
        $saved = Join-Path $dir ($base + '-合併' + $ext)
        $wb.SaveAs($saved, $fmt)
    }

    Write-Host ""
    Write-Host ("完成:補上名稱 {0} 筆、新增 {1} 列、原本相同 {2} 筆{3} — 更動皆為紅字" -f `
        $filled, $added, $same, $(if ($Overwrite) { "、覆蓋 $overwritten 筆" } else { "" })) -ForegroundColor Green
    Write-Host ("已存檔:{0}" -f $saved) -ForegroundColor Green
    if ($conflicts.Count -gt 0) {
        Write-Host ""
        Write-Host ("名稱不同、未覆蓋 {0} 筆(加 -Overwrite 可覆蓋成防火牆名稱):" -f $conflicts.Count) -ForegroundColor Yellow
        foreach ($cf in $conflicts) { Write-Host ("  " + $cf) }
    }
}
catch {
    Write-Error ("合併失敗:" + $_.Exception.Message)
    exit 2
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    foreach ($o in @($ws, $wb, $xl)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
