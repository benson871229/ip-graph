#requires -Version 5.1
<#
.SYNOPSIS
    把防火牆擷取的「ip,名稱,角色」清單合併進 Excel 資產表(可跨多個工作表):
    同時爬所有工作表判斷 IP 是否已存在;已存在就在「它所在的那個表」補上設備名稱,
    完全找不到的 IP 才在指定的工作表(預設第 2 個)新增。所有更動一律標成紅字。

.DESCRIPTION
    純 PowerShell 5.1 + Excel COM(機器上有裝 Excel 即可,零安裝、離線可用)。

    合併規則:
      * 任一工作表已有該 IP、名稱欄空白 -> 在該表補上名稱(紅字)
      * 任一工作表已有該 IP、名稱相同   -> 不動
      * 任一工作表已有該 IP、名稱不同   -> 預設不覆蓋、列入報告;-Overwrite 才覆蓋(紅字)
      * 所有工作表都沒有該 IP           -> 在 -AddToSheet(預設第 2 個)最後新增一列(紅字)

    每個工作表的欄位(IP / 名稱 / 角色)獨立自動偵測,所以三個分頁排版不同也沒關係。

    安全:預設不動原檔,另存 <原檔名>-合併.xlsx;要直接改原檔請加 -InPlace(建議先備份)。
    執行前請先關閉該 Excel 檔。

.PARAMETER FirewallCsv
    防火牆擷取的清單(get-asset.ps1 / get-asset-regex.ps1 的輸出),每行「ip,名稱[,角色]」。

.PARAMETER Excel
    你的資產表 .xlsx / .xlsm / .xls。

.PARAMETER Sheets
    要「搜尋」的工作表(名稱或序號,可多個)。預設 = 全部工作表。

.PARAMETER AddToSheet
    找不到的 IP 要新增到哪個工作表(名稱或序號)。預設 = 2(第 2 個工作表)。

.PARAMETER IpColumn / NameColumn / RoleColumn
    只套用在「新增用的工作表(-AddToSheet)」的欄位覆蓋(A=1, B=2...)。
    其餘工作表一律自動偵測。通常不需指定。

.PARAMETER HeaderRow
    標題列列號,預設 1(對所有工作表一致)。資料從下一列起算。

.PARAMETER Overwrite
    已有名稱但與防火牆不同時,覆蓋成防火牆名稱(紅字)。預設不覆蓋。

.PARAMETER InPlace
    直接存回原檔。預設另存 <原檔名>-合併.xlsx。

.EXAMPLE
    .\merge-assets.ps1 -FirewallCsv assets.csv -Excel 資產清冊.xlsx
    # 爬全部工作表,找不到的加到第 2 個工作表

.EXAMPLE
    .\merge-assets.ps1 -FirewallCsv assets.csv -Excel 資產清冊.xlsx `
        -AddToSheet "待確認" -Overwrite

.NOTES
    需要本機安裝 Excel(COM 自動化,不裝任何套件)。
    本檔需存成 UTF-8 with BOM,否則 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$FirewallCsv,
    [Parameter(Mandatory = $true, Position = 1)]
    [string]$Excel,
    [object[]]$Sheets = @(),
    [object]$AddToSheet = 2,
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

function Get-FirstIP([string]$s) {
    if (-not $s) { return $null }
    if ($s -match '(?<![\d.])(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?![\d.])') {
        if ([int]$Matches[1] -le 255 -and [int]$Matches[2] -le 255 -and
            [int]$Matches[3] -le 255 -and [int]$Matches[4] -le 255) { return $Matches[0] }
    }
    return $null
}

# --------------------------------------------------------------------------- #
#  讀防火牆清單
# --------------------------------------------------------------------------- #
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
    if ($fwSeen.Contains($ip)) { continue }
    [void]$fwSeen.Add($ip)
    $fw.Add([pscustomobject]@{ ip = $ip; name = $name; role = $role })
}
if ($fw.Count -eq 0) { Write-Error "防火牆清單沒有解析到任何「ip,名稱」資料。"; exit 1 }
Write-Host ("防火牆清單:{0} 筆不重複 IP" -f $fw.Count)

# --------------------------------------------------------------------------- #
#  取一個工作表的內容並偵測欄位,回傳 context
# --------------------------------------------------------------------------- #
function Cell($ctx, [int]$r, [int]$c) {
    if ($ctx.single) { if ($r -eq 1 -and $c -eq 1) { return [string]$ctx.vals } else { return '' } }
    if ($r -lt 1 -or $r -gt $ctx.rows -or $c -lt 1 -or $c -gt $ctx.cols) { return '' }
    $v = $ctx.vals[$r, $c]
    if ($null -eq $v) { return '' }
    return ([string]$v).Trim()
}
function New-SheetCtx($ws, [int]$ipOv, [int]$nameOv, [int]$roleOv) {
    $ur = $ws.UsedRange
    $ctx = @{
        Ws = $ws; Name = $ws.Name
        urRow = $ur.Row; urCol = $ur.Column
        rows = $ur.Rows.Count; cols = $ur.Columns.Count
        vals = $ur.Value2; single = $false
        IpCol = 0; NameCol = 0; RoleCol = 0; LastDataRow = $HeaderRow
    }
    if ($ctx.vals -isnot [System.Array]) { $ctx.single = $true }

    $hdrRel = $HeaderRow - $ctx.urRow + 1
    $ipCol = $ipOv; $nameCol = $nameOv; $roleCol = $roleOv
    if ($ipCol -le 0 -or $nameCol -le 0 -or $roleCol -le 0) {
        if ($hdrRel -ge 1 -and $hdrRel -le $ctx.rows) {
            for ($c = 1; $c -le $ctx.cols; $c++) {
                $h = (Cell $ctx $hdrRel $c).ToLower()
                if (-not $h) { continue }
                if ($ipCol -le 0 -and ($h -match '^ip$|ip ?address|ip位址|ip地址|^位址$')) { $ipCol = $ctx.urCol + $c - 1 }
                elseif ($nameCol -le 0 -and ($h -match '名稱|name|設備|裝置|主機|hostname')) { $nameCol = $ctx.urCol + $c - 1 }
                elseif ($roleCol -le 0 -and ($h -match '角色|role|用途|類型|type|說明|備註')) { $roleCol = $ctx.urCol + $c - 1 }
            }
        }
        if ($ipCol -le 0) {
            $best = 0; $bestCnt = 0
            for ($c = 1; $c -le $ctx.cols; $c++) {
                $cnt = 0
                for ($r = 1; $r -le [Math]::Min($ctx.rows, 200); $r++) { if (Get-FirstIP (Cell $ctx $r $c)) { $cnt++ } }
                if ($cnt -gt $bestCnt) { $bestCnt = $cnt; $best = $ctx.urCol + $c - 1 }
            }
            if ($bestCnt -gt 0) { $ipCol = $best }
        }
    }
    $ctx.IpCol = $ipCol; $ctx.NameCol = $nameCol; $ctx.RoleCol = $roleCol

    # 找最後一筆有資料的列(新增時接在後面)
    for ($r = 1; $r -le $ctx.rows; $r++) {
        $actual = $ctx.urRow + $r - 1
        if ($actual -le $HeaderRow) { continue }
        for ($c = 1; $c -le $ctx.cols; $c++) {
            if (Cell $ctx $r $c) { if ($actual -gt $ctx.LastDataRow) { $ctx.LastDataRow = $actual }; break }
        }
    }
    return $ctx
}

# --------------------------------------------------------------------------- #
#  Excel COM
# --------------------------------------------------------------------------- #
$xl = $null; $wb = $null
$sheetObjs = New-Object System.Collections.Generic.List[object]
try {
    $xl = New-Object -ComObject Excel.Application
} catch {
    Write-Error "無法啟動 Excel COM — 這台機器需要安裝 Excel 才能執行本工具。"
    exit 1
}
$xl.Visible = $false; $xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($ExcelFull)
    # 用 index 迴圈可靠地取工作表名稱(管線列舉 COM 集合在 5.1 常列不出東西)
    $wsCount = $wb.Worksheets.Count
    $wsNames = @()
    for ($i = 1; $i -le $wsCount; $i++) { $wsNames += [string]$wb.Worksheets.Item($i).Name }
    Write-Host ("活頁簿有 {0} 個工作表:{1}" -f $wsCount, ($wsNames -join '、'))

    # 明確區分「數字→索引」與「文字→名稱」;不靠 COM 猜(命令列傳入的 2 會是字串 "2",
    # 直接丟給 Item 會被當成「名叫 2 的工作表」而找不到)
    function Resolve-Sheet($id) {
        $asInt = 0
        if ($id -is [int]) { $asInt = $id }
        elseif ("$id" -match '^\s*\d+\s*$') { $asInt = [int]("$id".Trim()) }
        if ($asInt -ge 1) {
            if ($asInt -le $wsCount) { return $wb.Worksheets.Item($asInt) }
            throw "工作表索引 $asInt 超出範圍(這本共 $wsCount 個)。工作表:" + ($wsNames -join '、')
        }
        $ids = ("$id").Trim()
        for ($i = 1; $i -le $wsCount; $i++) {
            $nm = [string]$wb.Worksheets.Item($i).Name
            if ($nm -eq "$id" -or $nm.Trim() -eq $ids -or $nm.Trim().ToLower() -eq $ids.ToLower()) {
                return $wb.Worksheets.Item($i)
            }
        }
        try { return $wb.Worksheets.Item("$id") } catch {}   # 最後容錯:名稱含特殊字元時直接問 Excel
        throw "找不到工作表『$id』。這本活頁簿的工作表:" + ($wsNames -join '、')
    }

    # 要搜尋的工作表(預設全部);確保新增用的工作表也在搜尋範圍內
    $searchWs = @()
    if ($Sheets.Count -gt 0) { foreach ($s in $Sheets) { $searchWs += (Resolve-Sheet $s) } }
    else { for ($i = 1; $i -le $wb.Worksheets.Count; $i++) { $searchWs += $wb.Worksheets.Item($i) } }

    $addWs = Resolve-Sheet $AddToSheet

    # 建立每個搜尋表的 context,並組全域 IP -> 位置對照
    $globalMap = @{}       # ip -> @{ Ctx; Row }
    $skipSheets = New-Object System.Collections.Generic.List[string]
    foreach ($ws in $searchWs) {
        $ctx = New-SheetCtx $ws 0 0 0
        $sheetObjs.Add($ctx.Ws) | Out-Null
        if ($ctx.IpCol -le 0) { $skipSheets.Add($ctx.Name) | Out-Null; continue }
        $ipRel = $ctx.IpCol - $ctx.urCol + 1
        $found = 0
        for ($r = 1; $r -le $ctx.rows; $r++) {
            $actual = $ctx.urRow + $r - 1
            if ($actual -le $HeaderRow) { continue }
            $ipCell = Get-FirstIP (Cell $ctx $r $ipRel)
            if ($ipCell -and -not $globalMap.ContainsKey($ipCell)) {
                $globalMap[$ipCell] = @{ Ctx = $ctx; Row = $actual }; $found++
            }
        }
        Write-Host ("工作表『{0}』:IP 欄第 {1}{2},既有 {3} 個 IP" -f `
            $ctx.Name, $ctx.IpCol, $(if ($ctx.NameCol -gt 0) { "、名稱欄第 $($ctx.NameCol)" } else { "、無名稱欄" }), $found)
    }
    if ($skipSheets.Count -gt 0) {
        Write-Host ("(略過無法辨識 IP 欄的工作表:{0})" -f ($skipSheets -join '、')) -ForegroundColor DarkYellow
    }

    # 新增用工作表的 context(可用 -IpColumn/-NameColumn/-RoleColumn 覆蓋)
    $addCtx = New-SheetCtx $addWs $IpColumn $NameColumn $RoleColumn
    if ($addCtx.IpCol -le 0 -or $addCtx.NameCol -le 0) {
        $hdrs = @()
        $hrel = $HeaderRow - $addCtx.urRow + 1
        if ($hrel -ge 1) { for ($c = 1; $c -le $addCtx.cols; $c++) { $hdrs += ("{0}={1}" -f ($addCtx.urCol + $c - 1), (Cell $addCtx $hrel $c)) } }
        throw ("新增用的工作表『{0}』認不出 IP 或名稱欄,無法新增。請用 -IpColumn / -NameColumn 指定(A=1, B=2...)。目前標題列:{1}" -f $addCtx.Name, ($hdrs -join ' | '))
    }
    Write-Host ("新增目標:工作表『{0}』(IP 欄第 {1}、名稱欄第 {2}),接在第 {3} 列後新增" -f `
        $addCtx.Name, $addCtx.IpCol, $addCtx.NameCol, $addCtx.LastDataRow)

    # ----------------------------------------------------------------------- #
    #  合併(紅字 ColorIndex=3)
    # ----------------------------------------------------------------------- #
    $RED = 3
    $filled = 0; $added = 0; $same = 0; $overwritten = 0
    $conflicts = New-Object System.Collections.Generic.List[string]
    $noNameCol = New-Object System.Collections.Generic.List[string]

    foreach ($e in $fw) {
        if ($globalMap.ContainsKey($e.ip)) {
            $loc = $globalMap[$e.ip]; $ctx = $loc.Ctx; $row = $loc.Row
            if ($ctx.NameCol -le 0) { $noNameCol.Add(("{0}(在『{1}』)" -f $e.ip, $ctx.Name)) | Out-Null; continue }
            $rel = $row - $ctx.urRow + 1
            $nameRel = $ctx.NameCol - $ctx.urCol + 1
            $cur = Cell $ctx $rel $nameRel
            if (-not $cur) {
                $cell = $ctx.Ws.Cells.Item($row, $ctx.NameCol)
                $cell.Value2 = $e.name; $cell.Font.ColorIndex = $RED
                if ($ctx.RoleCol -gt 0 -and $e.role) {
                    $roleRel = $ctx.RoleCol - $ctx.urCol + 1
                    if (-not (Cell $ctx $rel $roleRel)) {
                        $rc = $ctx.Ws.Cells.Item($row, $ctx.RoleCol); $rc.Value2 = $e.role; $rc.Font.ColorIndex = $RED
                    }
                }
                $filled++
            }
            elseif ($cur -eq $e.name) { $same++ }
            elseif ($Overwrite) {
                $cell = $ctx.Ws.Cells.Item($row, $ctx.NameCol)
                $cell.Value2 = $e.name; $cell.Font.ColorIndex = $RED; $overwritten++
            }
            else {
                $conflicts.Add(("『{0}』第 {1} 列 {2}:表為『{3}』,防火牆為『{4}』" -f $ctx.Name, $row, $e.ip, $cur, $e.name)) | Out-Null
            }
        }
        else {
            $addCtx.LastDataRow++
            $r = $addCtx.LastDataRow
            $c1 = $addWs.Cells.Item($r, $addCtx.IpCol);   $c1.Value2 = $e.ip;   $c1.Font.ColorIndex = $RED
            $c2 = $addWs.Cells.Item($r, $addCtx.NameCol); $c2.Value2 = $e.name; $c2.Font.ColorIndex = $RED
            if ($addCtx.RoleCol -gt 0 -and $e.role) {
                $c3 = $addWs.Cells.Item($r, $addCtx.RoleCol); $c3.Value2 = $e.role; $c3.Font.ColorIndex = $RED
            }
            $added++
        }
    }

    # ----------------------------------------------------------------------- #
    #  存檔
    # ----------------------------------------------------------------------- #
    if ($InPlace) { $wb.Save(); $saved = $ExcelFull }
    else {
        $dir = [System.IO.Path]::GetDirectoryName($ExcelFull)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFull)
        $ext = [System.IO.Path]::GetExtension($ExcelFull).ToLower()
        $fmt = 51; if ($ext -eq '.xlsm') { $fmt = 52 } elseif ($ext -eq '.xls') { $fmt = 56 }
        $saved = Join-Path $dir ($base + '-合併' + $ext)
        $wb.SaveAs($saved, $fmt)
    }

    Write-Host ""
    Write-Host ("完成:補上名稱 {0} 筆、於『{1}』新增 {2} 列、原本相同 {3} 筆{4} — 更動皆為紅字" -f `
        $filled, $addCtx.Name, $added, $same, $(if ($Overwrite) { "、覆蓋 $overwritten 筆" } else { "" })) -ForegroundColor Green
    Write-Host ("已存檔:{0}" -f $saved) -ForegroundColor Green
    if ($conflicts.Count -gt 0) {
        Write-Host ""
        Write-Host ("名稱不同、未覆蓋 {0} 筆(加 -Overwrite 可覆蓋):" -f $conflicts.Count) -ForegroundColor Yellow
        foreach ($cf in $conflicts) { Write-Host ("  " + $cf) }
    }
    if ($noNameCol.Count -gt 0) {
        Write-Host ""
        Write-Host ("IP 已存在但該表無名稱欄、無法補名 {0} 筆:" -f $noNameCol.Count) -ForegroundColor Yellow
        foreach ($nn in $noNameCol) { Write-Host ("  " + $nn) }
    }
}
catch {
    Write-Error ("合併失敗:" + $_.Exception.Message)
    exit 2
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    foreach ($o in $sheetObjs) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
    foreach ($o in @($wb, $xl)) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
