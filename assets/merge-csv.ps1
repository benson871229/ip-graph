#requires -Version 5.1
<#
.SYNOPSIS
    把多份「欄位相同」的 CSV 合併成一份,並多加一欄記錄每列來自哪個檔案。

.DESCRIPTION
    純文字處理,不碰 Excel COM(所以不會踩到 COM 跨行程那些坑),
    Windows PowerShell 5.1 內建就能跑,零依賴、離線可用。

    合併規則:
      * 只保留第一份的表頭,其餘檔案的表頭列跳過
      * 同一個 IP 出現在多份檔案 -> 全部保留,不去重(這是刻意的)
      * 多加一欄「來源檔」,記錄每一列來自哪個檔案
      * 欄位以「欄名」對應,所以就算某份檔案欄序不同也接得起來;
        某份檔案多出來的欄會自動加在最後面,缺的欄留空,並在結果裡提醒你

    吃得下的輸入:
      UTF-8(有/沒有 BOM)、UTF-16 LE/BE、Big5;CRLF / LF / CR 混用;
      欄位含逗號、雙引號("" 逃逸)、含換行的引號欄位

    輸出一律 UTF-8 with BOM —— 沒有 BOM 的話 Excel 打開會中文亂碼。

    合併完要匯進 Excel 資產清冊,用 merge-csv.bas(VBA 巨集)比較省事:
    它可以直接選多個 CSV -> 合併 -> 附加到工作表,不用先產生中間檔。

.PARAMETER Path
    要合併的 CSV。可以給多個檔名、萬用字元(*.csv),或直接給一個資料夾。

.PARAMETER OutFile
    輸出檔名。預設 = 第一個來源檔所在資料夾底下的「合併-yyyyMMdd-HHmmss.csv」。

.PARAMETER SourceColumn
    來源檔欄位的名稱,預設「來源檔」。若與既有欄名撞名會自動改成「來源檔2」。

.PARAMETER NoSourceColumn
    不要加來源檔欄位。

.PARAMETER FullPath
    來源檔欄位記完整路徑(預設只記檔名)。

.EXAMPLE
    .\merge-csv.ps1 .\匯出\*.csv
    # 合併該資料夾所有 CSV,輸出到同一個資料夾

.EXAMPLE
    .\merge-csv.ps1 fw1.csv fw2.csv fw3.csv fw4.csv fw5.csv -OutFile 合併.csv

.EXAMPLE
    .\merge-csv.ps1 D:\匯出 -OutFile D:\合併.csv
    # 給資料夾 = 該資料夾下所有 *.csv

.NOTES
    本檔需存成「UTF-8 with BOM」,否則 PowerShell 5.1 會用 Big5 解讀,繁中變亂碼。
#>
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,
    [string]$OutFile,
    [string]$SourceColumn = '來源檔',
    [switch]$NoSourceColumn,
    [switch]$FullPath
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# --------------------------------------------------------------------------- #
#  編碼自動偵測
#  PowerShell 5.1 的 Get-Content -Encoding UTF8 不會認 UTF-16,這裡自己讀 bytes
#  判斷 BOM;無 BOM 時用 null byte 分布猜 UTF-16,再退回 UTF-8 / Big5。
# --------------------------------------------------------------------------- #
function Read-TextAuto([string]$p, [ref]$encName) {
    $bytes = [System.IO.File]::ReadAllBytes($p)
    if ($bytes.Length -eq 0) { $encName.Value = '空檔'; return '' }

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encName.Value = 'UTF-8 BOM'
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encName.Value = 'UTF-16 LE'
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encName.Value = 'UTF-16 BE'
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }

    $lim = [Math]::Min($bytes.Length, 4000)
    $evenNull = 0; $oddNull = 0
    for ($i = 0; $i -lt $lim; $i++) {
        if ($bytes[$i] -eq 0) {
            if ($i % 2 -eq 0) { $evenNull++ } else { $oddNull++ }
        }
    }
    if (($evenNull + $oddNull) -gt ($lim / 4)) {
        if ($oddNull -gt $evenNull) {
            $encName.Value = 'UTF-16 LE (無 BOM,推測)'
            return [System.Text.Encoding]::Unicode.GetString($bytes)
        }
        $encName.Value = 'UTF-16 BE (無 BOM,推測)'
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
    }

    $utf8 = New-Object System.Text.UTF8Encoding($false, $false)
    $text = $utf8.GetString($bytes)
    if ($text.IndexOf([char]0xFFFD) -ge 0) {
        $encName.Value = 'ANSI/Big5 (UTF-8 解碼失敗後退回)'
        return [System.Text.Encoding]::Default.GetString($bytes)
    }
    $encName.Value = 'UTF-8 (無 BOM)'
    return $text
}

# --------------------------------------------------------------------------- #
#  CSV 解析(狀態機)
#    - 只有出現在「欄位開頭」的雙引號才算引號欄位
#    - 引號欄位內的 "" = 一個雙引號字元,也可以直接包含換行
#    - CRLF / LF / CR 都當換行
#    - 全空白的列直接略過
#  回傳:List[object],每個元素是一個 string[](一列)
# --------------------------------------------------------------------------- #
function ConvertFrom-CsvText([string]$s) {
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $s) { return , $rows }
    if ($s[0] -eq [char]0xFEFF) { $s = $s.Substring(1) }   # 保險:再剝一次 BOM 字元
    if ($s.Length -eq 0) { return , $rows }

    $fields = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQ = $false
    $n = $s.Length
    $i = 0

    # 收一列(全空白就丟掉)
    $flush = {
        $arr = $fields.ToArray()
        $blank = $true
        foreach ($f in $arr) { if ($f.Trim().Length -gt 0) { $blank = $false; break } }
        if (-not $blank) { $rows.Add($arr) }
        $fields.Clear()
    }

    while ($i -lt $n) {
        $ch = $s[$i]
        if ($inQ) {
            if ($ch -eq '"') {
                $nxt = ''
                if ($i -lt ($n - 1)) { $nxt = $s[$i + 1] }
                if ($nxt -eq '"') {
                    [void]$sb.Append('"')
                    $i += 2
                }
                else {
                    $inQ = $false
                    $i++
                }
            }
            else {
                [void]$sb.Append($ch)
                $i++
            }
        }
        else {
            if ($ch -eq '"') {
                if ($sb.Length -eq 0) { $inQ = $true } else { [void]$sb.Append($ch) }
                $i++
            }
            elseif ($ch -eq ',') {
                $fields.Add($sb.ToString()); [void]$sb.Clear()
                $i++
            }
            elseif ($ch -eq "`r" -or $ch -eq "`n") {
                $fields.Add($sb.ToString()); [void]$sb.Clear()
                & $flush
                if ($ch -eq "`r" -and $i -lt ($n - 1) -and $s[$i + 1] -eq "`n") { $i++ }
                $i++
            }
            else {
                [void]$sb.Append($ch)
                $i++
            }
        }
    }
    if ($sb.Length -gt 0 -or $fields.Count -gt 0) {
        $fields.Add($sb.ToString()); [void]$sb.Clear()
        & $flush
    }
    # 逗號運算子:避免 5.1 把單列的 List 展開成純量
    return , $rows
}

# 欄名正規化:去空白、去底線/連字號、轉小寫
function Get-NormHeader([string]$s) {
    if ($null -eq $s) { return '' }
    $t = $s.Trim().ToLower()
    $t = $t -replace '[\s_\-　]', ''
    return $t
}

# 輸出 CSV 時的跳脫:含逗號/引號/換行/前後空白就加引號
function ConvertTo-CsvField([string]$s) {
    if ($null -eq $s) { return '' }
    if ($s.IndexOf(',') -ge 0 -or $s.IndexOf('"') -ge 0 -or
        $s.IndexOf("`r") -ge 0 -or $s.IndexOf("`n") -ge 0 -or $s -ne $s.Trim()) {
        return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}

# --------------------------------------------------------------------------- #
#  展開輸入路徑(檔名 / 萬用字元 / 資料夾),依檔名排序讓合併順序固定
# --------------------------------------------------------------------------- #
$inputFiles = New-Object System.Collections.Generic.List[string]
foreach ($p in $Path) {
    if (Test-Path -LiteralPath $p -PathType Container) {
        $found = @(Get-ChildItem -LiteralPath $p -Filter *.csv -File | Sort-Object Name)
        foreach ($f in $found) { $inputFiles.Add($f.FullName) }
        continue
    }
    $found = @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($found.Count -eq 0) {
        Write-Warning "找不到檔案:$p"
        continue
    }
    foreach ($f in $found) { $inputFiles.Add($f.FullName) }
}
# 去掉重複給到的同一個檔(保持順序)
$seenPath = New-Object 'System.Collections.Generic.HashSet[string]'
$files = New-Object System.Collections.Generic.List[string]
foreach ($f in $inputFiles) {
    if ($seenPath.Add($f.ToLower())) { $files.Add($f) }
}
if ($files.Count -eq 0) { Write-Error "沒有可合併的檔案。"; exit 1 }

# --------------------------------------------------------------------------- #
#  合併
# --------------------------------------------------------------------------- #
$header = New-Object System.Collections.Generic.List[string]   # 合併後的欄名(不含來源檔欄)
$data = New-Object System.Collections.Generic.List[object]     # 每列 = string[]
$srcs = New-Object System.Collections.Generic.List[string]     # 每列的來源檔
$baseKeys = $null
$mismatch = New-Object System.Collections.Generic.List[string]
$report = New-Object System.Collections.Generic.List[string]

foreach ($file in $files) {
    $enc = ''
    $text = Read-TextAuto $file ([ref]$enc)
    $rows = ConvertFrom-CsvText $text
    $base = [System.IO.Path]::GetFileName($file)
    if ($rows.Count -eq 0) {
        $report.Add(("  {0} ({1}):空檔,已略過" -f $base, $enc))
        continue
    }

    # 表頭 -> 合併欄位的對應(以欄名對應,欄序不同也接得起來)
    $fh = $rows[0]
    $map = New-Object 'System.Collections.Generic.List[int]'
    $usedIdx = New-Object 'System.Collections.Generic.HashSet[int]'
    $keys = New-Object System.Collections.Generic.List[string]
    for ($j = 0; $j -lt $fh.Count; $j++) {
        $hn = ([string]$fh[$j]).Trim()
        if ($hn.Length -eq 0) { $hn = "(未命名欄$($j + 1))" }
        $keys.Add((Get-NormHeader $hn))
        $found = -1
        for ($k = 0; $k -lt $header.Count; $k++) {
            if (-not $usedIdx.Contains($k) -and (Get-NormHeader $header[$k]) -eq (Get-NormHeader $hn)) {
                $found = $k; break
            }
        }
        if ($found -lt 0) {
            $header.Add($hn)
            $found = $header.Count - 1
        }
        [void]$usedIdx.Add($found)
        $map.Add($found)
    }

    $keyStr = ($keys -join '|')
    if ($null -eq $baseKeys) { $baseKeys = $keyStr }
    elseif ($keyStr -ne $baseKeys) { $mismatch.Add($base) }

    $tag = $base
    if ($FullPath) { $tag = $file }

    $cnt = 0
    for ($r = 1; $r -lt $rows.Count; $r++) {
        $src = $rows[$r]
        $out = New-Object 'string[]' $header.Count
        for ($j = 0; $j -lt $out.Length; $j++) { $out[$j] = '' }
        for ($j = 0; $j -lt $map.Count; $j++) {
            if ($j -lt $src.Count) { $out[$map[$j]] = [string]$src[$j] }
        }
        $data.Add($out)
        $srcs.Add($tag)
        $cnt++
    }
    $report.Add(("  {0} ({1}):{2} 列" -f $base, $enc, $cnt))
}

if ($header.Count -eq 0) { Write-Error "所有檔案都讀不到表頭,請確認第一列是欄位名稱。"; exit 1 }
if ($data.Count -eq 0) { Write-Error "合併後沒有任何資料列。"; exit 1 }

# 來源檔欄名不能跟既有欄名撞名
$srcName = ''
if (-not $NoSourceColumn) {
    $srcName = $SourceColumn
    $k = 2
    while ($true) {
        $hit = $false
        foreach ($h in $header) { if ((Get-NormHeader $h) -eq (Get-NormHeader $srcName)) { $hit = $true; break } }
        if (-not $hit) { break }
        $srcName = "$SourceColumn$k"
        $k++
        if ($k -gt 20) { break }
    }
}

# --------------------------------------------------------------------------- #
#  輸出(UTF-8 with BOM,Excel 直接打開才不會中文亂碼)
# --------------------------------------------------------------------------- #
if (-not $OutFile) {
    $dir = [System.IO.Path]::GetDirectoryName($files[0])
    $OutFile = Join-Path $dir ('合併-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.csv')
}

$sb = New-Object System.Text.StringBuilder
$hdrOut = New-Object System.Collections.Generic.List[string]
foreach ($h in $header) { $hdrOut.Add((ConvertTo-CsvField $h)) }
if ($srcName) { $hdrOut.Add((ConvertTo-CsvField $srcName)) }
[void]$sb.Append(($hdrOut -join ',')).Append("`r`n")

for ($r = 0; $r -lt $data.Count; $r++) {
    $row = $data[$r]
    $parts = New-Object 'string[]' ($header.Count + $(if ($srcName) { 1 } else { 0 }))
    for ($j = 0; $j -lt $header.Count; $j++) {
        $v = ''
        if ($j -lt $row.Length) { $v = $row[$j] }
        $parts[$j] = ConvertTo-CsvField $v
    }
    if ($srcName) { $parts[$header.Count] = ConvertTo-CsvField $srcs[$r] }
    [void]$sb.Append(($parts -join ',')).Append("`r`n")
}

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $utf8Bom)

# --------------------------------------------------------------------------- #
#  報告
# --------------------------------------------------------------------------- #
Write-Host ("讀入 {0} 個檔案:" -f $files.Count)
foreach ($line in $report) { Write-Host $line }
Write-Host ""
if ($mismatch.Count -gt 0) {
    Write-Host "注意:下列檔案的欄位與第一個檔案不完全相同,已用「欄名」對應(缺的欄留空、多的欄加在後面):" -ForegroundColor Yellow
    foreach ($m in $mismatch) { Write-Host ("  " + $m) }
    Write-Host ""
}
$colDesc = ($header -join '、')
if ($srcName) { $colDesc = $colDesc + '、' + $srcName }
Write-Host ("合併後:{0} 列 / {1} 欄(同 IP 不去重,全部保留)" -f $data.Count, $hdrOut.Count) -ForegroundColor Green
Write-Host ("欄位:{0}" -f $colDesc)
Write-Host ("已輸出(UTF-8 BOM):{0}" -f $OutFile) -ForegroundColor Green
Write-Host ""
Write-Host "下一步:要併進 Excel 資產清冊的話,用 assets\merge-csv.bas 巨集(可直接多選 CSV,不必先產生這個中間檔)。"
