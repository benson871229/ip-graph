Attribute VB_Name = "MergeAssets"
Option Explicit

' =====================================================================
'  合併資產表(Excel VBA 版)
'  把防火牆擷取的「ip,名稱[,角色]」清單合併進目前開啟的資產表活頁簿:
'    - 同時爬「所有」工作表判斷 IP 是否已存在
'    - 已存在且名稱空白 -> 在它所在的那個表補上名稱(紅字)
'    - 都找不到的 IP    -> 在 ADD_SHEET 指定的工作表底部新增(紅字)
'  所有更動一律紅字。巨集「不會自動存檔」,你檢查完紅字再自己另存。
'
'  用法:
'    1. 開啟你的資產表 .xlsx
'    2. Alt+F11 開 VBA 編輯器 -> 插入(Insert) -> 模組(Module)
'    3. 把這整段貼進去
'    4. 改下面 ADD_SHEET 成你要新增的工作表(名稱或索引數字皆可)
'    5. 按 F5 執行,選防火牆清單 CSV
' =====================================================================

' ====== 設定(依你的表調整這幾行)======
Const ADD_SHEET As String = "我看到的"   ' 找不到的 IP 加到哪個表:填名稱,或填索引數字如 "2"
Const HEADER_ROW As Long = 1              ' 標題列在第幾列(資料從下一列起算)
Const OVERWRITE As Boolean = False        ' True = 名稱不同時覆蓋成防火牆的名稱
Const CSV_CHARSET As String = "utf-8"     ' 防火牆清單編碼;若中文變亂碼改成 "big5"

Sub 合併資產表()
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    If wbk Is Nothing Then MsgBox "找不到開啟中的活頁簿。", vbExclamation: Exit Sub

    ' ---- 選防火牆清單檔 ----
    Dim fwPath As Variant
    fwPath = Application.GetOpenFilename( _
        "防火牆清單 (*.csv;*.txt),*.csv;*.txt,所有檔案 (*.*),*.*", , "選擇防火牆擷取的 ip,名稱 清單")
    If VarType(fwPath) = vbBoolean Then Exit Sub

    ' ---- 讀清單(用 ADODB.Stream 正確讀 UTF-8 中文)----
    Dim ips() As String, names() As String, roles() As String, n As Long
    ReDim ips(1 To 200000): ReDim names(1 To 200000): ReDim roles(1 To 200000): n = 0
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Dim stream As Object, content As String
    On Error GoTo readErr
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = CSV_CHARSET
    stream.Open
    stream.LoadFromFile CStr(fwPath)
    content = stream.ReadText(-1)
    stream.Close
    On Error GoTo 0

    Dim lines() As String, i As Long, ln As String, parts() As String, ip As String, nm As String, rl As String, k As Long
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 And Left(ln, 1) <> "#" Then
            parts = Split(ln, ",")
            If UBound(parts) >= 1 Then
                ip = ExtractIP(Trim(parts(0)))
                nm = Trim(parts(1))
                If Len(ip) > 0 And Len(nm) > 0 Then
                    If Not seen.Exists(ip) Then
                        seen.Add ip, True
                        rl = ""
                        If UBound(parts) >= 2 Then
                            rl = Trim(parts(2))
                            For k = 3 To UBound(parts): rl = rl & "," & parts(k): Next k
                        End If
                        n = n + 1: ips(n) = ip: names(n) = nm: roles(n) = rl
                    End If
                End If
            End If
        End If
    Next i
    If n = 0 Then MsgBox "防火牆清單沒有解析到任何 ip,名稱。", vbExclamation: Exit Sub

    ' ---- 每個工作表偵測欄位 + 建立全域 IP -> 位置對照 ----
    Dim wsCount As Long: wsCount = wbk.Worksheets.Count
    Dim ipCols() As Long, nameCols() As Long, roleCols() As Long
    ReDim ipCols(1 To wsCount): ReDim nameCols(1 To wsCount): ReDim roleCols(1 To wsCount)

    Dim loc As Object: Set loc = CreateObject("Scripting.Dictionary")   ' ip -> "wsIdx|row"
    Dim ws As Worksheet, w As Long, r As Long, lastR As Long, ic As Long, nc As Long, rc As Long, cellIP As String
    Dim skipped As String: skipped = ""

    For w = 1 To wsCount
        Set ws = wbk.Worksheets(w)
        DetectColumns ws, ic, nc, rc
        ipCols(w) = ic: nameCols(w) = nc: roleCols(w) = rc
        If ic = 0 Then
            skipped = skipped & "  " & ws.Name & vbCrLf
        Else
            lastR = SheetLastRow(ws)
            For r = HEADER_ROW + 1 To lastR
                cellIP = ExtractIP(CStr(ws.Cells(r, ic).Value))
                If Len(cellIP) > 0 Then
                    If Not loc.Exists(cellIP) Then loc.Add cellIP, w & "|" & r
                End If
            Next r
        End If
    Next w

    ' ---- 新增用工作表 ----
    Dim addWs As Worksheet
    Set addWs = ResolveSheet(wbk, ADD_SHEET)
    If addWs Is Nothing Then
        MsgBox "找不到要新增的工作表『" & ADD_SHEET & "』。" & vbCrLf & _
               "活頁簿的工作表:" & vbCrLf & ListSheets(wbk), vbExclamation
        Exit Sub
    End If
    Dim addIc As Long, addNc As Long, addRc As Long
    DetectColumns addWs, addIc, addNc, addRc
    If addIc = 0 Or addNc = 0 Then
        MsgBox "工作表『" & addWs.Name & "』認不出 IP 欄或名稱欄,無法新增。" & vbCrLf & _
               "請確認標題列(第 " & HEADER_ROW & " 列)有『ip』與『名稱/name/設備』之類的欄名。", vbExclamation
        Exit Sub
    End If
    Dim addRow As Long: addRow = SheetLastRow(addWs)

    ' ---- 合併(紅字)----
    Application.ScreenUpdating = False
    Dim filled As Long, added As Long, same As Long, overwritten As Long
    Dim conflicts As String, noNameCol As String
    Dim e As Long, cur As String, wsi As Long, rowi As Long, key As String, arr() As String

    For e = 1 To n
        If loc.Exists(ips(e)) Then
            arr = Split(loc(ips(e)), "|")
            wsi = CLng(arr(0)): rowi = CLng(arr(1))
            Set ws = wbk.Worksheets(wsi)
            nc = nameCols(wsi)
            If nc = 0 Then
                noNameCol = noNameCol & "  " & ips(e) & " (在『" & ws.Name & "』)" & vbCrLf
            Else
                cur = Trim(CStr(ws.Cells(rowi, nc).Value))
                If Len(cur) = 0 Then
                    ws.Cells(rowi, nc).Value = names(e)
                    ws.Cells(rowi, nc).Font.Color = vbRed
                    rc = roleCols(wsi)
                    If rc > 0 And Len(roles(e)) > 0 Then
                        If Len(Trim(CStr(ws.Cells(rowi, rc).Value))) = 0 Then
                            ws.Cells(rowi, rc).Value = roles(e)
                            ws.Cells(rowi, rc).Font.Color = vbRed
                        End If
                    End If
                    filled = filled + 1
                ElseIf cur = names(e) Then
                    same = same + 1
                ElseIf OVERWRITE Then
                    ws.Cells(rowi, nc).Value = names(e)
                    ws.Cells(rowi, nc).Font.Color = vbRed
                    overwritten = overwritten + 1
                Else
                    conflicts = conflicts & "  『" & ws.Name & "』第 " & rowi & " 列 " & ips(e) & _
                                ":表為『" & cur & "』,防火牆為『" & names(e) & "』" & vbCrLf
                End If
            End If
        Else
            addRow = addRow + 1
            addWs.Cells(addRow, addIc).Value = ips(e)
            addWs.Cells(addRow, addIc).Font.Color = vbRed
            addWs.Cells(addRow, addNc).Value = names(e)
            addWs.Cells(addRow, addNc).Font.Color = vbRed
            If addRc > 0 And Len(roles(e)) > 0 Then
                addWs.Cells(addRow, addRc).Value = roles(e)
                addWs.Cells(addRow, addRc).Font.Color = vbRed
            End If
            added = added + 1
        End If
    Next e
    Application.ScreenUpdating = True

    ' ---- 報告 ----
    Dim msg As String
    msg = "完成(更動皆為紅字,尚未存檔,請檢查後自行另存):" & vbCrLf & vbCrLf
    msg = msg & "補上名稱:" & filled & " 筆" & vbCrLf
    msg = msg & "於『" & addWs.Name & "』新增:" & added & " 列" & vbCrLf
    msg = msg & "原本就相同:" & same & " 筆" & vbCrLf
    If OVERWRITE Then msg = msg & "覆蓋:" & overwritten & " 筆" & vbCrLf
    If Len(skipped) > 0 Then msg = msg & vbCrLf & "略過(認不出 IP 欄)的工作表:" & vbCrLf & skipped
    If Len(conflicts) > 0 Then msg = msg & vbCrLf & "名稱不同、未覆蓋(把 OVERWRITE 改 True 可覆蓋):" & vbCrLf & TrimList(conflicts)
    If Len(noNameCol) > 0 Then msg = msg & vbCrLf & "IP 已存在但該表無名稱欄、無法補名:" & vbCrLf & TrimList(noNameCol)
    MsgBox msg, vbInformation, "合併資產表"
    Exit Sub

readErr:
    MsgBox "讀取清單失敗:" & Err.Description & vbCrLf & _
           "若是中文亂碼,把 CSV_CHARSET 改成 ""big5"" 再試。", vbExclamation
End Sub

' ---- 從字串抓第一個合法 IPv4 ----
Function ExtractIP(s As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})"
        re.Global = False
    End If
    ExtractIP = ""
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        Dim j As Long, ok As Boolean: ok = True
        For j = 0 To 3
            If CLng(m.SubMatches(j)) > 255 Then ok = False
        Next j
        If ok Then ExtractIP = m.Value
    End If
End Function

' ---- 偵測某工作表的 IP / 名稱 / 角色 欄(回傳欄號,0=找不到)----
Sub DetectColumns(ws As Worksheet, ByRef ipCol As Long, ByRef nameCol As Long, ByRef roleCol As Long)
    ipCol = 0: nameCol = 0: roleCol = 0
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    Dim c1 As Long, c2 As Long, r1 As Long, r2 As Long, c As Long, r As Long, h As String
    c1 = ur.Column: c2 = ur.Column + ur.Columns.Count - 1
    r1 = ur.Row: r2 = ur.Row + ur.Rows.Count - 1

    For c = c1 To c2
        h = LCase(Trim(CStr(ws.Cells(HEADER_ROW, c).Value)))
        If Len(h) > 0 Then
            If ipCol = 0 And (h = "ip" Or InStr(h, "位址") > 0 Or InStr(h, "地址") > 0 Or InStr(h, "ip address") > 0 Or InStr(h, "ipaddress") > 0) Then
                ipCol = c
            ElseIf nameCol = 0 And (InStr(h, "名稱") > 0 Or InStr(h, "name") > 0 Or InStr(h, "設備") > 0 Or InStr(h, "裝置") > 0 Or InStr(h, "主機") > 0 Or InStr(h, "hostname") > 0) Then
                nameCol = c
            ElseIf roleCol = 0 And (InStr(h, "角色") > 0 Or InStr(h, "role") > 0 Or InStr(h, "用途") > 0 Or InStr(h, "類型") > 0 Or InStr(h, "type") > 0 Or InStr(h, "說明") > 0 Or InStr(h, "備註") > 0) Then
                roleCol = c
            End If
        End If
    Next c

    If ipCol = 0 Then
        Dim best As Long, bestCnt As Long, cnt As Long, rmax As Long
        best = 0: bestCnt = 0
        rmax = r2: If rmax > r1 + 200 Then rmax = r1 + 200
        For c = c1 To c2
            cnt = 0
            For r = r1 To rmax
                If Len(ExtractIP(CStr(ws.Cells(r, c).Value))) > 0 Then cnt = cnt + 1
            Next r
            If cnt > bestCnt Then bestCnt = cnt: best = c
        Next c
        If bestCnt > 0 Then ipCol = best
    End If
End Sub

Function SheetLastRow(ws As Worksheet) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then SheetLastRow = HEADER_ROW Else SheetLastRow = ur.Row + ur.Rows.Count - 1
End Function

' ---- 依名稱或索引數字找工作表 ----
Function ResolveSheet(wbk As Workbook, id As String) As Worksheet
    Dim ws As Worksheet
    If IsNumeric(id) Then
        Dim idx As Long: idx = CLng(id)
        If idx >= 1 And idx <= wbk.Worksheets.Count Then Set ResolveSheet = wbk.Worksheets(idx): Exit Function
    End If
    For Each ws In wbk.Worksheets
        If Trim(ws.Name) = Trim(id) Or LCase(Trim(ws.Name)) = LCase(Trim(id)) Then Set ResolveSheet = ws: Exit Function
    Next ws
    On Error Resume Next
    Set ResolveSheet = wbk.Worksheets(id)
    On Error GoTo 0
End Function

Function ListSheets(wbk As Workbook) As String
    Dim ws As Worksheet, i As Long, s As String
    i = 0
    For Each ws In wbk.Worksheets
        i = i + 1: s = s & "  [" & i & "] " & ws.Name & vbCrLf
    Next ws
    ListSheets = s
End Function

' 報告清單過長時只留前 20 行
Function TrimList(s As String) As String
    Dim a() As String: a = Split(s, vbCrLf)
    If UBound(a) <= 20 Then TrimList = s: Exit Function
    Dim i As Long, o As String
    For i = 0 To 19: o = o & a(i) & vbCrLf: Next i
    TrimList = o & "  …(還有更多,共 " & (UBound(a)) & " 筆)" & vbCrLf
End Function
