Attribute VB_Name = "MergeAssets"
Option Explicit

' =====================================================================
'  合併資產表(Excel VBA 版)— 全部在同一個活頁簿內,不跨檔案
'
'  流程:
'    1. 讀「來源工作表」(SOURCE_SHEET,例如工作表3)所有的 ip -> 主機名稱
'    2. 在「比對工作表」(SEARCH_SHEETS)裡找每個 ip
'         - 有 match、名稱空白 -> 補上主機名稱(紅字)
'         - 都找不到           -> 統一在 ADD_SHEET 新增(紅字)
'    3. 所有更動一律紅字;巨集不自動存檔,檢查完自己另存
'
'  用法:Alt+F11 -> 插入 -> 模組 -> 貼上 -> 改下面設定 -> F5
' =====================================================================

' ====== 設定(依你的表名調整)======
Const SOURCE_SHEET As String = "工作表3"              ' 資料來源(防火牆抓的 ip+主機),可填名稱或索引數字
Const SEARCH_SHEETS As String = "表人給的,我看到的"    ' 要比對的工作表,逗號分隔
Const ADD_SHEET As String = "我看到的"                 ' 找不到的 ip 統一加到這張
Const HEADER_ROW As Long = 1                           ' 標題列在第幾列
Const OVERWRITE As Boolean = False                     ' True = 名稱不同時覆蓋成來源的名稱

Sub 合併資產表()
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    If wbk Is Nothing Then MsgBox "找不到開啟中的活頁簿。", vbExclamation: Exit Sub

    ' ---- 來源工作表 ----
    Dim srcWs As Worksheet
    Set srcWs = ResolveSheet(wbk, SOURCE_SHEET)
    If srcWs Is Nothing Then
        MsgBox "找不到來源工作表『" & SOURCE_SHEET & "』。" & vbCrLf & "工作表:" & vbCrLf & ListSheets(wbk), vbExclamation
        Exit Sub
    End If
    Dim sIc As Long, sNc As Long, sRc As Long
    DetectColumns srcWs, sIc, sNc, sRc
    If sIc = 0 Or sNc = 0 Then
        MsgBox "來源工作表『" & srcWs.Name & "』認不出 IP 欄或名稱欄。請確認第 " & HEADER_ROW & _
               " 列有『ip』與『名稱/name/主機』之類欄名。", vbExclamation
        Exit Sub
    End If

    ' ---- 讀來源 ip -> 名稱(取第一筆)----
    Dim ips() As String, names() As String, roles() As String, n As Long
    Dim lastR As Long: lastR = SheetLastRow(srcWs)
    ReDim ips(1 To lastR + 1): ReDim names(1 To lastR + 1): ReDim roles(1 To lastR + 1): n = 0
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long, ip As String, nm As String
    For r = HEADER_ROW + 1 To lastR
        ip = ExtractIP(CStr(srcWs.Cells(r, sIc).Value))
        If Len(ip) > 0 Then
            nm = Trim(CStr(srcWs.Cells(r, sNc).Value))
            If Len(nm) > 0 And Not seen.Exists(ip) Then
                seen.Add ip, True
                n = n + 1: ips(n) = ip: names(n) = nm
                If sRc > 0 Then roles(n) = Trim(CStr(srcWs.Cells(r, sRc).Value)) Else roles(n) = ""
            End If
        End If
    Next r
    If n = 0 Then MsgBox "來源工作表『" & srcWs.Name & "』沒讀到任何 ip+名稱。", vbExclamation: Exit Sub

    ' ---- 比對工作表:偵測欄位、建 ip -> 多個位置("wsName|row|nameCol|roleCol" 以分號串接)----
    Dim shNames() As String: shNames = Split(SEARCH_SHEETS, ",")
    Dim loc As Object: Set loc = CreateObject("Scripting.Dictionary")
    Dim missSheets As String, s As Long, ws As Worksheet, ic As Long, nc As Long, rc As Long, cellIP As String, keyStr As String
    For s = LBound(shNames) To UBound(shNames)
        Set ws = ResolveSheet(wbk, Trim(shNames(s)))
        If ws Is Nothing Then
            missSheets = missSheets & "  " & Trim(shNames(s)) & vbCrLf
        Else
            DetectColumns ws, ic, nc, rc
            If ic > 0 Then
                lastR = SheetLastRow(ws)
                For r = HEADER_ROW + 1 To lastR
                    cellIP = ExtractIP(CStr(ws.Cells(r, ic).Value))
                    If Len(cellIP) > 0 Then
                        keyStr = ws.Name & "|" & r & "|" & nc & "|" & rc
                        If loc.Exists(cellIP) Then loc(cellIP) = loc(cellIP) & ";" & keyStr Else loc.Add cellIP, keyStr
                    End If
                Next r
            End If
        End If
    Next s

    ' ---- 新增用工作表 ----
    Dim addWs As Worksheet
    Set addWs = ResolveSheet(wbk, ADD_SHEET)
    If addWs Is Nothing Then
        MsgBox "找不到新增用工作表『" & ADD_SHEET & "』。" & vbCrLf & "工作表:" & vbCrLf & ListSheets(wbk), vbExclamation
        Exit Sub
    End If
    Dim addIc As Long, addNc As Long, addRc As Long
    DetectColumns addWs, addIc, addNc, addRc
    If addIc = 0 Or addNc = 0 Then
        MsgBox "工作表『" & addWs.Name & "』認不出 IP 欄或名稱欄,無法新增。", vbExclamation
        Exit Sub
    End If
    Dim addRow As Long: addRow = SheetLastRow(addWs)

    ' ---- 合併(紅字)----
    Application.ScreenUpdating = False
    Dim filled As Long, added As Long, same As Long, overwritten As Long
    Dim conflicts As String, noNameCol As String
    Dim e As Long, locs() As String, p As Long, f() As String, cur As String
    Dim tWs As Worksheet, tRow As Long, tNc As Long, tRc As Long

    For e = 1 To n
        If loc.Exists(ips(e)) Then
            locs = Split(loc(ips(e)), ";")
            For p = LBound(locs) To UBound(locs)
                f = Split(locs(p), "|")
                Set tWs = wbk.Worksheets(f(0))
                tRow = CLng(f(1)): tNc = CLng(f(2)): tRc = CLng(f(3))
                If tNc = 0 Then
                    noNameCol = noNameCol & "  " & ips(e) & " (在『" & tWs.Name & "』)" & vbCrLf
                Else
                    cur = Trim(CStr(tWs.Cells(tRow, tNc).Value))
                    If Len(cur) = 0 Then
                        tWs.Cells(tRow, tNc).Value = names(e)
                        tWs.Cells(tRow, tNc).Font.Color = vbRed
                        If tRc > 0 And Len(roles(e)) > 0 Then
                            If Len(Trim(CStr(tWs.Cells(tRow, tRc).Value))) = 0 Then
                                tWs.Cells(tRow, tRc).Value = roles(e)
                                tWs.Cells(tRow, tRc).Font.Color = vbRed
                            End If
                        End If
                        filled = filled + 1
                    ElseIf cur = names(e) Then
                        same = same + 1
                    ElseIf OVERWRITE Then
                        tWs.Cells(tRow, tNc).Value = names(e)
                        tWs.Cells(tRow, tNc).Font.Color = vbRed
                        overwritten = overwritten + 1
                    Else
                        conflicts = conflicts & "  『" & tWs.Name & "』第 " & tRow & " 列 " & ips(e) & _
                                    ":表為『" & cur & "』,來源為『" & names(e) & "』" & vbCrLf
                    End If
                End If
            Next p
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
    msg = msg & "來源『" & srcWs.Name & "』讀到:" & n & " 個 ip" & vbCrLf
    msg = msg & "補上名稱:" & filled & " 筆" & vbCrLf
    msg = msg & "於『" & addWs.Name & "』新增:" & added & " 列" & vbCrLf
    msg = msg & "原本就相同:" & same & " 筆" & vbCrLf
    If OVERWRITE Then msg = msg & "覆蓋:" & overwritten & " 筆" & vbCrLf
    If Len(missSheets) > 0 Then msg = msg & vbCrLf & "找不到的比對工作表:" & vbCrLf & missSheets
    If Len(conflicts) > 0 Then msg = msg & vbCrLf & "名稱不同、未覆蓋(OVERWRITE 改 True 可覆蓋):" & vbCrLf & TrimList(conflicts)
    If Len(noNameCol) > 0 Then msg = msg & vbCrLf & "ip 已存在但該表無名稱欄:" & vbCrLf & TrimList(noNameCol)
    MsgBox msg, vbInformation, "合併資產表"
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

' ---- 偵測工作表的 IP / 名稱 / 角色 欄(0=找不到)----
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
End Function

Function ListSheets(wbk As Workbook) As String
    Dim ws As Worksheet, i As Long, o As String
    i = 0
    For Each ws In wbk.Worksheets
        i = i + 1: o = o & "  [" & i & "] " & ws.Name & vbCrLf
    Next ws
    ListSheets = o
End Function

Function TrimList(s As String) As String
    Dim a() As String: a = Split(s, vbCrLf)
    If UBound(a) <= 20 Then TrimList = s: Exit Function
    Dim i As Long, o As String
    For i = 0 To 19: o = o & a(i) & vbCrLf: Next i
    TrimList = o & "  …(共 " & (UBound(a)) & " 筆)" & vbCrLf
End Function
