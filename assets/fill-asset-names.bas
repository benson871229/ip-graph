Attribute VB_Name = "FillAssetNames"
Option Explicit

' =====================================================================
'  補資產表的「設備名稱 / 主機名稱」(Excel VBA)— 同一活頁簿內,不跨檔案
'
'  規則(每一欄各自獨立判斷):
'    資產表該格是空的  -> 用來源的值填進去,標紅字
'    資產表該格已有值  -> 儲存格底色標黃,原值不動
'    IP 完全找不到     -> 在 ADD_SHEET 新增一列(紅字)
'
'  巨集「不會自動存檔」。跑完自己檢查紅字與黃底,確認無誤再另存。
'
'  用法:Alt+F11 -> 插入(Insert) -> 模組(Module) -> 貼上 -> 改下面設定 -> F5
' =====================================================================

' ====== 設定(依你的表調整這幾行)======
Const SOURCE_SHEET  As String = "工作表3"             ' 合併後的資料放在哪張工作表
Const TARGET_SHEETS As String = "表人給的,我看到的"   ' 要更新的資產表,逗號分隔
Const ADD_SHEET     As String = "我看到的"            ' 找不到的 IP 新增到這張
Const COL_DEVICE    As String = "設備名稱"            ' 欄位標題(可只寫部分字,會做包含比對)
Const COL_HOST      As String = "主機名稱"
Const HEADER_ROW    As Long = 1                       ' 標題列在第幾列

Sub 補設備與主機名稱()
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    If wbk Is Nothing Then MsgBox "找不到開啟中的活頁簿。", vbExclamation: Exit Sub

    ' ---- 來源工作表 ----
    Dim srcWs As Worksheet
    Set srcWs = ResolveSheet2(wbk, SOURCE_SHEET)
    If srcWs Is Nothing Then
        MsgBox "找不到來源工作表『" & SOURCE_SHEET & "』。" & vbCrLf & _
               "這本活頁簿的工作表:" & vbCrLf & ListSheets2(wbk), vbExclamation
        Exit Sub
    End If

    Dim sIp As Long, sDev As Long, sHost As Long
    sIp = FindCol(srcWs, "ip")
    sDev = FindCol(srcWs, COL_DEVICE)
    sHost = FindCol(srcWs, COL_HOST)
    If sIp = 0 Then
        MsgBox "來源工作表『" & srcWs.Name & "』找不到 IP 欄。" & vbCrLf & _
               "第 " & HEADER_ROW & " 列的標題:" & vbCrLf & HeaderList(srcWs), vbExclamation
        Exit Sub
    End If
    If sDev = 0 And sHost = 0 Then
        MsgBox "來源工作表『" & srcWs.Name & "』找不到『" & COL_DEVICE & "』或『" & COL_HOST & "』欄。" & vbCrLf & _
               "第 " & HEADER_ROW & " 列的標題:" & vbCrLf & HeaderList(srcWs), vbExclamation
        Exit Sub
    End If

    ' ---- 讀來源:ip -> 設備名稱 / 主機名稱 ----
    Dim ips() As String, devs() As String, hosts() As String, n As Long
    Dim lastR As Long: lastR = SheetLastRow2(srcWs)
    ReDim ips(1 To lastR + 1): ReDim devs(1 To lastR + 1): ReDim hosts(1 To lastR + 1): n = 0
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long, ip As String
    For r = HEADER_ROW + 1 To lastR
        ip = ExtractIP2(CStr(srcWs.Cells(r, sIp).Value))
        If Len(ip) > 0 And Not seen.Exists(ip) Then
            seen.Add ip, True
            n = n + 1
            ips(n) = ip
            devs(n) = IIf(sDev > 0, Trim(CStr(srcWs.Cells(r, sDev).Value)), "")
            hosts(n) = IIf(sHost > 0, Trim(CStr(srcWs.Cells(r, sHost).Value)), "")
        End If
    Next r
    If n = 0 Then MsgBox "來源工作表『" & srcWs.Name & "』沒讀到任何 IP。", vbExclamation: Exit Sub

    ' ---- 建立目標工作表的 ip -> 位置對照(一個 IP 可能出現在多張表)----
    Dim shNames() As String: shNames = Split(TARGET_SHEETS, ",")
    Dim loc As Object: Set loc = CreateObject("Scripting.Dictionary")
    Dim missSheets As String, s As Long, ws As Worksheet
    Dim tIp As Long, tDev As Long, tHost As Long, cellIP As String
    For s = LBound(shNames) To UBound(shNames)
        Set ws = ResolveSheet2(wbk, Trim(shNames(s)))
        If ws Is Nothing Then
            missSheets = missSheets & "  " & Trim(shNames(s)) & vbCrLf
        Else
            tIp = FindCol(ws, "ip")
            tDev = FindCol(ws, COL_DEVICE)
            tHost = FindCol(ws, COL_HOST)
            If tIp > 0 Then
                lastR = SheetLastRow2(ws)
                For r = HEADER_ROW + 1 To lastR
                    cellIP = ExtractIP2(CStr(ws.Cells(r, tIp).Value))
                    If Len(cellIP) > 0 Then
                        ' 值格式:工作表名|列|設備欄|主機欄,多筆用分號串接
                        Dim kv As String
                        kv = ws.Name & "|" & r & "|" & tDev & "|" & tHost
                        If loc.Exists(cellIP) Then
                            loc(cellIP) = loc(cellIP) & ";" & kv
                        Else
                            loc.Add cellIP, kv
                        End If
                    End If
                Next r
            End If
        End If
    Next s

    ' ---- 新增用工作表 ----
    Dim addWs As Worksheet
    Set addWs = ResolveSheet2(wbk, ADD_SHEET)
    If addWs Is Nothing Then
        MsgBox "找不到新增用工作表『" & ADD_SHEET & "』。" & vbCrLf & _
               "這本活頁簿的工作表:" & vbCrLf & ListSheets2(wbk), vbExclamation
        Exit Sub
    End If
    Dim aIp As Long, aDev As Long, aHost As Long
    aIp = FindCol(addWs, "ip")
    aDev = FindCol(addWs, COL_DEVICE)
    aHost = FindCol(addWs, COL_HOST)
    If aIp = 0 Then
        MsgBox "工作表『" & addWs.Name & "』找不到 IP 欄,無法新增。", vbExclamation
        Exit Sub
    End If
    Dim addRow As Long: addRow = SheetLastRow2(addWs)

    ' ---- 主流程 ----
    Application.ScreenUpdating = False
    Dim filledDev As Long, filledHost As Long, markedDev As Long, markedHost As Long
    Dim added As Long, noSrc As Long
    Dim e As Long, locs() As String, p As Long, f() As String
    Dim tWs As Worksheet, tRow As Long

    For e = 1 To n
        If loc.Exists(ips(e)) Then
            locs = Split(loc(ips(e)), ";")
            For p = LBound(locs) To UBound(locs)
                f = Split(locs(p), "|")
                Set tWs = wbk.Worksheets(f(0))
                tRow = CLng(f(1))
                ' 設備名稱與主機名稱各自獨立判斷
                FillOrMark tWs, tRow, CLng(f(2)), devs(e), filledDev, markedDev
                FillOrMark tWs, tRow, CLng(f(3)), hosts(e), filledHost, markedHost
            Next p
        Else
            addRow = addRow + 1
            addWs.Cells(addRow, aIp).Value = ips(e)
            addWs.Cells(addRow, aIp).Font.Color = vbRed
            If aDev > 0 And Len(devs(e)) > 0 Then
                addWs.Cells(addRow, aDev).Value = devs(e)
                addWs.Cells(addRow, aDev).Font.Color = vbRed
            End If
            If aHost > 0 And Len(hosts(e)) > 0 Then
                addWs.Cells(addRow, aHost).Value = hosts(e)
                addWs.Cells(addRow, aHost).Font.Color = vbRed
            End If
            added = added + 1
        End If
    Next e
    Application.ScreenUpdating = True

    ' ---- 報告 ----
    Dim msg As String
    msg = "完成(尚未存檔,請檢查後自行另存):" & vbCrLf & vbCrLf
    msg = msg & "來源『" & srcWs.Name & "』讀到 " & n & " 個 IP" & vbCrLf & vbCrLf
    msg = msg & "填入(紅字):" & vbCrLf
    msg = msg & "  " & COL_DEVICE & " " & filledDev & " 格" & vbCrLf
    msg = msg & "  " & COL_HOST & " " & filledHost & " 格" & vbCrLf & vbCrLf
    msg = msg & "已有值、標黃底未動:" & vbCrLf
    msg = msg & "  " & COL_DEVICE & " " & markedDev & " 格" & vbCrLf
    msg = msg & "  " & COL_HOST & " " & markedHost & " 格" & vbCrLf & vbCrLf
    msg = msg & "於『" & addWs.Name & "』新增:" & added & " 列" & vbCrLf
    If Len(missSheets) > 0 Then msg = msg & vbCrLf & "找不到的目標工作表:" & vbCrLf & missSheets
    MsgBox msg, vbInformation, "補設備與主機名稱"
End Sub

' ---- 單一儲存格:空的就填(紅字),有值就標黃底不動 ----
Sub FillOrMark(ws As Worksheet, r As Long, c As Long, val As String, _
               ByRef nFilled As Long, ByRef nMarked As Long)
    If c = 0 Then Exit Sub                      ' 這張表沒有這一欄
    Dim cur As String
    cur = Trim(CStr(ws.Cells(r, c).Value))
    If Len(cur) = 0 Then
        If Len(val) = 0 Then Exit Sub           ' 來源也沒值,不做事
        ws.Cells(r, c).Value = val
        ws.Cells(r, c).Font.Color = vbRed
        nFilled = nFilled + 1
    Else
        ws.Cells(r, c).Interior.Color = vbYellow
        nMarked = nMarked + 1
    End If
End Sub

' ---- 依標題文字找欄(包含比對,大小寫不敏感)----
Function FindCol(ws As Worksheet, title As String) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    Dim c1 As Long, c2 As Long, c As Long, h As String, t As String
    c1 = ur.Column: c2 = ur.Column + ur.Columns.Count - 1
    t = LCase(Trim(title))
    ' 先找完全相同的
    For c = c1 To c2
        h = LCase(Trim(CStr(ws.Cells(HEADER_ROW, c).Value)))
        If h = t Then FindCol = c: Exit Function
    Next c
    ' 再找包含的(例如標題是「設備名稱(中文)」)
    For c = c1 To c2
        h = LCase(Trim(CStr(ws.Cells(HEADER_ROW, c).Value)))
        If Len(h) > 0 And InStr(h, t) > 0 Then FindCol = c: Exit Function
    Next c
End Function

Function HeaderList(ws As Worksheet) As String
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    Dim c As Long, o As String, h As String
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        h = Trim(CStr(ws.Cells(HEADER_ROW, c).Value))
        If Len(h) > 0 Then o = o & "  [" & c & "] " & h & vbCrLf
    Next c
    HeaderList = o
End Function

Function ExtractIP2(s As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})"
        re.Global = False
    End If
    ExtractIP2 = ""
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        Dim j As Long, ok As Boolean: ok = True
        For j = 0 To 3
            If CLng(m.SubMatches(j)) > 255 Then ok = False
        Next j
        If ok Then ExtractIP2 = m.Value
    End If
End Function

Function SheetLastRow2(ws As Worksheet) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then SheetLastRow2 = HEADER_ROW Else SheetLastRow2 = ur.Row + ur.Rows.Count - 1
End Function

Function ResolveSheet2(wbk As Workbook, id As String) As Worksheet
    Dim ws As Worksheet
    If IsNumeric(id) Then
        Dim idx As Long: idx = CLng(id)
        If idx >= 1 And idx <= wbk.Worksheets.Count Then Set ResolveSheet2 = wbk.Worksheets(idx): Exit Function
    End If
    For Each ws In wbk.Worksheets
        If Trim(ws.Name) = Trim(id) Or LCase(Trim(ws.Name)) = LCase(Trim(id)) Then Set ResolveSheet2 = ws: Exit Function
    Next ws
End Function

Function ListSheets2(wbk As Workbook) As String
    Dim ws As Worksheet, i As Long, o As String
    i = 0
    For Each ws In wbk.Worksheets
        i = i + 1: o = o & "  [" & i & "] " & ws.Name & vbCrLf
    Next ws
    ListSheets2 = o
End Function
