Attribute VB_Name = "MergeCsv"
Option Explicit

' =====================================================================
'  合併多份 CSV,並附加進 Excel 資產清冊(Excel VBA 版)
'
'  情境:同樣欄位的 CSV 有好幾份(例如 5 台防火牆各匯出一份),
'        要疊成一份再加進資產清冊。
'
'  兩個進入點(游標點進去按 F5 擇一執行):
'    合併CSV並附加到資產表   選檔 → 合併 → 直接附加到 ADD_SHEET(紅字、不存檔)
'    合併CSV另存新檔         選檔 → 合併 → 只輸出一份合併後的 CSV(UTF-8 BOM)
'
'  合併規則:
'    * 只保留第一份的表頭,其餘檔案的表頭列跳過
'    * 同一個 IP 出現在多份檔案 → 全部保留,不去重(這是刻意的)
'    * 多加一欄「來源檔」,記錄每一列來自哪個檔案,方便日後追查
'    * 欄位以「欄名」對應到工作表既有欄位;對不到的欄位加在最右邊(表頭也是紅字)
'    * port 欄預設不寫回工作表的 port 欄(該欄已知不準),改放獨立新欄保存原值
'    * 所有更動一律紅字;巨集不自動存檔,請自己檢查後另存
'
'  吃得下的輸入:
'    UTF-8(有/沒有 BOM)、UTF-16 LE/BE、Big5;CRLF / LF / CR 混用;
'    欄位含逗號、雙引號("" 逃逸)、含換行的引號欄位
'
'  相容性:本模組的共用函式全部宣告 Private 並加 Csv 前綴,
'          與 merge-assets.bas(MergeAssets 模組)放在同一個活頁簿也不會
'          出現「名稱模稜兩可 (Ambiguous name detected)」。
'
'  用法:Alt+F11 → 插入 → 模組 → 貼上 → 改下面設定 → F5
' =====================================================================

' ====== 設定(依你的資產清冊調整)======
Private Const ADD_SHEET As String = "我看到的"          ' 合併結果要附加到哪張工作表(名稱或索引數字)
Private Const HEADER_ROW As Long = 1                    ' 標題列在第幾列
Private Const SOURCE_COL_NAME As String = "來源檔"      ' 記錄來源檔名的欄位名稱;設成 "" 就不加這欄
Private Const SOURCE_FULL_PATH As Boolean = False       ' True = 來源檔記完整路徑,False = 只記檔名
Private Const ADD_MISSING_COLUMNS As Boolean = True     ' CSV 有、工作表沒有的欄位 → 加在最右邊
Private Const MAP_PORT As Boolean = False               ' CSV 的 port 欄要不要寫進工作表既有的 port 欄
                                                        ' (預設 False:該欄已知不準,不去動它)
Private Const WRITE_AS_TEXT As Boolean = True           ' 新增的儲存格先設成文字格式,
                                                        ' 避免 Excel 把 "1-5"、"03" 自作聰明轉成日期/數字
Private Const REPORT_EXISTING As Boolean = True         ' 報告有幾列的 IP 在活頁簿裡已經出現過(只報告,不做任何事)
Private Const OUT_CSV_FOLDER As String = ""             ' 另存合併 CSV 的資料夾;留空 = 跟第一個來源檔同資料夾

' =====================================================================
'  進入點 1:合併 CSV 並附加到資產表
' =====================================================================
Sub 合併CSV並附加到資產表()
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    If wbk Is Nothing Then MsgBox "找不到開啟中的活頁簿。", vbExclamation: Exit Sub

    Dim addWs As Worksheet
    Set addWs = CsvResolveSheet(wbk, ADD_SHEET)
    If addWs Is Nothing Then
        MsgBox "找不到工作表『" & ADD_SHEET & "』。" & vbCrLf & _
               "請把上面的 ADD_SHEET 改成下列其中一個:" & vbCrLf & CsvListSheets(wbk), vbExclamation
        Exit Sub
    End If

    ' ---- 選檔 + 合併 ----
    Dim files As Variant
    files = CsvPickFiles()
    If IsEmpty(files) Then Exit Sub

    Dim hdr() As String, nH As Long, data As Collection, srcs As Collection, info As String
    If Not CsvMergeFiles(files, hdr, nH, data, srcs, info) Then Exit Sub
    If data.Count = 0 Then MsgBox "合併後沒有任何資料列。" & vbCrLf & vbCrLf & info, vbExclamation: Exit Sub

    ' ---- 組出「合併後的欄名清單」(含來源檔欄)----
    Dim cols() As String, nC As Long
    nC = nH
    ReDim cols(0 To nH)          ' 先多留一格給來源檔欄
    Dim j As Long
    For j = 0 To nH - 1: cols(j) = hdr(j): Next j
    Dim srcCol As Long: srcCol = -1
    If Len(SOURCE_COL_NAME) > 0 Then
        cols(nC) = CsvUniqueName(SOURCE_COL_NAME, hdr, nH)
        srcCol = nC
        nC = nC + 1
    End If

    ' ---- 讀工作表標題列,做欄位對應 ----
    Dim tLastCol As Long, tHdr() As String
    tLastCol = CsvSheetLastCol(addWs)
    If tLastCol < 1 Then tLastCol = 1
    ReDim tHdr(1 To tLastCol)
    Dim c As Long
    For c = 1 To tLastCol: tHdr(c) = Trim$(CStr(addWs.Cells(HEADER_ROW, c).Value)): Next c

    Dim tUsed() As Boolean: ReDim tUsed(1 To tLastCol + nC)
    Dim map() As Long: ReDim map(0 To nC - 1)     ' 合併欄 -> 工作表欄(0 = 不寫入)
    Dim newCols As Long: newCols = 0
    Dim maxCol As Long: maxCol = tLastCol
    Dim mapLog As String, skipLog As String

    For j = 0 To nC - 1
        map(j) = CsvMatchTargetCol(cols(j), tHdr, tUsed, tLastCol)
        If map(j) > 0 Then
            tUsed(map(j)) = True
            mapLog = mapLog & "  " & cols(j) & "  →  第 " & map(j) & " 欄『" & tHdr(map(j)) & "』" & vbCrLf
        ElseIf ADD_MISSING_COLUMNS Then
            maxCol = maxCol + 1
            map(j) = maxCol
            tUsed(map(j)) = True
            newCols = newCols + 1
            mapLog = mapLog & "  " & cols(j) & "  →  第 " & map(j) & " 欄(新增)" & vbCrLf
        Else
            skipLog = skipLog & "  " & cols(j) & vbCrLf
        End If
    Next j

    ' ---- 先掃活頁簿既有 IP(只為了報告,不會改任何東西)----
    Dim seenIP As Object: Set seenIP = Nothing
    If REPORT_EXISTING Then Set seenIP = CsvCollectWorkbookIPs(wbk)

    ' ---- 準備要寫入的二維陣列(一次寫入,比逐格快很多)----
    Dim firstRow As Long: firstRow = CsvSheetLastRow(addWs) + 1
    If firstRow <= HEADER_ROW Then firstRow = HEADER_ROW + 1
    Dim nData As Long: nData = data.Count
    If firstRow + nData - 1 > addWs.Rows.Count Then
        MsgBox "資料太多,附加後會超過工作表的最大列數(" & addWs.Rows.Count & ")。", vbExclamation
        Exit Sub
    End If
    Dim outVals() As Variant
    ReDim outVals(1 To nData, 1 To maxCol)

    Dim ipColIdx As Long: ipColIdx = CsvFindIpColumn(cols, nC)
    Dim dupExisting As Long: dupExisting = 0
    Dim r As Long, rw As Variant, v As String, ipStr As String
    For r = 1 To nData
        rw = data(r)
        For j = 0 To nC - 1
            If map(j) > 0 Then
                If j = srcCol Then
                    v = CStr(srcs(r))
                Else
                    v = CsvField(rw, j)
                End If
                If Len(v) > 0 Then outVals(r, map(j)) = v
            End If
        Next j
        If Not seenIP Is Nothing And ipColIdx >= 0 Then
            ipStr = CsvExtractIP(CsvField(rw, ipColIdx))
            If Len(ipStr) > 0 Then
                If seenIP.Exists(ipStr) Then dupExisting = dupExisting + 1
            End If
        End If
    Next r

    ' ---- 寫入(全部紅字、不存檔)----
    Application.ScreenUpdating = False
    Dim blk As Range
    Set blk = addWs.Range(addWs.Cells(firstRow, 1), addWs.Cells(firstRow + nData - 1, maxCol))
    If WRITE_AS_TEXT Then blk.NumberFormat = "@"
    blk.Value = outVals
    blk.Font.Color = vbRed

    ' 新增欄位的表頭(也是紅字)
    For j = 0 To nC - 1
        If map(j) > tLastCol Then
            addWs.Cells(HEADER_ROW, map(j)).Value = cols(j)
            addWs.Cells(HEADER_ROW, map(j)).Font.Color = vbRed
        End If
    Next j
    Application.ScreenUpdating = True

    ' ---- 報告 ----
    Dim msg As String
    msg = "完成(更動皆為紅字,尚未存檔,請檢查後自行另存):" & vbCrLf & vbCrLf
    msg = msg & info & vbCrLf
    msg = msg & "合併後共 " & nData & " 列 / " & nC & " 欄(同 IP 不去重,全部保留)" & vbCrLf
    msg = msg & "已附加到『" & addWs.Name & "』第 " & firstRow & " ~ " & (firstRow + nData - 1) & " 列" & vbCrLf
    If newCols > 0 Then msg = msg & "在最右邊新增了 " & newCols & " 個欄位" & vbCrLf
    If REPORT_EXISTING And ipColIdx >= 0 Then
        msg = msg & "其中 " & dupExisting & " 列的 IP 在這本活頁簿裡原本就出現過(只是提醒,沒有做任何合併)" & vbCrLf
    End If
    msg = msg & vbCrLf & "欄位對應:" & vbCrLf & CsvTrimList(mapLog)
    If Len(skipLog) > 0 Then msg = msg & vbCrLf & "沒有對應、未寫入的欄位(ADD_MISSING_COLUMNS 改 True 可自動新增):" & vbCrLf & CsvTrimList(skipLog)
    If Not MAP_PORT Then msg = msg & vbCrLf & "註:port 欄不會寫進工作表原本的 port 欄(該欄已知不準),原值另存到新欄位。"
    MsgBox msg, vbInformation, "合併 CSV 到資產表"
End Sub

' =====================================================================
'  進入點 2:只合併成一份 CSV(UTF-8 BOM,Excel 直接打開不會亂碼)
' =====================================================================
Sub 合併CSV另存新檔()
    Dim files As Variant
    files = CsvPickFiles()
    If IsEmpty(files) Then Exit Sub

    Dim hdr() As String, nH As Long, data As Collection, srcs As Collection, info As String
    If Not CsvMergeFiles(files, hdr, nH, data, srcs, info) Then Exit Sub
    If data.Count = 0 Then MsgBox "合併後沒有任何資料列。" & vbCrLf & vbCrLf & info, vbExclamation: Exit Sub

    ' ---- 組文字 ----
    Dim sb As String, line As String, j As Long, r As Long, rw As Variant
    Dim srcName As String: srcName = ""
    If Len(SOURCE_COL_NAME) > 0 Then srcName = CsvUniqueName(SOURCE_COL_NAME, hdr, nH)

    For j = 0 To nH - 1
        If j > 0 Then line = line & ","
        line = line & CsvEscape(hdr(j))
    Next j
    If Len(srcName) > 0 Then line = line & "," & CsvEscape(srcName)
    sb = line & vbCrLf

    Dim nOut As Long
    nOut = nH
    If Len(srcName) > 0 Then nOut = nH + 1
    Dim parts() As String
    ReDim parts(0 To nOut - 1)
    Dim chunks As Collection: Set chunks = New Collection   ' 分段累積,避免大檔一直重配置字串
    For r = 1 To data.Count
        rw = data(r)
        For j = 0 To nH - 1: parts(j) = CsvEscape(CsvField(rw, j)): Next j
        If Len(srcName) > 0 Then parts(nH) = CsvEscape(CStr(srcs(r)))
        sb = sb & Join(parts, ",") & vbCrLf
        If Len(sb) > 500000 Then chunks.Add sb: sb = ""
    Next r
    chunks.Add sb
    sb = ""
    For r = 1 To chunks.Count: sb = sb & chunks(r): Next r

    ' ---- 輸出路徑 ----
    Dim folder As String
    If Len(OUT_CSV_FOLDER) > 0 Then
        folder = OUT_CSV_FOLDER
    Else
        folder = Left$(CStr(files(LBound(files))), InStrRev(CStr(files(LBound(files))), "\"))
    End If
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Dim outPath As String
    outPath = folder & "合併-" & Format$(Now, "yyyymmdd-HHmmss") & ".csv"

    ' UTF-8 with BOM:Excel 直接打開才不會中文亂碼
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText sb
    st.SaveToFile outPath, 2    ' adSaveCreateOverWrite
    st.Close

    MsgBox "完成:" & vbCrLf & vbCrLf & info & vbCrLf & _
           "合併後共 " & data.Count & " 列(同 IP 不去重,全部保留)" & vbCrLf & vbCrLf & _
           "已輸出(UTF-8 BOM):" & vbCrLf & outPath, vbInformation, "合併 CSV"
End Sub

' =====================================================================
'  選檔
' =====================================================================
Private Function CsvPickFiles() As Variant
    Dim f As Variant
    f = Application.GetOpenFilename( _
            "CSV / 文字檔 (*.csv;*.txt),*.csv;*.txt,所有檔案 (*.*),*.*", 1, _
            "選擇要合併的 CSV(可按住 Ctrl 或 Shift 多選)", , True)
    If VarType(f) = vbBoolean Then CsvPickFiles = Empty: Exit Function   ' 使用者按取消
    ' 依檔名排序,讓合併順序固定(選取順序是不可預期的)
    Dim i As Long, k As Long, tmp As String
    For i = LBound(f) To UBound(f) - 1
        For k = i + 1 To UBound(f)
            If LCase$(CsvBaseName(CStr(f(k)))) < LCase$(CsvBaseName(CStr(f(i)))) Then
                tmp = CStr(f(i)): f(i) = f(k): f(k) = tmp
            End If
        Next k
    Next i
    CsvPickFiles = f
End Function

' =====================================================================
'  合併核心
'    hdr / nH  = 合併後的欄名(不含來源檔欄)
'    data      = Collection of String(),每個元素是一列(0-based,長度可能小於 nH)
'    srcs      = Collection of String,與 data 一一對應的來源檔名
' =====================================================================
Private Function CsvMergeFiles(files As Variant, ByRef hdr() As String, ByRef nH As Long, _
                               ByRef data As Collection, ByRef srcs As Collection, _
                               ByRef info As String) As Boolean
    CsvMergeFiles = False
    Set data = New Collection
    Set srcs = New Collection
    nH = 0
    ReDim hdr(0 To 15)

    Dim warn As String
    Dim baseKeys As String: baseKeys = ""
    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim path As String: path = CStr(files(i))
        Dim enc As String, txt As String
        txt = CsvReadTextAuto(path, enc)

        Dim rows As Collection
        Set rows = CsvParse(txt)
        If rows.Count = 0 Then
            info = info & "  " & CsvBaseName(path) & "(" & enc & "):空檔,已略過" & vbCrLf
            GoTo NextFile
        End If

        ' ---- 表頭 -> 欄位對應(以欄名對應,所以就算欄序不同也接得起來)----
        Dim fh As Variant: fh = rows(1)
        Dim nf As Long: nf = UBound(fh) - LBound(fh) + 1
        Dim map() As Long: ReDim map(0 To nf - 1)
        Dim usedH() As Boolean: ReDim usedH(0 To nH + nf)
        Dim keys As String: keys = ""
        Dim j As Long, hn As String, k As Long, found As Long
        For j = 0 To nf - 1
            hn = Trim$(CStr(fh(j + LBound(fh))))
            If Len(hn) = 0 Then hn = "(未命名欄" & (j + 1) & ")"
            keys = keys & "|" & CsvNormHdr(hn)
            found = -1
            For k = 0 To nH - 1
                If Not usedH(k) Then
                    If CsvNormHdr(hdr(k)) = CsvNormHdr(hn) Then found = k: Exit For
                End If
            Next k
            If found < 0 Then
                If nH > UBound(hdr) Then ReDim Preserve hdr(0 To nH + 15)
                hdr(nH) = hn
                found = nH
                nH = nH + 1
            End If
            usedH(found) = True
            map(j) = found
        Next j

        If i = LBound(files) Then
            baseKeys = keys
        ElseIf keys <> baseKeys Then
            warn = warn & "  " & CsvBaseName(path) & vbCrLf
        End If

        ' ---- 資料列 ----
        Dim srcTag As String
        If SOURCE_FULL_PATH Then srcTag = path Else srcTag = CsvBaseName(path)
        Dim r As Long, cnt As Long: cnt = 0
        For r = 2 To rows.Count
            Dim src As Variant: src = rows(r)
            Dim out() As String
            ReDim out(0 To nH - 1)
            For j = 0 To nf - 1
                If (j + LBound(src)) <= UBound(src) Then out(map(j)) = CStr(src(j + LBound(src)))
            Next j
            data.Add out
            srcs.Add srcTag
            cnt = cnt + 1
        Next r
        info = info & "  " & CsvBaseName(path) & "(" & enc & "):" & cnt & " 列" & vbCrLf
NextFile:
    Next i

    If nH = 0 Then
        MsgBox "選到的檔案都讀不到表頭,請確認第一列是欄位名稱。" & vbCrLf & vbCrLf & info, vbExclamation
        Exit Function
    End If
    ReDim Preserve hdr(0 To nH - 1)
    info = "讀入 " & (UBound(files) - LBound(files) + 1) & " 個檔案:" & vbCrLf & info
    If Len(warn) > 0 Then
        info = info & "注意:下列檔案的欄位與第一個檔案不完全相同,已以「欄名」對應" & vbCrLf & _
               "(缺的欄留空、多的欄自動加在後面):" & vbCrLf & warn
    End If
    CsvMergeFiles = True
End Function

' =====================================================================
'  CSV 解析(狀態機:引號欄位、"" 逃逸、欄內換行、CRLF/LF/CR 都吃)
'  回傳 Collection of String()(0-based),全空白的列會略過
' =====================================================================
Private Function CsvParse(ByVal s As String) As Collection
    Dim rows As Collection: Set rows = New Collection
    Set CsvParse = rows
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = ChrW$(&HFEFF) Then s = Mid$(s, 2)     ' 保險:再剝一次 BOM 字元
    If Len(s) = 0 Then Exit Function

    ' 快路徑:整份檔案沒有雙引號時,直接切行切逗號(大檔快很多)
    If InStr(s, """") = 0 Then
        s = Replace(s, vbCrLf, vbLf)
        s = Replace(s, vbCr, vbLf)
        Dim lines() As String: lines = Split(s, vbLf)
        Dim i2 As Long
        For i2 = LBound(lines) To UBound(lines)
            If Len(Trim$(Replace(lines(i2), ",", ""))) > 0 Then rows.Add Split(lines(i2), ",")
        Next i2
        Exit Function
    End If

    Dim fields() As String, nf As Long
    ReDim fields(0 To 31)
    Dim fld As String, inQ As Boolean
    Dim i As Long, n As Long, ch As String, nxt As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n Then nxt = Mid$(s, i + 1, 1) Else nxt = ""
                If nxt = """" Then
                    fld = fld & """"
                    i = i + 2
                Else
                    inQ = False
                    i = i + 1
                End If
            Else
                fld = fld & ch
                i = i + 1
            End If
        Else
            Select Case ch
                Case """"
                    ' 只有出現在欄位開頭的引號才視為引號欄位;欄位中間的引號當一般字元
                    If Len(fld) = 0 Then inQ = True Else fld = fld & ch
                    i = i + 1
                Case ","
                    CsvPushField fields, nf, fld
                    fld = ""
                    i = i + 1
                Case vbCr, vbLf
                    CsvPushField fields, nf, fld
                    fld = ""
                    CsvPushRow rows, fields, nf
                    If ch = vbCr Then
                        If i < n Then
                            If Mid$(s, i + 1, 1) = vbLf Then i = i + 1
                        End If
                    End If
                    i = i + 1
                Case Else
                    fld = fld & ch
                    i = i + 1
            End Select
        End If
    Loop
    If Len(fld) > 0 Or nf > 0 Then
        CsvPushField fields, nf, fld
        CsvPushRow rows, fields, nf
    End If
End Function

Private Sub CsvPushField(ByRef fields() As String, ByRef nf As Long, ByVal v As String)
    If nf > UBound(fields) Then ReDim Preserve fields(0 To nf + 31)
    fields(nf) = v
    nf = nf + 1
End Sub

Private Sub CsvPushRow(ByRef rows As Collection, ByRef fields() As String, ByRef nf As Long)
    Dim j As Long, blank As Boolean: blank = True
    For j = 0 To nf - 1
        If Len(Trim$(fields(j))) > 0 Then blank = False: Exit For
    Next j
    If Not blank Then
        Dim out() As String
        ReDim out(0 To nf - 1)
        For j = 0 To nf - 1: out(j) = fields(j): Next j
        rows.Add out
    End If
    nf = 0
End Sub

' ---- 取欄位值;列比表頭短時回空字串 ----
Private Function CsvField(row As Variant, ByVal j As Long) As String
    If j < LBound(row) Or j > UBound(row) Then
        CsvField = ""
    Else
        CsvField = CStr(row(j))
    End If
End Function

' ---- 輸出 CSV 時的跳脫:含逗號/引號/換行/前後空白就加引號 ----
Private Function CsvEscape(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 _
       Or s <> Trim$(s) Then
        CsvEscape = """" & Replace(s, """", """""") & """"
    Else
        CsvEscape = s
    End If
End Function

' =====================================================================
'  讀檔 + 編碼自動偵測
'    BOM 優先;無 BOM 時用 null byte 分布猜 UTF-16;
'    再不行先試 UTF-8,出現替代字元 U+FFFD 就退回 Big5(ANSI/zh-TW)
'  註:VBA 直接 Open For Input 讀 UTF-8 會亂碼,一定要走 ADODB.Stream + Charset
' =====================================================================
Private Function CsvReadTextAuto(ByVal path As String, ByRef encName As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                 ' adTypeBinary
    st.Open
    st.LoadFromFile path
    Dim b() As Byte
    If st.Size = 0 Then
        st.Close
        encName = "空檔"
        CsvReadTextAuto = ""
        Exit Function
    End If
    b = st.Read
    st.Close

    Dim n As Long: n = UBound(b) - LBound(b) + 1

    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            encName = "UTF-8 BOM"
            CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "utf-8"))
            Exit Function
        End If
    End If
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then
            encName = "UTF-16 LE"
            CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "unicode"))
            Exit Function
        End If
        If b(0) = &HFE And b(1) = &HFF Then
            encName = "UTF-16 BE"
            CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "unicodeFFFE"))
            Exit Function
        End If
    End If

    ' 無 BOM:ASCII 字在 UTF-16 LE 的奇數位是 0、BE 的偶數位是 0
    Dim lim As Long: lim = n
    If lim > 4000 Then lim = 4000
    Dim i As Long, evenNull As Long, oddNull As Long
    For i = 0 To lim - 1
        If b(i) = 0 Then
            If i Mod 2 = 0 Then evenNull = evenNull + 1 Else oddNull = oddNull + 1
        End If
    Next i
    If (evenNull + oddNull) > (lim \ 4) Then
        If oddNull > evenNull Then
            encName = "UTF-16 LE(無 BOM,推測)"
            CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "unicode"))
        Else
            encName = "UTF-16 BE(無 BOM,推測)"
            CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "unicodeFFFE"))
        End If
        Exit Function
    End If

    Dim t As String
    t = CsvDecode(b, "utf-8")
    If InStr(t, ChrW$(&HFFFD)) > 0 Then
        encName = "Big5(UTF-8 解碼失敗後退回)"
        CsvReadTextAuto = CsvStripBOM(CsvDecode(b, "big5"))
    Else
        encName = "UTF-8"
        CsvReadTextAuto = CsvStripBOM(t)
    End If
End Function

Private Function CsvDecode(b() As Byte, ByVal charset As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                 ' adTypeBinary
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2                 ' adTypeText
    st.charset = charset
    CsvDecode = st.ReadText
    st.Close
End Function

Private Function CsvStripBOM(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW$(&HFEFF) Then s = Mid$(s, 2)
    End If
    CsvStripBOM = s
End Function

' =====================================================================
'  欄位對應
' =====================================================================
' 欄名正規化:去空白、去底線/連字號、轉小寫
Private Function CsvNormHdr(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    CsvNormHdr = s
End Function

' 把欄名歸到一個「同義群組」,讓 CSV 欄名跟工作表欄名不完全一樣時也對得起來。
' 注意判斷順序:「設備用途」要先被 用途 接走,否則會被 設備 誤判成名稱欄。
Private Function CsvAliasKey(ByVal s As String) As String
    Dim h As String: h = CsvNormHdr(s)
    CsvAliasKey = ""
    If Len(h) = 0 Then Exit Function
    If InStr(h, "用途") > 0 Or InStr(h, "角色") > 0 Or InStr(h, "role") > 0 Or _
       InStr(h, "類型") > 0 Or InStr(h, "說明") > 0 Or InStr(h, "備註") > 0 Then
        CsvAliasKey = "role": Exit Function
    End If
    If InStr(h, "網系") > 0 Or InStr(h, "網段") > 0 Or InStr(h, "vlan") > 0 Or _
       InStr(h, "segment") > 0 Or InStr(h, "subnet") > 0 Then
        CsvAliasKey = "net": Exit Function
    End If
    If InStr(h, "位置") > 0 Or InStr(h, "機房") > 0 Or InStr(h, "樓層") > 0 Or _
       InStr(h, "location") > 0 Then
        CsvAliasKey = "loc": Exit Function
    End If
    If h = "port" Or InStr(h, "連接埠") > 0 Or InStr(h, "埠號") > 0 Then
        CsvAliasKey = "port": Exit Function
    End If
    If h = "ip" Or h = "address" Or InStr(h, "ipaddress") > 0 Or _
       InStr(h, "位址") > 0 Or InStr(h, "地址") > 0 Then
        CsvAliasKey = "ip": Exit Function
    End If
    If InStr(h, "名稱") > 0 Or InStr(h, "name") > 0 Or InStr(h, "主機") > 0 Or _
       InStr(h, "hostname") > 0 Or InStr(h, "設備") > 0 Or InStr(h, "裝置") > 0 Then
        CsvAliasKey = "name": Exit Function
    End If
End Function

' 回傳工作表的欄號;0 = 對不到
Private Function CsvMatchTargetCol(ByVal name As String, ByRef tHdr() As String, _
                                   ByRef tUsed() As Boolean, ByVal lastCol As Long) As Long
    Dim c As Long
    CsvMatchTargetCol = 0
    If lastCol < 1 Then Exit Function
    ' 1) 欄名完全相同(正規化後)
    For c = 1 To lastCol
        If Not tUsed(c) Then
            If Len(tHdr(c)) > 0 Then
                If CsvNormHdr(tHdr(c)) = CsvNormHdr(name) Then CsvMatchTargetCol = c: Exit Function
            End If
        End If
    Next c
    ' 2) 同義群組
    Dim k As String: k = CsvAliasKey(name)
    If Len(k) = 0 Then Exit Function
    If k = "port" And Not MAP_PORT Then Exit Function   ' port 欄已知不準,預設不寫進去
    For c = 1 To lastCol
        If Not tUsed(c) Then
            If CsvAliasKey(tHdr(c)) = k Then CsvMatchTargetCol = c: Exit Function
        End If
    Next c
End Function

' 找合併結果裡的 IP 欄(只給「已存在提醒」用);-1 = 找不到
Private Function CsvFindIpColumn(ByRef cols() As String, ByVal nC As Long) As Long
    Dim j As Long
    CsvFindIpColumn = -1
    For j = 0 To nC - 1
        If CsvAliasKey(cols(j)) = "ip" Then CsvFindIpColumn = j: Exit Function
    Next j
End Function

' 避免來源檔欄名跟既有欄名撞名
Private Function CsvUniqueName(ByVal base As String, ByRef hdr() As String, ByVal nH As Long) As String
    Dim cand As String, k As Long, j As Long, hit As Boolean
    cand = base
    For k = 2 To 20
        hit = False
        For j = 0 To nH - 1
            If CsvNormHdr(hdr(j)) = CsvNormHdr(cand) Then hit = True: Exit For
        Next j
        If Not hit Then CsvUniqueName = cand: Exit Function
        cand = base & k
    Next k
    CsvUniqueName = cand
End Function

' =====================================================================
'  工作表小工具(全部 Private,不會跟 merge-assets.bas 撞名)
' =====================================================================
Private Function CsvExtractIP(ByVal s As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})"
        re.Global = False
    End If
    CsvExtractIP = ""
    If Len(s) = 0 Then Exit Function
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        Dim j As Long, ok As Boolean: ok = True
        For j = 0 To 3
            If CLng(m.SubMatches(j)) > 255 Then ok = False
        Next j
        If ok Then CsvExtractIP = m.Value
    End If
End Function

' 掃整本活頁簿既有的 IP(只讀不寫,純粹為了報告「這些 IP 你表裡本來就有」)
Private Function CsvCollectWorkbookIPs(wbk As Workbook) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, ur As Range, v As Variant
    Dim r As Long, c As Long, ip As String
    For Each ws In wbk.Worksheets
        Set ur = ws.UsedRange
        If Not ur Is Nothing Then
            If ur.Rows.Count * ur.Columns.Count <= 200000 Then     ' 太大就跳過,免得卡住
                v = ur.Value2
                If IsArray(v) Then
                    For r = LBound(v, 1) To UBound(v, 1)
                        For c = LBound(v, 2) To UBound(v, 2)
                            If VarType(v(r, c)) = vbString Then
                                ip = CsvExtractIP(CStr(v(r, c)))
                                If Len(ip) > 0 Then
                                    If Not d.Exists(ip) Then d.Add ip, ws.Name
                                End If
                            End If
                        Next c
                    Next r
                End If
            End If
        End If
    Next ws
    Set CsvCollectWorkbookIPs = d
End Function

Private Function CsvSheetLastRow(ws As Worksheet) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then
        CsvSheetLastRow = HEADER_ROW
    Else
        CsvSheetLastRow = ur.Row + ur.Rows.Count - 1
    End If
End Function

Private Function CsvSheetLastCol(ws As Worksheet) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then
        CsvSheetLastCol = 0
    Else
        CsvSheetLastCol = ur.Column + ur.Columns.Count - 1
    End If
End Function

Private Function CsvResolveSheet(wbk As Workbook, ByVal id As String) As Worksheet
    Dim ws As Worksheet
    If IsNumeric(id) Then
        Dim idx As Long: idx = CLng(id)
        If idx >= 1 And idx <= wbk.Worksheets.Count Then Set CsvResolveSheet = wbk.Worksheets(idx): Exit Function
    End If
    For Each ws In wbk.Worksheets
        If Trim$(ws.Name) = Trim$(id) Or LCase$(Trim$(ws.Name)) = LCase$(Trim$(id)) Then
            Set CsvResolveSheet = ws: Exit Function
        End If
    Next ws
End Function

Private Function CsvListSheets(wbk As Workbook) As String
    Dim ws As Worksheet, i As Long, o As String
    i = 0
    For Each ws In wbk.Worksheets
        i = i + 1: o = o & "  [" & i & "] " & ws.Name & vbCrLf
    Next ws
    CsvListSheets = o
End Function

Private Function CsvBaseName(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, "\")
    If k > 0 Then CsvBaseName = Mid$(p, k + 1) Else CsvBaseName = p
End Function

Private Function CsvTrimList(ByVal s As String) As String
    Dim a() As String: a = Split(s, vbCrLf)
    If UBound(a) <= 25 Then CsvTrimList = s: Exit Function
    Dim i As Long, o As String
    For i = 0 To 24: o = o & a(i) & vbCrLf: Next i
    CsvTrimList = o & "  …(共 " & UBound(a) & " 筆)" & vbCrLf
End Function
