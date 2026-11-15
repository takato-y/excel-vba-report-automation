Attribute VB_Name = "modImport"
'==============================================================================
' modImport : 入力フォルダのファイル読み込みと行ごとの検査
'==============================================================================
Option Explicit

Private Const COL_COUNT As Long = 6

Private Function ExpectedHeader() As Variant
    ExpectedHeader = Array("日付", "顧客名", "商品名", "数量", "単価", "金額")
End Function

' 正しい行だけを clsSalesRow の Collection で返す。不正な行はエラー一覧に記録して除外する
Public Function ImportFolder(ByRef cfg As ReportConfig) As Collection
    Dim files As Variant
    Dim result As New Collection
    Dim i As Long

    files = ListInputFiles(cfg.InputFolder)
    If Not IsArray(files) Then
        Err.Raise ERR_NO_FILES, , "入力フォルダに CSV / Excel ファイルがありません。" & vbCrLf & cfg.InputFolder
    End If

    For i = LBound(files) To UBound(files)
        ImportOneFile cfg.InputFolder & "\" & files(i), cfg.Charset, result
    Next i
    Set ImportFolder = result
End Function

' Dir の返す順番はファイルシステム任せのため、名前順に並べて結果を毎回同じにする
Private Function ListInputFiles(ByVal folder As String) As Variant
    Dim nm As String, ext As String
    Dim fileNames() As String
    Dim n As Long

    nm = Dir(folder & "\*.*")
    Do While nm <> ""
        ext = LCase$(Mid$(nm, InStrRev(nm, ".") + 1))
        If Left$(nm, 2) <> "~$" And (ext = "csv" Or ext = "xlsx" Or ext = "xls") Then
            ReDim Preserve fileNames(n)
            fileNames(n) = nm
            n = n + 1
        End If
        nm = Dir()
    Loop

    If n = 0 Then
        ListInputFiles = Empty
    Else
        SortStringsAsc fileNames
        ListInputFiles = fileNames
    End If
End Function

Private Sub ImportOneFile(ByVal filePath As String, ByVal csName As String, ByVal result As Collection)
    Dim data As Variant
    Dim fileName As String
    Dim r As Long, c As Long
    Dim hdr As Variant
    Dim msg As String
    Dim rec As clsSalesRow

    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)

    ' 読み込めないファイルは1件の失敗として記録し、他のファイルの処理は続ける
    On Error GoTo ReadFail
    If LCase$(Right$(filePath, 4)) = ".csv" Then
        data = ReadCsv(filePath, csName)
    Else
        data = ReadExcel(filePath)
    End If
    On Error GoTo 0

    If Not IsArray(data) Then
        LogError fileName, 0, "ファイルが空です。"
        Exit Sub
    End If
    If UBound(data, 2) < COL_COUNT Then
        LogError fileName, 1, "列の数が足りません（6列必要）。このファイルは取り込みませんでした。"
        Exit Sub
    End If

    ' 見出しが違うファイルは列の意味を取り違える恐れがあるため、丸ごと除外する
    hdr = ExpectedHeader()
    For c = 1 To COL_COUNT
        If Trim$(CellText(data(1, c))) <> hdr(c - 1) Then
            LogError fileName, 1, "見出し行が想定と異なるため、このファイルは取り込みませんでした。" & _
                                  "（想定：日付, 顧客名, 商品名, 数量, 単価, 金額）"
            Exit Sub
        End If
    Next c

    For r = 2 To UBound(data, 1)
        If Not IsBlankRow(data, r) Then
            Set rec = New clsSalesRow
            msg = ValidateRow(data, r, rec)
            If msg = "" Then
                rec.SourceFile = fileName
                result.Add rec
            Else
                LogError fileName, r, msg
            End If
        End If
    Next r
    Exit Sub

ReadFail:
    LogError fileName, 0, "ファイルを読み込めませんでした。他のアプリで開いていないか確認してください。（" & Err.Description & "）"
End Sub

Private Function IsBlankRow(ByRef data As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To COL_COUNT
        If Len(Trim$(CellText(data(r, c)))) > 0 Then
            IsBlankRow = False
            Exit Function
        End If
    Next c
    IsBlankRow = True
End Function

' 検査に通れば rec に値を詰めて空文字を返す。通らなければ理由を返す
Private Function ValidateRow(ByRef data As Variant, ByVal r As Long, ByVal rec As clsSalesRow) As String
    Dim v As Variant
    Dim d As Date

    v = data(r, 1)
    If VarType(v) = vbDate Then
        d = v
    ElseIf Len(Trim$(CellText(v))) > 0 And IsDate(v) Then
        d = CDate(v)
    Else
        ValidateRow = "日付の形式が正しくありません（値：" & CellText(v) & "）"
        Exit Function
    End If

    If Len(Trim$(CellText(data(r, 2)))) = 0 Then
        ValidateRow = "顧客名が空欄です"
        Exit Function
    End If
    If Len(Trim$(CellText(data(r, 3)))) = 0 Then
        ValidateRow = "商品名が空欄です"
        Exit Function
    End If

    ValidateRow = CheckNumber(data(r, 4), "数量")
    If ValidateRow <> "" Then Exit Function
    ValidateRow = CheckNumber(data(r, 5), "単価")
    If ValidateRow <> "" Then Exit Function
    ValidateRow = CheckNumber(data(r, 6), "金額")
    If ValidateRow <> "" Then Exit Function

    rec.SaleDate = d
    rec.MonthKey = Format$(d, "yyyy-mm")
    rec.Customer = Trim$(CellText(data(r, 2)))
    rec.Product = Trim$(CellText(data(r, 3)))
    rec.Qty = CDbl(data(r, 4))
    rec.UnitPrice = CDbl(data(r, 5))
    rec.Amount = CDbl(data(r, 6))

    ' 金額の手入力ミスや列ずれは集計を静かに狂わせるため、数量×単価との一致を確かめる
    If Abs(rec.Qty * rec.UnitPrice - rec.Amount) > 0.5 Then
        ValidateRow = "金額が 数量×単価 と一致しません（数量×単価：" & Format$(rec.Qty * rec.UnitPrice, "#,##0") & _
                      "、金額：" & Format$(rec.Amount, "#,##0") & "）"
    End If
End Function

' IsNumeric は空欄を数値とみなすため、空欄を先に判定する
Private Function CheckNumber(ByVal v As Variant, ByVal label As String) As String
    If Len(Trim$(CellText(v))) = 0 Then
        CheckNumber = label & "が空欄です"
    ElseIf Not IsNumeric(v) Then
        CheckNumber = label & "が数値ではありません（値：" & CellText(v) & "）"
    Else
        CheckNumber = ""
    End If
End Function

' Excel のエラー値（#N/A など）は CStr で実行時エラーになるため、文字に置き換えて検査を続ける
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "(エラー値)"
    Else
        CellText = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' CSV：文字コードを指定して全文を読み、値は文字列のまま受け取る（ADR-0001）
'------------------------------------------------------------------------------
Private Function ReadCsv(ByVal filePath As String, ByVal csName As String) As Variant
    Dim stm As Object
    Dim txt As String
    Dim lines As Variant
    Dim parsed() As Variant
    Dim n As Long, i As Long, c As Long, maxCols As Long
    Dim out() As Variant
    Dim f As Variant

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                        ' テキストとして読む
    If csName = "UTF-8" Then stm.Charset = "utf-8" Else stm.Charset = "shift_jis"
    stm.Open
    stm.LoadFromFile filePath
    txt = stm.ReadText(-1)
    stm.Close

    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)
    End If
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    Do While Right$(txt, 1) = vbLf
        txt = Left$(txt, Len(txt) - 1)
    Loop
    If Len(txt) = 0 Then
        ReadCsv = Empty
        Exit Function
    End If

    ' 行番号をファイルの行と一致させるため、空行も1行として残す
    lines = Split(txt, vbLf)
    n = UBound(lines) + 1
    ReDim parsed(1 To n)
    For i = 1 To n
        parsed(i) = ParseCsvLine(CStr(lines(i - 1)))
        If UBound(parsed(i)) + 1 > maxCols Then maxCols = UBound(parsed(i)) + 1
    Next i

    ReDim out(1 To n, 1 To maxCols)
    For i = 1 To n
        f = parsed(i)
        For c = 0 To UBound(f)
            out(i, c + 1) = f(c)
        Next c
    Next i
    ReadCsv = out
End Function

' ダブルクォートで囲まれた値の中のカンマと、"" による引用符のエスケープに対応する
Private Function ParseCsvLine(ByVal s As String) As Variant
    Dim fields() As String
    Dim n As Long, i As Long
    Dim ch As String, buf As String
    Dim inQuote As Boolean

    ReDim fields(0)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If inQuote Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                buf = buf & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuote = True
                Case ","
                    fields(n) = buf
                    buf = ""
                    n = n + 1
                    ReDim Preserve fields(n)
                Case Else
                    buf = buf & ch
            End Select
        End If
        i = i + 1
    Loop
    fields(n) = buf
    ParseCsvLine = fields
End Function

'------------------------------------------------------------------------------
' Excel：先頭シートの A1 から使用範囲の右下までを一括で読む
'------------------------------------------------------------------------------
Private Function ReadExcel(ByVal filePath As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCell As Range
    Dim v As Variant
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail
    Set wb = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(1)
    With ws.UsedRange
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    v = ws.Range(ws.Range("A1"), lastCell).Value
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If Not IsArray(v) Then
        If Len(Trim$(CellText(v))) = 0 Then
            ReadExcel = Empty
            Exit Function
        End If
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        v = one
    End If
    ReadExcel = v
    Exit Function

Fail:
    ' 開いたままにすると次回の実行で同名ブックの衝突が起きるため、必ず閉じてから上げ直す
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errNo, , errDesc
End Function
