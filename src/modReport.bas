Attribute VB_Name = "modReport"
'==============================================================================
' modReport : 明細シートとレポートシート（帳票）の書き出し
'==============================================================================
Option Explicit

Private Const COLOR_HEAD As Long = 7949855      ' RGB(31, 78, 121)
Private Const COLOR_TOTAL As Long = 15921906    ' RGB(242, 242, 242)
Private Const FMT_AMOUNT As String = "#,##0"
Private Const FMT_SHARE As String = "0.0%"

'------------------------------------------------------------------------------
' 明細シート：全行を配列に詰めて1回で書き込む（セル単位の書き込みは件数に比例して遅くなるため）
'------------------------------------------------------------------------------
Public Sub WriteDetailSheet(ByVal saleRows As Collection)
    Dim ws As Worksheet
    Dim arr() As Variant
    Dim rec As clsSalesRow
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_DETAIL)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Meiryo UI"
    ws.Cells.Font.Size = 10

    ws.Range("A1:H1").Value = Array("日付", "年月", "顧客名", "商品名", "数量", "単価", "金額", "取込元ファイル")
    StyleHeader ws.Range("A1:H1")

    ReDim arr(1 To saleRows.Count, 1 To 8)
    For Each rec In saleRows
        i = i + 1
        arr(i, 1) = rec.SaleDate
        arr(i, 2) = rec.MonthKey
        arr(i, 3) = rec.Customer
        arr(i, 4) = rec.Product
        arr(i, 5) = rec.Qty
        arr(i, 6) = rec.UnitPrice
        arr(i, 7) = rec.Amount
        arr(i, 8) = rec.SourceFile
    Next rec

    ' 「2026-04」を日付に自動変換させないよう、書き込む前に文字列書式にしておく
    ws.Range("B2").Resize(saleRows.Count, 1).NumberFormat = "@"
    ws.Range("A2").Resize(saleRows.Count, 8).Value = arr
    ws.Range("A2").Resize(saleRows.Count, 1).NumberFormat = "yyyy/mm/dd"
    ws.Range("E2").Resize(saleRows.Count, 3).NumberFormat = FMT_AMOUNT

    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E:G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 20
    ws.Range("A1").Resize(saleRows.Count + 1, 8).AutoFilter
End Sub

'------------------------------------------------------------------------------
' レポートシート（帳票）
'------------------------------------------------------------------------------
Public Sub BuildReport(ByVal agg As Object, ByRef cfg As ReportConfig, ByVal rowCount As Long, ByVal errCount As Long)
    Dim ws As Worksheet
    Dim ms As Variant, custs As Variant, prods As Variant
    Dim nM As Long, nC As Long, nP As Long
    Dim totalCol As Long, shareCol As Long, lastCol As Long
    Dim mHead As Long, mFirst As Long, mLast As Long, mTotal As Long
    Dim pHead As Long, pFirst As Long, pLast As Long, pTotal As Long
    Dim vals() As Variant
    Dim i As Long, c As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT)
    ws.Cells.FormatConditions.Delete
    ws.Cells.Clear
    ws.Cells.Font.Name = "Meiryo UI"
    ws.Cells.Font.Size = 10

    ms = SortedMonths(agg)
    custs = KeysByValueDesc(agg("custTotal"))
    prods = KeysByValueDesc(agg("prodAmt"))
    nM = UBound(ms) + 1
    nC = UBound(custs) + 1
    nP = UBound(prods) + 1

    ' 月の数で列数が変わるため、行と列の位置を先に計算してから書き込む
    totalCol = 2 + nM
    shareCol = totalCol + 1
    lastCol = shareCol
    If lastCol < 4 Then lastCol = 4
    mHead = 12
    mFirst = mHead + 1
    mLast = mFirst + nC - 1
    mTotal = mLast + 1
    pHead = mTotal + 4
    pFirst = pHead + 1
    pLast = pFirst + nP - 1
    pTotal = pLast + 1

    ' --- 表題 ---
    ws.Range("A1").Value = "売上集計レポート"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    If nM = 1 Then
        ws.Range("A2").Value = "対象期間：" & ms(0)
    Else
        ws.Range("A2").Value = "対象期間：" & ms(0) & " ～ " & ms(nM - 1)
    End If
    ws.Cells(2, lastCol).Value = "作成日時：" & Format$(Now, "yyyy/mm/dd hh:nn")
    ws.Cells(2, lastCol).HorizontalAlignment = xlRight
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_HEAD
    End With

    ' --- 集計サマリー ---
    ' 合計と消費税は数式で残し、帳票の上で検算できるようにする
    ws.Range("A4").Value = "■ 集計サマリー"
    ws.Range("A4").Font.Bold = True
    ws.Range("A5").Value = "売上合計（税抜）"
    ws.Range("B5").FormulaR1C1 = "=R" & mTotal & "C" & totalCol
    ws.Range("A6").Value = "消費税"
    ' 税率は帳票側にも値で残す。設定を後で変えても、作成済みの帳票の数字が変わらないようにするため
    ws.Range("C6").Value = cfg.TaxRate
    ws.Range("C6").NumberFormat = """税率 ""0%"
    ws.Range("B6").FormulaR1C1 = "=ROUNDDOWN(R5C2*R6C3,0)"
    ws.Range("A7").Value = "売上合計（税込）"
    ws.Range("B7").FormulaR1C1 = "=R5C2+R6C2"
    ws.Range("A8").Value = "取込明細件数"
    ws.Range("B8").Value = rowCount
    ws.Range("A9").Value = "エラー件数"
    ws.Range("B9").Value = errCount
    ws.Range("B5:B7").NumberFormat = FMT_AMOUNT
    ws.Range("B8:B9").NumberFormat = "#,##0"" 件"""
    ws.Range("A7:B7").Font.Bold = True
    ws.Range("A5:A9").Interior.Color = COLOR_TOTAL
    ws.Range("A5:B9").Borders.LineStyle = xlContinuous
    ws.Range("A5:B9").Borders.Color = RGB(166, 166, 166)
    ' 除外した行があることを見落とさないよう、帳票にも目立つ形で残す（ADR-0002）
    If errCount > 0 Then
        ws.Range("B9").Font.Color = vbRed
        ws.Range("B9").Font.Bold = True
        ws.Range("A10").Value = "※ 取り込めなかった行は「エラー一覧」シートを参照してください。"
        ws.Range("A10").Font.Color = vbRed
        ws.Range("A10").Font.Size = 9
    End If

    ' --- 顧客別・月別 ---
    ws.Cells(mHead - 1, 1).Value = "■ 顧客別・月別売上（税抜・円）"
    ws.Cells(mHead - 1, 1).Font.Bold = True
    ws.Cells(mHead, 1).Value = "顧客名"
    ws.Range(ws.Cells(mHead, 2), ws.Cells(mHead, 1 + nM)).NumberFormat = "@"
    For c = 0 To nM - 1
        ws.Cells(mHead, 2 + c).Value = ms(c)
    Next c
    ws.Cells(mHead, totalCol).Value = "合計"
    ws.Cells(mHead, shareCol).Value = "構成比"
    StyleHeader ws.Range(ws.Cells(mHead, 1), ws.Cells(mHead, shareCol))

    ReDim vals(1 To nC, 1 To 1 + nM)
    For i = 0 To nC - 1
        vals(i + 1, 1) = custs(i)
        For c = 0 To nM - 1
            vals(i + 1, 2 + c) = CustomerMonthAmount(agg, CStr(custs(i)), CStr(ms(c)))
        Next c
    Next i
    ws.Cells(mFirst, 1).Resize(nC, 1 + nM).Value = vals

    ' R1C1 形式で書き、行ごとに同じ数式を一度に入れる（A1 形式の番地組み立てによる誤りを避ける）
    ws.Range(ws.Cells(mFirst, totalCol), ws.Cells(mLast, totalCol)).FormulaR1C1 = "=SUM(RC2:RC" & (totalCol - 1) & ")"
    ws.Range(ws.Cells(mFirst, shareCol), ws.Cells(mLast, shareCol)).FormulaR1C1 = _
        "=IF(R" & mTotal & "C" & totalCol & "=0,0,RC" & totalCol & "/R" & mTotal & "C" & totalCol & ")"
    ws.Cells(mTotal, 1).Value = "合計"
    ws.Range(ws.Cells(mTotal, 2), ws.Cells(mTotal, shareCol)).FormulaR1C1 = "=SUM(R" & mFirst & "C:R" & mLast & "C)"
    StyleTable ws.Range(ws.Cells(mFirst, 1), ws.Cells(mTotal, shareCol))
    StyleTotal ws.Range(ws.Cells(mTotal, 1), ws.Cells(mTotal, shareCol))
    ws.Range(ws.Cells(mFirst, 2), ws.Cells(mTotal, totalCol)).NumberFormat = FMT_AMOUNT
    ws.Range(ws.Cells(mFirst, shareCol), ws.Cells(mTotal, shareCol)).NumberFormat = FMT_SHARE
    ws.Cells(mTotal + 1, 1).Value = "※ 色が濃いほど売上が大きく、赤字は前月より減少した月を示します。"
    ws.Cells(mTotal + 1, 1).Font.Size = 9
    ws.Cells(mTotal + 1, 1).Font.Color = RGB(89, 89, 89)
    AddMatrixFormats ws, mFirst, mLast, nM

    ' --- 商品別 ---
    ws.Cells(pHead - 1, 1).Value = "■ 商品別売上（税抜・円）"
    ws.Cells(pHead - 1, 1).Font.Bold = True
    ws.Range(ws.Cells(pHead, 1), ws.Cells(pHead, 4)).Value = Array("商品名", "数量", "売上金額", "構成比")
    StyleHeader ws.Range(ws.Cells(pHead, 1), ws.Cells(pHead, 4))

    ReDim vals(1 To nP, 1 To 3)
    For i = 0 To nP - 1
        vals(i + 1, 1) = prods(i)
        vals(i + 1, 2) = agg("prodQty")(prods(i))
        vals(i + 1, 3) = agg("prodAmt")(prods(i))
    Next i
    ws.Cells(pFirst, 1).Resize(nP, 3).Value = vals
    ws.Range(ws.Cells(pFirst, 4), ws.Cells(pLast, 4)).FormulaR1C1 = "=IF(R" & pTotal & "C3=0,0,RC3/R" & pTotal & "C3)"
    ws.Cells(pTotal, 1).Value = "合計"
    ws.Range(ws.Cells(pTotal, 2), ws.Cells(pTotal, 4)).FormulaR1C1 = "=SUM(R" & pFirst & "C:R" & pLast & "C)"
    StyleTable ws.Range(ws.Cells(pFirst, 1), ws.Cells(pTotal, 4))
    StyleTotal ws.Range(ws.Cells(pTotal, 1), ws.Cells(pTotal, 4))
    ws.Range(ws.Cells(pFirst, 2), ws.Cells(pTotal, 3)).NumberFormat = FMT_AMOUNT
    ws.Range(ws.Cells(pFirst, 4), ws.Cells(pTotal, 4)).NumberFormat = FMT_SHARE
    With ws.Range(ws.Cells(pFirst, 3), ws.Cells(pLast, 3)).FormatConditions.AddDatabar
        .BarColor.Color = RGB(99, 142, 198)
    End With

    ' --- 列幅 ---
    ws.Columns(1).ColumnWidth = 30
    ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).EntireColumn.ColumnWidth = 13
    ws.Calculate
End Sub

' 色の濃淡で大きさを、赤字で前月からの減少を示し、数字を読まずに傾向をつかめるようにする
Private Sub AddMatrixFormats(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal nM As Long)
    Dim rng As Range
    Dim cs As ColorScale
    Dim fc As FormatCondition
    Dim prevStyle As XlReferenceStyle

    Set rng = ws.Range(ws.Cells(firstRow, 2), ws.Cells(lastRow, 1 + nM))
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=2)
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(155, 194, 230)

    If nM < 2 Then Exit Sub
    ' A1 形式の相対参照は「選択中のセル」を起点に解釈されるため、起点に依存しない R1C1 形式で指定する
    prevStyle = Application.ReferenceStyle
    On Error GoTo Restore
    Application.ReferenceStyle = xlR1C1
    Set fc = ws.Range(ws.Cells(firstRow, 3), ws.Cells(lastRow, 1 + nM)).FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=RC<RC[-1]")
    fc.Font.Color = RGB(192, 0, 0)
    fc.Font.Bold = True
Restore:
    Application.ReferenceStyle = prevStyle
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = COLOR_HEAD
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub StyleTable(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(166, 166, 166)
    End With
End Sub

Private Sub StyleTotal(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = COLOR_TOTAL
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Color = RGB(89, 89, 89)
    End With
End Sub
