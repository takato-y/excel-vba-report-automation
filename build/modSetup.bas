Attribute VB_Name = "modSetup"
'==============================================================================
' modSetup : ブック作成時だけ使う初期化処理（完成したブックからは取り除く）
'==============================================================================
Option Explicit

Public Sub SetupWorkbook()
    Dim sheetNames As Variant
    Dim i As Long
    Dim ws As Worksheet

    sheetNames = Array(SHEET_SETTINGS, SHEET_REPORT, SHEET_DETAIL, SHEET_ERRORS)
    For i = 0 To UBound(sheetNames)
        If i < ThisWorkbook.Worksheets.Count Then
            Set ws = ThisWorkbook.Worksheets(i + 1)
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        End If
        ws.Name = sheetNames(i)
        ws.Cells.Font.Name = "Meiryo UI"
        ws.Cells.Font.Size = 10
    Next i

    BuildSettingsSheet ThisWorkbook.Worksheets(SHEET_SETTINGS)

    ThisWorkbook.Worksheets(SHEET_REPORT).Range("A1").Value = "「設定」シートの［レポート作成］ボタンを押すと、ここに帳票が作成されます。"
    ThisWorkbook.Worksheets(SHEET_DETAIL).Range("A1").Value = "取り込んだ明細がここに一覧で表示されます。"
    ClearErrorLog

    ThisWorkbook.Worksheets(SHEET_SETTINGS).Activate
End Sub

Private Sub BuildSettingsSheet(ByVal ws As Worksheet)
    Dim btn As Object
    Dim r As Long

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 58

    ws.Range("A1").Value = "売上集計レポート作成ツール"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "入力フォルダ内の売上ファイル（CSV / Excel）を取りまとめ、顧客別・商品別・月別の帳票と PDF を作成します。"

    ws.Range("A4:C4").Value = Array("項目", "設定値", "説明")
    With ws.Range("A4:C4")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
    End With

    ws.Range("A5:C5").Value = Array("入力フォルダ", "sample_data", "このブックからの相対パス、または C:\ から始まるフォルダ")
    ws.Range("A6:C6").Value = Array("出力フォルダ", "output", "PDF の保存先。存在しない場合は作成します")
    ws.Range("A7:C7").Value = Array("CSV文字コード", "UTF-8", "UTF-8 / Shift_JIS から選択")
    ws.Range("A8:C8").Value = Array("消費税率", 0.1, "10% のように入力（端数は切り捨て）")
    ws.Range("A9:C9").Value = Array("PDF出力", "する", "する / しない から選択")
    ws.Range("B8").NumberFormat = "0%"
    ws.Range("B5:B9").Interior.Color = RGB(255, 251, 230)
    ws.Range("B5:B9").HorizontalAlignment = xlLeft
    ws.Range("A4:C9").Borders.LineStyle = xlContinuous
    ws.Range("A4:C9").Borders.Color = RGB(166, 166, 166)

    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="UTF-8,Shift_JIS"
    End With
    With ws.Range("B9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="する,しない"
    End With
    With ws.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="0.99"
        .ErrorMessage = "0% 以上 100% 未満で入力してください。"
    End With

    ws.Range("A11").Value = "前回の実行結果"
    ws.Range("A11").Font.Bold = True
    ws.Range("A11").VerticalAlignment = xlTop
    ws.Range("B11:C11").Merge
    ws.Range("B11").WrapText = True
    ws.Rows(11).RowHeight = 30
    ws.Range("B11").VerticalAlignment = xlTop

    ThisWorkbook.Names.Add Name:=NAME_INPUT, RefersTo:="='" & ws.Name & "'!$B$5"
    ThisWorkbook.Names.Add Name:=NAME_OUTPUT, RefersTo:="='" & ws.Name & "'!$B$6"
    ThisWorkbook.Names.Add Name:=NAME_CHARSET, RefersTo:="='" & ws.Name & "'!$B$7"
    ThisWorkbook.Names.Add Name:=NAME_TAX, RefersTo:="='" & ws.Name & "'!$B$8"
    ThisWorkbook.Names.Add Name:=NAME_PDF, RefersTo:="='" & ws.Name & "'!$B$9"
    ThisWorkbook.Names.Add Name:=NAME_STATUS, RefersTo:="='" & ws.Name & "'!$B$11"

    r = 13
    Set btn = ws.Buttons.Add(ws.Range("A" & r).Left, ws.Range("A" & r).Top, 180, 34)
    btn.OnAction = "RunReport"
    btn.Caption = "レポート作成"
    btn.Font.Size = 12
    btn.Font.Bold = True
    btn.Name = "btnRunReport"

    ws.Range("A17").Value = "使い方"
    ws.Range("A17").Font.Bold = True
    ws.Range("A18").Value = "1. 入力フォルダに月別の売上ファイルを置きます（1行目の見出し：日付, 顧客名, 商品名, 数量, 単価, 金額）。"
    ws.Range("A19").Value = "2. ［レポート作成］ボタンを押します。"
    ws.Range("A20").Value = "3. 「レポート」に帳票、「明細」に取りまとめたデータ、「エラー一覧」に取り込めなかった行が表示されます。"
    ws.Range("A21").Value = "4. PDF出力が「する」の場合、出力フォルダに PDF が保存されます。"
End Sub
