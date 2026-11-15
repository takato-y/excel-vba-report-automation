Attribute VB_Name = "modUtil"
'==============================================================================
' modUtil : 定数・メッセージ表示・エラー一覧への記録
'==============================================================================
Option Explicit

' シート名はここだけで定義し、名前を変えるときの修正箇所を1か所にする
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_REPORT As String = "レポート"
Public Const SHEET_DETAIL As String = "明細"
Public Const SHEET_ERRORS As String = "エラー一覧"

' 設定値はセル番地ではなく名前で参照し、設定シートのレイアウト変更に強くする
Public Const NAME_INPUT As String = "cfg_InputFolder"
Public Const NAME_OUTPUT As String = "cfg_OutputFolder"
Public Const NAME_CHARSET As String = "cfg_Charset"
Public Const NAME_TAX As String = "cfg_TaxRate"
Public Const NAME_PDF As String = "cfg_ExportPdf"
Public Const NAME_STATUS As String = "cfg_Status"

' 想定内のエラーは独自の番号で上げ、予期しないエラーと表示を分ける
Public Const ERR_BASE As Long = vbObjectError + 1000
Public Const ERR_CONFIG As Long = ERR_BASE + 1
Public Const ERR_FOLDER As Long = ERR_BASE + 2
Public Const ERR_NO_FILES As Long = ERR_BASE + 3
Public Const ERR_NO_VALID_ROWS As Long = ERR_BASE + 4
Public Const ERR_PDF As Long = ERR_BASE + 5

Public Const STATUS_OK As String = "正常終了："
Public Const STATUS_NG As String = "エラー："

' 自動テストではメッセージボックスで処理が止まらないよう、表示を切り替える
Public gSilent As Boolean

Private mErrorCount As Long

Public Sub Notify(ByVal msg As String, ByVal style As VbMsgBoxStyle)
    If gSilent Then
        Debug.Print msg
    Else
        MsgBox msg, style, "売上集計レポート"
    End If
End Sub

Public Function NameRange(ByVal nm As String) As Range
    Set NameRange = ThisWorkbook.Names(nm).RefersToRange
End Function

Public Sub SetStatus(ByVal msg As String)
    NameRange(NAME_STATUS).Value = msg
End Sub

Public Function IsOwnError(ByVal errNo As Long) As Boolean
    IsOwnError = (errNo > ERR_BASE And errNo <= ERR_BASE + 100)
End Function

'------------------------------------------------------------------------------
' エラー一覧
'------------------------------------------------------------------------------
Public Sub ClearErrorLog()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_ERRORS)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("ファイル名", "行", "内容")
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 6
    ws.Columns("C").ColumnWidth = 80
    mErrorCount = 0
End Sub

' 行番号 0 はファイル全体に関するエラーを表す
Public Sub LogError(ByVal fileName As String, ByVal lineNo As Long, ByVal msg As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_ERRORS)
    mErrorCount = mErrorCount + 1
    ws.Cells(mErrorCount + 1, 1).Value = fileName
    If lineNo > 0 Then ws.Cells(mErrorCount + 1, 2).Value = lineNo Else ws.Cells(mErrorCount + 1, 2).Value = "-"
    ws.Cells(mErrorCount + 1, 3).Value = msg
End Sub

Public Function ErrorLogCount() As Long
    ErrorLogCount = mErrorCount
End Function

'------------------------------------------------------------------------------
' 並べ替え（件数が少ないため挿入ソートで十分。外部ライブラリに頼らない）
'------------------------------------------------------------------------------
Public Sub SortStringsAsc(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If StrComp(arr(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

' 値の大きい順。同額のときは名前順にして、実行のたびに並びが変わらないようにする
Public Function KeysByValueDesc(ByVal dict As Object) As Variant
    Dim ks As Variant, i As Long, j As Long, tmp As Variant
    ks = dict.Keys
    For i = LBound(ks) + 1 To UBound(ks)
        tmp = ks(i)
        j = i - 1
        Do While j >= LBound(ks)
            If dict(ks(j)) > dict(tmp) Then Exit Do
            If dict(ks(j)) = dict(tmp) And StrComp(ks(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
            ks(j + 1) = ks(j)
            j = j - 1
        Loop
        ks(j + 1) = tmp
    Next i
    KeysByValueDesc = ks
End Function
