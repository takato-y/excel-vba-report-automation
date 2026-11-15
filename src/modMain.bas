Attribute VB_Name = "modMain"
'==============================================================================
' modMain : ［レポート作成］ボタンから呼ばれる入口と、処理全体の流れ
'==============================================================================
Option Explicit

Public Sub RunReport()
    gSilent = False
    RunReportCore
End Sub

' 自動テスト用の入口。ボタンと同じ処理経路を通し、結果の文言を返す
Public Function RunReportForTest() As String
    gSilent = True
    RunReportCore
    gSilent = False
    RunReportForTest = CStr(NameRange(NAME_STATUS).Value)
End Function

Public Sub RunReportCore()
    Dim cfg As ReportConfig
    Dim saleRows As Collection
    Dim agg As Object
    Dim errCount As Long
    Dim msg As String
    Dim calcMode As XlCalculation
    Dim errNo As Long, errDesc As String

    calcMode = Application.Calculation
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    SetStatus "実行中..."

    cfg = LoadConfig()
    ClearErrorLog
    Set saleRows = ImportFolder(cfg)
    errCount = ErrorLogCount()
    If saleRows.Count = 0 Then
        Err.Raise ERR_NO_VALID_ROWS, , "取り込める明細が1件もありませんでした。「エラー一覧」シートで理由を確認してください。"
    End If

    ' 帳票は集計まで成功してから書き換える。途中で失敗したときに前回の帳票を壊さないため
    Set agg = Aggregate(saleRows)
    WriteDetailSheet saleRows
    BuildReport agg, cfg, saleRows.Count, errCount

    Application.Calculation = calcMode
    Application.ScreenUpdating = True

    msg = "レポートを作成しました。" & vbCrLf & _
          "取込明細：" & saleRows.Count & " 件　エラー：" & errCount & " 件"
    If errCount > 0 Then msg = msg & vbCrLf & "取り込めなかった行は「エラー一覧」シートを確認してください。"
    SetStatus STATUS_OK & Replace(msg, vbCrLf, " / ")
    If Not gSilent Then ThisWorkbook.Worksheets(SHEET_REPORT).Activate
    Notify msg, IIf(errCount > 0, vbExclamation, vbInformation)
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errDesc = Err.Description
    ' 失敗しても Excel の設定を元に戻し、画面が固まったように見える状態を残さない
    On Error Resume Next
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    SetStatus STATUS_NG & Replace(errDesc, vbCrLf, " ")
    Notify FriendlyMessage(errNo, errDesc), vbExclamation
End Sub

' 想定内のエラーはそのまま、想定外のエラーは番号と確認ポイントを添えて表示する
Private Function FriendlyMessage(ByVal errNo As Long, ByVal errDesc As String) As String
    If IsOwnError(errNo) Then
        FriendlyMessage = errDesc
    Else
        FriendlyMessage = "予期しないエラーが発生したため、処理を中止しました。" & vbCrLf & _
                          "（エラー番号 " & errNo & "：" & errDesc & "）" & vbCrLf & vbCrLf & _
                          "入力ファイルや出力先の PDF を開いたままにしていないか確認してください。"
    End If
End Function
