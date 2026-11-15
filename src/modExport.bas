Attribute VB_Name = "modExport"
'==============================================================================
' modExport : 帳票の印刷設定と PDF 出力
'==============================================================================
Option Explicit

' 保存した PDF のフルパスを返す
Public Function ExportReportPdf(ByRef cfg As ReportConfig) As String
    Dim ws As Worksheet
    Dim pdfPath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_REPORT)
    EnsureFolder cfg.OutputFolder
    ' 日時を名前に入れ、前回の PDF を上書きしない（過去の報告書を残すため）
    pdfPath = cfg.OutputFolder & "\売上集計レポート_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    ApplyPageSetup ws

    On Error GoTo ExportFail
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    If Dir(pdfPath) = "" Then
        Err.Raise ERR_PDF, , "PDF の作成を確認できませんでした。" & vbCrLf & pdfPath
    End If
    ExportReportPdf = pdfPath
    Exit Function

ExportFail:
    Err.Raise ERR_PDF, , "PDF を保存できませんでした。出力フォルダに書き込めるか、" & _
                         "同じ名前の PDF を開いていないか確認してください。" & vbCrLf & _
                         "（" & Err.Description & "）"
End Function

' 月の数が増えても1ページの幅に収まるよう、横は1ページに縮小し、縦は成り行きにする
Private Sub ApplyPageSetup(ByVal ws As Worksheet)
    Dim lastCell As Range
    With ws.UsedRange
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With

    On Error GoTo SetupFail
    ' プリンターとの通信をまとめて1回にし、設定のたびに待たされるのを避ける
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Range("A1"), lastCell).Address
        If lastCell.Column > 8 Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterFooter = "&P / &N"
    End With
    Application.PrintCommunication = True
    Exit Sub

SetupFail:
    Application.PrintCommunication = True
    Err.Raise ERR_PDF, , "印刷設定を反映できませんでした。Windows の「既定のプリンター」が設定されているか確認してください。" & _
                         vbCrLf & "（" & Err.Description & "）"
End Sub

' 出力フォルダが無ければ作る。親フォルダから作るのは誤入力で意図しない階層を作らないため、1階層のみとする
Private Sub EnsureFolder(ByVal p As String)
    If FolderExists(p) Then Exit Sub
    On Error GoTo MakeFail
    MkDir p
    Exit Sub
MakeFail:
    Err.Raise ERR_PDF, , "出力フォルダを作成できませんでした。上の階層のフォルダが存在するか確認してください。" & _
                         vbCrLf & p
End Sub
