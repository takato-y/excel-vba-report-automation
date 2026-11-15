Attribute VB_Name = "modConfig"
'==============================================================================
' modConfig : 設定シートの読み込みと検証
'==============================================================================
Option Explicit

Public Type ReportConfig
    InputFolder As String
    OutputFolder As String
    Charset As String
    TaxRate As Double
    ExportPdf As Boolean
End Type

' 処理を始める前に全項目を検査し、途中まで進んでから設定の誤りで止まることを避ける
Public Function LoadConfig() As ReportConfig
    Dim c As ReportConfig
    Dim v As Variant
    Dim s As String

    c.InputFolder = ResolvePath(CStr(NameRange(NAME_INPUT).Value), "sample_data")
    c.OutputFolder = ResolvePath(CStr(NameRange(NAME_OUTPUT).Value), "output")

    s = UCase$(Trim$(CStr(NameRange(NAME_CHARSET).Value)))
    If s = "" Then s = "UTF-8"
    If s <> "UTF-8" And s <> "SHIFT_JIS" Then
        Err.Raise ERR_CONFIG, , "「CSV文字コード」は UTF-8 または Shift_JIS を選んでください。（現在の値：" & s & "）"
    End If
    c.Charset = s

    v = NameRange(NAME_TAX).Value
    If IsEmpty(v) Or Not IsNumeric(v) Then
        Err.Raise ERR_CONFIG, , "「消費税率」は 10% のように数値で入力してください。"
    End If
    If CDbl(v) < 0 Or CDbl(v) >= 1 Then
        Err.Raise ERR_CONFIG, , "「消費税率」は 0% 以上 100% 未満で入力してください。"
    End If
    c.TaxRate = CDbl(v)

    s = Trim$(CStr(NameRange(NAME_PDF).Value))
    Select Case s
        Case "する": c.ExportPdf = True
        Case "しない": c.ExportPdf = False
        Case Else
            Err.Raise ERR_CONFIG, , "「PDF出力」は「する」または「しない」を選んでください。"
    End Select

    If Not FolderExists(c.InputFolder) Then
        Err.Raise ERR_FOLDER, , "入力フォルダが見つかりません。" & vbCrLf & c.InputFolder & vbCrLf & _
                                "設定シートの「入力フォルダ」を確認してください。"
    End If

    LoadConfig = c
End Function

' 相対パスはブックの場所を起点にする。ブックごとフォルダを移しても設定を直さずに済むようにするため
Public Function ResolvePath(ByVal p As String, ByVal defaultValue As String) As String
    Dim basePath As String
    p = Trim$(p)
    If p = "" Then p = defaultValue
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    If p Like "[A-Za-z]:*" Or Left$(p, 2) = "\\" Then
        ResolvePath = p
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If basePath = "" Then
        Err.Raise ERR_CONFIG, , "ブックが保存されていないため、相対パスを解決できません。先にブックを保存してください。"
    End If
    ' OneDrive の同期フォルダでは Path が URL になり、フォルダ操作が失敗するため先に止める
    If LCase$(Left$(basePath, 4)) = "http" Then
        Err.Raise ERR_CONFIG, , "ブックがクラウド上の場所にあるため、相対パスを使えません。" & vbCrLf & _
                                "設定シートに C:\ から始まるフォルダを入力してください。"
    End If
    ResolvePath = basePath & "\" & p
End Function

' Dir は隠し属性のフォルダを見落とすことがあるため、属性を直接調べる
Public Function FolderExists(ByVal p As String) As Boolean
    Dim attr As Long
    On Error GoTo NotFound
    attr = GetAttr(p)
    FolderExists = ((attr And vbDirectory) = vbDirectory)
    Exit Function
NotFound:
    FolderExists = False
End Function
