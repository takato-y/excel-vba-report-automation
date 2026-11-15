Attribute VB_Name = "modAggregate"
'==============================================================================
' modAggregate : 顧客別・月別・商品別の集計
'==============================================================================
Option Explicit

' 集計結果は Dictionary の入れ子で返す。帳票側はキー名だけを知っていればよく、
' 集計項目を増やしても呼び出し側の引数を変えずに済む
'   "months"    : 年月 → True
'   "custTotal" : 顧客 → 金額合計
'   "custMonth" : 顧客 & vbTab & 年月 → 金額
'   "prodQty"   : 商品 → 数量合計
'   "prodAmt"   : 商品 → 金額合計
'   "grand"     : 全体の金額合計（Double）
Public Function Aggregate(ByVal saleRows As Collection) As Object
    Dim agg As Object
    Dim months As Object, custTotal As Object, custMonth As Object
    Dim prodQty As Object, prodAmt As Object
    Dim rec As clsSalesRow
    Dim key As String
    Dim grand As Double

    Set months = NewDict()
    Set custTotal = NewDict()
    Set custMonth = NewDict()
    Set prodQty = NewDict()
    Set prodAmt = NewDict()

    For Each rec In saleRows
        months(rec.MonthKey) = True
        AddTo custTotal, rec.Customer, rec.Amount
        key = rec.Customer & vbTab & rec.MonthKey
        AddTo custMonth, key, rec.Amount
        AddTo prodQty, rec.Product, rec.Qty
        AddTo prodAmt, rec.Product, rec.Amount
        grand = grand + rec.Amount
    Next rec

    Set agg = NewDict()
    Set agg("months") = months
    Set agg("custTotal") = custTotal
    Set agg("custMonth") = custMonth
    Set agg("prodQty") = prodQty
    Set agg("prodAmt") = prodAmt
    agg("grand") = grand
    Set Aggregate = agg
End Function

' 年月のキーを昇順で返す（yyyy-mm 形式なので文字列順＝時系列順）
Public Function SortedMonths(ByVal agg As Object) As Variant
    Dim ks As Variant
    ks = agg("months").Keys
    SortStringsAsc ks
    SortedMonths = ks
End Function

Public Function CustomerMonthAmount(ByVal agg As Object, ByVal cust As String, ByVal mk As String) As Double
    Dim key As String
    key = cust & vbTab & mk
    If agg("custMonth").Exists(key) Then CustomerMonthAmount = agg("custMonth")(key)
End Function

' 全角半角・大文字小文字の違いを勝手に同一視しないよう完全一致で比較する
' （表記ゆれは帳票で別の行として見えるため、元データの修正に気づける）
Private Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    Set NewDict = d
End Function

Private Sub AddTo(ByVal dict As Object, ByVal key As String, ByVal addValue As Double)
    If dict.Exists(key) Then
        dict(key) = dict(key) + addValue
    Else
        dict.Add key, addValue
    End If
End Sub
