VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsSalesRow : 検査を通った明細1行
'   配列の列番号で扱うと取り違えが起きやすいため、項目名で参照できるようにする
'==============================================================================
Option Explicit

Public SaleDate As Date
Public MonthKey As String
Public Customer As String
Public Product As String
Public Qty As Double
Public UnitPrice As Double
Public Amount As Double
Public SourceFile As String
