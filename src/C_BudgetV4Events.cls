VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_BudgetV4Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ExcelApp As Application
Attribute ExcelApp.VB_VarHelpID = -1

Private Sub ExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo HandlerFailed
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    If Sh.Name <> "Budget_Detail" Then Exit Sub
    BudgetV4_OnDetailChange Target
    Exit Sub
HandlerFailed:
    MsgBox "Budget V4 event error (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub
