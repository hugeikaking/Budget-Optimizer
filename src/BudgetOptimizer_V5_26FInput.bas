Attribute VB_Name = "BudgetOptimizer_V5_26FInput"
Option Explicit

' User-entered 26F total baselines in an independent, persistent worksheet.
' Budget_26F_Input!B2: Sales Amt, B3: Set QTY.
' Budget_Overview!E5 and B5 link to these input cells.
' Only those 2 overview cells change. No trial, model edits, or source data
' are rebuilt or cleared. Requires V4's optional 26F E5 override to persist
' after a later RefreshBudgetSheet1Preview / EnsureOverview.

Private Const FY_INPUT As String = "Budget_26F_Input"
Private Const FY_OVERVIEW As String = "Budget_Overview"

Public Sub SetupBudgetV5_26FInput()
    Dim inp As Worksheet, ov As Worksheet, eventsBefore As Boolean
    Dim failureText As String
    eventsBefore = Application.EnableEvents
    On Error GoTo Failed
    Set ov = ThisWorkbook.Worksheets(FY_OVERVIEW)
    On Error Resume Next
    Set inp = ThisWorkbook.Worksheets(FY_INPUT)
    On Error GoTo Failed

    Application.EnableEvents = False
    If inp Is Nothing Then
        Set inp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        inp.Name = FY_INPUT
    End If
    With inp
        .Range("A1:B1").Value = Array("26F TTL BASELINE", "INPUT VALUE")
        .Range("A2").Value2 = "26F Sales Amt"
        .Range("A3").Value2 = "26F Set QTY"
        .Range("D2").Value2 = "Enter the 26F full-fiscal-year TOTAL (not a per-H40 figure)."
        .Range("D3").Value2 = "Use plain numeric values, not units like M/K or formulas with links."
        .Range("D4").Value2 = "Run ImportBudgetV5_26FTotals once to link Overview B5 and E5."
        .Range("D5").Value2 = "After linking, edits to these two cells update Overview automatically."
        .Range("D6").Value2 = "Existing values are preserved when setup is run again."
        .Range("A1:B1").Interior.Color = RGB(44, 74, 108)
        .Range("A1:B1").Font.Color = RGB(255, 255, 255)
        .Range("A1:B1").Font.Bold = True
        .Range("B2:B3").Interior.Color = RGB(255, 244, 193)
        .Range("B2:B3").NumberFormat = "#,##0;[Red]-#,##0;0"
        .Columns("A").ColumnWidth = 27
        .Columns("B").ColumnWidth = 24
        .Columns("D").ColumnWidth = 72
    End With

    Application.EnableEvents = eventsBefore
    MsgBox "26F input sheet ready. Enter 26F Sales Amt in B2 and Set QTY in B3.", vbInformation
    Exit Sub
Failed:
    failureText = "26F input setup failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = eventsBefore
    MsgBox failureText, vbExclamation
End Sub

Public Sub ImportBudgetV5_26FTotals()
    Dim inp As Worksheet, ov As Worksheet, amt As Variant, qty As Variant
    Dim eventsBefore As Boolean, failureText As String
    eventsBefore = Application.EnableEvents
    On Error GoTo Failed
    Set inp = ThisWorkbook.Worksheets(FY_INPUT)
    Set ov = ThisWorkbook.Worksheets(FY_OVERVIEW)

    amt = inp.Range("B2").Value2
    qty = inp.Range("B3").Value2
    If IsError(amt) Or IsError(qty) Then _
        Err.Raise vbObjectError + 801, , "26F input cells contain an Excel error."
    If Len(Trim$(CStr(amt))) = 0 Or Len(Trim$(CStr(qty))) = 0 Then _
        Err.Raise vbObjectError + 802, , "Enter both 26F Sales Amt (B2) and Set QTY (B3)."
    If Not IsNumeric(amt) Or Not IsNumeric(qty) Then _
        Err.Raise vbObjectError + 803, , "Both inputs must be numeric."
    If CDbl(amt) < 0 Then _
        Err.Raise vbObjectError + 804, , "26F Sales Amt must be nonnegative."
    If CDbl(qty) < 0 Or CDbl(qty) <> Fix(CDbl(qty)) Then _
        Err.Raise vbObjectError + 805, , "26F Set QTY must be a nonnegative integer."

    If MsgBox("Link Budget_Overview to these 26F totals?" & vbCrLf & _
              "26F Sales Amt: " & Format$(CDbl(amt), "#,##0") & vbCrLf & _
              "26F Set QTY: " & Format$(CDbl(qty), "#,##0") & vbCrLf & _
              "Only Overview E5 and B5 will change. Continue?", _
              vbQuestion + vbYesNo, "26F Baseline Import") <> vbYes Then Exit Sub

    Application.EnableEvents = False
    ' Links are intentional, so updated 26F input values flow to Overview
    ' without another import. The latest V4 code preserves the E5 link on
    ' refreshing Trial; legacy V4 may overwrite E5 and needs a repair import.
    ov.Range("E5").Formula = "='Budget_26F_Input'!$B$2"
    ov.Range("B5").Formula = "='Budget_26F_Input'!$B$3"
    ov.Range("E5").NumberFormat = "#,##0;[Red]-#,##0;0"
    ov.Range("B5").NumberFormat = "#,##0;[Red]-#,##0;0"
    ov.Range("B5").Calculate
    ov.Range("E5").Calculate
    Application.EnableEvents = eventsBefore

    MsgBox "26F totals linked to Budget_Overview B5 and E5." & vbCrLf & _
           "Future edits to Budget_26F_Input B2:B3 update Overview automatically.", vbInformation
    Exit Sub
Failed:
    failureText = "26F total import failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = eventsBefore
    MsgBox failureText, vbExclamation
End Sub
