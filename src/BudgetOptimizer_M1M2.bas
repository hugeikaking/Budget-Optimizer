Attribute VB_Name = "BudgetOptimizer_M1M2"
Option Explicit

' Budget Optimizer M1/M2 - DEMO ONLY
' Import this module into a COPY of Budget_Optimizer_M1M2_Demo.xlsx saved as .xlsm.
' Never use the test macro directly against an official budget workbook.
' The sample's S column contains fictitious H40 groups.
'
Private Const DATA_SHEET As String = "Budget_Data"
Private Const TRIAL_SHEET As String = "Budget_Trial"
Private Const PANEL_SHEET As String = "Budget_Panel"
Private Const PLAN_SHEET As String = "Edit_Plan"
Private Const BASE_SHEET As String = "Base26"
Private Const HEADER_ROW As Long = 7
Private Const FIRST_ROW As Long = 8
Private Const FIRST_MONTH_COL As Long = 24 ' X
Private Const MONTHS As Long = 20

Public Sub RefreshBudgetPreview()
    Dim source As Worksheet, trial As Worksheet, panel As Worksheet
    Dim plan As Worksheet, baselines As Worksheet
    Dim h40 As String, focusYear As Long, r As Long, k As Long, yr As Long
    Dim m As Long, j As Long, baseRow As Long, maxRow As Long
    Dim metricsNow(1 To 4) As Double, metrics26(1 To 4) As Double
    Dim metrics26Original(1 To 4) As Double, metricsAll27(1 To 4) As Double
    Dim tmp(1 To 4) As Double, annual(1 To 4) As Double
    Dim v As Variant, nChanges As Long
    On Error GoTo Failed
    Set source = ThisWorkbook.Worksheets(DATA_SHEET)
    Set panel = ThisWorkbook.Worksheets(PANEL_SHEET)
    Set plan = ThisWorkbook.Worksheets(PLAN_SHEET)
    Set baselines = ThisWorkbook.Worksheets(BASE_SHEET)
    h40 = Trim$(CStr(panel.Range("B3").Value2))
    focusYear = CLng(panel.Range("E3").Value2)
    If h40 = "" Then Err.Raise vbObjectError + 10, , "Select an H40 in Budget_Panel!B3"
    If focusYear < 2026 Or focusYear > 2100 Then Err.Raise vbObjectError + 11, , "Focus year is invalid"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(TRIAL_SHEET).Delete
    On Error GoTo Failed
    Application.DisplayAlerts = True
    source.Copy After:=source
    Set trial = ActiveSheet
    trial.Name = TRIAL_SHEET
    maxRow = trial.Cells(trial.Rows.Count, "S").End(xlUp).Row
    ' Trial modifications are applied only to a copied worksheet.
    For r = 2 To plan.Cells(plan.Rows.Count, "B").End(xlUp).Row
        If UCase$(Trim$(CStr(plan.Cells(r, "A").Value2))) = "Y" Then
            ApplyPlanRow trial, plan, r, maxRow
            nChanges = nChanges + 1
        End If
    Next r
    trial.Calculate
    For r = FIRST_ROW To maxRow
        If CStr(trial.Cells(r, "S").Value2) = h40 Then
            For m = 0 To MONTHS - 1
                yr = MonthYear(trial.Cells(HEADER_ROW, FIRST_MONTH_COL + m).Value2) 
                If yr > 0 Then
                    RowMonthMetrics trial, r, m, tmp
                    If yr = 2026 Then AddFour metrics26, tmp
                    If yr = focusYear Then AddFour metricsNow, tmp
                End If
            Next m
        End If
        For m = 0 To MONTHS - 1
            yr = MonthYear(trial.Cells(HEADER_ROW, FIRST_MONTH_COL + m).Value2)
            If yr = focusYear Then
                RowMonthMetrics trial, r, m, tmp
                AddFour metricsAll27, tmp
            End If
        Next m
    Next r
    ' Comparison: original in-window 26F from Budget_Data.
    For r = FIRST_ROW To source.Cells(source.Rows.Count, "S").End(xlUp).Row
        If CStr(source.Cells(r, "S").Value2) = h40 Then
            For m = 0 To MONTHS - 1
                If MonthYear(source.Cells(HEADER_ROW, FIRST_MONTH_COL + m).Value2) = 2026 Then
                    RowMonthMetrics source, r, m, tmp
                    AddFour metrics26Original, tmp
                End If
            Next m
        End If
    Next r
    ' Written base 26F is a manually persisted full-year baseline.
    baseRow = 0
    For r = 2 To baselines.Cells(baselines.Rows.Count, "A").End(xlUp).Row
        If CStr(baselines.Cells(r, "A").Value2) = h40 Then baseRow = r: Exit For
    Next r
    For k = 1 To 4
        panel.Cells(5 + k, "C").Value2 = metrics26(k)
        panel.Cells(5 + k, "D").Value2 = metricsNow(k)
        panel.Cells(5 + k, "B").ClearContents
        If baseRow > 0 Then
            v = baselines.Cells(baseRow, 1 + k).Value2
            If IsNumeric(v) And Len(CStr(v)) > 0 Then
                annual(k) = CDbl(v) + metrics26(k) - metrics26Original(k)
                panel.Cells(5 + k, "B").Value2 = annual(k)
            End If
        End If
    Next k
    panel.Range("E11").Value2 = metricsAll27(3)
    panel.Range("E12").Value2 = "Preview ready (" & nChanges & " plan rows)"
    panel.Range("B13").Formula = "=IFERROR(B12/B11-1,"""")"
    panel.Range("B11").Formula = "=B8"
    panel.Range("B12").Formula = "=D8"
    panel.Range("F6:F9").FormulaR1C1 = "=IF(RC[-1]="""", """", RC[-2]-RC[-1])"
    Application.Calculate
    MsgBox "Preview complete. Changes exist only in Budget_Trial, not in Budget_Data.", vbInformation
Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    MsgBox "Preview could not be completed: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub ApplyPlanRow(ByVal trial As Worksheet, ByVal plan As Worksheet, ByVal line As Long, ByVal maxRow As Long)
    Dim dataRow As Long, fromKey As Long, toKey As Long
    Dim m As Long, key As Long, x As Variant
    dataRow = CLng(Val(plan.Cells(line, "B").Value2))
    If dataRow < FIRST_ROW Or dataRow > maxRow Then Err.Raise vbObjectError + 22, , "Invalid data row in Edit_Plan row " & line
    fromKey = MonthKey(plan.Cells(line, "C").Value2)
    toKey = MonthKey(plan.Cells(line, "D").Value2)
    If fromKey = 0 Or toKey = 0 Or fromKey > toKey Then Err.Raise vbObjectError + 23, , "Invalid month interval in Edit_Plan row " & line
    If Len(Trim$(CStr(plan.Cells(line, "I").Value2))) > 0 Then
        If CStr(trial.Cells(dataRow, "S").Value2) <> CStr(plan.Cells(line, "I").Value2) Then Err.Raise vbObjectError + 24, , "H40 mismatch in Edit_Plan row " & line
    End If
    For m = 0 To MONTHS - 1
        key = MonthKey(trial.Cells(HEADER_ROW, FIRST_MONTH_COL + m).Value2)
        If key >= fromKey And key <= toKey Then
            x = plan.Cells(line, "E").Value2 ' Pcs/Set
            If Len(CStr(x)) > 0 Then
                If Not IsNumeric(x) Then Err.Raise vbObjectError + 25, , "Invalid Pcs/Set"
                If CDbl(x) < 0 Or CDbl(x) <> Fix(CDbl(x)) Then Err.Raise vbObjectError + 26, , "Pcs/Set must be a non-negative integer"
                trial.Cells(dataRow, 44 + m).Value2 = CDbl(x)
            End If
            x = plan.Cells(line, "F").Value2 ' S/P
            If Len(CStr(x)) > 0 Then
                If Not IsNumeric(x) Then Err.Raise vbObjectError + 27, , "Invalid S/P"
                If CDbl(x) < 0 Or Abs(CDbl(x) * 100000 - Round(CDbl(x) * 100000, 0)) > 0.000001 Then Err.Raise vbObjectError + 28, , "S/P must be non-negative with at most five decimals"
                trial.Cells(dataRow, 84 + m).Value2 = CDbl(x)
            End If
            x = plan.Cells(line, "G").Value2 ' Share
            If Len(CStr(x)) > 0 Then
                If Not IsNumeric(x) Then Err.Raise vbObjectError + 29, , "Invalid Share"
                If CDbl(x) < 0 Or CDbl(x) > 100 Then Err.Raise vbObjectError + 30, , "Share must be between 0 and 100"
                trial.Cells(dataRow, 104 + m).Value2 = CDbl(x)
            End If
            x = plan.Cells(line, "H").Value2 ' Comp Ratio
            If Len(CStr(x)) > 0 Then
                If Not IsNumeric(x) Then Err.Raise vbObjectError + 31, , "Invalid Comp Ratio"
                If CDbl(x) < 0 Or CDbl(x) > 100 Then Err.Raise vbObjectError + 32, , "Comp Ratio must be between 0 and 100"
                trial.Cells(dataRow, 64 + m).Value2 = CDbl(x)
            End If
        End If
    Next m
End Sub

Private Sub RowMonthMetrics(ByVal ws As Worksheet, ByVal r As Long, ByVal m As Long, ByRef result() As Double)
    Dim sets As Double, pieces As Double, ratio As Double, sp As Double, share As Double
    Dim demandQty As Double, salesQty As Double
    sets = Val(ws.Cells(r, 24 + m).Value2)
    pieces = Val(ws.Cells(r, 44 + m).Value2)
    ratio = Val(ws.Cells(r, 64 + m).Value2)
    sp = Val(ws.Cells(r, 84 + m).Value2)
    share = Val(ws.Cells(r, 104 + m).Value2)
    demandQty = sets * pieces * ratio / 100#
    salesQty = demandQty * share / 100#
    result(1) = WorksheetFunction.Round(demandQty * sp, 2)  ' Demand Amt
    result(2) = demandQty                                 ' Demand Qty
    ' Sales Amt is read from the recalculated trial workbook formula when present.
    If IsNumeric(ws.Cells(r, 127 + m).Value2) Then
        result(3) = CDbl(ws.Cells(r, 127 + m).Value2)
    Else
        result(3) = WorksheetFunction.Round(salesQty * sp, 2)
    End If
    result(4) = salesQty                                  ' Sales Qty
End Sub

Private Sub AddFour(ByRef dst() As Double, ByRef src() As Double)
    Dim k As Long
    For k = 1 To 4
        dst(k) = dst(k) + src(k)
    Next k
End Sub

Private Function MonthKey(ByVal rawValue As Variant) As Long
    On Error GoTo Invalid
    Dim st As String, y As Long, m As Long
    If IsDate(rawValue) Then
        MonthKey = Year(CDate(rawValue)) * 100 + Month(CDate(rawValue))
        Exit Function
    End If
    st = Replace(Trim$(CStr(rawValue)), "_", "/")
    st = Replace(st, "-", "/")
    y = CLng(Split(st, "/")(0))
    m = CLng(Split(st, "/")(1))
    If m < 1 Or m > 12 Or y < 2000 Then GoTo Invalid
    MonthKey = y * 100 + m
    Exit Function
Invalid:
    MonthKey = 0
End Function

Private Function MonthYear(ByVal rawValue As Variant) As Long
    MonthYear = MonthKey(rawValue) \ 100
End Function