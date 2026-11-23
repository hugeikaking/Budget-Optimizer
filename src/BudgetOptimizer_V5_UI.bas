Attribute VB_Name = "BudgetOptimizer_V5_UI"
Option Explicit

' V5 phase 1: UI enhancement only. Keep BudgetOptimizer_V4 and ThisWorkbook events.
' Layout preserves V4 coordinates: B3, E3, B6:F9, E11, R3:R5,
' and row 23+ columns A:M. No source or trial values are edited.

Private Const DETAIL_TAB As String = "Budget_Detail"
Private Const OVERVIEW_TAB As String = "Budget_Overview"
Private Const ROW_FIRST As Long = 23
Private Const ROW_LAST As Long = 4000

Public Sub SetupBudgetOptimizerV5UI()
    Dim detail As Worksheet, ov As Worksheet
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim failureText As String
    On Error GoTo Failed
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Set detail = ThisWorkbook.Worksheets(DETAIL_TAB)
    Set ov = ThisWorkbook.Worksheets(OVERVIEW_TAB)
    If Len(Trim$(CStr(detail.Range("A22").Value2))) = 0 Then _
        Err.Raise vbObjectError + 510, "SetupBudgetOptimizerV5UI", "V4 detail editor is missing. Run SetupBudgetOptimizerV4 first."

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Presentation only. Never clear the rows holding existing edit values.
    With detail
        .Range("A1:R1").Interior.Color = RGB(27, 49, 77)
        .Range("A1").Value2 = "BUDGET OPTIMIZER  |  H40 ADJUSTMENT WORKSPACE"
        .Range("A1:R1").Font.Name = "Aptos"
        .Range("A1:R1").Font.Bold = True
        .Range("A1:R1").Font.Color = RGB(255, 255, 255)
        .Range("A1:R1").Font.Size = 15
        .Rows(1).RowHeight = 33
        .Rows(2).RowHeight = 12

        .Range("A3:F3").Interior.Color = RGB(231, 239, 249)
        .Range("A3").Value2 = "H40 / Selected Group"
        .Range("D3").Value2 = "Fiscal Year (Apr-Mar)"
        .Range("A3:F3").Font.Bold = True
        .Range("B3").Interior.Color = RGB(255, 244, 193)
        .Range("B3").Font.Color = RGB(25, 64, 111)
        .Range("B3").Font.Bold = True
        .Range("B3").Borders.LineStyle = xlContinuous
        .Range("B3").Borders.Color = RGB(189, 208, 230)
        .Range("E3").Interior.Color = RGB(255, 244, 193)

        .Range("A5:F5").Interior.Color = RGB(47, 73, 105)
        .Range("A5:F5").Font.Color = RGB(255, 255, 255)
        .Range("A5:F5").Font.Bold = True
        .Rows(5).RowHeight = 28
        .Range("A6:F9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A6:F9").Borders(xlInsideHorizontal).Color = RGB(229, 235, 242)
        .Range("A6:A9").Font.Bold = True
        .Range("D6:D9").Interior.Color = RGB(228, 239, 250)
        .Range("D6:D9").Font.Bold = True
        .Range("E6:E9").Interior.Color = RGB(255, 244, 193)
        .Range("F6:F9").Font.Color = RGB(27, 80, 135)
        .Range("B6:F9").NumberFormat = "#,##0;[Red]-#,##0;0"

        .Range("H3:O3").Interior.Color = RGB(47, 73, 105)
        .Range("H3:O3").Font.Bold = True
        .Range("H3:O3").Font.Color = RGB(255, 255, 255)
        .Range("H3").Value2 = "LIVE FINANCIAL SNAPSHOT"
        .Range("H4").Value2 = "27F TTL Sales Amt"
        .Range("L4").Value2 = "Sales Amt YoY"
        .Range("H8").Value2 = "H40 Sales Amt (27F)"
        .Range("L8").Value2 = "H40 Sales Amt gap to target"
        .Range("H4:O4").Font.Color = RGB(91, 107, 124)
        .Range("H8:O8").Font.Color = RGB(91, 107, 124)
        .Range("H5:K6").Merge
        .Range("L5:O6").Merge
        .Range("H9:K10").Merge
        .Range("L9:O10").Merge
        .Range("H5").Formula = "='" & OVERVIEW_TAB & "'!E6"
        .Range("L5").Formula = "='" & OVERVIEW_TAB & "'!E7"
        .Range("H9").Formula = "=$D$8"
        .Range("L9").Formula = "=IF($E$8="""","""",$D$8-$E$8)"
        FormatMetricCell .Range("H5:K6"), False
        FormatMetricCell .Range("L5:O6"), True
        FormatMetricCell .Range("H9:K10"), False
        FormatMetricCell .Range("L9:O10"), False
        ' Current calculated sales amount is green, distinct from yellow manual targets.
        HighlightCurrentAmt .Range("H5:K6")
        HighlightCurrentAmt .Range("H9:K10")
        ' Global 27F sales amount target belongs to Overview, not individual H40.
        ' Keep this manually editable value in Overview!E9; never overwrite it on rerun.
        .Range("H12:O12").Merge
        .Range("H12").Value2 = "Green = current Sales Amt   |   Yellow = manual Target   |   Gray = source" 
        .Range("H12").Font.Color = RGB(95, 111, 131)
        .Range("H12").Font.Italic = True
        .Range("H13").Value2 = "27F TTL Sales Amt TARGET"
        .Range("L13").Value2 = "27F TTL Sales Amt GAP (actual - target)"
        .Range("H13:O13").Font.Color = RGB(91, 107, 124)
        .Range("H14:K15").Merge
        .Range("L14:O15").Merge
        .Range("H14").Formula = "=IF('" & OVERVIEW_TAB & "'!E9="""","""",'" & OVERVIEW_TAB & "'!E9)"
        .Range("L14").Formula = "=IF('" & OVERVIEW_TAB & "'!E9="""","""",'" & OVERVIEW_TAB & "'!E10)"
        FormatMetricCell .Range("H14:K15"), False
        FormatMetricCell .Range("L14:O15"), False
        HighlightTargetAmt .Range("H14:K15")

        .Range("A11").Font.Bold = True
        .Range("D11").Font.Bold = True
        .Range("E11").Font.Bold = True
        HighlightCurrentAmt .Range("E11")
        HighlightCurrentAmt .Range("D8")
        .Range("E6:E9").Interior.Color = RGB(255, 244, 193)
        .Range("A13").Font.Bold = True
        .Range("B13").NumberFormat = "0.00%;[Red]-0.00%;0.00%"

        .Range("A19:O19").Interior.Color = RGB(27, 49, 77)
        .Range("A19").Value2 = "H40 ROW EDITOR  |  Source values vs proposed edits"
        .Range("A19:O19").Font.Color = RGB(255, 255, 255)
        .Range("A19:O19").Font.Bold = True
        .Rows(19).RowHeight = 27
        .Range("A20").Value2 = "Select B3 (H40), then edit D:E month range and F:G parameters. H:I require R3 = YES."
        .Range("A21").Value2 = "Blank proposed values keep the original. Yellow background here means an input field; real changes are assessed on export."
        .Range("A20:O21").Font.Color = RGB(96, 113, 132)
        .Rows(20).RowHeight = 19
        .Rows(21).RowHeight = 19

        .Range("A22:N22").Value = Array("Source row", "Model", "H40", "From YYYY_MM", "To YYYY_MM", "New Pcs/Set", "New S/P", "New Share", "New Comp Ratio", "Original Pcs/Set*", "Original S/P*", "Original Share*", "Original Comp Ratio*", "Edit status")
        .Range("A22:N22").Interior.Color = RGB(47, 73, 105)
        .Range("A22:N22").Font.Bold = True
        .Range("A22:N22").Font.Color = RGB(255, 255, 255)
        .Range("A22:N22").WrapText = True
        .Rows(22).RowHeight = 35
        .Range("A" & ROW_FIRST & ":C" & ROW_LAST).Interior.Color = RGB(248, 250, 252)
        .Range("D" & ROW_FIRST & ":G" & ROW_LAST).Interior.Color = RGB(255, 248, 218)
        .Range("H" & ROW_FIRST & ":I" & ROW_LAST).Interior.Color = RGB(240, 242, 245)
        .Range("J" & ROW_FIRST & ":M" & ROW_LAST).Interior.Color = RGB(245, 247, 250)
        .Range("F" & ROW_FIRST & ":F" & ROW_LAST).NumberFormat = "#,##0"
        .Range("G" & ROW_FIRST & ":G" & ROW_LAST).NumberFormat = "#,##0.#####"
        .Range("H" & ROW_FIRST & ":I" & ROW_LAST).NumberFormat = "0"
        .Range("J" & ROW_FIRST & ":J" & ROW_LAST).NumberFormat = "#,##0"
        .Range("K" & ROW_FIRST & ":K" & ROW_LAST).NumberFormat = "#,##0.#####"
        .Range("L" & ROW_FIRST & ":M" & ROW_LAST).NumberFormat = "0"

        ' Status is a live formula and survives V4 clearing A:M when H40 changes.
        .Range("N" & ROW_FIRST & ":N" & ROW_LAST).FormulaR1C1 = _
            "=IF(RC1="""","""",IF(COUNTA(RC6:RC9)>0,""EDITED"",""ORIGINAL""))"
        .Range("N" & ROW_FIRST & ":N" & ROW_LAST).HorizontalAlignment = xlCenter
        .Range("N" & ROW_FIRST & ":N" & ROW_LAST).Interior.Color = RGB(244, 247, 250)
        .Range("F" & ROW_FIRST & ":I" & ROW_LAST).FormatConditions.Delete
        With .Range("F" & ROW_FIRST & ":I" & ROW_LAST).FormatConditions.Add( _
            Type:=xlExpression, Formula1:="=AND($A23<>"""",F23<>"""")")
            .Interior.Color = RGB(255, 218, 116)
            .Font.Color = RGB(88, 57, 8)
        End With
        .Range("H" & ROW_FIRST & ":I" & ROW_LAST).FormatConditions.Add _
            Type:=xlExpression, Formula1:="=AND($A23<>"""",$R$3=""YES"",H23="""")"
        .Range("H" & ROW_FIRST & ":I" & ROW_LAST).FormatConditions(1).Interior.Color = RGB(255, 248, 218)

        .Range("N" & ROW_FIRST & ":N" & ROW_LAST).FormatConditions.Delete
        With .Range("N" & ROW_FIRST & ":N" & ROW_LAST).FormatConditions.Add( _
            Type:=xlTextString, String:="EDITED", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 235, 168)
            .Font.Color = RGB(128, 84, 0)
        End With

        .Range("Q3:R5").Borders.LineStyle = xlContinuous
        .Range("Q3:R5").Borders.Color = RGB(220, 229, 239)
        .Range("Q3:Q5").Interior.Color = RGB(231, 239, 249)
        .Range("Q3:Q5").Font.Bold = True
        .Range("R3").Interior.Color = RGB(255, 244, 193)
        .Range("R4").Interior.Color = RGB(237, 247, 237)
        .Range("Q3").Value2 = "Unlock Share / Comp Ratio"
        .Range("Q4").Value2 = "Calculation status"
        .Range("Q5").Value2 = "Loaded H40 rows"

        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 24
        .Columns("C").ColumnWidth = 20
        .Columns("D:E").ColumnWidth = 18
        .Columns("F:I").ColumnWidth = 19
        .Columns("J:M").ColumnWidth = 21
        .Columns("N").ColumnWidth = 15
        .Columns("O").ColumnWidth = 14
        .Columns("P").ColumnWidth = 4
        .Columns("Q").ColumnWidth = 30
        .Columns("R").ColumnWidth = 23
        .Activate
        .Range("A23").Select
    End With

    ' Global sales amount target, separate from Overview B6 (Set QTY target).
    ' Do not clear or overwrite any existing E9 manual input.
    ov.Range("D9").Value2 = "27F TTL Sales Amt Target (manual)"
    ov.Range("D10").Value2 = "27F TTL Sales Amt Gap (actual - target)"
    ov.Range("E9").NumberFormat = "#,##0;[Red]-#,##0;0"
    HighlightTargetAmt ov.Range("E9")
    HighlightTargetAmt ov.Range("B6")
    On Error Resume Next
    ov.Range("E9").Validation.Delete
    On Error GoTo Failed
    With ov.Range("E9").Validation
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
            Operator:=xlGreaterEqual, Formula1:="0"
        .ErrorTitle = "Invalid sales amount target"
        .ErrorMessage = "Enter a nonnegative 27F TTL Sales Amt target."
        .ShowError = True
    End With
    ov.Range("E10").Formula = "=IF(E9="""","""",E6-E9)"
    ov.Range("E10").NumberFormat = "#,##0;[Red]-#,##0;0"
    ov.Range("E10").Interior.Color = RGB(228, 239, 250)
    HighlightCurrentAmt ov.Range("E6")
    HighlightTargetAmt ov.Range("B" & 40 & ":B" & ov.Cells(ov.Rows.Count, "S").End(xlUp).Row)
    ov.Range("D9:D10").Font.Bold = True

    ' The Overview wording must match the now automatic V4 editor behavior.
    ov.Range("A16").Value2 = "4. Choose H40 in Budget_Detail; changing Pcs/Set or S/P recalculates trial automatically."
    ov.Range("A17").Value2 = "Total Set QTY is counted once per model, not once per H40."

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "V5 UI ready. V4 calculation logic and all existing edits are retained.", vbInformation
    Exit Sub
Failed:
    failureText = "V5 UI failed (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox failureText, vbExclamation
End Sub

Private Sub FormatMetricCell(ByVal area As Range, ByVal isPercent As Boolean)
    With area
        .Interior.Color = RGB(228, 239, 250)
        .Font.Name = "Aptos Display"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(27, 72, 119)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        If isPercent Then
            .NumberFormat = "0.00%;[Red]-0.00%;0.00%"
        Else
            .NumberFormat = "#,##0;[Red]-#,##0;0"
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(202, 218, 236)
    End With
End Sub

' Color convention: actual/live Sales Amt is green; manual target is yellow.
' This affects presentation only and never changes values, formulas or Excel events.
Private Sub HighlightCurrentAmt(ByVal area As Range)
    With area
        .Interior.Color = RGB(222, 241, 225)
        .Font.Color = RGB(24, 96, 54)
        .Font.Bold = True
    End With
End Sub

Private Sub HighlightTargetAmt(ByVal area As Range)
    With area
        .Interior.Color = RGB(255, 239, 166)
        .Font.Color = RGB(102, 74, 12)
        .Font.Bold = True
    End With
End Sub
