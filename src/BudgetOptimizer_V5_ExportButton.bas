Attribute VB_Name = "BudgetOptimizer_V5_ExportButton"
Option Explicit

' V5 optional UX add-on: an export button on Budget_Overview.
' The verified ExportBudgetV5SimpleCopy remains unchanged.
' Run SetupBudgetV5ExportButton once in the development workbook.
' No change to ThisWorkbook, the V4 trial engine, or source data.

Private Const BTN_NAME As String = "btnBudgetV5SimpleExport"
Private Const BTN_OVERVIEW As String = "Budget_Overview"

Public Sub SetupBudgetV5ExportButton()
    Dim ws As Worksheet, sh As Shape
    Dim prevEvents As Boolean, errText As String

    prevEvents = Application.EnableEvents
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(BTN_OVERVIEW)
    Application.EnableEvents = False

    On Error Resume Next
    ws.Shapes(BTN_NAME).Delete
    On Error GoTo Failed

    ' Use H5:K7, above the existing seasonality chart (row 20+).
    ' Avoid moving/modifying any current budget input cells.
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("H5").Left, ws.Range("H5").Top, _
        ws.Range("H5:K7").Width, ws.Range("H5:K7").Height)
    With sh
        .Name = BTN_NAME
        .TextFrame.Characters.Text = "EXPORT BUDGET" & vbLf & _
            "Value-only / C flags / yellow edits"
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(30, 98, 74)
        .Line.ForeColor.RGB = RGB(22, 73, 55)
        .Placement = xlMoveAndSize
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
            "'!RunBudgetV5ExportButton"
    End With

    Application.EnableEvents = prevEvents
    MsgBox "Export Budget button added to Budget_Overview near H5.", vbInformation
    Exit Sub

Failed:
    errText = "Export button setup failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = prevEvents
    MsgBox errText, vbExclamation
End Sub

' Count effective differences using the same five input blocks and tolerance
' as the verified value-only exporter. Never rely only on the formula C marker:
' it may be stale if a manual calculation mode is selected.
Public Sub RunBudgetV5ExportButton()
    Dim src As Worksheet, trial As Worksheet
    Dim lastSrc As Long, lastTrial As Long, rr As Long, mm As Long, k As Long
    Dim bases As Variant, a As Variant, b As Variant
    Dim nCells As Long, nRows As Long, changed As Boolean, colNum As Long
    Dim h As String, errText As String

    On Error GoTo Failed
    On Error Resume Next
    Set src = ThisWorkbook.Worksheets("Budget_Data")
    If src Is Nothing Then Set src = ThisWorkbook.Worksheets("Sheet1")
    Set trial = ThisWorkbook.Worksheets("Budget_Trial")
    On Error GoTo Failed
    If src Is Nothing Or trial Is Nothing Then _
        Err.Raise vbObjectError + 721, , "Budget_Data/Sheet1 or Budget_Trial missing."

    lastSrc = src.Cells(src.Rows.Count, 19).End(xlUp).Row
    lastTrial = trial.Cells(trial.Rows.Count, 19).End(xlUp).Row
    If lastSrc <> lastTrial Then _
        Err.Raise vbObjectError + 722, , "Source and trial row counts differ. Refresh trial."

    bases = Array(24, 44, 64, 84, 104)
    For rr = 8 To lastSrc
        h = Trim$(CStr(src.Cells(rr, 19).Value2))
        If Len(h) > 0 And LCase$(h) <> "h40" And LCase$(h) <> "eop" Then
            If CStr(src.Cells(rr, 19).Value2) <> CStr(trial.Cells(rr, 19).Value2) Or _
               CStr(src.Cells(rr, 2).Value2) <> CStr(trial.Cells(rr, 2).Value2) Then _
                Err.Raise vbObjectError + 723, , "Row mismatch at " & rr
            changed = False
            For k = 0 To 4
                For mm = 0 To 19
                    colNum = CLng(bases(k)) + mm
                    a = src.Cells(rr, colNum).Value2
                    b = trial.Cells(rr, colNum).Value2
                    If BudgetV5ButtonDifferent(a, b) Then
                        changed = True
                        nCells = nCells + 1
                    End If
                Next mm
            Next k
            If changed Then nRows = nRows + 1
        End If
    Next rr

    If MsgBox("Current differences versus Budget_Data:" & vbCrLf & _
        "  Changed rows: " & Format$(nRows, "#,##0") & vbCrLf & _
        "  Changed input cells: " & Format$(nCells, "#,##0") & vbCrLf & vbCrLf & _
        "Continue to choose a new export file?" & vbCrLf & _
        "(The export macro will ask for final confirmation again.)", _
        vbQuestion + vbYesNo, "Export Budget Review") <> vbYes Then Exit Sub

    ExportBudgetV5SimpleCopy
    Exit Sub

Failed:
    errText = "Export preview failed (" & Err.Number & "): " & Err.Description
    MsgBox errText, vbExclamation
End Sub

Private Function BudgetV5ButtonDifferent(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then
        BudgetV5ButtonDifferent = True
        Exit Function
    End If
    If Len(CStr(a)) = 0 And Len(CStr(b)) = 0 Then Exit Function
    If IsNumeric(a) And IsNumeric(b) Then
        BudgetV5ButtonDifferent = Abs(CDbl(a) - CDbl(b)) > _
            0.00000001 * Application.Max(1#, Abs(CDbl(a)), Abs(CDbl(b)))
    Else
        BudgetV5ButtonDifferent = CStr(a) <> CStr(b)
    End If
End Function
