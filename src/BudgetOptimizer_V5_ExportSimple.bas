Attribute VB_Name = "BudgetOptimizer_V5_ExportSimple"
Option Explicit

' Simple value-only budget review export.
' Source formatting and row/column positions remain. Output contains no
' copied budget formulas or helper flags; changed input values are highlighted.
' This module does not edit source/trial/overview in the tool workbook.

Private Const S_MONTHS As Long = 20
Private Const S_LAST_INPUT As Long = 123 ' DS
Private Const S_H40_COL As Long = 19
Private Const S_DIFF_TOL As Double = 0.00000001

Public Sub ExportBudgetV5SimpleCopy()
    Dim src As Worksheet, trial As Worksheet, outBook As Workbook, dst As Worksheet
    Dim lastRow As Long, firstRow As Long, rr As Long, cc As Long
    Dim bases As Variant, k As Long, mm As Long, n As Long, marked As Long
    Dim originalValues As Variant, trialValues As Variant
    Dim firstData As Long, colInput As Long, rowRelative As Long
    Dim outName As Variant, suggested As String, stage As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean
    Dim outputCell As Range, fCells As Range, one As Range, area As Range
    Dim links As Variant, i As Long, nm As Name
    Dim rawH40 As Variant, originalValue As Variant, newValue As Variant
    Dim errorText As String

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed

    Set src = SimpleFindSheet("Budget_Data")
    If src Is Nothing Then Set src = SimpleFindSheet("Sheet1")
    Set trial = SimpleFindSheet("Budget_Trial")
    If src Is Nothing Or trial Is Nothing Then _
        Err.Raise vbObjectError + 701, , "Budget_Data and Budget_Trial required."
    lastRow = src.Cells(src.Rows.Count, S_H40_COL).End(xlUp).Row
    If lastRow <> trial.Cells(trial.Rows.Count, S_H40_COL).End(xlUp).Row Then _
        Err.Raise vbObjectError + 702, , "Source and trial rows differ. Refresh trial."
    firstRow = 8
    For rr = firstRow To lastRow
        If SimpleValidRow(src, rr) Then
            If CStr(src.Cells(rr, 19).Value2) <> CStr(trial.Cells(rr, 19).Value2) Or _
               CStr(src.Cells(rr, 2).Value2) <> CStr(trial.Cells(rr, 2).Value2) Then _
                Err.Raise vbObjectError + 703, , "Row identifier mismatch at row " & rr
        End If
    Next rr
    If MsgBox("Create a value-only budget copy (with C row flags and yellow changed inputs)?" & _
        vbCrLf & "This will NOT retain formulas or overwrite the original.", _
        vbQuestion + vbYesNo, "Simple Budget Export") <> vbYes Then Exit Sub

    suggested = "Budget_Review_Values_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    outName = Application.GetSaveAsFilename(InitialFileName:=suggested, _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx", Title:="Save value-only budget review")
    If VarType(outName) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(outName), 5)) <> ".xlsx" Then _
        Err.Raise vbObjectError + 704, , "Choose an .xlsx file."
    If Len(Dir$(CStr(outName))) <> 0 Then _
        Err.Raise vbObjectError + 705, , "Choose a new filename; overwrite not allowed."

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    stage = "Copying original worksheet formatting"
    src.Copy
    Set outBook = ActiveWorkbook
    Set dst = outBook.Worksheets(1)
    dst.Name = "Sheet1"

    ' IMPORTANT: Use already calculated SOURCE values, never recalculate copied
    ' workbook formulas that might have become external after worksheet copy.
    stage = "Freezing original formula cells into values"
    On Error Resume Next
    Set fCells = dst.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo Failed
    If Not fCells Is Nothing Then
        For Each area In fCells.Areas
            On Error Resume Next
            area.Value2 = src.Range(area.Address).Value2
            If Err.Number <> 0 Then
                Err.Clear
                For Each one In area.Cells
                    If one.HasFormula Then
                        one.Value2 = src.Range(one.Address).Value2
                    End If
                Next one
            End If
            On Error GoTo Failed
        Next area
    End If

    ' Do not delete every Excel-defined name. Built-in/hidden names may
    ' reject deletion with error 1004. Only the known V4 H40 dropdown helper
    ' can create a cross-workbook link when Sheet1 is copied alone.
    stage = "Removing copied BudgetH40Options helper from output only"
    SimpleRemoveCopiedH40Name outBook

    ' Compare the same five input blocks used by Budget_Trial row C flag.
    ' X:AQ, AR:BK, BL:CE, CF:CY, CZ:DS.
    stage = "Applying actual differences from Budget_Trial"
    bases = Array(24, 44, 64, 84, 104)
    For k = 0 To 4
        For mm = 0 To S_MONTHS - 1
            colInput = CLng(bases(k)) + mm
            originalValues = src.Range(src.Cells(firstRow, colInput), src.Cells(lastRow, colInput)).Value2
            trialValues = trial.Range(trial.Cells(firstRow, colInput), trial.Cells(lastRow, colInput)).Value2
            For rowRelative = 1 To lastRow - firstRow + 1
                rr = firstRow + rowRelative - 1
                If SimpleValidRow(src, rr) Then
                    originalValue = originalValues(rowRelative, 1)
                    newValue = trialValues(rowRelative, 1)
                    If SimpleDifferent(originalValue, newValue) Then
                        If IsError(newValue) Or Not IsNumeric(newValue) Then _
                            Err.Raise vbObjectError + 706, , _
                                "Invalid trial input at " & trial.Cells(rr, colInput).Address(False, False)
                        Set outputCell = dst.Cells(rr, colInput)
                        outputCell.Value2 = newValue
                        outputCell.Interior.Color = RGB(255, 235, 115)
                        If CStr(dst.Cells(rr, 1).Value2) <> "C" Then marked = marked + 1
                        dst.Cells(rr, 1).Value2 = "C"
                        n = n + 1
                    End If
                End If
            Next rowRelative
        Next mm
    Next k

    ' Sales Amt DW:EP is a calculated output, not a manually adjusted input.
    ' Freeze the current Trial amounts as static values without yellow marks.
    stage = "Copying calculated Sales Amt values from trial"
    For mm = 0 To S_MONTHS - 1
        colInput = 127 + mm
        trialValues = trial.Range(trial.Cells(firstRow, colInput), _
            trial.Cells(lastRow, colInput)).Value2
        For rowRelative = 1 To lastRow - firstRow + 1
            rr = firstRow + rowRelative - 1
            If SimpleValidRow(src, rr) Then
                newValue = trialValues(rowRelative, 1)
                If IsError(newValue) Then Err.Raise vbObjectError + 709, , _
                    "Trial Sales Amt error at " & trial.Cells(rr, colInput).Address(False, False)
                dst.Cells(rr, colInput).Value2 = newValue
            End If
        Next rowRelative
    Next mm

    ' Never carry a dynamic C formula into the exported worksheet. Rows that
    ' have no effective input differences retain the original source A value.
    stage = "Validating review copy"
    On Error Resume Next
    Set fCells = Nothing
    Set fCells = dst.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo Failed
    If Not fCells Is Nothing Then _
        Err.Raise vbObjectError + 707, , "Output still contains formula cells: export stopped."
    links = outBook.LinkSources(Type:=xlExcelLinks)
    If IsArray(links) Then _
        Err.Raise vbObjectError + 708, , _
            "Output still contains an external workbook link (chart, object or connection)."

    stage = "Saving output copy"
    outBook.SaveAs Filename:=CStr(outName), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    outBook.Close SaveChanges:=False
    Set outBook = Nothing

    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Value-only export completed." & vbCrLf & _
        "Changed cells highlighted: " & Format$(n, "#,##0") & vbCrLf & _
        "Rows marked C: " & Format$(marked, "#,##0") & vbCrLf & _
        "File: " & CStr(outName), vbInformation
    Exit Sub
Failed:
    errorText = "Value-only export failed at " & stage & " (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If Not outBook Is Nothing Then outBook.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox errorText & vbCrLf & "Source and trial were not changed.", vbExclamation
End Sub

Private Function SimpleFindSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SimpleFindSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function SimpleValidRow(ByVal ws As Worksheet, ByVal rr As Long) As Boolean
    Dim h As String
    h = Trim$(CStr(ws.Cells(rr, S_H40_COL).Value2))
    SimpleValidRow = Len(h) > 0 And LCase$(h) <> "h40" And LCase$(h) <> "eop"
End Function

Private Function SimpleDifferent(ByVal oldValue As Variant, ByVal newValue As Variant) As Boolean
    If IsError(oldValue) Or IsError(newValue) Then
        SimpleDifferent = True
        Exit Function
    End If
    If Len(CStr(oldValue)) = 0 And Len(CStr(newValue)) = 0 Then Exit Function
    If IsNumeric(oldValue) And IsNumeric(newValue) Then
        SimpleDifferent = Abs(CDbl(oldValue) - CDbl(newValue)) > _
            S_DIFF_TOL * Application.Max(1#, Abs(CDbl(oldValue)), Abs(CDbl(newValue)))
    Else
        SimpleDifferent = CStr(oldValue) <> CStr(newValue)
    End If
End Function


' Only remove V4's known internal dropdown name from the NEW output workbook.
' Other defined names are retained; later LinkSources validation is still strict.
Private Sub SimpleRemoveCopiedH40Name(ByVal copiedBook As Workbook)
    Dim i As Long, oneName As Name, localName As String
    Dim bangAt As Long, referText As String, deleteFailure As String

    For i = copiedBook.Names.Count To 1 Step -1
        Set oneName = copiedBook.Names(i)
        localName = oneName.Name
        bangAt = InStrRev(localName, "!")
        If bangAt > 0 Then localName = Mid$(localName, bangAt + 1)
        If StrComp(localName, "BudgetH40Options", vbTextCompare) = 0 Then
            referText = vbNullString
            On Error Resume Next
            referText = oneName.RefersTo
            On Error GoTo 0
            If InStr(1, referText, "[", vbBinaryCompare) > 0 And _
               InStr(1, referText, "]", vbBinaryCompare) > 0 Then
                On Error Resume Next
                oneName.Delete
                If Err.Number <> 0 Then
                    deleteFailure = Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
                If Len(deleteFailure) > 0 Then Err.Raise vbObjectError + 710, , _
                    "Could not delete copied BudgetH40Options helper: " & deleteFailure
            End If
        End If
    Next i
End Sub
