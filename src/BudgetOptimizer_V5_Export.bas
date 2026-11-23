Attribute VB_Name = "BudgetOptimizer_V5_Export"
Option Explicit

' V5 phase 2: standalone, reviewed output of the current V4 trial.
' This module only READS Budget_Data/Sheet1 and Budget_Trial.
' It never writes to either source or to the stable workbook.
' Export is limited to input blocks X:DS. DW:EP Sales Amt formulas remain intact.

Private Const EX_SOURCE As String = "Budget_Data"
Private Const EX_FALLBACK As String = "Sheet1"
Private Const EX_TRIAL As String = "Budget_Trial"
Private Const EX_OVERVIEW As String = "Budget_Overview"
Private Const EX_LOG As String = "Change_Log"
Private Const EX_MONTHS As Long = 20
Private Const EX_LAST_COL As Long = 146  ' EP, original source output width
Private Const EX_H40_COL As Long = 19
Private Const EX_SALES_COL As Long = 127 ' DW
Private Const EX_GREEN As Long = 0
Private Const EX_MARGIN As Double = 0.00000001

Public Sub ExportBudgetV5YellowCopy()
    Dim src As Worksheet, trial As Worksheet, ov As Worksheet, exported As Worksheet
    Dim outputBook As Workbook, logSheet As Worksheet
    Dim monthKeys(0 To EX_MONTHS - 1) As Long
    Dim header As Long, firstData As Long, lastData As Long
    Dim prevEvents As Boolean, prevScreen As Boolean, prevAlerts As Boolean
    Dim outputPath As Variant, suggested As String, failureMessage As String
    Dim n As Long, targetFiscalYear As Long, hasLink As Boolean
    Dim linkData As Variant
    Dim firstLogRow As Long, outRow As Long
    Dim blockBases As Variant, blockNames As Variant
    Dim blockIndex As Long, m As Long, rowNum As Long, colNum As Long
    Dim originalValue As Variant, adjustedValue As Variant
    Dim originalData As Variant, adjustedData As Variant
    Dim rowCount As Long, currentModel As String, currentKey As String
    Dim modelData As Variant, keysData As Variant, h40Data As Variant
    Dim destinationCell As Range
    Dim currentStage As String

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set src = ExportSourceSheet()
    Set trial = ExportFindSheet(EX_TRIAL)
    Set ov = ExportFindSheet(EX_OVERVIEW)
    If src Is Nothing Or trial Is Nothing Or ov Is Nothing Then _
        Err.Raise vbObjectError + 601, , _
        "Source, Budget_Trial or Budget_Overview missing. Complete a V4 trial first."

    currentStage = "Checking worksheet structure"
    lastData = src.Cells(src.Rows.Count, EX_H40_COL).End(xlUp).Row
    If lastData <> trial.Cells(trial.Rows.Count, EX_H40_COL).End(xlUp).Row Then _
        Err.Raise vbObjectError + 602, , "Trial/source last row mismatch. Refresh the trial first."
    If lastData < 8 Then Err.Raise vbObjectError + 603, , "Source has too few rows."
    ExportGetMonths src, monthKeys, header
    firstData = ExportFirstDataRow(src, header, lastData)
    ExportCheckAlignment src, trial, firstData, lastData
    targetFiscalYear = CLng(ov.Range("B3").Value2)
    If targetFiscalYear < 2026 Or targetFiscalYear > 2100 Then _
        Err.Raise vbObjectError + 604, , "Invalid fiscal year in Overview!B3."
    If MsgBox("Export a NEW standalone budget copy with yellow input changes and Change_Log?" & vbCrLf & _
             "The original Budget_Data and V4 trial are not overwritten.", _
             vbQuestion + vbYesNo, "Export V5 Budget") <> vbYes Then Exit Sub

    suggested = "Budget_27F_Reviewed_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    outputPath = Application.GetSaveAsFilename(InitialFileName:=suggested, _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx", _
        Title:="Save reviewed budget copy (new file only)")
    If VarType(outputPath) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(outputPath), 5)) <> ".xlsx" Then _
        Err.Raise vbObjectError + 605, , "Output must be an .xlsx file."
    If Len(Dir$(CStr(outputPath))) > 0 Then _
        Err.Raise vbObjectError + 606, , "The selected file exists. Choose a NEW file name."

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    currentStage = "Copying untouched original budget sheet"
    src.Copy
    Set outputBook = ActiveWorkbook
    Set exported = outputBook.Worksheets(1)
    exported.Name = "Sheet1"

    ' Excel may create workbook links when a single worksheet is copied.
    ' First inspect the actual formula references, not only LinkSources.
    ' If copied names alone caused phantom links, remove only demonstrably
    ' unused external names in the NEW workbook and validate again.
    currentStage = "Auditing copied sheet formulas and names"
    ExportAuditCopiedLinks src, outputBook, exported

    currentStage = "Creating change log"
    Set logSheet = outputBook.Worksheets.Add(After:=exported)
    logSheet.Name = EX_LOG
    logSheet.Range("A1:J1").Value = Array("Row", "Excel cell", "Model (M)", _
        "Model ID (B)", "H40 (S)", "Month", "Parameter", _
        "Original", "Exported", "Delta")
    logSheet.Range("A1:J1").Font.Bold = True
    logSheet.Range("A1:J1").Interior.Color = RGB(35, 57, 84)
    logSheet.Range("A1:J1").Font.Color = RGB(255, 255, 255)
    logSheet.Range("A2:J2").Merge
    logSheet.Range("A2").Value2 = "Only effective differences in the input blocks X:DS are highlighted yellow. Sales Amt DW:EP formulas are preserved."
    logSheet.Range("A2").Font.Italic = True
    logSheet.Range("A3:J3").Value = Array("Row", "Excel cell", "Model (M)", _
        "Model ID (B)", "H40 (S)", "Month", "Parameter", _
        "Original", "Exported", "Delta")
    logSheet.Range("A3:J3").Font.Bold = True
    firstLogRow = 4
    outRow = firstLogRow

    ' Exact-column mapping: 24 X, 44 AR, 64 BL, 84 CF, 104 CZ.
    ' No edits to Actual S/P DT:DV or Sales Amt DW:EP.
    blockBases = Array(24, 44, 64, 84, 104)
    blockNames = Array("Set QTY", "Pcs/Set", "Comp Ratio", "S/P", "Share")
    rowCount = lastData - firstData + 1

    modelData = src.Range(src.Cells(firstData, 13), src.Cells(lastData, 13)).Value2
    keysData = src.Range(src.Cells(firstData, 2), src.Cells(lastData, 2)).Value2
    h40Data = src.Range(src.Cells(firstData, EX_H40_COL), src.Cells(lastData, EX_H40_COL)).Value2

    currentStage = "Comparing 100 input-month columns"
    For blockIndex = 0 To 4
        For m = 0 To EX_MONTHS - 1
            colNum = CLng(blockBases(blockIndex)) + m
            originalData = src.Range(src.Cells(firstData, colNum), src.Cells(lastData, colNum)).Value2
            adjustedData = trial.Range(trial.Cells(firstData, colNum), trial.Cells(lastData, colNum)).Value2
            currentModel = vbNullString
            currentKey = vbNullString
            For rowNum = 1 To rowCount
                If Not IsError(modelData(rowNum, 1)) Then
                    If Len(Trim$(CStr(modelData(rowNum, 1)))) > 0 Then _
                        currentModel = Trim$(CStr(modelData(rowNum, 1)))
                End If
                If Not IsError(keysData(rowNum, 1)) Then
                    If Len(Trim$(CStr(keysData(rowNum, 1)))) > 0 Then _
                        currentKey = Trim$(CStr(keysData(rowNum, 1)))
                End If
                If Not IsError(h40Data(rowNum, 1)) Then
                    If Len(Trim$(CStr(h40Data(rowNum, 1)))) > 0 And _
                       StrComp(Trim$(CStr(h40Data(rowNum, 1))), "H40", vbTextCompare) <> 0 And _
                       StrComp(Trim$(CStr(h40Data(rowNum, 1))), "EOP", vbTextCompare) <> 0 Then
                        originalValue = originalData(rowNum, 1)
                        adjustedValue = adjustedData(rowNum, 1)
                        If ExportIsDifferent(originalValue, adjustedValue) Then
                            If IsError(adjustedValue) Or Not IsNumeric(adjustedValue) Then _
                                Err.Raise vbObjectError + 608, , _
                                "Non-numeric trial value at row " & firstData + rowNum - 1 & _
                                ", column " & colNum
                            Set destinationCell = exported.Cells(firstData + rowNum - 1, colNum)
                            destinationCell.Value2 = adjustedValue
                            destinationCell.Interior.Color = RGB(255, 235, 115)
                            ' Keep source style, structure and remaining formulas.
                            logSheet.Cells(outRow, 1).Value2 = firstData + rowNum - 1
                            logSheet.Cells(outRow, 2).Value2 = destinationCell.Address(False, False)
                            logSheet.Cells(outRow, 3).Value2 = currentModel
                            logSheet.Cells(outRow, 4).Value2 = currentKey
                            logSheet.Cells(outRow, 5).Value2 = CStr(h40Data(rowNum, 1))
                            logSheet.Cells(outRow, 6).Value2 = Format$(monthKeys(m) \ 100, "0000") & "_" & _
                                Format$(monthKeys(m) Mod 100, "00")
                            logSheet.Cells(outRow, 7).Value2 = CStr(blockNames(blockIndex))
                            If Not IsError(originalValue) Then logSheet.Cells(outRow, 8).Value2 = originalValue
                            logSheet.Cells(outRow, 9).Value2 = adjustedValue
                            If Not IsError(originalValue) Then
                                If IsNumeric(originalValue) Then _
                                    logSheet.Cells(outRow, 10).Value2 = CDbl(adjustedValue) - CDbl(originalValue)
                            End If
                            outRow = outRow + 1
                        End If
                    End If
                End If
            Next rowNum
        Next m
    Next blockIndex

    n = outRow - firstLogRow
    currentStage = "Formatting review output"
    logSheet.Range("L1").Value2 = "FY"
    logSheet.Range("M1").Value2 = targetFiscalYear
    logSheet.Range("L2").Value2 = "Effective changed cells"
    logSheet.Range("M2").Value2 = n
    logSheet.Range("L3").Value2 = "Exported at"
    logSheet.Range("M3").Value2 = Now
    logSheet.Range("M3").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    logSheet.Columns("A:J").AutoFit
    logSheet.Columns("H:J").NumberFormat = "#,##0.#####;[Red]-#,##0.#####;0"
    logSheet.Rows("1:3").WrapText = True
    logSheet.Range("A3:J3").AutoFilter
    exported.Activate
    exported.Range("A1").Select
    ' Workbook has no Calculate method in Excel VBA; Worksheet does.
    currentStage = "Recalculating exported Sheet1"
    exported.Calculate
    logSheet.Calculate

    currentStage = "Saving standalone .xlsx output"
    outputBook.SaveAs Filename:=CStr(outputPath), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    MsgBox "Budget export completed." & vbCrLf & _
           "Changed input cells highlighted: " & Format$(n, "#,##0") & vbCrLf & _
           "Saved to: " & CStr(outputPath), vbInformation, "V5 Export"
    Exit Sub

Failed:
    failureMessage = "V5 export failed during " & currentStage & _
        " (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    MsgBox failureMessage & vbCrLf & _
        "Source and trial were not modified; no successful export was reported.", vbExclamation
End Sub

Private Function ExportFindSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ExportFindSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ExportSourceSheet() As Worksheet
    Set ExportSourceSheet = ExportFindSheet(EX_SOURCE)
    If ExportSourceSheet Is Nothing Then Set ExportSourceSheet = ExportFindSheet(EX_FALLBACK)
End Function

Private Function ExportIsDifferent(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then
        ExportIsDifferent = (IsError(a) Xor IsError(b))
        Exit Function
    End If
    If Len(Trim$(CStr(a))) = 0 And Len(Trim$(CStr(b))) = 0 Then Exit Function
    If IsNumeric(a) And IsNumeric(b) Then
        ExportIsDifferent = (Abs(CDbl(a) - CDbl(b)) > EX_MARGIN * _
            Application.Max(1#, Abs(CDbl(a)), Abs(CDbl(b))))
    Else
        ExportIsDifferent = (CStr(a) <> CStr(b))
    End If
End Function

Private Sub ExportCheckAlignment(ByVal src As Worksheet, ByVal trial As Worksheet, _
    ByVal firstData As Long, ByVal lastData As Long)
    Dim r As Long
    For r = firstData To lastData
        If CStr(src.Cells(r, EX_H40_COL).Value2) <> CStr(trial.Cells(r, EX_H40_COL).Value2) Or _
           CStr(src.Cells(r, 2).Value2) <> CStr(trial.Cells(r, 2).Value2) Then _
            Err.Raise vbObjectError + 609, , _
                "Trial/source row alignment mismatch at row " & r & ". Refresh trial."
    Next r
End Sub

Private Function ExportParseMonth(ByVal v As Variant) As Long
    Dim re As Object, matches As Object, textValue As String
    On Error GoTo Invalid
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ExportParseMonth = Year(CDate(v)) * 100 + Month(CDate(v))
        Exit Function
    End If
    textValue = CStr(v)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(20[0-9]{2})[_/\-](1[0-2]|0?[1-9])([^0-9]|$)"
    If re.Test(textValue) Then
        Set matches = re.Execute(textValue)
        ExportParseMonth = CLng(matches(0).SubMatches(0)) * 100 + _
            CLng(matches(0).SubMatches(1))
    End If
    Exit Function
Invalid:
    ExportParseMonth = 0
End Function

Private Sub ExportGetMonths(ByVal src As Worksheet, ByRef keys() As Long, ByRef header As Long)
    Dim rr As Long, offset As Long, count As Long, best As Long, monthKey As Long
    For rr = 1 To 25
        count = 0
        For offset = 0 To EX_MONTHS - 1
            monthKey = ExportParseMonth(src.Cells(rr, EX_SALES_COL + offset).Value2)
            If monthKey = 0 Then monthKey = ExportParseMonth(src.Cells(rr, 24 + offset).Value2)
            If monthKey > 0 Then count = count + 1
        Next offset
        If count > best Then best = count: header = rr
    Next rr
    If best < 12 Then Err.Raise vbObjectError + 610, , "20-month header not found."
    For offset = 0 To EX_MONTHS - 1
        keys(offset) = ExportParseMonth(src.Cells(header, EX_SALES_COL + offset).Value2)
        If keys(offset) = 0 Then keys(offset) = ExportParseMonth(src.Cells(header, 24 + offset).Value2)
        If keys(offset) = 0 Then Err.Raise vbObjectError + 611, , _
            "Month missing in source header for offset " & offset
    Next offset
End Sub

Private Function ExportFirstDataRow(ByVal src As Worksheet, _
    ByVal header As Long, ByVal lastData As Long) As Long
    Dim r As Long, m As Long, block As Variant, v As Variant
    For r = header + 1 To lastData
        If Len(Trim$(CStr(src.Cells(r, EX_H40_COL).Value2))) > 0 Then
            For Each block In Array(24, 44, 64, 84, 104)
                For m = 0 To EX_MONTHS - 1
                    v = src.Cells(r, CLng(block) + m).Value2
                    If Not IsError(v) Then
                        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
                            ExportFirstDataRow = r
                            Exit Function
                        End If
                    End If
                Next m
            Next block
        End If
    Next r
    Err.Raise vbObjectError + 612, , "Could not find a numeric H40 data row."
End Function


' Audit external links arising from worksheet-only copy.
' Never break workbook links: BreakLink would turn real formulas into values.
' An external defined name is removed only if it is not referenced in ANY
' worksheet formula, validation formula, chart series or conditional format.
' For other links, stop with actionable diagnostic rather than corrupt output.
Private Sub ExportAuditCopiedLinks(ByVal src As Worksheet, _
    ByVal outputBook As Workbook, ByVal copiedSheet As Worksheet)
    Dim sourceLinks As Variant, destLinks As Variant
    Dim nm As Name, externalFormulaCount As Long, exampleCell As String
    Dim srcFormula As String, formulaCells As Range, c As Range
    Dim linkExample As String

    sourceLinks = src.Parent.LinkSources(Type:=xlExcelLinks)
    destLinks = outputBook.LinkSources(Type:=xlExcelLinks)
    If Not IsArray(destLinks) Then Exit Sub

    On Error Resume Next
    linkExample = CStr(destLinks(LBound(destLinks)))
    On Error GoTo 0

    ' Any external references in output cell formulas block automatic cleanup.
    On Error Resume Next
    Set formulaCells = copiedSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not formulaCells Is Nothing Then
        For Each c In formulaCells.Cells
            srcFormula = CStr(c.Formula)
            If InStr(1, srcFormula, "[", vbBinaryCompare) > 0 And _
               InStr(1, srcFormula, "]", vbBinaryCompare) > 0 Then
                externalFormulaCount = externalFormulaCount + 1
                If Len(exampleCell) = 0 Then exampleCell = c.Address(False, False)
            End If
        Next c
    End If
    If externalFormulaCount > 0 Then _
        Err.Raise vbObjectError + 607, , _
        "Copied Sheet1 contains " & externalFormulaCount & _
        " formulas with workbook-style references (example " & exampleCell & ")." & _
        " Check formula dependencies before exporting."

    If IsArray(sourceLinks) Then _
        Err.Raise vbObjectError + 613, , _
        "The source workbook itself lists external links (example: " & _
        Left$(linkExample, 180) & "). Review the original workbook dependencies."

    ' BudgetH40Options is a tool-only dropdown name defined by V4 for
    ' Budget_Detail!B3, not part of the customer's budget Sheet1.
    ' It can become an external name when copying Budget_Data alone.
    ' Remove this one exact known helper ONLY from the new output workbook.
    ' Do not use Workbooks.BreakLink or modify the source workbook.
    ExportRemoveToolOnlyH40Name outputBook
    destLinks = outputBook.LinkSources(Type:=xlExcelLinks)
    If Not IsArray(destLinks) Then Exit Sub

    ' Any other remaining copied name is not automatically discarded.
    ' Report its identity so that genuine budget dependencies stay intact.
    Dim nmSample As String, nmCount As Long, refersText As String
    For Each nm In outputBook.Names
        refersText = vbNullString
        On Error Resume Next
        refersText = nm.RefersTo
        On Error GoTo 0
        If InStr(1, refersText, "[", vbBinaryCompare) > 0 And _
           InStr(1, refersText, "]", vbBinaryCompare) > 0 Then
            nmCount = nmCount + 1
            If Len(nmSample) = 0 Then nmSample = nm.Name
        End If
    Next nm
    For Each nm In copiedSheet.Names
        refersText = vbNullString
        On Error Resume Next
        refersText = nm.RefersTo
        On Error GoTo 0
        If InStr(1, refersText, "[", vbBinaryCompare) > 0 And _
           InStr(1, refersText, "]", vbBinaryCompare) > 0 Then
            nmCount = nmCount + 1
            If Len(nmSample) = 0 Then nmSample = nm.Name
        End If
    Next nm
    If nmCount > 0 Then _
        Err.Raise vbObjectError + 614, , _
        "Copy introduced " & nmCount & " external defined-name references" & _
        " (example: " & nmSample & "). Not removed automatically because" & _
        " validation/chart dependencies could use them."

    Err.Raise vbObjectError + 615, , _
        "Copied sheet has external link(s) but no cell formula reference found." & _
        " First linked workbook: " & Left$(linkExample, 180) & _
        ". May be chart, object, connection or hidden name."
End Sub


Private Sub ExportRemoveToolOnlyH40Name(ByVal outBook As Workbook)
    Dim idx As Long, nm As Name, shortName As String
    Dim refersText As String, splitPos As Long

    ' Loop backwards because deleting a name changes the Names collection.
    For idx = outBook.Names.Count To 1 Step -1
        Set nm = outBook.Names(idx)
        shortName = nm.Name
        splitPos = InStrRev(shortName, "!")
        If splitPos > 0 Then shortName = Mid$(shortName, splitPos + 1)
        If StrComp(shortName, "BudgetH40Options", vbTextCompare) = 0 Then
            refersText = vbNullString
            On Error Resume Next
            refersText = nm.RefersTo
            On Error GoTo 0
            ' Only remove the helper when its copy points outside this workbook.
            If InStr(1, refersText, "[", vbBinaryCompare) > 0 And _
               InStr(1, refersText, "]", vbBinaryCompare) > 0 Then
                nm.Delete
            End If
        End If
    Next idx
End Sub
