Attribute VB_Name = "BudgetOptimizer_V4"
Option Explicit

' Budget Optimizer M1/M2 Sheet1 adapter - TEST VERSION
' All modifications are made on Budget_Trial, never on Sheet1/Budget_Data.
' Import into an .xlsm workbook. Do NOT upload real company data to GitHub.

Private Const SOURCE_LOCAL As String = "Sheet1"
Private Const SOURCE_IMPORTED As String = "Budget_Data"
Private Const TRIAL_SHEET As String = "Budget_Trial"
Private Const PANEL_SHEET As String = "Budget_Detail"
Private Const PLAN_SHEET As String = "Edit_Plan"
Private Const BASE_SHEET As String = "Base26"
Private Const LIST_SHEET As String = "H40_List"
Private Const FIRST_MONTH_COL As Long = 24      ' Set QTY: X
Private Const SALES_AMT_COL As Long = 127       ' Sales Amt: DW
Private Const H40_COL As Long = 19             ' S
Private Const MONTH_COUNT As Long = 20
Private Const OVERVIEW_SHEET As String = "Budget_Overview"
Private Const V4_TARGET_SHEET As String = "H40_Targets"
Private Const V4_HEADER_ROW As Long = 22
Private Const V4_START_ROW As Long = 23
Private Const V4_MAX_ROWS As Long = 4000
Private Const MODEL_START_ROW As Long = 40
Private Const MODEL_HEADER_ROW As Long = 39
Private Const HELP_FIRST As Long = 147
Private Const FMT_WHOLE As String = "#,##0;[Red]-#,##0;0"
Private Const FMT_PRICE As String = "#,##0.#####;[Red]-#,##0.#####;0"
Private Const FMT_PERCENT As String = "0.00%;[Red]-0.00%;0.00%"
Private mV3Stage As String
Public gBudgetV4Events As C_BudgetV4Events
Private gV4Busy As Boolean

' One-time initialization. Does not overwrite populated existing panels.
Public Sub SetupBudgetOptimizerSheet1()
    Dim src As Worksheet, p As Worksheet, pl As Worksheet, b As Worksheet
    On Error GoTo Failed
    MigrateOldPanel
    Set src = SourceSheet()
    If src Is Nothing Then
        MsgBox "No Sheet1 or Budget_Data found. Run ImportBudgetSheet1 first, or place the macro in the .xlsm containing Sheet1.", vbExclamation
        Exit Sub
    End If
    Set p = EnsureSheet(PANEL_SHEET)
    Set pl = EnsureSheet(PLAN_SHEET)
    Set b = EnsureSheet(BASE_SHEET)
    Call EnsureSheet(LIST_SHEET)
    If Len(CStr(p.Range("A1").Value2)) = 0 Then
        p.Range("A1").Value2 = "BUDGET OPTIMIZER  |  Sheet1 M1/M2 trial"
        p.Range("A3").Value2 = "H40"
        p.Range("D3").Value2 = "Fiscal year (Apr-Mar)"
        p.Range("E3").Value2 = 2027
        p.Range("A5:F5").Value = Array("Metric", "26F full-FY (manual + delta)", "26F in-window Apr-Mar", "Focus FY in-window", "Target (optional)", "Target difference")
        p.Range("A6:A9").Value = Application.Transpose(Array("Demand Amt", "Demand Qty", "Sales Amt", "Sales Qty"))
        p.Range("A11").Value2 = "26F full Sales Amt"
        p.Range("A12").Value2 = "Focus FY Sales Amt"
        p.Range("A13").Value2 = "YoY"
        p.Range("D11").Value2 = "All H40 focus-FY Sales Amt"
        p.Range("D12").Value2 = "Run status"
        p.Range("A15").Value2 = "Edit 26F annual baselines in Base26 B:E. Only current in-window differences are added."
        p.Range("A16").Value2 = "Trial changes come from Edit_Plan; the source Sheet1/Budget_Data is never edited."
        p.Range("A17").Value2 = "Rolling month changes between file versions are NOT yet reconciled automatically."
        p.Columns("A").ColumnWidth = 31
        p.Columns("B:F").ColumnWidth = 23
        p.Range("A1:F1").Font.Bold = True
        p.Range("A5:F5").Font.Bold = True
        p.Range("B3").Interior.Color = RGB(255, 250, 200)
        p.Range("E3").Interior.Color = RGB(255, 250, 200)
        p.Range("E6:E9").Interior.Color = RGB(255, 250, 200)
        p.Range("B13").NumberFormat = "0.00%"
    End If
    If Len(CStr(pl.Range("A1").Value2)) = 0 Then
        pl.Range("A1:I1").Value = Array("Enable Y", "Sheet row", "From YYYY_MM", "To YYYY_MM", "New Pcs/Set", "New S/P", "New Share", "New Comp Ratio", "Expected H40")
        pl.Rows(1).Font.Bold = True
        pl.Columns("A:I").ColumnWidth = 19
        pl.Columns("I").ColumnWidth = 30
    End If
    If Len(CStr(b.Range("A1").Value2)) = 0 Then
        b.Range("A1:E1").Value = Array("H40", "26F Demand Amt", "26F Demand Qty", "26F Sales Amt", "26F Sales Qty")
        b.Rows(1).Font.Bold = True
        b.Columns("A:E").ColumnWidth = 25
    End If
    RefreshH40List src, p, b
    EnsureOverview src
    FormatBudgetV3Display
    p.Activate
    MsgBox "Setup complete. Choose H40 in Budget_Panel!B3, enter 26F annual baseline in Base26, then run RefreshBudgetSheet1Preview.", vbInformation
    Exit Sub
Failed:
    MsgBox "Setup failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Optional: pull a pasted Sheet1 from ANOTHER Excel workbook into this tool.
' Source file is opened read-only. Any old imported Budget_Data is replaced only
' after explicit confirmation. Neither external source nor user's Sheet1 is edited.
Public Sub ImportBudgetSheet1()
    Dim f As Variant, wb As Workbook, ws As Worksheet, old As Worksheet
    Dim prevAlerts As Boolean, prevScreen As Boolean
    On Error GoTo Failed
    f = Application.GetOpenFilename("Excel Workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "Choose workbook containing Sheet1")
    If VarType(f) = vbBoolean Then Exit Sub
    If CStr(f) = ThisWorkbook.FullName Then
        MsgBox "You selected the tool workbook itself. No import needed if Sheet1 is already here.", vbInformation
        Exit Sub
    End If
    Set old = FindSheet(SOURCE_IMPORTED)
    If Not old Is Nothing Then
        If MsgBox("Replace existing Budget_Data snapshot? Base26 and Edit_Plan are preserved.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(SOURCE_LOCAL)
    On Error GoTo Failed
    If ws Is Nothing Then Err.Raise vbObjectError + 100, , "Selected workbook has no Sheet1"
    If Not old Is Nothing Then
        Application.DisplayAlerts = False
        old.Delete
        Application.DisplayAlerts = prevAlerts
    End If
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = SOURCE_IMPORTED
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.ScreenUpdating = prevScreen
    SetupBudgetOptimizerSheet1
    MsgBox "Sheet1 imported into Budget_Data. Cross-workbook formulas, if any, still need verification.", vbInformation
    Exit Sub
Failed:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Import failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Recalculate a fully separate trial copy. No direct writes to source.
Public Sub RefreshBudgetSheet1Preview(Optional ByVal silent As Boolean = False)
    Dim src As Worksheet, trial As Worksheet, panel As Worksheet, pl As Worksheet, b As Worksheet
    Dim keys(0 To MONTH_COUNT - 1) As Long, headerRow As Long, firstDataRow As Long, lastRow As Long
    Dim h40 As String, focusYear As Long, r As Long, m As Long, k As Long, br As Long
    Dim cur26(1 To 4) As Double, old26(1 To 4) As Double
    Dim curYear(1 To 4) As Double, allYear(1 To 4) As Double
    Dim vals(1 To 4) As Double, numEdits As Long, found As Boolean, n As Long
    Dim baseVal As Variant, origAlerts As Boolean
    On Error GoTo Failed
    Set src = SourceSheet()
    If src Is Nothing Then Err.Raise vbObjectError + 101, , "Sheet1/Budget_Data not found. Run ImportBudgetSheet1 or add Sheet1."
    Set panel = FindSheet(PANEL_SHEET)
    Set pl = FindSheet(PLAN_SHEET)
    Set b = FindSheet(BASE_SHEET)
    If panel Is Nothing Or pl Is Nothing Or b Is Nothing Then
        SetupBudgetOptimizerSheet1
        Set panel = FindSheet(PANEL_SHEET)
        Set pl = FindSheet(PLAN_SHEET)
        Set b = FindSheet(BASE_SHEET)
    End If
    ' Upgrade labels in existing panels too (Setup skips pre-existing sheets).
    panel.Range("D3").Value2 = "Fiscal year (Apr-Mar)"
    panel.Range("B5").Value2 = "26F full-FY (manual + delta)"
    panel.Range("C5").Value2 = "26F in-window Apr-Mar"
    panel.Range("D5").Value2 = "Focus FY in-window"
    panel.Range("A12").Value2 = "Focus FY Sales Amt"
    panel.Range("D11").Value2 = "All H40 focus-FY Sales Amt"
    h40 = Trim$(CStr(panel.Range("B3").Value2))
    If Len(h40) = 0 Then Err.Raise vbObjectError + 102, , "Choose H40 in Budget_Panel!B3"
    If Not IsNumeric(panel.Range("E3").Value2) Then Err.Raise vbObjectError + 103, , "Focus year must be numeric"
    focusYear = CLng(panel.Range("E3").Value2)
    If focusYear < 2026 Or focusYear > 2100 Then Err.Raise vbObjectError + 104, , "Invalid focus year"
    FindRollingMonths src, keys, headerRow
    firstDataRow = FindFirstBudgetDataRow(src, headerRow)
    lastRow = src.Cells(src.Rows.Count, H40_COL).End(xlUp).Row
    If lastRow < firstDataRow Then Err.Raise vbObjectError + 105, , "No data rows below month header"
    RefreshH40List src, panel, b
    EnsureOverview src
    FormatBudgetV3Display
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    origAlerts = Application.DisplayAlerts
    Set trial = FindSheet(TRIAL_SHEET)
    If Not trial Is Nothing Then
        Application.DisplayAlerts = False
        trial.Delete
        Application.DisplayAlerts = origAlerts
    End If
    src.Copy After:=src
    Set trial = ActiveSheet
    trial.Name = TRIAL_SHEET
    For r = 2 To pl.Cells(pl.Rows.Count, "B").End(xlUp).Row
        If UCase$(Trim$(CStr(pl.Cells(r, "A").Value2))) = "Y" Then
            ApplyPlan trial, pl, r, firstDataRow, lastRow, keys
            numEdits = numEdits + 1
        End If
    Next r
    Dim appliedSetCells As Long
    appliedSetCells = ApplyOverviewSetQty(trial, src, firstDataRow, lastRow, keys)
    BuildTrialSummaryFormulas trial, panel, firstDataRow, lastRow, keys, focusYear
    trial.Calculate
    src.Calculate
    For r = firstDataRow To lastRow
        If IsBudgetDataRow(trial, r) And StrComp(Trim$(CStr(trial.Cells(r, H40_COL).Value2)), h40, vbTextCompare) = 0 Then
            found = True
            For m = 0 To MONTH_COUNT - 1
                ReadMetrics trial, r, m, vals
                If FiscalYear(keys(m)) = 2026 Then AddMetrics cur26, vals
                If FiscalYear(keys(m)) = focusYear Then AddMetrics curYear, vals
            Next m
        End If
        If IsBudgetDataRow(trial, r) Then
        For m = 0 To MONTH_COUNT - 1
            If FiscalYear(keys(m)) = focusYear Then
                ReadMetrics trial, r, m, vals
                AddMetrics allYear, vals
            End If
        Next m
        End If
    Next r
    If Not found Then Err.Raise vbObjectError + 106, , "H40 not found in S column: " & h40
    For r = firstDataRow To lastRow
        If IsBudgetDataRow(src, r) And StrComp(Trim$(CStr(src.Cells(r, H40_COL).Value2)), h40, vbTextCompare) = 0 Then
            For m = 0 To MONTH_COUNT - 1
                If FiscalYear(keys(m)) = 2026 Then
                    ReadMetrics src, r, m, vals
                    AddMetrics old26, vals
                End If
            Next m
        End If
    Next r
    br = FindBaselineRow(b, h40)
    For k = 1 To 4
        panel.Cells(k + 5, "C").Value2 = cur26(k)
        panel.Cells(k + 5, "D").Value2 = curYear(k)
        panel.Cells(k + 5, "B").ClearContents
        If br > 0 Then
            baseVal = b.Cells(br, k + 1).Value2
            If Not IsError(baseVal) And Len(Trim$(CStr(baseVal))) > 0 Then
                If IsNumeric(baseVal) Then
                    panel.Cells(k + 5, "B").Value2 = CDbl(baseVal) + cur26(k) - old26(k)
                End If
            End If
        End If
        panel.Cells(k + 5, "F").Formula = "=IF(E" & (k + 5) & "="""","""",D" & (k + 5) & "-E" & (k + 5) & ")"
    Next k
    panel.Range("B11").Formula = "=B8"
    panel.Range("B12").Formula = "=D8"
    panel.Range("B13").Formula = "=IFERROR(B12/B11-1,"""")"
    panel.Range("E11").Formula = "=SUM('" & TRIAL_SHEET & "'!" & trial.Range(trial.Cells(firstDataRow, HELP_FIRST + 3), trial.Cells(lastRow, HELP_FIRST + 3)).Address(False, False) & ")"
    LinkLiveDetail panel, trial, firstDataRow, lastRow
    LinkOverviewSummary trial, firstDataRow, lastRow
    panel.Range("E12").Value2 = "Trial refreshed; " & numEdits & " edit row(s); Set QTY cells linked: " & appliedSetCells
    If Not FindSheet(OVERVIEW_SHEET) Is Nothing Then FindSheet(OVERVIEW_SHEET).Range("B11").Value2 = "Set QTY cells linked: " & appliedSetCells
    Application.Calculate
    If Not silent Then
        panel.Activate
        MsgBox "Preview complete. Budget_Overview and Budget_Detail share the same trial; source unchanged.", vbInformation
    End If
Cleanup:
    Application.DisplayAlerts = origAlerts
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Dim message As String
    message = "Preview failed (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If silent Then
        Err.Raise vbObjectError + 209, "RefreshBudgetSheet1Preview", message
    Else
        MsgBox message, vbExclamation
    End If
End Sub

' Detect first actual data row after multiple date/title/memo header rows.
' Only matches rows with a nonblank H40 in S and at least two numeric budget inputs.
Private Function FindFirstBudgetDataRow(ByVal ws As Worksheet, ByVal monthHeaderRow As Long) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, H40_COL).End(xlUp).Row
    For r = monthHeaderRow + 1 To Application.Min(lastR, monthHeaderRow + 100)
        If IsBudgetDataRow(ws, r) Then
            FindFirstBudgetDataRow = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 143, , "Could not locate first numeric budget row below the month header at row " & monthHeaderRow & ". Check S H40 and budget columns."
End Function

' A row is eligible when S holds a H40 and two independent input blocks
' contain numeric values. Prevents labels like Set Number +2M being read as data.
Private Function IsBudgetDataRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim s As String, m As Long, countBlocks As Long, blockNo As Long
    Dim bases As Variant, v As Variant
    s = Trim$(CStr(ws.Cells(r, H40_COL).Value2))
    If Len(s) = 0 Then Exit Function
    If LCase$(s) = "h40" Or LCase$(s) = "eop" Then Exit Function
    bases = Array(24, 44, 64, 84, 104)
    For blockNo = 0 To 4
        For m = 0 To MONTH_COUNT - 1
            v = ws.Cells(r, CLng(bases(blockNo)) + m).Value2
            If Not IsError(v) Then
                If Len(Trim$(CStr(v))) > 0 And IsNumeric(v) Then
                    countBlocks = countBlocks + 1
                    Exit For
                End If
            End If
        Next m
        If countBlocks >= 2 Then
            IsBudgetDataRow = True
            Exit Function
        End If
    Next blockNo
End Function

Private Sub ApplyPlan(ByVal ws As Worksheet, ByVal pl As Worksheet, ByVal planRow As Long, ByVal firstData As Long, ByVal lastData As Long, ByRef months() As Long)
    Dim rowNum As Long, fromMonth As Long, toMonth As Long, mm As Long
    Dim raw As Variant, h40 As String
    rowNum = CLng(Val(pl.Cells(planRow, "B").Value2))
    If rowNum < firstData Or rowNum > lastData Then Err.Raise vbObjectError + 120, , "Invalid row in Edit_Plan row " & planRow
    fromMonth = ParseMonth(pl.Cells(planRow, "C").Value2)
    toMonth = ParseMonth(pl.Cells(planRow, "D").Value2)
    If fromMonth = 0 Or toMonth = 0 Or fromMonth > toMonth Then Err.Raise vbObjectError + 121, , "Invalid month range in Edit_Plan row " & planRow
    h40 = Trim$(CStr(pl.Cells(planRow, "I").Value2))
    If Len(h40) > 0 Then
        If StrComp(Trim$(CStr(ws.Cells(rowNum, H40_COL).Value2)), h40, vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 122, , "H40 does not match in Edit_Plan row " & planRow
    End If
    For mm = 0 To MONTH_COUNT - 1
        If months(mm) >= fromMonth And months(mm) <= toMonth Then
            raw = pl.Cells(planRow, "E").Value2
            If Len(Trim$(CStr(raw))) > 0 Then
                If Not IsNumeric(raw) Then Err.Raise vbObjectError + 123, , "Invalid Pcs/Set"
                If CDbl(raw) < 0 Or CDbl(raw) <> Fix(CDbl(raw)) Then Err.Raise vbObjectError + 124, , "Pcs/Set must be nonnegative integer"
                ws.Cells(rowNum, 44 + mm).Value2 = CDbl(raw)
            End If
            raw = pl.Cells(planRow, "F").Value2
            If Len(Trim$(CStr(raw))) > 0 Then
                If Not IsNumeric(raw) Then Err.Raise vbObjectError + 125, , "Invalid S/P"
                If CDbl(raw) < 0 Or Abs(CDbl(raw) * 100000 - Round(CDbl(raw) * 100000, 0)) > 0.000001 Then _
                    Err.Raise vbObjectError + 126, , "S/P must be nonnegative, <=5 decimals"
                ws.Cells(rowNum, 84 + mm).Value2 = CDbl(raw)
            End If
            raw = pl.Cells(planRow, "G").Value2
            If Len(Trim$(CStr(raw))) > 0 Then
                If Not IsNumeric(raw) Then Err.Raise vbObjectError + 127, , "Invalid Share"
                If CDbl(raw) < 0 Or CDbl(raw) > 100 Or CDbl(raw) <> Fix(CDbl(raw)) Then _
                    Err.Raise vbObjectError + 128, , "Share must be a whole number from 0 to 100"
                ws.Cells(rowNum, 104 + mm).Value2 = CDbl(raw)
            End If
            raw = pl.Cells(planRow, "H").Value2
            If Len(Trim$(CStr(raw))) > 0 Then
                If Not IsNumeric(raw) Then Err.Raise vbObjectError + 129, , "Invalid Comp Ratio"
                If CDbl(raw) < 0 Or CDbl(raw) > 100 Or CDbl(raw) <> Fix(CDbl(raw)) Then _
                    Err.Raise vbObjectError + 130, , "Comp Ratio must be a whole number from 0 to 100"
                ws.Cells(rowNum, 64 + mm).Value2 = CDbl(raw)
            End If
        End If
    Next mm
End Sub

Private Sub ReadMetrics(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal mm As Long, ByRef o() As Double)
    Dim sets As Double, pieces As Double, comp As Double, sp As Double, share As Double
    Dim dq As Double, sq As Double, amount As Variant
    sets = SafeNumber(ws.Cells(rowNum, 24 + mm).Value2)
    pieces = SafeNumber(ws.Cells(rowNum, 44 + mm).Value2)
    comp = SafeNumber(ws.Cells(rowNum, 64 + mm).Value2)
    sp = SafeNumber(ws.Cells(rowNum, 84 + mm).Value2)
    share = SafeNumber(ws.Cells(rowNum, 104 + mm).Value2)
    dq = sets * pieces * comp / 100#
    sq = dq * share / 100#
    o(1) = WorksheetFunction.Round(dq * sp, 2)
    o(2) = dq
    amount = ws.Cells(rowNum, SALES_AMT_COL + mm).Value2
    If IsError(amount) Then
        Err.Raise vbObjectError + 131, , "Sales Amt Excel error at " & ws.Cells(rowNum, SALES_AMT_COL + mm).Address(False, False)
    End If
    If IsNumeric(amount) And Len(CStr(amount)) > 0 Then
        o(3) = CDbl(amount)
    Else
        o(3) = WorksheetFunction.Round(sq * sp, 2)
    End If
    o(4) = sq
End Sub

Private Function SafeNumber(ByVal v As Variant) As Double
    If IsError(v) Then Err.Raise vbObjectError + 132, , "Source cell has an Excel error"
    If Len(Trim$(CStr(v))) = 0 Then SafeNumber = 0: Exit Function
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 133, , "Non-numeric value in budget numeric block: " & Left$(CStr(v), 50)
    SafeNumber = CDbl(v)
End Function

Private Sub AddMetrics(ByRef target() As Double, ByRef v() As Double)
    Dim k As Long
    For k = 1 To 4
        target(k) = target(k) + v(k)
    Next k
End Sub

Private Sub FindRollingMonths(ByVal ws As Worksheet, ByRef months() As Long, ByRef headerRow As Long)
    Dim r As Long, m As Long, cnt As Long, best As Long, bestRow As Long, key As Long
    best = -1
    For r = 1 To 25
        cnt = 0
        For m = 0 To MONTH_COUNT - 1
            key = ParseMonth(ws.Cells(r, SALES_AMT_COL + m).Value2)
            If key = 0 Then key = ParseMonth(ws.Cells(r, FIRST_MONTH_COL + m).Value2)
            If key <> 0 Then cnt = cnt + 1
        Next m
        If cnt > best Then best = cnt: bestRow = r
    Next r
    If best < 12 Then Err.Raise vbObjectError + 140, , "Could not find rolling 20-month header in rows 1:25 (found " & best & "/20)"
    headerRow = bestRow
    For m = 0 To MONTH_COUNT - 1
        months(m) = ParseMonth(ws.Cells(bestRow, SALES_AMT_COL + m).Value2)
        If months(m) = 0 Then months(m) = ParseMonth(ws.Cells(bestRow, FIRST_MONTH_COL + m).Value2)
        If months(m) = 0 Then Err.Raise vbObjectError + 141, , "Missing month header at offset " & m & " (header row " & bestRow & ")"
        If m > 0 Then
            If months(m) <> NextMonth(months(m - 1)) Then _
                Err.Raise vbObjectError + 142, , "Month header sequence not consecutive near offset " & m
        End If
    Next m
End Sub

' Fiscal year F begins in April of YYYY and ends in March of YYYY+1.
' E.g. 2027/03 = 26F, 2027/04 = 27F, 2028/03 = 27F.
Private Function FiscalYear(ByVal key As Long) As Long
    Dim y As Long, mm As Long
    y = key \ 100
    mm = key Mod 100
    If mm < 1 Or mm > 12 Then Err.Raise vbObjectError + 144, , "Invalid YYYYMM: " & key
    If mm >= 4 Then
        FiscalYear = y
    Else
        FiscalYear = y - 1
    End If
End Function

Private Function NextMonth(ByVal key As Long) As Long
    If key Mod 100 = 12 Then
        NextMonth = (key \ 100 + 1) * 100 + 1
    Else
        NextMonth = key + 1
    End If
End Function

Private Function ParseMonth(ByVal v As Variant) As Long
    Dim re As Object, ms As Object, s As String, yy As Long, mm As Long
    On Error GoTo Invalid
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ParseMonth = Year(CDate(v)) * 100 + Month(CDate(v))
        Exit Function
    End If
    s = Replace(CStr(v), vbLf, " ")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(20[0-9]{2})[_/\-](1[0-2]|0?[1-9])([^0-9]|$)"
    re.Global = False
    If re.Test(s) Then
        Set ms = re.Execute(s)
        yy = CLng(ms(0).SubMatches(0))
        mm = CLng(ms(0).SubMatches(1))
        ParseMonth = yy * 100 + mm
    End If
    Exit Function
Invalid:
    ParseMonth = 0
End Function

Private Function SourceSheet() As Worksheet
    ' Prefer latest explicitly imported snapshot, then a pasted local Sheet1.
    Set SourceSheet = FindSheet(SOURCE_IMPORTED)
    If SourceSheet Is Nothing Then Set SourceSheet = FindSheet(SOURCE_LOCAL)
End Function

Private Function FindSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set FindSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    Set EnsureSheet = FindSheet(name)
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function FindBaselineRow(ByVal ws As Worksheet, ByVal h40 As String) As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, "A").Value2)), h40, vbTextCompare) = 0 Then
            FindBaselineRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub RefreshH40List(ByVal src As Worksheet, ByVal p As Worksheet, ByVal b As Worksheet)
    Dim list As Worksheet, dict As Object, r As Long, last As Long
    Dim key As String, i As Long, keys As Variant, n As Long
    Set list = EnsureSheet(LIST_SHEET)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    last = src.Cells(src.Rows.Count, "S").End(xlUp).Row
    For r = 2 To last
        key = Trim$(CStr(src.Cells(r, "S").Value2))
        If IsBudgetDataRow(src, r) And Len(key) > 0 And LCase$(key) <> "h40" And LCase$(key) <> "eop" Then
            If Not dict.Exists(key) Then dict.Add key, key
        End If
    Next r
    If dict.Count = 0 Then Err.Raise vbObjectError + 160, , "No H40 found in S column"
    list.Cells.Clear
    list.Range("A1").Value2 = "H40"
    keys = dict.Keys
    For i = 0 To dict.Count - 1
        list.Cells(i + 2, "A").Value2 = keys(i)
    Next i
    On Error Resume Next
    ThisWorkbook.Names("BudgetH40Options").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="BudgetH40Options", RefersTo:="='" & LIST_SHEET & "'!$A$2:$A$" & (dict.Count + 1)
    On Error Resume Next
    p.Range("B3").Validation.Delete
    p.Range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BudgetH40Options"
    On Error GoTo 0
    If Len(Trim$(CStr(p.Range("B3").Value2))) = 0 Then p.Range("B3").Value2 = keys(0)
    For i = 0 To dict.Count - 1
        If FindBaselineRow(b, CStr(keys(i))) = 0 Then
            n = b.Cells(b.Rows.Count, "A").End(xlUp).Row + 1
            b.Cells(n, "A").Value2 = CStr(keys(i))
        End If
    Next i
End Sub


' ===== M2.5 / V3: model-level April-March Set QTY planning =====
' Overview B5: MANUAL 26F TTL Set QTY (never calculated from source)
' Overview B6: 27F TTL target; B7: 27F suggested plan (deduplicated model)
' Machine ID comes from column B; each machine's H40 rows link to E:P plan.
' User can edit E:P directly after generating. Trial references them via formulas.

Public Sub SetupBudgetOptimizerV3()
    On Error GoTo Failed
    MigrateOldPanel
    SetupBudgetOptimizerSheet1
    EnsureOverview SourceSheet()
    FindSheet(OVERVIEW_SHEET).Activate
    MsgBox "V3 ready: Budget_Overview (Set QTY + TTL Sales Amt) and Budget_Detail (H40). Enter 26F TTL Set QTY in B5; 27F target in B6.", vbInformation
    Exit Sub
Failed:
    MsgBox "V3 setup failed (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Private Sub MigrateOldPanel()
    Dim old As Worksheet
    If FindSheet(PANEL_SHEET) Is Nothing Then
        Set old = FindSheet("Budget_Panel")
        If Not old Is Nothing Then old.Name = PANEL_SHEET
    End If
End Sub

' Preserve existing monthly input values when upgrading V3 layout 36/37 -> 39/40.
' This is done only once. Do not clear/rebuild target values.
Private Sub UpgradeOverviewLayout(ByVal ov As Worksheet)
    Dim oldLast As Long, rr As Long, valueKey As String
    If CStr(ov.Range("S2").Value2) = "V3_LAYOUT_39" Then Exit Sub
    If CStr(ov.Range("A36").Value2) = "Model (source B)" Then
        oldLast = ov.Cells(ov.Rows.Count, "A").End(xlUp).Row
        ov.Rows("37:39").Insert Shift:=xlDown
        ' Old data now begins at row 40; keep its B and E:P values unchanged.
        If oldLast >= 37 Then
            For rr = MODEL_START_ROW To oldLast + 3
                valueKey = Trim$(CStr(ov.Cells(rr, 1).Value2))
                If Len(valueKey) > 0 Then ov.Cells(rr, 19).Value2 = valueKey
            Next rr
        End If
        ov.Range("A36:R38").ClearContents
    End If
    ov.Range("S2").Value2 = "V3_LAYOUT_39"
End Sub

' Presentation only: do not round underlying budget values or source cells.
' Reapply on existing sheets because Setup does not rebuild populated panels.
Private Sub FormatBudgetV3Display()
    Dim ov As Worksheet, detail As Worksheet, base As Worksheet, plan As Worksheet
    Set ov = FindSheet(OVERVIEW_SHEET)
    Set detail = FindSheet(PANEL_SHEET)
    Set base = FindSheet(BASE_SHEET)
    Set plan = FindSheet(PLAN_SHEET)

    If Not ov Is Nothing Then
        ov.Range("B5:B7").NumberFormat = FMT_WHOLE
        ov.Range("B8:B9").NumberFormat = FMT_PERCENT
        ov.Range("E5:E6").NumberFormat = FMT_WHOLE
        ov.Range("E7").NumberFormat = FMT_PERCENT
        ov.Range("E8").NumberFormat = FMT_WHOLE
        ov.Range("B" & MODEL_START_ROW & ":D" & ov.Rows.Count).NumberFormat = FMT_WHOLE
        ov.Range("E" & MODEL_START_ROW & ":Q" & ov.Rows.Count).NumberFormat = FMT_WHOLE
        ov.Range("U2:U13").NumberFormat = FMT_WHOLE
    End If
    If Not detail Is Nothing Then
        detail.Range("B6:F9").NumberFormat = FMT_WHOLE
        detail.Range("B11:B12").NumberFormat = FMT_WHOLE
        detail.Range("B13").NumberFormat = FMT_PERCENT
        detail.Range("E11").NumberFormat = FMT_WHOLE
    End If
    If Not base Is Nothing Then base.Range("B2:E" & base.Rows.Count).NumberFormat = FMT_WHOLE
    If Not plan Is Nothing Then
        plan.Range("E2:E" & plan.Rows.Count).NumberFormat = FMT_WHOLE  ' Pcs/Set
        plan.Range("F2:F" & plan.Rows.Count).NumberFormat = FMT_PRICE  ' S/P up to 5 decimals
        plan.Range("G2:H" & plan.Rows.Count).NumberFormat = FMT_WHOLE  ' Share and Comp Ratio
    End If
End Sub

Private Sub EnsureOverview(ByVal src As Worksheet)
    Dim ov As Worksheet, r As Long, last As Long, machine As String, seen As Object
    Dim nextR As Long, currentCode As String, currentLabel As String, labelByCode As Object, keys(0 To MONTH_COUNT - 1) As Long
    Dim header As Long, m As Long, fy As Long, first As Long, original As Double
    mV3Stage = "O01 ensure Overview worksheet"
    Set ov = EnsureSheet(OVERVIEW_SHEET)
    mV3Stage = "O02 initialize Overview headings"
    If Len(CStr(ov.Range("A1").Value2)) = 0 Then
        ov.Range("A1").Value2 = "BUDGET OVERVIEW  |  MODEL SET QTY + TTL SALES AMT"
        ov.Range("A3").Value2 = "FY (Apr-Mar)"
        ' A fresh V3 install may not have Budget_Detail yet.
        ' Use FY 2027 as a safe starting default; Setup/refresh can sync later.
        If FindSheet(PANEL_SHEET) Is Nothing Then
            ov.Range("B3").Value2 = 2027
        Else
            ov.Range("B3").Value2 = FindSheet(PANEL_SHEET).Range("E3").Value2
        End If
        ov.Range("A5").Value2 = "26F TTL Set QTY (manual)"
        ov.Range("A6").Value2 = "27F TTL target (manual)"
        ov.Range("A7").Value2 = "27F model plan TTL (unique)"
        ov.Range("A8").Value2 = "Set QTY YoY vs manual 26F"
        ov.Range("A9").Value2 = "Plan vs TTL target"
        ov.Range("A10").Value2 = "Plan status"
        ov.Range("D5").Value2 = "26F TTL Sales Amt (Base26)"
        ov.Range("D6").Value2 = "27F TTL Sales Amt (LIVE)"
        ov.Range("D7").Value2 = "Sales Amt YoY"
        ov.Range("D8").Value2 = "Trial vs original Sales Amt"
        ov.Range("A12").Value2 = "SEASONAL: Apr-Sep ascending; Oct drop; Nov-Jan ascending; Feb drop; Mar recovery."
        ov.Range("A13").Value2 = "1. Input 26F TTL Set QTY in B5 (only one manual total)."
        ov.Range("A14").Value2 = "2. Input 27F TTL target in B6; optional per-model targets in column B below."
        ov.Range("A15").Value2 = "3. Run GenerateSeasonalSetQtyPlanV3. Edit monthly E:P to fine-tune; cells link LIVE."
        ov.Range("A16").Value2 = "4. For H40 fine-tuning use Budget_Detail / Edit_Plan and run RefreshBudgetSheet1Preview."
        ov.Range("A17").Value2 = "Plan quantity is counted once per model, not once per H40."
        ov.Range("A18").Value2 = "Do not modify source Sheet1 / Budget_Data. Do not upload real data to public GitHub."
        ov.Range("A1:P1").Font.Bold = True
        ov.Range("B5:B6").Interior.Color = RGB(255, 250, 200)
        ov.Range("B3").Interior.Color = RGB(255, 250, 200)
        ov.Range("B8").NumberFormat = "0.00%"
        ov.Range("E7").NumberFormat = "0.00%"
        ov.Columns("A").ColumnWidth = 28
        ov.Columns("B:D").ColumnWidth = 20
        ov.Columns("E:P").ColumnWidth = 13
        ov.Columns("Q").ColumnWidth = 18
        ov.Columns("R").Hidden = True
    End If
    mV3Stage = "O03 build model table header"
    UpgradeOverviewLayout ov
    ov.Range("A39:D39").Value = Array("Model (source M)", "27F target (manual)", "Original 27F qty", "Planned 27F qty")
    For m = 0 To 11
        ov.Cells(MODEL_HEADER_ROW, 5 + m).Value2 = Format(DateSerial(CLng(ov.Range("B3").Value2), 4 + m, 1), "yyyy_mm")
    Next m
    ov.Range("Q39").Value2 = "Plan vs model target"
    ov.Range("R39").Value2 = "Source row"
    ov.Range("S39").Value2 = "Machine key (source B)"
    ov.Columns("S").Hidden = True
    ov.Range("A39:Q39").Font.Bold = True
    mV3Stage = "O04 initialize model dictionary"
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set labelByCode = CreateObject("Scripting.Dictionary")
    labelByCode.CompareMode = vbTextCompare
    last = src.Cells(src.Rows.Count, H40_COL).End(xlUp).Row
    mV3Stage = "O05 detect rolling month headers"
    FindRollingMonths src, keys, header
    fy = CLng(ov.Range("B3").Value2)
    nextR = MODEL_START_ROW
    currentCode = vbNullString
    mV3Stage = "O06 scan B column machine IDs"
    For r = header + 1 To last
        If Len(Trim$(CStr(src.Cells(r, 2).Value2))) > 0 Then _
            currentCode = Trim$(CStr(src.Cells(r, 2).Value2))
        If Len(currentCode) > 0 And Len(Trim$(CStr(src.Cells(r, 13).Value2))) > 0 Then _
            labelByCode(currentCode) = Trim$(CStr(src.Cells(r, 13).Value2))
        If IsBudgetDataRow(src, r) And Len(currentCode) > 0 Then
            If Not seen.Exists(currentCode) Then
                seen.Add currentCode, nextR
                If Len(Trim$(CStr(ov.Cells(nextR, 19).Value2))) = 0 Then
                    ov.Cells(nextR, 19).Value2 = currentCode
                    ov.Cells(nextR, 18).Value2 = r
                ElseIf StrComp(Trim$(CStr(ov.Cells(nextR, 19).Value2)), currentCode, vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 170, , "Model order changed. Preserve existing plan; machine key mismatch at Overview row " & nextR & "."
                End If
                currentLabel = currentCode
                If labelByCode.Exists(currentCode) Then currentLabel = labelByCode(currentCode)
                ov.Cells(nextR, 1).Value2 = currentLabel
                original = 0
                For m = 0 To MONTH_COUNT - 1
                    If FiscalYear(keys(m)) = fy Then original = original + SafeNumber(src.Cells(r, FIRST_MONTH_COL + m).Value2)
                Next m
                ov.Cells(nextR, 3).Value2 = original
                ov.Cells(nextR, 4).Formula = "=SUM(E" & nextR & ":P" & nextR & ")"
                ov.Cells(nextR, 17).Formula = "=IF(B" & nextR & "="""","""",D" & nextR & "-B" & nextR & ")"
                nextR = nextR + 1
            End If
        End If
    Next r
    ' Resolve Model names even when source M is populated later within the same B group.
    For m = MODEL_START_ROW To nextR - 1
        currentCode = Trim$(CStr(ov.Cells(m, 19).Value2))
        If labelByCode.Exists(currentCode) Then
            ov.Cells(m, 1).Value2 = labelByCode(currentCode)
        Else
            ov.Cells(m, 1).Value2 = "(Model missing in M)"
        End If
    Next m
    If nextR = MODEL_START_ROW Then Err.Raise vbObjectError + 171, , "No model IDs found in B for numeric H40 detail rows"
    mV3Stage = "O07 set integer validations"
    ApplyWholeNumberValidation ov.Range("B5:B6")
    ApplyWholeNumberValidation ov.Range("B" & MODEL_START_ROW & ":B" & nextR - 1)
    ApplyWholeNumberValidation ov.Range("E" & MODEL_START_ROW & ":P" & nextR - 1)
    ov.Range("B7").Formula = "=SUM(D" & MODEL_START_ROW & ":D" & nextR - 1 & ")"
    ov.Range("B8").Formula = "=IFERROR(B7/B5-1,"""")"
    ov.Range("B9").Formula = "=IFERROR(B7/B6-1,"""")"
    ov.Range("B8:B9").NumberFormat = "0.00%"
    ov.Range("E5").Formula = "=IF(COUNT(Base26!D2:D10000)=0,"""",SUM(Base26!D2:D10000))"
    ov.Range("E7").Formula = "=IFERROR(E6/E5-1,"""")"
    ov.Range("E7").NumberFormat = "0.00%"
    mV3Stage = "O08 prepare monthly chart data"
    PrepareMonthlyTotalData ov, nextR - 1
    mV3Stage = "O09 create trend chart"
    CreateSetQtyTrendChart ov
    FormatBudgetV3Display
End Sub

Private Sub PrepareMonthlyTotalData(ByVal ov As Worksheet, ByVal lastModel As Long)
    Dim m As Long, txt As String
    ov.Range("T1").Value2 = "Month"
    ov.Range("U1").Value2 = "27F TTL Set QTY"
    For m = 0 To 11
        txt = CStr(ov.Cells(MODEL_HEADER_ROW, m + 5).Value2)
        ov.Cells(m + 2, 20).Value2 = txt
        ov.Cells(m + 2, 21).Formula = "=SUM(" & ov.Cells(MODEL_START_ROW, m + 5).Address(False, False) & ":" & ov.Cells(lastModel, m + 5).Address(False, False) & ")"
    Next m
End Sub

Private Sub ApplyWholeNumberValidation(ByVal rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="999999999"
    rng.Validation.ErrorTitle = "Set QTY must be a whole number"
    rng.Validation.ErrorMessage = "Only nonnegative integer quantities are allowed."
    rng.Validation.ShowError = True
    On Error GoTo 0
End Sub

Private Sub CreateSetQtyTrendChart(ByVal ov As Worksheet)
    Dim co As ChartObject
    Dim topPos As Double, spaceBelow As Double, fitHeight As Double, fitWidth As Double
    On Error Resume Next
    Set co = ov.ChartObjects("V3_SetQTY_Trend")
    On Error GoTo 0

    topPos = ov.Range("A20").Top
    ' Keep a visible gap above row 39 where the model table starts.
    spaceBelow = ov.Range("A39").Top - topPos - 16#
    If spaceBelow < 110# Then
        Err.Raise vbObjectError + 195, , "Not enough vertical space for the Set QTY chart above row 39. Check row heights 20:38."
    End If
    fitHeight = Application.Min(230#, spaceBelow)
    fitWidth = Application.Min(860#, ov.Range("Q20").Left + ov.Range("Q20").Width - ov.Range("A20").Left - 12#)

    If co Is Nothing Then
        Set co = ov.ChartObjects.Add(ov.Range("A20").Left, topPos, fitWidth, fitHeight)
        co.Name = "V3_SetQTY_Trend"
    End If

    ' Always resize/reposition EXISTING charts too. The old version returned
    ' early, leaving a stretched chart overlaying the row-39 model table.
    ' xlFreeFloating prevents resizing when rows/columns are inserted.
    With co
        .Placement = xlFreeFloating
        .Left = ov.Range("A20").Left
        .Top = topPos
        .Width = fitWidth
        .Height = fitHeight
    End With
    With co.Chart
        .ChartType = xlLineMarkers
        .SetSourceData Source:=ov.Range("T1:U13")
        .HasTitle = True
        .ChartTitle.Text = "27F TTL Set QTY | monthly seasonality"
        .HasLegend = False
    End With
End Sub

Public Sub GenerateSeasonalSetQtyPlanV3()
    Dim src As Worksheet, ov As Worksheet, fy As Long, lastR As Long, r As Long, m As Long, i As Long
    Dim yearTarget As Double, fixedTotal As Double, remaining As Double, weightedTotal As Double
    Dim machineGoal As Double, totalSuggested As Double, targets() As Double, hasTarget() As Boolean
    Dim weight As Double, unassigned As Long, targetCell As Variant, msg As String
    Dim w As Variant, ws As Worksheet
    On Error GoTo Failed
    mV3Stage = "G01 locate source Sheet1 / Budget_Data"
    Set src = SourceSheet()
    If src Is Nothing Then Err.Raise vbObjectError + 180, , "No imported source data"
    mV3Stage = "G02 migrate or create Budget_Detail"
    MigrateOldPanel
    ' Make this macro safe to run directly without first running Setup.
    If FindSheet(PANEL_SHEET) Is Nothing Or FindSheet(PLAN_SHEET) Is Nothing Or FindSheet(BASE_SHEET) Is Nothing Then
        SetupBudgetOptimizerSheet1
    End If
    If FindSheet(PANEL_SHEET) Is Nothing Then Err.Raise vbObjectError + 193, , "Budget_Detail initialization failed. Run SetupBudgetOptimizerV3."
    mV3Stage = "G03 build Budget_Overview and model list"
    EnsureOverview src
    mV3Stage = "G04 obtain Budget_Overview worksheet"
    Set ov = FindSheet(OVERVIEW_SHEET)
    If ov Is Nothing Then Err.Raise vbObjectError + 194, , "Budget_Overview initialization failed."
    fy = CLng(ov.Range("B3").Value2)
    mV3Stage = "G05 sync fiscal year to Budget_Detail"
    Dim detail As Worksheet
    Set detail = FindSheet(PANEL_SHEET)
    If detail Is Nothing Then Err.Raise vbObjectError + 193, , "Budget_Detail missing after initialization"
    detail.Range("E3").Value2 = fy
    mV3Stage = "G06 locate model rows"
    lastR = ov.Cells(ov.Rows.Count, "S").End(xlUp).Row
    If lastR < MODEL_START_ROW Then Err.Raise vbObjectError + 181, , "No models in Overview"
    If Not IsNumeric(ov.Range("B6").Value2) Then Err.Raise vbObjectError + 182, , "Enter 27F TTL target in B6"
    yearTarget = CDbl(ov.Range("B6").Value2)
    If yearTarget <= 0 Or yearTarget <> Fix(yearTarget) Then Err.Raise vbObjectError + 183, , "27F TTL target must be a positive integer"
    mV3Stage = "G07 read per-model annual goals"
    ReDim targets(MODEL_START_ROW To lastR)
    ReDim hasTarget(MODEL_START_ROW To lastR)
    For r = MODEL_START_ROW To lastR
        targetCell = ov.Cells(r, 2).Value2
        If Len(Trim$(CStr(targetCell))) > 0 Then
            If Not IsNumeric(targetCell) Then Err.Raise vbObjectError + 184, , "Invalid model target at row " & r
            If CDbl(targetCell) < 0 Or CDbl(targetCell) <> Fix(CDbl(targetCell)) Then _
                Err.Raise vbObjectError + 185, , "Model target must be a nonnegative integer at row " & r
            targets(r) = CDbl(targetCell)
            hasTarget(r) = True
            fixedTotal = fixedTotal + targets(r)
        Else
            unassigned = unassigned + 1
            weightedTotal = weightedTotal + Application.Max(0, CDbl(ov.Cells(r, 3).Value2))
        End If
    Next r
    If fixedTotal > yearTarget Then Err.Raise vbObjectError + 186, , "Specified model goals exceed 27F TTL target"
    remaining = yearTarget - fixedTotal
    If unassigned = 0 And remaining > 0 Then
        If MsgBox("All model targets specified but total differs from TTL target by " & Format(remaining, "#,##0") & ". Proceed without silently allocating the difference?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    If UCase$(Trim$(CStr(ov.Range("B10").Value2))) = "ACTIVE" Then
        If MsgBox("Regenerate seasonal plan? This replaces all existing E:P monthly manual adjustments.", _
             vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    mV3Stage = "G08 generate seasonal 12-month distribution"
    w = Array(69#, 74#, 80#, 86#, 94#, 100#, 72#, 80#, 89#, 97#, 71#, 88#)
    For r = MODEL_START_ROW To lastR
        machineGoal = targets(r)
        If Not hasTarget(r) Then
            If weightedTotal > 0 Then
                weight = Application.Max(0, CDbl(ov.Cells(r, 3).Value2)) / weightedTotal
            Else
                weight = 1# / Application.Max(1, unassigned)
            End If
            machineGoal = remaining * weight
        End If
        For m = 0 To 11
            ov.Cells(r, 5 + m).Value2 = RoundTwoSignificant(machineGoal * CDbl(w(m)) / 1000#)
        Next m
    Next r
    mV3Stage = "G09 activate plan"
    ov.Range("B10").Value2 = "ACTIVE"
    mV3Stage = "G10 refresh Budget_Trial"
    RefreshBudgetSheet1Preview
    mV3Stage = "G11 activate Overview after refresh"
    ov.Activate
    msg = "Seasonal plan generated; model quantities linked to Budget_Trial." & vbCrLf & _
          "Adjust E:P directly to see live amount changes (Excel calculation must be Automatic)."
    MsgBox msg, vbInformation
    Exit Sub
Failed:
    MsgBox "Seasonal plan failed (" & Err.Number & ")" & vbCrLf & _
           "Stage: " & mV3Stage & vbCrLf & _
           "Model row: " & r & " / Month index: " & m & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function RoundTwoSignificant(ByVal x As Double) As Double
    Dim power As Double
    If x <= 0 Then Exit Function
    power = 10 ^ (Int(Log(x) / Log(10#)) - 1)
    RoundTwoSignificant = WorksheetFunction.Round(x / power, 0) * power
    ' Quantity is always whole numbers.
    If RoundTwoSignificant <> Fix(RoundTwoSignificant) Then RoundTwoSignificant = WorksheetFunction.Round(RoundTwoSignificant, 0)
End Function

' Re-apply model-level Set QTY after EVERY source-to-trial rebuild.
' Return how many FY cells were linked; fail visibly if a previously active plan
' is incomplete instead of silently returning original Set QTY.
Private Function ApplyOverviewSetQty(ByVal trial As Worksheet, ByVal src As Worksheet, _
    ByVal firstR As Long, ByVal lastR As Long, ByRef keys() As Long) As Long
    Dim ov As Worksheet, dict As Object, r As Long, rr As Long, m As Long
    Dim carry As String, fy As Long, monthNo As Long, v As Variant
    Dim planStatus As String, plannedRows As Long, filled As Long, validModels As Long
    Dim lastModel As Long, expected As Long, firstCheckRow As Long, firstCheckMonth As Long
    Dim qtyFormula As String

    Set ov = FindSheet(OVERVIEW_SHEET)
    If ov Is Nothing Then Exit Function
    lastModel = ov.Cells(ov.Rows.Count, 19).End(xlUp).Row
    If lastModel < MODEL_START_ROW Then Exit Function
    planStatus = UCase$(Trim$(CStr(ov.Range("B10").Value2)))

    ' Determine availability from actual monthly plan values, not only a label.
    For rr = MODEL_START_ROW To lastModel
        If Len(Trim$(CStr(ov.Cells(rr, 19).Value2))) > 0 Then
            validModels = validModels + 1
            For monthNo = 0 To 11
                v = ov.Cells(rr, 5 + monthNo).Value2
                If IsError(v) Then Err.Raise vbObjectError + 196, , _
                    "Set QTY plan contains an Excel error at " & ov.Cells(rr, 5 + monthNo).Address(False, False)
                If Len(Trim$(CStr(v))) > 0 Then filled = filled + 1
            Next monthNo
        End If
    Next rr

    If filled = 0 Then
        If planStatus = "ACTIVE" Then Err.Raise vbObjectError + 197, , _
            "Seasonal plan is marked ACTIVE but all monthly Set QTY cells are blank. Generate the plan again."
        ov.Range("B11").Value2 = "No 27F Set QTY plan generated"
        Exit Function
    End If
    If filled <> validModels * 12 Then
        Err.Raise vbObjectError + 198, , "Monthly Set QTY plan incomplete: " & _
            filled & " of " & validModels * 12 & " cells filled. Review Budget_Overview E:P before refreshing."
    End If

    fy = CLng(ov.Range("B3").Value2)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For rr = MODEL_START_ROW To lastModel
        carry = Trim$(CStr(ov.Cells(rr, 19).Value2))
        If Len(carry) > 0 Then
            If dict.Exists(carry) Then Err.Raise vbObjectError + 199, , _
                "Duplicate machine key in Budget_Overview S column: " & carry
            dict.Add carry, rr
        End If
    Next rr

    carry = vbNullString
    For r = firstR To lastR
        If Len(Trim$(CStr(src.Cells(r, 2).Value2))) > 0 Then carry = Trim$(CStr(src.Cells(r, 2).Value2))
        If IsBudgetDataRow(src, r) Then
            If Len(carry) = 0 Or Not dict.Exists(carry) Then _
                Err.Raise vbObjectError + 187, , "Missing model code for budget row " & r
            rr = CLng(dict(carry))
            For m = 0 To MONTH_COUNT - 1
                If FiscalYear(keys(m)) = fy Then
                    monthNo = (keys(m) Mod 100 + 8) Mod 12
                    v = ov.Cells(rr, 5 + monthNo).Value2
                    If Not IsNumeric(v) Then Err.Raise vbObjectError + 189, , _
                        "Missing numeric monthly Set QTY for model " & carry & ", month " & keys(m)
                    If CDbl(v) < 0 Or CDbl(v) <> Fix(CDbl(v)) Then _
                        Err.Raise vbObjectError + 190, , "Set QTY must be a nonnegative integer: " & carry
                    qtyFormula = "='" & OVERVIEW_SHEET & "'!" & ov.Cells(rr, 5 + monthNo).Address(True, True)
                    trial.Cells(r, FIRST_MONTH_COL + m).Formula = qtyFormula
                    ApplyOverviewSetQty = ApplyOverviewSetQty + 1
                    If firstCheckRow = 0 Then
                        firstCheckRow = r
                        firstCheckMonth = m
                    End If
                End If
            Next m
        End If
    Next r
    expected = ApplyOverviewSetQty
    If expected = 0 Then Err.Raise vbObjectError + 200, , _
        "Plan data exists but no Set QTY cells were linked; verify selected FY and rolling month headers."
    If Not trial.Cells(firstCheckRow, FIRST_MONTH_COL + firstCheckMonth).HasFormula Then _
        Err.Raise vbObjectError + 201, , "Set QTY link was not written to trial row " & firstCheckRow

    ov.Range("B10").Value2 = "ACTIVE"
    ov.Range("B11").Value2 = "Set QTY cells linked: " & expected
End Function

Private Sub BuildTrialSummaryFormulas(ByVal trial As Worksheet, ByVal panel As Worksheet, _
    ByVal firstR As Long, ByVal lastR As Long, ByRef keys() As Long, ByVal fy As Long)
    Dim m As Long, startIndex As Long, r As Long, n As Long, terms As String
    Dim arng As String, brng As String, crng As String, drng As String, erng As String, srng As String
    startIndex = -1
    For m = 0 To MONTH_COUNT - 1
        If FiscalYear(keys(m)) = fy Then
            startIndex = m
            Exit For
        End If
    Next m
    If startIndex < 0 Or startIndex + 11 >= MONTH_COUNT Then _
        Err.Raise vbObjectError + 191, , "Focus FY does not include all 12 months inside current window"
    If FiscalYear(keys(startIndex + 11)) <> fy Then _
        Err.Raise vbObjectError + 192, , "Focus FY months not contiguous in source"
    For r = firstR To lastR
        If IsBudgetDataRow(trial, r) Then
            arng = trial.Range(trial.Cells(r, FIRST_MONTH_COL + startIndex), trial.Cells(r, FIRST_MONTH_COL + startIndex + 11)).Address(False, False)
            brng = trial.Range(trial.Cells(r, 44 + startIndex), trial.Cells(r, 44 + startIndex + 11)).Address(False, False)
            crng = trial.Range(trial.Cells(r, 64 + startIndex), trial.Cells(r, 64 + startIndex + 11)).Address(False, False)
            drng = trial.Range(trial.Cells(r, 84 + startIndex), trial.Cells(r, 84 + startIndex + 11)).Address(False, False)
            erng = trial.Range(trial.Cells(r, 104 + startIndex), trial.Cells(r, 104 + startIndex + 11)).Address(False, False)
            srng = trial.Range(trial.Cells(r, SALES_AMT_COL + startIndex), trial.Cells(r, SALES_AMT_COL + startIndex + 11)).Address(False, False)
            trial.Cells(r, HELP_FIRST).Formula = "=IFERROR(SUMPRODUCT(" & arng & "," & brng & "," & crng & "," & drng & ")/100,0)"
            trial.Cells(r, HELP_FIRST + 1).Formula = "=IFERROR(SUMPRODUCT(" & arng & "," & brng & "," & crng & ")/100,0)"
            trial.Cells(r, HELP_FIRST + 2).Formula = "=IFERROR(SUMPRODUCT(" & arng & "," & brng & "," & crng & "," & erng & ")/10000,0)"
            trial.Cells(r, HELP_FIRST + 3).Formula = "=SUM(" & srng & ")"
        Else
            trial.Range(trial.Cells(r, HELP_FIRST), trial.Cells(r, HELP_FIRST + 3)).ClearContents
        End If
    Next r
    trial.Cells(firstR - 1, HELP_FIRST).Resize(1, 4).Value = Array("FY Demand Amt", "FY Demand Qty", "FY Sales Qty", "FY Sales Amt")
    trial.Columns(HELP_FIRST).Resize(, 4).Hidden = True
End Sub

Private Sub LinkLiveDetail(ByVal panel As Worksheet, ByVal trial As Worksheet, ByVal firstR As Long, ByVal lastR As Long)
    Dim helper As Variant, k As Long, col As Long, rng As String, criteria As String
    helper = Array(HELP_FIRST, HELP_FIRST + 1, HELP_FIRST + 3, HELP_FIRST + 2)
    criteria = "'" & TRIAL_SHEET & "'!$S$" & firstR & ":$S$" & lastR
    For k = 1 To 4
        col = CLng(helper(k - 1))
        rng = "'" & TRIAL_SHEET & "'!" & trial.Range(trial.Cells(firstR, col), trial.Cells(lastR, col)).Address(True, True)
        panel.Cells(k + 5, 4).Formula = "=SUMIF(" & criteria & ",$B$3," & rng & ")"
    Next k
End Sub

Private Sub LinkOverviewSummary(ByVal trial As Worksheet, ByVal firstR As Long, ByVal lastR As Long)
    Dim ov As Worksheet, src As Worksheet, m As Long, keys(0 To MONTH_COUNT - 1) As Long
    Dim headerRow As Long, n As Long, amt As Double, r As Long, lastModel As Long, orig As Double
    Set ov = FindSheet(OVERVIEW_SHEET)
    Set src = SourceSheet()
    If ov Is Nothing Then Exit Sub
    ov.Range("E6").Formula = "=SUM('" & TRIAL_SHEET & "'!" & trial.Range(trial.Cells(firstR, HELP_FIRST + 3), trial.Cells(lastR, HELP_FIRST + 3)).Address(False, False) & ")"
    FindRollingMonths src, keys, headerRow
    orig = 0
    For r = firstR To lastR
        If IsBudgetDataRow(src, r) Then
            For m = 0 To MONTH_COUNT - 1
                If FiscalYear(keys(m)) = CLng(ov.Range("B3").Value2) Then _
                    orig = orig + SafeNumber(src.Cells(r, SALES_AMT_COL + m).Value2)
            Next m
        End If
    Next r
    ov.Range("R5").Value2 = orig
    ov.Range("E8").Formula = "=E6-R5"
End Sub


' ===== V4 interaction layer =====
' The class C_BudgetV4Events routes Application.SheetChange to
' BudgetV4_OnDetailChange. Auto_Open initializes the event listener.
' Source Budget_Data / Sheet1 is never changed.

Public Sub Auto_Open()
    On Error Resume Next
    StartBudgetV4Events
    On Error GoTo 0
End Sub

Public Sub StartBudgetV4Events()
    If gBudgetV4Events Is Nothing Then
        Set gBudgetV4Events = New C_BudgetV4Events
    End If
    Set gBudgetV4Events.ExcelApp = Application
End Sub

Public Sub SetupBudgetOptimizerV4()
    Dim prevEvents As Boolean
    Dim panel As Worksheet, src As Worksheet
    On Error GoTo Failed
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    SetupBudgetOptimizerV3
    Set panel = FindSheet(PANEL_SHEET)
    Set src = SourceSheet()
    If panel Is Nothing Or src Is Nothing Then Err.Raise vbObjectError + 210, , "Missing data or Budget_Detail."
    SetupDetailEditorV4 panel
    Call EnsureSheet(V4_TARGET_SHEET)
    ' Ensure trial is available, but never replace an existing user's plan on setup.
    If FindSheet(TRIAL_SHEET) Is Nothing Then RefreshBudgetSheet1Preview True
    LoadDetailEditorV4
    StartBudgetV4Events
    Application.EnableEvents = prevEvents
    panel.Activate
    MsgBox "V4 enabled. Choose H40 in Budget_Detail!B3; edit proposed F:I in the row editor. Changes refresh automatically.", vbInformation
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "V4 setup failed (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Private Sub SetupDetailEditorV4(ByVal panel As Worksheet)
    panel.Range("A19").Value2 = "V4 | H40 EDITOR (automatically loads on H40 selection)"
    panel.Range("A20").Value2 = "Select H40 in B3. Enter YYYY_MM month range in D:E; proposed Pcs/Set, S/P, Share, Comp Ratio in F:I."
    panel.Range("A21").Value2 = "Yellow cells are editable. Blank parameter = keep original. Share/Comp Ratio only with unlock (R3 = YES)."
    panel.Range("A22:M22").Value = Array("Source row", "Model", "H40", "From YYYY_MM", "To YYYY_MM", "New Pcs/Set", "New S/P", "New Share", "New Comp Ratio", "Current Pcs/Set", "Current S/P", "Current Share", "Current Comp Ratio")
    panel.Range("A22:M22").Font.Bold = True
    panel.Range("Q3").Value2 = "Unlock Share / Comp Ratio"
    If Len(Trim$(CStr(panel.Range("R3").Value2))) = 0 Then panel.Range("R3").Value2 = "NO"
    panel.Range("R3").Validation.Delete
    panel.Range("R3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="YES,NO"
    panel.Range("Q4").Value2 = "V4 status"
    panel.Range("Q5").Value2 = "H40 editor rows"
    panel.Range("R4").Value2 = "READY"
    panel.Range("R3").Interior.Color = RGB(255, 250, 200)
    panel.Range("D" & V4_START_ROW & ":G" & V4_MAX_ROWS).Interior.Color = RGB(255, 250, 200)
    panel.Range("H" & V4_START_ROW & ":I" & V4_MAX_ROWS).Interior.Color = RGB(245, 245, 245)
    panel.Range("F" & V4_START_ROW & ":F" & V4_MAX_ROWS).NumberFormat = FMT_WHOLE
    panel.Range("G" & V4_START_ROW & ":G" & V4_MAX_ROWS).NumberFormat = FMT_PRICE
    panel.Range("H" & V4_START_ROW & ":I" & V4_MAX_ROWS).NumberFormat = FMT_WHOLE
    panel.Range("J" & V4_START_ROW & ":J" & V4_MAX_ROWS).NumberFormat = FMT_WHOLE
    panel.Range("K" & V4_START_ROW & ":K" & V4_MAX_ROWS).NumberFormat = FMT_PRICE
    panel.Range("L" & V4_START_ROW & ":M" & V4_MAX_ROWS).NumberFormat = FMT_WHOLE
    panel.Columns("A").ColumnWidth = 30
    panel.Columns("B").ColumnWidth = 24
    panel.Columns("C:M").ColumnWidth = 19
    panel.Columns("Q").ColumnWidth = 29
    panel.Columns("R").ColumnWidth = 21
    panel.Range("B4").NumberFormat = "@"
    panel.Range("B4").Font.Color = panel.Range("B4").Interior.Color
End Sub

Private Sub SaveDetailTargetsV4(ByVal h40 As String)
    Dim targetSheet As Worksheet, p As Worksheet, rr As Long, k As Long
    If Len(h40) = 0 Then Exit Sub
    Set targetSheet = EnsureSheet(V4_TARGET_SHEET)
    Set p = FindSheet(PANEL_SHEET)
    rr = FindV4TargetRow(targetSheet, h40)
    If rr = 0 Then
        rr = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1
        If rr < 2 Then rr = 2
        targetSheet.Cells(rr, 1).Value2 = h40
    End If
    For k = 1 To 4
        targetSheet.Cells(rr, k + 1).Value2 = p.Cells(k + 5, 5).Value2
    Next k
End Sub

Private Function FindV4TargetRow(ByVal ws As Worksheet, ByVal h40 As String) As Long
    Dim rr As Long
    For rr = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(rr, 1).Value2)), h40, vbTextCompare) = 0 Then
            FindV4TargetRow = rr
            Exit Function
        End If
    Next rr
End Function

Private Sub LoadDetailTargetsV4(ByVal h40 As String)
    Dim tgt As Worksheet, p As Worksheet, rr As Long, k As Long
    Set tgt = EnsureSheet(V4_TARGET_SHEET)
    Set p = FindSheet(PANEL_SHEET)
    rr = FindV4TargetRow(tgt, h40)
    For k = 1 To 4
        p.Cells(k + 5, 5).ClearContents
        If rr > 0 Then
            If Len(Trim$(CStr(tgt.Cells(rr, k + 1).Value2))) > 0 Then
                p.Cells(k + 5, 5).Value2 = tgt.Cells(rr, k + 1).Value2
            End If
        End If
    Next k
End Sub

Private Function FindV4EditPlanRow(ByVal pl As Worksheet, ByVal sourceRow As Long) As Long
    Dim rr As Long
    For rr = 2 To pl.Cells(pl.Rows.Count, "J").End(xlUp).Row
        If CStr(pl.Cells(rr, "J").Value2) = "V4" And Val(pl.Cells(rr, "B").Value2) = sourceRow Then
            FindV4EditPlanRow = rr
            Exit Function
        End If
    Next rr
End Function

Public Sub LoadDetailEditorV4()
    Dim p As Worksheet, src As Worksheet, tr As Worksheet, pl As Worksheet
    Dim h40 As String, oldH40 As String, r As Long, outR As Long
    Dim keys(0 To MONTH_COUNT - 1) As Long, header As Long, firstR As Long, lastR As Long
    Dim firstFYIndex As Long, i As Long, fy As Long, planRow As Long
    Dim currentModel As String, targetSheet As Worksheet
    On Error GoTo Failed
    Set p = FindSheet(PANEL_SHEET)
    Set src = SourceSheet()
    Set tr = FindSheet(TRIAL_SHEET)
    Set pl = FindSheet(PLAN_SHEET)
    If p Is Nothing Or src Is Nothing Or pl Is Nothing Then Err.Raise vbObjectError + 211, , "Missing source or detail sheet"
    h40 = Trim$(CStr(p.Range("B3").Value2))
    If Len(h40) = 0 Then Exit Sub
    oldH40 = Trim$(CStr(p.Range("B4").Value2))
    If oldH40 <> h40 Then
        If Len(oldH40) > 0 Then SaveDetailTargetsV4 oldH40
        LoadDetailTargetsV4 h40
    End If
    p.Range("B4").Value2 = h40
    p.Range("A" & V4_START_ROW & ":M" & V4_MAX_ROWS).ClearContents
    FindRollingMonths src, keys, header
    firstR = FindFirstBudgetDataRow(src, header)
    lastR = src.Cells(src.Rows.Count, H40_COL).End(xlUp).Row
    fy = CLng(p.Range("E3").Value2)
    firstFYIndex = -1
    For i = 0 To MONTH_COUNT - 1
        If FiscalYear(keys(i)) = fy Then
            firstFYIndex = i
            Exit For
        End If
    Next i
    If firstFYIndex < 0 Or firstFYIndex + 11 >= MONTH_COUNT Then _
        Err.Raise vbObjectError + 212, , "Focus FY not covered by full 12-month window"
    outR = V4_START_ROW
    currentModel = vbNullString
    For r = firstR To lastR
        If Len(Trim$(CStr(src.Cells(r, 13).Value2))) > 0 Then currentModel = Trim$(CStr(src.Cells(r, 13).Value2))
        If IsBudgetDataRow(src, r) And StrComp(Trim$(CStr(src.Cells(r, H40_COL).Value2)), h40, vbTextCompare) = 0 Then
            If outR > V4_MAX_ROWS Then Err.Raise vbObjectError + 213, , "Too many detail rows for selected H40"
            p.Cells(outR, 1).Value2 = r
            p.Cells(outR, 2).Value2 = currentModel
            p.Cells(outR, 3).Value2 = h40
            p.Cells(outR, 4).Value2 = Format(DateSerial(fy, 4, 1), "yyyy_mm")
            p.Cells(outR, 5).Value2 = Format(DateSerial(fy + 1, 3, 1), "yyyy_mm")
            p.Cells(outR, 10).Value2 = src.Cells(r, 44 + firstFYIndex).Value2
            p.Cells(outR, 11).Value2 = src.Cells(r, 84 + firstFYIndex).Value2
            p.Cells(outR, 12).Value2 = src.Cells(r, 104 + firstFYIndex).Value2
            p.Cells(outR, 13).Value2 = src.Cells(r, 64 + firstFYIndex).Value2
            planRow = FindV4EditPlanRow(pl, r)
            If planRow > 0 Then
                p.Cells(outR, 4).Value2 = pl.Cells(planRow, 3).Value2
                p.Cells(outR, 5).Value2 = pl.Cells(planRow, 4).Value2
                p.Range(p.Cells(outR, 6), p.Cells(outR, 9)).Value = pl.Range(pl.Cells(planRow, 5), pl.Cells(planRow, 8)).Value
            End If
            outR = outR + 1
        End If
    Next r
    p.Range("R5").Value2 = outR - V4_START_ROW
    Update26FForSelectedH40V4
    p.Range("R4").Value2 = "H40 ready: " & h40
    Exit Sub
Failed:
    Err.Raise vbObjectError + 214, "LoadDetailEditorV4", "Load H40 failed (" & Err.Number & "): " & Err.Description
End Sub

Private Sub Update26FForSelectedH40V4()
    Dim src As Worksheet, tr As Worksheet, p As Worksheet, b As Worksheet
    Dim keys(0 To MONTH_COUNT - 1) As Long, header As Long, firstR As Long, lastR As Long
    Dim h40 As String, r As Long, m As Long, k As Long, br As Long
    Dim v(1 To 4) As Double, old26(1 To 4) As Double, cur26(1 To 4) As Double
    Dim raw As Variant
    Set src = SourceSheet()
    Set tr = FindSheet(TRIAL_SHEET)
    Set p = FindSheet(PANEL_SHEET)
    Set b = FindSheet(BASE_SHEET)
    If tr Is Nothing Then Exit Sub
    h40 = Trim$(CStr(p.Range("B3").Value2))
    FindRollingMonths src, keys, header
    firstR = FindFirstBudgetDataRow(src, header)
    lastR = src.Cells(src.Rows.Count, H40_COL).End(xlUp).Row
    For r = firstR To lastR
        If IsBudgetDataRow(src, r) And StrComp(Trim$(CStr(src.Cells(r, H40_COL).Value2)), h40, vbTextCompare) = 0 Then
            For m = 0 To MONTH_COUNT - 1
                If FiscalYear(keys(m)) = 2026 Then
                    ReadMetrics src, r, m, v
                    AddMetrics old26, v
                    ReadMetrics tr, r, m, v
                    AddMetrics cur26, v
                End If
            Next m
        End If
    Next r
    br = FindBaselineRow(b, h40)
    For k = 1 To 4
        p.Cells(k + 5, 3).Value2 = cur26(k)
        p.Cells(k + 5, 2).ClearContents
        If br > 0 Then
            raw = b.Cells(br, k + 1).Value2
            If Not IsError(raw) And Len(Trim$(CStr(raw))) > 0 Then
                If IsNumeric(raw) Then p.Cells(k + 5, 2).Value2 = CDbl(raw) + cur26(k) - old26(k)
            End If
        End If
    Next k
End Sub

Private Sub SyncDetailRowToPlanV4(ByVal p As Worksheet, ByVal detailRow As Long)
    Dim pl As Worksheet, planRow As Long, sourceRow As Long, i As Long
    Dim monthFrom As String, monthTo As String, raw As Variant, h40 As String
    Dim hasChanges As Boolean, qty As Double, scale As Double
    Set pl = FindSheet(PLAN_SHEET)
    sourceRow = CLng(Val(p.Cells(detailRow, 1).Value2))
    If sourceRow < 1 Then Exit Sub
    h40 = Trim$(CStr(p.Cells(detailRow, 3).Value2))
    monthFrom = Trim$(CStr(p.Cells(detailRow, 4).Value2))
    monthTo = Trim$(CStr(p.Cells(detailRow, 5).Value2))
    If ParseMonth(monthFrom) = 0 Or ParseMonth(monthTo) = 0 Or ParseMonth(monthFrom) > ParseMonth(monthTo) Then _
        Err.Raise vbObjectError + 215, , "Invalid YYYY_MM range at Budget_Detail row " & detailRow
    For i = 6 To 9
        raw = p.Cells(detailRow, i).Value2
        If IsError(raw) Then Err.Raise vbObjectError + 216, , "Invalid input at row " & detailRow
        If Len(Trim$(CStr(raw))) > 0 Then
            If Not IsNumeric(raw) Then Err.Raise vbObjectError + 217, , "Numeric value required at row " & detailRow
            qty = CDbl(raw)
            If qty < 0 Then Err.Raise vbObjectError + 218, , "Negative value not allowed"
            Select Case i
                Case 6
                    If qty <> Fix(qty) Then Err.Raise vbObjectError + 219, , "Pcs/Set must be an integer"
                Case 7
                    If Abs(qty * 100000# - Round(qty * 100000#, 0)) > 0.000001 Then _
                        Err.Raise vbObjectError + 220, , "S/P supports at most 5 decimals"
                Case 8, 9
                    If qty > 100 Or qty <> Fix(qty) Then Err.Raise vbObjectError + 221, , "Share/Comp Ratio must be integers 0-100"
            End Select
            hasChanges = True
        End If
    Next i
    If (Len(Trim$(CStr(p.Cells(detailRow, 8).Value2))) > 0 Or _
        Len(Trim$(CStr(p.Cells(detailRow, 9).Value2))) > 0) And _
        UCase$(Trim$(CStr(p.Range("R3").Value2))) <> "YES" Then _
        Err.Raise vbObjectError + 222, , "Unlock Share/Comp Ratio in R3 before editing H or I"
    planRow = FindV4EditPlanRow(pl, sourceRow)
    If planRow = 0 And hasChanges Then
        planRow = pl.Cells(pl.Rows.Count, "B").End(xlUp).Row + 1
        If planRow < 2 Then planRow = 2
    End If
    If planRow = 0 Then Exit Sub
    If hasChanges Then
        pl.Cells(planRow, 1).Value2 = "Y"
        pl.Cells(planRow, 2).Value2 = sourceRow
        pl.Cells(planRow, 3).Value2 = monthFrom
        pl.Cells(planRow, 4).Value2 = monthTo
        pl.Range(pl.Cells(planRow, 5), pl.Cells(planRow, 8)).Value = p.Range(p.Cells(detailRow, 6), p.Cells(detailRow, 9)).Value
        pl.Cells(planRow, 9).Value2 = h40
        pl.Cells(planRow, 10).Value2 = "V4"
    Else
        pl.Cells(planRow, 1).ClearContents
        pl.Range(pl.Cells(planRow, 5), pl.Cells(planRow, 8)).ClearContents
    End If
End Sub

' Called by application event handler after a REAL user edit.
' If editing fails, existing trial remains until a successful refresh.
Public Sub BudgetV4_OnDetailChange(ByVal target As Range)
    Dim p As Worksheet, previousEvents As Boolean, changed As Range, cell As Range
    Dim seen As Object, rr As Long, key As Variant, oldH40 As String
    If gV4Busy Then Exit Sub
    Set p = FindSheet(PANEL_SHEET)
    If p Is Nothing Then Exit Sub
    If target.CountLarge > 500 Then Exit Sub
    On Error GoTo Failed
    gV4Busy = True
    previousEvents = Application.EnableEvents
    Application.EnableEvents = False

    If Not Intersect(target, p.Range("B3")) Is Nothing Then
        LoadDetailEditorV4
    ElseIf Not Intersect(target, p.Range("E6:E9")) Is Nothing Then
        SaveDetailTargetsV4 Trim$(CStr(p.Range("B3").Value2))
    ElseIf Not Intersect(target, p.Range("R3")) Is Nothing Then
        If UCase$(Trim$(CStr(p.Range("R3").Value2))) <> "YES" Then p.Range("R3").Value2 = "NO"
    Else
        Set changed = Intersect(target, p.Range("D" & V4_START_ROW & ":I" & V4_MAX_ROWS))
        If Not changed Is Nothing Then
            Set seen = CreateObject("Scripting.Dictionary")
            For Each cell In changed.Cells
                rr = cell.Row
                If Len(Trim$(CStr(p.Cells(rr, 1).Value2))) > 0 Then seen(CStr(rr)) = rr
            Next cell
            For Each key In seen.Keys
                SyncDetailRowToPlanV4 p, CLng(seen(key))
            Next key
            If seen.Count > 0 Then
                p.Range("R4").Value2 = "Recalculating..."
                RefreshBudgetSheet1Preview True
                Update26FForSelectedH40V4
                p.Range("R4").Value2 = "Updated automatically"
            End If
        End If
    End If
Done:
    Application.EnableEvents = previousEvents
    gV4Busy = False
    Exit Sub
Failed:
    Dim message As String
    message = "V4 auto-update failed (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    p.Range("R4").Value2 = "ERROR - check inputs"
    Application.EnableEvents = True
    gV4Busy = False
    MsgBox message, vbExclamation
End Sub
