Attribute VB_Name = "BudgetOptimizer_V5_Target4Sheets"
Option Explicit

' One-time setup + repeatable import from an external workbook containing
' FOUR 2027F target worksheets. Each target sheet: H40 = B4:B..., 2027F = F4:F...
' The user configures actual sheet names once in H40_Target_Source!B2:B5.
' Reuses the existing V4 H40_Targets table (A:H40, B:E four targets).
' Does NOT touch Budget_Data, Budget_Trial, Edit_Plan, or monthly Set QTY.

Private Const T4_CONFIG As String = "H40_Target_Source"
Private Const T4_STORE As String = "H40_Targets"
Private Const T4_DETAIL As String = "Budget_Detail"
Private Const T4_LIST As String = "H40_List"
Private Const T4_DATA_FIRST As Long = 4

Public Sub SetupBudgetV5FourSheetTargets()
    Dim cfg As Worksheet, detail As Worksheet, already As Boolean
    Dim prevEvents As Boolean, problem As String
    prevEvents = Application.EnableEvents
    On Error GoTo Failed
    Set detail = ThisWorkbook.Worksheets(T4_DETAIL)
    On Error Resume Next
    Set cfg = ThisWorkbook.Worksheets(T4_CONFIG)
    On Error GoTo Failed
    Application.EnableEvents = False
    If cfg Is Nothing Then
        Set cfg = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        cfg.Name = T4_CONFIG
    End If
    With cfg
        .Range("A1:C1").Value = Array("Metric / Detail row", "SOURCE SHEET NAME", "Source columns")
        .Range("A2").Value2 = "Demand Amt / E6"
        .Range("A3").Value2 = "Demand Qty / E7"
        .Range("A4").Value2 = "Sales Amt / E8"
        .Range("A5").Value2 = "Sales Qty / E9"
        .Range("C2:C5").Value2 = "H40 = B, 2027F = F, row >= 4"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(44, 74, 108)
        .Range("A1:C1").Font.Color = RGB(255, 255, 255)
        .Range("B2:B5").Interior.Color = RGB(255, 244, 193)
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 34
        .Columns("C").ColumnWidth = 32
        .Range("A7").Value2 = "1. Type the four EXACT source sheet names in B2:B5."
        .Range("A8").Value2 = "2. Run ImportBudgetV5FourSheetTargets. Select the updated downloaded Excel file."
        .Range("A9").Value2 = "3. Unknown source H40s are skipped; no existing other targets are erased."
        .Range("A10").Value2 = "4. Change the source names here only when its four sheet names change."
    End With

    ' No V4 editor cells are moved. This is an independent calculated column.
    With detail
        .Range("G5").Value2 = "Current / Target"
        .Range("G5").Interior.Color = RGB(47, 73, 105)
        .Range("G5").Font.Color = RGB(255, 255, 255)
        .Range("G5").Font.Bold = True
        .Range("G6:G9").NumberFormat = "0%"
        .Columns("G").ColumnWidth = 18
        .Range("G6").Formula = "=IF(OR(E6="""",E6=0),"""",D6/E6)"
        .Range("G7").Formula = "=IF(OR(E7="""",E7=0),"""",D7/E7)"
        .Range("G8").Formula = "=IF(OR(E8="""",E8=0),"""",D8/E8)"
        .Range("G9").Formula = "=IF(OR(E9="""",E9=0),"""",D9/E9)"
        .Range("G6:G9").Interior.Color = RGB(231, 239, 249)
    End With
    Application.EnableEvents = prevEvents
    MsgBox "Target source mapping ready. Set four sheet names in H40_Target_Source!B2:B5.", vbInformation
    Exit Sub
Failed:
    problem = "Four-sheet target setup failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = prevEvents
    MsgBox problem, vbExclamation
End Sub

Public Sub ImportBudgetV5FourSheetTargets()
    Dim cfg As Worksheet, store As Worksheet, detail As Worksheet, listSheet As Worksheet
    Dim srcBook As Workbook, srcSheet As Worksheet
    Dim known As Object, existing As Object, allRows As Object, sheetSeen As Object
    Dim inputFile As Variant, sheetName As String, sourceKey As String, key As String
    Dim valuesByMetric(1 To 4) As Object
    Dim rr As Long, k As Long, j As Long, lastRow As Long, destRow As Long
    Dim totalRead As Long, unknowns As Long, matches As Long, skippedBlanks As Long
    Dim rowCount As Long, filled(1 To 4) As Long, arr As Variant, entry As Variant
    Dim raw As Variant, header As String, currentKey As String, previousKey As String
    Dim copied As Long, prevEvents As Boolean, problem As String, newKey As Variant
    Dim alreadyOpen As Boolean, wb As Workbook, seenName As Object

    prevEvents = Application.EnableEvents
    On Error GoTo Failed
    Set cfg = ThisWorkbook.Worksheets(T4_CONFIG)
    Set detail = ThisWorkbook.Worksheets(T4_DETAIL)
    Set listSheet = ThisWorkbook.Worksheets(T4_LIST)
    Set store = ThisWorkbook.Worksheets(T4_STORE)

    currentKey = T4Key(detail.Range("B3").Value2)
    previousKey = T4Key(detail.Range("B4").Value2)
    If currentKey <> previousKey Then _
        Err.Raise vbObjectError + 781, , "Current H40 B3/B4 not synchronized. Select it again first."

    Set seenName = CreateObject("Scripting.Dictionary")
    seenName.CompareMode = vbTextCompare
    For k = 1 To 4
        sheetName = Trim$(CStr(cfg.Cells(k + 1, 2).Value2))
        If Len(sheetName) = 0 Then _
            Err.Raise vbObjectError + 782, , "Enter the source sheet name in H40_Target_Source!B" & k + 1
        If seenName.Exists(sheetName) Then _
            Err.Raise vbObjectError + 783, , "One sheet cannot be mapped to two different metrics: " & sheetName
        seenName.Add sheetName, True
    Next k

    inputFile = Application.GetOpenFilename("Excel Workbooks (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
                                            , "Choose downloaded 4-sheet H40 targets workbook")
    If VarType(inputFile) = vbBoolean Then Exit Sub
    If CStr(inputFile) = ThisWorkbook.FullName Then _
        Err.Raise vbObjectError + 784, , "Choose external downloaded budget workbook, not the optimizer."

    ' Reuse already open workbook if present. Never close another workbook
    ' the user had open before the macro.
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(inputFile), vbTextCompare) = 0 Then
            Set srcBook = wb
            alreadyOpen = True
            Exit For
        End If
    Next wb
    If srcBook Is Nothing Then
        Set srcBook = Application.Workbooks.Open(Filename:=CStr(inputFile), _
                                             ReadOnly:=True, UpdateLinks:=0)
    End If

    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare
    For rr = 2 To listSheet.Cells(listSheet.Rows.Count, "A").End(xlUp).Row
        key = T4Key(listSheet.Cells(rr, 1).Value2)
        If Len(key) > 0 Then known(key) = True
    Next rr
    If known.Count = 0 Then Err.Raise vbObjectError + 785, , "No H40 identifiers found in H40_List."

    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    For rr = 2 To store.Cells(store.Rows.Count, "A").End(xlUp).Row
        key = T4Key(store.Cells(rr, 1).Value2)
        If Len(key) > 0 Then
            If existing.Exists(key) Then _
                Err.Raise vbObjectError + 786, , "Duplicate stored H40: " & key
            existing.Add key, rr
        End If
    Next rr

    Set allRows = CreateObject("Scripting.Dictionary")
    allRows.CompareMode = vbTextCompare
    For k = 1 To 4
        Set valuesByMetric(k) = CreateObject("Scripting.Dictionary")
        valuesByMetric(k).CompareMode = vbTextCompare
        sheetName = Trim$(CStr(cfg.Cells(k + 1, 2).Value2))
        Set srcSheet = Nothing
        On Error Resume Next
        Set srcSheet = srcBook.Worksheets(sheetName)
        On Error GoTo Failed
        If srcSheet Is Nothing Then _
            Err.Raise vbObjectError + 787, , "Sheet not found in downloaded file: " & sheetName

        ' Validate that F is the intended forecast year, not the historical FY.
        header = Trim$(CStr(srcSheet.Range("F2").Value2))
        If InStr(1, header, "2027", vbTextCompare) = 0 Then _
            Err.Raise vbObjectError + 788, , _
                "Expected 2027F in " & sheetName & "!F2, found: " & header

        lastRow = srcSheet.Cells(srcSheet.Rows.Count, "B").End(xlUp).Row
        If lastRow < T4_DATA_FIRST Then _
            Err.Raise vbObjectError + 789, , "No H40 data in source sheet: " & sheetName
        For rr = T4_DATA_FIRST To lastRow
            key = T4Key(srcSheet.Cells(rr, 2).Value2)
            raw = srcSheet.Cells(rr, 6).Value2
            If Len(key) > 0 Then
                ' Only H40s that exist in the current budget are imported.
                ' Source file can contain other customers/H40s; skip them.
                If known.Exists(key) Then
                    If valuesByMetric(k).Exists(key) Then _
                        Err.Raise vbObjectError + 790, , _
                            "Duplicate H40 " & key & " in " & sheetName & " row " & rr
                    If IsError(raw) Then _
                        Err.Raise vbObjectError + 791, , _
                            "Invalid 2027F error in " & sheetName & "!F" & rr
                    If Len(Trim$(CStr(raw))) > 0 Then
                        If Not IsNumeric(raw) Then _
                            Err.Raise vbObjectError + 792, , _
                                "Non-numeric 2027F value at " & sheetName & "!F" & rr
                        If CDbl(raw) < 0 Then _
                            Err.Raise vbObjectError + 793, , _
                                "Negative 2027F target at " & sheetName & "!F" & rr
                        valuesByMetric(k).Add key, CDbl(raw)
                        allRows(key) = True
                        filled(k) = filled(k) + 1
                    Else
                        valuesByMetric(k).Add key, Empty
                        skippedBlanks = skippedBlanks + 1
                    End If
                Else
                    unknowns = unknowns + 1
                End If
            End If
        Next rr
    Next k

    If allRows.Count = 0 Then _
        Err.Raise vbObjectError + 794, , "No matching H40 2027F values found in the four sheets."

    If MsgBox("Four 2027F target sheets validated." & vbCrLf & _
        "Unique matched H40s: " & allRows.Count & vbCrLf & _
        "Demand Amt / Qty: " & filled(1) & " / " & filled(2) & vbCrLf & _
        "Sales Amt / Qty: " & filled(3) & " / " & filled(4) & vbCrLf & _
        "Source H40 rows not in current budget (across sheets): " & unknowns & vbCrLf & _
        "Empty F cells preserved (not overwritten): " & skippedBlanks & vbCrLf & _
        "Continue updating H40_Targets?", _
        vbQuestion + vbYesNo, "Import 2027F H40 Targets") <> vbYes Then GoTo CleanExit

    Application.EnableEvents = False
    ' Revalidate all data before making changes. Write only numeric F values.
    For Each newKey In allRows.Keys
        key = CStr(newKey)
        If existing.Exists(key) Then
            destRow = CLng(existing(key))
        Else
            destRow = store.Cells(store.Rows.Count, "A").End(xlUp).Row + 1
            If destRow < 2 Then destRow = 2
            store.Cells(destRow, "A").NumberFormat = "@"
            store.Cells(destRow, "A").Value2 = key
            existing.Add key, destRow
        End If
        For k = 1 To 4
            If valuesByMetric(k).Exists(key) Then
                raw = valuesByMetric(k)(key)
                If Not IsEmpty(raw) Then
                    store.Cells(destRow, k + 1).Value2 = CDbl(raw)
                    copied = copied + 1
                End If
            End If
        Next k
    Next newKey

    ' V4 saves visible Detail E6:E9 on the next B3 selection. Prevent its
    ' current on-screen values overwriting a newly imported current H40.
    If existing.Exists(currentKey) Then
        destRow = CLng(existing(currentKey))
        For k = 1 To 4
            If valuesByMetric(k).Exists(currentKey) Then
                raw = valuesByMetric(k)(currentKey)
                If Not IsEmpty(raw) Then detail.Cells(k + 5, 5).Value2 = CDbl(raw)
            End If
        Next k
    End If
    detail.Range("G6:G9").Calculate

CleanExit:
    Application.EnableEvents = prevEvents
    If Not srcBook Is Nothing And Not alreadyOpen Then srcBook.Close SaveChanges:=False
    If copied > 0 Then
        MsgBox "Imported " & copied & " H40 target values." & vbCrLf & _
            "Select H40 in Budget_Detail!B3 to see the saved E6:E9 targets.", vbInformation
    End If
    Exit Sub

Failed:
    problem = "4-sheet H40 target import failed (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    Application.EnableEvents = prevEvents
    If Not srcBook Is Nothing And Not alreadyOpen Then srcBook.Close SaveChanges:=False
    MsgBox problem & vbCrLf & "No budget source or trial input was intentionally changed.", vbExclamation
End Sub

Private Function T4Key(ByVal v As Variant) As String
    If IsError(v) Then Err.Raise vbObjectError + 795, , "Invalid H40 Excel error"
    T4Key = Trim$(Replace(CStr(v), ChrW$(160), " "))
End Function
