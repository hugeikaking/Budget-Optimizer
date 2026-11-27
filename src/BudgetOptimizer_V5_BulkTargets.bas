Attribute VB_Name = "BudgetOptimizer_V5_BulkTargets"
Option Explicit

' Optional V5 enhancement: bulk import four H40 targets and Detail G6:G9
' current / target percentage. Leaves V4 logic and other modules unchanged.
' Paste data into H40_Target_Import!A:E (header row 1), then run import.
' Existing H40_Targets is the V4 authoritative target store.

Private Const TB_IMPORT As String = "H40_Target_Import"
Private Const TB_STORE As String = "H40_Targets"
Private Const TB_DETAIL As String = "Budget_Detail"
Private Const TB_SOURCE As String = "Budget_Data"
Private Const TB_FALLBACK As String = "Sheet1"

Public Sub SetupBudgetV5BulkTargets()
    Dim ws As Worksheet, p As Worksheet
    On Error GoTo Failed
    Set p = ThisWorkbook.Worksheets(TB_DETAIL)
    Set ws = BulkEnsureSheet(TB_IMPORT)
    ws.Range("A1:E1").Value = Array("H40", "Demand Amt Target", _
        "Demand Qty Target", "Sales Amt Target", "Sales Qty Target")
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(44, 71, 101)
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A").NumberFormat = "@"
    ws.Columns("A").ColumnWidth = 19
    ws.Columns("B:E").ColumnWidth = 23
    ws.Range("B2:E10000").NumberFormat = "#,##0.#####;[Red]-#,##0.#####;0"

    p.Range("G5").Value2 = "Current / Target"
    p.Range("G5").Interior.Color = RGB(47, 73, 105)
    p.Range("G5").Font.Color = RGB(255, 255, 255)
    p.Range("G5").Font.Bold = True
    p.Columns("G").ColumnWidth = 21
    p.Range("G6:G9").NumberFormat = "0%"
    p.Range("G6:G9").Interior.Color = RGB(228, 239, 250)
    BulkInstallRatioFormulas p
    MsgBox "Template ready in H40_Target_Import (A:E)." & vbCrLf & _
           "Paste targets from row 2, then run ImportBudgetV5H40Targets.", vbInformation
    Exit Sub
Failed:
    MsgBox "Target template setup failed (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Public Sub ImportBudgetV5H40Targets()
    Dim stage As String, staging As Worksheet, storage As Worksheet
    Dim src As Worksheet, p As Worksheet, known As Object, rows As Object
    Dim pending As Object, rowIndex As Long, lastImport As Long, lastStore As Long
    Dim lastSource As Long, i As Long, j As Long, key As String
    Dim values As Variant, raw As Variant, rowData As Variant
    Dim existing As Long, created As Long, overwritten As Long
    Dim oldEvents As Boolean, msg As String, destRow As Long, sourceKey As String
    Dim mapRows As Object, targetRow As Long, targetValue As Variant

    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    stage = "Locating sheets"
    Set staging = ThisWorkbook.Worksheets(TB_IMPORT)
    Set storage = ThisWorkbook.Worksheets(TB_STORE)
    Set p = ThisWorkbook.Worksheets(TB_DETAIL)
    On Error Resume Next
    Set src = ThisWorkbook.Worksheets(TB_SOURCE)
    If src Is Nothing Then Set src = ThisWorkbook.Worksheets(TB_FALLBACK)
    On Error GoTo Failed
    If src Is Nothing Then Err.Raise vbObjectError + 751, , "Budget_Data/Sheet1 not found."
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare
    Set pending = CreateObject("Scripting.Dictionary")
    pending.CompareMode = vbTextCompare
    Set mapRows = CreateObject("Scripting.Dictionary")
    mapRows.CompareMode = vbTextCompare

    stage = "Reading known H40 codes"
    lastSource = src.Cells(src.Rows.Count, 19).End(xlUp).Row
    For i = 2 To lastSource
        If IsError(src.Cells(i, 19).Value2) Then Err.Raise vbObjectError + 752, , _
            "Source H40 error at row " & i
        sourceKey = Trim$(CStr(src.Cells(i, 19).Value2))
        If Len(sourceKey) > 0 And LCase$(sourceKey) <> "h40" And _
            LCase$(sourceKey) <> "eop" Then known(sourceKey) = True
    Next i

    stage = "Validating import data"
    lastImport = staging.Cells(staging.Rows.Count, "A").End(xlUp).Row
    If lastImport < 2 Then Err.Raise vbObjectError + 753, , _
        "Paste at least one H40 into H40_Target_Import!A2:E."
    For i = 2 To lastImport
        raw = staging.Cells(i, 1).Value2
        If IsError(raw) Then Err.Raise vbObjectError + 754, , "Invalid H40 at row " & i
        key = Trim$(CStr(raw))
        If Len(key) = 0 Then Err.Raise vbObjectError + 755, , "Blank H40 at row " & i
        If Not known.Exists(key) Then Err.Raise vbObjectError + 756, , _
            "H40 not present in source at import row " & i & ": " & key
        If pending.Exists(key) Then Err.Raise vbObjectError + 757, , _
            "Duplicate H40 in import at row " & i & ": " & key
        ReDim values(1 To 4)
        For j = 1 To 4
            raw = staging.Cells(i, j + 1).Value2
            If IsError(raw) Then Err.Raise vbObjectError + 758, , _
                "Excel error in target at import row " & i & ", column " & j + 1
            If Len(Trim$(CStr(raw))) > 0 Then
                If Not IsNumeric(raw) Then Err.Raise vbObjectError + 759, , _
                    "Target must be numeric at import row " & i & ", column " & j + 1
                If CDbl(raw) < 0 Then Err.Raise vbObjectError + 760, , _
                    "Negative target at import row " & i & ", column " & j + 1
                values(j) = CDbl(raw)
            Else
                values(j) = Empty
            End If
        Next j
        pending.Add key, values
    Next i

    stage = "Validating existing saved target rows"
    lastStore = storage.Cells(storage.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastStore
        raw = storage.Cells(i, 1).Value2
        If IsError(raw) Then Err.Raise vbObjectError + 761, , "Stored H40 error at row " & i
        key = Trim$(CStr(raw))
        If Len(key) > 0 Then
            If mapRows.Exists(key) Then Err.Raise vbObjectError + 762, , _
                "Duplicate H40 already stored: " & key & ". Resolve before import."
            mapRows.Add key, i
        End If
    Next i

    If MsgBox("Validated " & pending.Count & " unique H40 target rows." & vbCrLf & _
        "Existing H40 targets will be updated only for imported codes." & vbCrLf & _
        "Blank import target fields will CLEAR that target." & vbCrLf & _
        "Unlisted H40 targets and all budget input edits remain unchanged." & vbCrLf & _
        "Apply now?", vbQuestion + vbYesNo, "Import H40 Targets") <> vbYes Then Exit Sub

    Application.EnableEvents = False
    stage = "Writing validated targets"
    For i = 2 To lastImport
        key = Trim$(CStr(staging.Cells(i, 1).Value2))
        rowData = pending(key)
        If mapRows.Exists(key) Then
            destRow = CLng(mapRows(key))
            overwritten = overwritten + 1
        Else
            lastStore = lastStore + 1
            If lastStore < 2 Then lastStore = 2
            destRow = lastStore
            mapRows.Add key, destRow
            created = created + 1
        End If
        storage.Cells(destRow, 1).NumberFormat = "@"
        storage.Cells(destRow, 1).Value2 = key
        For j = 1 To 4
            If IsEmpty(rowData(j)) Then
                storage.Cells(destRow, j + 1).ClearContents
            Else
                storage.Cells(destRow, j + 1).Value2 = rowData(j)
            End If
        Next j
    Next i

    ' The V4 editor automatically reads H40_Targets on B3 changes.
    ' To update the current displayed H40 without changing B3, clear the
    ' hidden "previous H40" tracker B4 temporarily. This prevents V4
    ' from overwriting freshly imported targets with the old E6:E9 values.
    stage = "Refreshing selected H40 without overwriting import"
    p.Range("B4").ClearContents
    LoadDetailEditorV4
    BulkInstallRatioFormulas p
    Application.EnableEvents = oldEvents
    MsgBox "Imported " & pending.Count & " H40 target rows." & vbCrLf & _
        "New: " & created & "  Updated: " & overwritten & vbCrLf & _
        "Current H40 targets E6:E9 and ratios G6:G9 refreshed.", vbInformation
    Exit Sub

Failed:
    msg = "H40 target import failed at " & stage & " (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    Application.EnableEvents = oldEvents
    MsgBox msg, vbExclamation
End Sub

Private Sub BulkInstallRatioFormulas(ByVal p As Worksheet)
    Dim rr As Long
    For rr = 6 To 9
        p.Cells(rr, 7).Formula = "=IFERROR(IF(OR(E" & rr & _
            "="""",E" & rr & "=0),"""",D" & rr & "/E" & rr & "),"""")"
    Next rr
    p.Range("G6:G9").NumberFormat = "0%"
End Sub

Private Function BulkEnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set BulkEnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If BulkEnsureSheet Is Nothing Then
        Set BulkEnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        BulkEnsureSheet.Name = name
    End If
End Function
