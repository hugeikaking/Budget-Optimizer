Attribute VB_Name = "BudgetOptimizer_V5_TargetImport"
Option Explicit

' Optional add-on for a V5 workbook already configured with V4 and V5.
' Staging H40_Target_Import!A:E -> existing H40_Targets!A:E.
' Preserves other targets and every existing Edit_Plan/Budget_Trial value.
' No changes to ThisWorkbook event handlers or V4 module are required.

Private Const TI_STAGE As String = "H40_Target_Import"
Private Const TI_STORE As String = "H40_Targets"
Private Const TI_DETAIL As String = "Budget_Detail"
Private Const TI_H40_LIST As String = "H40_List"

Public Sub SetupBudgetV5TargetsImport()
    Dim ws As Worksheet, detail As Worksheet, oldEvents As Boolean
    Dim problem As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Set detail = ThisWorkbook.Worksheets(TI_DETAIL)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TI_STAGE)
    On Error GoTo Failed
    Application.EnableEvents = False
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = TI_STAGE
    End If

    ' Never clear existing paste area on repeat setup.
    With ws
        .Range("A1:E1").Value = Array("H40", "Demand Amt Target", "Demand Qty Target", _
                                     "Sales Amt Target", "Sales Qty Target")
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(44, 74, 108)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        .Columns("A").NumberFormat = "@"
        .Columns("A").ColumnWidth = 20
        .Columns("B:E").ColumnWidth = 23
        .Range("B2:E10000").NumberFormat = "#,##0.#####;[Red]-#,##0.#####;0"
        .Range("G1").Value2 = "PASTE TARGETS"
        .Range("G2").Value2 = "Paste H40 in A, four targets in B:E, starting row 2."
        .Range("G3").Value2 = "Blank target cell = preserve existing value."
        .Range("G4").Value2 = "Zero = a valid target, shown as blank % denominator."
        .Range("G5").Value2 = "Run ImportBudgetV5Targets to validate and apply."
        .Range("G6").Value2 = "Existing H40s not included are unchanged."
        .Columns("G").ColumnWidth = 62
    End With

    ' G is outside the existing V4/V5 financial-input area A:F.
    With detail
        .Range("G5").Value2 = "Current / Target"
        .Range("G5").Interior.Color = RGB(47, 73, 105)
        .Range("G5").Font.Color = RGB(255, 255, 255)
        .Range("G5").Font.Bold = True
        .Columns("G").ColumnWidth = 18
        .Range("G6:G9").NumberFormat = "0%"
        .Range("G6").Formula = "=IF(OR(E6="""",E6=0),"""",D6/E6)"
        .Range("G7").Formula = "=IF(OR(E7="""",E7=0),"""",D7/E7)"
        .Range("G8").Formula = "=IF(OR(E8="""",E8=0),"""",D8/E8)"
        .Range("G9").Formula = "=IF(OR(E9="""",E9=0),"""",D9/E9)"
        .Range("G6:G9").Interior.Color = RGB(231, 239, 249)
        .Range("G6:G9").Font.Bold = True
    End With

    Application.EnableEvents = oldEvents
    MsgBox "H40 target paste sheet is ready. Paste A:E, then run ImportBudgetV5Targets.", vbInformation
    Exit Sub
Failed:
    problem = "Target import setup failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = oldEvents
    MsgBox problem, vbExclamation
End Sub

Public Sub ImportBudgetV5Targets()
    Dim stageSheet As Worksheet, storeSheet As Worksheet, detail As Worksheet
    Dim listSheet As Worksheet, known As Object, staged As Object, existing As Object
    Dim lastRow As Long, r As Long, k As Long, matched As Long, updated As Long
    Dim key As String, raw As Variant, record As Variant, h As Variant
    Dim rowTarget As Long, currentKey As String, previousKey As String
    Dim hasValue As Boolean, v(1 To 4) As Variant, failure As String
    Dim oldEvents As Boolean, applyCount As Long

    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Set stageSheet = ThisWorkbook.Worksheets(TI_STAGE)
    Set detail = ThisWorkbook.Worksheets(TI_DETAIL)
    Set listSheet = ThisWorkbook.Worksheets(TI_H40_LIST)
    On Error Resume Next
    Set storeSheet = ThisWorkbook.Worksheets(TI_STORE)
    On Error GoTo Failed
    If storeSheet Is Nothing Then _
        Err.Raise vbObjectError + 751, , "H40_Targets missing. Set up V4 before importing."

    currentKey = TargetKey(detail.Range("B3").Value2)
    previousKey = TargetKey(detail.Range("B4").Value2)
    If currentKey <> previousKey Then _
        Err.Raise vbObjectError + 752, , _
            "Current H40 selection is not synchronized. Select H40 again before import."

    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare
    For r = 2 To listSheet.Cells(listSheet.Rows.Count, 1).End(xlUp).Row
        key = TargetKey(listSheet.Cells(r, 1).Value2)
        If Len(key) > 0 Then known(key) = True
    Next r

    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    For r = 2 To storeSheet.Cells(storeSheet.Rows.Count, 1).End(xlUp).Row
        key = TargetKey(storeSheet.Cells(r, 1).Value2)
        If Len(key) > 0 Then
            If existing.Exists(key) Then _
                Err.Raise vbObjectError + 753, , _
                    "Duplicate H40 in H40_Targets rows " & existing(key) & " and " & r
            existing.Add key, r
        End If
    Next r

    Set staged = CreateObject("Scripting.Dictionary")
    staged.CompareMode = vbTextCompare
    lastRow = stageSheet.Cells(stageSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then _
        Err.Raise vbObjectError + 754, , "Paste at least one H40 target row into A2:E."
    ' Validate the complete staging table before ANY target values are changed.
    For r = 2 To lastRow
        key = TargetKey(stageSheet.Cells(r, 1).Value2)
        If Len(key) > 0 Then
            If Not known.Exists(key) Then _
                Err.Raise vbObjectError + 755, , _
                    "Unknown H40 at staging row " & r & ": " & key
            If staged.Exists(key) Then _
                Err.Raise vbObjectError + 756, , _
                    "Duplicate H40 in pasted data: " & key
            hasValue = False
            For k = 1 To 4
                raw = stageSheet.Cells(r, k + 1).Value2
                If IsError(raw) Then _
                    Err.Raise vbObjectError + 757, , "Excel error at row " & r & ", column " & k + 1
                If Len(Trim$(CStr(raw))) > 0 Then
                    If Not IsNumeric(raw) Then _
                        Err.Raise vbObjectError + 758, , _
                            "Non-numeric target at row " & r & ", column " & k + 1
                    If CDbl(raw) < 0 Then _
                        Err.Raise vbObjectError + 759, , _
                            "Negative target at row " & r & ", column " & k + 1
                    hasValue = True
                End If
            Next k
            If Not hasValue Then _
                Err.Raise vbObjectError + 760, , "All four targets blank at row " & r
            staged.Add key, r
        Else
            For k = 2 To 5
                raw = stageSheet.Cells(r, k).Value2
                If IsError(raw) Then _
                    Err.Raise vbObjectError + 761, , "Error value without H40 in row " & r
                If Len(Trim$(CStr(raw))) > 0 Then _
                    Err.Raise vbObjectError + 762, , "Target value without H40 in row " & r
            Next k
        End If
    Next r
    If staged.Count = 0 Then Err.Raise vbObjectError + 763, , "No valid H40 targets found."

    If MsgBox("Import targets for " & staged.Count & " H40(s)?" & vbCrLf & _
              "Blank numeric cells preserve their existing targets." & vbCrLf & _
              "Other H40 targets and budget edits remain unchanged.", _
              vbQuestion + vbYesNo, "Confirm H40 target import") <> vbYes Then Exit Sub

    Application.EnableEvents = False
    For Each h In staged.Keys
        key = CStr(h)
        r = CLng(staged(h))
        If existing.Exists(key) Then
            rowTarget = CLng(existing(key))
        Else
            rowTarget = storeSheet.Cells(storeSheet.Rows.Count, 1).End(xlUp).Row + 1
            If rowTarget < 2 Then rowTarget = 2
            storeSheet.Cells(rowTarget, 1).NumberFormat = "@"
            storeSheet.Cells(rowTarget, 1).Value2 = key
            existing.Add key, rowTarget
        End If
        For k = 1 To 4
            raw = stageSheet.Cells(r, k + 1).Value2
            If Len(Trim$(CStr(raw))) > 0 Then
                storeSheet.Cells(rowTarget, k + 1).Value2 = CDbl(raw)
                applyCount = applyCount + 1
            End If
        Next k
        updated = updated + 1
    Next h

    ' Critical: V4 saves the currently selected E6:E9 on the NEXT H40 switch.
    ' Synchronize those cells now, or the old screen values would overwrite
    ' the newly imported targets for that same H40.
    If Len(currentKey) > 0 And existing.Exists(currentKey) Then
        rowTarget = CLng(existing(currentKey))
        If staged.Exists(currentKey) Then
            For k = 1 To 4
                raw = storeSheet.Cells(rowTarget, k + 1).Value2
                detail.Cells(k + 5, 5).Value2 = raw
            Next k
            matched = 1
        End If
    End If
    detail.Range("G6:G9").Calculate
    Application.EnableEvents = oldEvents
    MsgBox "Targets imported for " & updated & " H40(s), " & applyCount & _
           " nonblank metric values." & vbCrLf & _
           "Current H40 synchronized: " & IIf(matched = 1, "Yes", "Not in batch"), vbInformation
    Exit Sub
Failed:
    failure = "H40 target import failed (" & Err.Number & "): " & Err.Description
    Application.EnableEvents = oldEvents
    MsgBox failure & vbCrLf & "Check the paste sheet and backup before retrying.", vbExclamation
End Sub

Private Function TargetKey(ByVal raw As Variant) As String
    If IsError(raw) Then Err.Raise vbObjectError + 764, , "H40 cell contains Excel error."
    TargetKey = Trim$(Replace(CStr(raw), ChrW$(160), " "))
End Function
