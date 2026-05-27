'==============================================================================
' LEGAL CONTRACTS MIS TOOL — Complete VBA Solution
'
' HOW TO USE:
'   1. Open your new .xlsm file (or create one: File > New > Blank Workbook,
'      then Save As > Excel Macro-Enabled Workbook .xlsm)
'   2. Press Alt + F11 to open the VBA Editor
'   3. In the left panel, right-click your workbook name > Insert > Module
'   4. Name it "modMIS" (click the module, then in Properties window set Name)
'   5. Paste the entire contents of this file into that module
'   6. Separately, paste the SHEET CODE section into the Master sheet's
'      code module (double-click "Sheet: Master" in the left panel)
'   7. Paste the WORKBOOK CODE section into "ThisWorkbook"
'   8. Close VBA Editor, go back to Excel
'   9. Run SetupSheets first (Alt+F8 > SetupSheets > Run)
'  10. Run RemoveAllCF to nuke any old conditional formatting
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' COLUMN POSITIONS — edit these if your columns ever change order
'------------------------------------------------------------------------------
Public Const C_SNO       As Integer = 1    ' S. No.
Public Const C_DOCNAME   As Integer = 2    ' Document Name
Public Const C_VENDOR    As Integer = 3    ' Vendor / Party
Public Const C_SPOC      As Integer = 4    ' Business SPOC
Public Const C_DIVISION  As Integer = 5    ' Business Division
Public Const C_REV1      As Integer = 6    ' First Reviewer
Public Const C_REV2      As Integer = 7    ' Second Reviewer
Public Const C_DATERECD  As Integer = 8    ' Date of Receipt
Public Const C_DATEREPLY As Integer = 9    ' Date of Reply
Public Const C_VALUE     As Integer = 10   ' Contract Value
Public Const C_STATUS    As Integer = 11   ' Status
Public Const C_DUPFLAG   As Integer = 12   ' Duplicate flag (paste sheets only)

Public Const ROW_HEADER  As Integer = 1
Public Const ROW_DATA    As Integer = 2

'==============================================================================
' SECTION 1 — COLOR ENGINE
' Replaces all Conditional Formatting with direct VBA coloring.
' Zero CF rules = no file corruption.
'==============================================================================

Public Function StatusColor(sStatus As String) As Long
    Select Case Trim(sStatus)
        Case "Executed / Closed", "N/A - Advisory"
            StatusColor = RGB(198, 239, 206)   ' Light Green
        Case "Reverted to Counter-party", "Pending with Counter-party", _
             "Pending Execution", "Pending with Business"
            StatusColor = RGB(255, 199, 131)   ' Light Orange
        Case "Under Review", "Not Responded"
            StatusColor = RGB(255, 199, 199)   ' Light Red/Pink
        Case "On Hold", "Abandoned"
            StatusColor = RGB(189, 215, 238)   ' Light Blue
        Case Else
            StatusColor = RGB(255, 255, 255)   ' White (no status / blank)
    End Select
End Function

' Colors all data rows on a single sheet based on Status column
Public Sub ColorSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_DOCNAME).End(xlUp).Row
    If lastRow < ROW_DATA Then Exit Sub

    Application.ScreenUpdating = False
    Dim i As Long
    For i = ROW_DATA To lastRow
        Dim sVal As String
        sVal = Trim(ws.Cells(i, C_STATUS).Value)
        If sVal <> "" Then
            ws.Rows(i).Interior.Color = StatusColor(sVal)
        Else
            ws.Rows(i).Interior.ColorIndex = xlNone
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

' Recolors all data sheets — called on file open, save, and by button
Public Sub ColorAllSheets()
    Application.ScreenUpdating = False
    Dim names As Variant
    names = Array("Master", "Paste_NewWork", "Paste_StatusUpdates", "CombinedView")
    Dim i As Integer
    For i = 0 To UBound(names)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(names(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ColorSheet ws
            Set ws = Nothing
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

' Button-friendly wrapper that shows a confirmation message
Public Sub ColorAllSheets_Button()
    Call ColorAllSheets
    MsgBox "All sheets recolored by status.", vbInformation, "MIS Tool"
End Sub

' ONE-TIME CLEANUP: removes all CF rules from every sheet in this file.
' Run this once on your old file before migrating data to the new file.
Public Sub RemoveAllCF()
    If MsgBox("Remove ALL Conditional Formatting rules from every sheet?" & vbNewLine & _
              "This cannot be undone. Continue?", _
              vbYesNo + vbQuestion, "Remove CF Rules") = vbNo Then Exit Sub

    Dim ws As Worksheet
    Dim total As Long
    For Each ws In ThisWorkbook.Worksheets
        total = total + ws.Cells.FormatConditions.Count
        ws.Cells.FormatConditions.Delete
    Next ws
    MsgBox "Done. Removed " & total & " conditional formatting rule(s).", _
           vbInformation, "CF Removed"
End Sub

'==============================================================================
' SECTION 2 — AUDIT LOG
' Every status change is recorded: timestamp, user, document, old value, new value.
' The AuditLog sheet is hidden (xlSheetVeryHidden) — use ViewAuditLog to see it.
'==============================================================================

Private Const AUDIT_SHEET As String = "AuditLog"

Private Sub EnsureAuditSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(AUDIT_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = AUDIT_SHEET
    ws.Visible = xlSheetVeryHidden
    ws.Cells(1, 1).Value = "Timestamp"
    ws.Cells(1, 2).Value = "User"
    ws.Cells(1, 3).Value = "Sheet"
    ws.Cells(1, 4).Value = "Row"
    ws.Cells(1, 5).Value = "Document Name"
    ws.Cells(1, 6).Value = "Vendor / Party"
    ws.Cells(1, 7).Value = "Field Changed"
    ws.Cells(1, 8).Value = "Old Value"
    ws.Cells(1, 9).Value = "New Value"
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(68, 114, 196)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
End Sub

Public Sub LogChange(sheetName As String, rowNum As Long, docName As String, _
                     vendor As String, fieldName As String, _
                     oldVal As String, newVal As String)
    Call EnsureAuditSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(AUDIT_SHEET)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = Application.UserName
    ws.Cells(nextRow, 3).Value = sheetName
    ws.Cells(nextRow, 4).Value = rowNum
    ws.Cells(nextRow, 5).Value = docName
    ws.Cells(nextRow, 6).Value = vendor
    ws.Cells(nextRow, 7).Value = fieldName
    ws.Cells(nextRow, 8).Value = oldVal
    ws.Cells(nextRow, 9).Value = newVal
End Sub

Public Sub ViewAuditLog()
    Call EnsureAuditSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(AUDIT_SHEET)
    ws.Visible = xlSheetVisible
    ws.Columns.AutoFit
    ws.Activate
End Sub

Public Sub HideAuditLog()
    On Error Resume Next
    ThisWorkbook.Sheets(AUDIT_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
    ThisWorkbook.Sheets("Master").Activate
End Sub

'==============================================================================
' SECTION 3 — DUPLICATE CHECK
' Compares Document Name + Vendor against the Master sheet.
' Run CheckForDuplicates after pasting into Paste_NewWork.
'==============================================================================

' Returns the row number in Master where docName+vendor match, or 0 if not found
Public Function FindInMaster(docName As String, vendor As String) As Long
    Dim masterWs As Worksheet
    On Error Resume Next
    Set masterWs = ThisWorkbook.Sheets("Master")
    On Error GoTo 0
    If masterWs Is Nothing Then FindInMaster = 0 : Exit Function

    Dim lastRow As Long
    lastRow = masterWs.Cells(masterWs.Rows.Count, C_DOCNAME).End(xlUp).Row
    Dim searchDoc As String : searchDoc = Trim(LCase(docName))
    Dim searchVnd As String : searchVnd = Trim(LCase(vendor))

    Dim i As Long
    For i = ROW_DATA To lastRow
        If Trim(LCase(masterWs.Cells(i, C_DOCNAME).Value)) = searchDoc And _
           Trim(LCase(masterWs.Cells(i, C_VENDOR).Value)) = searchVnd Then
            FindInMaster = i
            Exit Function
        End If
    Next i
    FindInMaster = 0
End Function

' Flags each row in Paste_NewWork as "DUPLICATE" or "New" in column L
Public Sub CheckForDuplicates()
    Dim pasteWs As Worksheet
    On Error Resume Next
    Set pasteWs = ThisWorkbook.Sheets("Paste_NewWork")
    On Error GoTo 0
    If pasteWs Is Nothing Then
        MsgBox "Paste_NewWork sheet not found.", vbExclamation : Exit Sub
    End If

    Dim lastRow As Long
    lastRow = pasteWs.Cells(pasteWs.Rows.Count, C_DOCNAME).End(xlUp).Row
    If lastRow < ROW_DATA Then
        MsgBox "No data found in Paste_NewWork.", vbInformation : Exit Sub
    End If

    ' Clear previous flags
    pasteWs.Columns(C_DUPFLAG).ClearContents
    pasteWs.Columns(C_DUPFLAG).Interior.ColorIndex = xlNone
    pasteWs.Cells(ROW_HEADER, C_DUPFLAG).Value = "Duplicate Check"
    pasteWs.Cells(ROW_HEADER, C_DUPFLAG).Font.Bold = True

    Dim dupCount As Integer : dupCount = 0
    Dim i As Long
    For i = ROW_DATA To lastRow
        Dim docName As String : docName = Trim(pasteWs.Cells(i, C_DOCNAME).Value)
        Dim vendor As String  : vendor  = Trim(pasteWs.Cells(i, C_VENDOR).Value)
        If docName = "" Then GoTo NextDupRow

        If FindInMaster(docName, vendor) > 0 Then
            pasteWs.Cells(i, C_DUPFLAG).Value = "DUPLICATE"
            pasteWs.Cells(i, C_DUPFLAG).Interior.Color = RGB(255, 199, 199)
            pasteWs.Cells(i, C_DUPFLAG).Font.Bold = True
            dupCount = dupCount + 1
        Else
            pasteWs.Cells(i, C_DUPFLAG).Value = "New"
            pasteWs.Cells(i, C_DUPFLAG).Interior.Color = RGB(198, 239, 206)
        End If
NextDupRow:
    Next i

    pasteWs.Columns(C_DUPFLAG).AutoFit

    If dupCount > 0 Then
        MsgBox dupCount & " duplicate(s) found — check column L before adding to Master.", _
               vbExclamation, "Duplicates Found"
    Else
        MsgBox "No duplicates. All " & (lastRow - ROW_HEADER) & " entries are new.", _
               vbInformation, "Duplicate Check"
    End If
End Sub

'==============================================================================
' SECTION 4 — STATUS RECONCILIATION
' Step 1: ReconcileStatusUpdates — compares Paste_StatusUpdates against Master
' Step 2: Review the StatusReconciliation sheet; change "Pending" to "Skip" to ignore
' Step 3: ApplyApprovedUpdates — writes approved changes to Master + logs them
'==============================================================================

Public Sub ReconcileStatusUpdates()
    Dim masterWs As Worksheet, updateWs As Worksheet, reconWs As Worksheet
    On Error Resume Next
    Set masterWs = ThisWorkbook.Sheets("Master")
    Set updateWs = ThisWorkbook.Sheets("Paste_StatusUpdates")
    Set reconWs  = ThisWorkbook.Sheets("StatusReconciliation")
    On Error GoTo 0
    If masterWs Is Nothing Or updateWs Is Nothing Or reconWs Is Nothing Then
        MsgBox "Could not find Master, Paste_StatusUpdates, or StatusReconciliation sheet.", _
               vbExclamation : Exit Sub
    End If

    ' Clear old reconciliation data
    reconWs.Rows(ROW_DATA & ":" & reconWs.Rows.Count).ClearContents
    reconWs.Rows(ROW_DATA & ":" & reconWs.Rows.Count).Interior.ColorIndex = xlNone

    ' Set headers
    Dim hdr As Variant
    hdr = Array("Master Row", "Document Name", "Vendor / Party", _
                "Status in Master", "Proposed Status", "Changed?", "Action")
    Dim h As Integer
    For h = 0 To UBound(hdr)
        reconWs.Cells(ROW_HEADER, h + 1).Value = CStr(hdr(h))
    Next h
    reconWs.Rows(ROW_HEADER).Font.Bold = True
    reconWs.Rows(ROW_HEADER).Interior.Color = RGB(68, 114, 196)
    reconWs.Rows(ROW_HEADER).Font.Color = RGB(255, 255, 255)

    Dim lastRow As Long
    lastRow = updateWs.Cells(updateWs.Rows.Count, C_DOCNAME).End(xlUp).Row

    Dim reconRow As Long : reconRow = ROW_DATA
    Dim changedCnt As Integer : changedCnt = 0
    Dim notFoundCnt As Integer : notFoundCnt = 0

    Dim i As Long
    For i = ROW_DATA To lastRow
        Dim docName As String : docName = Trim(updateWs.Cells(i, C_DOCNAME).Value)
        Dim vendor As String  : vendor  = Trim(updateWs.Cells(i, C_VENDOR).Value)
        Dim newSt As String   : newSt   = Trim(updateWs.Cells(i, C_STATUS).Value)
        If docName = "" Then GoTo NextReconRow

        Dim mRow As Long : mRow = FindInMaster(docName, vendor)

        reconWs.Cells(reconRow, 1).Value = mRow
        reconWs.Cells(reconRow, 2).Value = docName
        reconWs.Cells(reconRow, 3).Value = vendor
        reconWs.Cells(reconRow, 5).Value = newSt

        If mRow > 0 Then
            Dim curSt As String
            curSt = Trim(masterWs.Cells(mRow, C_STATUS).Value)
            reconWs.Cells(reconRow, 4).Value = curSt
            If curSt <> newSt And newSt <> "" Then
                reconWs.Cells(reconRow, 6).Value = "YES"
                reconWs.Cells(reconRow, 6).Interior.Color = RGB(255, 199, 131)
                reconWs.Cells(reconRow, 7).Value = "Pending"
                changedCnt = changedCnt + 1
            Else
                reconWs.Cells(reconRow, 6).Value = "No Change"
                reconWs.Cells(reconRow, 6).Interior.Color = RGB(235, 235, 235)
                reconWs.Cells(reconRow, 7).Value = "-"
            End If
        Else
            reconWs.Cells(reconRow, 4).Value = "NOT FOUND IN MASTER"
            reconWs.Cells(reconRow, 4).Interior.Color = RGB(255, 199, 199)
            reconWs.Cells(reconRow, 6).Value = "NOT FOUND"
            reconWs.Cells(reconRow, 7).Value = "Review Manually"
            notFoundCnt = notFoundCnt + 1
        End If

        reconRow = reconRow + 1
NextReconRow:
    Next i

    reconWs.Columns.AutoFit
    reconWs.Activate
    MsgBox "Reconciliation complete." & vbNewLine & _
           "  Status changes: " & changedCnt & vbNewLine & _
           "  Not found in Master: " & notFoundCnt & vbNewLine & vbNewLine & _
           "Review the Action column. Change 'Pending' to 'Skip' for any rows " & _
           "you don't want applied. Then run Apply Approved Updates.", _
           vbInformation, "Reconciliation"
End Sub

' Applies all rows where Action = "Pending" from StatusReconciliation to Master
Public Sub ApplyApprovedUpdates()
    Dim masterWs As Worksheet, reconWs As Worksheet
    On Error Resume Next
    Set masterWs = ThisWorkbook.Sheets("Master")
    Set reconWs  = ThisWorkbook.Sheets("StatusReconciliation")
    On Error GoTo 0
    If masterWs Is Nothing Or reconWs Is Nothing Then
        MsgBox "Required sheets not found.", vbExclamation : Exit Sub
    End If

    Dim lastRow As Long
    lastRow = reconWs.Cells(reconWs.Rows.Count, 2).End(xlUp).Row

    Dim pendingCnt As Integer : pendingCnt = 0
    Dim i As Long
    For i = ROW_DATA To lastRow
        If Trim(reconWs.Cells(i, 7).Value) = "Pending" Then pendingCnt = pendingCnt + 1
    Next i

    If pendingCnt = 0 Then
        MsgBox "No rows with Action = 'Pending'. Nothing to apply.", vbInformation : Exit Sub
    End If

    If MsgBox("Apply " & pendingCnt & " status update(s) to the Master sheet?" & vbNewLine & _
              "All changes will be recorded in the Audit Log.", _
              vbYesNo + vbQuestion, "Apply Updates") = vbNo Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim appliedCnt As Integer : appliedCnt = 0
    For i = ROW_DATA To lastRow
        If Trim(reconWs.Cells(i, 7).Value) = "Pending" Then
            Dim mRow As Long      : mRow    = reconWs.Cells(i, 1).Value
            Dim docName As String : docName = reconWs.Cells(i, 2).Value
            Dim vendor As String  : vendor  = reconWs.Cells(i, 3).Value
            Dim oldSt As String   : oldSt   = reconWs.Cells(i, 4).Value
            Dim newSt As String   : newSt   = reconWs.Cells(i, 5).Value

            If mRow > 0 Then
                Call LogChange("Master", mRow, docName, vendor, "Status", oldSt, newSt)
                masterWs.Cells(mRow, C_STATUS).Value = newSt
                masterWs.Rows(mRow).Interior.Color = StatusColor(newSt)
                reconWs.Cells(i, 7).Value = "Applied " & Format(Now(), "dd-mmm hh:mm")
                reconWs.Cells(i, 7).Interior.Color = RGB(198, 239, 206)
                appliedCnt = appliedCnt + 1
            End If
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox appliedCnt & " update(s) applied to Master successfully.", vbInformation, "Done"
End Sub

'==============================================================================
' SECTION 5 — MONTHLY SUMMARY
' Refreshes MonthlySummary sheet with:
'   - Status breakdown with counts and percentages
'   - Reviewer workload (total / closed / active)
'   - Overdue contracts (receipt >30 days ago, still active status)
'==============================================================================

Public Sub RefreshMonthlySummary()
    Dim masterWs As Worksheet, summaryWs As Worksheet
    On Error Resume Next
    Set masterWs  = ThisWorkbook.Sheets("Master")
    Set summaryWs = ThisWorkbook.Sheets("MonthlySummary")
    On Error GoTo 0
    If masterWs Is Nothing Or summaryWs Is Nothing Then
        MsgBox "Master or MonthlySummary sheet not found.", vbExclamation : Exit Sub
    End If

    Application.ScreenUpdating = False
    summaryWs.Cells.ClearContents
    summaryWs.Cells.Interior.ColorIndex = xlNone

    Dim lastRow As Long
    lastRow = masterWs.Cells(masterWs.Rows.Count, C_DOCNAME).End(xlUp).Row
    Dim totalRows As Long : totalRows = lastRow - ROW_HEADER
    Dim r As Integer : r = 1

    '--- Status Breakdown ---
    summaryWs.Cells(r, 1).Value = "STATUS BREAKDOWN  (as of " & Format(Date, "dd-mmm-yyyy") & ")"
    summaryWs.Cells(r, 1).Font.Bold = True
    summaryWs.Cells(r, 1).Font.Size = 12
    r = r + 1
    summaryWs.Cells(r, 1).Value = "Status"
    summaryWs.Cells(r, 2).Value = "Count"
    summaryWs.Cells(r, 3).Value = "% of Total"
    summaryWs.Rows(r).Font.Bold = True
    r = r + 1

    Dim statuses As Variant
    statuses = Array("Executed / Closed", "N/A - Advisory", _
                     "Reverted to Counter-party", "Pending with Counter-party", _
                     "Under Review", "Pending Execution", "Pending with Business", _
                     "Not Responded", "On Hold", "Abandoned")

    Dim s As Integer
    For s = 0 To UBound(statuses)
        Dim sName As String : sName = CStr(statuses(s))
        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIf(masterWs.Columns(C_STATUS), sName)
        summaryWs.Cells(r, 1).Value = sName
        summaryWs.Cells(r, 2).Value = cnt
        If totalRows > 0 Then
            summaryWs.Cells(r, 3).Value = cnt / totalRows
            summaryWs.Cells(r, 3).NumberFormat = "0.0%"
        End If
        summaryWs.Rows(r).Interior.Color = StatusColor(sName)
        r = r + 1
    Next s
    summaryWs.Cells(r, 1).Value = "TOTAL"
    summaryWs.Cells(r, 2).Value = totalRows
    summaryWs.Cells(r, 1).Font.Bold = True
    summaryWs.Cells(r, 2).Font.Bold = True
    r = r + 2

    '--- Reviewer Workload ---
    summaryWs.Cells(r, 1).Value = "REVIEWER WORKLOAD"
    summaryWs.Cells(r, 1).Font.Bold = True
    summaryWs.Cells(r, 1).Font.Size = 12
    r = r + 1
    summaryWs.Cells(r, 1).Value = "Reviewer"
    summaryWs.Cells(r, 2).Value = "Total Assigned"
    summaryWs.Cells(r, 3).Value = "Closed / Advisory"
    summaryWs.Cells(r, 4).Value = "Active"
    summaryWs.Rows(r).Font.Bold = True
    r = r + 1

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = ROW_DATA To lastRow
        Dim rev1 As String : rev1 = Trim(masterWs.Cells(j, C_REV1).Value)
        If rev1 <> "" And Not dict.Exists(rev1) Then dict.Add rev1, 0
    Next j

    Dim key As Variant
    For Each key In dict.Keys
        Dim revName As String : revName = CStr(key)
        Dim total As Long
        Dim closed As Long
        total = Application.WorksheetFunction.CountIf(masterWs.Columns(C_REV1), revName)
        closed = Application.WorksheetFunction.CountIfs( _
                     masterWs.Columns(C_REV1), revName, _
                     masterWs.Columns(C_STATUS), "Executed / Closed") + _
                 Application.WorksheetFunction.CountIfs( _
                     masterWs.Columns(C_REV1), revName, _
                     masterWs.Columns(C_STATUS), "N/A - Advisory")
        summaryWs.Cells(r, 1).Value = revName
        summaryWs.Cells(r, 2).Value = total
        summaryWs.Cells(r, 3).Value = closed
        summaryWs.Cells(r, 4).Value = total - closed
        r = r + 1
    Next key
    r = r + 1

    '--- Overdue Contracts (>30 days open, still active) ---
    summaryWs.Cells(r, 1).Value = "OVERDUE CONTRACTS  (receipt >30 days ago, still active)"
    summaryWs.Cells(r, 1).Font.Bold = True
    summaryWs.Cells(r, 1).Font.Size = 12
    r = r + 1
    summaryWs.Cells(r, 1).Value = "Document Name"
    summaryWs.Cells(r, 2).Value = "Vendor / Party"
    summaryWs.Cells(r, 3).Value = "Days Open"
    summaryWs.Cells(r, 4).Value = "Status"
    summaryWs.Cells(r, 5).Value = "First Reviewer"
    summaryWs.Rows(r).Font.Bold = True
    r = r + 1

    Dim activeStatuses As String
    activeStatuses = "Under Review|Pending with Business|Pending Execution|" & _
                     "Not Responded|Pending with Counter-party"

    For j = ROW_DATA To lastRow
        Dim st As String : st = Trim(masterWs.Cells(j, C_STATUS).Value)
        If InStr(activeStatuses, st) > 0 Then
            Dim dateRecd As Date
            Dim daysOpen As Long
            On Error Resume Next
            dateRecd = CDate(masterWs.Cells(j, C_DATERECD).Value)
            On Error GoTo 0
            If dateRecd > 0 Then
                daysOpen = Date - dateRecd
                If daysOpen > 30 Then
                    summaryWs.Cells(r, 1).Value = masterWs.Cells(j, C_DOCNAME).Value
                    summaryWs.Cells(r, 2).Value = masterWs.Cells(j, C_VENDOR).Value
                    summaryWs.Cells(r, 3).Value = daysOpen
                    summaryWs.Cells(r, 4).Value = st
                    summaryWs.Cells(r, 5).Value = masterWs.Cells(j, C_REV1).Value
                    summaryWs.Rows(r).Interior.Color = RGB(255, 199, 199)
                    r = r + 1
                End If
            End If
        End If
    Next j

    summaryWs.Columns.AutoFit
    summaryWs.Activate
    Application.ScreenUpdating = True
    MsgBox "Monthly Summary refreshed.", vbInformation, "MIS Tool"
End Sub

'==============================================================================
' SECTION 6 — SETUP UTILITIES
' Run SetupSheets once when you create a new file.
' Run AddStatusDropdowns to add/refresh the Status dropdown on all data sheets.
'==============================================================================

Public Sub SetupSheets()
    ' Creates all required sheets if they don't already exist
    Dim sheetList As Variant
    sheetList = Array("Master", "Paste_NewWork", "Paste_StatusUpdates", _
                      "StatusReconciliation", "CombinedView", "MonthlySummary")

    Dim i As Integer
    For i = 0 To UBound(sheetList)
        Dim wsName As String : wsName = CStr(sheetList(i))
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(wsName)
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = wsName
        End If
        Set ws = Nothing
    Next i

    ' Apply standard headers to the three data sheets
    Dim dataSheets As Variant
    dataSheets = Array("Master", "Paste_NewWork", "Paste_StatusUpdates")
    Dim headers As Variant
    headers = Array("S. No.", "Document Name", "Vendor / Party", "Business SPOC", _
                    "Business Division", "First Reviewer", "Second Reviewer", _
                    "Date of Receipt", "Date of Reply", "Contract Value", "Status")

    Dim di As Integer
    For di = 0 To UBound(dataSheets)
        Dim dataWs As Worksheet
        Set dataWs = ThisWorkbook.Sheets(CStr(dataSheets(di)))
        Dim col As Integer
        For col = 0 To UBound(headers)
            dataWs.Cells(ROW_HEADER, col + 1).Value = CStr(headers(col))
        Next col
        With dataWs.Rows(ROW_HEADER)
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
        Set dataWs = Nothing
    Next di

    Call EnsureAuditSheet
    Call AddStatusDropdowns

    MsgBox "Setup complete!" & vbNewLine & vbNewLine & _
           "Sheets created: Master, Paste_NewWork, Paste_StatusUpdates," & vbNewLine & _
           "StatusReconciliation, CombinedView, MonthlySummary, AuditLog (hidden)." & vbNewLine & vbNewLine & _
           "Next step: paste your 700+ rows of data into the Master sheet.", _
           vbInformation, "Setup Complete"
End Sub

Public Sub AddStatusDropdowns()
    Dim statusList As String
    statusList = "Executed / Closed,Reverted to Counter-party,Pending with Counter-party," & _
                 "Under Review,Pending Execution,Pending with Business," & _
                 "Not Responded,On Hold,Abandoned,N/A - Advisory"

    Dim sheetNames As Variant
    sheetNames = Array("Master", "Paste_NewWork", "Paste_StatusUpdates")
    Dim i As Integer
    For i = 0 To UBound(sheetNames)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(ROW_DATA, C_STATUS), ws.Cells(ws.Rows.Count, C_STATUS))
            rng.Validation.Delete
            With rng.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=statusList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowError = True
                .ErrorTitle = "Invalid Status"
                .ErrorMessage = "Please select a valid status from the list."
            End With
            Set ws = Nothing
        End If
    Next i
End Sub


'==============================================================================
' *** SHEET CODE — paste this into the Master sheet's code module ***
'
' How to get there:
'   Alt+F11 > in left panel, expand your workbook > double-click "Master"
'   (it may show as "Sheet1 (Master)" — double-click that)
'   Paste the block below into the white area on the right.
'==============================================================================
'
' Option Explicit
'
' Private oldStatusVal As String
'
' ' Captures the current cell value BEFORE the user changes it
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'     If Target.Cells.Count = 1 And Target.Column = C_STATUS Then
'         oldStatusVal = Trim(Target.Value)
'     End If
' End Sub
'
' ' Fires after the user edits a cell — recolors the row and logs the change
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Target.Column <> C_STATUS Then Exit Sub
'     If Target.Row <= ROW_HEADER Then Exit Sub
'     If Target.Cells.Count > 1 Then Exit Sub   ' ignore bulk paste
'
'     Dim newVal As String : newVal = Trim(Target.Value)
'
'     Application.EnableEvents = False
'     Target.EntireRow.Interior.Color = StatusColor(newVal)
'     Application.EnableEvents = True
'
'     Dim docName As String : docName = Me.Cells(Target.Row, C_DOCNAME).Value
'     Dim vendor As String  : vendor  = Me.Cells(Target.Row, C_VENDOR).Value
'     Call LogChange(Me.Name, Target.Row, docName, vendor, "Status", oldStatusVal, newVal)
' End Sub
'
'==============================================================================
' *** WORKBOOK CODE — paste this into ThisWorkbook's code module ***
'
' How to get there:
'   Alt+F11 > in left panel, double-click "ThisWorkbook"
'   Paste the block below.
'==============================================================================
'
' Option Explicit
'
' Private Sub Workbook_Open()
'     Call ColorAllSheets
' End Sub
'
' Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'     Call ColorAllSheets   ' silently recolors before every save
' End Sub
'
'==============================================================================
