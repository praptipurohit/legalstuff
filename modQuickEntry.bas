'==============================================================================
' modQuickEntry — Legal Contracts MIS Quick Entry
' Paste this entire block into a new module named "modQuickEntry"
'
' Shortcuts (registered in ThisWorkbook):
'   Ctrl+Shift+S  Stage one entry
'   Ctrl+Shift+A  Commit all staged entries to Master
'   Ctrl+Shift+F  Clear the form
'   Ctrl+Shift+X  Clear the staging area
'==============================================================================

Option Explicit

Private Const QE_SHEET     As String = "Quick Entry"
Private Const QE_FORM_COL  As Long   = 3     ' Column C holds form inputs
Private Const QE_STAGE_HDR As Long   = 16    ' Staging header row
Private Const QE_STAGE_MAX As Long   = 30    ' Max staged rows

' Form input rows (all values entered in column C)
Private Const QF_DOCNAME   As Long = 4
Private Const QF_VENDOR    As Long = 5
Private Const QF_SPOC      As Long = 6
Private Const QF_DIVISION  As Long = 7
Private Const QF_REV1      As Long = 8
Private Const QF_REV2      As Long = 9
Private Const QF_DATERECD  As Long = 10
Private Const QF_DATEREPLY As Long = 11
Private Const QF_VALUE     As Long = 12
Private Const QF_STATUS    As Long = 13

'------------------------------------------------------------------------------
' QE_AddToStaging
' Validates form, checks for duplicates, writes to staging area.
' Shortcut: Ctrl+Shift+S
'------------------------------------------------------------------------------
Public Sub QE_AddToStaging()
    Dim ws As Worksheet
    On Error GoTo SheetMissing
    Set ws = ThisWorkbook.Sheets(QE_SHEET)
    On Error GoTo 0

    ' Read all form values
    Dim docName  As String : docName  = Trim(CStr(ws.Cells(QF_DOCNAME,   QE_FORM_COL).Value))
    Dim vendor   As String : vendor   = Trim(CStr(ws.Cells(QF_VENDOR,    QE_FORM_COL).Value))
    Dim spoc     As String : spoc     = Trim(CStr(ws.Cells(QF_SPOC,      QE_FORM_COL).Value))
    Dim div      As String : div      = Trim(CStr(ws.Cells(QF_DIVISION,  QE_FORM_COL).Value))
    Dim rev1     As String : rev1     = Trim(CStr(ws.Cells(QF_REV1,      QE_FORM_COL).Value))
    Dim rev2     As String : rev2     = Trim(CStr(ws.Cells(QF_REV2,      QE_FORM_COL).Value))
    Dim dateRcd  As String : dateRcd  = Trim(CStr(ws.Cells(QF_DATERECD,  QE_FORM_COL).Value))
    Dim dateRep  As String : dateRep  = Trim(CStr(ws.Cells(QF_DATEREPLY, QE_FORM_COL).Value))
    Dim contVal  As String : contVal  = Trim(CStr(ws.Cells(QF_VALUE,     QE_FORM_COL).Value))
    Dim status   As String : status   = Trim(CStr(ws.Cells(QF_STATUS,    QE_FORM_COL).Value))

    ' Validate required fields
    Dim missing As String : missing = ""
    If docName = "" Then missing = missing & vbCrLf & "  - Document Name"
    If vendor  = "" Then missing = missing & vbCrLf & "  - Vendor / Party"
    If status  = "" Then missing = missing & vbCrLf & "  - Status"
    If missing <> "" Then
        MsgBox "Please fill these required fields:" & missing, vbExclamation, "Missing Fields"
        Exit Sub
    End If

    ' Duplicate check against Master
    If FindInMaster(docName, vendor) > 0 Then
        If MsgBox("This contract already exists in Master:" & vbCrLf & vbCrLf & _
                  "  " & docName & vbCrLf & _
                  "  " & vendor & vbCrLf & vbCrLf & _
                  "Stage it anyway?", _
                  vbYesNo + vbQuestion, "Duplicate Found") = vbNo Then
            Exit Sub
        End If
    End If

    ' Find next empty staging row (keyed on Document Name column = col 2)
    Dim r As Long, stageRow As Long : stageRow = 0
    For r = QE_STAGE_HDR + 1 To QE_STAGE_HDR + QE_STAGE_MAX
        If Trim(CStr(ws.Cells(r, 2).Value)) = "" Then
            stageRow = r
            Exit For
        End If
    Next r

    If stageRow = 0 Then
        MsgBox "Staging area is full (" & QE_STAGE_MAX & " rows)." & vbCrLf & _
               "Commit to Master (Ctrl+Shift+A) or clear staging (Ctrl+Shift+X) first.", _
               vbExclamation, "Staging Full"
        Exit Sub
    End If

    ' Write row to staging (columns A=1 through K=11)
    Application.ScreenUpdating = False
    With ws
        .Cells(stageRow, 1).Value  = ""       ' S.No assigned on commit
        .Cells(stageRow, 2).Value  = docName
        .Cells(stageRow, 3).Value  = vendor
        .Cells(stageRow, 4).Value  = spoc
        .Cells(stageRow, 5).Value  = div
        .Cells(stageRow, 6).Value  = rev1
        .Cells(stageRow, 7).Value  = rev2
        .Cells(stageRow, 8).Value  = dateRcd
        .Cells(stageRow, 9).Value  = dateRep
        .Cells(stageRow, 10).Value = contVal
        .Cells(stageRow, 11).Value = status
        .Rows(stageRow).Interior.Color = StatusColor(status)
    End With

    QE_ClearForm
    ws.Cells(QF_DOCNAME, QE_FORM_COL).Select
    Application.ScreenUpdating = True

    Application.StatusBar = "Staged " & (stageRow - QE_STAGE_HDR) & _
                            " of " & QE_STAGE_MAX & ":  " & docName & "  /  " & vendor
    Exit Sub

SheetMissing:
    MsgBox "Sheet '" & QE_SHEET & "' not found.", vbCritical, "Sheet Missing"
End Sub


'------------------------------------------------------------------------------
' QE_ClearForm
' Empties all form input cells.
' Shortcut: Ctrl+Shift+F
'------------------------------------------------------------------------------
Public Sub QE_ClearForm()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(QE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = QF_DOCNAME To QF_STATUS
        ws.Cells(r, QE_FORM_COL).ClearContents
    Next r
    ws.Cells(QF_DOCNAME, QE_FORM_COL).Select
    Application.StatusBar = False
End Sub


'------------------------------------------------------------------------------
' QE_CommitAll
' Moves all filled staging rows to Master, logs each to AuditLog.
' Shortcut: Ctrl+Shift+A
'------------------------------------------------------------------------------
Public Sub QE_CommitAll()
    Dim wsQE As Worksheet, masterWs As Worksheet
    On Error GoTo SheetMissing
    Set wsQE     = ThisWorkbook.Sheets(QE_SHEET)
    Set masterWs = ThisWorkbook.Sheets("Master")
    On Error GoTo 0

    ' Find next empty Master row
    Dim masterRow As Long
    masterRow = masterWs.Cells(masterWs.Rows.Count, C_DOCNAME).End(xlUp).Row + 1

    ' Highest existing S.No — continue from there
    Dim nextSNo As Long
    On Error Resume Next
    nextSNo = Application.WorksheetFunction.Max(masterWs.Columns(C_SNO))
    On Error GoTo 0
    If nextSNo < 0 Then nextSNo = 0

    Dim committed As Long : committed = 0
    Dim r As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For r = QE_STAGE_HDR + 1 To QE_STAGE_HDR + QE_STAGE_MAX
        Dim docName As String : docName = Trim(CStr(wsQE.Cells(r, 2).Value))
        If docName = "" Then GoTo NextRow

        Dim vendor  As String : vendor  = Trim(CStr(wsQE.Cells(r, 3).Value))
        Dim status  As String : status  = Trim(CStr(wsQE.Cells(r, 11).Value))
        nextSNo = nextSNo + 1

        With masterWs
            .Cells(masterRow, C_SNO).Value       = nextSNo
            .Cells(masterRow, C_DOCNAME).Value   = docName
            .Cells(masterRow, C_VENDOR).Value    = vendor
            .Cells(masterRow, C_SPOC).Value      = wsQE.Cells(r, 4).Value
            .Cells(masterRow, C_DIVISION).Value  = wsQE.Cells(r, 5).Value
            .Cells(masterRow, C_REV1).Value      = wsQE.Cells(r, 6).Value
            .Cells(masterRow, C_REV2).Value      = wsQE.Cells(r, 7).Value
            .Cells(masterRow, C_DATERECD).Value  = wsQE.Cells(r, 8).Value
            .Cells(masterRow, C_DATEREPLY).Value = wsQE.Cells(r, 9).Value
            .Cells(masterRow, C_VALUE).Value     = wsQE.Cells(r, 10).Value
            .Cells(masterRow, C_STATUS).Value    = status
            .Rows(masterRow).Interior.Color = StatusColor(status)
        End With

        Call LogChange("Master", masterRow, docName, vendor, _
                       "New Entry via Quick Entry", "", status)

        masterRow = masterRow + 1
        committed = committed + 1
NextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If committed = 0 Then
        MsgBox "Staging is empty. Use Ctrl+Shift+S to stage entries first.", _
               vbInformation, "Nothing to Commit"
        Exit Sub
    End If

    If MsgBox(committed & " contract(s) added to Master." & vbCrLf & vbCrLf & _
              "Clear staging area now?", _
              vbInformation + vbYesNo, "Committed Successfully") = vbYes Then
        Call QE_ClearStaging
    End If
    Exit Sub

SheetMissing:
    MsgBox "Required sheets not found ('Quick Entry' and 'Master').", vbCritical
End Sub


'------------------------------------------------------------------------------
' QE_ClearStaging
' Empties all staging rows and removes colour.
' Shortcut: Ctrl+Shift+X
'------------------------------------------------------------------------------
Public Sub QE_ClearStaging()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(QE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Dim r As Long, c As Long
    For r = QE_STAGE_HDR + 1 To QE_STAGE_HDR + QE_STAGE_MAX
        For c = 1 To 11
            ws.Cells(r, c).ClearContents
        Next c
        ws.Rows(r).Interior.ColorIndex = xlNone
    Next r
    ws.Cells(QF_DOCNAME, QE_FORM_COL).Select
    Application.ScreenUpdating = True
    Application.StatusBar = "Staging cleared."
End Sub
