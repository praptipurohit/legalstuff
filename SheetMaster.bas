'==============================================================================
' Sheet: Master — paste this into the Master sheet's code module
' (Alt+F11 → double-click "Sheet2 (Master)" in left panel → paste here)
'==============================================================================

Option Explicit

Private oldStatusVal As String

' Captures the current cell value BEFORE the user edits it
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 And Target.Column = C_STATUS Then
        oldStatusVal = Trim(Target.Value)
    End If
End Sub

' Fires after a cell is edited — recolors the row instantly and logs the change
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column <> C_STATUS Then Exit Sub
    If Target.Row <= ROW_HEADER Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub   ' ignore bulk paste

    Dim newVal As String : newVal = Trim(Target.Value)

    Application.EnableEvents = False
    Target.EntireRow.Interior.Color = StatusColor(newVal)
    Application.EnableEvents = True

    Dim docName As String : docName = Me.Cells(Target.Row, C_DOCNAME).Value
    Dim vendor  As String : vendor  = Me.Cells(Target.Row, C_VENDOR).Value
    Call LogChange(Me.Name, Target.Row, docName, vendor, "Status", oldStatusVal, newVal)
End Sub
