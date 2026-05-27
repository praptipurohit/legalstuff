'==============================================================================
' ThisWorkbook — paste this into the ThisWorkbook code module
' (Alt+F11 → double-click ThisWorkbook in left panel → paste here)
'==============================================================================

Option Explicit

Private Sub Workbook_Open()
    ' Recolor all data sheets on open
    Call ColorAllSheets

    ' Register Quick Entry keyboard shortcuts
    Application.OnKey "^+s", "QE_AddToStaging"
    Application.OnKey "^+a", "QE_CommitAll"
    Application.OnKey "^+f", "QE_ClearForm"
    Application.OnKey "^+x", "QE_ClearStaging"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Silently recolor before every save — no popup
    Call ColorAllSheets
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Release keyboard shortcuts so they don't affect other workbooks
    Application.OnKey "^+s"
    Application.OnKey "^+a"
    Application.OnKey "^+f"
    Application.OnKey "^+x"
End Sub
