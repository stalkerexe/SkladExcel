' Component: Лист10  [Главная]
' Type: Document (Sheet / ThisWorkbook)
Option Explicit

Private Sub Worksheet_Deactivate()
On Error Resume Next
Unload DOP_ot
Unload DOP_spr
Unload frm_Mnn
DoEvents
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Unload DOP_ot
Unload DOP_sv
Unload DOP_spr
Unload frm_Mnn
End Sub