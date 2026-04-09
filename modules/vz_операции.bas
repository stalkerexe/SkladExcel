Attribute VB_Name = "vz_операции"
Option Explicit

Public Sub zvk_open_archive()
On Error GoTo ErrHandler
zvSelect.Show
Exit Sub
ErrHandler:
ReportVbaError "zvk_open_archive", Err.Number, Err.Description, "Архив"
End Sub

Public Sub zvk_commit_return(ByVal frm As Object)
On Error GoTo ErrHandler

If frm Is Nothing Then Exit Sub

Call dann_zvk

If Trim(CStr(marker)) = "" Then marker = "vz_" & Format(Now, "yyyymmdd_hhnnss")
If Trim(CStr(nomer)) = "" Then nomer = 0

iVid = "vz"
Call save_nk

MsgBox "Возврат сохранен в архив.", vbInformation, "Возврат"
Unload frm
Exit Sub

ErrHandler:
ReportVbaError "zvk_commit_return", Err.Number, Err.Description, "Возврат"
End Sub

Public Sub zvk_cancel_return(ByVal frm As Object)
On Error GoTo ErrHandler

If frm Is Nothing Then Exit Sub

If MsgBox("Отменить изменения по возврату?", vbYesNo + vbQuestion, "Возврат") = vbNo Then Exit Sub

Unload frm
Exit Sub

ErrHandler:
ReportVbaError "zvk_cancel_return", Err.Number, Err.Description, "Возврат"
End Sub
