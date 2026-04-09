Attribute VB_Name = "я_экран"
Option Explicit


Public Sub ScreenOff()
With Application
.ScreenUpdating = False
.EnableEvents = False
End With
End Sub
Public Sub ScreenOn()
With Application
.ScreenUpdating = True
.EnableEvents = True
End With
End Sub

Public Sub doScreenOff()
Waite.Show: DoEvents
Call ScreenOff: Application.DisplayAlerts = False
End Sub
Public Sub doScreenOn()
Call ScreenOn: Application.DisplayAlerts = True
Unload Waite
End Sub


Public Sub polnekr()
On Error Resume Next
Application.DisplayFormulaBar = False
ActiveWindow.DisplayHeadings = False
Application.DisplayFullScreen = True
End Sub
Public Sub verekr()
On Error Resume Next
Application.DisplayFullScreen = False
Application.DisplayFormulaBar = True
ActiveWindow.DisplayHeadings = True
End Sub
Public Sub ekr()
On Error GoTo ErrHandler

Dim skladSheet As Worksheet

If Application.DisplayFullScreen = False Then
    Call polnekr
Else
    Call verekr
End If

If Not RequireSheet(SHEET_SKLAD, skladSheet, "ekr") Then Exit Sub

skladSheet.Select
Exit Sub

ErrHandler:
ReportVbaError "ekr", Err.Number, Err.Description
End Sub
Public Sub AutoFilter_delete()
On Error GoTo ErrHandler

With ActiveSheet
    If .AutoFilterMode = True Then
        .Cells.AutoFilter
    End If
End With

Exit Sub
ErrHandler:
ReportVbaError "AutoFilter_delete", Err.Number, Err.Description
End Sub
Public Sub msg_demo()
MsgBox "Функция недоступна в текущем контексте. Проверьте права пользователя или настройки документа.", vbExclamation, "Ограничение доступа"
End Sub

Public Function nom_nk(clmn As Integer)
On Error GoTo ErrHandler

Dim ws As Worksheet
If Not RequireSheet("nummm", ws, "nom_nk") Then Exit Function

With ws
    .Cells(2, clmn) = .Cells(2, clmn) + 1
    nomer = .Cells(2, clmn)
End With

Exit Function
ErrHandler:
ReportVbaError "nom_nk", Err.Number, Err.Description
End Function
Public Sub clearBf()
On Error GoTo ErrHandler

Dim ws As Worksheet
If Not RequireSheet("буфер", ws, "clearBf") Then Exit Sub

With ws
    .Cells.ClearContents
End With

Exit Sub
ErrHandler:
ReportVbaError "clearBf", Err.Number, Err.Description
End Sub
