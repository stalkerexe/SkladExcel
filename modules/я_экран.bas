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
Application.DisplayFullScreen = False
Application.DisplayFormulaBar = True
ActiveWindow.DisplayHeadings = True
End Sub
Public Sub ekr()
On Error Resume Next
If Application.DisplayFullScreen = False Then
Call polnekr
Else
Call verekr
End If
Sheets("Cклад").Select
End Sub
Public Sub AutoFilter_delete()
On Error Resume Next
With ActiveSheet
If .AutoFilterMode = True Then
.Cells.AutoFilter
End If
End With
End Sub
Public Sub msg_demo()
MsgBox "Данная функция доступна только для полной версии программы!", 64, "Демо-версия"
End Sub

Public Function nom_nk(clmn As Integer)
On Error Resume Next
With ThisWorkbook.Sheets("nummm")
.Cells(2, clmn) = .Cells(2, clmn) + 1
nomer = .Cells(2, clmn)
End With
End Function
Public Sub clearBf()
On Error Resume Next
With ThisWorkbook.Sheets("буфер")
.Cells.ClearContents
End With
End Sub



