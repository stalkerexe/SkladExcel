Attribute VB_Name = "ф_файлы_2"
Option Explicit


Public Sub clearpr3()
On Error Resume Next
Call unload_mn_vid_pr
DoEvents
If MsgBox("Удалить все позиции из накладной?", vbOKCancel + vbQuestion, "Очистить") = vbCancel Then Exit Sub
With ThisWorkbook.Sheets("Приход")
r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
.Range(.Cells(rwZv, 2), .Cells(r7 + 44, 2)).EntireRow.Delete
.Cells(rwzvSm, prSm) = ""
.Range("a1") = ""
.Cells(1, prComm) = ""
End With

        Call режим_редактирования_off_pr("Приход")

End Sub


Public Sub clear_box()
With ThisWorkbook.Sheets("корзина")
r7 = .UsedRange.Rows.Count + .UsedRange.Row - 1
.Range(.Cells(rwZv, 2), .Cells(r7 + 44, 2)).EntireRow.Delete
.Cells(rwzvSm, zvSm) = ""
End With
With ThisWorkbook.Sheets("Склад")
.Cells(3, iBox1) = 0
.Cells(3, iBox2) = 0
End With
End Sub

Public Sub sh_Praise()
On Error Resume Next
Call unload_mn_vid
DoEvents
Praise.Show
End Sub

Public Sub sh_vvodZv()
On Error Resume Next
Call unload_mn_vid: DoEvents
vvodZv.Show
End Sub

Public Sub sh_vvodPr()
On Error Resume Next
Call unload_mn_vid_pr: DoEvents
vvodPr.Show
End Sub

Function Replace_symbols(ByVal txt As String) As String
On Error Resume Next
Dim st As String
st = "~:<>!@/\#$%^&*=|`"""
For i = 1 To Len(st)
txt = Replace(txt, VBA.Mid(st, i, 1), "_")
Next
Replace_symbols = txt
End Function

Public Sub sh_fSet()
        On Error Resume Next
        Call do_unload_dop
        frm_Set.Show
End Sub

Public Sub sh_fSet_zk()
frm_Set_zk.Show
End Sub


Public Sub s_pr()
Sheets("Приход").Select
End Sub
Public Sub s_zv()
Sheets("Расход").Select
End Sub
Public Sub s_sk()
Sheets("Склад").Select
End Sub
Public Sub s_gl()
Sheets("Главная").Select
End Sub
Public Sub s_ot_rs()
Sheets("Отложено_расход").Select
End Sub
Public Sub s_ot_pr()
Sheets("Отложено_приход").Select
End Sub

Public Sub sh_frm_sk()
frm_sk.Show
End Sub


Public Sub sh_DOP_arh()
        On Error Resume Next
        Call do_unload_dop
        zvSelect.Show
End Sub

Public Sub sh_DOP_ot()
On Error Resume Next
Call do_unload_dop
DOP_ot.Show
End Sub

Public Sub sh_DOP_spr()
On Error Resume Next
Call do_unload_dop
DOP_spr.Show
End Sub

Public Sub sh_DOP_sv()
On Error Resume Next
Call do_unload_dop
DOP_sv.Show
End Sub

Public Sub do_unload_dop()
On Error Resume Next
Unload DOP_ot
Unload DOP_sv
Unload DOP_spr
DoEvents
End Sub

Public Sub sh_Form_sklads()
        On Error Resume Next
        Call do_unload_dop
        Form_sklads.Show
End Sub

Public Sub sh_frm_Show()
On Error Resume Next
Dim a As Double
Dim b As Double
With frm_Show
.Show
.StartUpPosition = 0
.Left = Application.Width
.Top = ActiveSheet.Shapes("grCmbBox").Top + ActiveSheet.Shapes("grCmbBox").Height + 20
Call добавить_контролы
a = Application.Width
b = ActiveSheet.Shapes("grCmbBox").Left
For i = a To b Step -30
.Left = i
Next
End With
End Sub

Public Sub sh_frm_Find()
On Error Resume Next
frm_Find.Show
End Sub

Public Sub sh_frm_Find_zk()
On Error Resume Next
frm_Find_zk.Show
End Sub

Public Sub sh_Form_()
Form_.Show
End Sub

Public Sub sh_frm_Find_pr()
On Error Resume Next
frm_Find_pr.Show
End Sub

Public Sub remove_green()
On Error Resume Next
Application.ErrorCheckingOptions.BackgroundChecking = False
End Sub

Public Sub sh_frmHelp()
frmHelp.Show
End Sub


