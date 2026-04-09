VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "Подсказка"
   ClientHeight    =   12810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   OleObjectBlob   =   "frmHelp.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub lb_link1_Click()
On Error Resume Next
Waite.Show
Waite.Label1.Caption = "Интернет соединение..."
DoEvents
ThisWorkbook.FollowHyperlink "https://sklad-excel.ru/help_mgp_4/"
If Err Then Unload Waite: DoEvents: MsgBox "Интернет не подключен!", 48, " Интернет-Подключение"
Unload Waite
End Sub
Private Sub lb_link1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_link1.ForeColor = &HC000&
End Sub
Private Sub tb_1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_link1.ForeColor = RGB(0, 0, 0)
End Sub




Private Sub lb_link2_Click()
On Error Resume Next
Waite.Show
Waite.Label1.Caption = "Интернет соединение..."
DoEvents
ThisWorkbook.FollowHyperlink "https://youtu.be/HvG0fFgNnpU"
If Err Then Unload Waite: DoEvents: MsgBox "Интернет не подключен!", 48, " Интернет-Подключение"
Unload Waite
End Sub
Private Sub lb_link2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_link2.ForeColor = &HC000&
End Sub
Private Sub tb_6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_link2.ForeColor = RGB(0, 0, 0)
End Sub




Private Sub UserForm_Initialize()
Me.StartUpPosition = 0
Me.Top = Application.Top + 15
Me.Left = Application.Width - Me.Width - 15

For i = 1 To 6
Me.Controls("tb_" & i).Locked = True
Next

If ActiveSheet.Name = "Главная" Then Me.MultiPage1.value = 0
If ActiveSheet.Name = "Расход" Then Me.MultiPage1.value = 1
If ActiveSheet.Name = "Отложено_расход" Then Me.MultiPage1.value = 2
If ActiveSheet.Name = "Приход" Then Me.MultiPage1.value = 3
If ActiveSheet.Name = "Отложено_приход" Then Me.MultiPage1.value = 4
If ActiveSheet.Name = "Склад" Then Me.MultiPage1.value = 5
End Sub

Private Sub MultiPage1_Change()
With Me.MultiPage1
If .value = 0 Then Sheets("Главная").Select
If .value = 1 Then Sheets("Расход").Select
If .value = 2 Then Sheets("Отложено_расход").Select
If .value = 3 Then Sheets("Приход").Select
If .value = 4 Then Sheets("Отложено_приход").Select
If .value = 5 Then Sheets("Склад").Select
End With
End Sub


Private Sub ico_1_Click()
Unload Me
End Sub
Private Sub ico_2_Click()
Unload Me
End Sub
Private Sub ico_3_Click()
Unload Me
End Sub
Private Sub ico_4_Click()
Unload Me
End Sub
Private Sub ico_5_Click()
Unload Me
End Sub
Private Sub ico_6_Click()
Unload Me
End Sub
