VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ 
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "Form_.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_nm.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_nm.ForeColor = RGB(0, 0, 255)
lb_cn.ForeColor = RGB(0, 0, 0)
lb_me.ForeColor = RGB(0, 0, 0)
End Sub


Private Sub Frame_me_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_cn.ForeColor = RGB(0, 0, 0)
lb_me.ForeColor = RGB(0, 0, 0)
End Sub
Private Sub Frame_cn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_cn.ForeColor = RGB(0, 0, 0)
lb_me.ForeColor = RGB(0, 0, 0)
End Sub



Private Sub lb_nm_Click()
On Error Resume Next
Waite.Show: Waite.Label1.Caption = "Интернет соединение...": DoEvents
ThisWorkbook.FollowHyperlink "https://sklad-excel.ru/uchet-prodaj/"
If Err Then Unload Waite: DoEvents: MsgBox "Интернет не подключен!", 48, " Интернет-Подключение"
Unload Waite
End Sub

Private Sub lb_nm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_nm.ForeColor = &HC000&
End Sub


Private Sub lb_cn_Click()
On Error Resume Next
Waite.Show: Waite.Label1.Caption = "Интернет соединение...": DoEvents
ThisWorkbook.FollowHyperlink "https://sklad-excel.ru/price/"
If Err Then Unload Waite: DoEvents: MsgBox "Интернет не подключен!", 48, " Интернет-Подключение"
Unload Waite
End Sub
Private Sub lb_cn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_cn.ForeColor = &HC000&
End Sub


Private Sub lb_me_Click()
On Error Resume Next
Waite.Show: Waite.Label1.Caption = "Интернет соединение...": DoEvents
ThisWorkbook.FollowHyperlink "https://sklad-excel.ru/me/"
If Err Then Unload Waite: DoEvents: MsgBox "Интернет не подключен!", 48, " Интернет-Подключение"
Unload Waite
End Sub

Private Sub lb_me_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
lb_me.ForeColor = &HC000&
End Sub


