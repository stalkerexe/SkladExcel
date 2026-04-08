VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_zkz 
   Caption         =   "Новый заказчик"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   OleObjectBlob   =   "add_zkz.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_zkz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub OK_Click()
On Error Resume Next
If TextBox1.text = "" Then
MsgBox "Введите данные заказчика!", 64, "Данные"
TextBox1.SetFocus
Exit Sub
End If
frm_msg.Show
End Sub

Private Sub add_zkzz()
Call dann
Call do_add
End Sub

Private Sub dann()
On Error Resume Next
sZkz = TextBox1.text
sAdr = TextBox2.text
sTlf = TextBox3.text
sMail = tb_mail.text
End Sub

Private Sub do_add()
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
Me.Top = vvodZv.Top
Me.Left = vvodZv.Left
End Sub

Private Sub NO_Click()
Unload Me
End Sub
