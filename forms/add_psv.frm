VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_psv 
   Caption         =   "Новый поставщик"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   OleObjectBlob   =   "add_psv.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_psv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OK_Click()
        If tb_psv.Text = "" Then
        MsgBox "Введите данные поставщика!", 64, "Данные"
        tb_psv.SetFocus
        Exit Sub
        End If
        frm_msg.Show
End Sub



Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0
Me.Top = vvodPr.Top
Me.Left = vvodPr.Left
End Sub
Private Sub NO_Click()
Unload Me
End Sub

