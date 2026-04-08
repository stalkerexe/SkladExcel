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
On Error GoTo ErrHandler
If Trim(TextBox1.Text) = "" Then
    MsgBox "Введите данные заказчика!", vbInformation, "Данные"
    TextBox1.SetFocus
    Exit Sub
End If

Call add_zkzz
Exit Sub

ErrHandler:
ReportVbaError "add_zkz.OK_Click", Err.Number, Err.Description, "Заказчик"
End Sub

Private Sub add_zkzz()
On Error GoTo ErrHandler
Call dann
Call do_add
Exit Sub

ErrHandler:
ReportVbaError "add_zkz.add_zkzz", Err.Number, Err.Description, "Заказчик"
End Sub

Private Sub dann()
On Error GoTo ErrHandler
sZkz = TextBox1.Text
sAdr = TextBox2.Text
sTlf = TextBox3.Text
sMail = tb_mail.Text
Exit Sub

ErrHandler:
ReportVbaError "add_zkz.dann", Err.Number, Err.Description, "Заказчик"
End Sub

Private Sub do_add()
On Error GoTo ErrHandler

    If save_zkz_to_spr(sZkz, sAdr, sTlf, sMail) Then
        Call refresh_vvod_forms_sources
        Unload Me
    End If

Exit Sub
ErrHandler:
ReportVbaError "add_zkz.do_add", Err.Number, Err.Description, "Заказчик"
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
Me.StartUpPosition = 0
Me.Top = vvodZv.Top
Me.Left = vvodZv.Left
Exit Sub

ErrHandler:
ReportVbaError "add_zkz.UserForm_Initialize", Err.Number, Err.Description, "Заказчик"
End Sub

Private Sub NO_Click()
Unload Me
End Sub
