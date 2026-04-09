VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_psv 
   Caption         =   "Новый поставщик"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   OleObjectBlob   =   "add_psv.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_psv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OK_Click()
        If Trim(tb_psv.Text) = "" Then
            MsgBox "Введите данные поставщика!", vbInformation, "Данные"
            tb_psv.SetFocus
            Exit Sub
        End If

        If save_psv_to_spr(tb_psv.Text) Then
            Call refresh_vvod_forms_sources
            Unload Me
        End If
End Sub



Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
Me.StartUpPosition = 0
Me.Top = vvodPr.Top
Me.Left = vvodPr.Left
Exit Sub

ErrHandler:
ReportVbaError "add_psv.UserForm_Initialize", Err.Number, Err.Description, "Поставщик"
End Sub
Private Sub NO_Click()
Unload Me
End Sub
