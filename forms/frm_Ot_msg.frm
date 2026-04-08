VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Ot_msg 
   Caption         =   " Создать отчет"
   ClientHeight    =   1830
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6360
   OleObjectBlob   =   "frm_Ot_msg.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Ot_msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub NO_Click()
Unload Me
End Sub

Private Sub OK_Click()
Call otchet_do
Unload Me
End Sub
