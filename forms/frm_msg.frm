VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_msg 
   Caption         =   "Демо-версия"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6345
   OleObjectBlob   =   "frm_msg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub NO_Click()
Unload Me
End Sub

Private Sub OK_Click()
On Error Resume Next
Unload Me
DoEvents
Form_.Show
Form_.MultiPage1.Value = 1
End Sub
