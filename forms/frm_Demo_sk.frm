VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Demo_sk 
   Caption         =   "Склад"
   ClientHeight    =   9660.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9345.001
   OleObjectBlob   =   "frm_Demo_sk.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Demo_sk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub OK_Click()
On Error Resume Next
Unload Me
DoEvents
Form_.Show
Form_.MultiPage1.value = 1
End Sub

Private Sub NO_Click()
Unload Me
End Sub

