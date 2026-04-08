VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_sklads
   Caption         =   "Склады"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   OleObjectBlob   =   "Form_sklads.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_sklads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Call open_sklad
End Sub
Private Sub CommandButton2_Click()
Call добавить_склад
End Sub
Private Sub CommandButton3_Click()
Call rename_sk
End Sub
Private Sub CommandButton4_Click()
Call delete_sk
End Sub

Private Sub CommandButton5_Click()
Call open_sklad
End Sub


Private Sub CommandButton7_Click()
Call добавить_склад
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
If ListBox1.ListIndex = -1 Then Exit Sub
Application.CommandBars("MyContextMenu_ListBox").Delete
With Application.CommandBars.Add(Name:="MyContextMenu_ListBox", Position:=msoBarPopup, Temporary:=True)
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 9255
.Caption = "Открыть склад"
.OnAction = "open_sklad"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 137
.Caption = "Добавить склад"
.OnAction = "добавить_склад"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 162
.Caption = "Переименовать склад"
.OnAction = "rename_sk"
End With
With .Controls.Add(Type:=msoControlButton)
.Style = msoButtonIconAndCaption
.FaceId = 923
.Caption = "Удалить склад"
.OnAction = "delete_sk"
End With
End With
CommandBars("MyContextMenu_ListBox").ShowPopup
End Sub
Private Sub NO1_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Call load
End Sub
Private Sub load()
On Error Resume Next
Call load_sk
For i = 0 To dic_sk.Count - 1
ListBox1.AddItem dic_sk.Item(i)
Next
End Sub
