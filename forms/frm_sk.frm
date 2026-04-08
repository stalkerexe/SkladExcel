VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_sk 
   Caption         =   "Склады"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5490
   OleObjectBlob   =   "frm_sk.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_sk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub OK_Click()

    Call check_
    If iCol = 0 Then MsgBox "Выберите позиции!", 64, "Склад": Exit Sub
    
    Call do_filter
    
    Unload Me
    
End Sub

Private Sub do_filter()

    On Error Resume Next

    ReDim sk(1 To 3, 1 To 1)

    j = 1

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            sk(j, 1) = ListBox1.List(i, 0)
            j = j + 1
        End If
    Next

    Call sklad_show
    
    Erase c
    
End Sub

Private Sub CheckBox1_Click()
        On Error Resume Next

    With ListBox1
        For i = 0 To .ListCount - 1
            If CheckBox1.Value = True Then
                .Selected(i) = True
            Else
                .Selected(i) = False
            End If
        Next
    End With
End Sub

Private Sub check_()
        On Error Resume Next

    iCol = 0
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            iCol = iCol + 1
        End If
    Next
End Sub




Private Sub UserForm_Initialize()
        On Error Resume Next
        
        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("cmb_sk")
        Me.Top = .Top + .Height + 20
        Me.Left = .Left
        End With

        ListBox1.ListStyle = fmListStyleOption
        ListBox1.MultiSelect = fmMultiSelectMulti

        Call load_sklads
End Sub

Private Sub load_sklads()
On Error Resume Next
ListBox1.AddItem "Материалы"
ListBox1.AddItem "Металлопрокат"
ListBox1.AddItem "Спецодежда"
End Sub


Private Sub NO_Click()
    Unload Me
End Sub

