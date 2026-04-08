VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Oplata 
   Caption         =   "Способ оплаты"
   ClientHeight    =   3135
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3540
   OleObjectBlob   =   "frm_Oplata.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Oplata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub OK_Click()

        Call check_
        If iCol = 0 Then MsgBox "Выберите способ оплаты!!", 64, "Способ оплаты": Exit Sub

        Call do_ok

        Unload Me

End Sub

Private Sub do_ok()

        ThisWorkbook.Activate
        Sheets("Расход").Select
        
        iOpl = ListBox1.Value
        
        Cells(rwZv_mj, zvSm).Value = iOpl

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
        Call doScreenOff
        Call load_all
        Call doScreenOn
End Sub

Private Sub load_all()

        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("cmb_oplata")
            Me.Top = .Top + .Height + 20
            Me.Left = .Left
        End With
        
        Call load_spr

        Me.ListBox1.ListStyle = fmListStyleOption
        Me.ListBox1.MultiSelect = fmMultiSelectSingle
    
End Sub

Private Sub load_spr()
        On Error Resume Next
        ListBox1.AddItem "Наличный"
        ListBox1.AddItem "Безналичный"
        ListBox1.AddItem "Картой"
        ListBox1.AddItem "Перевод"
End Sub

Private Sub NO_Click()
        Unload Me
End Sub
