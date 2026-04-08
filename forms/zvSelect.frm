VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zvSelect 
   Caption         =   " Накладные"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865.001
   OleObjectBlob   =   "zvSelect.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zvSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sData As String
Dim sCombData As String

Private Const iCm As Integer = 5
Dim a()
Dim b()




Private Sub comb_vid_Click()
        On Error Resume Next

        Call load_arh
        
        If ListBox1.ListCount > 0 Then
            comb_dt.Enabled = True
        Else
            comb_dt.Enabled = False
        End If
        
        If iVid = "Возврат" Then
                Me.lb_mj.Caption = "Накладная"
            Else
                Me.lb_mj.Caption = "Сотрудник"
        End If
        
        ListBox1.SetFocus
        
End Sub


Private Sub comb_dt_Click()
        Call parse_nk
End Sub

Private Sub comb_zkz_Click()
        Call parse_nk
End Sub

Private Sub comb_mj_Click()
        Call parse_nk
End Sub


Private Sub parse_nk()
        On Error Resume Next
        
        ListBox1.clear
        
        sMj = Me.comb_Mj.Value
        sZkz = Me.comb_zkz.Value
        sCombData = Me.comb_dt.Value
        
If (Not Not c) = 0 Then Exit Sub
        
        ReDim b(LBound(c) To UBound(c), 1 To iCm)
        
        j = 1
        For i = LBound(c) To UBound(c)
        
            If sCombData = "Все" Then GoTo 11
            If sCombData = "" Then GoTo 11
            
            If sCombData = "Сегодня" Then sData = VBA.Format(VBA.Date, "dd.mm.yyyy")
            If sCombData = "Вчера" Then sData = VBA.Format(VBA.Date - 1, "dd.mm.yyyy")
            
            If c(i, 4) = sData Then
11
            If sZkz = "" Then GoTo 22
            If sZkz = "Все" Then GoTo 22
            If c(i, 3) = sZkz Then
22
            If sMj = "" Then GoTo 33
            If sMj = "Все" Then GoTo 33
            If c(i, 5) = sMj Then
            
33

                    For cm = 1 To iCm
                        b(j, cm) = c(i, cm)
                    Next
                    j = j + 1
                
            End If
            End If
            End If
            
        Next
        
        ListBox1.List = b
        
End Sub







Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        Call load_nk_from_arh
End Sub





Private Sub UserForm_Initialize()
        Call load_all
End Sub

Private Sub load_all()
        On Error Resume Next
        
        Call combo
        Call forma
        
        comb_vid.ListIndex = 0
        
        Call load_arh
        
        Call load_mj
        
        If ListBox1.ListCount > 0 Then
            comb_dt.Enabled = True
        Else
            comb_dt.Enabled = False
        End If
        
End Sub

Private Sub forma()
        On Error Resume Next
        ListBox1.ColumnCount = iCm
        ListBox1.ColumnWidths = "0;44;190;85;94"
Me.Height = Application.Height / 3 * 2
        ListBox1.Height = Me.Height - 70
End Sub



Private Sub load_arh()
        Call doScreenOff
        Call do_load_arh
        Call doScreenOn
        
        If ListBox1.ListCount = 0 Then
            With Frame_msg
                .SpecialEffect = 0
                .Visible = True
                .Left = 96
                .Top = 80
                .ZOrder 0
            End With
        Else
            Frame_msg.Visible = False
        End If
        
End Sub

Private Sub do_load_arh()
        On Error Resume Next

        ListBox1.clear
        Erase c

        iVid = Me.comb_vid.Value
        Call find_path_vid

        Call arr_arh_all
        
        Call parse_arh
        
        Call load_mj
77
        If ListBox1.ListCount > 0 Then
            comb_zkz.Enabled = True
            comb_Mj.Enabled = True
        Else
            comb_zkz.Enabled = False
            comb_Mj.Enabled = False
        End If
        
End Sub

Private Sub parse_arh()
        On Error Resume Next
        
        If iCol = 0 Then Exit Sub
        
        ReDim c(1 To iCol + 2, 1 To iCm)
        
        n = 1
        
        For i = LBound(mk) To UBound(mk)
            If mk(i, 1) <> "" Then
            c(n, 1) = mk(i, 1)
            c(n, 2) = Format(nom(i, 1), "00000")
c(n, 3) = zkz(i, 1)
            c(n, 4) = Format(dt(i, 1), "dd.mm.yyyy")
            c(n, 5) = mj(i, 1): If iVid = "Возврат" Then c(n, 5) = doc(i, 1)
            n = n + 1
            End If
        Next
        
        Me.ListBox1.List = c
End Sub


Private Sub combo()
On Error Resume Next

With comb_vid
.AddItem "Приход"
.AddItem "Отгрузка"
End With

With comb_year
.Value = VBA.Year(VBA.Now)
.Enabled = False
End With

With comb_dt
.Left = lb_dt.Left
.Top = lb_dt.Top
.Width = lb_dt.Width + 13
.ZOrder 1
.AddItem "Вчера"
.AddItem "Сегодня"
.AddItem "Все"
End With

With comb_Mj
.Left = lb_mj.Left
.Top = lb_mj.Top
.Width = lb_mj.Width + 13
.ZOrder 1
End With

With comb_zkz
.Left = lb_zkz.Left
.Top = lb_zkz.Top
.Width = lb_zkz.Width + 13
.ZOrder 1
End With
End Sub


Private Sub load_mj()
        On Error Resume Next
        Call RemoveDuplicates
        
        With ThisWorkbook.Sheets("буфер")
            r7 = .Cells(Rows.Count, "c").End(xlUp).Row + 1
            a = .Range("c1:c" & r7).Value
            comb_zkz.List = a
            i = comb_zkz.ListCount - 1
            comb_zkz.List(i, 0) = "Все"
            r7 = .Cells(Rows.Count, "e").End(xlUp).Row + 1
            a = .Range("e1:e" & r7).Value
            comb_Mj.List = a
            i = comb_Mj.ListCount - 1
            comb_Mj.List(i, 0) = "Все"
        End With
        
        Call clearBf
End Sub

Private Sub RemoveDuplicates()
        On Error Resume Next

If (Not Not c) = 0 Then Exit Sub

        Call clearBf

        With ThisWorkbook.Sheets("буфер")
            .Cells(1, "a").Resize(UBound(c), 5) = c
            r7 = .Cells(Rows.Count, "a").End(xlUp).Row + 1
            .Range("c1:c" & r7).RemoveDuplicates Columns:=1, Header:=xlNo
            .Range("e1:e" & r7).RemoveDuplicates Columns:=1, Header:=xlNo
        End With
        
Call sort_do(3, 3)
Call sort_do(5, 5)
        
End Sub



Private Sub lb_mj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If ListBox1.ListCount > 0 Then comb_Mj.DropDown
End Sub

Private Sub lb_zkz_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If ListBox1.ListCount > 0 Then comb_zkz.DropDown
End Sub

Private Sub lb_dt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        On Error Resume Next
        If comb_year.Value <> CStr(VBA.Year(VBA.Now)) Then Exit Sub
        
        If comb_year.Value = CStr(VBA.Year(VBA.Now)) Then
            If ListBox1.ListCount > 0 Then
                comb_dt.DropDown
            End If
        End If
End Sub



Private Sub UserForm_Terminate()
        On Error Resume Next
        Call clearBf
        Erase dt: Erase c: Erase nom: Erase b: Erase a
End Sub

Private Sub NO_Click()
        Unload Me
End Sub
