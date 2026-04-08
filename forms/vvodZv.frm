VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vvodZv 
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795.001
   OleObjectBlob   =   "vvodZv.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vvodZv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tb_zkz_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        comb_zkz.DropDown
End Sub



Private Sub tb_Zkz_Change()
        Call do_find
End Sub

Private Sub do_find()
    On Error Resume Next

        If flag = 1 Then Exit Sub
        
        comb_zkz.SetFocus
        Call find_zkz
        
        If Val(comb_find.ListCount) > 0 Then comb_find.DropDown
        
        If tb_Zkz.Text = "" Then
            If Val(comb_zkz.ListCount) > 0 Then
                comb_zkz.DropDown
                Exit Sub
            End If
        End If

End Sub


Private Sub find_zkz()
    On Error Resume Next
    
    str_ = tb_Zkz.Text
    
    If VBA.Len(str_) = 0 Then
        comb_find.clear
        comb_find.SetFocus
        tb_Zkz.SetFocus
        Exit Sub
    End If
   
    iCol = 0
    
c = zkz
    
    ReDim cc(LBound(c) To UBound(c), 1 To 2)
    
    textS = UCase(str_)
    
    For i = LBound(c) To UBound(c)
        sNm = c(i, 1)
        
        If Len(str_) = 1 Then
            If UCase(VBA.Left(sNm, 1)) = textS Then
                cc(i, 1) = i
                cc(i, 2) = c(i, 1)
                iCol = iCol + 1
            End If
        Else
            If InStr(1, UCase(sNm), textS) > 0 Then
                cc(i, 1) = i
                cc(i, 2) = c(i, 1)
                iCol = iCol + 1
            End If
        End If
    Next
    
    If iCol = 0 Then
        comb_find.clear
        comb_find.SetFocus
        tb_Zkz.SetFocus
        DoEvents
        Exit Sub
    End If

    ReDim w(1 To iCol, 1 To 2)
    j = 1
    For i = LBound(cc) To UBound(cc)
    If cc(i, 1) <> "" Then
    w(j, 1) = cc(i, 1)
    w(j, 2) = cc(i, 2)
    j = j + 1
    End If
    Next
    
    comb_find.clear
    comb_find.SetFocus
    tb_Zkz.SetFocus
    DoEvents
    comb_find.List = w
    
End Sub


Private Sub comb_find_Click()
    On Error Resume Next
        
        flag = 1
        
        With comb_find
            i = .ListIndex
            If i = -1 Then GoTo 99
            sZkz = .List(i, 1)
            sAdr = CStr(adr(.List(i, 0), 1))
            sTlf = CStr(tlf(.List(i, 0), 1))
        End With
        
        tb_Zkz.Text = sZkz
        tb_adr.Text = sAdr
        tb_tlf.Text = sTlf
        
        comb_zkz.ListIndex = -1
        
        flag = 0
99
End Sub


Private Sub comb_zkz_Click()
        flag = 1
        

        ind = comb_zkz.ListIndex
        If ind = -1 Then Exit Sub
        
        With comb_zkz
            tb_Zkz.Text = CStr(.List(ind, 0))
            tb_adr.Text = CStr(adr(ind + 1, 1))
            tb_tlf.Text = CStr(tlf(ind + 1, 1))
        End With
        
        
        flag = 0
End Sub



Private Sub comb_mj_Click()
        On Error Resume Next
        tb_mj.Text = comb_Mj.Value
End Sub



Private Sub OK_Click()
        On Error Resume Next
        ThisWorkbook.Activate
        Sheets("Расход").Select
        
        Cells(rwZv_zkz, 4).Value = Me.tb_Zkz.Text
        Cells(rwZv_adr, 4).Value = Me.tb_adr.Text
        Cells(rwZv_tlf, 4).Value = Me.tb_tlf.Text
        Cells(rwZv_mj, 4).Value = Me.tb_mj.Text
        Cells(rwZv_dt, 4).Value = Me.tb_dt1.Text
        Cells(rwZv_dt2, 4).Value = Me.tb_dt2.Text
        
        Unload Me
End Sub



Private Sub UserForm_Initialize()
        flag = 1
        Call doScreenOff
        Call load_dann
        Call doScreenOn
        flag = 0
End Sub

Private Sub load_dann()
    On Error Resume Next

        With ThisWorkbook.Sheets("Расход")
            tb_Zkz.Text = .Cells(rwZv_zkz, 4).Value
            tb_adr.Text = .Cells(rwZv_adr, 4).Value
            tb_tlf.Text = .Cells(rwZv_tlf, 4).Value
            tb_mj.Text = .Cells(rwZv_mj, 4).Value
            tb_dt1.Text = .Cells(rwZv_dt, 4).Value
            tb_dt2.Text = .Cells(rwZv_dt2, 4).Value
            DoEvents
        End With
        
        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("cmb_d")
            Me.Top = .Top + .Height + 20
            Me.Left = .Left
        End With
        
        Call combo
        Call load_spr
        
        OK.BackColor = RGB(58, 110, 165)
        OK.ForeColor = RGB(255, 255, 255)
        NO.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub combo()
        On Error Resume Next
        
        With comb_zkz
        .Left = tb_Zkz.Left
        .Top = tb_Zkz.Top
        .Width = tb_Zkz.Width
        .ZOrder 1
        End With
        
        With comb_find
        .Left = tb_Zkz.Left
        .Top = tb_Zkz.Top
        .Width = tb_Zkz.Width
        .ColumnCount = 2
        .ColumnWidths = 0
        .ZOrder 1
        End With
        
        With comb_Mj
        .Left = tb_mj.Left
        .Top = tb_mj.Top
        .Width = tb_mj.Width
        .ZOrder 1
        End With
End Sub

Private Sub load_spr()
        On Error Resume Next
        
        Call load_mjj_all
        comb_Mj.List = mj
        
        Call load_zkz
        
End Sub

Private Sub load_zkz()
On Error Resume Next
ReDim zkz(1 To 5, 1 To 1)
ReDim adr(1 To 5, 1 To 1)
ReDim tlf(1 To 5, 1 To 1)

zkz(1, 1) = "ООО «ЕФТ ГРУПП»"
adr(1, 1) = "г.Москва, ул.Российская 17"
tlf(1, 1) = "890473748"

zkz(2, 1) = "ИП СеверМет г.Уфа"
adr(2, 1) = "г.Уфа, ул.Гагарина 26"
tlf(2, 1) = "890473748"

zkz(3, 1) = "ООО ТК «ТЕХНОРЕСУРС»"
adr(3, 1) = "г.Москва, ул.Лесная 57"
tlf(3, 1) = "890473748"

zkz(4, 1) = "ООО ГК «АЛЬФА-СПК-ДЖИТЕЙЧ»"
adr(4, 1) = "г.Оренбург, ул.Ухтомского 12"
tlf(4, 1) = "890473748"

zkz(5, 1) = "ИП Левникова Ю.П."
adr(5, 1) = "г.Пермь, ул.Северная 45"
tlf(5, 1) = "890473748"

comb_zkz.List = zkz

End Sub


Private Sub UserForm_Terminate()
flag = 1
tb_Zkz.Text = ""
tb_mj.Text = ""
tb_dt1.Text = ""
tb_dt2.Text = ""
flag = 0
End Sub

Private Sub tb_mj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
comb_Mj.DropDown
End Sub
Private Sub tb_Zkz_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
comb_zkz.DropDown
End Sub
Private Sub tb_dt1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
iTb = 0
Set iForm = Me
Call ShowForm4
End Sub
Private Sub tb_dt2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
iTb = 1
Set iForm = Me
Call ShowForm4
End Sub
Private Sub ico_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
ico.BackColor = &H80000005
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
ico.BackColor = &H8000000F
End Sub

Private Sub ico_Click()
add_zkz.Show
End Sub

Private Sub NO_Click()
Unload Me
End Sub

