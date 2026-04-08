VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vvodPr 
   Caption         =   "Приход"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   OleObjectBlob   =   "vvodPr.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vvodPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub tb_psv_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            comb_psv.DropDown
End Sub



Private Sub tb_psv_Change()
        Call do_find
End Sub

Private Sub do_find()
        On Error Resume Next

        If flag = 1 Then Exit Sub
        
        comb_psv.SetFocus
        Call find_psv
        
        If Val(comb_find.ListCount) > 0 Then comb_find.DropDown
        
        If tb_psv.Text = "" Then
            If Val(comb_psv.ListCount) > 0 Then
                comb_psv.DropDown
                Exit Sub
            End If
        End If

End Sub


Private Sub find_psv()
    On Error Resume Next
    
    str_ = tb_psv.Text
    
    If VBA.Len(str_) = 0 Then
        comb_find.clear
        comb_find.SetFocus
        tb_psv.SetFocus
        Exit Sub
    End If
   
    iCol = 0
    
    c = comb_psv.List
    
    ReDim cc(LBound(c) To UBound(c), 1 To 1)
    
    textS = UCase(str_)
    
    For i = LBound(c) To UBound(c)
        sNm = c(i, 0)
        
        If Len(str_) = 1 Then
            If UCase(VBA.Left(sNm, 1)) = textS Then
                cc(i, 1) = c(i, 0)
                iCol = iCol + 1
            End If
        Else
            If InStr(1, UCase(sNm), textS) > 0 Then
                cc(i, 1) = c(i, 0)
                iCol = iCol + 1
            End If
        End If
    Next
    
    If iCol = 0 Then
        comb_find.clear
        comb_find.SetFocus
        tb_psv.SetFocus
        DoEvents
        Exit Sub
    End If

    ReDim w(1 To iCol, 1 To 1)
    j = 1
    For i = LBound(cc) To UBound(cc)
    If cc(i, 1) <> "" Then
    w(j, 1) = cc(i, 1)
    j = j + 1
    End If
    Next
    
    comb_find.clear
    comb_find.SetFocus
    tb_psv.SetFocus
    DoEvents
    comb_find.List = w
    
End Sub


Private Sub comb_find_Click()
    On Error Resume Next
        
        With comb_find
            i = .ListIndex
            If i = -1 Then GoTo 99
            sZkz = .List(i, 0)
        End With
        
        flag = 1
        
        tb_psv.Text = sZkz
        comb_psv.ListIndex = -1
        
        flag = 0
99
End Sub


Private Sub comb_psv_Click()
        flag = 1
        tb_psv.Text = comb_psv.Value
        flag = 0
End Sub








Private Sub OK_Click()
        On Error Resume Next
        
        ThisWorkbook.Activate
        Sheets("Приход").Select
        
        Cells(rwPr_zkz, 4).Value = tb_psv.Text
        Cells(rwPr_mj, 4).Value = tb_mj.Text
        Cells(rwPr_dt, 4).Value = tb_dt1.Text
        
        Cells(1, prDoc).Value = tb_doc.Text
        Cells(1, prDocN).Value = "'" & tb_docN.Text
        Cells(1, prDocDt).Value = tb_dt2.Text
        
        Cells(rwPr_doc, 4).Value = tb_doc.Text & " № " & tb_docN.Text & " от " & tb_dt2.Text
        
        Unload Me
End Sub



Private Sub UserForm_Initialize()
        flag = 1
        Call load_dann
        flag = 0
End Sub

Private Sub load_dann()
        On Error Resume Next
        Call forma
        Call combo
        Call load_spr
End Sub

Private Sub forma()
        On Error Resume Next

        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("cmb_d")
            Me.Top = .Top + .Height + 20
            Me.Left = .Left
        End With
        
        With ThisWorkbook.Sheets("Приход")
            tb_psv.Text = .Cells(rwPr_zkz, 4).Value
            tb_mj.Text = .Cells(rwPr_mj, 4).Value
            tb_dt1.Text = .Cells(rwPr_dt, 4).Value
            
            tb_doc.Text = Cells(1, prDoc).Value
            tb_docN.Text = Cells(1, prDocN).Value
            tb_dt2.Text = Cells(1, prDocDt).Value
        End With
        
        OK.BackColor = RGB(58, 110, 165)
        OK.ForeColor = RGB(255, 255, 255)
        NO.ForeColor = RGB(255, 255, 255)
        
        If ThisWorkbook.Sheets("setting").Range("b35") = 1 Then
        Frame_doc.Height = 20
        Else
        Frame_doc.Height = 0
        End If
        
        With Frame_doc
            Frame_button.Top = .Top + .Height + 3
        End With


End Sub

Private Sub load_spr()
        On Error Resume Next

        Call load_zkz_all
        comb_psv.List = zkz
        
        Call load_mjj_all
        comb_Mj.List = mj
        
        Call load_doc_all
        comb_osn.List = doc
        
End Sub

Private Sub combo()
        On Error Resume Next
        
        With comb_Mj
        .Left = tb_mj.Left
        .Top = tb_mj.Top
        .Width = tb_mj.Width
        .ZOrder 1
        End With
        
        With comb_psv
        .Left = tb_psv.Left
        .Top = tb_psv.Top
        .Width = tb_psv.Width
        .ZOrder 1
        End With
        
        With comb_find
        .Left = tb_psv.Left
        .Top = tb_psv.Top
        .Width = tb_psv.Width
        .ZOrder 1
        End With
                
        With comb_osn
        .Left = tb_doc.Left
        .Top = tb_doc.Top
.Width = tb_doc.Width
        .ZOrder 1
        End With
End Sub

Private Sub NO_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub comb_osn_Click()
On Error Resume Next
tb_doc.Text = comb_osn.Value
tb_docN.SetFocus
End Sub
Private Sub comb_mj_Click()
On Error Resume Next
tb_mj.Text = comb_Mj.Value
End Sub



Private Sub tb_doc_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_osn.DropDown
End Sub
Private Sub tb_mj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_Mj.DropDown
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


Private Sub tb_psv_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_psv.DropDown
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
On Error Resume Next
add_psv.Show
End Sub


