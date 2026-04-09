VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Skidka 
   Caption         =   "Скидка"
   ClientHeight    =   1200
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4065
   OleObjectBlob   =   "frm_Skidka.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Skidka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim cn()




Private Sub OK_Click()

        If Me.tb_nm.value = "" Then Me.tb_nm.value = 0
'        If Me.tb_nm.Value = "" Then
'            MsgBox "Введите процент скидки!!", 64, "Скидка"
'            tb_nm.SetFocus
'            DoEvents
'            Exit Sub
'        End If

        Call do_ok
        
        Unload Me
End Sub

Private Sub do_ok()
        On Error Resume Next

        ThisWorkbook.Activate
        Sheets("Расход").Select
        
        iSkid = CDbl(tb_nm.value)
        
        Call arr_zv

        With ThisWorkbook.Sheets("Расход")
            r7 = .Cells(Rows.Count, zvNm).End(xlUp).Row + 2
cn = Range(.Cells(rwZv, zvCn), .Cells(r7, zvCn)).value
        End With
        
        ReDim cnR(LBound(nm) To UBound(nm), 1 To 4)
        ReDim sm(LBound(nm) To UBound(nm), 1 To 4)

        For i = LBound(nm) To UBound(nm)
            If nm(i, 1) <> "" Then
            
                    sCol = col(i, 1)
            
                    sCn = cn(i, 1)
                    
                    sCnR = sCn - (sCn * iSkid / 100)
                    
                    iSm = sCol * sCnR
                    
                    cnR(i, 1) = sCnR
                    sm(i, 1) = iSm
                    
        
            End If
        Next
        
        
        With ThisWorkbook.Sheets("Расход")
            .Cells(rwZv, zvCnR).Resize(UBound(sm), 1) = cnR
            .Cells(rwZv, zvSm).Resize(UBound(sm), 1) = sm
            
            .Cells(rwZv_mj, zvOst).value = iSkid
        End With


        r7 = Cells(Rows.Count, zvNm).End(xlUp).Row
        Cells(rwzvSm, zvSm) = Application.Sum(Range(Cells(rwZv, zvSm), Cells(r7 + 4, zvSm)))

End Sub









Private Sub UserForm_Initialize()

        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("cmb_skidka")
            Me.Top = .Top + .Height + 20
            Me.Left = .Left - Me.Width + 20
        End With
        
        Call combo

        tb_nm.value = Cells(rwZv_mj, zvOst)

End Sub

Private Sub combo()
        On Error Resume Next
        With comb_nm
        .Left = tb_nm.Left
        .Top = tb_nm.Top
        .Width = tb_nm.Width + 13
        .AddItem 3
        .AddItem 5
        .AddItem 7
        .AddItem 10
        .AddItem 15
        .AddItem 20
        .AddItem 30
        .ZOrder 1
        End With
End Sub


Private Sub tb_nm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        comb_nm.DropDown
End Sub

Private Sub comb_nm_Click()
        tb_nm.value = comb_nm.value
End Sub

Private Sub NO_Click()
Unload Me
End Sub
