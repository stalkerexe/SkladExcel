VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Set_zk 
   Caption         =   "Настройки формата"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4080
   OleObjectBlob   =   "frm_Set_zk.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Set_zk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub OK_Click()

        sheetNm = ActiveSheet.Name

        With ThisWorkbook.Sheets("setting")
        
            .Range("o24").Value = tb_rz.Text
            
            .Range("o26").Value = tb_m.Text
            
            If CheckBox_rz.Value = True Then .Range("o25").Value = 1
            If CheckBox_rz.Value = False Then .Range("o25").Value = 0
            
        End With
        
        Call setting_do
        
        
        Unload Me:    DoEvents

End Sub

Private Sub setting_do()
        Call doScreenOff
        Call do_obnov_pr
        Call do_obnov
Call setting_m
        Sheets(sheetNm).Select
        Call doScreenOn
End Sub

Private Sub setting_m()

    iMsh = ThisWorkbook.Sheets("setting").Range("o26").Value

    Sheets("Отложено_приход").Select
    ActiveWindow.Zoom = iMsh

    Sheets("Отложено_расход").Select
    ActiveWindow.Zoom = iMsh

End Sub



Private Sub UserForm_Initialize()
        Call load_dann
End Sub

Private Sub load_dann()
    On Error Resume Next

        With ThisWorkbook.Sheets("setting")
        
            tb_rz.Text = .Range("o24").Value
            
            tb_m.Text = .Range("o26").Value
            
            If .Range("o25").Value = 1 Then CheckBox_rz.Value = True
            If .Range("o25").Value = 0 Then CheckBox_rz.Value = False
            
        End With

        Me.StartUpPosition = 0
        With ActiveSheet.Shapes("pic_sett")
            Me.Top = .Top + .Height + 20
            Me.Left = .Left
        End With
        Call combo

End Sub

Private Sub combo()
        On Error Resume Next
        With comb_rz
        .Left = tb_rz.Left
        .Top = tb_rz.Top
        .Width = tb_rz.Width + 13
        .ZOrder 1
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        End With
        
        With comb_m
        .Left = tb_m.Left
        .Top = tb_m.Top
        .Width = tb_m.Width + 13
        .ZOrder 1
        .AddItem "75"
        .AddItem "80"
        .AddItem "90"
        .AddItem "100"
        .AddItem "110"
        .AddItem "120"
        End With
        
End Sub



Private Sub tb_rz_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_rz.DropDown
End Sub

Private Sub tb_m_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_m.DropDown
End Sub

Private Sub comb_rz_Click()
tb_rz.Text = comb_rz.Value
End Sub

Private Sub comb_m_Click()
tb_m.Text = comb_m.Value
End Sub



Private Sub NO_Click()
Unload Me
End Sub
