VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Doc 
   Caption         =   "Документ"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   OleObjectBlob   =   "frm_Doc.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Doc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OK_Click()
        Call ent_doc
End Sub

Private Sub ent_doc()
        sheetNm = "Отложено_приход"
        cm = pzkOsn
        Sheets(sheetNm).Select
        sFiles = "zkz_prihod.xlsx"
        Call dann
        Call add_dann_hear
        Unload Me
End Sub

Private Sub dann()
        With Me
            marker = .tb_mk.Text
            sDoc = .tb_doc.Text
            sDocN = "'" & .tb_docN.Text
            sDocDt = .tb_dt1.Text
            sOsn = sDoc & " № " & sDocN & " от " & sDocDt
            iRow = .tb_row.value
        End With
End Sub

Private Sub add_dann_hear()
        Cells(iRow, pzkDoc) = sDoc
        Cells(iRow, pzkDocN) = sDocN
        Cells(iRow, pzkDocDt) = sDocDt
        Cells(iRow + 1, pzkOsn) = sOsn
        Cells(iRow, cm).Select
        Call remove_green
End Sub



Private Sub UserForm_Initialize()
        On Error Resume Next
        Call doScreenOff
        Me.StartUpPosition = 0
        Me.Top = 200
        Me.Left = Cells(2, pzkOsn).Left
        Call combo
        Call doScreenOn
End Sub

Private Sub combo()
        On Error Resume Next
        With comb_osn
            .Left = tb_doc.Left
            .Top = tb_doc.Top
            .Width = tb_doc.Width
            .Height = tb_doc.Height
            .ZOrder 1
        End With
        
        Call load_doc_all
        comb_osn.List = doc
99
End Sub

Private Sub tb_doc_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
comb_osn.DropDown
End Sub
Private Sub comb_osn_Click()
tb_doc.Text = comb_osn.value
End Sub

Private Sub tb_dt1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
iTb = 0
Set iForm = Me
Call ShowForm4
End Sub


Private Sub NO_Click()
Unload Me
End Sub

