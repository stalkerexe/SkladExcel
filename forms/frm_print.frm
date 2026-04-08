VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_print 
   Caption         =   "Печать"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   OleObjectBlob   =   "frm_print.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub OK_Click()
        On Error Resume Next
        
        sPrinter = Me.Label2.Caption
        iCol_Print = Me.ComboBox1.Value

        If Me.TextBox1.text = 1 Then Call prnt_zv
        If Me.TextBox1.text = 2 Then Call prnt_pr
        
If Me.TextBox1.text = 3 Then Call prnt_zk
        If Me.TextBox1.text = 4 Then Call prnt_zk
        
        If Me.TextBox1.text = 5 Then Call print_ZVK_arh
        
        If Me.TextBox1.text = 7 Then Call prnt_sk

        Unload Me: DoEvents
End Sub

Private Sub prnt_sk()
On Error Resume Next
With ThisWorkbook.Sheets("Склад")
r7 = .Cells(Rows.Count, skNm).End(xlUp).Row
With .PageSetup
.PrintTitleRows = "$3:$5"
.PrintArea = "c1:k" & r7
End With
.PrintOut Copies:=iCol_Print, ActivePrinter:=sPrinter
End With
End Sub



Private Sub UserForm_Initialize()
On Error Resume Next
OK.BackColor = RGB(55, 96, 145)
OK.ForeColor = RGB(255, 255, 255)
NO.ForeColor = RGB(255, 255, 255)
Me.ComboBox1.AddItem "1"
Me.ComboBox1.AddItem "2"
Me.ComboBox1.AddItem "3"
ComboBox1.ListIndex = 0
Call list_print
End Sub

Private Sub list_print()
On Error Resume Next
Dim nn
With CreateObject("Shell.Application").Namespace(4).Items
For nn = 1 To .Count - 1
Me.ComboBox2.AddItem .Item(nn).Name
Next
End With
Label2.Caption = Application.ActivePrinter
Me.ComboBox2.Value = Label2.Caption
End Sub

Private Sub ComboBox2_Click()
On Error Resume Next
Label2.Caption = ComboBox2.Value
TextBox1.SetFocus
End Sub

Private Sub ComboBox1_Change()
On Error Resume Next
OK.SetFocus
End Sub

Private Sub NO_Click()
Unload Me
End Sub
