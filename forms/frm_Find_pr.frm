VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Find_pr 
   Caption         =   "Приходы"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   OleObjectBlob   =   "frm_Find_pr.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Find_pr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r As Long

Private Sub ListBox1_Click()
On Error Resume Next
ind = ListBox1.ListIndex
rw = ListBox1.List(ind, 0)
Range(Cells(rw, 1), Cells(rw, 22)).Select
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
r = Cells(Rows.Count, pzkNm).End(xlUp).Row + 1
nn = Range(Cells(4, 1), Cells(r, 1)).Value
nom = Range(Cells(4, pzkNom), Cells(r, pzkNom)).Value
nm = Range(Cells(4, pzkNm), Cells(r, pzkNm)).Value
dt = Range(Cells(4, pzkDt), Cells(r, pzkDt)).Value
mj = Range(Cells(4, pzkMj), Cells(r, pzkMj)).Value
iCol = Application.CountIf(Range("a4:a" & r), "<>")
If iCol = 0 Then Exit Sub
ReDim c(1 To iCol, 1 To 5)
n = 1
For i = LBound(nm) To UBound(nm)
If nn(i, 1) <> "" Then
c(n, 1) = i + 3
c(n, 2) = nm(i, 1)
c(n, 3) = Format(nom(i, 1), "00000")
c(n, 4) = Format(dt(i, 1), "dd.mm.yyyy")
c(n, 5) = mj(i, 1)
n = n + 1
End If
Next
ListBox1.List = c
ListBox1.ColumnWidths = "0;140;50;80"
End Sub
Private Sub ico_Click()
Unload Me
End Sub

