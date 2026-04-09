VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Find_zk 
   Caption         =   "Текущие заказы"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   OleObjectBlob   =   "frm_Find_zk.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Find_zk"
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
r = Cells(Rows.Count, zkNm).End(xlUp).Row + 1
nn = Range(Cells(4, 1), Cells(r, 1)).value
nom = Range(Cells(4, zkNom), Cells(r, zkNom)).value
nm = Range(Cells(4, zkNm), Cells(r, zkNm)).value
dt = Range(Cells(4, zkDt1), Cells(r, zkDt1)).value
mj = Range(Cells(4, zkMj), Cells(r, zkMj)).value
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
