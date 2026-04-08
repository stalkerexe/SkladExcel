Attribute VB_Name = "ш_меню"
Option Explicit

Dim hh As Double
Dim h As Double
Public Sub show_mn_vid()
On Error Resume Next
With ActiveSheet.Shapes("mn_vid")
.Height = 10
.Visible = True
.Top = ActiveSheet.Shapes("cmb_vd").Top + ActiveSheet.Shapes("cmb_vd").Height + 4
.Left = ActiveSheet.Shapes("cmb_vd").Left
End With
Call show_mn_vid_
End Sub
Private Sub show_mn_vid_()
On Error Resume Next
With ActiveSheet.Shapes("mn_vid")
h = 112
.Visible = True
For hh = 0 To h Step h / 2
.Height = hh
DoEvents
Next hh
End With
End Sub
Public Sub unload_mn_vid()
On Error Resume Next
With ThisWorkbook.Sheets("Расход").Shapes("mn_vid")
.Height = 10
.Top = 10
.Visible = False
DoEvents
End With
End Sub


Public Sub show_mn_vid_pr()
On Error Resume Next
With ActiveSheet.Shapes("mn_vid_pr")
.Height = 10
.Visible = True
.Top = ActiveSheet.Shapes("cmb_vd").Top + ActiveSheet.Shapes("cmb_vd").Height + 4
.Left = ActiveSheet.Shapes("cmb_vd").Left
End With
Call show_mn_vid_pr_
End Sub
Private Sub show_mn_vid_pr_()
On Error Resume Next
With ActiveSheet.Shapes("mn_vid_pr")
h = 112
.Visible = True
For hh = 0 To h Step h / 4
.Height = hh
Next hh
DoEvents
End With
End Sub
Public Sub unload_mn_vid_pr()
On Error Resume Next
With ThisWorkbook.Sheets("Приход").Shapes("mn_vid_pr")
h = .Height
.Height = hh
.Top = 10
.Visible = False
End With
DoEvents
End Sub
Private Sub sh_mn_mn()
On Error Resume Next
frm_Mnn.Show
End Sub
Public Sub unload_mn_mn()
On Error Resume Next
Unload frm_Mnn
End Sub


