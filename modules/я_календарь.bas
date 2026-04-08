Attribute VB_Name = "я_календарь"
Option Explicit

Public dt_1 As Date
Public iForm As Object
Public iTb As Byte


Public Sub ShowForm4()
On Error Resume Next
dt_1 = VBA.Date
Form_SelectDate.Show
End Sub


