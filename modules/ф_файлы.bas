Attribute VB_Name = "ф_файлы"
Option Explicit

Public Sub load_sk()
On Error Resume Next
Set dic_sk = CreateObject("Scripting.Dictionary")
dic_sk.Add 0, "Материалы"
dic_sk.Add 1, "Металлопрокат"
dic_sk.Add 2, "Спецодежда"
End Sub

Public Sub sh_frm_Skidka()
        On Error Resume Next
        Call unload_mn_vid
        Unload frm_Oplata
        frm_Skidka.Show
End Sub

Public Sub sh_frm_Oplata()
        On Error Resume Next
        Call unload_mn_vid
        Unload frm_Skidka
        frm_Oplata.Show
End Sub

