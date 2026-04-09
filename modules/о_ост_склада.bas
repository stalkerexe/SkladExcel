Attribute VB_Name = "о_ост_склада"
Option Explicit
Public Sub ost_skds()
        On Error Resume Next

        For i = LBound(nm) To UBound(nm)
            If nm(i, 1) <> "" Then

                    sSk = sk(i, 1)
                    rw = id(i, 1) + 1

                    Call find_cm_sk
                    Call sk_ost_do

            End If
        Next i

End Sub

' ИСПРАВЛЕНИЕ #5: добавлена явная проверка наличия листа sk_123.
' В оригинале On Error Resume Next молча проглатывал ошибку если лист
' отсутствует или переименован — остатки не обновлялись без предупреждения.
' Теперь используется RequireSheet: при отсутствии листа пользователь
' видит сообщение, а не молчаливые неверные данные.
Private Sub sk_ost_do()
        Dim ws As Worksheet
        If Not RequireSheet("sk_123", ws, "sk_ost_do") Then Exit Sub

        On Error Resume Next
        With ws
            If iOperation = "zv" Then .Cells(rw, cm) = .Cells(rw, cm) - CDbl(col(i, 1))
            If iOperation = "pr" Then .Cells(rw, cm) = .Cells(rw, cm) + CDbl(col(i, 1))
            If iOperation = "vz" Then .Cells(rw, cm) = .Cells(rw, cm) + CDbl(col(i, 1))
        End With
End Sub


Public Sub find_cm_sk()
        If sSk = "Материалы" Then cm = 2
        If sSk = "Металлопрокат" Then cm = 4
        If sSk = "Спецодежда" Then cm = 6
End Sub


