Attribute VB_Name = "Add_Damp"
Option Explicit

Public Sub ExportVBAProjectToTxt_UTF8()
    Dim filePath As String
    Dim comp As Object
    Dim i As Long
    Dim stream As Object
    Dim lineContent As String
    
    ' Путь к файлу
    filePath = ThisWorkbook.Path & "\VBA_Sklad_Dump_UTF8.txt"

    ' Создаем объект потока
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2 ' adTypeText
        .Charset = "utf-8" ' Устанавливаем UTF-8
        .Open
        
        ' Пишем заголовок (добавим кириллицу для проверки кодировки)
        .WriteText "===== ЭКСПОРТ ПРОЕКТА VBA (UTF-8) =====" & vbCrLf
        .WriteText "Книга: " & ThisWorkbook.Name & vbCrLf
        .WriteText "Дата: " & Now & vbCrLf
        .WriteText String(50, "=") & vbCrLf & vbCrLf

        ' ==== ДЕРЕВО ПРОЕКТА ====
        .WriteText "### СТРУКТУРА ПРОЕКТА ###" & vbCrLf
        For Each comp In ThisWorkbook.VBProject.VBComponents
            .WriteText "- " & comp.Name & " (" & ComponentTypeName(comp.Type) & ")" & vbCrLf
        Next comp
        .WriteText String(50, "-") & vbCrLf & vbCrLf

        ' ==== КОД МОДУЛЕЙ ====
        For Each comp In ThisWorkbook.VBProject.VBComponents
            .WriteText "=== КОМПОНЕНТ: " & comp.Name & " ===" & vbCrLf
            .WriteText "ТИП: " & ComponentTypeName(comp.Type) & vbCrLf
            .WriteText String(50, "-") & vbCrLf

            If comp.CodeModule.CountOfLines > 0 Then
                For i = 1 To comp.CodeModule.CountOfLines
                    ' Читаем строку кода
                    lineContent = comp.CodeModule.Lines(i, 1)
                    ' Записываем с номером строки
                    .WriteText Format(i, "0000") & " | " & lineContent & vbCrLf
                Next i
            Else
                .WriteText "[Пустой модуль]" & vbCrLf
            End If

            .WriteText vbCrLf & vbCrLf
        Next comp

        ' Сохраняем (2 = перезаписать файл)
        .SaveToFile filePath, 2
        .Close
    End With

    MsgBox "Готово! Файл сохранен в UTF-8:" & vbCrLf & filePath, vbInformation
End Sub

Private Function ComponentTypeName(ByVal t As Long) As String
    Select Case t
        Case 1: ComponentTypeName = "Standard Module"
        Case 2: ComponentTypeName = "Class Module"
        Case 3: ComponentTypeName = "UserForm"
        Case 100: ComponentTypeName = "ThisWorkbook / Sheet"
        Case Else: ComponentTypeName = "Unknown"
    End Select
End Function
