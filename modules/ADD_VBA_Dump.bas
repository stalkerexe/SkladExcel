Attribute VB_Name = "ADD_VBA_Dump"
Option Explicit
Public Sub ExportProjectToUTF8()
    Dim comp As Object
    Dim stream As Object
    Dim i As Long
    Dim filePath As String
    Dim lineContent As String
    
    ' Путь сохранения (в папку с книгой)
    filePath = ThisWorkbook.path & "\VBA_Project_Dump_UTF8.txt"
    
    ' Создаем объект для работы с кодировкой (ADODB.Stream)
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        MsgBox "Ошибка: Не удалось создать ADODB.Stream. Проверьте права системы.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    
    ' Формируем "шапку" файла
    stream.WriteText "===== VBA PROJECT EXPORT =====" & vbCrLf
    stream.WriteText "Workbook: " & ThisWorkbook.Name & vbCrLf
    stream.WriteText "Exported: " & Now & vbCrLf
    stream.WriteText String(50, "=") & vbCrLf & vbCrLf

    ' Проходим по всем компонентам (модули, классы, формы)
    For Each comp In ThisWorkbook.VBProject.VBComponents
        stream.WriteText "=== COMPONENT: " & comp.Name & " ===" & vbCrLf
        stream.WriteText "TYPE: " & GetCompTypeName(comp.Type) & vbCrLf
        stream.WriteText String(50, "-") & vbCrLf
        
        If comp.CodeModule.CountOfLines > 0 Then
            For i = 1 To comp.CodeModule.CountOfLines
                ' Просто берем строку кода без добавления номера и разделителей
                lineContent = comp.CodeModule.Lines(i, 1)
                stream.WriteText lineContent & vbCrLf
            Next i
        Else
            stream.WriteText "' [No Code in Module]" & vbCrLf
        End If
        
        stream.WriteText vbCrLf & vbCrLf
    Next comp

    ' Сохраняем и закрываем
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    
    MsgBox "Готово! Файл сохранен здесь:" & vbCrLf & filePath, vbInformation
End Sub

' Та самая "потерянная" функция для определения типа компонента
Private Function GetCompTypeName(ByVal compType As Long) As String
    Select Case compType
        Case 1: GetCompTypeName = "Standard Module"
        Case 2: GetCompTypeName = "Class Module"
        Case 3: GetCompTypeName = "Microsoft Form"
        Case 11: GetCompTypeName = "ActiveX Designer"
        Case 100: GetCompTypeName = "Document (Sheet/ThisWorkbook)"
        Case Else: GetCompTypeName = "Unknown (" & compType & ")"
    End Select
End Function

