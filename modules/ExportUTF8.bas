Attribute VB_Name = "ExportUTF8"
Option Explicit

' ================================================================
' МОДУЛЬ: modGitExport
' Экспорт VBA-компонентов в папку srcCloude/ для GitHub/Claude
' ================================================================
' Структура:
'   src/
'     modules/    — стандартные модули (.bas)
'     classes/    — классы (.cls)
'     forms/      — формы (.frm)
'     sheets/     — код листов и ЭтойКниги (.bas, utf-8 BOM)
'
' .gitignore: *.frx
' ================================================================

Private Const EXPORT_ROOT As String = "srcCloudeUTF8"

Private Type ExportResult
    Exported As Long
    Skipped  As Long
    Errors   As Long
    ErrorLog As String
End Type


' ================================================================
' ТОЧКА ВХОДА
' ================================================================
Public Sub ExportAllModulesUTF8()

    Dim result   As ExportResult
    Dim rootPath As String

    rootPath = ThisWorkbook.Path & "\" & EXPORT_ROOT & "\"

    EnsureFolder rootPath
    EnsureFolder rootPath & "modules\"
    EnsureFolder rootPath & "classes\"
    EnsureFolder rootPath & "forms\"
    EnsureFolder rootPath & "sheets\"

    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ExportComponent comp, rootPath, result
    Next comp

    Dim msg As String
    msg = "Экспорт завершён:" & vbCrLf & _
          "  Экспортировано: " & result.Exported & vbCrLf & _
          "  Пропущено:      " & result.Skipped

    If result.Errors > 0 Then
        msg = msg & vbCrLf & "  Ошибок: " & result.Errors & _
              vbCrLf & vbCrLf & result.ErrorLog
        MsgBox msg, vbExclamation, "Экспорт VBA"
    Else
        MsgBox msg & vbCrLf & vbCrLf & rootPath, vbInformation, "Экспорт VBA"
    End If

End Sub


' ================================================================
' МАРШРУТИЗАЦИЯ КОМПОНЕНТА
' ================================================================
Private Sub ExportComponent( _
    ByVal comp As Object, _
    ByVal rootPath As String, _
    ByRef result As ExportResult _
)
    ' Сам себя не экспортируем
    If comp.Name = "modGitExport" Or comp.Name = "ADD_VBA_Dump" Then
        result.Skipped = result.Skipped + 1
        Exit Sub
    End If

    Select Case comp.Type

        Case 1   ' StdModule
            ExportViaAPI comp, rootPath & "modules\" & comp.Name & ".bas", result

        Case 2   ' ClassModule
            ExportViaAPI comp, rootPath & "classes\" & comp.Name & ".cls", result

        Case 3   ' MSForm
            ExportViaAPI comp, rootPath & "forms\" & comp.Name & ".frm", result

        Case 100 ' Document (листы, ЭтаКнига)
            ExportDocumentModule comp, rootPath & "sheets\", result

        Case Else
            result.Skipped = result.Skipped + 1

    End Select
End Sub


' ================================================================
' ЭКСПОРТ ЧЕРЕЗ comp.Export (модули, классы, формы)
' ================================================================
Private Sub ExportViaAPI( _
    ByVal comp As Object, _
    ByVal filePath As String, _
    ByRef result As ExportResult _
)
    On Error GoTo Fail
    comp.Export filePath
    result.Exported = result.Exported + 1
    Exit Sub
Fail:
    result.Errors = result.Errors + 1
    result.ErrorLog = result.ErrorLog & comp.Name & ": " & Err.Description & vbCrLf
    On Error GoTo 0
End Sub


' ================================================================
' ЭКСПОРТ КОДА ЛИСТА / ЭТОЙКНИГИ ЧЕРЕЗ ADODB.Stream (UTF-8 BOM)
'
' Почему нельзя comp.Export:
'   Document-компоненты привязаны к объекту книги, Excel
'   не позволяет экспортировать их как отдельный файл —
'   метод Export для Type=100 выбрасывает ошибку.
'
' Почему UTF-8 с BOM:
'   BOM (EF BB BF) позволяет GitHub, VSCode и Claude корректно
'   определять кодировку. VBA при повторном импорте .bas/.cls
'   читает файлы в системной кодировке (cp1251), поэтому
'   такие файлы предназначены только для просмотра/ревью,
'   не для прямого реимпорта в VBE.
' ================================================================
Private Sub ExportDocumentModule( _
    ByVal comp As Object, _
    ByVal folderPath As String, _
    ByRef result As ExportResult _
)
    ' Пропускаем компоненты без кода
    If comp.CodeModule.CountOfLines = 0 Then
        result.Skipped = result.Skipped + 1
        Exit Sub
    End If

    Dim i       As Long
    Dim content As String

    ' Заголовок-комментарий: имя компонента и тип объекта
    Dim objName As String
    On Error Resume Next
    objName = comp.Properties("Name").Value
    On Error GoTo 0
    If objName = "" Then objName = comp.Name

    content = "' Component: " & comp.Name & "  [" & objName & "]" & vbCrLf & _
              "' Type: Document (Sheet / ThisWorkbook)" & vbCrLf & _
              "Option Explicit" & vbCrLf & vbCrLf

    ' Пропускаем первую строку если это уже "Option Explicit"
    Dim startLine As Long
    startLine = 1
    If LCase(Trim(comp.CodeModule.Lines(1, 1))) = "option explicit" Then
        startLine = 2
    End If

    For i = startLine To comp.CodeModule.CountOfLines
        content = content & comp.CodeModule.Lines(i, 1) & vbCrLf
    Next i

    ' Имя файла: имя компонента + "_" + имя вкладки (если отличается)
    Dim fileName As String
    If objName <> comp.Name And objName <> "" Then
        fileName = comp.Name & "_" & CleanFileName(objName) & ".bas"
    Else
        fileName = comp.Name & ".bas"
    End If

    Dim filePath As String
    filePath = folderPath & fileName

    ' Запись через ADODB.Stream в UTF-8 с BOM
    On Error GoTo Fail
    WriteFileUtf8Bom filePath, content
    On Error GoTo 0

    result.Exported = result.Exported + 1
    Exit Sub

Fail:
    result.Errors = result.Errors + 1
    result.ErrorLog = result.ErrorLog & comp.Name & ": " & Err.Description & vbCrLf
    On Error GoTo 0
End Sub


' ================================================================
' ЗАПИСЬ ТЕКСТА В ФАЙЛ В КОДИРОВКЕ UTF-8 С BOM (EF BB BF)
'
' Алгоритм:
'   1. Пишем текст в текстовый Stream с charset = utf-8
'      (ADODB сам вставляет BOM при utf-8)
'   2. Переключаем Stream в бинарный режим
'   3. Сохраняем в файл
'
' Примечание: ADODB.Stream с Charset="utf-8" автоматически
' добавляет BOM при переключении в бинарный режим через
' CopyTo — именно этим трюком мы и пользуемся.
' ================================================================
Private Sub WriteFileUtf8Bom(ByVal filePath As String, ByVal content As String)

    ' --- текстовый поток (UTF-8, ADODB сам добавит BOM) ---
    Dim tsText As Object
    Set tsText = CreateObject("ADODB.Stream")
    tsText.Type = 2             ' adTypeText
    tsText.Charset = "utf-8"
    tsText.Open
    tsText.WriteText content

    ' --- бинарный поток для сохранения ---
    Dim tsBin As Object
    Set tsBin = CreateObject("ADODB.Stream")
    tsBin.Type = 1              ' adTypeBinary
    tsBin.Open

    ' Перемотка в начало и копирование (BOM уже включён)
    tsText.Position = 0
    tsText.CopyTo tsBin

    tsBin.SaveToFile filePath, 2    ' adSaveCreateOverWrite

    tsText.Close
    tsBin.Close
    Set tsText = Nothing
    Set tsBin = Nothing

End Sub


' ================================================================
' ОЧИСТКА ИМЕНИ ФАЙЛА ОТ НЕДОПУСТИМЫХ СИМВОЛОВ
' ================================================================
Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant
    Dim i   As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    CleanFileName = s
End Function


' ================================================================
' СОЗДАТЬ ПАПКУ (если не существует)
' ================================================================
Private Sub EnsureFolder(ByVal Path As String)
    If Dir(Path, vbDirectory) = "" Then
        On Error Resume Next
        MkDir Path
        On Error GoTo 0
    End If
End Sub


