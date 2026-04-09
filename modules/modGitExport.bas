Attribute VB_Name = "modGitExport"

Option Explicit

Private Const EXPORT_ROOT As String = "srcGIT"

Private Type ExportResult
    Exported As Long
    Skipped  As Long
    Errors   As Long
    ErrorLog As String
End Type

' ================================================================
' ЭКСПОРТ
' ================================================================
Public Sub ExportAllModulesGIT()
    Dim result As ExportResult
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
          "  Пропущено:      " & result.Skipped & vbCrLf & _
          "  Ошибок:         " & result.Errors

    If result.ErrorLog <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Лог ошибок:" & vbCrLf & result.ErrorLog
    End If

    MsgBox msg, vbInformation, "Экспорт VBA"
End Sub

Private Sub ExportComponent(ByVal comp As Object, ByVal rootPath As String, ByRef result As ExportResult)
    If comp.Name = "modGitExport" Then
        result.Skipped = result.Skipped + 1
        Exit Sub
    End If

    Select Case comp.Type
        Case 1:   ExportViaAPI comp, rootPath & "modules\" & comp.Name & ".bas", result
        Case 2:   ExportViaAPI comp, rootPath & "classes\" & comp.Name & ".cls", result
        Case 3:   ExportViaAPI comp, rootPath & "forms\" & comp.Name & ".frm", result
        Case 100: ExportDocumentModule comp, rootPath & "sheets\", result
        Case Else: result.Skipped = result.Skipped + 1
    End Select
End Sub

Private Sub ExportViaAPI(ByVal comp As Object, ByVal filePath As String, ByRef result As ExportResult)
    Dim tempPath As String: tempPath = ""
    On Error GoTo Fail
    tempPath = filePath & ".tmp"
    comp.Export tempPath
    Dim content As String: content = ReadFileAnsi(tempPath)
    WriteFileUtf8Bom filePath, content
    SafeKill tempPath
    result.Exported = result.Exported + 1
    Exit Sub
Fail:
    result.Errors = result.Errors + 1
    result.ErrorLog = result.ErrorLog & comp.Name & ": " & Err.Description & vbCrLf
    SafeKill tempPath
    On Error GoTo 0
End Sub

Private Sub ExportDocumentModule(ByVal comp As Object, ByVal folderPath As String, ByRef result As ExportResult)
    If comp.CodeModule.CountOfLines = 0 Then
        result.Skipped = result.Skipped + 1
        Exit Sub
    End If
    On Error GoTo Fail
    Dim objName As String: objName = ""
    On Error Resume Next
    objName = comp.Properties("Name").value
    On Error GoTo Fail
    If objName = "" Then objName = comp.Name
    Dim startLine As Long: startLine = 1
    If comp.CodeModule.CountOfLines > 0 Then
        If LCase(Trim(comp.CodeModule.lines(1, 1))) = "option explicit" Then startLine = 2
    End If
    Dim codeLines As String: codeLines = ""
    If comp.CodeModule.CountOfLines >= startLine Then
        codeLines = comp.CodeModule.lines(startLine, comp.CodeModule.CountOfLines - startLine + 1)
    End If
    Dim content As String
    content = "' Component: " & comp.Name & "  [" & objName & "]" & vbCrLf & _
              "' Type: Document (Sheet / ThisWorkbook)" & vbCrLf & _
              "Option Explicit" & vbCrLf & vbCrLf & codeLines
    Dim fileName As String
    If objName <> comp.Name And objName <> "" Then
        fileName = comp.Name & "_" & CleanFileName(objName) & ".bas"
    Else
        fileName = comp.Name & ".bas"
    End If
    WriteFileUtf8Bom folderPath & fileName, content
    result.Exported = result.Exported + 1
    Exit Sub
Fail:
    result.Errors = result.Errors + 1
    result.ErrorLog = result.ErrorLog & comp.Name & " (doc): " & Err.Description & vbCrLf
    On Error GoTo 0
End Sub

' ================================================================
' ИМПОРТ
' ================================================================
Public Sub A_ImportAllModules()
    Dim rootPath As String
    rootPath = ThisWorkbook.Path & "\" & EXPORT_ROOT & "\"
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Папка не найдена: " & rootPath, vbExclamation, "Импорт"
        Exit Sub
    End If
    Dim imported As Long: imported = 0
    Dim Errors As Long: Errors = 0
    Dim ErrorLog As String: ErrorLog = ""
    ImportFolder rootPath & "modules\", imported, Errors, ErrorLog
    ImportFolder rootPath & "classes\", imported, Errors, ErrorLog
    ImportFolder rootPath & "forms\", imported, Errors, ErrorLog
    Dim msg As String
    msg = "Импорт завершён:" & vbCrLf & "  Импортировано: " & imported & vbCrLf & "  Ошибок: " & Errors
    If ErrorLog <> "" Then msg = msg & vbCrLf & vbCrLf & "Лог ошибок:" & vbCrLf & ErrorLog
    MsgBox msg, vbInformation, "Импорт VBA"
End Sub

Private Sub ImportFolder(ByVal folderPath As String, ByRef imported As Long, ByRef Errors As Long, ByRef ErrorLog As String)
    If Dir(folderPath, vbDirectory) = "" Then Exit Sub
    Dim files() As String
    Dim fileCount As Long: fileCount = 0
    ReDim files(0 To 0)
    Dim f As String: f = Dir(folderPath & "*.*")
    Do While f <> ""
        ReDim Preserve files(0 To fileCount)
        files(fileCount) = f
        fileCount = fileCount + 1
        f = Dir
    Loop
    If fileCount = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To fileCount - 1
        Dim fileName As String: fileName = files(i)
        Dim ext As String: ext = LCase(Right(fileName, 4))
        If ext = ".bas" Or ext = ".cls" Then
            ImportSingleFile folderPath & fileName, imported, Errors, ErrorLog
        ElseIf ext = ".frm" Then
            Dim fullPath As String: fullPath = folderPath & fileName
            If FileExists(fullPath & ".frx") Or FileExists(Left(fullPath, Len(fullPath) - 4) & ".frx") Or FileExists(fullPath & ".frm.frx") Then
                ImportSingleFile fullPath, imported, Errors, ErrorLog
            Else
                Errors = Errors + 1
                ErrorLog = ErrorLog & fileName & ": нет парного .frx файла" & vbCrLf
            End If
        End If
    Next i
End Sub

Private Sub ImportSingleFile(ByVal filePath As String, ByRef imported As Long, ByRef Errors As Long, ByRef ErrorLog As String)
    Dim tempPath As String: tempPath = ""
    Dim tempFrxPath As String: tempFrxPath = ""
    On Error GoTo Fail
    Dim content As String: content = ReadFileUtf8NoBom(filePath)
    If content = "" Then
        Errors = Errors + 1
        ErrorLog = ErrorLog & GetFileName(filePath) & ": файл пустой" & vbCrLf
        Exit Sub
    End If
    Dim compName As String: compName = ExtractVBName(content)
    If compName = "" Then compName = GetFileNameNoExt(filePath)
    Dim origExt As String: origExt = LCase(GetFileExt(filePath))

    ' Имя файла в TEMP должно совпадать с именем компонента
    tempPath = Environ("TEMP") & "\" & compName & origExt
    WriteFileAnsi tempPath, content

    If origExt = ".frm" Then
        ' Если ссылка внутри: "add_psv.frm.frx", то имя в TEMP должно быть таким же
        tempFrxPath = tempPath & ".frx"
        
        Dim srcFrx As String: srcFrx = ""
        ' Проверяем варианты: Form.frm.frx или Form.frx
        If FileExists(filePath & ".frx") Then
            srcFrx = filePath & ".frx"
        ElseIf FileExists(Left(filePath, Len(filePath) - 4) & ".frx") Then
            srcFrx = Left(filePath, Len(filePath) - 4) & ".frx"
        End If
        
        If srcFrx <> "" Then FileCopy srcFrx, tempFrxPath
    End If

    RemoveComponentIfExists compName
    ThisWorkbook.VBProject.VBComponents.Import tempPath
    
    SafeKill tempPath
    SafeKill tempFrxPath
    imported = imported + 1
    Exit Sub
Fail:
    Errors = Errors + 1
    ErrorLog = ErrorLog & GetFileName(filePath) & ": " & Err.Description & vbCrLf
    SafeKill tempPath: SafeKill tempFrxPath
    On Error GoTo 0
End Sub

' ================================================================
' FILE I/O
' ================================================================
Private Function ReadFileAnsi(ByVal filePath As String) As String
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    With s: .Type = 2: .Charset = "windows-1251": .Open: .LoadFromFile filePath: ReadFileAnsi = .ReadText: .Close: End With
End Function

Private Function ReadFileUtf8NoBom(ByVal filePath As String) As String
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    Dim raw As String
    With s: .Type = 2: .Charset = "utf-8": .Open: .LoadFromFile filePath: raw = .ReadText: .Close: End With
    If Len(raw) > 0 Then If AscW(Left(raw, 1)) = 65279 Then raw = Mid(raw, 2)
    ReadFileUtf8NoBom = raw
End Function

Private Sub WriteFileUtf8Bom(ByVal filePath As String, ByVal content As String)
    Dim sText As Object: Set sText = CreateObject("ADODB.Stream")
    Dim sBin As Object: Set sBin = CreateObject("ADODB.Stream")
    sBin.Type = 1: sBin.Open
    With sText: .Type = 2: .Charset = "utf-8": .Open: .WriteText content: .Position = 0: .CopyTo sBin: .Close: End With
    sBin.SaveToFile filePath, 2: sBin.Close
End Sub

Private Sub WriteFileAnsi(ByVal filePath As String, ByVal content As String)
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    With s: .Type = 2: .Charset = "windows-1251": .Open: .WriteText content: .SaveToFile filePath, 2: .Close: End With
End Sub

' ================================================================
' ВСПОМОГАТЕЛЬНЫЕ
' ================================================================
Private Sub SafeKill(ByVal filePath As String)
    On Error Resume Next: If filePath <> "" Then If Dir(filePath) <> "" Then Kill filePath: On Error GoTo 0
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next: FileExists = (Len(filePath) > 0) And (Len(Dir(filePath)) > 0): On Error GoTo 0
End Function

Private Function ExtractVBName(ByVal content As String) As String
    Dim lines() As String: lines = Split(Replace(content, vbCrLf, vbLf), vbLf)
    Dim i As Long: For i = 0 To IIf(UBound(lines) > 20, 20, UBound(lines))
        If LCase(Left(Trim(lines(i)), 19)) = "attribute vb_name =" Then
            Dim parts() As String: parts = Split(lines(i), """")
            If UBound(parts) >= 1 Then ExtractVBName = parts(1): Exit Function
        End If
    Next i
End Function

Private Sub RemoveComponentIfExists(ByVal compName As String)
    On Error Resume Next
    Dim comp As Object: Set comp = ThisWorkbook.VBProject.VBComponents(compName)
    If Not comp Is Nothing Then If comp.Type <> 100 Then ThisWorkbook.VBProject.VBComponents.Remove comp
    On Error GoTo 0
End Sub

Private Function GetFileName(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, "\"): GetFileName = IIf(pos > 0, Mid(filePath, pos + 1), filePath)
End Function

Private Function GetFileNameNoExt(ByVal filePath As String) As String
    Dim fn As String: fn = GetFileName(filePath): Dim pos As Long: pos = InStrRev(fn, "."): GetFileNameNoExt = IIf(pos > 1, Left(fn, pos - 1), fn)
End Function

Private Function GetFileExt(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, "."): GetFileExt = IIf(pos > 0, Mid(filePath, pos), "")
End Function

Private Function CleanFileName(ByVal s As String) As String
    Dim bad, i As Long: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For i = 0 To UBound(bad): s = Replace(s, bad(i), "_"): Next: CleanFileName = s
End Function

Private Sub EnsureFolder(ByVal Path As String)
    If Dir(Path, vbDirectory) = "" Then On Error Resume Next: MkDir Path: On Error GoTo 0
End Sub

