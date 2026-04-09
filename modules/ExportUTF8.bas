Attribute VB_Name = "ExportUTF8"
Option Explicit

Private Const EXPORT_ROOT As String = "srcCloudeUTF8"

Private Type ExportResult
    Exported As Long
    Skipped  As Long
    Errors   As Long
    ErrorLog As String
End Type


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


Private Sub ExportComponent( _
    ByVal comp As Object, _
    ByVal rootPath As String, _
    ByRef result As ExportResult _
)
    If comp.Name = "modGitExport" Or comp.Name = "ADD_VBA_Dump" Then
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


' ИСПРАВЛЕНИЕ #13: временный .tmp файл теперь удаляется и в блоке Fail.
' В оригинале Kill tempPath выполнялся только в нормальном потоке.
' При ошибке в ReadFileAnsi или WriteFileUtf8Bom файл .tmp оставался
' на диске рядом с .bas файлами навсегда.
Private Sub ExportViaAPI( _
    ByVal comp As Object, _
    ByVal filePath As String, _
    ByRef result As ExportResult _
)
    On Error GoTo Fail

    Dim tempPath As String
    tempPath = filePath & ".tmp"

    comp.Export tempPath

    Dim content As String
    content = ReadFileAnsi(tempPath)

    WriteFileUtf8Bom filePath, content

    Kill tempPath

    result.Exported = result.Exported + 1
    Exit Sub

Fail:
    result.Errors = result.Errors + 1
    result.ErrorLog = result.ErrorLog & comp.Name & ": " & Err.Description & vbCrLf

    ' ИСПРАВЛЕНИЕ: убираем .tmp даже при ошибке
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If Dir(tempPath) <> "" Then Kill tempPath
    End If
    On Error GoTo 0
End Sub


Private Sub ExportDocumentModule( _
    ByVal comp As Object, _
    ByVal folderPath As String, _
    ByRef result As ExportResult _
)
    If comp.CodeModule.CountOfLines = 0 Then
        result.Skipped = result.Skipped + 1
        Exit Sub
    End If

    Dim i       As Long
    Dim content As String

    Dim objName As String
    On Error Resume Next
    objName = comp.Properties("Name").value
    On Error GoTo 0
    If objName = "" Then objName = comp.Name

    content = "' Component: " & comp.Name & "  [" & objName & "]" & vbCrLf & _
              "' Type: Document" & vbCrLf & _
              "Option Explicit" & vbCrLf & vbCrLf

    Dim startLine As Long
    startLine = 1
    If LCase(Trim(comp.CodeModule.lines(1, 1))) = "option explicit" Then
        startLine = 2
    End If

    For i = startLine To comp.CodeModule.CountOfLines
        content = content & comp.CodeModule.lines(i, 1) & vbCrLf
    Next i

    Dim fileName As String
    If objName <> comp.Name And objName <> "" Then
        fileName = comp.Name & "_" & CleanFileName(objName) & ".bas"
    Else
        fileName = comp.Name & ".bas"
    End If

    Dim filePath As String
    filePath = folderPath & fileName

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


Private Function ReadFileAnsi(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "windows-1251"
    stm.Open
    stm.LoadFromFile filePath
    ReadFileAnsi = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function


Private Sub WriteFileUtf8Bom(ByVal filePath As String, ByVal content As String)
    Dim tsText As Object
    Set tsText = CreateObject("ADODB.Stream")
    tsText.Type = 2
    tsText.Charset = "utf-8"
    tsText.Open
    tsText.WriteText content

    Dim tsBin As Object
    Set tsBin = CreateObject("ADODB.Stream")
    tsBin.Type = 1
    tsBin.Open

    tsText.Position = 0
    tsText.CopyTo tsBin

    tsBin.SaveToFile filePath, 2

    tsText.Close
    tsBin.Close
    Set tsText = Nothing
    Set tsBin = Nothing
End Sub


Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant
    Dim i   As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    CleanFileName = s
End Function


Private Sub EnsureFolder(ByVal Path As String)
    If Dir(Path, vbDirectory) = "" Then
        On Error Resume Next
        MkDir Path
        On Error GoTo 0
    End If
End Sub


