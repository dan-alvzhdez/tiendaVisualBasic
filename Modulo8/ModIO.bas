Attribute VB_Name = "ModIO"
Option Explicit

Public Function ReadAllText(ByVal filePath As String) As String
    On Error GoTo EH
    Dim f As Integer: f = FreeFile
    Dim s As String
    Open filePath For Binary As #f
    s = Input$(LOF(f), #f)
    Close #f
    ReadAllText = s
    Exit Function
EH:
    If f <> 0 Then Close #f
    Err.Raise Err.Number, "ReadAllText", Err.Description
End Function

Public Sub WriteAllText(ByVal filePath As String, ByVal content As String, Optional ByVal appendMode As Boolean = False)
    On Error GoTo EH
    Dim f As Integer: f = FreeFile
    If appendMode Then
        Open filePath For Append As #f
        Print #f, content
    Else
        Open filePath For Output As #f
        Print #f, content
    End If
    Close #f
    Exit Sub
EH:
    If f <> 0 Then Close #f
    Err.Raise Err.Number, "WriteAllText", Err.Description
End Sub

Public Function SafeFileCopy(ByVal src As String, ByVal dst As String, Optional ByVal failIfExists As Boolean = False) As Boolean
    On Error GoTo VBfallback
    Dim ok As Long
    ok = CopyFile(src, dst, IIf(failIfExists, 1, 0))
    SafeFileCopy = (ok <> 0)
    Exit Function
VBfallback:
    On Error GoTo 0
    FileCopy src, dst
    SafeFileCopy = True
End Function

Public Function EnsureFolder(ByVal path As String) As Boolean
    On Error Resume Next
    MkDir path
    EnsureFolder = (Err.Number = 0 Or Dir$(path, vbDirectory) <> "")
    Err.Clear
End Function

