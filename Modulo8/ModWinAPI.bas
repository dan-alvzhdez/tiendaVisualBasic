Attribute VB_Name = "ModWinAPI"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
  (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Function TempPath() As String
    Dim s As String * 260, n As Long
    n = GetTempPath(260, s)
    If n > 0 Then
        TempPath = Left$(s, n)
    Else
        TempPath = App.path & "\"
    End If
End Function

Public Sub OpenWithDefault(ByVal filePath As String)
    Call ShellExecute(0&, "open", filePath, vbNullString, vbNullString, 1)
End Sub


