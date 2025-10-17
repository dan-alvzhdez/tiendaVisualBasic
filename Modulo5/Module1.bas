Attribute VB_Name = "modUtils"
Option Explicit
' Nombre de la app
Public Const APP_NAME As String = "AgendaVB6"

' colección global (objetos CContacto)
Public Contacts As Collection

' ---------- Validaciones ----------
Public Function IsValidEmail(ByVal s As String) As Boolean
    s = Trim$(s)
    IsValidEmail = (Len(s) > 3 And InStr(1, s, "@") > 1 And InStrRev(s, ".") > InStr(1, s, "@") + 1)
End Function
' ---------- Validacion de solo digitos numericos ----------
Public Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String
    Dim out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next
    OnlyDigits = out
End Function

' ---------- Persistencia CSV con manejo de errores ----------
Public Function SaveContactsCSV(ByVal path As String) As Boolean
    On Error GoTo EH
    Dim i As Long
    Dim c As cContacto

    If Contacts Is Nothing Or Contacts.Count = 0 Then
        Err.Raise 1001, APP_NAME, "No hay contactos para exportar."
    End If

    Open path For Output As #1
    Print #1, "Nombre,Email,Telefono1,Telefono2,Telefono3"

    For Each c In Contacts
        Print #1, c.ToCSV()
    Next
    Close #1

    SaveContactsCSV = True
    Exit Function
EH:
    SaveContactsCSV = False
    If Err.Number <> 0 Then
        MsgBox "Error al exportar (" & Err.Number & "): " & Err.Description, vbCritical, APP_NAME
        On Error Resume Next: Close #1
    End If
End Function

' ---------- Utilidad de depuración ----------
Public Sub DebugDump()
    Dim c As cContacto, idx As Long
    Debug.Print "---- Dump (" & Now & ") ----"
    If Contacts Is Nothing Then
        Debug.Print "Contacts = Nothing": Exit Sub
    End If
    Debug.Print "Total: "; Contacts.Count
    For Each c In Contacts
        idx = idx + 1
        Debug.Print idx & ") " & c.nombre & " | " & c.email & " | " & Join(c.Phones, " / ")
    Next
End Sub

