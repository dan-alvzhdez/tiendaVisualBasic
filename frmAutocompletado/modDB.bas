Attribute VB_Name = "modDB"
' modDB.bas
Option Explicit

Public cn As ADODB.Connection

' Abre la conexión si no está abierta. Devuelve True si quedó abierta.
Public Function OpenConnection() As Boolean
    On Error GoTo EH

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then OpenConnection = True: Exit Function
    End If

    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 5
    cn.CommandTimeout = 30
    cn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=127.0.0.1;" & _
        "Initial Catalog=bancoppel;" & _
        "User ID=sa;Password=Temporal2025*;" & _
        "Persist Security Info=False;"
    cn.Open

    OpenConnection = (cn.State = adStateOpen)
    Exit Function
EH:
    OpenConnection = False
    MsgBox "Error al abrir conexión: " & Err.Number & " - " & Err.Description, vbCritical, "Conexión"
End Function

' Garantiza que la conexión esté abierta antes de ejecutar consultas
Public Function EnsureConnectionOpen() As Boolean
    On Error GoTo EH
    If cn Is Nothing Then Set cn = New ADODB.Connection
    If cn.State <> adStateOpen Then
        cn.ConnectionTimeout = 5
        cn.CommandTimeout = 30
        cn.ConnectionString = _
            "Provider=SQLOLEDB;Data Source=127.0.0.1;Initial Catalog=bancoppel;User ID=sa;Password=Temporal2025*;Persist Security Info=False;"
        cn.Open
    End If
    EnsureConnectionOpen = (cn.State = adStateOpen)
    Exit Function
EH:
    EnsureConnectionOpen = False
End Function

' Cierra y limpia la conexión
Public Sub CloseConnection()
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
End Sub

