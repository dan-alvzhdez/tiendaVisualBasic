Attribute VB_Name = "modAutocomplete"
' modAutocomplete.bas
Option Explicit

' Escapa comillas para usar con LIKE
Public Function SqlLikeSafe(ByVal s As String) As String
    SqlLikeSafe = Replace(s, "'", "''")
End Function

' Carga el combo de Actividades
Public Sub LoadActividades(ByRef cbo As ComboBox)
    If Not EnsureConnectionOpen() Then
        MsgBox "Conexión cerrada. No se pudo cargar Actividades.", vbExclamation
        Exit Sub
    End If

    Dim rs As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = "SELECT idActividad, nombre FROM dbo.TB_Actividad ORDER BY nombre"
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    cbo.Clear
    Do While Not rs.EOF
        cbo.AddItem rs!nombre
        cbo.ItemData(cbo.NewIndex) = rs!idActividad
        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing
End Sub

' Carga SubActividades por actividad con filtro "contiene"
Public Sub LoadSubActividades(ByRef cbo As ComboBox, ByVal idActividad As Long, ByVal filtro As String)
    If Not EnsureConnectionOpen() Then
        MsgBox "Conexión cerrada. No se pudo cargar SubActividades.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    isFillingSub = True

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim textoEscrito As String

    filtro = Trim$(filtro)

    sql = "SELECT idSubActividad, nombre FROM dbo.TB_SubActividad WHERE idActividad=" & idActividad
    If Len(filtro) > 0 Then
        sql = sql & " AND nombre LIKE '%" & SqlLikeSafe(filtro) & "%'"
        ' Si quieres ignorar acentos (si tu collation no lo hace):
        ' sql = sql & " AND nombre COLLATE Latin1_General_CI_AI LIKE '%" & SqlLikeSafe(filtro) & "%'"
    End If
    sql = sql & " ORDER BY nombre"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    textoEscrito = cbo.Text

    cbo.Clear
    Do While Not rs.EOF
        cbo.AddItem rs!nombre
        cbo.ItemData(cbo.NewIndex) = rs!idSubActividad
        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing

    ' Restaurar texto sin provocar recursión
    If cbo.Text <> textoEscrito Then cbo.Text = textoEscrito
    cbo.SelStart = Len(cbo.Text)
    cbo.SelLength = 0

FIN:
    isFillingSub = False
    Exit Sub
EH:
    isFillingSub = False
    MsgBox "Error cargando subactividades: " & Err.Description, vbExclamation
    Resume FIN
End Sub

' Devuelve el idSubActividad si el texto coincide exactamente con un item, o -1
Public Function GetSelectedSubActividadId(ByRef cbo As ComboBox) As Long
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), cbo.Text, vbTextCompare) = 0 Then
            GetSelectedSubActividadId = cbo.ItemData(i)
            Exit Function
        End If
    Next i
    GetSelectedSubActividadId = -1
End Function

