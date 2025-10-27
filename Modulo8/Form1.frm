VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtContenido 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton cmdLimpiarOLE 
      Caption         =   "Limpiar OLE"
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdGuardarOLE 
      Caption         =   "Guardar OLE"
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdEditarOLE 
      Caption         =   "Editar OLE"
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdInsertarOLE 
      Caption         =   "Insertar OLE"
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "Exportar a Excel"
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbrirDefauilt 
      Caption         =   "Abrir con App"
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Archivo"
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopiarBackup 
      Caption         =   "Copia (FileCopy)"
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "Anexar a"
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGuardarComo 
      Caption         =   "Guardar Como"
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   10440
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OLE oleDoc 
      Class           =   "Word.OpenDocumentText.12"
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "Form1.frx":0006
      TabIndex        =   0
      Top             =   3120
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OLEVERB_OPEN As Long = -1
Private currentFile As String

Private Sub Form_Load()
    currentFile = ""
    Me.Caption = "Bloc de notas + OLE (Demo Módulo 8)"
End Sub

' ======== NUEVO/ABRIR/GUARDAR/ANEXAR ========

Private Sub cmdNuevo_Click(Index As Integer)
    txtContenido.Text = ""
    currentFile = ""
    Me.Caption = "Nuevo archivo - Bloc de notas + OLE"
End Sub

Private Sub cmdAbrir_Click(Index As Integer)
    On Error GoTo EH
    With cdlg
        .CancelError = True
        .Filter = "Texto (*.txt)|*.txt|Todos (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.path
        .ShowOpen
        currentFile = .FileName
    End With
    txtContenido.Text = ReadAllText(currentFile)
    Me.Caption = "Editando: " & currentFile
    Exit Sub
EH:
    If Err.Number <> 32755 And Err.Number <> 0 Then
        MsgBox "Error al abrir: " & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdGuardar_Click(Index As Integer)
    If LenB(currentFile) = 0 Then
        cmdGuardar_Click
        Exit Sub
    End If
    On Error GoTo EH
    WriteAllText currentFile, txtContenido.Text, False
    MsgBox "Archivo guardado.", vbInformation
    Exit Sub
EH:
    MsgBox "Error al guardar: " & Err.Description, vbCritical
End Sub

Private Sub cmdGuardarComo_Click(Index As Integer)
    On Error GoTo EH
    With cdlg
        .CancelError = True
        .Filter = "Texto (*.txt)|*.txt|Todos (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.path
        .ShowSave
        currentFile = .FileName
    End With
    WriteAllText currentFile, txtContenido.Text, False
    Me.Caption = "Editando: " & currentFile
    MsgBox "Guardado como: " & currentFile, vbInformation
    Exit Sub
EH:
    If Err.Number <> 32755 And Err.Number <> 0 Then
        MsgBox "Error en Guardar como: " & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdAppend_Click(Index As Integer)
    On Error GoTo EH
    With cdlg
        .CancelError = True
        .Filter = "Texto (*.txt)|*.txt|Todos (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.path
        .ShowOpen
    End With
    WriteAllText cdlg.FileName, vbCrLf & txtContenido.Text, True
    MsgBox "Contenido anexado a: " & cdlg.FileName, vbInformation
    Exit Sub
EH:
    If Err.Number <> 32755 And Err.Number <> 0 Then
        MsgBox "Error al anexar: " & Err.Description, vbCritical
    End If
End Sub

' ======== COPIAR/ELIMINAR/ABRIR CON APP (API) ========

Private Sub cmdCopiarBackup_Click(Index As Integer)
    On Error GoTo EH
    If LenB(currentFile) = 0 Then
        MsgBox "Guarda o abre un archivo primero.", vbExclamation
        Exit Sub
    End If
    Dim backupFolder As String
    backupFolder = TempPath() & "VB6DemoBackup\"
    EnsureFolder backupFolder

    Dim backupName As String
    backupName = backupFolder & Format$(Now, "yyyymmdd_hhnnss") & "_" & Dir$(currentFile)

    If SafeFileCopy(currentFile, backupName, False) Then
        MsgBox "Copia realizada en: " & backupName, vbInformation
    Else
        MsgBox "No se pudo copiar.", vbExclamation
    End If
    Exit Sub
EH:
    MsgBox "Error en copia: " & Err.Description, vbCritical
End Sub

Private Sub cmdEliminar_Click(Index As Integer)
    On Error GoTo EH
    If LenB(currentFile) = 0 Then
        MsgBox "No hay archivo actual.", vbExclamation
        Exit Sub
    End If
    If MsgBox("¿Eliminar " & currentFile & "?", vbQuestion + vbYesNo) = vbYes Then
        Kill currentFile
        txtContenido.Text = ""
        Me.Caption = "Archivo eliminado"
        currentFile = ""
    End If
    Exit Sub
EH:
    MsgBox "Error al eliminar: " & Err.Description, vbCritical
End Sub

Private Sub cmdAbrirDefault_Click(Index As Integer)
    If LenB(currentFile) = 0 Then
        MsgBox "No hay archivo para abrir con la aplicación predeterminada.", vbExclamation
        Exit Sub
    End If
    OpenWithDefault currentFile
End Sub

' ======== AUTOMATIZACIÓN OLE ? EXPORTAR A EXCEL ========

Private Sub cmdExportarExcel_Click(Index As Integer)
    On Error GoTo EH
    If LenB(txtContenido.Text) = 0 Then
        MsgBox "No hay contenido para exportar.", vbExclamation
        Exit Sub
    End If

    With cdlg
        .CancelError = True
        .Filter = "Libro de Excel 97-2003 (*.xls)|*.xls|Todos (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.path
        .ShowSave
    End With

    Dim xlApp As Object, wb As Object, ws As Object
    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Worksheets(1)

    Dim lines() As String, i As Long
    lines = Split(txtContenido.Text, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        ws.Cells(i + 1, 1).Value = lines(i)
    Next i
    ws.Columns("A:A").EntireColumn.AutoFit

    wb.SaveAs cdlg.FileName
    wb.Close False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing

    MsgBox "Exportado a Excel: " & cdlg.FileName, vbInformation
    Exit Sub
EH:
    If Err.Number = 32755 Then Exit Sub
    MsgBox "Error al exportar a Excel: " & Err.Description, vbCritical
End Sub

' ======== OLE CONTAINER: INSERTAR/EDITAR/GUARDAR/LIMPIAR ========

Private Sub cmdInsertarOLE_Click(Index As Integer)
    On Error GoTo EH
    oleDoc.InsertObjDlg     ' Nuevo (Word/Excel/etc.) o desde archivo (PDF/imagen/etc.)
    Exit Sub
EH:
    MsgBox "Error al insertar OLE: " & Err.Description, vbCritical
End Sub

Private Sub cmdEditarOLE_Click(Index As Integer)
    On Error GoTo EH
    If oleDoc.ObjType = 0 Then
        MsgBox "No hay objeto OLE insertado.", vbExclamation
        Exit Sub
    End If
    oleDoc.DoVerb OLEVERB_OPEN ' Edición in-place si aplica
    Exit Sub
EH:
    MsgBox "Error al editar OLE: " & Err.Description, vbCritical
End Sub

Private Sub cmdGuardarOLE_Click(Index As Integer)
    On Error GoTo EH
    If oleDoc.ObjType = 0 Then
        MsgBox "No hay objeto OLE para guardar.", vbExclamation
        Exit Sub
    End If
    With cdlg
        .CancelError = True
        .Filter = "Objeto OLE (*.ole)|*.ole|Todos (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.path
        .ShowSave
        oleDoc.SaveToFile .FileName ' Persistencia binaria
    End With
    MsgBox "Objeto OLE guardado en: " & cdlg.FileName, vbInformation
    Exit Sub
EH:
    If Err.Number <> 32755 And Err.Number <> 0 Then
        MsgBox "Error al guardar OLE: " & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdLimpiarOLE_Click(Index As Integer)
    On Error Resume Next
    oleDoc.Close
    ' Reemplaza por un objeto "vacío" tipo Package para limpiar visualmente
    oleDoc.CreateEmbed "Package", ""
End Sub

