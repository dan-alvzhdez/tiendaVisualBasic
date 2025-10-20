VERSION 5.00
Begin VB.Form frmAgenda 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTelefono 
      Height          =   375
      Index           =   2
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtTelefono 
      Height          =   375
      Index           =   1
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Salir"
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton btnExportarCSV 
      Caption         =   "&Exportar CSV"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame fraAgenda 
      Caption         =   "Agenda"
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1695
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   3735
         Begin VB.TextBox txtTelefono 
            Height          =   375
            Index           =   0
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   10
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Telefono"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ListBox lstAgenda 
         Height          =   2790
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   8175
      End
      Begin VB.CommandButton btnAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lvlEmail 
         Caption         =   "Email"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const APP_NAME = "Agenda"

Private Sub Form_Load()
    If Contacts Is Nothing Then Set Contacts = New Collection
    txtNombre.Text = ""
    txtEmail.Text = ""
    txtTelefono(0).Text = ""
    txtTelefono(1).Text = ""
    txtTelefono(2).Text = ""
End Sub

Private Sub btnAgregar_Click()
    On Error GoTo EH

    Dim tel() As String
    Dim i As Integer, l As Integer
    Dim c As cContacto
    Dim msg As String

    ReDim tel(0)
    l = -1
    For i = 0 To 2
        If Len(Trim$(txtTelefono(i).Text)) > 0 Then
            l = l + 1
            If l = 0 Then
                ReDim tel(0)
            Else
                ReDim Preserve tel(l)
            End If
            tel(l) = txtTelefono(i).Text
        End If
    Next

    Set c = New cContacto
    Call c.Init(txtNombre.Text, txtEmail.Text, tel)

    If Not c.IsValid(msg) Then
        MsgBox msg, vbExclamation, APP_NAME
        Exit Sub
    End If

    Contacts.Add c
    lstAgenda.AddItem c.nombre & " | " & c.email & " | " & IIf(UBound(tel) >= 0, Join(tel, " / "), "-")

    Debug.Print "Agregado: "; c.nombre; "  Total="; Contacts.Count
    LimpiarForm
    Exit Sub
EH:
    MsgBox "Error al agregar (" & Err.Number & "): " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub LimpiarForm()
    txtNombre.Text = ""
    txtEmail.Text = ""
    Dim i As Integer
    For i = 0 To 2
        txtTelefono(i).Text = ""
    Next
    txtNombre.SetFocus
End Sub

Private Sub txtTelefono_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelefono_Change(Index As Integer)
    Dim t As String, cur As Long
    cur = txtTelefono(Index).SelStart
    t = OnlyDigits(txtTelefono(Index).Text)
    txtTelefono(Index).Text = t
    txtTelefono(Index).SelStart = cur
End Sub

Private Sub btnExportarCSV_Click(Index As Integer)
    Dim ruta As String, ok As Boolean

    On Error Resume Next

    Dim cdl As Object
    Set cdl = CreateObject("MSComDlg.CommonDialog")

    If Not cdl Is Nothing Then
        cdl.CancelError = True
        cdl.DialogTitle = "Exportar agenda"
        cdl.Filter = "CSV (*.csv)|*.csv"
        cdl.FileName = "agenda.csv"
        cdl.ShowSave
        If Err.Number = 32755 Then Exit Sub
        ruta = cdl.FileName
        Err.Clear
    End If
    On Error GoTo 0

    If Len(ruta) = 0 Then ruta = App.path & "\agenda.csv"

    ok = SaveContactsCSV(ruta)
    If ok Then MsgBox "Exportado a: " & ruta, vbInformation, APP_NAME
End Sub

Private Sub btnClose_Click(Index As Integer)
    Unload Me
End Sub

