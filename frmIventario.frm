VERSION 5.00
Begin VB.Form formus 
   Caption         =   "Inventario"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraProductos 
      Caption         =   "Registro de productos"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.CommandButton btnInventario 
         Caption         =   "Ver inventario"
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4560
         Width           =   1575
      End
      Begin VB.ComboBox cboCategoria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmIventario.frx":0000
         Left            =   1560
         List            =   "frmIventario.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton btnRegistrar 
         Caption         =   "&Guardar"
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible a la venta"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optGranel 
         Caption         =   "Venta a granel"
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   3360
         Width           =   1695
      End
      Begin VB.OptionButton optIndividual 
         Caption         =   "Venta individual"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   615
         Left            =   1800
         TabIndex        =   6
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtPrecioVenta 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtProducto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3840
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label3 
         Caption         =   "Categoría"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Precio de venta"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.Menu mnuProductos 
      Caption         =   "&Registro de productos"
      Index           =   0
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "formus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim inventario As Collection

Private Sub Form_Load()

    Set inventario = New Collection

    cboCategoria.AddItem "Embutidos"
    cboCategoria.AddItem "Lácteos"
    cboCategoria.AddItem "Quesos"
    cboCategoria.AddItem "Refrescos"
    cboCategoria.AddItem "Sabritas"

    chkVisible.Value = 1
End Sub

Private Sub btnLimpiar_Click()
    funcionLimpiar
End Sub

Private Sub btnRegistrar_Click()
    Dim producto As String
    Dim precio As Double
    Dim cantidad As Integer
    Dim categoria As String
    Dim tipoVenta As String
    Dim visible As Boolean
    Dim registro As String

    producto = txtProducto.Text
    precio = Val(txtPrecioVenta.Text)
    cantidad = Val(txtCantidad.Text)
    categoria = cboCategoria.Text
    tipoVenta = IIf(optIndividual.Value, "Individual", "Granel")
    visible = (chkVisible.Value = 1)

    registro = producto & " | $" & precio & " | Cant: " & cantidad & " | " & categoria & " | " & tipoVenta & " | " & IIf(visible, "Visible", "Oculto")
    inventario.Add registro
    MsgBox "Producto guardado correctamente.", vbInformation
    
    funcionLimpiar

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        ' Si ya hay un punto, bloquear otro
        If InStr(txtCantidad.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' válido
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub funcionLimpiar()
    txtProducto.Text = ""         ' Producto
    txtPrecioVenta.Text = ""         ' Precio de venta
    txtCantidad.Text = ""         ' Cantidad
    cboCategoria.ListIndex = -1   ' Deseleccionar categoría
    optIndividual.Value = False   ' Venta individual
    optGranel.Value = False   ' Venta a granel
    chkVisible.Value = 0        ' Volver a activar el checkbox
End Sub

Private Sub txtPrecioVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        ' Si ya hay un punto, bloquear otro
        If InStr(txtPrecioVenta.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' válido
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub btnInventario_Click()
    If inventario.Count = 0 Then
        MsgBox "No hay productos registrados.", vbExclamation, "Inventario vacío"
    Else
        frmListadoProductos.lstRegistrados.Clear
        Dim item As Variant
        For Each item In inventario
            frmListadoProductos.lstRegistrados.AddItem item
        Next
        frmListadoProductos.Show
    End If
End Sub
