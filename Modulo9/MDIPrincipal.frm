VERSION 5.00
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11820
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuProductos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ventana"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "TileVertical"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MDIPrincipal (MDIForm)

Private Sub mnuAcercaDe_Click()
    MsgBox "Ejemplo MDI - Módulo 9: Aplicaciones MDI y depuración avanzada" & vbCrLf & _
           "Autor: TechTigerFTW" & vbCrLf & "Fecha: " & Date, vbInformation, "Acerca de"
End Sub

' ---------- Eventos del menú ----------
Private Sub mnuClientes_Click()
    On Error GoTo EH
    ' Si ya existe la instancia, simplemente mostrar otra.
    Dim f As frmClientes
    Set f = New frmClientes
    f.Show
    Debug.Print "Se abrió formulario: Clientes - " & Now
    Exit Sub
EH:
    Debug.Print "Error en mnuClientes_Click: " & Err.Number & " - " & Err.Description
End Sub

Private Sub mnuProductos_Click()
    On Error GoTo EH
    Dim f As frmProductos
    Set f = New frmProductos
    f.Show
    Debug.Print "Se abrió formulario: Productos - " & Now
    Exit Sub
EH:
    Debug.Print "Error en mnuProductos_Click: " & Err.Number & " - " & Err.Description
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub


'----------------- QueryUnload Confirmar Cierre --------------------

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim respuesta As VbMsgBoxResult
    
    respuesta = MsgBox("Seguro que desea salir de la aplicación?", vbYesNo + vbQuestion, "Confirmar para salir")
    
    If respuesta = vbNo Then
        Cancel = True
        Debug.Print "Cancelado por el usuario - " & Now
    Else
        Debug.Print "Aplicacion cerrada por el usuario - " & Now
    End If
End Sub
