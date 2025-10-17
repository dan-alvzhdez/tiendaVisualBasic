VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
Private cNombre As String
Private cEmail As String
Private cPhone() As String

' ------------ Propiedades ------------
Public Property Get nombre() As String
    nombre = cNombre
End Property

Public Property Let nombre(ByVal v As String)
    cNombre = Trim$(v)
End Property

Public Property Get email() As String
    email = cEmail
End Property

Public Property Let email(ByVal v As String)
    cEmail = Trim$(v)
End Property
Public Property Get Phone() As String()
    Phone = cPhone
End Property

Public Property Let Phone(ByRef arr() As String)
    cPhone = arr
End Property


' ------------ inicializador ------
Public Sub Init(ByVal nombre As String, ByVal email As String, ByRef telefonos() As String)
    cNombre = Trim$(nombre)
    cEmail = Trim$(email)
    cPhone() = telefonos
End Sub



