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
Option Explicit

Private mNombre As String
Private mEmail As String
' arreglo dinámico de teléfonos
Private mPhones() As String

' --------- Propiedades ---------
Public Property Get nombre() As String
    nombre = mNombre
End Property
Public Property Let nombre(ByVal v As String)
    mNombre = Trim$(v)
End Property

Public Property Get email() As String
    email = mEmail
End Property
Public Property Let email(ByVal v As String)
    mEmail = Trim$(v)
End Property

Public Property Get Phones() As String()
    Phones = mPhones
End Property
Public Property Let Phones(ByRef arr() As String)
    mPhones = arr
End Property

' --------- Inicializador cómodo ---------
Public Sub Init(ByVal nombre As String, ByVal email As String, ByRef telefonos() As String)
    mNombre = Trim$(nombre)
    mEmail = Trim$(email)
    mPhones = telefonos
End Sub

' --------- Validación de negocio ---------
Public Function IsValid(ByRef mensaje As String) As Boolean
    If Len(mNombre) = 0 Then
        mensaje = "El nombre es obligatorio.": Exit Function
    End If
    If Not IsValidEmail(mEmail) Then
        mensaje = "Email inválido.": Exit Function
    End If
    IsValid = True
End Function

' --------- Serialización a CSV ---------
Public Function ToCSV() As String
    Dim p1 As String, p2 As String, p3 As String
    If UBound(mPhones) >= 0 Then p1 = mPhones(0)
    If UBound(mPhones) >= 1 Then p2 = mPhones(1)
    If UBound(mPhones) >= 2 Then p3 = mPhones(2)
    ToCSV = """" & mNombre & """,""" & mEmail & """,""" & p1 & """,""" & p2 & """,""" & p3 & """"
End Function

