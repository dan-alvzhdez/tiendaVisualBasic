VERSION 5.00
Begin VB.Form frmListadoProductos 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRegistrados 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Productos registrados"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmListadoProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
