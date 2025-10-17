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
   Begin VB.TextBox txtTelefonoTres 
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtTelefonoDos 
      Height          =   375
      Index           =   1
      Left            =   6480
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
         Begin VB.TextBox txtTelefonoUno 
            Height          =   375
            Index           =   0
            Left            =   1200
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
         Index           =   0
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
Private Sub Label1_Click()

End Sub
