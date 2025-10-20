VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   1335
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2355
      _Version        =   393216
      Format          =   142213121
      CurrentDate     =   45950
   End
   Begin ComCtl2.Animation Animation2 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   327681
      AutoPlay        =   -1  'True
      FullWidth       =   65
      FullHeight      =   25
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      FullWidth       =   81
      FullHeight      =   25
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   480
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000001&
      DrawMode        =   1  'Blackness
      Height          =   2775
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrastrando As Boolean
    Dim offsetX As Integer
    Dim offsetY As Integer
    
Private Sub shp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    arrastrando = True
    offsetX = X
    offsetY = Y

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If arrastrando Then
        shp.Left = X - offsetX
        shp.Top = Y - offsetY
    End If
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    arrastrando = False
End Sub


Private Sub Timer1_Timer()
    Dim r As Integer, g As Integer, b As Integer
    r = Int(Rnd * 256)
    g = Int(Rnd * 256)
    b = Int(Rnd * 256)
    
    shp.FillColor = RGB(r, g, b)
    
    
End Sub
