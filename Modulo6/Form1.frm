VERSION 5.00
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const PIXEL = 15 ' Aproximadamente 1 pixel = 15 twips

    ' Crear un Shape dinámicamente
    Dim miShape As Shape
    Set miShape = Me.Controls.Add("VB.Shape", "miShape" & Me.Controls.Count)
    With miShape
        .Shape = vbShapeRectangle
        .Left = X
        .Top = Y
        .Width = 300 * PIXEL ' ~300 px
        .Height = 150 * PIXEL ' ~150 px
        .BorderColor = vbRed
        .BorderWidth = 2
        .Visible = True
    End With

    ' Crear una línea dinámica
    Dim miLinea As Line
    Set miLinea = Me.Controls.Add("VB.Line", "miLinea" & Me.Controls.Count)
    With miLinea
        .X1 = X
        .Y1 = Y
        .X2 = X + 300 * PIXEL
        .Y2 = Y + 150 * PIXEL
        .BorderColor = vbBlue
        .BorderWidth = 2
        .Visible = True
    End With
End Sub
