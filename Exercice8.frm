VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fougère"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   6600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim angle As Double
'a1 , vert
Function a1x(x, y)
    a1x = Cos(-3 * angle) * 0.85 * x - Sin(-3 * angle) * 0.85 * y
End Function
Function a1y(x, y)
    a1y = Sin(-3 * angle) * 0.85 * x + Cos(-3 * angle) * 0.85 * y + (420 / 100) * 15
End Function
'a2, rouge
Function a2x(x, y)
    a2x = Cos(45 * angle) * 0.3 * x - Sin(45 * angle) * 0.3 * y
End Function
Function a2y(x, y)
    a2y = Sin(45 * angle) * 0.3 * x + Cos(45 * angle) * 0.3 * y + (420 / 100) * 15
End Function
'a3 , bleu
Function a3x(x, y)
    a3x = Cos(-50 * angle) * 0.3 * (-x) - Sin(-50 * angle) * 0.3 * y
End Function
   Function a3y(x, y)
    a3y = Sin(-50 * angle) * 0.3 * (-x) + Cos(-50 * angle) * 0.3 * y + (420 / 100) * 7.5
   End Function
'a4
Function a4x(x, y)
    a4x = 0.001 * x
End Function
Function a4y(x, y)
    a4y = 0.15 * y
End Function

Function xg(x)
    xg = (420 / 400) * x + 210
End Function
Function yg(y)
    yg = -420 / 400 * y + 420
End Function

Private Sub Picture1_Click()
pi = 4 * Atn(1)
'Point initial P
x = 400 * Rnd - 200: y = 400 * Rnd
Randomize

'angle
Let angle = pi / 180


For k = 1 To 100000
    'Choix de la transformation
    p = Rnd
    If 0 <= p And p < 0.79 Then
        nx = a1x(x, y)
        ny = a1y(x, y)
        col = RGB(0, 255, 0)
    ElseIf 0.79 <= p And p < 0.89 Then
        nx = a2x(x, y)
        ny = a2y(x, y)
        col = RGB(0, 0, 255)
    ElseIf 0.89 <= p And p < 0.99 Then
        nx = a3x(x, y)
        ny = a3y(x, y)
        col = RGB(255, 0, 0)
    ElseIf 0.99 <= p Then
        nx = a4x(x, y)
        ny = a4y(x, y)
        col = RGB(0, 0, 0)
    End If
    'Calcul des coordonnées de l'image P' de P
    nxg = xg(nx)
    nyg = yg(ny)
    'Dessin du point P'
    Picture1.PSet (nxg, nyg), col
    'Le point P' devient le point P
    x = nx
    y = ny
    
Next k
Label1.Caption = "FIN"
End Sub


