VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Triangle de Sierpinski"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Sierpinski"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6360
      Left            =   120
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   120
      Width           =   6360
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
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function a1x(x, y)
a1x = 0.5 * x
End Function
Function a1y(x, y)
 a1y = 0.5 * y + 200
End Function
Function a2x(x, y)
a2x = 0.5 * x - 100 ' parce que c'est le centre qu'on d�place. donc pas 200!!s
End Function
Function a2y(x, y)
a2y = 0.5 * y
End Function
Function a3x(x, y)
a3x = 0.5 * x + 100
End Function
Function a3y(x, y)
a3y = 0.5 * y
End Function
Function xg(x)
    xg = (420 / 400) * x + 210
End Function
Function yg(y)
    yg = -420 / 400 * y + 420
End Function

Private Sub Command1_Click()
'base du triangle 400 pixels, hauteur du triangle 200 pixels
Picture1.Cls
'origine au point (200,400) de la picturebox
'choix d'un premier point au hasard
x = 400 * Rnd - 200
y = 400 * Rnd
'dessin de points
For j = 1 To 200000
    'tirage d'un nombre au hasard
    p = Rnd
    'calcul de l'image du point
    If 0 <= p And p < (1 / 3) Then
        nx = a1x(x, y)
        ny = a1y(x, y)
    ElseIf (1 / 3) <= p And p < (2 / 3) Then
        nx = a2x(x, y)
        ny = a2y(x, y)
    ElseIf (2 / 3) <= p And p < 1 Then
        nx = a3x(x, y)
        ny = a3y(x, y)
    End If
    'dessin du point
    nxg = xg(nx)
    nyg = yg(ny)
    'un bord de 10 pixels est pr�vu pour faire joli
    Picture1.PSet (nxg, nyg)
    'nouveau point devient ancien point
    x = nx
    y = ny
Next j
'pour savoir quand le programme est termin�
Label1.Caption = "Fin"
End Sub
