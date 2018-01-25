VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16215
   LinkTopic       =   "Form1"
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1081
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   3075
      Left            =   12600
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   5
      Top             =   120
      Width           =   3075
   End
   Begin VB.PictureBox Picture4 
      Height          =   3075
      Left            =   9480
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   4
      Top             =   120
      Width           =   3075
   End
   Begin VB.PictureBox Picture3 
      Height          =   3075
      Left            =   6360
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   3
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dessin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   3075
      Left            =   3240
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   120
      Width           =   3075
   End
   Begin VB.PictureBox Picture1 
      Height          =   3075
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture1.Picture = LoadPicture("matterhorn201.bmp")
End Sub

Function arg(x, y) '-pi < arg <= pi
pi = 4 * Atn(1)
If x > 0 Then
    arg = Atn(y / x)
End If
If x = 0 Then
    arg = Sgn(y) * pi / 2
End If
If x < 0 Then
    If y >= 0 Then
        arg = Atn(y / x) + pi
    Else
        arg = Atn(y / x) - pi
    End If
End If
End Function

Function h(phi)
    If Abs(Cos(phi)) > Abs(Sin(phi)) Then
         h = 100 / Abs(Cos(phi))
    Else
        h = 100 / Abs(Sin(phi))
    End If
End Function
Function toRho(x, y)
toRho = Sqr(x ^ 2 + y ^ 2)
End Function
Function toPhi(x, y)
toPhi = arg(x, y)
End Function
Function toX(rho, phi)
toX = rho * Cos(phi)
End Function
Function toY(rho, phi)
toY = rho * Sin(phi)
End Function
Function xg(x)
xg = x + 100
End Function
Function yg(y)
yg = -y + 100
End Function
Function r1(phi)
r1 = 0.5 * Cos(4 * phi) + 0.5 ' le oméga représent le nombre de feiolles
End Function
Function r2(phi)
r2 = 0.25 * -Cos(8 * phi) + 0.75
End Function
Private Sub Command1_Click()
'parcourt les pixels de départ
rayon = 100 ' ou 100 pour la consigne
For x = -100 To 100
    For y = -100 To 100
        rho = toRho(x, y)
         phi = toPhi(x, y)
         nrho = rho * rayon / h(phi) ' rayon divsé par h(phi) la distance
         c = Picture1.Point(xg(x), yg(y))
        
         nx = toX(nrho, phi)
         ny = toY(nrho, phi)
        
        Picture2.PSet (xg(nx), yg(ny)), c
        
    Next y
Next x
 
 For x = -100 To 100
    For y = -100 To 100
        phi = toPhi(x, y)
        rho = toRho(x, y)
        c = Picture2.Point(xg(toX(rho, phi)), yg(toY(rho, phi)))
        nx = rho * r1(phi) * Cos(phi)
        ny = rho * r1(phi) * Sin(phi)
        
        Picture3.PSet (xg(nx), yg(ny)), c
        
    Next y
Next x
        
        
For x = -100 To 100
    For y = -100 To 100
        phi = toPhi(x, y)
        rho = toRho(x, y)
        c = Picture2.Point(xg(toX(rho, phi)), yg(toY(rho, phi)))
        nx = rho * r2(phi) * Cos(phi)
        ny = rho * r2(phi) * Sin(phi)
        
        Picture4.PSet (xg(nx), yg(ny)), c
        
    Next y
Next x

End Sub
