VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Triangles"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   3075
      Left            =   6360
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   2
      Top             =   120
      Width           =   3075
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
      AutoSize        =   -1  'True
      Height          =   3075
      Left            =   45
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
Function xg(x)
xg = x + 100
End Function
Function yg(y)
yg = -y + 100
End Function
Function kx(x)
kx = (1 / 200) * x + 1 / 2
End Function
Function ky(y)
ky = -(1 / 200) * y + 1 / 2
End Function

Private Sub Picture2_Click()
For x = -100 To 100
    For y = -100 To 100
        Let ny = kx(x) * y
        Let c = Picture1.Point(xg(x), yg(y))
        Picture2.PSet (xg(x), yg(ny)), c
    Next y
Next x

End Sub

Private Sub Picture3_Click()

For y = -100 To 100
    For x = -100 To 100
        Let nx = ky(y) * x
        Let c = Picture1.Point(xg(x), yg(y))
        Picture3.PSet (xg(nx), yg(y)), c
    Next x
Next y

End Sub
