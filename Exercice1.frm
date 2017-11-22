VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Exercice 2"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Dessin"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   240
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   108
      Width           =   3075
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   3360
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   1
      Top             =   120
      Width           =   6075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For x = 0 To 200
     For y = 0 To 200
        c = Picture1.Point(x, y)
        'image en haut à gauche
        nxhg = 200 - x + 200
        nyhg = 200 - y
        Picture2.PSet (nxhg, nyhg), c
        
        'image en bas à gauche
        nxbg = 200 - x + 200
        nybg = 200 + y
        Picture2.PSet (nxbg, nybg), c
        
        ' image en bas à droite
        nxbd = x
        nybd = 200 + y
        Picture2.PSet (nxbd, nybd), c
        
        'image en haut à droite
        nxhd = x
        nyhd = 200 - y
        Picture2.PSet (nxhd, nyhd), c
        
    Next y
Next x
        
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture("matterhorn201.bmp")
End Sub
