VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Vagues"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   3075
      Left            =   3360
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
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label Label4 
      Caption         =   "Amplitude variable, parcourt l'image d'arrivée"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   7680
      Width           =   3255
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
Function k(x)
k = -35 * Sin((3.1415 / 49) * x) ' B est nul car c'est déjà le y . ( on ajoute des pixels à chaques position y )
End Function

Private Sub Picture1_Click()
    For x = -100 To 100
        For y = -100 To 100
             ny = y + k(x) ' il ne faut pas diviser par 100, on additionne des pixels!!!
             xg = x + 100
             yg = -ny + 100
             c = Picture1.Point(x + 100, -y + 100)
             Picture2.PSet (xg, yg), c
        Next y
    Next x
    
             
End Sub

