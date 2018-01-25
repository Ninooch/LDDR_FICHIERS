VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rotation"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2400
      TabIndex        =   0
      Top             =   3240
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dessin avec réciproque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   3165
   End
   Begin VB.PictureBox Picture2 
      Height          =   4575
      Left            =   3405
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   2
      Top             =   75
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3075
      Left            =   75
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   75
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Angle en degrés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   165
      TabIndex        =   4
      Top             =   3240
      Width           =   2220
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

Private Sub Command1_Click()
Picture2.Cls
Let angle = Val(Text1.Text) * 3.1415 / 180

'Transformation directe

cx = (200 / 2) ' le centre de l'image de 200 par 200
cy = (200 / 2)


' ce qu'il faut faire : déplacer le centre (100,100) en haut à gauche de l'image
' car c'est son origine et que les formules ne permettent de tourner une
' image qu'autour de l'origine!!
' ensuite : effectuer la rotation
' ensuite : déplacer le centre de l'image au millieu du canvas de 300 par 300
For x = 0 To 200
    For y = 0 To 200
        Let c = Picture1.Point(x, y)
        Let nx = Cos(angle) * (x - cx) - Sin(angle) * (y - cy) + 150 '150 est le centre de l'image de 300 par 300 !
        Let ny = Sin(angle) * (x - cx) + Cos(angle) * (y - cy) + 150
        Picture2.PSet (nx, ny), c
    Next y
Next x



End Sub

Private Sub Command2_Click()
Let angle = Val(Text1.Text) * 3.1415 / 180

'Coloriage de la PictureBox d'arrivée au moyen de la transformation réciproque
'La réciproque est une rotation d'angle -a dans système usuel,
'donc d'angle a dans le système des PictureBoxes
Picture2.Cls
For nx = 0 To 300
    For ny = 0 To 300
    Let x = Cos(angle) * (nx - 150) + Sin(angle) * (ny - 150) + 100
    Let y = -Sin(angle) * (nx - 150) + Cos(angle) * (ny - 150) + 100
    Let c = Picture1.Point(x, y)
    
    Picture2.PSet (nx, ny), c
    Next ny
Next nx
    
End Sub
