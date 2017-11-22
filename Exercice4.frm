VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   540
   ClientTop       =   870
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   630
   Begin VB.CommandButton Command3 
      Caption         =   "Dessin avec fonction réciproque"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dessin demi-pixel par demi-pixel"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dessin pixel par pixel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   2985
   End
   Begin VB.PictureBox Picture2 
      Height          =   4575
      Left            =   3240
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   1
      Top             =   120
      Width           =   6075
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture1.Picture = LoadPicture("Gauguin201.bmp")
End Sub
Private Sub Command1_Click() 'transformation pixel par pixel de Picture1 vers Picture2
Picture2.Cls
For x = 0 To 200
    For y = 0 To 200
        c = Picture1.Point(x, y)
        nx = x * 400 / 200
        ny = y * 300 / 200
        Picture2.PSet (nx, ny), c
    Next y
Next x
End Sub

Private Sub Command2_Click() 'transformation demi-pixel par demi-pixel de Picture1 vers Picture2
Picture2.Cls
For x = 0 To 200 Step 0.5
    For y = 0 To 200 Step 0.5
        c = Picture1.Point(x, y)
        nx = x * 400 / 200
        ny = y * 300 / 200
        Picture2.PSet (nx, ny), c
    Next y
Next x

End Sub

Private Sub Command3_Click() 'transformation réciproque, de Picture2 vers Picture1
Picture2.Cls
For nx = 0 To 400
    For ny = 0 To 300
    
        x = nx * 200 / 400
        y = ny * 200 / 300
        
        c = Picture1.Point(x, y)
        
        Picture2.PSet (nx, ny), c
    Next ny
Next nx

End Sub

