VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   StartUpPosition =   3  'Windows Default
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
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   3135
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   4575
      Left            =   3330
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Double, px(2) As Double, py(2) As Double

Private Sub Form_Load()
Let n = -1
Picture1.Picture = LoadPicture("matterhorn201.bmp")
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
n = n + 1 'n mémorise le nombre de points
px(n) = X: py(n) = Y 'image des sommets, l'origine est le point (px(0),py(0))
Picture2.Circle (X, Y), 3
End Sub

Private Sub Command1_Click()
Picture2.Cls
For k = 0 To 2
    Picture2.Circle (px(k), py(k)), 3
Next k
End Sub

Private Sub Command2_Click()
Picture2.Cls
For k = 0 To 2
    Picture2.Circle (px(k), py(k)), 3
Next k
End Sub

