VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   2505
   ClientTop       =   2685
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3360
      Left            =   120
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   3
      Top             =   120
      Width           =   3360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<----------"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "---------->"
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
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   3360
      Left            =   3600
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   0
      Top             =   120
      Width           =   3360
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
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Command1_Click()
For x = 0 To 219
    For y = 0 To 219 ' important qu'elle aille à 219
        
        c = Picture1.Point(x, y)
        
        If x / 2 = Int(x / 2) Then
            xp = True
        Else
            xp = False
        End If
        
        If y / 2 = Int(y / 2) Then
            yp = True
        Else
            yp = False
        End If
        
        
        'x est pair , y est pair --> haut à gauche
        If xp And yp Then
            nx = x / 2
            ny = y / 2
            Picture2.PSet (nx, ny), c
        
        'x est impair, y est pair --> haut à droite
        ElseIf yp And (Not xp) Then
            nx = Int(x / 2) + 110
            ny = y / 2
            Picture2.PSet (nx, ny), c
            
        'x est impair, y est impair --> bas à droite
        ElseIf (Not xp) And (Not yp) Then
            nx = Int(x / 2) + 110
            ny = Int(y / 2) + 110
            Picture2.PSet (nx, ny), c
            
        ' x est pair, y est impair --> bas à gauche
        ElseIf xp And (Not yp) Then
            nx = x / 2
            ny = Int(y / 2) + 110
            Picture2.PSet (nx, ny), c
        End If
    Next y
Next x

n = n + 1
Label1.Caption = n
End Sub

Private Sub Command2_Click()
For x = 0 To 219
    For y = 0 To 219 ' important qu'elle aille à 219
        
        c = Picture2.Point(x, y)
        
        If x / 2 = Int(x / 2) Then
            xp = True
        Else
            xp = False
        End If
        
        If y / 2 = Int(y / 2) Then
            yp = True
        Else
            yp = False
        End If
        
        
        'x est pair , y est pair --> haut à gauche
        If xp And yp Then
            nx = x / 2
            ny = y / 2
            Picture1.PSet (nx, ny), c ' si on laisse picture 1 , ça fait une fractale
        
        'x est impair, y est pair --> haut à droite
        ElseIf yp And (Not xp) Then
            nx = Int(x / 2) + 110
            ny = y / 2
            Picture1.PSet (nx, ny), c
            
        'x est impair, y est impair --> bas à droite
        ElseIf (Not xp) And (Not yp) Then
            nx = Int(x / 2) + 110
            ny = Int(y / 2) + 110
            Picture1.PSet (nx, ny), c
            
        ' x est pair, y est impair --> bas à gauche
        ElseIf xp And (Not yp) Then
            nx = x / 2
            ny = Int(y / 2) + 110
            Picture1.PSet (nx, ny), c
        End If
    Next y
Next x

n = n + 1
Label1.Caption = n

End Sub

Private Sub Form_Load()
n = 0
Picture1.Picture = LoadPicture("matterhorn220.bmp")
End Sub
