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
      Top             =   120
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
Private Sub Picture1_Click()
pi = 4 * Atn(1)
'Point initial P
x = 400 * Rnd - 200: y = 400 * Rnd
Randomize

' changement de coordonnée :
Let mx = 420 / 400
Let hx = 210


For k = 1 To 100000
    'Choix de la transformation
    p = Rnd
    'Calcul des coordonnées de l'image P' de P
    
    'Dessin du point P'
    
    'Le point P' devient le point P
    
Next k
Label1.Caption = "FIN"
End Sub
