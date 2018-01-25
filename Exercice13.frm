VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   2445
   ClientTop       =   3225
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3075
      Left            =   45
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   60
      Width           =   3075
   End
   Begin VB.PictureBox Picture2 
      Height          =   3075
      Left            =   3240
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
Private Sub Form_Load()
Picture1.Picture = LoadPicture("matterhorn201.bmp")
End Sub
Function k(y)
k = -0.25 * Sin((3.1415 / 100) * y) + 0.75 'attention!!! il faut diviser par 100!!!
End Function

Private Sub Picture1_Click()
For x = -100 To 100
        For y = -100 To 100
        nx = k(y) * x
        xg = nx + 100
        yg = -y + 100
        c = Picture1.Point(x + 100, -y + 100)
        Picture2.PSet (xg, yg), c
    Next y
Next x
    
End Sub
