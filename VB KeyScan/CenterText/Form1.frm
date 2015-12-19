VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Center Text"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Cls
With Picture1
    .CurrentX = (.ScaleWidth / 2) - (.TextWidth("Centered Text") / 2)
    .CurrentY = (.ScaleHeight / 2) - (.TextHeight("Centered Text") / 2)
    Picture1.Print "Centered Text"
End With
End Sub

Private Sub Form_Load()
Me.Show
Picture1.Print "Centered Text"
End Sub
