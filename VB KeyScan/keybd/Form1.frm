VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   7470
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "16"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "24"
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Text2.Text = Val(Text2.Text) + 1
End Sub

Private Sub Command2_Click()
Text2.Text = Val(Text2.Text) - 1
End Sub

Private Sub Command3_Click()
Text1.Text = Val(Text1.Text) + 1
End Sub

Private Sub Command4_Click()
Text1.Text = Val(Text1.Text) - 1
End Sub

Private Sub Form_Load()
    SetKeyboardHook
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RemoveKeyboardHook
End Sub

