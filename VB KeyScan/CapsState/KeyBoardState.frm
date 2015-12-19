VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8280
      Top             =   480
   End
   Begin VB.CommandButton CommandSCROLLLock 
      Caption         =   "SCROLL Lock"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton CommandNUMLocks 
      Caption         =   "NUM Locks"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton CommandCapLocks 
      Caption         =   "CAP Locks"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub Timer1_Timer()
    Dim statusCapsLock As Integer, statusNumLock As Integer, statusScrollLock As Integer
    statusCapsLock = GetKeyState(vbKeyCapital)
    statusNumLock = GetKeyState(vbKeyNumlock)
    statusScrollLock = GetKeyState(vbKeyScrollLock)
    
    Label1 = statusCapsLock
    Label2 = statusNumLock
    Label3 = statusScrollLock
End Sub
