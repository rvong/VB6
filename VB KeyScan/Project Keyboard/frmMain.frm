VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   Caption         =   "KBD Sentinel"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10560
      TabIndex        =   115
      Top             =   2760
      Width           =   975
   End
   Begin VB.PictureBox contMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   9225
      TabIndex        =   112
      Top             =   360
      Width           =   9255
      Begin VB.TextBox txtRollover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   7935
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   114
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.ListBox lstLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   120
      TabIndex        =   108
      Top             =   2880
      Width           =   9255
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   106
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   107
      Top             =   6840
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   104
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   106
      Top             =   6840
      Width           =   735
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   102
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   105
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   101
      Left            =   7200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   104
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   100
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   103
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   99
      Left            =   7920
      ScaleHeight     =   705
      ScaleWidth      =   345
      TabIndex        =   102
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   98
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   101
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   97
      Left            =   7200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   100
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   96
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   99
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   94
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   98
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   93
      Left            =   7200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   97
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   92
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   96
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   91
      Left            =   7920
      ScaleHeight     =   705
      ScaleWidth      =   345
      TabIndex        =   95
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   90
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   94
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   89
      Left            =   7200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   93
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   85
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   92
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   83
      Left            =   7920
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   91
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   88
      Left            =   6360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   90
      Top             =   6840
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   87
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   89
      Top             =   6840
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   86
      Left            =   5640
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   88
      Top             =   6840
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   84
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   87
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   82
      Left            =   6360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   86
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   81
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   85
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   80
      Left            =   5640
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   84
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   79
      Left            =   6360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   83
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   78
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   82
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   77
      Left            =   5640
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   81
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   76
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   80
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   75
      Left            =   5040
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   79
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   74
      Left            =   3600
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   78
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   73
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   77
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   72
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   76
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   71
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   2025
      TabIndex        =   75
      Top             =   6840
      Width           =   2055
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   70
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   74
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   69
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   73
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   68
      Left            =   4200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   72
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   67
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   71
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   66
      Left            =   3480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   70
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   65
      Left            =   3840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   69
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   64
      Left            =   2760
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   68
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   63
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   67
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   62
      Left            =   2040
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   66
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   61
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   65
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   60
      Left            =   1320
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   64
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   59
      Left            =   1680
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   63
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   58
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   825
      TabIndex        =   62
      Top             =   6480
      Width           =   855
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   56
      Left            =   960
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   61
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   57
      Left            =   4680
      ScaleHeight     =   345
      ScaleWidth      =   825
      TabIndex        =   60
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   55
      Left            =   3960
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   59
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   54
      Left            =   4320
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   58
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   53
      Left            =   3240
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   57
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   52
      Left            =   3600
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   56
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   51
      Left            =   2520
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   55
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   50
      Left            =   2880
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   54
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   49
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   53
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   48
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   52
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   47
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   51
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   46
      Left            =   1440
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   50
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   45
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   49
      Top             =   6120
      Width           =   615
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   44
      Left            =   720
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   48
      Top             =   6120
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   43
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   47
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   42
      Left            =   4920
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   46
      Top             =   5760
      Width           =   615
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   41
      Left            =   3840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   45
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   40
      Left            =   4200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   44
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   39
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   43
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   38
      Left            =   3480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   42
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   37
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   41
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   36
      Left            =   2760
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   40
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   35
      Left            =   1680
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   39
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   34
      Left            =   2040
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   38
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   33
      Left            =   960
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   37
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   32
      Left            =   1320
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   36
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   31
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   35
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   30
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   34
      Top             =   5760
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   29
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   33
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   28
      Left            =   4800
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   32
      Top             =   5400
      Width           =   735
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   27
      Left            =   3720
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   31
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   26
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   30
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   25
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   29
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   24
      Left            =   3360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   28
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   23
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   27
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   22
      Left            =   2640
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   26
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   1920
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   24
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   840
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   23
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   1200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   21
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   20
      Top             =   5400
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   5640
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   19
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   18
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   6360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   17
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   15
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   4800
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   3480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   12
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   10
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2760
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   9
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1440
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   720
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox pbKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label cmdTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   111
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label cmdTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   110
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label cmdTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   109
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label cmdTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   9720
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   9720
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type KeyInfoType
    IsKeyDown(1 To 254) As Boolean
    
    NotFirstPress(1 To 254) As Boolean
    
    KeyRepeatCount(1 To 254) As Long
    KeyPressCount(1 To 254) As Long
    
    KeyDownTime(1 To 254) As Double
    LastDownTime As Double
    
    CurrentKeyPressCount As Long
    TotalKeyPressCount As Long
End Type


Dim KeyInfo As KeyInfoType
Dim RolloverCount As Integer

Dim WindowHandle As Long

Public WithEvents Class As clsKeyStats
Attribute Class.VB_VarHelpID = -1

Private Sub Command2_Click()

End Sub

Private Sub Class_Boom(Num As Long)
    Debug.Print "boom"
End Sub

Private Sub cmdTab_Click(Index As Integer)

End Sub

Private Sub Command1_Click()
    Call Class.Bang
End Sub

'Debug.Print GetPerfCount
'If GetActiveWindow = WindowHandle Then Debug.Print "asdf" & Rnd




''''''''''''''''''''''''
Private Sub Form_Load()
    If modKBHook.HookMe = False Then
        MsgBox "Keyboar Hook Failed", vbCritical, "KBD Sentinel"
    End If
    
    WindowHandle = Me.hWnd
    Set Class = New clsKeyStats

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call modKBHook.UnhookMe
End Sub
'''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''
Private Sub txtRollover_KeyPress(KeyAscii As Integer)
    If RolloverCount = 0 Then txtRollover.Text = vbNullString

    If InStrB(1, txtRollover.Text, ChrW$(KeyAscii), vbTextCompare) Then
        KeyAscii = 0
    Else
        RolloverCount = RolloverCount + 1
        lbl = RolloverCount
    End If
End Sub

Private Sub txtRollover_KeyUp(KeyCode As Integer, Shift As Integer)
    If RolloverCount > 0 Then RolloverCount = RolloverCount - 1
End Sub
''''''''''''''''''''''''''''''''''''





' Process Key Input Information
''''''''''''''''''''''''''''''''''''''''''
Public Sub ProcAllKeys(vkCode As Long, ScanCode As Long, Flags As Long, Time As Double)
'
End Sub

Public Sub ProcKeyDown(vkCode As Long, ScanCode As Long, Flags As Long, Time As Double)
    If KeyInfo.IsKeyDown(vkCode) = False Then KeyInfo.IsKeyDown(vkCode) = True
    
    If KeyInfo.NotFirstPress(vkCode) = False Then
        KeyInfo.KeyDownTime(vkCode) = Time
        KeyInfo.CurrentKeyPressCount = KeyInfo.CurrentKeyPressCount + 1
        Label1.Caption = KeyInfo.CurrentKeyPressCount
        Label2.Caption = Time - KeyInfo.LastDownTime
        KeyInfo.NotFirstPress(vkCode) = True
    Else
        KeyInfo.KeyRepeatCount(vkCode) = KeyInfo.KeyRepeatCount(vkCode) + 1
    End If
    

    KeyInfo.LastDownTime = Time

    KeyInfo.KeyPressCount(vkCode) = KeyInfo.KeyPressCount(vkCode) + 1
    KeyInfo.TotalKeyPressCount = KeyInfo.TotalKeyPressCount + 1
    
    
    If KeyInfo.NotFirstPress(vkCode) = False Then
        'List1.AddItem KeyInfo.KeyDownTime(vkCode)
        List1.AddItem vbTab & vkCode & vbTab & ScanCode & vbTab & Flags & vbTab & Time
        List1.TopIndex = List1.NewIndex
    End If
End Sub

Public Sub ProcKeyUp(vkCode As Long, ScanCode As Long, Flags As Long, Time As Double)
    List1.AddItem Time - KeyInfo.KeyDownTime(vkCode)
    
    KeyInfo.NotFirstPress(vkCode) = False
    KeyInfo.CurrentKeyPressCount = KeyInfo.CurrentKeyPressCount - 1
    
    Label1.Caption = KeyInfo.CurrentKeyPressCount
    
    List1.TopIndex = List1.NewIndex
End Sub
''''''''''''''''''''''''''''''''''''''''''
' /Process Key Input Information





Public Sub AddKeyInfoToList(ListBox As ListBox, vkCode As Long, ScanCode As Long, Flags As Long, Time As Double)

    ListBox.AddItem vbTab & vkCode & vbTab & _
                            ScanCode & vbTab & _
                            Flags & vbTab & _
                            Time
                            
    ListBox.TopIndex = ListBox.NewIndex
End Sub
