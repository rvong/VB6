VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KeyScan"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox frmKeyboard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   3615
      TabIndex        =   152
      Top             =   4320
      Width           =   3615
      Begin VB.CheckBox chkMarkKeysPressed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mark Keys Pressed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1780
      End
      Begin VB.CheckBox chkBlockKeyboardInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Block Keyboard Input"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1800
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   166
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   165
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Layout:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   164
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Function Keys:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   163
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblKeyPress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Key Press:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   162
         Top             =   1350
         Width           =   1815
      End
      Begin VB.Label lblKeyboardName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   161
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lblKeyCount 
         BackStyle       =   0  'Transparent
         Caption         =   "Key Count:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   160
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label lblRepeatRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   159
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label lblRepeatDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Delay:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   158
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label lblKeyboardDescription 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   157
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblKeyboardLayout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   156
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblFunctionKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1800
         TabIndex        =   155
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.PictureBox frmHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   3645
      TabIndex        =   167
      Top             =   4320
      Visible         =   0   'False
      Width           =   3645
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Block keyboard input, system-wide."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   177
         Top             =   1935
         Width           =   3735
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Block Keyboard Input"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   176
         Top             =   1755
         Width           =   1815
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark keys that have been pressed.  Click again to reset or disable."
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   11
         Left            =   0
         TabIndex        =   175
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Keys Pressed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   174
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard repeat-delay when a key is pressed consecutively.  Approx. 250ms to 1 sec delay."
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   173
         Top             =   680
         Width           =   3735
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " - range 0-3 (shortest-longest)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   172
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Delay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   171
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard repeat-speed when a key is held down."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   170
         Top             =   200
         Width           =   3735
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " - range 0-31 (slowest-fastest)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   169
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   168
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox frmAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   3615
      TabIndex        =   178
      Top             =   4320
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for using KeyScan!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   186
         Top             =   1860
         Width           =   2775
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This application is freeware."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   840
         TabIndex        =   185
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date: 7/2009"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   840
         TabIndex        =   184
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version: v1.0 beta"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   840
         TabIndex        =   183
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Richard V."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   840
         TabIndex        =   182
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblWebsite 
         BackStyle       =   0  'Transparent
         Caption         =   "Website"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2880
         TabIndex        =   181
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Diagnostics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   840
         TabIndex        =   180
         Top             =   440
         Width           =   2535
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KeyScan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   14
         Left            =   840
         TabIndex        =   179
         Top             =   -80
         Width           =   2535
      End
      Begin VB.Image imgAboutIcon 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   735
      End
   End
   Begin ComctlLib.TabStrip tbsMain 
      Height          =   2655
      Left            =   120
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   3915
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   4683
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Keyboard"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lvwKeysPressed 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Key Data"
      Top             =   6600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virtual-key Constant"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virtual Key"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Scan Code"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Key Flags"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line line 
      BorderColor     =   &H00000000&
      Index           =   9
      X1              =   6080
      X2              =   6260
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Line line 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   6080
      X2              =   6240
      Y1              =   3530
      Y2              =   3530
   End
   Begin VB.Line line 
      BorderColor     =   &H00000000&
      Index           =   7
      X1              =   6080
      X2              =   6260
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   5
      Left            =   9600
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   4
      Left            =   9840
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   3
      Left            =   10800
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   2
      Left            =   10560
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   1
      Left            =   10320
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape shpBlocks 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Index           =   0
      Left            =   10080
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "num"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   143
      Left            =   9000
      TabIndex        =   150
      Top             =   1100
      Width           =   645
   End
   Begin VB.Shape shpLED 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   0
      Left            =   9195
      Top             =   960
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "caps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   144
      Left            =   9645
      TabIndex        =   149
      Top             =   1100
      Width           =   645
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scroll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   145
      Left            =   10275
      TabIndex        =   148
      Top             =   1100
      Width           =   660
   End
   Begin VB.Shape shpLED 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   9840
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape shpLED 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   2
      Left            =   10485
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   9645
      Top             =   840
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   10275
      Top             =   840
      Width           =   660
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   9000
      Top             =   840
      Width           =   660
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   142
      Left            =   9960
      TabIndex        =   147
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   141
      Left            =   9000
      TabIndex        =   146
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   140
      Left            =   9480
      TabIndex        =   145
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   14
      Left            =   9480
      TabIndex        =   144
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblSysInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   143
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label lblOS 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   142
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblKeyboardDiagnostics 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Diagnostics"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   141
      Top             =   435
      Width           =   1935
   End
   Begin VB.Label lblKeyScan 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KeyScan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   105
      TabIndex        =   140
      Top             =   0
      Width           =   1935
   End
   Begin VB.Line line 
      Index           =   6
      X1              =   8460
      X2              =   8840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Del"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   139
      Left            =   9960
      TabIndex        =   139
      Top             =   3640
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Ins"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   138
      Left            =   9000
      TabIndex        =   138
      Top             =   3640
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " PgDn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   137
      Left            =   9960
      TabIndex        =   137
      Top             =   3160
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   136
      Left            =   9000
      TabIndex        =   136
      Top             =   3160
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " PgUp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   135
      Left            =   9960
      TabIndex        =   135
      Top             =   2200
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   134
      Left            =   9000
      TabIndex        =   134
      Top             =   2200
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Enter"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   133
      Left            =   10440
      TabIndex        =   133
      Top             =   2910
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   132
      Left            =   9945
      TabIndex        =   132
      Top             =   3195
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   131
      Left            =   9000
      TabIndex        =   131
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   130
      Left            =   10440
      TabIndex        =   130
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   129
      Left            =   9960
      TabIndex        =   129
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   128
      Left            =   9480
      TabIndex        =   128
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   127
      Left            =   9000
      TabIndex        =   127
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   126
      Left            =   9960
      TabIndex        =   126
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   125
      Left            =   9480
      TabIndex        =   125
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   124
      Left            =   9000
      TabIndex        =   124
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   123
      Left            =   9960
      TabIndex        =   123
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   122
      Left            =   9480
      TabIndex        =   122
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   121
      Left            =   9000
      TabIndex        =   121
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  _"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   120
      Left            =   10440
      TabIndex        =   120
      Top             =   1350
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   119
      Left            =   10080
      TabIndex        =   119
      Top             =   1460
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   118
      Left            =   9600
      TabIndex        =   118
      Top             =   1470
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Num"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   117
      Left            =   9000
      TabIndex        =   117
      Top             =   1515
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Lock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   116
      Left            =   9000
      TabIndex        =   116
      Top             =   1660
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   115
      Left            =   8400
      TabIndex        =   115
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   114
      Left            =   7920
      TabIndex        =   114
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   113
      Left            =   7440
      TabIndex        =   113
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   112
      Left            =   7920
      TabIndex        =   112
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   111
      Left            =   8400
      TabIndex        =   111
      Top             =   2140
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   110
      Left            =   8400
      TabIndex        =   110
      Top             =   2000
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   109
      Left            =   7920
      TabIndex        =   109
      Top             =   2060
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   108
      Left            =   7440
      TabIndex        =   108
      Top             =   2060
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   107
      Left            =   8400
      TabIndex        =   107
      Top             =   1660
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   106
      Left            =   8400
      TabIndex        =   106
      Top             =   1520
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   105
      Left            =   7920
      TabIndex        =   105
      Top             =   1580
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Insert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   104
      Left            =   7440
      TabIndex        =   104
      Top             =   1580
      Width           =   495
   End
   Begin VB.Line line 
      Index           =   5
      X1              =   6720
      X2              =   6720
      Y1              =   2620
      Y2              =   2700
   End
   Begin VB.Line line 
      Index           =   4
      X1              =   6360
      X2              =   6720
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Line line 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   6480
      X2              =   6880
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Line line 
      Index           =   2
      X1              =   240
      X2              =   600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   98
      Left            =   6120
      TabIndex        =   103
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   103
      Left            =   480
      TabIndex        =   102
      Top             =   2160
      Width           =   255
   End
   Begin VB.Line line 
      Index           =   1
      X1              =   240
      X2              =   600
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Image imgWinKey 
      Height          =   195
      Index           =   1
      Left            =   5490
      Picture         =   "frmMain.frx":2EFA
      Top             =   3440
      Width           =   225
   End
   Begin VB.Image imgWinKey 
      Height          =   195
      Index           =   0
      Left            =   920
      Picture         =   "frmMain.frx":32C5
      Top             =   3440
      Width           =   225
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   200
      Left            =   6080
      Top             =   3430
      Width           =   180
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Ctrl"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   102
      Left            =   6600
      TabIndex        =   101
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label keyLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Alt"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   101
      Left            =   4800
      TabIndex        =   100
      Top             =   3405
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Alt"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   100
      Left            =   1440
      TabIndex        =   99
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label keyLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Ctrl"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   99
      Left            =   120
      TabIndex        =   98
      Top             =   3405
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   97
      Left            =   6120
      TabIndex        =   97
      Top             =   2925
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   96
      Left            =   180
      TabIndex        =   96
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Shift"
      Height          =   255
      Index           =   95
      Left            =   120
      TabIndex        =   95
      Top             =   2925
      Width           =   615
   End
   Begin VB.Label keyLbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   94
      Left            =   6240
      TabIndex        =   94
      Top             =   2560
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Enter"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   93
      Left            =   6240
      TabIndex        =   93
      Top             =   2445
      Width           =   615
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  '"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   92
      Left            =   5760
      TabIndex        =   92
      Top             =   2715
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  """
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   91
      Left            =   5750
      TabIndex        =   91
      Top             =   2430
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ;"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   90
      Left            =   5280
      TabIndex        =   90
      Top             =   2610
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   76
      Left            =   5280
      TabIndex        =   89
      Top             =   2385
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  /"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   89
      Left            =   5520
      TabIndex        =   88
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   88
      Left            =   5520
      TabIndex        =   87
      Top             =   2895
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   87
      Left            =   5040
      TabIndex        =   86
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   86
      Left            =   5040
      TabIndex        =   85
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   85
      Left            =   4560
      TabIndex        =   84
      Top             =   2970
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  <"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   84
      Left            =   4560
      TabIndex        =   83
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   83
      Left            =   4080
      TabIndex        =   82
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   82
      Left            =   3600
      TabIndex        =   81
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   81
      Left            =   3120
      TabIndex        =   80
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   80
      Left            =   2640
      TabIndex        =   79
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   79
      Left            =   2160
      TabIndex        =   78
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   78
      Left            =   1680
      TabIndex        =   77
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   77
      Left            =   1200
      TabIndex        =   76
      Top             =   2900
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   75
      Left            =   4800
      TabIndex        =   75
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   74
      Left            =   4320
      TabIndex        =   74
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   73
      Left            =   3840
      TabIndex        =   73
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   72
      Left            =   3360
      TabIndex        =   72
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   71
      Left            =   2880
      TabIndex        =   71
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   70
      Left            =   2400
      TabIndex        =   70
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   69
      Left            =   1920
      TabIndex        =   69
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   68
      Left            =   1440
      TabIndex        =   68
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   67
      Left            =   960
      TabIndex        =   67
      Top             =   2420
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caps Lock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   66
      Left            =   105
      TabIndex        =   66
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  \"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   65
      Left            =   6600
      TabIndex        =   65
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  |"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   64
      Left            =   6600
      TabIndex        =   64
      Top             =   1935
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   63
      Left            =   6120
      TabIndex        =   63
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  }"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   62
      Left            =   6120
      TabIndex        =   62
      Top             =   1935
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  ["
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   61
      Left            =   5640
      TabIndex        =   61
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  {"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   60
      Left            =   5620
      TabIndex        =   60
      Top             =   1935
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   59
      Left            =   5160
      TabIndex        =   59
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   58
      Left            =   4680
      TabIndex        =   58
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   57
      Left            =   4200
      TabIndex        =   57
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   56
      Left            =   3720
      TabIndex        =   56
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   55
      Left            =   3240
      TabIndex        =   55
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   54
      Left            =   2760
      TabIndex        =   54
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   53
      Left            =   2280
      TabIndex        =   53
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   52
      Left            =   1800
      TabIndex        =   52
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   51
      Left            =   1320
      TabIndex        =   51
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   50
      Left            =   840
      TabIndex        =   50
      Top             =   1940
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Tab"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   49
      Left            =   120
      TabIndex        =   49
      Top             =   1935
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   48
      Left            =   180
      TabIndex        =   48
      Top             =   2060
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Backspace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   47
      Left            =   6360
      TabIndex        =   47
      Top             =   1480
      Width           =   975
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   46
      Left            =   6360
      TabIndex        =   46
      Top             =   1625
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   45
      Left            =   5880
      TabIndex        =   45
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   44
      Left            =   5880
      TabIndex        =   44
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " _"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   43
      Left            =   5400
      TabIndex        =   43
      Top             =   1220
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  _"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   42
      Left            =   5415
      TabIndex        =   42
      Top             =   1605
      Width           =   240
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  )"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   41
      Left            =   4920
      TabIndex        =   41
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   40
      Left            =   4920
      TabIndex        =   40
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  ("
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   39
      Left            =   4440
      TabIndex        =   39
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   38
      Left            =   4440
      TabIndex        =   38
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   37
      Left            =   3960
      TabIndex        =   37
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   36
      Left            =   3960
      TabIndex        =   36
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  &&"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   35
      Left            =   3480
      TabIndex        =   35
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   34
      Left            =   3480
      TabIndex        =   34
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  ^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   33
      Left            =   3000
      TabIndex        =   33
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   32
      Left            =   3000
      TabIndex        =   32
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  %"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   31
      Left            =   2520
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   30
      Left            =   2520
      TabIndex        =   30
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  $"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   29
      Left            =   2040
      TabIndex        =   29
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  4"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   28
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  #"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   27
      Left            =   1560
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  3"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   26
      Left            =   1560
      TabIndex        =   26
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  2"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   25
      Left            =   1080
      TabIndex        =   25
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   24
      Left            =   1080
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  1"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   23
      Left            =   600
      TabIndex        =   23
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  !"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   " `"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  ~"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Break"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   8400
      TabIndex        =   19
      Top             =   1120
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   18
      Left            =   8400
      TabIndex        =   18
      Top             =   885
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   7920
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   7920
      TabIndex        =   16
      Top             =   920
      Width           =   495
   End
   Begin VB.Line line 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   7500
      X2              =   7880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SysRq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   15
      Top             =   1120
      Width           =   495
   End
   Begin VB.Label keyLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PrtScr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   13
      Left            =   7440
      TabIndex        =   14
      Top             =   885
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F12"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   6840
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F11"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F10"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F9"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F8"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F7"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F6"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F5"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F4"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F3"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F2"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  F1"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label keyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "  Esc"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   -120
      Top             =   -120
      Width           =   11295
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   110
      Left            =   9960
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   96
      Left            =   9000
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   975
      Index           =   1013
      Left            =   10440
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   99
      Left            =   9960
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   98
      Left            =   9480
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   97
      Left            =   9000
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   102
      Left            =   9960
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   101
      Left            =   9480
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   100
      Left            =   9000
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   975
      Index           =   107
      Left            =   10440
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   105
      Left            =   9960
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   104
      Left            =   9480
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   103
      Left            =   9000
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   109
      Left            =   10440
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   106
      Left            =   9960
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   111
      Left            =   9480
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   144
      Left            =   9000
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   38
      Left            =   7920
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   39
      Left            =   8400
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   40
      Left            =   7920
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   37
      Left            =   7440
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   34
      Left            =   8400
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   35
      Left            =   7920
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   46
      Left            =   7440
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   33
      Left            =   8400
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   36
      Left            =   7920
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   45
      Left            =   7440
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   163
      Left            =   6600
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   93
      Left            =   6000
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   92
      Left            =   5400
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   165
      Left            =   4800
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   32
      Left            =   2040
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   164
      Left            =   1440
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   91
      Left            =   840
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   162
      Left            =   120
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   161
      Left            =   6000
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   191
      Left            =   5520
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   190
      Left            =   5040
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   188
      Left            =   4560
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   77
      Left            =   4080
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   78
      Left            =   3600
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   66
      Left            =   3120
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   86
      Left            =   2640
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   67
      Left            =   2160
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   88
      Left            =   1680
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   90
      Left            =   1200
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   160
      Left            =   120
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   13
      Left            =   6240
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   222
      Left            =   5760
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   186
      Left            =   5280
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   76
      Left            =   4800
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   75
      Left            =   4320
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   74
      Left            =   3840
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   72
      Left            =   3360
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   71
      Left            =   2880
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   70
      Left            =   2400
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   68
      Left            =   1920
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   83
      Left            =   1440
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   65
      Left            =   960
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   20
      Left            =   120
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   220
      Left            =   6600
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   221
      Left            =   6120
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   219
      Left            =   5640
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   80
      Left            =   5160
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   79
      Left            =   4680
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   73
      Left            =   4200
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   85
      Left            =   3720
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   89
      Left            =   3240
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   84
      Left            =   2760
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   82
      Left            =   2280
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   69
      Left            =   1800
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   87
      Left            =   1320
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   81
      Left            =   840
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   120
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   6360
      Top             =   1440
      Width           =   975
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   187
      Left            =   5880
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   189
      Left            =   5400
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   48
      Left            =   4920
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   57
      Left            =   4440
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   56
      Left            =   3960
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   55
      Left            =   3480
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   54
      Left            =   3000
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   53
      Left            =   2520
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   52
      Left            =   2040
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   51
      Left            =   1560
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   50
      Left            =   1080
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   49
      Left            =   600
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   192
      Left            =   120
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   19
      Left            =   8400
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   145
      Left            =   7920
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   44
      Left            =   7440
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   123
      Left            =   6840
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   122
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   121
      Left            =   5880
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   120
      Left            =   5400
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   119
      Left            =   4680
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   118
      Left            =   4200
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   117
      Left            =   3720
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   116
      Left            =   3240
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   115
      Left            =   2520
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   114
      Left            =   2040
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   113
      Left            =   1560
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   112
      Left            =   1080
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   27
      Left            =   120
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Launch Control Panel Keyboard Setting
'Shell "control keyboard"
Dim vk_list_dec() As String  'Used for unmarking keys shapes

Private Sub GetDevice()
  Dim os_archi As String
  Dim mem_capacity As Double
  
  Dim wbemObject As Object
  Dim RegAccess As clsRegistry
  
  Set RegAccess = New clsRegistry
  
  'On Error Resume Next
  
  'Get & Format OS, Build Number, Service Pack, & Architecture (Bit)
  Set wbemObject = GetObject("winmgmts:\\").InstancesOf("Win32_OperatingSystem")
  For Each wbemObject In wbemObject
      'RICHARD-PC = wbemObject.CSName
      lblOS.Caption = Trim$(wbemObject.Caption) & " | " & Trim$(wbemObject.Version) & " | " & Trim$(wbemObject.CSDVersion)
  Next
  
  'Get Architecture
  os_archi = LCase$(Trim$(RegAccess.GetData(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE")))
  
  If os_archi = "x86" Then os_archi = "32-bit"
  If os_archi = "x64" Then os_archi = "64-bit"
  
  lblOS.Caption = lblOS.Caption & " | " & os_archi
  
  'Get & Format Processor Name
  lblSysInfo.Caption = FormatProcessorName(RegAccess.GetData(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString"))

  'Get & Format Physical Memory
  Set wbemObject = GetObject("winmgmts:\\").InstancesOf("Win32_PhysicalMemory")
  For Each wbemObject In wbemObject
      mem_capacity = mem_capacity + (Val(wbemObject.Capacity) / (1024& * 1024& * 1024&))
  Next
      lblSysInfo.Caption = lblSysInfo.Caption & ", " & mem_capacity & "GB RAM"
  
  'Get & Format Video Card
  Set wbemObject = GetObject("winmgmts:\\").InstancesOf("Win32_VideoController")
  For Each wbemObject In wbemObject
      lblSysInfo.Caption = lblSysInfo.Caption & ", " & wbemObject.Name
  Next
  
  'Get & Format Keyboard
  Set wbemObject = GetObject("winmgmts:\\").InstancesOf("Win32_Keyboard")
  
  For Each wbemObject In wbemObject
      lblKeyboardName.Caption = wbemObject.Name
      lblKeyboardDescription.Caption = wbemObject.Description
      lblKeyboardLayout.Caption = wbemObject.Layout & " " & LangIdent(wbemObject.Layout)
      lblFunctionKeys.Caption = wbemObject.NumberOfFunctionKeys
  Next
  
  'KeyboardSpeed (Repeat Rate), & Repeat Delay
  lblRepeatRate.Caption = "Repeat Rate:  " & RegAccess.GetData(HKEY_CURRENT_USER, "Control Panel\Keyboard", "KeyboardSpeed")
  lblRepeatDelay.Caption = "Repeat Delay:  " & RegAccess.GetData(HKEY_CURRENT_USER, "Control Panel\Keyboard", "KeyboardDelay")
  
  Set wbemObject = Nothing
  Set RegAccess = Nothing
End Sub

Private Sub LoadVKLists()
  Dim vk_name() As String, vk_dsc() As String, vk_dec() As String
  Dim x As Long
  
  Set vk_collection = New Collection
  
  vk_name() = Split(StrConv(LoadResData("NAME", "VK_LIST"), vbUnicode), vbNewLine)
  vk_dsc() = Split(StrConv(LoadResData("DSC", "VK_LIST"), vbUnicode), vbNewLine)
  vk_dec() = Split(StrConv(LoadResData("DEC", "VK_LIST"), vbUnicode), vbNewLine)
  vk_list_dec = vk_dec
  
  For x = 0 To UBound(vk_name)
    vk_collection.Add vk_name(x) & vbTab & vk_dsc(x), vk_dec(x)
  Next x
End Sub

Private Sub chkMarkKeysPressed_Click()
  Dim x As Integer
  
  If chkMarkKeysPressed.Value = vbUnchecked Then
    For x = LBound(vk_list_dec) To UBound(vk_list_dec)
      If HasIndex(shpKey, Val(vk_list_dec(x))) = True Then shpKey(Val(vk_list_dec(x))).BackColor = &HFFFFFF
    Next x
  End If
End Sub

'' Form Loads
Private Sub Form_Initialize()
  Call InitCommonControls
  
  Call HookMe
  
  Call GetDevice
  Call LoadVKLists
  
  Call UpdateLED(vbKeyNumlock, True)
  Call UpdateLED(vbKeyCapital, True)
  Call UpdateLED(vbKeyScrollLock, True)
End Sub

Private Sub Form_Activate()
  frmKeyboard.SetFocus
  imgAboutIcon.Picture = Me.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set vk_collection = Nothing
  
  Call UnhookMe
End Sub
''Form Ends

Private Sub lblWebsite_Click()
  Call OpenURL("http://sites.google.com/site/inertially/keyscan")
End Sub

Private Sub chkMarkKeysPressed_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  frmKeyboard.SetFocus
End Sub

Private Sub chkBlockKeyboardInput_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  frmKeyboard.SetFocus
End Sub

Private Sub lvwKeysPressed_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  frmKeyboard.SetFocus
End Sub


'' Tab Switcher
Private Sub tbsMain_Click()
  Select Case tbsMain.SelectedItem.Index
    Case 1 'Keyboard
      frmKeyboard.Visible = True
      frmHelp.Visible = False
      frmAbout.Visible = False
    Case 2 'Help
      frmKeyboard.Visible = False
      frmHelp.Visible = True
      frmAbout.Visible = False
    Case 3 'About
      frmKeyboard.Visible = False
      frmHelp.Visible = False
      frmAbout.Visible = True
  End Select
End Sub
