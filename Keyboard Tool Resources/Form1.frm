VERSION 5.00
Begin VB.Form AutoClicker 
   Caption         =   "Auto Clicker"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Clicker 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   2280
   End
   Begin VB.Timer ClickCounter 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3840
      Top             =   2280
   End
   Begin VB.Timer Hotkey_XY 
      Interval        =   50
      Left            =   3360
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "1000"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click Counter"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "Click"
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 Clicks Per Sec"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y: 0"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X: 0"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status: Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Delay (in milliseconds):"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "AutoClicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const LEFT_DOWN = &H2
Private Const LEFT_UP = &H4

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim POINT_API As POINTAPI

Dim X As Long, Y As Long, Z As Long, A As Long

Private Sub LeftClick()
    mouse_event LEFT_DOWN, 0&, 0&, X, Y
    mouse_event LEFT_UP, 0&, 0&, X, Y
End Sub

Private Function GetMousePosX() As Long
    Z = GetCursorPos(POINT_API)
    GetMousePosX = POINT_API.X
End Function

Private Function GetMousePosY() As Long
    Z = GetCursorPos(POINT_API)
    GetMousePosY = POINT_API.Y
End Function

Private Function GetKeyState(Key As Integer) As Boolean
    GetKeyState = CBool(GetAsyncKeyState(Key))
End Function

Private Sub StartClick()
    If Clicker.Enabled = False Then
        Clicker.Interval = CInt(Val(Text1.Text))
        Clicker.Enabled = True
        
        Label2.Caption = " Status: Clicking..."
    End If
End Sub

Private Sub StopClick()
    If Clicker.Enabled = True Then
        Clicker.Enabled = False
    
        Label2.Caption = " Status: Clicking Stopped!"
    End If
End Sub

Private Sub Clicker_Timer()
    Call LeftClick
    DoEvents
End Sub

Private Sub Command1_Click()
    Call StartClick
End Sub

Private Sub Command2_Click()
    Call StopClick
End Sub

Private Sub Hotkey_XY_Timer()
    If GetKeyState(vbKeyF9) = True Then Call StartClick
    If GetKeyState(vbKeyF10) = True Then Call StopClick
    
    Label4.Caption = "X: " & CStr(GetMousePosX)
    Label5.Caption = "Y: " & CStr(GetMousePosY)
    
    DoEvents
End Sub

Private Sub ClickCounter_Timer()
    If A = 0 Then
        ClickCounter.Enabled = False
        Command3.Caption = "Click"
        Label3.Caption = "0 Clicks Per Sec"
    Else
        Label3.Caption = CStr(A * 0.25) & " Clicks Per Sec"
        A = 0
    End If
End Sub

Private Sub Command3_Click()
    If ClickCounter.Enabled = False Then ClickCounter.Enabled = True
    If Label3.Caption <> "Counting..." Then Label3.Caption = "Counting..."
    
    A = A + 1
    Command3.Caption = CStr(A)
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "Application running.", vbExclamation, "Application"

        Unload Me
        End
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    Unload Me
    End
End Sub

