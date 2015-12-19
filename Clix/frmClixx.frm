VERSION 5.00
Begin VB.Form Clixx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clix"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClixx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboHotkey 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   29
      Text            =   "F3"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Spammer"
      Height          =   255
      Left            =   8400
      TabIndex        =   23
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame frmSpam 
      Caption         =   "Spammer"
      Height          =   2895
      Left            =   3960
      TabIndex        =   19
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox frmSpamOptions 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   1815
         TabIndex        =   24
         Top             =   2520
         Width           =   1815
         Begin VB.CheckBox chkSpamA 
            Caption         =   "@.."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   495
         End
         Begin VB.CheckBox chkSpamFind 
            Caption         =   "/find"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1280
            TabIndex        =   26
            Top             =   0
            Width           =   615
         End
         Begin VB.CheckBox chkSpamA2 
            Caption         =   "@ @.."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   560
            TabIndex        =   25
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.CommandButton cmdClearSpam 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoadSpam 
         Caption         =   "Load"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtSpam 
         Height          =   2205
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame frmHotkey 
      Caption         =   "Hotkeys"
      Height          =   1695
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   1695
      Begin VB.ComboBox cboHotkey 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         ItemData        =   "frmClixx.frx":0CCA
         Left            =   960
         List            =   "frmClixx.frx":0CCC
         TabIndex        =   15
         Text            =   "F1"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboHotkey 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Text            =   "F2"
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox cboHotkey 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   13
         Text            =   "F3"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Auto Loot:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Auto Click:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Auto Atk:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   480
   End
   Begin VB.Timer tmrLoot 
      Interval        =   50
      Left            =   3600
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      Begin VB.CheckBox chkSystemTray 
         Caption         =   "System Tray"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "1"
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay (s):"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Stay On Top"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Enter"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "20"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Spammer:"
      Height          =   255
      Index           =   2
      Left            =   8520
      TabIndex        =   30
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H009F9F9F&
      X1              =   9000
      X2              =   9000
      Y1              =   2160
      Y2              =   2335
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   28
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Launcher 
      Height          =   240
      Left            =   3360
      Picture         =   "frmClixx.frx":0CCE
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblWebsite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Website"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3080
      TabIndex        =   11
      Top             =   2775
      Width           =   615
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   2775
      Width           =   1695
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H009F9F9F&
      Height          =   255
      Left            =   -120
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Clicks Per Sec:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Delay in milliseconds"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuSystemTray 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Clixx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n As Long

Dim HotkeyClick As Integer, HotkeyAttack As Integer, HotkeyLoot As Integer
Dim Keybd As clsKeyboard, Mouse As clsMouse, Registry As clsRegistry

Dim AutoAttack As Boolean, AutoLoot As Boolean

Dim FiveCount As Integer 'Every 250ms, 4 times a sec

Dim Run As String
    
Private Function CheckValue(Num As String) As Boolean
    n = Val(Num)
    
    If IsNumeric(n) = True Then
        If n > 0 Then
            If n > 1000 Then n = 1000
            CheckValue = True
        Else
            CheckValue = False
        End If
    Else
        CheckValue = False
    End If
End Function

Private Sub StartClick()
    If CheckValue(Text1.Text) = True Then
    
        Beeper 1000, 80, True
        
        If Timer1.Enabled = False Then
        
        Text1.Text = Val(Text1.Text)
        Timer1.Interval = Round(1000 / n, 0)
        
        If Check4.Value = vbChecked Then
            If CheckValue(Text2.Text) = True Then
                Text2.Text = Val(Text2.Text)
                lblStatus.Caption = "Delay On, " & Text2.Text & " Seconds"

                Pause Text2.Text
            Else
                MsgBox "Invalid Delay", vbOKCancel, "Invalid"
                Exit Sub
            End If
        End If

            Timer1.Enabled = True
            Text1.Enabled = False
            
            lblStatus.Caption = "Auto Clicker On!"
        End If
    Else
        MsgBox "Invalid Clicks Per Second", vbOKCancel, "Invalid"
    End If
End Sub

Private Sub StopClick()
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        Text1.Enabled = True
        
        If Text2.Enabled = False Then Text2.Enabled = True
        
        lblStatus.Caption = "Auto Clicker Off!"
        
        Beeper 1000, 80, False
    End If
End Sub

Private Sub cboHotkey_Click(Index As Integer)
Dim tmpHotkey As Long

If cboHotkey(0).Text = cboHotkey(1).Text _
  Or cboHotkey(1).Text = cboHotkey(2).Text _
  Or cboHotkey(2).Text = cboHotkey(0).Text Then

    MsgBox "Can't have same hotkeys.", vbInformation + vbOKCancel, "Hotkey"
    
    cboHotkey(0).Text = "F1"
    cboHotkey(1).Text = "F2"
    cboHotkey(2).Text = "F3"
    
    HotkeyClick = vbKeyF1
    HotkeyAttack = vbKeyF2
    HotkeyLoot = vbKeyF3
    
    Exit Sub
End If

Select Case cboHotkey(Index).Text
    Case "F1"
        tmpHotkey = vbKeyF1
    Case "F2"
        tmpHotkey = vbKeyF2
    Case "F3"
        tmpHotkey = vbKeyF3
    Case "F4"
        tmpHotkey = vbKeyF4
    Case "F5"
        tmpHotkey = vbKeyF5
    Case "F6"
        tmpHotkey = vbKeyF6
    Case "F7"
        tmpHotkey = vbKeyF7
    Case "F8"
        tmpHotkey = vbKeyF8
    Case "F9"
        tmpHotkey = vbKeyF9
    Case "F10"
        tmpHotkey = vbKeyF10
    Case "F11"
        tmpHotkey = vbKeyF11
    Case "F12"
        tmpHotkey = vbKeyF12
End Select

Select Case Index
    Case 0
        HotkeyClick = tmpHotkey
    Case 1
        HotkeyAttack = tmpHotkey
    Case 2
        HotkeyLoot = tmpHotkey
End Select
End Sub

Private Sub Check3_Click()
    If Check3.Value = vbChecked Then
        SetOnTop Me.hWnd, True
    Else
        SetOnTop Me.hWnd, False
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = vbChecked Then
        Text2.Enabled = True
    Else
        Text2.Enabled = False
    End If
End Sub

Private Sub chkSystemTray_Click()
    If chkSystemTray.Value = vbChecked Then
        CreateTray Me
    Else
        DeleteTray
    End If
End Sub

Private Sub Command1_Click()
    Call StartClick
End Sub

Private Sub Command2_Click()
    Call StopClick
    
    AutoAttack = False
    AutoLoot = False
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then CloseMe 'One Instance
End Sub

Private Sub Form_Initialize()
    Dim X As Integer, Y As Integer
    
    InitCommonControls 'System Style
    
    Set Keybd = New clsKeyboard 'Send Key
    Set Mouse = New clsMouse
    Set Registry = New clsRegistry
    
    'Settings
    Run = Val(Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Run"))
    
    If Val(Run) > 0 Then
        Text1.Text = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Text1")
        Text2.Text = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Text2")
        Check1.Value = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Check1")
        Check3.Value = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Check3")
        Check4.Value = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "Check4")
        
        chkSystemTray.Value = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "chkSystemTray")
        
        HotkeyClick = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyClick")
        HotkeyAttack = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyAttack")
        HotkeyLoot = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyLoot")
        
        cboHotkey(0).Text = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(0)")
        cboHotkey(1).Text = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(1)")
        cboHotkey(2).Text = Registry.GetData(HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(2)")
    Else
        cboHotkey(0).Text = "F1"
        cboHotkey(1).Text = "F2"
        cboHotkey(2).Text = "F3"
        
        HotkeyClick = vbKeyF1
        HotkeyAttack = vbKeyF2
        HotkeyLoot = vbKeyF3
    End If

    For X = 1 To 12
        For Y = 0 To 2
            cboHotkey(Y).AddItem "F" & X
        Next Y
    Next X
    
    If chkSystemTray.Value = vbChecked Then CreateTray Me
    Call Check3_Click 'on top
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkSystemTray.Value = vbChecked Then CreatePopMenu Me, X
End Sub

Private Sub Form_Resize()
    If chkSystemTray.Value = vbChecked Then
        If Me.WindowState = 1 Then
            mnuShow.Caption = "Show"
        Else
            mnuShow.Caption = "Hide"
        End If
        
        SendToTray Me
    End If
End Sub

Private Sub CloseMe()
    Dim frmAll As Form
    
    If chkSystemTray.Value = vbChecked Then DeleteTray
        
    Set Keybd = Nothing
    Set Mouse = Nothing
    
    For Each frmAll In Forms
        Unload frmAll
        Set frmAll = Nothing
    Next

    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Text1", Text1.Text
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Text2", Text2.Text
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Check1", Check1.Value
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Check3", Check3.Value
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Check4", Check4.Value
    
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "chkSystemTray", chkSystemTray.Value
    
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyClick", HotkeyClick
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyAttack", HotkeyAttack
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "HotkeyLoot", HotkeyLoot
    
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(0)", cboHotkey(0).Text
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(1)", cboHotkey(1).Text
    Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "cboHotkey(2)", cboHotkey(2).Text

    If Run = 0 Then
        Registry.SetData HKEY_CURRENT_USER, "Software\ClixAC", "Run", "1"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseMe
End Sub

Private Sub Launcher_Click()
Dim strLocation As String
strLocation = Registry.GetData(HKEY_LOCAL_MACHINE, "SOFTWARE\Wizet\MapleStory", "ExecPath")

If LenB(strLocation) > 0 Then
    If Left(LCase$(Trim$(App.EXEName)), 8) <> "joytokey" Then
        MsgBox "Exe name must start with ""JoyToKey"" to work with MapleStory", vbInformation, "JoyToKey"
        Exit Sub
    End If
    
    If Run = 0 Then
        MsgBox "Requires a BYPASS to use with MapleStory.", vbInformation, "Bypass"
    End If
    
    Shell strLocation & "\MapleStory.exe"
End If
End Sub

Private Sub lblWebsite_Click()
    OpenBrowser Me.hWnd, "http://vertion.fused.ws/"
End Sub

Private Sub mnuExit_Click()
    Call CloseMe
End Sub

Private Sub mnuShow_Click()
    If mnuShow.Caption = "Show" Then
        mnuShow.Caption = "Hide"
        SendToDesktop Me
    Else
        mnuShow.Caption = "Show"
        Me.WindowState = 1
        SendToTray Me
    End If
End Sub

Private Sub Text2_Change()
    If Val(Text2.Text) > 60 Then Text2.Text = "60"
    If Val(Text2.Text) < 1 Then Text2.Text = "1"
End Sub

Private Sub Timer1_Timer()
    Mouse.DoubleClick ClickLeft
    
    If Check1.Value = vbChecked Then Keybd.PressKeyVK keyReturn
End Sub

Private Sub tmrLoot_Timer()
    FiveCount = FiveCount + 1
    
    If AutoLoot = True Then Keybd.PressKeyVK keyNumPad0 '50ms interval, 20 a sec
    
    If FiveCount = 5 Then
    
    If AutoAttack = True Then Keybd.PressKeyVK keyControl

    'Hotkey
    
    If GetKeyState(HotkeyClick) = True Then 'Click
        If Timer1.Enabled = False Then
            Call StartClick
        Else
            Call StopClick
        End If
    End If
    
    If GetKeyState(HotkeyAttack) = True Then 'Attack
        If AutoAttack = False Then
            Beeper 1500, 80, True
            AutoAttack = True
            lblStatus.Caption = "Auto Attack On!"
        Else
            Beeper 1500, 80, False
            AutoAttack = False
            lblStatus.Caption = "Auto Attack Off!"
        End If
    End If
    
    If GetKeyState(HotkeyLoot) = True Then 'Loot
        If AutoLoot = False Then
            Beeper 1200, 80, True
            AutoLoot = True
            lblStatus.Caption = "Auto Loot On!"
        Else
            Beeper 1200, 80, False
            AutoLoot = False
            lblStatus.Caption = "Auto Loot Off!"
        End If
    End If
    
    FiveCount = 0
    End If
End Sub

Private Sub txtSpam_Change()
'max char = 70

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'stack
'@@@@@@@ @ @@@@@@@ @ @@@@@@@ @ @@@@@@@ @ @@@@@@@ @ @@@@@@@ @ @@@@@@@@@@
End Sub
