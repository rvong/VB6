VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Yahoo! Password Changer"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetNPW 
      Caption         =   "Set New Password as Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboEditSvr 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "frmMain.frx":1CFA
      Left            =   3000
      List            =   "frmMain.frx":1D0A
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "edit.yahoo.com"
      Top             =   3880
      Width           =   2535
   End
   Begin VB.ComboBox cboLoginSvr 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "frmMain.frx":1D5F
      Left            =   3000
      List            =   "frmMain.frx":1D6C
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "login.yahoo.com"
      Top             =   3400
      Width           =   2535
   End
   Begin VB.HScrollBar hsbOpacity 
      Height          =   255
      LargeChange     =   20
      Left            =   1080
      Max             =   255
      Min             =   50
      SmallChange     =   10
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3720
      Value           =   255
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2665
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2665
      Width           =   1095
   End
   Begin VB.CheckBox chkKeepOnTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Keep on Top"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2665
      Width           =   1095
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2665
      Width           =   1095
   End
   Begin VB.ListBox lstResults 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      IntegralHeight  =   0   'False
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtNPW 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtOPW 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin SHDocVwCtl.WebBrowser Nav 
      CausesValidation=   0   'False
      Height          =   5175
      Left            =   6000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label cmdMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5160
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5400
      TabIndex        =   22
      Top             =   80
      Width           =   330
   End
   Begin VB.Label lblStatus2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo! Account Password Changer"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   75
      Width           =   5535
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opacity"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   3705
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4560
      Left            =   0
      Top             =   405
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00808080&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Image img 
      Height          =   420
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":1DB0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1215
      Left            =   120
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Log"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo! ID"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Image img 
      Height          =   4860
      Index           =   1
      Left            =   -120
      Picture         =   "frmMain.frx":1FD4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const YLogoutURL As String = "http://login.yahoo.com/login?logout=1"

'Trim cboText
Dim LoginDomain As String, EditDomain As String


'Change Password
Private Sub cmdChange_Click()

  'Check filled ID & OPW
  If Len(txtID) = 0 Or Len(txtOPW) = 0 Then
      lblStatus = "Please verify your ID/Password"
      
  'Check valid NPW
  ElseIf Len(txtNPW) < 6 Or txtOPW = txtNPW Then _
      lblStatus = "Invalid New Password"
      
  'Input OK!
  Else

    LoginDomain = Trim$(cboLoginSvr)
    EditDomain = Trim$(cboEditSvr)
    
    If LInstr(LoginDomain, ".yahoo.com") = False Then
    
      lblStatus = "Invalid Login Server"
      
    ElseIf LInstr(EditDomain, ".yahoo.com") = False Then
    
      lblStatus = "Invalid Edit Server"
      
    Else ' Input OK!, Server OK!
      
      lblStatus = "Signing in..."
      lblStatus2 = "Connecting..."
      
      Call EnableControls(False)
      Call Nav.Navigate(LoginURL(txtID, txtOPW, LoginDomain, EditDomain))
    End If
    
  End If
End Sub


''''''Commands'''''''''
'''''''''''''''''''''''
'''''''''Clear Text Boxes'''''''
Private Sub cmdReset_Click()
  txtID = vbNullString
  txtOPW = vbNullString
  txtNPW = vbNullString
End Sub

'''''''''Empty List Box Log'''''''
Private Sub cmdClear_Click()
  If lstResults.ListCount > 0 Then lstResults.Clear
End Sub

''''''''''Save Results'''''''
Private Sub cmdSave_Click()
  Dim i As Long
  
  On Error GoTo Error
  
  With cDialog
    .CancelError = True
    .DialogTitle = "Save As"
    .Filter = "Text Document (*.txt)|*.txt"
    .ShowSave
    
    If .Flags = 0 Then Exit Sub
    
    Open .FileName For Output As #1
      For i = 0 To lstResults.ListCount - 1
        Print #1, lstResults.List(i)
      Next i
    Close #1
  End With
  
Error: Err.Clear
End Sub

'Keep form on top
Private Sub chkKeepOnTop_Click()
  If chkKeepOnTop.Value = vbChecked Then
    Call SetOnTop(Me, True)
  Else
    Call SetOnTop(Me, False)
  End If
End Sub

'Change Opacity
Private Sub hsbOpacity_Change()
  Call Opacity(hsbOpacity.Value, Me)
End Sub

'Set Default NPW
Private Sub cmdSetNPW_Click()
  Call SetDefaultPassword
End Sub
'''''''''''''''''''''''''
'''''''End Commands''''''




''''''''Web Browser'''''''''
''''''''''''''''''''''''''''
'Page loaded
Private Sub Nav_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  'Call ProcessHTML
  If Nav.LocationURL <> "about:blank" And Nav.ReadyState = READYSTATE_COMPLETE Then
    Call ProcessHTML(Nav.Document.Body.InnerHTML)
  End If
End Sub

'''Does things
Private Sub ProcessHTML(Expression As String)
    lblStatus2 = "Processing data..."
    'Text1 = Nav.LocationURL & vbNewLine & Expression
    
    Dim navURL As String
    navURL = Nav.LocationURL
    
    'Success, Password Changed!
    If LInstr(navURL, ".confirm=1") Then
      lblStatus = "Password Changed!"
      lstResults.AddItem txtID & " - " & txtNPW
      
    'Confirm logout
    ElseIf LInstr(Expression, "signed out of the yahoo! network") Then
      'Nav.Stop
      Nav.Navigate "about:blank"
      lblStatus2 = "Signed Out"
      
      Call EnableControls(True)
      Exit Sub '<-- Already signed out
    
    'Bad Login, Invalid ID or password
    ElseIf LInstr(Expression, "Invalid ID or password") Then
      lblStatus = "Invalid ID or Password"
      lblStatus2 = "Data Processed!"
      
      Call EnableControls(True)
      Exit Sub '<-- Didn't login, already logged out
    
    'ID Not Exist
    ElseIf LInstr(Expression, "This ID is not yet taken") Then
      lblStatus = "ID Does Not Exist"
      lblStatus2 = "Data Processed!"
      
      Call EnableControls(True)
      Exit Sub
      
    'Already signed in, Bad Login
    ElseIf InStr(Expression, "Why am I being asked for my password") Then
      lblStatus = "Sign-out Error"
      
    'Banned from server
    ElseIf LInstr(Expression, "error 999") Then
      lblStatus = "999 Blocked: " & Left$(GetBetween(navURL, "://", "/"), 21)
      
    'Wrong redirect
    ElseIf LInstr(navURL, "my.yahoo") Or LInstr(navURL, "www.yahoo") Then
    
      If LInstr(Expression, "Sign Out") Then
      
        lblStatus = "Redirecting..."
        lblStatus2 = "Connecting..."
        Nav.Navigate "https://" & EditDomain & "/config/change_pw?.redir_from=LOGIN"
        
        DoEvents
        Exit Sub
        
      Else
        lblStatus = "Redirect Error"
      End If

    'International
    ElseIf LInstr(navURL, "replica_agree?.done=http") Then
    
      lblStatus = "Redirecting..."
      lblStatus2 = "Connecting..."
      
      Nav.Navigate "https://" & EditDomain & "/config/change_pw?.redir_from=LOGIN"
      
      DoEvents
      Exit Sub

    'Check password page
    ElseIf InStr(navURL, EditDomain) And LInstr(Expression, "Change your Yahoo! Password") Then
          
          'Wrong new password
          If LInstr(Expression, "new password you entered is not valid") Then
              lblStatus = "Invalid New Password"
              
          'Wrong current password
          ElseIf LInstr(Expression, "Please specify the correct current password") Then
              lblStatus = "Incorrect Current Password"
              
          'Check page, if good -> Change
          ElseIf LInstr(Expression, ".opw") And LInstr(Expression, "ContinueBtn") Then
              lblStatus = "Changing Password..."
              
              'Fill text boxes
              Nav.Document.All(".opw").Value = txtOPW
              Nav.Document.All(".pw1").Value = txtNPW
              Nav.Document.All(".pw2").Value = txtNPW
                
              'MsgBox "Click"
              Nav.Document.All("ContinueBtn").Click
              lblStatus2 = "Please Wait..."
              Exit Sub  '<-- Don't sign out
          End If
      Else
          'Not 999, Or Correct page, Something changed
          lblStatus = "Page Error #1"
      End If
    Else
      lblStatus = "Page Error #2"   'A lot changed
    End If
    
    'Sign out
    Nav.Navigate YLogoutURL
    DoEvents
End Sub
'''''''''''''''''''''''
'''End Process HTML''''




'''''''''Subs''''''''
'''''''''''''''''''''
'Sign out
Private Sub LogoutYahoo()
  lblStatus2 = "Signing out of Yahoo! network"
  DoEvents
  Nav.Navigate YLogoutURL
End Sub

'Default NPW
Private Sub GetDefaultPassword()
  txtNPW = GetSetting("YahooPasswordChanger", "Settings", "DefaultNPW", vbNullString)
End Sub
Private Sub SetDefaultPassword()
  Call SaveSetting("YahooPasswordChanger", "Settings", "DefaultNPW", txtNPW)
End Sub

'Disable Controls While Changing/Processing
Private Sub EnableControls(Enable As Boolean)
  If Enable = True Then
    cmdChange.Enabled = True
    cmdReset.Enabled = True
    
    txtID.Enabled = True
    txtOPW.Enabled = True
    txtNPW.Enabled = True
    
    cboLoginSvr.Enabled = True
    cboEditSvr.Enabled = True
  Else
    cmdChange.Enabled = False
    cmdReset.Enabled = False
    
    txtID.Enabled = False
    txtOPW.Enabled = False
    txtNPW.Enabled = False
    
    cboLoginSvr.Enabled = False
    cboEditSvr.Enabled = False
  End If
End Sub
''''''''''''''''''''
''''''''Subs''''''''




'''''''''Functions''''''''
''''''''''''''''''''''''''
'Generate Login URL
'https ->. sometimes certificate takes a dump
Private Function LoginURL(ID As String, OPW As String, LoginServer As String, EditServer As String) As String
  LoginURL = "http://" & LoginServer & "/config/login?login=" & ID & "&passwd=" & OPW & "&.done=https%3A%2F%2F" & EditServer & "%2Fconfig%2Fchange_pw%3F.redir_from%3DLOGIN"
End Function

'Case-insensitive Instr
Private Function LInstr(String1 As String, String2 As String) As Boolean
  LInstr = InStrB(1, String1, String2, vbTextCompare)
End Function

'Parsing, Domain from URL
Private Function GetBetween(Expression As String, TagA As String, TagB As String) As String
  Dim a As Long, b As Long
  
  a = InStrB(1, Expression, TagA) + LenB(TagA)
  If a = LenB(TagA) Then Exit Function
  
  b = InStrB(a, Expression, TagB)
  If b < a Then Exit Function
  
  GetBetween = MidB$(Expression, a, b - a)
End Function
'''''''''''''''''''''''''''
'''''''End Functions''''''




''''''''''''Start Form''''''''''''
'''''''''''''''''''''''''''''''''''
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then Call cmdChange_Click: KeyAscii = 0
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
  
  Nav.Navigate "about:blank"
  Call GetDefaultPassword
End Sub

'Movable Form
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then: ReleaseCapture: Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then: ReleaseCapture: Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

'Min, Close
Private Sub cmdMin_Click()
  Me.WindowState = vbMinimized
End Sub
Private Sub cmdClose_Click()
  Call Unload(Me)
End Sub

'Select Text
Private Sub txtID_GotFocus()
  txtID.SelStart = 0
  txtID.SelLength = Len(txtID)
End Sub
Private Sub txtOPW_GotFocus()
  txtOPW.SelStart = 0
  txtOPW.SelLength = Len(txtOPW)
End Sub
Private Sub txtNPW_GotFocus()
  txtNPW.SelStart = 0
  txtNPW.SelLength = Len(txtNPW)
End Sub
'''''''''''''''''''''''''''''''''
''''''''''''End Form'''''''''''''
