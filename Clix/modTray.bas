Attribute VB_Name = "modTray"
Option Explicit

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public NID As NOTIFYICONDATA

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
       
Public Sub CreateTray(frmForm As Form)
    With NID
        .cbSize = Len(NID)
        .hWnd = frmForm.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmForm.Icon
        .szTip = frmForm.Caption & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_ADD, NID
    'frmForm.mnufile.Visible = False
    App.TaskVisible = False
End Sub

Public Sub CreatePopMenu(frmForm As Form, X As Single)
    Dim Msg As Long
    
    Msg = X / Screen.TwipsPerPixelX
    
    Select Case Msg
        Case WM_LBUTTONDBLCLK:
            frmForm.WindowState = 0
            SetForegroundWindow frmForm.hWnd
            frmForm.Visible = True
        Case WM_RBUTTONUP:
            SetForegroundWindow frmForm.hWnd
            frmForm.PopupMenu frmForm.mnuSystemTray
    End Select
End Sub

Public Sub DeleteTray()
    Shell_NotifyIcon NIM_DELETE, NID
End Sub

Public Sub SendToTray(frmForm As Form)
    If frmForm.WindowState = 1 Then frmForm.Visible = False
End Sub

Public Sub SendToDesktop(frmForm As Form)
    If frmForm.Visible = False Then
        frmForm.WindowState = 0
        frmForm.Visible = True
    End If
End Sub
