Attribute VB_Name = "modActivate"
Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Const GWL_WNDPROC = -4
Const HSHELL_WINDOWACTIVATED = 4

Dim myPid As Long, lpWndProc As Long, WM_SHELLHOOK As Long

Sub Setup()
    WM_SHELLHOOK = RegisterWindowMessage("SHELLHOOK")
    lpWndProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WndProc)
    myPid = GetCurrentProcessId
    MsgBox "setup"
End Sub

Private Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If msg = WM_SHELLHOOK Then
        If wParam = HSHELL_WINDOWACTIVATED Then
            Dim S As String * 100, Pid As Long
            GetWindowText lParam, S, Len(S)
            GetWindowThreadProcessId lParam, Pid
            frmMain.Caption = S
            frmMain.BackColor = IIf(Pid = myPid, vbGreen, vbBlack)
        End If
    End If
            Debug.Print wParam
    WndProc = CallWindowProc(lpWndProc, hwnd, msg, wParam, lParam)
End Function
