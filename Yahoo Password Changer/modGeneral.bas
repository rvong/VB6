Attribute VB_Name = "modGeneral"
Option Explicit


'Public
'System Theme Manifest
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Movable Form
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2&
'/Movable Form

'Form on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Const hWnd_TOPMOST = -1
Private Const hWnd_NOTOPMOST = -2
'/Form on top

'Opacity
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Function Opacity(Value As Byte, Frm As Form)
  On Error GoTo ErrorHandler
  
  Dim MaxVal As Byte, MinVal As Byte
      
  MinVal = 20: MaxVal = 255
      
  If Value > MaxVal Then Value = MaxVal
  If Value < MinVal Then Value = MinVal
      
  SetWindowLongA Frm.hWnd, GWL_EXSTYLE, GetWindowLongA(Frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
  SetLayeredWindowAttributes Frm.hWnd, 0, Value, LWA_ALPHA
      
ErrorHandler:     Exit Function
End Function
'/Opacity


'Form on top
Public Function SetOnTop(Form As Form, Optional ByVal GetOnTop As Boolean = True)
  If GetOnTop = True Then
      SetWindowPos Form.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
  Else
      SetWindowPos Form.hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
  End If
End Function

