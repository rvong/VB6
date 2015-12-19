Attribute VB_Name = "modKeybdHook"
Option Explicit

'Monitor keyboard input events about to be posted in a thread input queue
Private Const WH_KEYBOARD_LL = 13&


'Keyboard msg identifier in wParam
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105


'wParam and lParam parameters contain information about a keyboard message
Private Const HC_ACTION = 0&
                            
Private Type KBDLLHOOKSTRUCT
  vkCode As Long                       'Virtual-key code, 1 to 254
  scanCode As Long                     'Hardware scan code for key
  flags As Long                        'Keystroke flags
  time As Long                         'Time stamp for msg
  dwExtraInfo As Long                  'Extra info associated with msg
End Type

Private Const LLKHF_EXTENDED = &H1     'Extended-key flag
Private Const LLKHF_INJECTED = &H10    'Event-injected flag
Private Const LLKHF_ALTDOWN = &H20     'Context code
Private Const LLKHF_UP = &H80          'Transition-state flag

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)

'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private m_hDllKbdHook As Long          'Handle to the hook procedure

Public vk_collection As Collection

Private KeyCount As Long
Private KeyPress As Long
Private is_vk_down(0 To 255) As Integer

Public Sub HookMe()
  m_hDllKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
  
  If m_hDllKbdHook = 0 Then
    MsgBox "Low-level keyboard hook failed - " & Err.LastDllError
    
    Unload frmMain
    End
  End If
End Sub

Public Sub UnhookMe()
  If m_hDllKbdHook <> 0 Then Call UnhookWindowsHookEx(m_hDllKbdHook)
End Sub

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Static kbdllhs As KBDLLHOOKSTRUCT
  
  If nCode = HC_ACTION Then
    Call CopyMemory(kbdllhs, ByVal lParam, Len(kbdllhs))
    
    Debug.Print wParam
    
    Select Case wParam
    End Select
    
    Debug.Print kbdllhs.vkCode & vbNewLine & kbdllhs.scanCode & kbdllhs.flags & kbdllhs.time & kbdllhs.dwExtraInfo
    
    'LowLevelKeyboardProc = 1: Exit Sub  'Block Input
  End If
  
  LowLevelKeyboardProc = CallNextHookEx(m_hDllKbdHook, nCode, wParam, lParam)
End Function


