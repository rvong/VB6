VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalKeyHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As KBDLLHOOKSTRUCT) As Long

'Monitor keyboard input events about to be posted in a thread input queue
'wParam and lParam parameters contain information about a keyboard message
Private Const WH_KEYBOARD_LL = 13

'Contains information about a low-level keyboard input event
Private Type KBDLLHOOKSTRUCT
    vkCode As Long                      'Virtual-key code, 1 to 254
    ScanCode As Long                    'Hardware scan code for key
    Flags As Long                       'Keystroke flags
    Time As Long                        'Time stamp for msg
    dwExtraInfo As Long                 'Extra info associated with msg
End Type

'Keybd msg identifier in wParam
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104     'Sys Key = ALT + Other Key
Private Const WM_SYSKEYUP = &H105

'Keystroke flags
Private Const LLKHF_EXTENDED = &H1       'Extended-key flag
Private Const LLKHF_INJECTED = &H10      'Event-injected flag
Private Const LLKHF_ALTDOWN = &H20       'Context code
Private Const LLKHF_UP = &H80            'Transition-state flag

Private hKBHook As Long                 'Handle to the hook procedure

Public Function HookMe() As Boolean
    hKBHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
    HookMe = hKBHook '0 = Failed
    Debug.Print hKBHook
End Function

Public Sub UnhookMe()
    If (hKBHook <> 0) Then Call UnhookWindowsHookEx(hKBHook)
End Sub

Private Function LowLevelKeyboardProc(nCode As Long, wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    If (nCode = 0) Then
        Static kbdllhs As KBDLLHOOKSTRUCT
        kbdllhs = lParam
        
        Select Case wParam  'Keybd msg identifier 'kbdllhs.Time is imprecise, use GetPerfCount
            Case WM_KEYDOWN, WM_SYSKEYDOWN
                'If kbdllhs.Flags = LLKHF_INJECTED Then frmMain.List1.AddItem kbdllhs.vkCode
                'Debug.Print "bit:" & RBitShift(kbdllhs.Flags, 4)
                'If kbdllhs.Flags = LLKHF_EXTENDED Then Debug.Print "ext1"
                'Debug.Print "key down"
                'Call frmMain.ProcKeyDown(kbdllhs.vkCode, kbdllhs.ScanCode, kbdllhs.Flags, GetPerfCount)
                'Debug.Print GetKeyText(kbdllhs.ScanCode, kbdllhs.Flags)
            Case WM_KEYUP, WM_SYSKEYUP
                If kbdllhs.Flags = (LLKHF_UP + LLKHF_EXTENDED) Then Debug.Print "ext2"
                'Debug.Print "bit:" & RBitShift(kbdllhs.Flags, 6)
                Debug.Print kbdllhs.vkCode
                'Debug.Print "key up"
                'Call frmMain.ProcKeyUp(kbdllhs.vkCode, kbdllhs.ScanCode, kbdllhs.Flags, GetPerfCount)
        End Select
    
    'Block Input --> LowLevelKeyboardProc = 1: Exit Function
    End If
    
    LowLevelKeyboardProc = CallNextHookEx(hKBHook, nCode, wParam, lParam)

End Function
