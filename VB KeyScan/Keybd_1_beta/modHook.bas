Attribute VB_Name = "modKeybdHook"
Option Explicit

Private Const WH_KEYBOARD_LL = 13&     'enables monitoring of keyboard
                                       'input events about to be posted
                                       'in a thread input queue
                                       
Private Const HC_ACTION = 0&           'wParam and lParam parameters
                                       'contain information about a
                                       'keyboard message
Private Type KBDLLHOOKSTRUCT
  vkCode As Long        'a virtual-key code in the range 1 to 254
  scanCode As Long      'hardware scan code for the key
  flags As Long         'specifies the extended-key flag,
                        'event-injected flag, context code,
                        'and transition-state flag
  time As Long          'time stamp for this message
  dwExtraInfo As Long   'extra info associated with the message
End Type

Private Type TESTmeh
  wParam As Long
  lParam As Long
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private m_hDllKbdHook As Long  'private variable holding the handle to the hook procedure

Public vk_collection As Collection

Private KeyCount As Long
Private KeyPress As Long
Private is_vk_down(0 To 255) As Integer

Public Sub HookMe()
  m_hDllKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
  
  If m_hDllKbdHook = 0 Then
    MsgBox "Keyboard hook failed - " & Err.LastDllError
    
    Unload frmMain
    End
  End If
End Sub

Public Sub UnhookMe()
  Call UnhookWindowsHookEx(m_hDllKbdHook)
End Sub

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Static kbdllhs As KBDLLHOOKSTRUCT
  
  If nCode = HC_ACTION Then
  'Debug.Print wParam
    Call CopyMemory(kbdllhs, ByVal lParam, Len(kbdllhs))
    Call ProcessKeys(kbdllhs.vkCode, kbdllhs.flags, kbdllhs.scanCode)
    
    If frmMain.chkBlockKeyboardInput = vbChecked Then LowLevelKeyboardProc = 1: Exit Function
  End If
  
  LowLevelKeyboardProc = CallNextHookEx(m_hDllKbdHook, nCode, wParam, lParam)
End Function

Private Sub ProcessKeys(ByVal pVKey As Integer, ByVal pFlags As Integer, ByVal pScanCode As Integer)
  Dim vk_item() As String
  
  vk_item() = Split(vk_collection.Item(CStr(pVKey)), vbTab)
  
  'Add Data to ListView
  Call AddKeysPressed(vk_item(0), vk_item(1), pVKey & " (0x" & Hex$(pVKey) & ")", pScanCode, pFlags)
  
  'Counting
  If pFlags < 100 Then 'Down, Inc Count
    KeyCount = KeyCount + 1
    frmMain.lblKeyCount.Caption = "Key Count:  " & KeyCount
    
    'Keypress
    If is_vk_down(pVKey) <> 1 Then 'Not Repeat
      is_vk_down(pVKey) = 1
      KeyPress = KeyPress + 1
    End If
  Else
    If is_vk_down(pVKey) <> 0 Then 'UP
      is_vk_down(pVKey) = 0
      KeyPress = KeyPress - 1
    End If
  End If

  frmMain.lblKeyPress = "Key Press:  " & KeyPress
   
  'Highlighting
  'Is Control Index Loaded (Exists)
  If HasIndex(frmMain.shpKey, pVKey) = True Then

    Select Case pVKey
      Case 13 'Num Pad Enter
        If pFlags = 1 Or pFlags = 129 Then pVKey = 1013
    End Select
    
    'Update LEDs
    If frmMain.chkBlockKeyboardInput.Value = vbUnchecked Then
      Select Case pVKey
        Case 20  'Caps Lock
          If pFlags = 0 Then Call UpdateLED(20)
        Case 144 'Num Lock
          If pFlags = 1 Then Call UpdateLED(144)
        Case 145 'Scroll Lock
          If pFlags = 0 Then Call UpdateLED(145)
      End Select
    End If

    
    If pFlags < 100 Then 'Less than 100 = Down, 100+ = Up?
        If frmMain.shpKey(pVKey).BackColor <> &HAAFFAA Then
          frmMain.shpKey(pVKey).BackColor = &HAAFFAA ' &HC0FFC0
          
          'KeyPress = KeyPress + 1
        End If
    Else
        If frmMain.shpKey(pVKey).BackColor <> &HFFFFFF And _
        frmMain.shpKey(pVKey).BackColor <> &HDDDDDD Then
          
          'Key Marker
          If frmMain.chkMarkKeysPressed.Value = vbChecked Then
            frmMain.shpKey(pVKey).BackColor = &HDDDDDD
          Else
            frmMain.shpKey(pVKey).BackColor = &HFFFFFF
          End If
          
          'KeyPress = KeyPress - 1
        End If
    End If
    
    'frmMain.lblKeyPress = "Key Press:  " & KeyPress
  End If

End Sub

Private Function FormatVirtualKey(VirtKey As Integer) As Integer
'
End Function

