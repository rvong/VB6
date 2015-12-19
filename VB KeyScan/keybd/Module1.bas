Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    
Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" ( _
    ByVal wCode As Long, ByVal wMapType As Long) As Long
        
Private Declare Function GetKeyNameText Lib "user32.dll" Alias "GetKeyNameTextA" ( _
    ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const WH_KEYBOARD_LL = 13
Private Const HC_ACTION = 0

Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    ScanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private hHook As Long
Private IsHooked As Boolean
Public cButton As CommandButton

Public Sub SetKeyboardHook()
    If Not IsHooked Then
        hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
        IsHooked = True
    End If
End Sub

Public Sub RemoveKeyboardHook()
    If IsHooked Then
        UnhookWindowsHookEx hHook
        IsHooked = False
    End If
End Sub

Private Function LowLevelKeyboardProc(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    Dim sBuffer As String, lRet As Long
    If uCode >= 0 Then
        If uCode = HC_ACTION Then
            'Debug.Print "asdf"
            Form1.List1.AddItem GetKeyText(lParam.ScanCode, lParam.Flags)
            LowLevelKeyboardProc = 1
            Exit Function
        End If
    End If
    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)
End Function

Public Function GetKeyText(ScanCode As Long, Flags As Long) As String
  Dim Buffer As String, lRet As Long
  On Error GoTo Error
  
  Buffer = Space$(255)
  lRet = GetKeyNameText(RBitShift(ScanCode, Val(Form1.Text2.Text)) Or RBitShift(Flags, Val(Form1.Text1.Text)), Buffer, 255&)
  
  GetKeyText = Left$(Buffer, lRet)
  
  Exit Function
Error: Debug.Print Err.Number & " - " & Err.Description
End Function

Function RBitShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then RBitShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    RBitShift = (lNum And (2 ^ (31 - lBits) - 1)) * IfLng(lBits = 31, &H80000000, 2 ^ lBits) _
                 Or IfLng((lNum And 2 ^ (31 - lBits)) = 2 ^ (31 - lBits), &H80000000, 0)
End Function

Public Function LBitShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then LBitShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    If lNum < 0 Then
        LBitShift = (lNum And &H7FFFFFFF) \ (2 ^ lBits) Or 2 ^ (31 - lBits)
    Else
        LBitShift = lNum \ (2 ^ lBits)
    End If
End Function

Public Function IfByte(ByVal Expression As Boolean, ByVal TruePart As Byte, ByVal FalsePart As Byte) As Byte
    If Expression Then IfByte = TruePart Else IfByte = FalsePart
End Function

Public Function IfInt(ByVal Expression As Boolean, ByVal TruePart As Integer, ByVal FalsePart As Integer) As Integer
    If Expression Then IfInt = TruePart Else IfInt = FalsePart
End Function

Public Function IfLng(ByVal Expression As Boolean, ByVal TruePart As Long, ByVal FalsePart As Long) As Long
    If Expression Then IfLng = TruePart Else IfLng = FalsePart
End Function

Public Function IfVar(ByVal Expression As Boolean, ByVal TruePart As Variant, ByVal FalsePart As Variant) As Variant
    If Expression Then IfVar = TruePart Else IfVar = FalsePart
End Function

