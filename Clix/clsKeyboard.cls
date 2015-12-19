VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------- Class Name: clsKeys

Option Explicit

Private Declare Function MapVirtualKey Lib "user32" Alias _
   "MapVirtualKeyA" (ByVal wCode As Long, _
   ByVal wMapType As Long) As Long

Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal _
   cChar As Byte) As Integer
   
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
   bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
   
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As _
   Long) As Integer

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Public Enum enumKeys
    keyBackspace = &H8
    keyTab = &H9
    keyReturn = &HD
    
    keyShift = &H10
    keyControl = &H11
    keyAlt = &H12
    
    keyPause = &H13
    keyEscape = &H1B
    keySpace = &H20
    keyEnd = &H23
    keyHome = &H24
    keyLeft = &H25
    KeyUp = &H26
    keyRight = &H27
    KeyDown = &H28
    keyInsert = &H2D
    keyDelete = &H2E
    
    keyF1 = &H70
    KeyF2 = &H71
    keyF3 = &H72
    keyF4 = &H73
    keyF5 = &H74
    keyF6 = &H75
    keyF7 = &H76
    keyF8 = &H77
    keyF9 = &H78
    keyF10 = &H79
    keyF11 = &H7A
    keyF12 = &H7B
    
    keyScrollLock = &H91
    keyCapsLock = &H14
    
    keyNumPad0 = &H60 'Numerickeypad0key
    keyNumPad1 = &H61 'Numerickeypad1key
    keyNumPad2 = &H62 'Numerickeypad2key
    keyNumPad3 = &H63 'Numerickeypad3key
    keyNumPad4 = &H64 'Numerickeypad4key
    keyNumPad5 = &H65 'Numerickeypad5key
    keyNumPad6 = &H66 'Numerickeypad6key
    keyNumPad7 = &H67 'Numerickeypad7key
    keyNumPad8 = &H68 'Numerickeypad8key
    keyNumPad9 = &H69 'Numerickeypad9key

    keyMultiply = &H6A 'Multiplykey
    keyAdd = &H6B 'Addkey
    keySubtract = &H6D 'Subtractkey
    keyDecimal = &H6E 'Decimalkey
    keyDivide = &H6F 'Dividekey
    keyNumLock = &H90 'Presses the single key represented by sKey
End Enum


Public Sub PressKey(sKey As String, Optional bHold As Boolean, Optional _
   bRelease As Boolean)

    Dim nVK As Long
    nVK = VkKeyScan(Asc(sKey))
    
    If nVK = 0 Then
        Exit Sub
    End If
    
    Dim nScan As Long
    Dim nExtended As Long
        
    nScan = MapVirtualKey(nVK, 2)
    nExtended = 0
    If nScan = 0 Then
        nExtended = KEYEVENTF_EXTENDEDKEY
    End If
    nScan = MapVirtualKey(nVK, 0)
    
    Dim bShift As Boolean
    Dim bCtrl As Boolean
    Dim bAlt As Boolean
    
    bShift = (nVK And &H100)
    bCtrl = (nVK And &H200)
    bAlt = (nVK And &H400)
    
    nVK = (nVK And &HFF)
    
    If Not bRelease Then
        If bShift Then
            keybd_event enumKeys.keyShift, 0, 0, 0
        End If
        If bCtrl Then
            keybd_event enumKeys.keyControl, 0, 0, 0
        End If
        If bAlt Then
            keybd_event enumKeys.keyAlt, 0, 0, 0
        End If
    
        keybd_event nVK, nScan, nExtended, 0
    End If
    
    If Not bHold Then
        keybd_event nVK, nScan, KEYEVENTF_KEYUP Or nExtended, 0
    
        If bShift Then
            keybd_event enumKeys.keyShift, 0, KEYEVENTF_KEYUP, 0
        End If
        If bCtrl Then
            keybd_event enumKeys.keyControl, 0, KEYEVENTF_KEYUP, 0
        End If
        If bAlt Then
            keybd_event enumKeys.keyAlt, 0, KEYEVENTF_KEYUP, 0
        End If
    End If

End Sub

'Loop through a string and calls PressKey for each character (Does not
' parse strings like SendKeys)
Public Sub PressString(ByVal sString As String, Optional bDoEvents As Boolean = True)

    Do While sString <> vbNullString '""
        PressKey Mid(sString, 1, 1)
        
        'Sleep 20
        If bDoEvents Then
            DoEvents
        End If
        
        sString = Mid(sString, 2)
    Loop

End Sub

'Presses a specific key (this is used for keys that don't have a
' ascii equilivant)
Public Sub PressKeyVK(keyPress As enumKeys, Optional bHold As Boolean, _
   Optional bRelease As Boolean, Optional bCompatible As Boolean)
    
    Dim nScan As Long
    Dim nExtended As Long
        
    nScan = MapVirtualKey(keyPress, 2)
    nExtended = 0
    If nScan = 0 Then
        nExtended = KEYEVENTF_EXTENDEDKEY
    End If
    nScan = MapVirtualKey(keyPress, 0)
    
    If bCompatible Then
        nExtended = 0
    End If
    
    If Not bRelease Then
        keybd_event keyPress, nScan, nExtended, 0
    End If
    
    If Not bHold Then
        keybd_event keyPress, nScan, KEYEVENTF_KEYUP Or nExtended, 0
    End If

End Sub

'Returns (in the boolean variables) the status of the various Lock keys
Public Sub GetLockStatus(bCapsLock As Boolean, bNumLock As Boolean, _
   bScrollLock As Boolean)

    bCapsLock = GetKeyState(enumKeys.keyCapsLock)
    bNumLock = GetKeyState(enumKeys.keyNumLock)
    bScrollLock = GetKeyState(enumKeys.keyScrollLock)

End Sub

