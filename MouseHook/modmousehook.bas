Attribute VB_Name = "modMouseHook"
Option Explicit
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_MOUSE = 7
Private hHook As Long
Public Const HC_NOREMOVE = 3
Public Const HC_ACTION = 0

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type MOUSEHOOKSTRUCT
    
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long

End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function MouseProc(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim udtMouseHook As MOUSEHOOKSTRUCT
Dim ptrMOUSEHOOKSTRUCT As Long

    If nCode < 0 Then
    
        'We must call the next hook without any further processing
        MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam)
        
    ElseIf nCode = HC_ACTION Then 'Something happened with the mouse
    
        'Get the address of our MOUSEHOOKSTRUCT
        ptrMOUSEHOOKSTRUCT = VarPtr(udtMouseHook)
        'Fill the udt by copying the memory as the address stored in lParam
        CopyMemory ByVal ptrMOUSEHOOKSTRUCT, ByVal lParam, Len(udtMouseHook)
        
        If wParam = WM_LBUTTONDBLCLK Then
            
            'Just to test, give the handle of the control that was dblclicked on
            Debug.Print "DBL " & udtMouseHook.hwnd
        
        ElseIf wParam = WM_LBUTTONDOWN Then
        
            Debug.Print "Mouse down on : " & udtMouseHook.hwnd

        End If
        
        MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam)
        
    ElseIf nCode = HC_NOREMOVE Then
    
        'In this case, it means that an application used the PeekMessage function to
        'take a look at the messages in the message queue and did not remove the message.
        'If you process this message now, you will receive it again when the application use GetMessage.
        'If you don't want your code to execute twice, ignore this and wait for the other
        MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam)
        
    End If
    
End Function

Public Sub hookmouse()

    hHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, App.hInstance, App.ThreadID)
    
End Sub

Public Sub unhookmouse()

    UnhookWindowsHookEx hHook
    
End Sub
