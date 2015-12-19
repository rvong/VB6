Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'this file needs to be on the comp
'XP MANIFEST ===========================
'Private Sub Form_Initialize()
'    InitCommonControls
'End Sub

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function Beep Lib "Kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Function GetKeyState(Key As Integer) As Boolean
    GetKeyState = CBool(GetAsyncKeyState(Key)) '192 = ~
End Function

Public Function SetOnTop(hWin As Long, Optional ByVal GetOnTop As Boolean = True)
If GetOnTop = True Then
    SetWindowPos hWin, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
Else
    SetWindowPos hWin, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End If
End Function

Public Sub Pause(ByVal Seconds As Double, Optional Milliseconds As Boolean = False)
    Dim cTime As Double
    Dim Count As Integer
    
    If Seconds = 0 Then Exit Sub
    If Milliseconds = True Then Seconds = Val(Seconds) * 0.001
    
    cTime = Timer
    Do While Timer < cTime + Seconds
        Count = Count + 1
            
        If Count = 1000 Then
            Count = 0
            Sleep 1
            DoEvents
        End If
    Loop
End Sub

Public Function FileExists(FilePath As String) As Boolean
    If Len(FilePath) = 0 Then
        FileExists = False
    Else
        If Len(Dir$(FilePath)) > 0 Then
            If (GetAttr(FilePath) And vbDirectory) <> vbDirectory Then
                FileExists = True
            Else
                FileExists = False
            End If
        Else
            FileExists = False
        End If
    End If
End Function

Public Function DirectoryExists(DirectoryPath As String) As Boolean
    If Len(DirectoryPath) = 0 Then
        DirectoryExists = False
    Else
        If Len(Dir$(DirectoryPath, vbDirectory)) > 0 Then
            If (GetAttr(DirectoryPath) And vbDirectory) = vbDirectory Then
                DirectoryExists = True
            Else
                DirectoryExists = False
            End If
        Else
            DirectoryExists = False
        End If
    End If
End Function

Public Sub Beeper(Frequency As Integer, Duration As Integer, IsOn As Boolean)
If IsOn = True Then
    Beep Frequency, Duration
    Beep Frequency + 500, Duration
Else
    Beep Frequency + 500, Duration
    Beep Frequency - 500, Duration
End If
End Sub

Public Sub OpenBrowser(hWindow As Long, URL As String)
ShellExecute hWindow, "OPEN", URL, vbNullString, vbNullString, vbNormalFocus
End Sub
