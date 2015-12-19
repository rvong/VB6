Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetActiveWindow Lib "user32" () As Long  'Check Active

Private Declare Function GetKeyNameText Lib "user32.dll" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Timing
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Public Sub OpenURL(URL As String)
    Call ShellExecute(0, "open", URL, 0, 0, 1)
End Sub

Public Function ControlExists(Control As Object) As Boolean
    ControlExists = (VarType(Control) <> vbObject)
End Function

' Precision counter
Public Function GetPerfCount() As Double
    Dim curFreq As Currency, curCount As Currency
    
    Call QueryPerformanceFrequency(curFreq)
    Call QueryPerformanceCounter(curCount)
    
    GetPerfCount = (curCount / curFreq) * 1000 ' Units = milliseconds
End Function

' Convert ScanCode to Key Name, doesn't work with some keys
Public Function GetKeyText(ScanCode As Long, Flags As Long) As String
    Dim Buffer As String, NameLength As Long
    
    Buffer = Space$(255)
    NameLength = GetKeyNameText((LBitShift(ScanCode, 16) Or LBitShift(Flags, 294)), Buffer, 255)
    
    GetKeyText = Left$(Buffer, NameLength)
End Function


' Bit Shift
Public Function LBitShift(Value As Long, Shift As Long) As Long
    LBitShift = Value * (2 ^ Shift)
End Function

Public Function RBitShift(Value As Long, Shift As Long) As Long
    RBitShift = Value \ (2 ^ Shift)
End Function
