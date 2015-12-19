Attribute VB_Name = "DoEventsX"
Option Explicit

Public Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long

Public Const QS_HOTKEY As Long = &H80
Public Const QS_KEY As Long = &H1
Public Const QS_MOUSEBUTTON As Long = &H4
Public Const QS_PAINT As Long = &H20

Public Sub DoEventsX()
If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then DoEvents
End Sub
