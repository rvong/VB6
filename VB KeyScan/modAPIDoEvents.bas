Attribute VB_Name = "modAPIDoEvents"
Option Explicit

Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

' API Constants and declare
' Constants used by GetQueueStatus API function
Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSEMOVE = &H2
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8
Private Const QS_SENDMESSAGE = &H40
Private Const QS_TIMER = &H10
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const QS_MESSAGES = (QS_POSTMESSAGE Or QS_SENDMESSAGE)           ' Not MS standard constant
Private Const QS_STANDARD = (QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT)   ' Not MS standard constant

Public Sub APIDoEvents()
If GetQueueStatus(QS_STANDARD + QS_MESSAGES) <> 0 Then DoEvents ' Speciy Events
' If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then DoEvents
End Sub


