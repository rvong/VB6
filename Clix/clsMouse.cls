VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean

Public Enum cType
    ClickLeft
    ClickMiddle
    ClickRight
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim PosXY As POINTAPI

Public Sub Click(ByVal cType As cType)
Select Case cType
    Case ClickLeft
        mouse_event &H2, 0&, 0&, 0&, 0&
        mouse_event &H4, 0&, 0&, 0&, 0&
        
    Case ClickMiddle
        mouse_event &H20, 0&, 0&, 0&, 0&
        mouse_event &H40, 0&, 0&, 0&, 0&
        
    Case ClickRight
        mouse_event &H8, 0&, 0&, 0&, 0&
        mouse_event &H10, 0&, 0&, 0&, 0&
End Select
End Sub

Public Sub DoubleClick(ByVal cType As cType)
Select Case cType
    Case ClickLeft
        mouse_event &H2, 0&, 0&, 0&, 0&
        mouse_event &H4, 0&, 0&, 0&, 0&
        mouse_event &H2, 0&, 0&, 0&, 0&
        mouse_event &H4, 0&, 0&, 0&, 0&
        
    Case ClickMiddle
        mouse_event &H20, 0&, 0&, 0&, 0&
        mouse_event &H40, 0&, 0&, 0&, 0&
        mouse_event &H20, 0&, 0&, 0&, 0&
        mouse_event &H40, 0&, 0&, 0&, 0&
        
    Case ClickRight
        mouse_event &H8, 0&, 0&, 0&, 0&
        mouse_event &H10, 0&, 0&, 0&, 0&
        mouse_event &H8, 0&, 0&, 0&, 0&
        mouse_event &H10, 0&, 0&, 0&, 0&
End Select
End Sub

Public Property Get MouseX() As Long
    GetCursorPos PosXY
    MouseX = PosXY.X
End Property

Public Property Get MouseY() As Long
    GetCursorPos PosXY
    MouseY = PosXY.Y
End Property

Public Sub MoveMouse(Horizontal As Long, Vertical As Long)
    mouse_event &H1, Horizontal, Vertical, 0&, 0&
End Sub

Public Sub SetMouseXY(X As Long, Y As Long)
    SetCursorPos X, Y
End Sub
