VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Binary Search"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   7440
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search!"
      Height          =   435
      Left            =   4680
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Items"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   7800
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   7440
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Function SearchList(ToSearch As String, lstList As ListBox) As Integer
Const LB_FINDSTRINGEXACT = &H1A2
SearchList = SendMessage(lstList.hwnd, LB_FINDSTRINGEXACT, 0&, ByVal ToSearch)
End Function

Private Sub Command3_Click()
Debug.Print BinarySearch(Text1, List1)
End Sub

Private Sub Form_Load()
Randomize
End Sub

Public Function RandomNumber(ByVal MaxValue As Long, Optional ByVal MinValue As Long = 0)
  On Error Resume Next
  Randomize Timer
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue
End Function

Private Sub Command1_Click()
Dim x As Long
List1.Visible = False
List2.Visible = False

For x = 1 To 300
  List1.AddItem RandomNumber(10000, 1)
  List2.AddItem x
Next x

List1.Visible = True
List2.Visible = True
End Sub

Private Function BinarySearch(Expression As String, ListBox As ListBox) As Long
  Dim minPos As Long, midPos As Long, maxPos As Long
  Dim swapPos As Long, lastPos As Long
  Dim currentItem As String

  If ListBox.ListCount = 0 Then BinarySearch = -1: Exit Function
  maxPos = ListBox.ListCount
  midPos = maxPos / 2
  
  Do
    currentItem = ListBox.List(midPos)
    swapPos = midPos
    
    If Expression > currentItem Then         'Go Higher
      midPos = (midPos + maxPos) / 2
      minPos = swapPos
    ElseIf Expression < currentItem Then     'Go Lower
      midPos = (midPos + minPos) / 2
      maxPos = swapPos
    Else                                     'Found
      BinarySearch = midPos: Exit Do
    End If
    
    If midPos = lastPos Then BinarySearch = -1: List1.ListIndex = midPos: Exit Do 'Not Found
    lastPos = midPos
  Loop
End Function

Private Sub List1_Scroll()
    List2.TopIndex = List1.TopIndex
End Sub
Private Sub List2_Scroll()
    List1.TopIndex = List2.TopIndex
End Sub
Private Sub List1_Click()
  List2.ListIndex = List1.ListIndex
End Sub
Private Sub List2_Click()
  List1.ListIndex = List2.ListIndex
End Sub

