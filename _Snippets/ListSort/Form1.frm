VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuicksort 
      Caption         =   "Quicksort"
      Height          =   735
      Left            =   4800
      TabIndex        =   11
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdShell 
      Caption         =   "Shell Sort"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelection 
      Caption         =   "Selection Sort"
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsertion 
      Caption         =   "Insertion Sort"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CommandButton cmdShaker 
      Caption         =   "Shaker Sort"
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   5520
      ItemData        =   "Form1.frx":0000
      Left            =   7560
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List Sort"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBubble 
      Caption         =   "Bubble Sort"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   3960
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim t As Double

Private Sub cmdClear_Click()
List1.Clear
End Sub

Private Sub cmdInsertion_Click()
Dim x As Long, j As Long
Dim temp As String

t = Timer

'start with second item
For x = 1 To List1.ListCount - 1
  temp = List1.List(x) 'swap when ready
  j = x - 1 'item before current
  
  While j >= 0 And List1.List(j) > temp 'move items higher in list, until found item < temp
    List1.List(j + 1) = List1.List(j)
    j = j - 1
  Wend

  List1.List(j + 1) = temp  'stick temp in the correct spot
Next x   'repeat for next item

MsgBox Timer - t
End Sub


Private Sub cmdQuicksort_Click()
  t = Timer
  
  'List1.Visible = False
  
  Call QuickSort(List1, 0, List1.ListCount - 1)
  
  'List1.Visible = True
  
  MsgBox Timer - t
End Sub


Private Sub QuickSort(ListBox As ListBox, indexLow As Long, indexHigh As Long)
Dim Low As Long, High As Long
Dim Pivot As String, Swap As String

Low = indexLow
High = indexHigh

Pivot = ListBox.List((indexLow + indexHigh) \ 2)

While (Low <= High)

    While (ListBox.List(Low) < Pivot) And (Low < indexHigh)
      Low = Low + 1
    Wend
    
    While (ListBox.List(High) > Pivot) And (High > indexLow)
      High = High - 1
    Wend
  
    If (Low <= High) Then
      Swap = ListBox.List(Low)
      ListBox.List(Low) = ListBox.List(High)
      ListBox.List(High) = Swap
      
      Low = Low + 1
      High = High - 1
    End If
    
    'Sleep 100
    DoEvents
    
   If (indexLow < High) Then Call QuickSort(ListBox, indexLow, High)
   If (Low < indexHigh) Then Call QuickSort(ListBox, Low, indexHigh)
Wend
End Sub


Private Sub QuickSort(Arr() As String, lBnd As Long, uBnd As Long)
    Dim Lo As Long, Hi As Long
    Dim Pivot As String, SwapBuffer As String
    
    Lo = lBnd
    Hi = uBnd
    Pivot = Arr((Lo + Hi) \ 2)
    
    While (Lo <= Hi)
        While (Arr(Lo) < Pivot) And (Lo < uBnd)
            Lo = Lo + 1
        Wend
        
        While (Arr(Hi) > Pivot) And (Hi > lBnd)
            Hi = Hi - 1
        Wend
      
        If (Lo <= Hi) Then
            SwapBuffer = Arr(Lo)
            Arr(Lo) = Arr(Hi)
            Arr(Hi) = SwapBuffer
          
            Lo = Lo + 1
            Hi = Hi - 1
        End If
        
       If (lBnd < Hi) Then Call QuickSort(Arr(), lBnd, Hi)
       If (Lo < uBnd) Then Call QuickSort(Arr(), Lo, uBnd)
    Wend
End Sub

'''Looks through the list, finds the smallest value, swaps with first item,
'''then finds second smallest value (which is then the smallest avail,
'''doesn't look at first item again) and swaps with second item in list.
'''and so on...
Private Sub cmdSelection_Click()
Dim a As Long, b As Long, i As Long
Dim s As String

t = Timer

For a = 0 To List1.ListCount - 1
  i = a
  
  For b = a + 1 To List1.ListCount - 1
    If List1.List(b) < List1.List(i) Then i = b
  Next b
  
  If a <> i Then
    s = List1.List(a)
    List1.List(a) = List1.List(i)
    List1.List(i) = s
  End If
Next a

MsgBox Timer - t
End Sub

Private Sub cmdShaker_Click()
Dim x As Long, y As Long
Dim temp As String
Dim swapped As Boolean

t = Timer

Do
  swapped = False
  
  For x = 0 To List1.ListCount - 2 'Forwards
    If List1.List(x) > List1.List(x + 1) Then
      temp = List1.List(x + 1)
      
      List1.List(x + 1) = List1.List(x)
      List1.List(x) = temp
      If swapped = False Then swapped = True
    End If
  Next x
  
  If swapped = False Then Exit Do
  
  For x = List1.ListCount - 2 To 0 Step -1 'Reverse
    If List1.List(x) < List1.List(x - 1) Then
      temp = List1.List(x - 1)
      List1.List(x - 1) = List1.List(x)
      List1.List(x) = temp
      
      If swapped = False Then swapped = True
    End If
  Next x
Loop Until swapped = False

MsgBox Timer - t
End Sub

Private Sub cmdBubble_Click()
Dim x As Long, y As Long
Dim temp As String
Dim swapped As Boolean

t = Timer

Do
  swapped = False

  For x = 0 To List1.ListCount - 2
    If List1.List(x) > List1.List(x + 1) Then
      temp = List1.List(x + 1)
      List1.List(x + 1) = List1.List(x)
      List1.List(x) = temp
      
      If swapped = False Then swapped = True
    End If
  Next x
Loop Until swapped = False

MsgBox Timer - t
End Sub

Private Sub cmdShell_Click()
Dim inc As Long, i As Long, j As Long
Dim temp As String

t = Timer

inc = Round(List1.ListCount / 2)

While inc > 0
  For i = inc To List1.ListCount - 1
    temp = List1.List(i)
    j = i
    
    While j >= inc And List1.List(j - inc) > temp
      List1.List(j) = List1.List(j - inc)
      j = j - inc
    Wend
    
    List1.List(j) = temp
  Next i
inc = Round(inc / 2.2)
Wend

MsgBox Timer - t
End Sub

Private Sub Command1_Click()
ShellSort2 List1
End Sub

Private Sub Command2_Click()
Dim x As Integer

Randomize Timer
For x = 0 To 100
  List1.AddItem Round(((9 - 1) * Rnd) + 1, 0)
Next x
End Sub

Private Sub cmdList_Click()
Dim x As Long

t = Timer

List1.Visible = False

For x = 0 To List1.ListCount - 1
  List2.AddItem List1.List(x)
Next x

List1.Clear

For x = 0 To List2.ListCount - 1
  List1.AddItem List2.List(x)
Next x

List2.Clear

List1.Visible = True

MsgBox Timer - t
End Sub

Private Sub Command5_Click()
Dim x As Integer
For x = 10 To 1 Step -1
  Debug.Print x
Next x
End Sub


Private Sub ShellSortNumbers(varray As Variant)
   Dim cnt As Long, tmp As Long
   Dim nHold As Long, nHValue As Long

   nHValue = LBound(varray)
   
   Do
      nHValue = 3 * nHValue + 1
   Loop Until nHValue > UBound(varray)

   Do
      nHValue = nHValue / 3
      
      For cnt = nHValue + LBound(varray) To UBound(varray)
         tmp = varray(cnt)
         nHold = cnt

         Do While varray(nHold - nHValue) > tmp
            varray(nHold) = varray(nHold - nHValue)
            nHold = nHold - nHValue

            If nHold < nHValue Then Exit Do
         Loop

         varray(nHold) = tmp
      Next cnt
   Loop Until nHValue = LBound(varray)
End Sub


Private Sub ShellSort2(ListBox As ListBox)
  Dim i As Long, h As Long, v As Long
  Dim s As String
  
  t = Timer
  
  Do
    v = 3 * v + 1
  Loop Until v > List1.ListCount - 1
  
  Do
    v = v / 3
    
    For i = v To List1.ListCount - 1
      s = List1.List(i)
      h = i
      
      Do While List1.List(h - v) > s
        List1.List(h) = List1.List(h - v)
        h = h - v
        If h < v Then Exit Do
      Loop
      
      List1.List(h) = s
    Next i
  Loop Until v = 0
  
  MsgBox Timer - t
End Sub

