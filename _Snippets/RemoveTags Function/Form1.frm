VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   5040
      TabIndex        =   1
      Top             =   7680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Text1.Text = RemoveAllTags(Text1.Text, "<", ">")
Text1.Text = ParseIt(Text1.Text, "<img", ">")
End Sub




Function RemoveTag(str As String, tagA As String, tagB As String) As String
  Dim A As Long, B As Long
  
  A = InStr(1, str, tagA)
  
  If A > 0 Then
    B = InStr(A, str, tagB)
    
    If B > A Then
      RemoveTag = Mid$(str, 1, A - Len(tagA)) & Mid$(str, B + Len(tagB), Len(str) - (B))
    Else
      RemoveTag = str
    End If
    
  Else
    RemoveTag = str
  End If
End Function


Function RemAllTags(str As String, tagA As String, tagB As String) As String

'do until instr(str,
End Function

Function RemoveAllTags(str As String, tagA As String, tagB As String) As String
  Dim A As Long, B As Long
  
  A = 1
  B = 1
  
  Do Until A = 0
    A = InStr(A, str, tagA)
    
    If A > 0 Then
      B = InStr(A, str, tagB)
      
      If B > A Then
        str = Mid$(str, A, A - Len(tagA)) & Mid$(str, B + Len(tagB), Len(str) - (B))
      End If
    End If
  Loop
  
  RemoveAllTags = str
End Function


Public Function ParseIt(Expr As String, tagA As String, tagB As String) As String
Dim A As Long, B As Long
A = 1

While InStrB(A, Expr, tagA) > 0
  A = InStrB(A, Expr, tagA)
  B = InStrB(A, Expr, tagB) + LenB(tagB)
  Expr = Replace$(Expr, MidB$(Expr, A, B - A), vbNullString)
Wend

ParseIt = Expr
End Function
