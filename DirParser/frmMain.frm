VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dir Parser"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   9720
      TabIndex        =   4
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3720
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strDir As String

Private Function GetWithTags(Expression As String, TagA As String, TagB As String) As String
    Dim a As Long, b As Long
    
    a = InStrB(1, Expression, TagA)
    If (a = 0) Then Exit Function
    
    b = InStrB(a, Expression, TagB) + LenB(TagB)
    If (b > a) Then GetWithTags = MidB$(Expression, a, b - a)
End Function

Private Function GetWithoutTags(Expression As String, TagA As String, TagB As String) As String
    Dim a As Long, b As Long
    
    a = InStrB(1, Expression, TagA) + LenB(TagA)
    If a = LenB(TagA) Then Exit Function
    
    b = InStrB(a, Expression, TagB)
    If (b > a) Then GetWithoutTags = MidB$(Expression, a, b - a)
End Function

Private Function GetBetweenTags(Expression As String, TagA As String, TagB As String, Optional IncludeTags As Boolean = False) As String
    Dim a As Long, b As Long
    
    a = InStrB(1, Expression, TagA)
    If (a = 0) Then Exit Function
    If IncludeTags = False Then a = a + LenB(TagA)

    b = InStrB(a, Expression, TagB)
    If IncludeTags = True Then b = b + LenB(TagB)
    
    If (b > a) Then GetBetweenTags = MidB$(Expression, a, b - a)
End Function

Private Sub Command1_Click()
    'Text1 = GetWithTags(Text1, " Directory of", "bytes")
    'strDir = GetWithoutTags(Text1, " Directory of", vbNewLine)
    'Call GetLines(Text1, List1)
    Call GetSection(Text1, " Directory of", "bytes", False)
End Sub

Private Sub GetSection(Expression As String, TagA As String, TagB As String, Optional IncludeTags As Boolean = False)
    Dim a As Long, b As Long
    a = 1
    
    Do
        a = InStrB(a, Expression, TagA)
        If (a = 0) Then Exit Do
        If IncludeTags = False Then a = a + LenB(TagA)
        
        b = InStrB(a, Expression, TagB)
        If (b <= a) Then Exit Do
        If IncludeTags = True Then b = b + LenB(TagB)
        
        Text2 = Text2 & MidB$(Expression, a, b - a)
        DoEvents
        'a = b
        'Call GetLines(MidB$(Expression, a, b - a), List1)
    Loop Until a = 0 Or b < a
End Sub

Private Sub GetLines(Expression As String, ListBox As ListBox)
    Dim a As Long, b As Long
    Dim IsLastItem As Boolean
    a = 1
    
    Do
        If (b > 1) Then a = InStrB(a, Expression, vbNewLine) + 4 'LenB(vbNewLine)
        If (a = 4) Then Exit Do '0 -> 4
        b = InStrB(a, Expression, vbNewLine)
      
        If (b = 0) Then b = LenB(Expression): IsLastItem = True
        If (b < a) Then Exit Do
        
        'Text2.Text = Text2.Text & vbCrLf & MidB$(Expression, a, b - a)
        Call AddList(MidB$(Expression, a, b - a), List1)
    Loop Until IsLastItem = True
End Sub

Private Sub AddList(Expression As String, ListBox As ListBox) 'Parse Line
    Const DateLength = 10 'DateBegin = First Item, Using Left$
    Const TimeBegin = 13, TimeLength = 8
    Const TypeBegin = 25, TypeLength = 15
    Const NameBegin = 40  'NameLength = Len(Expression)
    
    Dim strDate As String, strTime As String, strType As String, strName As String
    
    If InStr(Expression, "PM") > 0 Or InStr(Expression, "AM") > 0 Then
    
        strDate = Left$(Expression, DateLength)
        strTime = Mid$(Expression, TimeBegin, TimeLength)
        strType = Mid$(Expression, TypeBegin, TypeLength)
        strName = Mid$(Expression, NameBegin, Len(Expression) - NameBegin)
        
        List1.AddItem Expression
        'List1.AddItem strDir & "\" & strName
        'List1.AddItem strDate & "  " & strTime & "  " & strType & "  " & strName
    End If
End Sub

Private Sub Command2_Click()
MsgBox GetBetweenTags("<a>asdf</a>", "<a>", "</a>", False)
End Sub
