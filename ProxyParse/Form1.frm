VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Hello"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Parse"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Count:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim x As Integer
On Error GoTo Error
With CommonDialog1
    .DialogTitle = "Save As"
    .Filter = "Text Document (*.txt)|*.txt"
    .ShowSave
    Open .FileName For Output As #1
        For x = 0 To List1.ListCount - 1
        Print #1, List1.List(x)
        Next x
    Close #1
End With
Error: Exit Sub
End Sub

Private Sub Command2_Click()
If List1.ListCount > 0 Then List1.Clear
End Sub

Private Sub Command3_Click()
Winsock1.Close
Winsock1.Connect "www.cybersyndrome.net", "80"
End Sub

Private Sub Winsock1_Connect()
Dim HTTP_Thingy As String

HTTP_Thingy = "GET /pld.html HTTP/1.0" & vbCrLf
HTTP_Thingy = HTTP_Thingy & "Host: www.cybersyndrome.net" & vbCrLf
HTTP_Thingy = HTTP_Thingy & "Accept: */*" & vbCrLf
HTTP_Thingy = HTTP_Thingy & "Accept-Language: en-us, en" & vbCrLf
HTTP_Thingy = HTTP_Thingy & "Accept-Encoding: gzip, deflate" & vbCrLf
HTTP_Thingy = HTTP_Thingy & "Connection: keep-alive" & vbCrLf & vbCrLf

Winsock1.SendData HTTP_Thingy
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim DaData As String
Dim A As Long, B As Long

Winsock1.GetData DaData, vbString, bytesTotal
RichTextBox1.Text = RichTextBox1.Text & DaData

If InStr(RichTextBox1.Text, "</html>") Then
Winsock1.Close
ParseIt RichTextBox1.Text, "onMouseOut=" & Chr(34) & "d" & Chr(40) & Chr(41) & Chr(34) & ">", "</a>", List1
Label1.Caption = "Count: " & List1.ListCount
End If
End Sub

Function ParseIt(Expression As String, DelimiterA As String, DelimiterB As String, LstBx As ListBox)
Dim A As Long, B As Long
A = 1
While InStr(A, Expression, DelimiterA) > 0
  A = InStr(A, Expression, DelimiterA) + Len(DelimiterA)
  B = InStr(A, Expression, DelimiterB)
  LstBx.AddItem Mid$(Expression, A, B - A)
Wend
End Function
