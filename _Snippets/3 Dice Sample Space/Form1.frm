VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3 Dice Sample Space (0-18)"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear List"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "0"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List Sample Space"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Probability:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Sample Space Count:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Dice Val"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'June 18, 2009
'High School Trig/Statistics, Compute Sample Space for 3 standard dice (w/ values 1-6) w/ probability.
'Possible values are 3-18 (1*3 to 6*3)

Private Sub Command1_Click()
Dim d1 As Byte, d2 As Byte, d3 As Byte, dVal As Byte

dVal = Val(Text1.Text)
List1.Clear

For d1 = 1 To 6
    For d2 = 1 To 6
        For d3 = 1 To 6
            If d1 + d2 + d3 = dVal Then
                List1.AddItem "(" & d1 & ", " & d2 & ", " & d3 & ")"
            End If
        Next d3
    Next d2
Next d1

Label2.Caption = "Sample Space Count: " & List1.ListCount

If List1.ListCount > 0 Then
    Label3.Caption = "Probability: " & Round(((List1.ListCount / 216) * 100), 3) & "%"
Else
    Label3.Caption = "Probability: 0"
End If

MsgBox "Finished" & vbCrLf & "Count: " & List1.ListCount, , "Yay!"
End Sub

Private Sub Command2_Click()
List1.Clear
Label2.Caption = "Sample Space Count: " & List1.ListCount
Label3.Caption = "Probability: ?"
End Sub



