VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "RemDupes"
   ClientHeight    =   2670
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   3240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   117
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   " Output File"
      ToolTipText     =   "Output Location"
      Top             =   465
      Width           =   2938
   End
   Begin VB.CommandButton cmdRemDupes 
      Caption         =   "Remove Duplicates!"
      Height          =   364
      Left            =   117
      TabIndex        =   4
      ToolTipText     =   "Remove Duplicates"
      Top             =   2220
      Width           =   3640
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   3159
      TabIndex        =   3
      ToolTipText     =   "Select Output File"
      Top             =   435
      Width           =   598
   End
   Begin VB.CheckBox chkAlpha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alphabetize"
      ForeColor       =   &H80000008&
      Height          =   247
      Left            =   2457
      TabIndex        =   2
      ToolTipText     =   "Alphabetize Word List"
      Top             =   930
      Value           =   1  'Checked
      Width           =   1183
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   285
      Left            =   3159
      TabIndex        =   1
      ToolTipText     =   "Select Input File"
      Top             =   75
      Width           =   598
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   117
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " Input File"
      ToolTipText     =   "Input Location"
      Top             =   105
      Width           =   2938
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Status"
      Top             =   1860
      Width           =   3645
   End
   Begin VB.Label lblUnique 
      BackStyle       =   0  'Transparent
      Caption         =   "Unique Items:"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Unique"
      Top             =   1395
      Width           =   2115
   End
   Begin VB.Label lblDuplicates 
      BackStyle       =   0  'Transparent
      Caption         =   "Duplicates:"
      Height          =   240
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Number of Duplicates"
      Top             =   1170
      Width           =   2115
   End
   Begin VB.Label lblLines 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines:"
      Height          =   240
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Line Count"
      Top             =   930
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   945
      Left            =   120
      Top             =   810
      Width           =   3645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Removes duplicates from large (milllions) wordlists quickly
'
'Sorts wordlists using the quicksort algorithm, then
'removes duplicates in a single pass, O(n), by comparing
'each item against the preceding item.
'
'Items can be "unsorted" back to the order of the original
'wordlist with the use of an array of indexes, a new
'array containing the original position of each item.
'
'Stable Quicksort algorithms by Rde.
'http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=63941&lngWId=1
'
'Much(!) faster Split replacement, "Quick Split," by Merri @ VBForums.com.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Note: memory usage issues

Option Explicit
           
Dim InputOK As Boolean, OutputOK As Boolean

Dim ArrBuffer() As String, ArrIndexes() As Long

Dim LineCount As Long, Unique As Long
Dim Min As Long, Max As Long
    
Private Sub cmdOpen_Click()
On Error GoTo Cancel
    With cDialog
        .CancelError = True
        .FileName = vbNullString
        .Filter = "Text Document (*.txt)|*.txt"
        .ShowOpen
        
        If FileExists(.FileName) Then
            txtInput = .FileName
            txtInput.Tag = .FileTitle
            Call LoadWordList
        Else
            InputOK = False
            lblStatus = "Invalid input file"
            MsgBox "Input file does not exist.", vbInformation, "Error"
        End If
    End With
Cancel:
End Sub

Private Sub cmdSave_Click()
On Error GoTo Cancel
    With cDialog
        .CancelError = True
        .FileName = vbNullString
        .Filter = "Text Document (*.txt)|*.txt"
        .ShowSave

        If LenB(.FileName) Then
            OutputOK = True
            txtOutput = .FileName
            lblStatus = "Output file selected!"
        Else
            OutputOK = False
        End If
    End With
Cancel:
End Sub

Private Sub LoadWordList()
    cmdRemDupes.Enabled = False
    lblStatus = "Loading word list...": DoEvents
    
    Dim strBuffer As String
    Call BinOpen(txtInput, strBuffer)   'File Contents -> strBuffer
    Call QuickSplit(strBuffer, ArrBuffer(), vbNewLine) 'strBuffer -> String Array ArrBuffer()
    
    InputOK = True
    
    Min = LBound(ArrBuffer)
    Max = UBound(ArrBuffer)
    LineCount = (Max + 1) - Min
    
    lblLines = "Lines: " & LineCount
    lblDuplicates = "Duplicates:"
    lblUnique = "Unique Items:"
    
    lblStatus = txtInput.Tag & " loaded!"
    cmdRemDupes.Enabled = True
End Sub

Private Sub Reset()
    Erase ArrBuffer()   'Clear Memory
    Erase ArrIndexes()
    InputOK = False     'Reset status
    OutputOK = False
    txtInput = " Input File"
    txtOutput = " Ouptut File"
End Sub

Private Sub ProcessWordList()
    cmdRemDupes.Enabled = False
    
    Dim t As Double: t = Timer

    lblDuplicates = "Duplicates:"
    lblUnique = "Unique Items:"
    lblStatus = "Removing duplicates...": DoEvents
    
    If chkAlpha.Value = vbChecked Then
        Call strStableSort2(ArrBuffer(), Min, Max)
        Call PrintUnique 'Using Indexed Sort -> Call IndexPrintUnique
    Else
        ReDim ArrIndexes(Min To Max)
        Call strStableSort2Indexed(ArrBuffer(), ArrIndexes(), Min, Max)
        Call IndexPrintUniqueOrig
    End If

    lblDuplicates = "Duplicates: " & LineCount - Unique
    lblUnique = "Unique Items: " & Unique
    lblStatus = "Dupes removed in " & Round(Timer - t, 3) & " secs!"
    
    Call Reset
    cmdRemDupes.Enabled = True
End Sub

Private Sub cmdRemDupes_Click()
    If InputOK = False Or OutputOK = False Then
        lblStatus = "Invalid Input/Output File(s)"
        MsgBox "Invalid Input/Output File(s)", vbInformation, "Error"
    Else
        If LineCount > 0 Then
            Call ProcessWordList
        Else
            lblStatus = "Input file is blank"
        End If
    End If
End Sub

Private Sub PrintUnique() 'From string
    Dim FF As Integer: FF = FreeFile
    
    Open txtOutput For Output As FF
        Print #FF, ArrBuffer(Min): Unique = 1 'First Item
        
        Dim i As Long
        For i = Min + 1 To Max          'Start w/ 2nd item, compare to previous item
            If ArrBuffer(i) <> ArrBuffer(i - 1) Then
                Print #FF, ArrBuffer(i) 'Print #FF, ArrBuffer(i); 'Don't add CRLF
                Unique = Unique + 1
            End If
        Next i
    Close FF
End Sub

Private Sub IndexPrintUniqueOrig()
    Dim uIndex() As Long: ReDim uIndex(Min To Max) As Long

    uIndex(ArrIndexes(Min)) = ArrIndexes(Min): Unique = 1 'First Item
    
    Dim i As Long
    For i = Min + 1 To Max
        If ArrBuffer(ArrIndexes(i)) <> ArrBuffer(ArrIndexes(i - 1)) Then
            uIndex(ArrIndexes(i)) = ArrIndexes(i)
            Unique = Unique + 1
        Else
            uIndex(ArrIndexes(i)) = -1
        End If
    Next i
    
    Dim FF As Integer: FF = FreeFile
    Open txtOutput For Output As FF
        For i = Min To Max
            If uIndex(i) > -1 Then Print #FF, ArrBuffer(uIndex(i))
        Next i
    Close FF
End Sub

'Private Sub IndexPrintUnique() 'From Index
'    Dim FF As Integer: FF = FreeFile
'
'    Open txtOutput For Output As FF
'        Print #FF, ArrBuffer(ArrIndexes(Min))       'First item
'        Unique = 1
'
'        Dim i As Long
'        For i = Min + 1 To Max          'Start w/ 2nd item, compare to previous item
'
'            If ArrBuffer(ArrIndexes(i)) <> ArrBuffer(ArrIndexes(i - 1)) Then
'                Unique = Unique + 1
'
'                If i < Max Then
'                    Print #FF, ArrBuffer(ArrIndexes(i))
'                Else
'                    Print #FF, ArrBuffer(ArrIndexes(i)); 'Don't add CRLF
'                End If
'            End If
'
'        Next i
'    Close FF
'End Sub

'Private Sub Col_RemDupes() 'Determine/Remove duplicate items with a Collection
'    Dim col As Collection: Set col = New Collection
'    Dim FF As Integer: FF = FreeFile
'
'    On Error Resume Next
'
'    Open txtOutput For Output As FF
'
'        Dim i As Long
'        For i = LBound(ArrBuffer) To UBound(ArrBuffer)
'            col.Add ArrBuffer(i), ArrBuffer(i)
'
'            If Err.Number <> 457 Then
'                Print #FF, ArrBuffer(i)
'            Else
'                Err.Clear  '457 = Item already exists in the collection
'            End If
'        Next i
'
'    Close FF
'
'    Set col = Nothing
'End Sub
