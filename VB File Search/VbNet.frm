VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   9975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Text            =   "0"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search For Files"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   555
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Text            =   "Found"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8400
      TabIndex        =   6
      Text            =   "*.*"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recurse"
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label4 
      Caption         =   "File Type"
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Elapsed:"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Start Path"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Command1_Click()

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   
  'setting the list visibility to false
  'increases the load time
   Text3.Text = ""
   List1.Clear
   List1.Visible = False
   
  'set up search params
   With FP
      .sFileRoot = Text1.Text       'start path
      .sFileNameExt = Text2.Text    'file type of interest
      .bRecurse = Check1.Value = 1  '1 = recursive search
   End With
   
  'get start time, get files, and get finish time
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()
   
   List1.Visible = True
   
  'show the results
   Text3.Text = Format$(FP.Count, "###,###,###,##0") & _
                        " found (" & _
                        FP.sFileNameExt & ")"
                   
   Text4.Text = FormatNumber((tend - tstart) / 1000, 2) & "  seconds"
                                    
End Sub


Private Sub GetFileInformation(FP As FILE_PARAMS)

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  'FP.sFileRoot contains the path to search.
  'FP.sFileNameExt contains the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
         
        'Even though this routine uses file specs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then

           'this is where you add code to store
           'or display the returned file listing.
           '
           'if you want the file name only, save 'sTmp'.
           'if you want the full path, save 'sRoot & sTmp'

           'remove trailing nulls
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            List1.AddItem sRoot & sTmp

         End If
         
      Loop While FindNextFile(hFile, WFD)
      
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If

End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'This is where the method obtains the file
     'list and data for the folder passed.
      Call GetFileInformation(FP)

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           '..and the Recurse flag was specified
            If FP.bRecurse Then
            
              'and if the folder is not the default
              'self and parent folders (a . or ..)
               If Asc(WFD.cFileName) <> vbDot Then
               
                 '..then the item is a real folder, which
                 'may contain other sub folders, so assign
                 'the new folder name to FP.sFileRoot and
                 'recursively call this function again with
                 'the amended information.

                 'remove trailing nulls
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
                  
               End If
               
            End If
            
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub


Private Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
      
End Function


Private Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

Private Sub Form_Load()

End Sub
