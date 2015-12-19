Attribute VB_Name = "modCookies"
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameW" (ByVal lpBuffer As Long, ByRef nSize As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

Private Const CSIDL_COOKIES As Long = &H21

Public Sub ClearIECookies()
  On Error Resume Next
  Dim sPath As String
  
  sPath = Space$(260)
  Call SHGetSpecialFolderPath(0, sPath, CSIDL_COOKIES, False)
  sPath = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\*.txt*"
  Kill sPath
End Sub
