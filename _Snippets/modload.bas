Attribute VB_Name = "modload"
Option Explicit

Public Function ReadBinFile(ByVal fileName As String) As Byte()
   Dim fnum As Integer
   Dim data() As Byte

   fnum = FreeFile(0)
   Open fileName For Binary Access Read As #fnum
   ReDim data(LOF(fnum) - 1)
   Get #fnum, , data
   Close #fnum
   ReadBinFile = data
End Function


Public Function GetFileText(ByVal fileName As String) As String
  Dim bytText() As Byte
  Dim intFile As Integer

  intFile = FreeFile

  Open fileName For Binary Access Read As intFile
    ReDim bytText(LOF(intFile) - 1)
    Get #intFile, 1, bytText
  Close intFile
  
  GetFileText = bytText
End Function

Public Sub FileTextBIN(ByVal fileName As String, strBuffer As String)
    Dim fileNum As Byte
    ' ensure that the file exists
    If Len(Dir$(fileName)) = 0 Then Err.Raise 53  ' File not found
    
    fileNum = FreeFile
    
    Open fileName For Binary As fileNum
        strBuffer = Space$(LOF(fileNum))
        Get fileNum, , strBuffer
    Close fileNum
End Sub
