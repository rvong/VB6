Attribute VB_Name = "modGeneral"
Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Function FileExists(Path As String) As Boolean
    FileExists = CBool(PathFileExists(Path))
End Function

Public Sub BinOpen(Path As String, Buffer As String)
    Dim FF As Integer: FF = FreeFile
    
    Open Path For Binary Access Read As FF
        Buffer = Space$(LOF(FF))
        Get FF, , Buffer
    Close FF
End Sub

'Faster Split - @Merri
Public Sub QuickSplit(Expression As String, ResultSplit() As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByRef IgnoreDelimiterWithin As String = vbNullString)
    Dim lngA As Long, lngB As Long, lngCount As Long, lngDelLen As Long, lngExpLen As Long, lngIgnLen As Long, lngResults() As Long
    lngExpLen = LenB(Expression)
    lngDelLen = LenB(Delimiter)
    If lngExpLen > 0 And lngDelLen > 0 And (Limit > 0 Or Limit = -1&) Then
        lngIgnLen = LenB(IgnoreDelimiterWithin)
        If lngIgnLen Then
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            lngB = InStrB(1, Expression, IgnoreDelimiterWithin, Compare)
            Do Until (lngB And 1) Or (lngB = 0)
                lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
            Loop
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        Do Until (lngA And 1) Or (lngA = 0)
                            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                        Loop
                        lngCount = lngCount + 1
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                        Loop
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            Do Until (lngA And 1) Or (lngA = 0)
                                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                            Loop
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                                Do Until (lngB And 1) Or (lngB = 0)
                                    lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                                Loop
                            End If
                        End If
                    End If
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        Do Until (lngA And 1) Or (lngA = 0)
                            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                        Loop
                        lngCount = lngCount + 1
                        If lngCount = Limit Then Exit Do
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                        Loop
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            Do Until (lngA And 1) Or (lngA = 0)
                                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                            Loop
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                                Do Until (lngB And 1) Or (lngB = 0)
                                    lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                                Loop
                            End If
                        End If
                    End If
                Loop
            End If
        Else
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                    Loop
                    lngCount = lngCount + 1
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0 And lngCount < Limit
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                    Loop
                    lngCount = lngCount + 1
                Loop
            End If
        End If
        ReDim Preserve ResultSplit(0 To lngCount)
        If lngCount = 0 Then
            ResultSplit(0) = Expression
        Else
            ResultSplit(0) = LeftB$(Expression, lngResults(0) - 1)
            For lngCount = 0 To lngCount - 2
                ResultSplit(lngCount + 1) = MidB$(Expression, lngResults(lngCount) + lngDelLen, lngResults(lngCount + 1) - lngResults(lngCount) - lngDelLen)
            Next lngCount
            ResultSplit(lngCount + 1) = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
        End If
    Else
        ResultSplit = VBA.Split(vbNullString)
    End If
End Sub
