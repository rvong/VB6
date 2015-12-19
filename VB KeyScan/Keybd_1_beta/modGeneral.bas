Attribute VB_Name = "modGeneral"
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'this file needs to be on the comp

Public Sub OpenURL(sURL As String)
    ShellExecute 0&, "open", sURL, 0&, 0&, 1&
End Sub

Public Function GetKeyStatus(KeyCode As Integer, Optional bInit As Boolean = False) As Boolean
  'Stupid stupid stupid, Alternates On/off >_<
  ' -127 = repeated on   -128 = repeated off

    If GetKeyState(KeyCode) = 0 Or GetKeyState(KeyCode) = -127 Then 'ON
        If bInit = False Then 'Switch for alternative Init BS
            GetKeyStatus = True
        Else
            GetKeyStatus = False
        End If
        
    ElseIf GetKeyState(KeyCode) = 1 Or GetKeyState(KeyCode) = -128 Then 'OFF
        If bInit = False Then
            GetKeyStatus = False
        Else
            GetKeyStatus = True
        End If
        
    End If
End Function

Public Sub UpdateLED(LEDKey As Integer, Optional bInit As Boolean = False)
  Select Case LEDKey
    Case vbKeyNumlock
      If GetKeyStatus(vbKeyNumlock, bInit) = True Then
        frmMain.shpLED(0).BackColor = &H55FECA
      Else
        frmMain.shpLED(0).BackColor = &H8000000A
      End If
    Case vbKeyCapital
      If GetKeyStatus(vbKeyCapital, bInit) = True Then
        frmMain.shpLED(1).BackColor = &H55FECA
      Else
        frmMain.shpLED(1).BackColor = &H8000000A
      End If
    Case vbKeyScrollLock
      If GetKeyStatus(vbKeyScrollLock, bInit) = True Then
        frmMain.shpLED(2).BackColor = &H55FECA
      Else
        frmMain.shpLED(2).BackColor = &H8000000A
      End If
  End Select
End Sub

Public Function HasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
  HasIndex = (VarType(ControlArray(Index)) <> vbObject)
End Function

Public Function FormatProcessorName(Expr As String) As String
  Expr = RidOfSpaces(Expr)
  Expr = Replace$(Expr, "(TM)", "™")
  Expr = Replace$(Expr, "(R)", "®")
  
  FormatProcessorName = Expr
End Function

Public Function RidOfSpaces(Expr As String) As String
  Do Until InStrB(Expr, "  ") = 0
    Expr = Replace$(Expr, "  ", " ")
  Loop
  
  RidOfSpaces = Trim$(Expr)
End Function

Public Function LangIdent(keyLayout As String) As String
  Dim KeybdLayoutList As String
  
  KeybdLayoutList = StrConv(LoadResData("LAYOUT", "KEYBD"), vbUnicode)
  LangIdent = ParseIt(KeybdLayoutList, Trim$(keyLayout) & ":", vbNewLine)
End Function

Public Function ParseIt(Expr As String, Del_A As String, Del_B As String) As String
  Dim A As Long, B As Long
  
  A = InStrB(1, Expr, Del_A) + LenB(Del_A)
  
  If A - LenB(Del_A) > 0 Then
    B = InStrB(A, Expr, Del_B)
    ParseIt = MidB$(Expr, A, B - A)
  End If
End Function

Public Sub AddKeysPressed(akpName As String, akpDSC As String, akpVK As String, akpSC As Integer, akpFlags As Integer)
    Exit Sub
    Dim ItemToAdd As ListItem
        
    Set ItemToAdd = frmMain.lvwKeysPressed.ListItems.Add(, , akpName)
    ItemToAdd.SubItems(1) = akpDSC
    ItemToAdd.SubItems(2) = akpVK
    ItemToAdd.SubItems(3) = akpSC
    ItemToAdd.SubItems(4) = akpFlags

    ItemToAdd.Selected = True
    ItemToAdd.EnsureVisible
    ItemToAdd.Selected = False

    Set ItemToAdd = Nothing

''''''''''''
'  With frmMain.lvwKeysPressed.ListItems
'    .Add , , akpName
'    .Item(.Count).SubItems(1) = akpDSC
'    .Item(.Count).SubItems(2) = akpVK
'    .Item(.Count).SubItems(3) = akpSC
'    .Item(.Count).SubItems(4) = akpFlags
'
'    'Scroll Down
'
'    'frmMain.lvwKeysPressed.ListItems(.Count).Selected = True
'
'    'Don't over populate Listview
'    If .Count > 100 Then .Remove 1
'  End With
End Sub
