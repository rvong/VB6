VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Pixel Color"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   3960
      Top             =   0
   End
   Begin VB.Frame Frame4 
      Caption         =   "RGB"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Hex"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1935
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Long"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   1785
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim PosXY As POINTAPI

Private Function LongToRGB(lColor As Long) As String
    Dim iRed As Long, iGreen As Long, iBlue As Long
    
    iRed = lColor Mod 256
    iGreen = ((lColor And &HFF00) / 256&) Mod 256&
    iBlue = (lColor And &HFF0000) / 65536
    
    LongToRGB = Format$(iRed, "000") & ", " & Format$(iGreen, "000") & ", " & Format$(iBlue, "000")
End Function

Private Function GetPixelColor(x As Long, y As Long, Optional HexCode As Boolean = False) As String
    Dim DC As Long
    
    DC = CreateDC("DISPLAY", vbNullString, vbNullString, 0&)
    
    GetPixelColor = Format$(GetPixel(DC, x, y), "00000000") 'long
    
    If HexCode = True Then GetPixelColor = LongToHexColor(GetPixelColor)
    
    DeleteDC DC
End Function

Private Function LongToHexColor(ByVal lngColor As Long) As String
Dim hColor As String

    hColor = Right$("000000" & Hex(lngColor), 6)
    LongToHexColor = Mid$(hColor, 5, 2) & Mid$(hColor, 3, 2) & Mid$(hColor, 1, 2)
End Function

Private Function GetMousePosX() As Long
    GetCursorPos PosXY
    GetMousePosX = PosXY.x
End Function

Private Function GetMousePosY() As Long
    GetCursorPos PosXY
    GetMousePosY = PosXY.y
End Function

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Picture1.BackColor = GetPixelColor(GetMousePosX, GetMousePosY)
    
    Label1.Caption = GetPixelColor(GetMousePosX, GetMousePosY)
    Label2.Caption = GetPixelColor(GetMousePosX, GetMousePosY, True)
    Label3.Caption = LongToRGB(GetPixelColor(GetMousePosX, GetMousePosY))
End Sub
