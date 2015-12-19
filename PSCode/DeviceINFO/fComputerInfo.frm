VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fComputerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hardware information"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "fComputerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   345
      Left            =   8400
      TabIndex        =   4
      Top             =   5760
      Width           =   1125
   End
   Begin VB.PictureBox picWait 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   2880
      ScaleHeight     =   1635
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   4275
      Begin MSComctlLib.ProgressBar prbDetecting 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Please wait while detecting hardware..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         TabIndex        =   10
         Top             =   300
         Width           =   4155
      End
      Begin VB.Shape Shape1 
         Height          =   1635
         Left            =   0
         Top             =   0
         Width           =   4275
      End
   End
   Begin MSComctlLib.TreeView trvComputer 
      Height          =   4995
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   8811
      _Version        =   393217
      Indentation     =   635
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh hardware list"
      Height          =   345
      Left            =   2580
      TabIndex        =   5
      Top             =   180
      Width           =   2475
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Device properties:"
      Height          =   5415
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   4395
      Begin MSComctlLib.ListView lsvProperties 
         Height          =   4035
         Left            =   120
         TabIndex        =   2
         Top             =   1260
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   7117
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtDevice 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4155
      End
      Begin VB.Label Label2 
         Caption         =   "Device:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Device:"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Doctor VB sample application: www.dr-vb.co.il"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   5760
      Width           =   6975
   End
   Begin VB.Label Label4 
      Caption         =   "Hardware list:"
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "fComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************************************************
' Sample code:
' Getting hardware information using WMI (Windows Management Instrumentation)
' Copyright ©2003 Yaniv Drukman, All rights reserved (http://www.dr-vb.co.il).
' This sample code is distributed under the GPL license (http://www.gnu.org).

' If the "Microsft WMI Scripting V1.x Library" is missing, please download it
' from the following link:
' http://msdn.microsoft.com/library/default.asp?url=/downloads/list/wmi.asp
' ******************************************************************************

Private Function GetDevice(DeviceName As String) As Variant
' In this function we will get the devices referring to the given class name
Dim DeviceSet As SWbemObjectSet
Dim Device As SWbemObject
Dim iCount As Integer
Dim sTemp As String

    On Error Resume Next
    ' Set the SWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf(DeviceName)
    
    ' Get the devices captions
    For Each Device In DeviceSet
        sTemp = sTemp & Device.Caption & "|"
    Next Device
    ' Remove the '|' character at the end of the string
    If Right(sTemp, 1) = "|" Then sTemp = Left(sTemp, Len(sTemp) - 1)
    ' Return an array (variant) with the devices captions
    GetDevice = Split(sTemp, "|")
End Function

Private Sub cmdAbout_Click()
    ' Show about box
    MsgBox "Sample code:" & vbCrLf & _
           "Getting hardware information using WMI (Windows Management Instrumentation)" & vbCrLf & _
           "Copyright ©2003 Yaniv Drukman, All rights reserved." & vbCrLf & _
           "This sample code is distributed under the GPL license (http://www.gnu.org).", vbInformation + vbOKOnly, "About"
End Sub

Private Sub cmdExit_Click()
    ' Exit the application
    End
End Sub

Private Sub cmdRefresh_Click()
' In this sub we'll get the hardware list
Dim DevicesNames As Variant
Dim Device As Variant
Dim TempDevice As Variant
Dim nodeComputer As Node
Dim nodeDevice As Node
Dim nodeSubDevice As Node
Dim NumOfDevices As Integer
Dim ProgressInterval As Integer

    ' Disable the TreeView
    trvComputer.Enabled = False
    ' Initialize the ProgressBar
    prbDetecting.Value = 0
    ' Clear the TreeView
    trvComputer.Nodes.Clear
    ' Move the "Wait" PictureBox to the center of the form
    picWait.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    ' Make the "Wait" PictureBox visible
    picWait.Visible = True

    ' Find the computer name
    For Each Device In GetDevice("Win32_ComputerSystem")
        Set nodeComputer = trvComputer.Nodes.Add(, , , Device)
    Next
    ' Make sure that the main node (the computer) is extended.
    trvComputer.Nodes(1).Expanded = True
    ' Show the form
    Me.Show
  
    ' This array contains the Computer System Hardware Classes names
    DevicesNames = Array("Win32_Processor", "Win32_Keyboard", "Win32_PointingDevice", "Win32_CDROMDrive", "Win32_DiskDrive", _
                        "Win32_FloppyDrive", "Win32_TapeDrive", "Win32_1394Controller", "Win32_BaseBoard", "Win32_MotherboardDevice", _
                        "Win32_BIOS", "Win32_Bus", "Win32_CacheMemory", "Win32_FloppyController", "Win32_IDEController", _
                        "Win32_InfraredDevice", "Win32_IRQResource", "Win32_MemoryArray", "Win32_MemoryDevice", "Win32_OnBoardDevice", "Win32_ParallelPort", "Win32_PCMCIAController", _
                        "Win32_PhysicalMemory", "Win32_PhysicalMemoryArray", "Win32_PNPEntity", "Win32_PortConnector", "Win32_PortResource", _
                        "Win32_SCSIController", "Win32_SerialPort", "Win32_SerialPortConfiguration", "Win32_SMBIOSMemory", _
                        "Win32_SoundDevice", "Win32_SystemEnclosure", "Win32_SystemMemoryResource", "Win32_SystemSlot", _
                        "Win32_USBController", "Win32_USBHub", "Win32_NetworkAdapter", _
                        "Win32_AssociatedBattery", "Win32_Battery", "Win32_CurrentProbe", "Win32_PortableBattery", "Win32_UninterruptiblePowerSupply", _
                        "Win32_VoltageProbe", "Win32_Printer", "Win32_POTSModem", _
                        "Win32_POTSModemToSerialPort", "Win32_DesktopMonitor", "Win32_DisplayConfiguration", "Win32_DisplayControllerConfiguration", _
                        "Win32_VideoConfiguration", "Win32_VideoController", "Win32_Fan", "Win32_HeatPipe", "Win32_Refrigeration", "Win32_TemperatureProbe")
    
    ' Find the number of hardware classes
    NumOfDevices = UBound(DevicesNames)
    ' Calculate the ProgressBar's interval
    ProgressInterval = 100 / NumOfDevices
    
    ' Find all the hardware devices
    For Each Device In DevicesNames
        ' Check the ProgressBar state
        If prbDetecting.Value < 100 Then
            On Error Resume Next
            ' Increase the current value of the ProgressBar
            prbDetecting.Value = prbDetecting.Value + ProgressInterval
        End If
        ' Make sure that the operating system can process other events
        DoEvents
        ' Add the hardware class name (without the win32_ prefix) to the TreeView
        Set nodeDevice = trvComputer.Nodes.Add(nodeComputer, tvwChild, , Right(Device, Len(Device) - 6))
            For Each TempDevice In GetDevice(CStr(Device))
                ' Add the device name to the TreeView
                Set nodeSubDevice = trvComputer.Nodes.Add(nodeDevice, tvwChild, , TempDevice)
            Next
    Next
    ' Hide the "Wait" PictureBox
    picWait.Visible = False
    ' Enable the TreeView
    trvComputer.Enabled = True
    
End Sub

Private Sub Form_Load()
    ' Initialize the ListView
    lsvProperties.ColumnHeaders.Add , , "Property"
    lsvProperties.ColumnHeaders.Add , , "Value"
    lsvProperties.View = lvwReport
    lsvProperties.ColumnHeaders(1).Width = lsvProperties.Width / 2 - 40
    lsvProperties.ColumnHeaders(2).Width = lsvProperties.Width / 2 - 40
    ' Get the hardware list
    cmdRefresh_Click
End Sub

Private Sub trvComputer_Click()
Dim vFullPath As Variant
Dim vItems As Variant
Dim vTemp As Variant

    ' Put the path parts of selected item into a variant array
    vFullPath = Split(trvComputer.SelectedItem.FullPath, "\")
    ' Check whether the user choose a device name
    If UBound(vFullPath) = 2 Then
        ' Update the TextBox with the chosen device name
        txtDevice.Text = vFullPath(2)
        On Error Resume Next
        ' Clear the ListView
        lsvProperties.ListItems.Clear
        ' Populate the ListView with the device's properties
        For Each vTemp In GetProperties(vFullPath)
            On Error Resume Next
            vItems = Split(vTemp, "^")
            lsvProperties.ListItems.Add(, , CStr(vItems(0))).SubItems(1) = vItems(1)
        Next vTemp
        ' Resize the columns width (in the ListView)
        Call AutosizeColumns(lsvProperties)
    End If
End Sub

Private Function GetProperties(vPath As Variant) As Variant
' This function returns all the properties of a specific device
Dim DeviceSet As SWbemObjectSet
Dim Device As SWbemObject
Dim iCount As Integer
Dim vTemp As Variant
Dim sTemp As String

    On Error Resume Next
    ' Set theSWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf("Win32_" & vPath(1))
    For Each Device In DeviceSet
        ' Check if the current device in the chosen device
        If Device.Caption = vPath(2) Then
            ' Get all the properties of the chosen device
            For Each vTemp In Device.Properties_
                On Error Resume Next
                If vTemp <> "" And vTemp <> vbNull Then
                    ' Add the property name and its value to the temporary string
                    sTemp = sTemp & vTemp.Name & "^" & vTemp & "|"
                End If
            Next
            ' Remove the '|' character at the end of the string
            If Right(sTemp, 1) = "|" Then
                sTemp = Left(sTemp, Len(sTemp) - 1)
            End If
        End If
    Next Device
    ' Return an array containing the device properties
    GetProperties = Split(sTemp, "|")
End Function
