VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Registry Events
Public Event RegistryError(ByVal Number As Long, ByVal Description As String)

'Registry API declarations
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'Kernel32 declaration
Private Declare Function FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'Operating System Platform ID constants
Private Const VER_PLATFORM_WIN32_NT = 2

'FormatMessage constants
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

'Registry Root Key constants
Public Enum RegRootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

'Registry Data Type constants
Public Enum RegDataType
    REG_SZ = 1
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_QWORD = 11
    REG_QWORD_LITTLE_ENDIAN = 11
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
    REG_NONE = 0
End Enum

'Registry base error code
Private Const ERROR_BASE = 50000

'Windows Error Code constants
Public Enum RegErrorCode
    ERROR_SUCCESS = 0&
    ERROR_MORE_DATA = ERROR_BASE + 234
    ERROR_INVALID_FUNCTION = ERROR_BASE + 1
    ERROR_FILE_NOT_FOUND = ERROR_BASE + 2&
    ERROR_ACCESS_DENIED = ERROR_BASE + 5&
    ERROR_INVALID_HANDLE = ERROR_BASE + 6&
    ERROR_INVALID_ACCESS = ERROR_BASE + 12&
    ERROR_OUTOFMEMORY = ERROR_BASE + 14&
    ERROR_NETNAME_DELETED = ERROR_BASE + 64&
    ERROR_NETWORK_ACCESS_DENIED = ERROR_BASE + 65&
    ERROR_TOO_MANY_SESS = ERROR_BASE + 69&
    ERROR_INVALID_PARAMETER = ERROR_BASE + 87
    ERROR_BAD_PATHNAME = ERROR_BASE + 161&
End Enum

'Property variables
Private pSubKey      As String
Private pRootKey     As RegRootKey
Private pDataType    As RegDataType
Private pValueName   As String
Private pForceWrite  As Boolean
Private pReportError As Boolean

Private Function IsWindowsNT() As Boolean
    Dim lpOSVerInfo As OSVERSIONINFO
    
    'Initialize the structure
    lpOSVerInfo.dwOSVersionInfoSize = Len(lpOSVerInfo)
    
    'Check Windows Platform ID
    GetVersionEx lpOSVerInfo
    
    If lpOSVerInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        'This is Windows NT
        IsWindowsNT = True
    End If
    
End Function

Private Sub RaiseError(ByVal Number As Long)
    Dim Desc    As String * 255
    Dim nChar   As Long
    
    If Err.Number <> 0 Then
        'Replace with local error if no error occurred in API
        Number = Err.Number
        Desc = Err.Description
    Else
        'Format the error code into description
        nChar = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Number, 0, Desc, 255, ByVal 0&)
        
        Desc = Left(Desc, nChar)
    End If
    
    If ReportError Then
        'Raise error
        Err.Raise Number, "Registry", Desc
    End If
    
    RaiseEvent RegistryError(Number, Desc)
    
End Sub

Public Property Get ReportError() As Boolean
    ReportError = pReportError
End Property

Public Property Let ReportError(ByVal Value As Boolean)
    pReportError = Value
End Property

Public Property Get ForceWrite() As Boolean
    ForceWrite = pForceWrite
End Property

Public Property Let ForceWrite(ByVal Value As Boolean)
    pForceWrite = Value
End Property

Public Property Get RootKey() As RegRootKey
    RootKey = pRootKey
End Property

Public Property Let RootKey(ByVal Value As RegRootKey)
    pRootKey = Value
End Property

Public Property Get SubKey() As String
    SubKey = pSubKey
End Property

Public Property Let SubKey(ByVal Value As String)
    pSubKey = Value
End Property

Public Property Get ValueName() As String
    ValueName = pValueName
End Property

Public Property Let ValueName(ByVal Value As String)
    pValueName = Value
End Property

Public Property Get DataType() As RegDataType
    DataType = pDataType
End Property

Public Property Let DataType(ByVal Value As RegDataType)
    pDataType = Value
End Property

Public Property Get Data() As Variant
    Dim lpBufferCurrency    As Currency
    Dim lpBufferBinary()    As Byte
    Dim lpBufferString      As String
    Dim lpBufferLong        As Long
    Dim nSize               As Long
    Dim hKey                As Long
    Dim lRet                As Long
    
    On Local Error GoTo Finally
    
    'Open a handle to the required registry key
    lRet = RegOpenKey(pRootKey, pSubKey, hKey)
    
    If lRet = ERROR_SUCCESS Then
        'Retrieve the data type and data size
        RegQueryValueEx hKey, pValueName, 0&, pDataType, ByVal 0&, nSize
        
        'Convert the buffer to the data type we wanted
        Select Case pDataType
            Case REG_BINARY 'Binary data
                'Initialize buffer to retrieve data
                ReDim lpBufferBinary(1 To nSize)
                
                'Retrieve the binary data
                lRet = RegQueryValueEx(hKey, pValueName, 0&, pDataType, lpBufferBinary(1), nSize)
                
                If lRet = ERROR_SUCCESS Then
                    'Return the value
                    Data = lpBufferBinary
                End If
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ 'String
                'Initialize buffer to retrieve string data
                lpBufferString = Space(nSize)
                
                'Retrieve the string data
                lRet = RegQueryValueEx(hKey, pValueName, 0&, pDataType, ByVal lpBufferString, nSize)
                
                If lRet = ERROR_SUCCESS Then
                    'Return the data
                    Data = Left(lpBufferString, nSize - 1)
                End If
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN 'Long/Int32
                'Retrieve the DWord value
                lRet = RegQueryValueEx(hKey, pValueName, 0&, pDataType, lpBufferLong, nSize)
                
                If lRet = ERROR_SUCCESS Then
                    'Return the value
                    Data = lpBufferLong
                End If
            Case REG_QWORD 'Currency/Int64
                'Retrieve the QWord value
                lRet = RegQueryValueEx(hKey, pValueName, 0&, pDataType, lpBufferCurrency, nSize)
                
                If lRet = ERROR_SUCCESS Then
                    'Return the value
                    Data = lpBufferCurrency
                End If
        End Select
    End If
    
Finally:
    'Reset error information
    RaiseError lRet
    
    'Release handle
    RegCloseKey hKey
    
End Property

Public Property Let Data(ByVal Value As Variant)
    Dim hKey As Long
    Dim lRet As Long
    
    'Open a handle to the required key
    lRet = RegCreateKey(pRootKey, pSubKey, hKey)
    If lRet <> ERROR_SUCCESS Then RaiseError lRet
    
    'Automatically determine the data type
    Select Case VarType(Value)
        Case vbByte, vbInteger, vbLong, vbBoolean
            'DWord value
            pDataType = REG_DWORD
            lRet = RegSetValueEx(hKey, pValueName, 0&, pDataType, ByVal VarPtr(CLng(Value)), 4)
        Case vbString
            'LPStr value
            pDataType = REG_SZ
            
            If IsWindowsNT() Then
                lRet = RegSetValueEx(hKey, pValueName, 0&, pDataType, ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
            Else
                lRet = RegSetValueEx(hKey, pValueName, 0&, pDataType, ByVal StrPtr(StrConv(Value, vbFromUnicode)), (Len(Value) \ 2))
            End If
            
        Case vbCurrency
            'QWord value
            pDataType = REG_QWORD
            lRet = RegSetValueEx(hKey, pValueName, 0&, pDataType, ByVal VarPtr(CCur(Value)), 8)
        Case Else
            'Unreconised data type
            GoTo Finally
    End Select
        
    'Set the value of the key
    If lRet <> ERROR_SUCCESS Then RaiseError lRet
    
    If pForceWrite Then
        'Force the operation system to write the changes
        RegFlushKey hKey
    End If
    
Finally:
    'Close the handle
    RegCloseKey hKey
    
End Property

Public Function GetData(RootKey As RegRootKey, SubKey As String, ValueName As String) As Variant
    'Set optional property
    If Not IsMissing(SubKey) Then Me.SubKey = SubKey
    If Not IsMissing(RootKey) Then Me.RootKey = RootKey
    If Not IsMissing(ValueName) Then Me.ValueName = ValueName
    
    'Retrieve data
    GetData = Me.Data
End Function

Public Sub SetData(RootKey As RegRootKey, SubKey As String, ValueName As String, ByVal Data As Variant)
    'Set optional property
    If Not IsMissing(SubKey) Then Me.SubKey = SubKey
    If Not IsMissing(RootKey) Then Me.RootKey = RootKey
    If Not IsMissing(ValueName) Then Me.ValueName = ValueName
    
    'Set data
    Me.Data = Data
End Sub
