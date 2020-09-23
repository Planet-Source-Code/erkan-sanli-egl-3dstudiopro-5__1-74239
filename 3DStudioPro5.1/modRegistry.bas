Attribute VB_Name = "modRegistry"
Option Explicit
 
'<<<<<<<<<<<<<<<
    Public Const REFFILES = "Software\EGL\A3DStudioV5"
    
    Public Type SETUPPARAMETERS
        First               As Boolean
        AutoRefresh         As Byte
'        FacePath(6)         As String   '0-5 face pictures , 6 background picture
'        BackGroundOption    As Byte
'        EffectOption        As String
'        ScreenResolution    As String
'        Interval            As Integer  'Timer interval for speed
'        CubeSize            As Integer
'        Mask                As Byte     'chkMask ; apply Mask
'        MaskColor           As Long     'Mask color
'        Opacity             As Byte
'        CubeType            As Byte     'Type 0:Picture, 1:Clock
'        ClockFaceIndex      As Byte     'Clock Type
'        MouseMove           As Byte
    End Type
    
    Public Params As SETUPPARAMETERS

'<<<<<<<<<<<<<<<<


'    Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
'    Const HKEY_LOCAL_MACHINE = &H80000002
'    Const HKEY_USERS = &H80000003
'    Const HKEY_CURRENT_CONFIG = &H80000005
'    Const HKEY_DYN_DATA = &H80000006
    
    Const STANDARD_RIGHTS_ALL = &H1F0000
    Const KEY_QUERY_VALUE = &H1
    Const KEY_SET_VALUE = &H2
    Const KEY_CREATE_SUB_KEY = &H4
    Const KEY_ENUMERATE_SUB_KEYS = &H8
    Const KEY_NOTIFY = &H10
    Const KEY_CREATE_LINK = &H20
    Const SYNCHRONIZE = &H100000
         'KEY_ALL_ACCESS = &H3F
    Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
                                    KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
                                    KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
                                    KEY_CREATE_LINK) And (Not SYNCHRONIZE))
                                    
    'Const REG_CREATED_NEW_KEY = &H1
    'Const REG_OPENED_EXISTING_KEY = &H2
                                    
    'Const REG_NONE = 0
    Public Const REG_SZ = 1
    'Const REG_EXPAND_SZ = 2
    'Const REG_BINARY = 3
    Public Const REG_DWORD = 4
    'Const REG_DWORD_LITTLE_ENDIAN = 4
    'Const REG_DWORD_BIG_ENDIAN = 5
    'Const REG_LINK = 6
    'Const REG_MULTI_SZ = 7
    'Const REG_RESOURCE_LIST = 8
    
    Const REG_OPTION_NON_VOLATILE = 0
    
    Const ERROR_SUCCESS = 0&
    Const ERROR_NONE = 0
    Const ERROR_BADDB = 1
    Const ERROR_BADKEY = 2
    Const ERROR_CANTOPEN = 3
    Const ERROR_CANTREAD = 4
    Const ERROR_CANTWRITE = 5
    Const ERROR_OUTOFMEMORY = 6
    Const ERROR_INVALID_PARAMETER = 7
    Const ERROR_ACCESS_DENIED = 8
    Const ERROR_INVALID_PARAMETERS = 87
    Const ERROR_MORE_DATA = 234
    Const ERROR_NO_MORE_ITEMS = 259&

    Private Enum ETYPE
       cREG_Unknown = 0
       cREG_String = 1
       cREG_EnvString = 2
       cREG_Integer = 3
       cREG_BigEndian = 4
       cREG_Binary = 5
    End Enum
    
    Public Type SECURITY_ATTRIBUTES
       nLength As Long
       lpSecurityDescriptor As Long
       bInheritHandle As Long
    End Type
    
    Public Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
    End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
        (ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        lpcbName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpClass As String, _
        lpcbClass As Long, _
        lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
        (ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpType As Long, _
        ByVal lpData As Byte, _
        ByVal lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
        (ByVal hKey As Long, _
        ByVal lpClass As String, _
        lpcbClass As Long, _
        ByVal lpReserved As Long, _
        lpcSubKeys As Long, _
        lpcbMaxSubKeyLen As Long, _
        lpcbMaxClassLen As Long, _
        lpcValues As Long, _
        lpcbMaxValueNameLen As Long, _
        lpcbMaxValueLen As Long, _
        lpcbSecurityDescriptor As Long, _
        lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Any, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        ByVal lpData As String, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Long, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        ByVal lpData As Long, _
        lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        lpData As Any, _
        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        lpData As Byte, _
        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        lpData As Long, _
        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal lpData As String, _
        ByVal cbData As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" _
        (ByVal hKey As Long, _
        ByVal lpFile As String, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES) _
        As Long
Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal lpFile As String) As Long
        
Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
    
    Dim lRetval As Long
    Dim hKey As Long
    
    lRetval = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegDeleteKey(lPredefinedKey, sKeyName)
    RegCloseKey (hKey)

End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    
    Dim lRetval As Long
    Dim hKey As Long
    
    lRetval = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegDeleteValue(hKey, sValueName)
    RegCloseKey (hKey)

End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select

End Function

Public Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    
    Dim cch         As Long
    Dim lrc         As Long
    Dim lType       As Long
    Dim lValue      As Long
    Dim sValue      As String

    On Error GoTo QueryValueExError
    
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
    Select Case lType
        Case REG_SZ             ' For Strings
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD          ' For DWORDS
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    vValue = -1
    Resume QueryValueExExit
    
End Function

Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    
    Dim hNewKey As Long
    Dim lRetval As Long
    Dim sa As SECURITY_ATTRIBUTES
    
    lRetval = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, sa, hNewKey, lRetval)
    RegCloseKey (hNewKey)

End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    
    Dim lRetval As Long
    Dim hKey As Long
    
    lRetval = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetval = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    
    Dim lRetval  As Long
    Dim hKey     As Long
    Dim vValue   As Variant
    
    lRetval = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetval = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
    
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><

Public Sub RegWrite_AutoRefresh(NewValue As Long)
    
    SetKeyValue HKEY_CURRENT_USER, REFFILES, "AutoRefresh", NewValue, REG_DWORD

End Sub

Public Function RegRead_AutoRefresh() As Long
    
    Dim retval As Long
    
    retval = QueryValue(HKEY_CURRENT_USER, REFFILES, "AutoRefresh")
    If retval = -1 Then
       CreateNewKey HKEY_CURRENT_USER, REFFILES
       SetKeyValue HKEY_CURRENT_USER, REFFILES, "AutoRefresh", "1", REG_DWORD
       retval = QueryValue(HKEY_CURRENT_USER, REFFILES, "AutoRefresh")
    End If
    RegRead_AutoRefresh = retval

End Function
