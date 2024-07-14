Attribute VB_Name = "modRegistry"
Option Explicit

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_SZ = 1 'Unicode null terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const REG_OPTION_BACKUP_RESTORE = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_OPTION_VOLATILE = 1
Public Const REG_CREATED_NEW_KEY = &H1        ' A new key was created
Public Const REG_OPENED_EXISTING_KEY = &H2    ' An existing key was opened

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&         ' Registry path does not exist
Public Const ERROR_ACCESS_DENIED = 5&          ' Requested permissions not available
Public Const ERROR_INVALID_HANDLE = 6&         ' Invalid handle or top-level key
Public Const ERROR_BAD_NETPATH = 53            ' Network path not found
Public Const ERROR_INVALID_PARAMETER = 87      ' Bad parameter to a Win32 API function
Public Const ERROR_CALL_NOT_IMPLEMENTED = 120& ' Function valid only in WinNT/2000?XP
Public Const ERROR_INSUFFICIENT_BUFFER = 122   ' Buffer too small to hold data
Public Const ERROR_BAD_PATHNAME = 161          ' Registry path does not exist
Public Const ERROR_NO_MORE_ITEMS = 259&        ' Invalid enumerated value
Public Const ERROR_BADDB = 1009                ' Corrupted registry
Public Const ERROR_BADKEY = 1010               ' Invalid registry key
Public Const ERROR_CANTOPEN = 1011&            ' Cannot open registry key
Public Const ERROR_CANTREAD = 1012&            ' Cannot read from registry key
Public Const ERROR_CANTWRITE = 1013&           ' Cannot write to registry key
Public Const ERROR_REGISTRY_RECOVERED = 1014&  ' Recovery of part of registry successful
Public Const ERROR_REGISTRY_CORRUPT = 1015&    ' Corrupted registry
Public Const ERROR_REGISTRY_IO_FAILED = 1016&  ' Input/output operation failed
Public Const ERROR_NOT_REGISTRY_FILE = 1017&   ' Input file not in registry file format
Public Const ERROR_KEY_DELETED = 1018&         ' Key already deleted
Public Const ERROR_KEY_HAS_CHILDREN = 1020&    ' Key has subkeys & cannot be deleted

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = ((KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK))

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal pszSubKey As String) As Long
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
Dim lRegResult As Long

    lRegResult = RegDeleteKey(hKey, strPath)
End Sub
Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
    
    'Set up default value
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
        End If
    Else
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

    'Set up default value
    If Not IsEmpty(Default) Then
        GetSettingLong = Default
    Else
        GetSettingLong = 0
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4 '4 bytes = 32 bits = long

    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetSettingLong = lBuffer
        End If
    Else
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)

    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long

    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetSettingByte = Default
        Else
            GetSettingByte = 0
        End If
    Else
        GetSettingByte = 0
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)

    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            ReDim byBuffer(lDataBufferSize - 1) As Byte
            lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
            GetSettingByte = byBuffer
        End If
    Else
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
Dim lRegResult As Long
Dim hCurKey As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)

    lRegResult = RegCloseKey(hCurKey)
End Sub
