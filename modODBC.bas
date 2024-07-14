Attribute VB_Name = "modODBC"
Option Explicit

' ODBC API
' -- ODBC Commands
Public Const ODBC_ADD_DSN = 1&
Public Const ODBC_CONFIG_DSN = 2&
Public Const ODBC_REMOVE_DSN = 3&
Public Const ODBC_ADD_SYS_DSN = 4&
Public Const ODBC_CONFIG_SYS_DSN = 5&
Public Const ODBC_REMOVE_SYS_DSN = 6&
Public Const ODBC_REMOVE_DEFAULT_DSN = 7&
' -- ODBC Error Codes
Public Const ODBC_ERROR_GENERAL_ERR = 1
Public Const ODBC_ERROR_INVALID_BUFF_LEN = 2
Public Const ODBC_ERROR_INVALID_HWND = 3
Public Const ODBC_ERROR_INVALID_STR = 4
Public Const ODBC_ERROR_INVALID_REQUEST_TYPE = 5
Public Const ODBC_ERROR_COMPONENT_NOT_FOUND = 6
Public Const ODBC_ERROR_INVALID_NAME = 7
Public Const ODBC_ERROR_INVALID_KEYWORD_VALUE = 8
Public Const ODBC_ERROR_INVALID_DSN = 9
Public Const ODBC_ERROR_INVALID_INF = 10
Public Const ODBC_ERROR_REQUEST_FAILED = 11
Public Const ODBC_ERROR_INVALID_PATH = 12
Public Const ODBC_ERROR_LOAD_LIB_FAILED = 13
Public Const ODBC_ERROR_INVALID_PARAM_SEQUENCE = 14
Public Const ODBC_ERROR_INVALID_LOG_FILE = 15
Public Const ODBC_ERROR_USER_CANCELED = 16
Public Const ODBC_ERROR_USAGE_UPDATE_FAILED = 17
Public Const ODBC_ERROR_CREATE_DSN_FAILED = 18
Public Const ODBC_ERROR_WRITING_SYSINFO_FAILED = 19
Public Const ODBC_ERROR_REMOVE_DSN_FAILED = 20
Public Const ODBC_ERROR_OUT_OF_MEM = 21
Public Const ODBC_ERROR_OUTPUT_STRING_TRUNCATED = 22

Public Const vbAPINull As Long = 0&
'API to modify/Edit/Create a Data Source Name
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwnd As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Public Sub registerODBC5(pDSN As String, pDriver As String, pDatabase As String, pIP As String)
Dim DriverPath As String

    DriverPath = modRegistry.GetSettingString(modRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBCINST.INI\" & pDriver, "Driver", "")
    modRegistry.SaveSettingString modRegistry.HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", pDSN, pDriver
    modRegistry.SaveSettingString modRegistry.HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & pDSN, "Driver", DriverPath
    modRegistry.SaveSettingString modRegistry.HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & pDSN, "Server", pIP
    modRegistry.SaveSettingString modRegistry.HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & pDSN, "Database", pDatabase
    modRegistry.SaveSettingString modRegistry.HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & pDSN, "Port", "3306"
    
End Sub


