VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub configureDB()

    loadParametersDB
    
End Sub

Private Sub loadParametersDB()
Dim lngSuccess As Long
    
Dim Reply As ICMP_ECHO_REPLY

    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "DSN") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "DB") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "UID") = "" Then defaultParametersDB
    If GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "IP") = "" Then defaultParametersDB
    
    dbapp.dsn = GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "DSN")
    dbapp.database = GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "DB")
    dbapp.uid = GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "UID")
    dbapp.pwd = GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "PWD")
    dbapp.ip = GetSettingString(HKEY_LOCAL_MACHINE, "Software\agua", "IP")
    
    ' Verifica la IP
    lngSuccess = ping(dbapp.ip, Reply)
    If lngSuccess <> 0 Then
        MsgBox "ERROR: Negocio '" & dbapp.ip & "' " & evaluatePingResponse(lngSuccess)
        End
    End If
    
End Sub

Private Sub defaultParametersDB()
    
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\agua", "DSN", "agua"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\agua", "DB", "agua"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\agua", "UID", "sistemas"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\agua", "PWD", "uvspes"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\agua", "IP", "127.0.0.1"

End Sub

