Attribute VB_Name = "modProperties"
Option Explicit

Global properties As Dictionary

Public Sub loadProperties()
Dim linea As String
Dim suffix As String

Dim keyvalue As Variant

    suffix = Command$
    If suffix = "" Then suffix = "prod"
    
    Set properties = New Dictionary

    Open App.path & "\application-" & suffix & ".properties" For Input As #1
    
    Do
        Line Input #1, linea
        If Left(linea, 1) <> "#" Then
            keyvalue = Split(linea, "=")
            properties.add Trim(keyvalue(0)), Trim(keyvalue(1))
        End If
    Loop Until EOF(1)
    Close #1
    
End Sub

