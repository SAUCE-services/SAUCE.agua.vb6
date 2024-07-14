Attribute VB_Name = "modDB"
Option Explicit

Public Function validateNullField(pField As Variant) As String

    If IsNull(pField) Then
        validateNullField = "Null"
    Else
        validateNullField = "'" & pField & "'"
    End If
    
End Function

Public Function replaceSpecialSymbols(ByVal pValue As String) As String
    
    pValue = Replace(pValue, "\", "\\")
    pValue = Replace(pValue, "'", "''")
    pValue = Replace(pValue, Chr(34), Chr(34) & Chr(34))
    
    replaceSpecialSymbols = pValue

End Function

Public Function fechaDB(pField As Variant) As String
    
    If IsNull(pField) Then
        fechaDB = "Null"
    Else
        fechaDB = "'" & Format(pField, "yyyy-mm-dd") & "'"
    End If

End Function

Public Function horaDB(pField As Variant) As String
    
    If IsNull(pField) Then
        horaDB = "Null"
    Else
        horaDB = "'" & Format(pField, "HH:mm:ss") & "'"
    End If

End Function

Public Function fechaHoraDB(pField As Variant) As String

    If IsNull(pField) Then
        fechaHoraDB = "Null"
    Else
        fechaHoraDB = "'" & Format(pField, "yyyy/mm/dd HH:mm:ss") & "'"
    End If

End Function

Public Function toReportDate(pFecha As Date) As String

    toReportDate = "DATE(" & Year(pFecha) & ", " & Month(pFecha) & ", " & Day(pFecha) & ")"
    
End Function



