VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlPagoFacil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function codigopf(liquidacion As clsMODFactura) As String
Dim request As MSXML2.XMLHTTP

Dim url As String
Dim json_text As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturacion/codigopf"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"
    
    json_text = modJsonConverter.ConvertToJson(liquidacion.makeParams)

    request.send json_text
    
    codigopf = request.responseText

End Function

Public Function codigoI2of5(codigo As String) As String
Dim request As MSXML2.XMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturacion/i2of5/" & codigo

    request.Open "GET", url
    request.send
    
    codigoI2of5 = request.responseText

End Function

Public Function loadFile(pFile As String, db As clsDB) As Collection
Dim objPFHF As clsPFHeaderFile
Dim objPFHB As clsPFHeaderBatch
Dim pfrecord As clsPFRecord

Dim colPFR As New Collection

Dim objPFRec As New clsMyAPFRecord

    file2database pFile, db

    For Each objPFRec In objPFRec.collectionByFileName(modConv.parseFilename(pFile), db)
        With objPFRec
            Select Case Left(.line, 1)
                Case "1":
                    Set objPFHF = New clsPFHeaderFile
                    objPFHF.fillByLine .line
                Case "3":
                    Set objPFHB = New clsPFHeaderBatch
                    objPFHB.fillByLine .line
                Case "5":
                    Set pfrecord = New clsPFRecord
                    pfrecord.PFRec5.fillByLine .line
                    colPFR.add pfrecord
                Case "6":
                    pfrecord.PFRec6.fillByLine .line
                Case "7":
                    pfrecord.PFRec7.fillByLine .line
            End Select
        End With
    Next
    
    Set loadFile = colPFR
    
End Function

Private Sub file2database(pFile As String, db As clsDB)
Dim pffile As New clsMyAPFFile
Dim pfrecord As clsMyAPFRecord

Dim strFile As String
Dim strPath As String
Dim strLine As String

    strFile = modConv.parseFilename(pFile, strPath)
    
    pffile.filename = strFile
    pffile.findByPrimaryKey db
    
    If pffile.collectionByFileName(strFile, db).Count > 0 Then Exit Sub
    
    pffile.filename = strFile
    pffile.path = strPath
    pffile.uid = "admin"
    
    pffile.save db
    
    Open pFile For Input As #1
    
    Do
        Line Input #1, strLine
        
        Set pfrecord = New clsMyAPFRecord
        
        pfrecord.filename = strFile
        pfrecord.line = strLine
        pfrecord.uid = "admin"
        
        pfrecord.save db
    Loop While Not EOF(1)
    
    Close #1
    
End Sub
