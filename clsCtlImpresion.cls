VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub printReport(crystal As CrystalReport, report_name As String, string_connection As String, Optional subreports As Variant, Optional parameters As Variant, Optional sorts As Variant, Optional default As Boolean = False, Optional formulas As Variant)
Dim subreport As Variant
Dim parameter As Variant
Dim formula As Variant
Dim sort As Variant

Dim exist_subreport As Boolean

Dim secuencia As Integer

On Error GoTo handler

    exist_subreport = False
    
    crystal.Reset
    crystal.ReportFileName = App.path + "\" & report_name & ".rpt"
    crystal.Connect = string_connection
    
    If Not IsMissing(subreports) Then
        For Each subreport In subreports
            exist_subreport = True
            crystal.SubreportToChange = subreport
            crystal.Connect = string_connection
        Next
    End If
    
    If exist_subreport Then crystal.SubreportToChange = ""
    
    If Not IsMissing(parameters) Then
        secuencia = 0
        For Each parameter In parameters
            crystal.ParameterFields(secuencia) = parameter(0) & ";" & parameter(1) & ";TRUE"
            secuencia = secuencia + 1
        Next
    End If
    
    If Not IsMissing(formulas) Then
        secuencia = 0
        For Each formula In formulas
            crystal.formulas(secuencia) = formula(0) & "= '" & formula(1) & "'"
            secuencia = secuencia + 1
        Next
    End If
    
    If Not IsMissing(sorts) Then
        secuencia = 0
        For Each sort In sorts
            crystal.SortFields(secuencia) = sort
            secuencia = secuencia + 1
        Next
    End If

    If default Then
        crystal.Destination = crptToPrinter
        crystal.Action = 1
    Else
        frmImpresora.Show vbModal
        frmImpresora.cargar crystal
        If Not frmImpresora.Cancel Then crystal.Action = 1
    End If
    
    Exit Sub
    
handler:
    MsgBox "Error de Impresión: " & Err.Number & " - " & Err.description
    Exit Sub
    
End Sub

