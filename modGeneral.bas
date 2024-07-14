Attribute VB_Name = "modGeneral"
Option Explicit

Global dbapp As New clsDB

Global Const cntCertificado = "uvspes"

Global Const cntNotificacion15 = 1
Global Const cntNotificacionOC = 2
Global Const cntNotificacionCorte = 3

Global Const cntPagoManual = 0
Global Const cntPagoPagoFacil = 1
Global Const cntPagoRapiPago = 2

Global Const filePMC = 10

Global cntIVA As Variant

Public Sub marcarseleccion(pTextBox As TextBox)
    
    pTextBox.SelStart = 0
    pTextBox.SelLength = Len(pTextBox.Text)

End Sub

