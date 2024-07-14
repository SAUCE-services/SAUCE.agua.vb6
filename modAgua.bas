Attribute VB_Name = "modAgua"
Option Explicit

Global dbmy As New clsDBMy

Global Const cntCertificado = "uvspes"

Global cntIVA As Variant

Public Sub marcarseleccion(pTextBox As TextBox)
    
    pTextBox.SelStart = 0
    pTextBox.SelLength = Len(pTextBox.Text)

End Sub

