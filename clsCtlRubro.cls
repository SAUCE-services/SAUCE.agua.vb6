VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub updateVigente(db As clsDB)
Dim rubro As New clsMyARubro
Dim rubrovigente As New clsMyARubroVigente

    For Each rubrovigente In rubrovigente.collectionAll(db)
        rubrovigente.delete db
    Next
    
    For Each rubro In rubro.collectionSinRepeticion(db)
        Set rubrovigente = New clsMyARubroVigente
        rubrovigente.rubroID = rubro.rubroID
        rubrovigente.fecha = rubro.fecha
        rubrovigente.rangoID = rubro.rangoID
        rubrovigente.concepto = rubro.concepto
        rubrovigente.precioUnitario = rubro.precioUnitario
        rubrovigente.IVA = rubro.IVA
        rubrovigente.comun = rubro.comun
        rubrovigente.comunSocio = rubro.comunSocio
        rubrovigente.cobro = rubro.cobro
        rubrovigente.desconectado = rubro.desconectado
        rubrovigente.uid = rubro.uid
        rubrovigente.save db
    Next
    
End Sub
