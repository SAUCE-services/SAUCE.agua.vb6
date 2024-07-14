VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function datosConsumo(clienteId As Long, periodoId As Integer, medidorID As String, fechaEmision As Date, fechaActual As Variant, estadoActual As Long, fechaAnterior As Variant, estadoAnterior As Long, db As clsDB) As Long
Dim desconexion As New clsMyADesconexion
Dim periodo As New clsRESTPeriodo
Dim lectura As New clsMyALectura
Dim medidor As New clsMyAMedidor
Dim medicion As clsMyAMedicion

Dim desconectado As Boolean

    fechaActual = CDate("01/01/1980")
    fechaAnterior = CDate("01/01/1980")
    estadoActual = 0
    estadoAnterior = 0
    datosConsumo = 0
    
    desconectado = False
    
    periodo.periodoId = periodoId
    periodo.findByPrimaryKey
    
    desconexion.clienteId = clienteId
    desconexion.fechaDesconexion = fechaEmision
    
    desconexion.findByPrimaryKey db
    
    If desconexion.autoID <> 0 Then
        If IsNull(desconexion.fechaReconexion) Then desconectado = True
        If desconexion.fechaReconexion > periodo.fechaFin Then If desconexion.fechaDesconexion <= periodo.fechaInicio Then desconectado = True
    End If
    
    If desconectado Then Exit Function
    
    medidor.clienteId = clienteId
    medidor.findByClienteID db
    
    lectura.medidorID = medidorID
    lectura.periodoId = periodoId
    lectura.findByPrimaryKey db
    
    fechaActual = lectura.fechaLectura
    estadoActual = medidor.estadoInicio
    If lectura.autoID > 0 Then
        If medidor.fechaColocacion <= periodo.fechaInicio Then
            estadoActual = lectura.estado
                            
            Set medicion = New clsMyAMedicion
            medicion.clienteId = clienteId
            medicion.periodoId = lectura.periodoId
            medicion.medidorID = medidorID
            medicion.fechaLectura = lectura.fechaLectura
            medicion.estado = lectura.estado
            
            medicion.save db
        End If
    Else
        Set medicion = New clsMyAMedicion
        medicion.clienteId = clienteId
        medicion.periodoId = periodoId
        medicion.medidorID = medidorID
        medicion.fechaLectura = Date - 365
        medicion.estado = 0
        
        medicion.save db
    End If
    
    lectura.medidorID = medidorID
    lectura.periodoId = periodoId
    lectura.findByMedidorIDPrev db
    
    fechaAnterior = lectura.fechaLectura
    estadoAnterior = medidor.estadoInicio
    If lectura.autoID > 0 Then
        If medidor.fechaColocacion <= periodo.fechaFin Then
            estadoAnterior = lectura.estado
                            
            Set medicion = New clsMyAMedicion
            medicion.clienteId = clienteId
            medicion.periodoId = periodoId - 1
            medicion.medidorID = medidorID
            medicion.fechaLectura = lectura.fechaLectura
            medicion.estado = lectura.estado
            
            medicion.save db
        End If
    Else
        Set medicion = New clsMyAMedicion
        medicion.clienteId = clienteId
        medicion.periodoId = periodoId - 1
        medicion.medidorID = medidorID
        medicion.fechaLectura = Date - 365
        medicion.estado = 0
        
        medicion.save db
    End If
    
    datosConsumo = estadoActual - estadoAnterior

End Function

Public Sub drawConsumo(pConsumo As PictureBox, clienteId As Long, periodoId As Integer, medidorID As String, db As clsDB)
Dim lx As Long
Dim ly As Long
Dim cx As Long
Dim cy As Long
Dim uy As Long
Dim ux As Long
Dim mx As Long
Dim vx As Long

Dim ct As Integer
Dim np As Integer
Dim per As Integer
Dim anio As Integer

Dim cons(12) As Currency
Dim Max As Currency
Dim est As Currency
Dim vax As Currency

Dim periodo As New clsRESTPeriodo
Dim objMLec As New clsMyALectura
Dim objMOpe As New clsMyAOperador
    
    objMOpe.findLast db
    
    pConsumo.Cls
    lx = pConsumo.Width
    ly = pConsumo.Height
    cx = lx / 8
    cy = ly - ly / 5
    uy = ly / 10
    ux = lx / 16
    For ct = 1 To 12
        cons(ct) = 0
    Next ct
    np = periodoId
    ct = 12
    Max = 0
    Do
        periodo.periodoId = np
        periodo.findByPrimaryKey
        If periodo.uid <> "" Then
            objMLec.medidorID = medidorID
            objMLec.periodoId = np
            objMLec.findByPrimaryKey db
            If objMLec.autoID > 0 Then
                est = objMLec.estado
                objMLec.periodoId = np - 1
                objMLec.findByPrimaryKey db
                If objMLec.autoID = 0 Then
                    cons(ct) = est
                Else
                    cons(ct) = est - objMLec.estado
                End If
                If np = 1 Then cons(ct) = 0
                If cons(ct) > Max Then Max = cons(ct)
            End If
        End If
        ct = ct - 1
        np = np - 1
    Loop While ct > 0 And objMLec.autoID > 0
    mx = (Int(Max / 10) + 1) * 10
    vax = mx / 7
    vx = Int(vax)
    mx = vx * 7
    pConsumo.CurrentX = 100
    pConsumo.CurrentY = cy - uy * 8
    pConsumo.Print "( m ³ )"
    pConsumo.CurrentX = cx + ux * 12 + 250
    pConsumo.CurrentY = cy + uy - 300
    If objMOpe.periodoFactura = 1 Then
        pConsumo.Print "Bimestre"
    Else
        pConsumo.Print "Mes"
    End If
    pConsumo.CurrentX = cx + ux * 12 + 250
    pConsumo.CurrentY = cy + uy
    pConsumo.Print "Año"
    pConsumo.Line (cx, cy)-(cx, cy - uy * 7)
    pConsumo.Line (cx, cy)-(cx + ux * 12, cy)
    For ct = 0 To 7
        pConsumo.CurrentX = 100
        pConsumo.CurrentY = cy - uy * ct - 100
        pConsumo.Print Right("     " & vx * ct, 5)
        pConsumo.Line (cx, cy - uy * ct)-(cx + ux * 12, cy - uy * ct)
    Next ct
    np = periodoId
    per = 1
    periodo.periodoId = np
    periodo.findByPrimaryKey
    If periodo.uid <> "" Then
        anio = Year(periodo.fechaInicio)
        np = np - 1
        Do
            periodo.periodoId = np
            periodo.findByPrimaryKey
            If periodo.uid <> 0 Then If anio = Year(periodo.fechaInicio) Then per = per + 1
            np = np - 1
        Loop While anio = Year(periodo.fechaInicio) And periodo.uid <> ""
    End If
    np = periodoId
    ct = 12
    Do
        periodo.periodoId = np
        periodo.findByPrimaryKey
        If periodo.uid <> "" And Max Then pConsumo.Line (cx + ux * ct - 75, cy)-(cx + ux * ct - 250, cy - Int((cons(ct) / vx * uy))), , BF
        pConsumo.CurrentX = cx + ux * ct - 300
        pConsumo.CurrentY = cy + uy - 300
        pConsumo.Print per
        If per = 1 Then
            pConsumo.CurrentX = cx + ux * ct - 300
            pConsumo.CurrentY = cy + uy
            pConsumo.Print anio
            anio = anio - 1
        End If
        per = per - 1
        If per = 0 Then
            If objMOpe.periodoFactura = 1 Then
                per = 6
            Else
                per = 12
            End If
        End If
        ct = ct - 1
        np = np - 1
    Loop While ct > 0
    
    SavePicture pConsumo.Image, "c:\temp\consumo.bmp"

End Sub

Public Function lastLecturaByClienteId(clienteId As Long) As Currency
Dim estado As Currency

Dim medidor As New clsMyAMedidor
Dim lectura As New clsMyALectura

    medidor.clienteId = clienteId
    medidor.findColocadoByClienteID dbapp
    If medidor.autoID = 0 Then
        lastLecturaByClienteId = 0
        Exit Function
    End If
    
    estado = 0
    lectura.medidorID = medidor.medidorID
    lectura.findLast dbapp
    If lectura.autoID > 0 Then
        estado = lectura.estado
    Else
        medidor.medidorID = medidor.medidorID
        medidor.findByMedidorID dbapp
        If medidor.autoID > 0 Then estado = medidor.estadoInicio
    End If
    
    lastLecturaByClienteId = estado

End Function
