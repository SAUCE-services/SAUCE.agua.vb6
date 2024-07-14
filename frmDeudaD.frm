VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeudaD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Deuda"
   ClientHeight    =   5085
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11775
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "Recalcular"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtTotalFacturas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDescortar 
      Cancel          =   -1  'True
      Caption         =   "Levantar CORTE"
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin Crystal.CrystalReport crpDeuda 
      Left            =   6000
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   3360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdLiquidacion 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdCuota 
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Plan de Cuotas"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Liquidaciones"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmDeudaD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cliente As New clsMODCliente

Private Sub fillDeuda()
Dim factura As New clsMyAFactura
Dim imputado As New clsMyAImputado
Dim recibo As New clsMyARecibo
Dim ncredito As New clsMyANCredito
Dim periodo As New clsRESTPeriodo
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

Dim periodos As Collection
Dim facturas As Collection
Dim cuotas As Collection

Dim interes_factura As Currency
Dim total_factura As Currency
Dim total_recibos As Currency
Dim total_notas As Currency
Dim total_facturas As Currency
Dim total_cuota As Currency

    Me.txtTotalFacturas.Text = "0.00"
    
    Set periodos = periodo.collectionAll
    
    Me.grdLiquidacion.Rows = 1
    Me.grdCuota.Rows = 1

    Set facturas = factura.collectionDeudaByClienteId(cliente.clienteId, dbapp)
    
    If facturas.Count = 0 Then
        Me.grdLiquidacion.AddItem modGrid.array2itemGrid(Array("", "No Registra Deuda", "", ""))
    Else
        total_facturas = 0
        For Each factura In facturas
            total_factura = factura.total
            total_recibos = 0
            For Each imputado In imputado.collectionByComprobante(1, factura.puntoVta, factura.nroComprob, dbapp, factura.clienteId)
                recibo.serieId = imputado.serieId
                recibo.numero = imputado.numeroID
                recibo.findByPrimaryKey dbapp
                total_recibos = total_recibos + recibo.total
            Next
            total_notas = 0
            For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
                total_notas = total_notas + ncredito.total
            Next
            
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & factura.periodoId) Then Set periodo = periodos("k." & factura.periodoId)
            
            interes_factura = interes(total_factura - total_recibos - total_notas, factura.tasa, periodo.fechaPrimero, Date)
            Me.grdLiquidacion.AddItem modGrid.array2itemGrid(Array(periodo.fechaPrimero, Format(total_factura, "0.00"), Format(total_recibos, "0.00"), Format(total_notas, "0.00"), Format(interes_factura, "0.00"), Format(total_factura - total_recibos - total_notas + interes_factura, "0.00")))
            total_facturas = total_facturas + total_factura + interes_factura - total_recibos - total_notas
        Next
        Me.txtTotalFacturas.Text = Format(total_facturas, "0.00")
    End If
    
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If deuda.planID = 0 Then
        Me.grdCuota.AddItem modGrid.array2itemGrid(Array("", "No Registra Deuda", ""))
        Exit Sub
    End If
    
    If deuda.planID <> 0 Then
        If deuda.pagado <> 0 Then
            Me.grdCuota.AddItem modGrid.array2itemGrid(Array("", "No Registra Deuda", ""))
            Exit Sub
        End If
    End If
    
    Set cuotas = cuota.collectionPendienteByPlanID(deuda.clienteId, deuda.planID, dbapp)
    
    If cuotas.Count = 0 Then
        Me.grdCuota.AddItem modGrid.array2itemGrid(Array("", "No Registra Deuda", ""))
        Exit Sub
    End If
    
    For Each cuota In cuotas
        total_cuota = cuota.importe
        
        For Each imputado In imputado.collectionByComprobante(2, cuota.planID, cuota.cuotaID, dbapp, cuota.clienteId)
            recibo.serieId = imputado.serieId
            recibo.numero = imputado.numeroID
            recibo.findByPrimaryKey dbapp
            
            If recibo.total <> 0 Then total_cuota = total_cuota - recibo.total
        Next
        
        Me.grdCuota.AddItem modGrid.array2itemGrid(Array(cuota.cuotaID, cuota.fechaVencimiento, Format(total_cuota, "0.00")))
    Next

End Sub

Private Sub cmdDescortar_Click()
Dim anotador As clsMyAAnotador

Dim clienteRep As New clsREPCliente
        
    If cliente.clienteId = 0 Then Exit Sub
    
    ' Se agrega novedad en el anotador
    Set anotador = New clsMyAAnotador
    anotador.clienteId = cliente.clienteId
    anotador.anotacion = "Se LEVANTA Corte con Fecha " & Date
    If Not anotador.add(dbapp) Then MsgBox "No se pudo ANOTAR la Novedad"
    Set cliente = clienteRep.findLastByClienteId(cliente.clienteId)
    cliente.cortado = 0
    Set cliente = clienteRep.save(cliente)
    Me.txtCliente.Text = cliente.textFound
    
    MsgBox "Corte LEVANTADO"

End Sub

Private Sub cmdImprimir_Click()
Dim consul As String
Dim cuit As String
Dim mens As String
Dim importeInteres As Currency
Dim importeDeuda As Currency
Dim fecha As Date

Dim listado As New clsMyAListado
Dim operador As New clsMyAOperador
Dim factura As New clsMyAFactura
Dim imputado As New clsMyAImputado
Dim recibo As New clsMyARecibo
Dim ncredito As New clsMyANCredito
Dim periodo As New clsRESTPeriodo
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

Dim clienteRep As New clsREPCliente

Dim impresionService As New clsCtlImpresion

On Error Resume Next
    
    frmDeudaD.MousePointer = 11
    
    listado.truncate dbapp
    operador.findLast dbapp
    If clienteRep.collectionActivos.Count = 0 Then
        frmDeudaD.MousePointer = 0
        Exit Sub
    End If
    
    cuit = Left(operador.cuit, 2) & "-" & Mid(operador.cuit, 3, 8) & "-" & Right(operador.cuit, 1)
    Select Case operador.situacionIVA
        Case 1:
            mens = "Resp. Inscripto"
        Case 2:
            mens = "Resp. No Inscripto"
        Case 3:
            mens = "Cons. Final"
        Case 4:
            mens = "Exento"
        Case 5:
            mens = "No Responsable"
    End Select
    
    For Each factura In factura.collectionDeudaByClienteId(cliente.clienteId, dbapp)
        importeDeuda = factura.total
        For Each imputado In imputado.collectionByComprobante(1, factura.puntoVta, factura.nroComprob, dbapp)
            recibo.serieId = imputado.serieId
            recibo.numero = imputado.numeroID
            recibo.findByPrimaryKey dbapp
            If recibo.autoID > 0 Then importeDeuda = importeDeuda - recibo.total
        Next
        For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
            importeDeuda = importeDeuda - ncredito.total
        Next
        periodo.periodoId = factura.periodoId
        periodo.findByPrimaryKey
        If periodo.fechaPrimero < Date Then
            Set listado = New clsMyAListado
            listado.c1 = "Factura"
            listado.c2 = Right("0000" & factura.puntoVta, 4) & "-" & Right("00000000" & factura.nroComprob, 8)
            listado.c3 = periodo.fechaPrimero
            listado.n1 = Format(importeDeuda, "#,###,##0.00")
            listado.n2 = Format(interes(importeDeuda, factura.tasa, periodo.fechaPrimero, Date), "#,###,##0.00")
            listado.n3 = Format(importeDeuda + interes(importeDeuda, factura.tasa, periodo.fechaPrimero, Date), "#,###,##0.00")
            listado.add dbapp
        End If
    Next
    
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If deuda.autoID > 0 Then
        If Not deuda.pagado Then
            For Each cuota In cuota.collectionDeudaByPlanID(cliente.clienteId, deuda.planID, dbapp)
                importeDeuda = cuota.importe
                For Each imputado In imputado.collectionByComprobante(2, cuota.planID, cuota.cuotaID, dbapp, cliente.clienteId)
                    recibo.serieId = imputado.serieId
                    recibo.numero = imputado.numeroID
                    recibo.findByPrimaryKey dbapp
                    If recibo.autoID > 0 Then importeDeuda = importeDeuda - recibo.total
                Next
                
                Set listado = New clsMyAListado
                listado.c1 = "Cuota"
                listado.c2 = deuda.planID & "-" & Right("000000" & cuota.cuotaID, 6)
                listado.c3 = cuota.fechaVencimiento
                listado.n1 = Format(importeDeuda, "#,###,##0.00")
                importeInteres = 0
                If cuota.fechaVencimiento < Date Then
                    importeInteres = Format(interes(importeDeuda, deuda.tasa, cuota.fechaVencimiento, Date), "#,###,##0.00")
                    If importeInteres < 0 Then importeInteres = 0
                End If
                listado.n2 = Format(importeInteres, "#,###,##0.00")
                listado.n3 = Format(importeDeuda + importeInteres, "#,###,##0.00")
                listado.add dbapp
            Next
        End If
    End If
    
    impresionService.printReport Me.crpDeuda, "rptDeuda", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Resumen de Deuda por Conexión"), _
        Array("conex", cliente.clienteId), _
        Array("clien", Me.txtCliente.Text))
    
    frmDeudaD.MousePointer = 0

End Sub

Private Sub cmdRecalcular_Click()

    fillDeuda
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    modGrid.makeGrid Me.grdLiquidacion, Array(Array("Vencimiento", 1700), Array("Importe Factura", 1800), Array("Recibos", 1800), Array("Notas de Crédito", 1800), Array("Intereses", 1800), Array("Importe con Intereses", 1800)), 0, 1, flexSelectionByRow
    modGrid.makeGrid Me.grdCuota, Array(Array("Cuota", 1650), Array("Vencimiento", 1650), Array("Importe", 1650)), 0, 1, flexSelectionByRow

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    fillDeuda

End Sub

