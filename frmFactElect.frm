VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFactElect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación Electrónica"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11775
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtIva27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtNeto27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5760
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   7800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crpFactura 
      Left            =   10800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtExento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtNeto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtIva 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetalle 
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.TextBox txtTotalMarcado 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdInvertir 
      Caption         =   "Invertir"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtTotalFacturas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109379585
      CurrentDate     =   42575
   End
   Begin MSFlexGridLib.MSFlexGrid grdFacturas 
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "Facturar"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109379585
      CurrentDate     =   42575
   End
   Begin MSFlexGridLib.MSFlexGrid grdNotaCredito 
      Height          =   1335
      Left            =   240
      TabIndex        =   28
      Top             =   6360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "(Doble Click en el CAE para Imprimir)"
      Height          =   195
      Index           =   11
      Left            =   9000
      TabIndex        =   30
      Top             =   840
      Width           =   2580
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Notas de Crédito"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   29
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA 27%"
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   26
      Top             =   5520
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Neto 27%"
      Height          =   195
      Index           =   8
      Left            =   2160
      TabIndex        =   24
      Top             =   5520
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Exento"
      Height          =   195
      Index           =   7
      Left            =   7920
      TabIndex        =   20
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Neto 21%"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   5520
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA 21%"
      Height          =   195
      Index           =   5
      Left            =   4080
      TabIndex        =   16
      Top             =   5520
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Index           =   4
      Left            =   9840
      TabIndex        =   14
      Top             =   5520
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Detalle Factura"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Facturas"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmFactElect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub calcularMarcados()
Dim curMarcados As Currency

Dim intCiclo As Integer

    curMarcados = 0
    
    For intCiclo = 1 To Me.grdFacturas.Rows - 1
        If modGrid.getCheckCell(Me.grdFacturas, intCiclo, 5) Then curMarcados = curMarcados + Val(Me.grdFacturas.TextMatrix(intCiclo, 4))
    Next intCiclo
    
    Me.txtTotalMarcado.Text = Format(curMarcados, "0.00")
    
End Sub

Private Sub fillDetalle()
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

Dim fefactura As clsMyAFEFactura
Dim fedetalle As New clsMyAFEDetalle

Dim ncredito As New clsMyANCredito

Dim ctlFac As New clsCtlFactura

Dim blnIva As Boolean

Dim key As String

Dim puntoVta As Integer

Dim total As Currency

Dim nroComprob As Long

    key = Trim(Str(Me.grdFacturas.RowData(Me.grdFacturas.row)))
    puntoVta = Val(Left(key, Len(key) - 8))
    nroComprob = Val(Right(key, 8))
    factura.puntoVta = puntoVta
    factura.nroComprob = nroComprob
    factura.findByPrimaryKey dbapp
    
    periodo.periodoId = factura.periodoId
    periodo.findByPrimaryKey
    
    Me.grdDetalle.Rows = 1
    Me.grdDetalle.Redraw = False
    For Each fedetalle In ctlFac.detalles2detartic(factura, dbapp, fefactura)
        If fedetalle.cantidad * fedetalle.unitarioSinIva <> 0 Then
            Me.grdDetalle.AddItem modGrid.array2itemGrid(Array(Format(fedetalle.rubroID, "00"), fedetalle.concepto, fedetalle.cantidad, Format(fedetalle.unitarioSinIva, "0.00"), Format(fedetalle.unitarioSinIva * fedetalle.cantidad, "0.00")))
            blnIva = False
            If fedetalle.unitarioConIva > fedetalle.unitarioSinIva Then blnIva = True
            modGrid.letCheckCell Me.grdDetalle, Me.grdDetalle.Rows - 1, 5, blnIva
        End If
    Next
    Me.grdDetalle.Redraw = True
    
    Me.txtNeto27.Text = Format(fefactura.neto27, "0.00")
    Me.txtNeto.Text = Format(fefactura.neto, "0.00")
    Me.txtIva27.Text = Format(fefactura.iva27, "0.00")
    Me.txtIva.Text = Format(fefactura.IVA, "0.00")
    Me.txtExento.Text = Format(fefactura.exento, "0.00")
    Me.txtTotal.Text = Format(fefactura.importe, "0.00")
    
    total = 0
    Me.grdNotaCredito.Rows = 1
    Me.grdNotaCredito.Redraw = False
    For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
        If ncredito.anulado = 0 Then
            total = total + ncredito.total
            Me.grdNotaCredito.AddItem modGrid.array2itemGrid(Array(modConv.formatNumComprobante(ncredito.serieId, ncredito.numero), ncredito.fecha, Format(ncredito.total, "0.00"), Format(ncredito.ivacf, "0.00"), Format(ncredito.ivari, "0.00"), Format(ncredito.ivarn, "0.00"), Format(total, "0.00")))
        End If
    Next
    Me.grdNotaCredito.Redraw = True
    
End Sub

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.value = Me.dtpDesde.value
    
End Sub

Private Sub cmdFacturar_Click()
Dim ciclo As Integer
Dim puntoVenta As Integer

Dim numeroComprobante As Long

Dim cae As String
Dim key As String

Dim ctlFac As New clsCtlFactura
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = Me.grdFacturas.Rows
    Me.prbProgreso.value = 1

    For ciclo = 1 To Me.grdFacturas.Rows - 1
        DoEvents
        
        If Me.prbProgreso.Max > Me.prbProgreso.value Then
            Me.prbProgreso.value = Me.prbProgreso.value + 1
            Me.prbProgreso.Refresh
        End If
        
        If modGrid.getCheckCell(Me.grdFacturas, ciclo, 5) Then
            key = Trim(Str(Me.grdFacturas.RowData(ciclo)))
            puntoVenta = Val(Left(key, Len(key) - 8))
            numeroComprobante = Val(Right(key, 8))
            cae = ctlFac.makeFactura(puntoVenta, numeroComprobante, dbapp)
            If cae <> "" Then
                Me.grdFacturas.TextMatrix(ciclo, 6) = cae
                modGrid.letCheckCell Me.grdFacturas, ciclo, 5, False
            End If
        End If
    Next ciclo
    
End Sub

Private Sub cmdInvertir_Click()
Dim intCiclo As Integer

    If Me.grdFacturas.Rows = 1 Then Exit Sub
    
    For intCiclo = 1 To Me.grdFacturas.Rows - 1
        If Me.grdFacturas.TextMatrix(intCiclo, 6) = "" Then modGrid.letCheckCell Me.grdFacturas, intCiclo, 5, Not modGrid.getCheckCell(Me.grdFacturas, intCiclo, 5)
    Next intCiclo
    
    calcularMarcados
    
End Sub

Private Sub cmdRevisar_Click()
Dim factura As New clsMyAFactura
Dim cliente As clsMODCliente
Dim periodo As New clsRESTPeriodo
Dim ncredito As New clsMyANCredito
Dim fefactura As New clsMyAFEFactura

Dim clienteRep As New clsREPCliente

Dim clienteId As Long

Dim periodoId As Integer

Dim comprobante As String

Dim totalFactura As Currency
Dim total As Currency
Dim importe As Currency
Dim interes As Currency

    clienteId = 0
    periodoId = 0
    totalFactura = 0
    
    Me.MousePointer = 11

    Me.grdFacturas.Rows = 1
    Me.grdFacturas.Redraw = False
    For Each factura In factura.collectionByPeriodoPago(Me.dtpDesde.value, Me.dtpHasta.value, dbapp)
        If clienteId <> factura.clienteId Then
            Set cliente = clienteRep.findLastByClienteId(factura.clienteId)
        End If
        If periodoId <> factura.periodoId Then
            periodo.periodoId = factura.periodoId
            periodo.findByPrimaryKey
        End If
        clienteId = factura.clienteId
        periodoId = factura.periodoId
        comprobante = "Liquidacion " & Format(factura.puntoVta, "0000") & "-" & Format(factura.nroComprob, "00000000")
        interes = 0
        ' Descuenta el monto de las notas de crédito
        total = factura.total
        For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
            total = total - ncredito.total
        Next
        
        If factura.fechapago > periodo.fechaPrimero Then interes = modInteres.interes(total, factura.tasa, periodo.fechaPrimero, periodo.fechaSegundo)
        importe = total + interes
        fefactura.puntoVta = factura.puntoVta
        fefactura.nroComprob = factura.nroComprob
        fefactura.findByLiquidacion dbapp
        Me.grdFacturas.AddItem modGrid.array2itemGrid(Array(cliente.clienteId, cliente.apellidonombre, comprobante, periodo.descripcion, Format(importe, "0.00"), "", fefactura.cae))
        Me.grdFacturas.RowData(Me.grdFacturas.Rows - 1) = factura.puntoVta * 100000000 + factura.nroComprob
        totalFactura = totalFactura + importe
        modGrid.letCheckCell Me.grdFacturas, Me.grdFacturas.Rows - 1, 5, False
    Next
    Me.grdFacturas.Redraw = True
    
    Me.txtTotalFacturas.Text = Format(totalFactura, "0.00")
    Me.txtTotalMarcado.Text = "0.00"
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpDesde_Change()

    Me.grdFacturas.Rows = 1
    Me.grdDetalle.Rows = 1
    Me.grdNotaCredito.Rows = 1
    Me.txtTotalFacturas.Text = "0.00"
    Me.txtTotalMarcado.Text = "0.00"
    Me.txtNeto.Text = "0.00"
    Me.txtTotal.Text = "0.00"
    Me.txtExento.Text = "0.00"
    Me.txtIva.Text = "0.00"

End Sub

Private Sub dtpHasta_Change()

    Me.grdFacturas.Rows = 1
    Me.grdDetalle.Rows = 1
    Me.grdNotaCredito.Rows = 1
    Me.txtTotalFacturas.Text = "0.00"
    Me.txtTotalMarcado.Text = "0.00"
    Me.txtNeto.Text = "0.00"
    Me.txtTotal.Text = "0.00"
    Me.txtExento.Text = "0.00"
    Me.txtIva.Text = "0.00"

End Sub

Private Sub Form_Load()

    modGrid.makeGrid Me.grdFacturas, Array(Array("#ID", 500), Array("Cliente", 3600), Array("Comprobante", 2100), Array("Periodo", 1500), Array("Total", 1000), Array("", 300), Array("CAE", 1900)), 0, 1, flexSelectionFree
    modGrid.makeGrid Me.grdDetalle, Array(Array("Rub", 500), Array("Concepto", 4700), Array("Cantidad", 1000), Array("Prec. Unitario", 1200), Array("Imp. Parciales", 1200), Array("IVA", 400)), 0, 1, flexSelectionByRow
    modGrid.makeGrid Me.grdNotaCredito, Array(Array("Comprobante", 2000), Array("Fecha", 1400), Array("Total", 1500), Array("IVA CF", 1500), Array("IVA 21%", 1500), Array("IVA 10.5%", 1500), Array("Acumulado", 1500)), 0, 1, flexSelectionByRow

    Me.dtpDesde.value = Date
    Me.dtpHasta.value = Date
    
End Sub

Private Sub grdFacturas_DblClick()
Dim impresionService As New clsCtlImpresion

Dim factura As New clsMyAFactura
Dim fefactura As New clsMyAFEFactura

Dim key As String

Dim puntoVta As Integer

Dim nroComprob As Long

    If Me.grdFacturas.col = 2 Then
        fillDetalle
        Exit Sub
    End If

    If Me.grdFacturas.TextMatrix(Me.grdFacturas.row, 6) <> "" Then
        If MsgBox("Imprime ?", vbYesNo, "Impresión Factura") = vbNo Then Exit Sub
        
        With factura
            key = Trim(Str(Me.grdFacturas.RowData(Me.grdFacturas.row)))
            puntoVta = 0
            If Len(key) > 8 Then puntoVta = Val(Left(key, Len(key) - 8))
            nroComprob = Val(Right(key, 8))
            .puntoVta = puntoVta
            .nroComprob = nroComprob
            .findByPrimaryKey dbapp
        End With
        
        fefactura.puntoVta = factura.puntoVta
        fefactura.nroComprob = factura.nroComprob
        fefactura.findByLiquidacion dbapp
        
        If fefactura.tipoCompro = "A" Then
            impresionService.printReport Me.crpFactura, "rptFEFactA", dbapp.stringConnection, Array("sFactura", "sFactura - 01", "sFactura - 02"), Array(Array("pAutoID", fefactura.autoID))
        End If
        If fefactura.tipoCompro = "B" Then
            impresionService.printReport Me.crpFactura, "rptFEFactB", dbapp.stringConnection, Array("sFactura", "sFactura - 01"), Array(Array("pAutoID", fefactura.autoID))
        End If
        Exit Sub
    End If
    
    If Me.grdFacturas.col <> 5 Then Exit Sub
    
    modGrid.letCheckCell Me.grdFacturas, Me.grdFacturas.row, 5, Not modGrid.getCheckCell(Me.grdFacturas, Me.grdFacturas.row, 5)
    
    calcularMarcados

End Sub

