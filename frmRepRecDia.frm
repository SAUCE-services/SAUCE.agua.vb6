VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepRecDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recaudación Diaria"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7920
   Begin Crystal.CrystalReport crpReporte 
      Left            =   5280
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdComprobantes 
      Height          =   5655
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42575
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total Recaudación"
      Height          =   195
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobantes"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Pago"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmRepRecDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim impresionService As New clsCtlImpresion

Dim operador As New clsMyAOperador

Dim cuit As String
Dim mens As String

    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    operador.findLast dbapp
    
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
    
    impresionService.printReport Me.crpReporte, "rptRecDiaria", dbapp.stringConnection, Array("sCuota"), Array(Array("pFechaPago", toReportDate(Me.dtpFecha.value)), Array("pTotal", Val(Me.txtTotal.Text))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Recaudación Diaria"), _
        Array("info1", "Fecha Pago : " & Me.dtpFecha.value))
    
    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdRevisar_Click()
Dim curInteres As Currency
Dim curTotalDia As Currency

Dim clienteId As Long

Dim periodoId As Integer

Dim factura As New clsMyAFactura
Dim cliente As clsMODCliente
Dim periodo As New clsRESTPeriodo
Dim cuota As New clsMyACuota
Dim ncredito As New clsMyANCredito
Dim recibo As New clsMyARecibo

Dim clienteRep As clsREPCliente

    clienteId = 0
    periodoId = 0
    curTotalDia = 0

    Me.grdComprobantes.Rows = 1
    Me.grdComprobantes.Redraw = False
    For Each factura In factura.collectionByPago(Me.dtpFecha.value, dbapp)
        If clienteId <> factura.clienteId Then
            Set clienteRep = New clsREPCliente
            Set cliente = clienteRep.findLastByClienteId(factura.clienteId)
            Set clienteRep = Nothing
        End If
        If periodoId <> factura.periodoId Then
            periodo.periodoId = factura.periodoId
            periodo.findByPrimaryKey
        End If
        
        curInteres = 0
        If factura.fechapago > periodo.fechaPrimero Then curInteres = modInteres.interes(factura.total, factura.tasa, periodo.fechaPrimero, periodo.fechaSegundo)
        factura.interes = curInteres
        factura.save dbapp
        
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Liquidación", cliente.apellidonombre, factura.puntoVta & "/" & factura.nroComprob, Format(factura.total + curInteres, "0.00")))
        curTotalDia = curTotalDia + factura.total + curInteres
    Next
    
    For Each cuota In cuota.collectionByPago(Me.dtpFecha.value, dbapp)
        If clienteId <> cuota.clienteId Then
            Set clienteRep = New clsREPCliente
            Set cliente = clienteRep.findLastByClienteId(cuota.clienteId)
            Set clienteRep = Nothing
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Cuota", cliente.apellidonombre, cuota.planID & "/" & cuota.cuotaID, Format(cuota.importe, "0.00")))
        curTotalDia = curTotalDia + cuota.importe
    Next
    
    For Each ncredito In ncredito.collectionByPago(Me.dtpFecha.value, dbapp)
        If clienteId <> ncredito.clienteId Then
            Set clienteRep = New clsREPCliente
            Set cliente = clienteRep.findLastByClienteId(ncredito.clienteId)
            Set clienteRep = Nothing
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("N.Crédito", cliente.apellidonombre, ncredito.serieId & "/" & ncredito.numero, Format(-ncredito.total, "0.00")))
        curTotalDia = curTotalDia - ncredito.total
    Next
    
    For Each recibo In recibo.collectionByPago(Me.dtpFecha.value, dbapp)
        If clienteId <> recibo.clienteId Then
            Set clienteRep = New clsREPCliente
            Set cliente = clienteRep.findLastByClienteId(recibo.clienteId)
            Set clienteRep = Nothing
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Recibo", cliente.apellidonombre, recibo.serieId & "/" & recibo.numero, Format(recibo.total, "0.00")))
        curTotalDia = curTotalDia + recibo.total
    Next
    Me.grdComprobantes.Redraw = True
    
    Me.txtTotal.Text = Format(curTotalDia, "0.00")
    
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpFecha_Change()

    Me.grdComprobantes.Rows = 1
    Me.txtTotal.Text = ""
    Me.cmdImprimir.Enabled = False
    
End Sub

Private Sub Form_Load()

    Me.dtpFecha.value = Date
    
    modGrid.makeGrid Me.grdComprobantes, Array(Array("Tipo", 1000), Array("Cliente", 3500), Array("Número", 1200), Array("Total", 1200)), 0, 1, flexSelectionByRow
    
    Me.cmdImprimir.Enabled = False
    
End Sub
