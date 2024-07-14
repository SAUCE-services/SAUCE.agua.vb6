VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImpLiqIndivD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Liquidación Individual"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9885
   Begin VB.TextBox txtDomicilio 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1560
      Width           =   9375
   End
   Begin VB.TextBox txtEmision 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtLiquidacion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picConsumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   11880
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
   End
   Begin Crystal.CrystalReport crpLiquidacion 
      Left            =   7200
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   7455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      ToolTipText     =   "Fin de la TAREA"
      Top             =   5760
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetalle 
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   39
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "SubTotal"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   37
      Top             =   4920
      Width           =   645
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblIvaRI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   35
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblIvaRN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   34
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Resp. Inscripto"
      Height          =   195
      Index           =   14
      Left            =   4080
      TabIndex        =   33
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Resp. No Insc."
      Height          =   195
      Index           =   15
      Left            =   6000
      TabIndex        =   32
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "a Pagar - 1er Vto"
      Height          =   195
      Index           =   16
      Left            =   4080
      TabIndex        =   31
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "a Pagar - 2do Vto"
      Height          =   195
      Index           =   17
      Left            =   6000
      TabIndex        =   30
      Top             =   5520
      Width           =   1245
   End
   Begin VB.Label lblPrimero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblSegundo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblIvaCF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Cons. Final"
      Height          =   195
      Index           =   19
      Left            =   2160
      TabIndex        =   26
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Medidor"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label lblMedidor 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Medición ACTUAL"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   23
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Medición ANTERIOR"
      Height          =   195
      Index           =   7
      Left            =   2400
      TabIndex        =   22
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label lblEstadoActual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblEstadoAnterior 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de LECTURA"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   19
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Label lblFechaActual 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblFechaAnterior 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   16
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "REGISTRADO (m³)"
      Height          =   195
      Index           =   10
      Left            =   7920
      TabIndex        =   15
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Label lblConsumo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Consumo"
      Height          =   195
      Index           =   11
      Left            =   7920
      TabIndex        =   13
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Emisión"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Liquidacion"
      Height          =   195
      Index           =   0
      Left            =   7920
      TabIndex        =   10
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Rubros Facturados"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   7920
      TabIndex        =   3
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "frmImpLiqIndivD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub fillDetalle()
Dim blnIva As Boolean

Dim consumoService As New clsCtlConsumo
Dim ctlFac As New clsCtlFactura

Dim factura As New clsMyAFactura
Dim detalle As New clsMyADetalle
Dim medidor As New clsMyAMedidor
Dim periodo As New clsRESTPeriodo

Dim lngConsumo As Long
Dim lngAnterior As Long
Dim lngActual As Long

Dim datAnterior As Date
Dim datActual As Date
Dim datFecha As Date

    Me.grdDetalle.Rows = 1
    Me.txtLiquidacion.Text = ""
    Me.txtEmision.Text = Date
    Me.lblMedidor.Caption = ""
    Me.lblConsumo.Caption = ""
    Me.lblFechaActual.Caption = ""
    Me.lblFechaAnterior.Caption = ""
    Me.lblEstadoActual.Caption = ""
    Me.lblEstadoAnterior.Caption = ""
    Me.lblSubtotal.Caption = ""
    Me.lblIvaCF.Caption = ""
    Me.lblIvaRI.Caption = ""
    Me.lblIvaRN.Caption = ""
    Me.lblPrimero.Caption = ""
    Me.lblSegundo.Caption = ""

    With factura
        .clienteId = cliente.clienteId
        .periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
        
        .findByClientePeriodo dbapp
    End With
    
    With medidor
        .clienteId = cliente.clienteId
        
        .findByClienteID dbapp
    End With
    
    periodo.periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
    periodo.findByPrimaryKey
    
    datFecha = Date
    If factura.autoID > 0 Then datFecha = factura.fecha
    
    lngConsumo = consumoService.datosConsumo(medidor.clienteId, periodo.periodoId, medidor.medidorID, datFecha, datActual, lngActual, datAnterior, lngAnterior, dbapp)
    
    Me.lblMedidor.Caption = medidor.medidorID
    Me.lblConsumo.Caption = lngConsumo
    Me.lblFechaActual.Caption = datActual
    Me.lblFechaAnterior.Caption = datAnterior
    Me.lblEstadoActual.Caption = lngActual
    Me.lblEstadoAnterior.Caption = lngAnterior
    
    If factura.autoID = 0 Then Exit Sub
    
    With factura
        Me.txtLiquidacion.Text = .puntoVta & "/" & .nroComprob
        Me.txtEmision.Text = .fecha
    End With
    
    Me.grdDetalle.Redraw = False
    For Each detalle In detalle.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
        Me.grdDetalle.AddItem modGrid.array2itemGrid(Array(Format(detalle.rubroID, "00"), detalle.concepto, detalle.cantidad, Format(detalle.precioUnitario, "0.00"), Format(detalle.precioUnitario * detalle.cantidad, "0.00")))
        blnIva = False
        If detalle.IVA <> 0 Then blnIva = True
        modGrid.letCheckCell Me.grdDetalle, Me.grdDetalle.Rows - 1, 5, blnIva
    Next
    Me.grdDetalle.Redraw = True
    
    Me.lblSubtotal.Caption = Format(factura.total - factura.ivacf - factura.ivari - factura.ivarn, "0.00")
    Me.lblIvaCF.Caption = Format(factura.ivacf, "0.00")
    Me.lblIvaRI.Caption = Format(factura.ivari, "0.00")
    Me.lblIvaRN.Caption = Format(factura.ivarn, "0.00")
    Me.lblPrimero.Caption = Format(factura.total, "0.00")
    Me.lblSegundo.Caption = Format(factura.total + interes(factura.total, factura.tasa, periodo.fechaPrimero, periodo.fechaSegundo), "0.00")

End Sub

Private Sub cboPeriodo_Click()

    fillDetalle
    
End Sub

Private Sub cmdImprimir_Click()
Dim factura As New clsMyAFactura

Dim consumoService As New clsCtlConsumo
Dim liquidacion_service As New clsCtlLiquidacion

    With factura
        .clienteId = cliente.clienteId
        .periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
        
        .findByClientePeriodo dbapp
        
        If .autoID = 0 Then
            MsgBox "ERROR: Período NO Liquidado . . ."
            Exit Sub
        End If
    End With
    
    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    liquidacion_service.printLiquidacion Me.hWnd, factura.puntoVta, factura.nroComprob, dbapp, Me.picConsumo, Me.crpLiquidacion

    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    modGrid.makeGrid Me.grdDetalle, Array(Array("Rub", 500), Array("Concepto", 4700), Array("Cantidad", 1000), Array("Prec. Unitario", 1200), Array("Imp. Parciales", 1200), Array("iva", 400)), 0, 1, flexSelectionByRow
    
    periodo.fillCombo Me.cboPeriodo
    
    Me.txtEmision.Text = Date
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    Me.txtDomicilio.Text = cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & " " & cliente.inmuebleLocalidad
    
    KeyAscii = 0
    
    fillDetalle

End Sub

