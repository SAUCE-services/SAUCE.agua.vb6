VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepFacPagFec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Pagadas entre Fechas"
   ClientHeight    =   1410
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin Crystal.CrystalReport crpFacturas 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmRepFacPagFec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.value = Me.dtpDesde.value
    
End Sub

Private Sub cmdImprimir_Click()
Dim factura As New clsMyAFactura
Dim listado As New clsMyAListado
Dim cliente As New clsMODCliente
Dim periodo As New clsRESTPeriodo
Dim operador As New clsMyAOperador

Dim clienteRep As New clsREPCliente

Dim clientes As Collection
Dim periodos As Collection

Dim cuit As String
Dim mens As String

Dim impresionService As New clsCtlImpresion

    Me.MousePointer = 11
    
    listado.truncate dbapp
    
    Set clientes = clienteRep.collectionActivos
    Set periodos = periodo.collectionAll
    
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
    
    For Each factura In factura.collectionPagadasByPeriodo(Me.dtpDesde.value, Me.dtpHasta.value, dbapp)
        Set cliente = New clsMODCliente
        Set periodo = New clsRESTPeriodo
        If modCollection.collectionExistElement(clientes, "k." & factura.clienteId) Then Set cliente = clientes("k." & factura.clienteId)
        If modCollection.collectionExistElement(periodos, "k." & factura.periodoId) Then Set periodo = periodos("k." & factura.periodoId)
        
        Set listado = New clsMyAListado
        
        listado.n1 = factura.clienteId
        listado.c1 = Left(cliente.apellidonombre, 25)
        listado.c2 = Right("0000" & factura.puntoVta, 4) & "-" & Right("00000000" & factura.nroComprob, 8)
        listado.c3 = factura.fechapago
        listado.c4 = periodo.descripcion
        If periodo.periodoId = 0 Then
            listado.n2 = factura.total
        Else
            listado.c4 = periodo.descripcion
            If factura.fechapago > periodo.fechaPrimero Then
                listado.n2 = factura.total + interes(factura.total, periodo.tasa, periodo.fechaPrimero, periodo.fechaSegundo)
            Else
                listado.n2 = factura.total
            End If
        End If
        listado.n3 = factura.ivacf
        listado.n4 = factura.ivari
        listado.n5 = factura.ivarn
        listado.add dbapp
    Next
    
    impresionService.printReport Me.crpFacturas, "rptFacturas", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Facturas Pagadas entre Fechas"), _
        Array("info1", "Pago : " & Me.dtpDesde.value & " - " & Me.dtpHasta.value))
        
    Set clientes = Nothing
    Set periodos = Nothing
    
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.dtpDesde.value = Date
    Me.dtpHasta.value = Date
    
End Sub
