VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepVolFactu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volumen Facturado por Período"
   ClientHeight    =   1350
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   6030
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   3615
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
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmRepVolFactu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clientes As Collection
Private periodos As Collection

Private Sub cmdImprimir_Click()
Dim listado As New clsMyAListado
Dim cliente As New clsMODCliente
Dim operador As New clsMyAOperador
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura
Dim medidor As New clsMyAMedidor
Dim lectura As New clsMyALectura
Dim clientevol As clsMyAClienteVolumen

Dim cuit As String
Dim Mensaje As String
Dim medidorID As String

Dim fechaActual As Variant
Dim fechaAnterior As Variant

Dim estadoActual As Currency
Dim estadoAnterior As Currency
Dim consumoRegistrado As Long

Dim impresionService As New clsCtlImpresion

    If Me.cboPeriodo.ListIndex < 0 Then Exit Sub

    Me.MousePointer = 11
    
    listado.truncate dbapp
    
    operador.findLast dbapp
    
    cuit = Left(operador.cuit, 2) & "-" & Mid(operador.cuit, 3, 8) & "-" & Right(operador.cuit, 1)
    Select Case operador.situacionIVA
        Case 1:
            Mensaje = "Resp. Inscripto"
        Case 2:
            Mensaje = "Resp. No Inscripto"
        Case 3:
            Mensaje = "Cons. Final"
        Case 4:
            Mensaje = "Exento"
        Case 5:
            Mensaje = "No Responsable"
    End Select

    For Each factura In factura.collectionByPeriodoId_(Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex), dbapp)
        DoEvents
        Set cliente = New clsMODCliente
        If modCollection.collectionExistElement(clientes, "k." & factura.clienteId) Then Set cliente = clientes("k." & factura.clienteId)
        Set periodo = New clsRESTPeriodo
        If modCollection.collectionExistElement(periodos, "k." & factura.periodoId) Then Set periodo = periodos("k." & factura.periodoId)
        
        Set listado = New clsMyAListado
        Set clientevol = New clsMyAClienteVolumen
        
        listado.n1 = factura.clienteId
        listado.c1 = Left(cliente.apellidonombre, 25)
        medidorID = 0
        medidor.clienteId = factura.clienteId
        medidor.findColocadoByClienteID dbapp
        If medidor.autoID > 0 Then medidorID = medidor.medidorID
        medidor.medidorID = medidorID
        medidor.findLast dbapp
        
        lectura.medidorID = medidor.medidorID
        lectura.periodoId = periodo.periodoId
        lectura.findByPrimaryKey dbapp
        
        If lectura.autoID = 0 Then
            fechaActual = Null
            estadoActual = medidor.estadoInicio
        Else
            fechaActual = lectura.fechaLectura
            estadoActual = medidor.estadoInicio
            If medidor.fechaColocacion <= periodo.fechaInicio Then estadoActual = lectura.estado
        End If
        
        clientevol.clienteId = factura.clienteId
        clientevol.periodoId = factura.periodoId
        clientevol.medidorIDActual = medidor.medidorID
        clientevol.estadoActual = estadoActual
        
        lectura.medidorID = medidor.medidorID
        lectura.periodoId = periodo.periodoId
        lectura.findByMedidorIDPrev dbapp
        If lectura.autoID = 0 Then
            fechaAnterior = Null
            estadoAnterior = 0
            If medidor.autoID > 0 Then estadoAnterior = medidor.estadoInicio
        Else
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & lectura.periodoId) Then Set periodo = periodos("k." & lectura.periodoId)
            fechaAnterior = lectura.fechaLectura
            estadoAnterior = medidor.estadoInicio
            If medidor.fechaColocacion <= periodo.fechaFin Then estadoAnterior = lectura.estado
        End If
        
        clientevol.medidorIDAnterior = medidor.medidorID
        clientevol.estadoAnterior = estadoAnterior
        
        consumoRegistrado = estadoActual - estadoAnterior
        If consumoRegistrado < 0 Then consumoRegistrado = 0
        
        clientevol.consumido = consumoRegistrado
        clientevol.save dbapp
        
        listado.n2 = consumoRegistrado
        If consumoRegistrado > 0 Then listado.add dbapp
    Next
    
    Set periodo = periodos("k." & Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex))
    
    impresionService.printReport Me.crpFacturas, "rptVolumen", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & Mensaje & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("info1", "Período : " & periodo.descripcion), _
        Array("titulo", "Volumen Facturado por Período"))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
Dim clienteRep As clsREPCliente
Dim periodo As New clsRESTPeriodo

    Set clienteRep = New clsREPCliente
    Set clientes = clienteRep.collectionActivos
    Set clienteRep = Nothing
    Set periodos = periodo.collectionAll

End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    periodo.fillCombo Me.cboPeriodo
    
End Sub
