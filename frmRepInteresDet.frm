VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepInteresDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Intereses"
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
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin Crystal.CrystalReport crpFacturas 
      Left            =   1680
      Top             =   840
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
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmRepInteresDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub calculaCliente()
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

Dim periodos As Collection

Dim pagofacil_service As New clsCtlPagoFacil
Dim liquidacion_service As New clsCtlLiquidacion

    Set periodos = periodo.collectionAll
    
    For Each factura In factura.collectionInteresByClienteID(cliente.clienteId, dbapp)
        If Not IsNull(factura.fechapago) Then
            Set periodo = periodos("k." & factura.periodoId)
            factura.interes = modInteres.interes(factura.total, factura.tasa, periodo.fechaPrimero, factura.fechapago)
            factura.pfcodigo = pagofacil_service.codigopf(liquidacion_service.oldFactura2newFactura(factura))
            factura.save dbapp
        End If
    Next
    
End Sub

Private Sub cmdImprimir_Click()
Dim operador As New clsMyAOperador

Dim cuit As String
Dim mens As String

Dim impresionService As New clsCtlImpresion

    Me.MousePointer = 11
    
    calculaCliente
    
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
    
    impresionService.printReport Me.crpFacturas, "rptInteresDet", dbapp.stringConnection, , _
        Array(Array("cliente_id", cliente.clienteId)), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
End Sub



