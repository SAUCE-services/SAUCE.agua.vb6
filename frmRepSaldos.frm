VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos por Conexión"
   ClientHeight    =   1410
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4110
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
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpReferencia 
      Height          =   375
      Left            =   240
      TabIndex        =   2
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
      Caption         =   "Referencia"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmRepSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim operador As New clsMyAOperador
Dim cliente As New clsMODCliente
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim cuota As New clsMyACuota
Dim listado As New clsMyAListado

Dim clienteRep As New clsREPCliente

Dim cuit As String
Dim mens As String

Dim t1 As Currency
Dim t2 As Currency
Dim t3 As Currency
Dim tt As Currency

Dim impresionService As New clsCtlImpresion

Dim periodos As Collection

    Me.MousePointer = 11
    
    listado.truncate dbapp
    
    operador.findLast dbapp
    
    Set periodos = periodo.collectionAll
    
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
    
    For Each cliente In clienteRep.collectionActivos(True)
        t1 = 0
        t2 = 0
        t3 = 0
        tt = 0
        For Each factura In factura.collectionDeudaDiferidaByClienteID(cliente.clienteId, Me.dtpReferencia.value, dbapp)
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & factura.periodoId) Then Set periodo = periodos("k." & factura.periodoId)
            t1 = t1 + factura.total
            t2 = t2 + interes(factura.total, factura.tasa, periodo.fechaPrimero, Me.dtpReferencia.value)
            tt = tt + factura.total + interes(factura.total, factura.tasa, periodo.fechaPrimero, Me.dtpReferencia.value)
        Next
        
        cuota.clienteId = cliente.clienteId
        cuota.findLastByClienteId dbapp
        For Each cuota In cuota.collectionDeudaByPlanID(cuota.clienteId, cuota.planID, dbapp, Me.dtpReferencia.value)
            t3 = t3 + cuota.importe
            tt = tt + cuota.importe
        Next
        If tt > 0 Then
            Set listado = New clsMyAListado
            listado.n1 = cliente.clienteId
            listado.c1 = Left(cliente.apellidonombre, 25)
            listado.n2 = tt
            listado.n3 = t1
            listado.n4 = t2
            listado.n5 = t3
            listado.add dbapp
        End If
    Next
 
    impresionService.printReport Me.crpFacturas, "rptSaldos", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Saldos por Conexión"), _
        Array("info1", "Deuda al : " & Me.dtpReferencia.value))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.dtpReferencia.value = Date
    
End Sub
