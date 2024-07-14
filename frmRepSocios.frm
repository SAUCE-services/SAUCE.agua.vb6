VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socios"
   ClientHeight    =   810
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4095
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
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmRepSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim listado As New clsMyAListado
Dim cliente As New clsMODCliente
Dim operador As New clsMyAOperador
Dim medidor As New clsMyAMedidor

Dim clienteRep As New clsREPCliente

Dim cuit As String
Dim mens As String

Dim impresionService As New clsCtlImpresion

    Me.MousePointer = 11
    
    listado.truncate dbapp
    
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
    
    For Each cliente In clienteRep.collectionSociosActivos
        Set listado = New clsMyAListado
        
        listado.n1 = cliente.clienteId
        listado.c1 = cliente.apellido
        listado.c2 = cliente.nombre
        listado.c3 = cliente.inmuebleCalle
        listado.c4 = Left(cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto, 25)
        listado.n3 = cliente.numeroSocio
        Select Case cliente.cobro
            Case 1:
                listado.c5 = "Servicio Medido"
            Case 2:
                listado.c5 = "Cuota Fija en Trans."
            Case 3:
                listado.c5 = "Cuota Fija"
        End Select
        If cliente.cobro < 3 Then
            medidor.clienteId = cliente.clienteId
            medidor.findByClienteId dbapp
            If medidor.autoID > 0 Then listado.c6 = medidor.medidorID
        End If
        listado.add dbapp
    Next
    
    impresionService.printReport Me.crpFacturas, "rptSocios", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Socios (General)"))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

