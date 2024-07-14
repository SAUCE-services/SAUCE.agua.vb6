VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepRestPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restricciones por Per�odo"
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
      Caption         =   "Per�odo"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmRepRestPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim listado As New clsMyAListado
Dim cliente As New clsMODCliente
Dim operador As New clsMyAOperador
Dim suspension As New clsMyASuspension
Dim periodo As New clsRESTPeriodo

Dim clienterep As clsREPCliente

Dim cuit As String
Dim mens As String

Dim ctlImp As New clsCtlImpresion

Dim clientes As Collection

    If Me.cboPeriodo.ListIndex < 0 Then Exit Sub

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

    Set clienterep = New clsREPCliente
    Set clientes = clienterep.collectionActivos
    Set clienterep = Nothing
    
    periodo.periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
    periodo.findByPrimaryKey

    For Each suspension In suspension.collectionByPeriodoID("R", Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex), dbapp)
        Set cliente = New clsMODCliente
        If modCollection.collectionExistElement(clientes, "k." & suspension.clienteId) Then Set cliente = clientes("k." & suspension.clienteId)
        
        Set listado = New clsMyAListado
        
        listado.n1 = suspension.numero
        listado.n2 = suspension.clienteId
        listado.c1 = cliente.apellido
        listado.c2 = cliente.nombre
        listado.c3 = suspension.fecha
        listado.add dbapp
    Next
    
    ctlImp.printReport Me.crpFacturas, "rptSuspen", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("info1", "Per�odo : " & periodo.descripcion), _
        Array("titulo", "Restricciones por Per�odo"))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    periodo.fillCombo Me.cboPeriodo
    
End Sub