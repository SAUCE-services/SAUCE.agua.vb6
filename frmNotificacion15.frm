VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotificacion15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificación 15 días"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11070
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1575
   End
   Begin Crystal.CrystalReport crpNotificacion 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdInvertir 
      Caption         =   "Invertir"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      ToolTipText     =   "Fin de la TAREA"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Fin de la TAREA"
      Top             =   4920
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   43281
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   4920
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6800
      _Version        =   393216
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Clientes"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Notificación"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmNotificacion15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private notificados As Collection

Private Sub fillClientes()
Dim cliente As clsMODCliente
Dim clientenotif As New clsMyAClienteNotif

Dim notificacion_service As New clsCtlNotificacion

Dim clientenotifs As Collection

    Set notificados = Nothing
    Set notificados = New Collection
    Set clientenotifs = clientenotif.collectionAll(dbapp)

    Me.grdClientes.Rows = 1
    Me.grdClientes.Redraw = False
    For Each cliente In notificacion_service.makeCollectionNotificacion15
        Set clientenotif = New clsMyAClienteNotif
        If modCollection.collectionExistElement(clientenotifs, "k." & cliente.clienteId) Then Set clientenotif = clientenotifs("k." & cliente.clienteId)
        Me.grdClientes.AddItem modGrid.array2itemGrid(Array(cliente.clienteId, cliente.apellidonombre, cliente.zona, cliente.ruta, cliente.orden, clientenotif.ultimaNotificacion15, clientenotif.ultimaNotificacion48, clientenotif.ultimaNotificacionCorte))
        Me.grdClientes.RowData(Me.grdClientes.Rows - 1) = cliente.clienteId
        modGrid.letCheckCell Me.grdClientes, Me.grdClientes.Rows - 1, 8, False
    Next
    Me.grdClientes.Redraw = True

End Sub

Private Sub cmdGenerar_Click()
Dim notificacion_service As New clsCtlNotificacion

    If notificados.Count = 0 Then
        MsgBox "ERROR: Nada para GENERAR"
        Exit Sub
    End If
    
    Me.cmdGenerar.Enabled = False
    Me.MousePointer = 11
    
    If notificacion_service.makeNotificacion15(Me.dtpFecha.value, notificados, dbapp) Then MsgBox "Notificaciones GENERADAS"
    
    Me.MousePointer = 0
    Me.cmdGenerar.Enabled = True
    
    fillClientes
    
End Sub

Private Sub cmdImprimir_Click()
Dim operador As New clsMyAOperador

Dim impresionService As New clsCtlImpresion
Dim notificacion_service As New clsCtlNotificacion

Dim cuit As String
Dim mens As String

    Me.cmdImprimir.Enabled = False
    
    Me.MousePointer = 11

    If Not notificacion_service.updateInteresesByFecha(Me.dtpFecha.value, dbapp) Then
        Me.MousePointer = 0
        Me.cmdImprimir.Enabled = True
        Exit Sub
    End If

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
    
    impresionService.printReport Me.crpNotificacion, "rptNotificacion15", dbapp.stringConnection, , Array(Array("fecha", toReportDate(Me.dtpFecha.value))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas))
        
    Me.MousePointer = 0
    
    Me.cmdImprimir.Enabled = True
        
End Sub

Private Sub cmdInvertir_Click()
Dim ciclo As Integer

    If Me.grdClientes.Rows = 1 Then Exit Sub
    
    For ciclo = 1 To Me.grdClientes.Rows - 1
        Me.grdClientes.col = 8
        Me.grdClientes.row = ciclo
        grdClientes_Click
    Next ciclo
    
End Sub

Private Sub cmdRevisar_Click()

    Me.cmdRevisar.Enabled = False
    
    Me.MousePointer = 11

    fillClientes
    
    Me.MousePointer = 0
    
    Me.cmdRevisar.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpFecha_Change()

    Me.grdClientes.Rows = 1
    
End Sub

Private Sub Form_Load()

    modGrid.makeGrid Me.grdClientes, Array(Array("Cliente", 600), Array("Apellido, Nombre", 4500), Array("Zona", 600), Array("Ruta", 600), Array("Orden", 600), Array("Ultima Not15", 1000), Array("Ultima AC", 1000), Array("Ultimo Corte", 1000), Array("", 300)), 0, 1, flexSelectionFree

    Me.dtpFecha.value = Date
    
End Sub

Private Sub grdClientes_Click()
Dim clientenotif As New clsMyAClienteNotif

    If Me.grdClientes.row < 1 Then Exit Sub
    If Me.grdClientes.col <> 8 Then Exit Sub
    
    modGrid.letCheckCell Me.grdClientes, Me.grdClientes.row, 8, Not modGrid.getCheckCell(Me.grdClientes, Me.grdClientes.row, 8)
    
    If Not modGrid.getCheckCell(Me.grdClientes, Me.grdClientes.row, 8) Then
        notificados.Remove "k." & Me.grdClientes.RowData(Me.grdClientes.row)
    Else
        clientenotif.clienteId = Me.grdClientes.RowData(Me.grdClientes.row)
        clientenotif.findByPrimaryKey dbapp
        
        clientenotif.ultimaNotificacion15 = Me.dtpFecha.value
        
        notificados.add clientenotif, "k." & Me.grdClientes.RowData(Me.grdClientes.row)
    End If
    
End Sub

