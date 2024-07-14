VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   8070
   ClientLeft      =   2085
   ClientTop       =   2805
   ClientWidth     =   9840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9840
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   4200
      TabIndex        =   47
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Enviar por Correo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      ToolTipText     =   "Imprime la FACTURA Actual"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimirDigital 
      Caption         =   "Imprimir Digital"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      ToolTipText     =   "Imprime la FACTURA Actual"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   720
      Width           =   7095
   End
   Begin VB.ComboBox pdef 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox desc 
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton fin 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton impri 
      Caption         =   "&Imprimir Factura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      ToolTipText     =   "Imprime la FACTURA Actual"
      Top             =   5760
      Width           =   1575
   End
   Begin Crystal.CrystalReport crpLiquidacion 
      Left            =   480
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:factura.rpt"
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton anul 
      Caption         =   "&Anular Factura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "Permite ANULAR la Factura actual"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox efec 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton factu 
      Caption         =   "&Facturar e Imprimir"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      ToolTipText     =   "Factura e Imprime los DATOS de la CONEXION actual"
      Top             =   5160
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid rfac 
      Height          =   1575
      Left            =   480
      TabIndex        =   29
      Top             =   3480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
   End
   Begin VB.Frame dmed 
      Caption         =   "Datos del MEDIDOR"
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   9375
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Consumo"
         Height          =   195
         Index           =   11
         Left            =   7560
         TabIndex        =   27
         Top             =   360
         Width           =   660
      End
      Begin VB.Label creg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7560
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "REGISTRADO (m³)"
         Height          =   195
         Index           =   10
         Left            =   7560
         TabIndex        =   25
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Index           =   9
         Left            =   5760
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Label fant 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label fact 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de LECTURA"
         Height          =   195
         Index           =   8
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label eant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label eact 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Medición ANTERIOR"
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Medición ACTUAL"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label nmed 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.PictureBox consumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   7515
      TabIndex        =   41
      Top             =   7920
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "e-mail"
      Height          =   195
      Index           =   2
      Left            =   4200
      TabIndex        =   48
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label mcli 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   480
      TabIndex        =   45
      Top             =   1320
      Width           =   8895
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio"
      Height          =   195
      Index           =   21
      Left            =   480
      TabIndex        =   44
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   4
      Left            =   7800
      TabIndex        =   13
      Top             =   480
      Width           =   570
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Cons. Final"
      Height          =   195
      Index           =   19
      Left            =   2400
      TabIndex        =   43
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Label tcfl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   42
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Descuentos o Bonificaciones"
      Height          =   195
      Index           =   18
      Left            =   3720
      TabIndex        =   40
      Top             =   5160
      Width           =   2070
   End
   Begin VB.Label tseg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   39
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label tpri 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "a Pagar - 2do Vto"
      Height          =   195
      Index           =   17
      Left            =   6000
      TabIndex        =   37
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "a Pagar - 1er Vto"
      Height          =   195
      Index           =   16
      Left            =   4200
      TabIndex        =   36
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Resp. No Insc."
      Height          =   195
      Index           =   15
      Left            =   6000
      TabIndex        =   35
      Top             =   5520
      Width           =   1365
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA Resp. Inscripto"
      Height          =   195
      Index           =   14
      Left            =   4200
      TabIndex        =   34
      Top             =   5520
      Width           =   1365
   End
   Begin VB.Label trni 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label tiva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label tsub 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "SubTotal"
      Height          =   195
      Index           =   13
      Left            =   600
      TabIndex        =   30
      Top             =   5520
      Width           =   645
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Rubros FACTURADOS"
      Height          =   195
      Index           =   12
      Left            =   480
      TabIndex        =   28
      Top             =   3240
      Width           =   1650
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   480
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label nrofac 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de Factura"
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prefijoId As Integer
Private facturaId As Long
Private lrub As Integer
Private rubroID As Integer
Private opserv, opper As Integer
Private cons, aCF, aIVA, aRNI As Currency
Private opini As Date
Private usuarioDesconectado As Boolean

Private loading As Boolean

Private cliente As New clsMODCliente

Public Sub buscar_fac()
Dim fini, ffin As Date
Dim regi As Variant
Dim exi, exs As Integer

Dim operador As New clsMyAOperador
Dim suspfactura As New clsMyASuspFactura
Dim periodo As New clsRESTPeriodo
Dim medidor As New clsMyAMedidor
Dim desconexion As New clsMyADesconexion
Dim factura As New clsMyAFactura

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If pdef.ListIndex < 0 Then Exit Sub
    
    loading = True
    
    opserv = 3
    opini = CDate("01/02/1998")
    opper = 1
    
    operador.findLast dbapp
    opserv = operador.servicio
    opper = operador.periodoFactura
    If Not IsNull(operador.fechaInicio) Then opini = operador.fechaInicio
    
    suspfactura.clienteId = cliente.clienteId
    suspfactura.periodoIDInicio = Me.pdef.ItemData(Me.pdef.ListIndex)
    suspfactura.findSuspendidaByClienteID dbapp
    If suspfactura.autoID > 0 Then
        If IsNull(suspfactura.periodoIdfin) Or suspfactura.periodoIdfin >= Me.pdef.ItemData(Me.pdef.ListIndex) Then
            MsgBox "Este Cliente tiene la Facturación SUSPENDIDA . . ."
            Exit Sub
        End If
    End If
    periodo.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    periodo.findByPrimaryKey
    fini = periodo.fechaInicio
    ffin = periodo.fechaFin
    medidor.clienteId = cliente.clienteId
    medidor.findByClienteId dbapp
    factu.Enabled = False
    desc.Enabled = False
    If medidor.autoID = 0 Then
        Set cliente = clienteRep.findLastByClienteId(cliente.clienteId)
        If cliente.cobro < 3 Then
            MsgBox "No tiene ningún MEDIDOR asociado"
            anul.Enabled = False
            impri.Enabled = False
            Me.cmdImprimirDigital.Enabled = False
            Me.cmdSendMail.Enabled = False
            nmed.Caption = 0
            Exit Sub
        End If
    End If
    
    desconexion.clienteId = cliente.clienteId
    desconexion.fechaDesconexion = fini
    desconexion.findDesconectado dbapp
    usuarioDesconectado = False
    If desconexion.autoID > 0 Then
        If IsNull(desconexion.fechaReconexion) Then usuarioDesconectado = True
        If desconexion.fechaReconexion > ffin Then If desconexion.fechaDesconexion <= fini Then usuarioDesconectado = True
    End If
    
    If usuarioDesconectado Then MsgBox "CLIENTE desconectado para este período"
    
    If factura.collectionAny(dbapp).Count > 0 Then
        factura.clienteId = cliente.clienteId
        factura.periodoId = pdef.ItemData(pdef.ListIndex)
        factura.findByClientePeriodo dbapp
        If factura.autoID = 0 Then
' Busca alguna factura de períodos múltiples
            factura.clienteId = cliente.clienteId
            factura.periodoId = pdef.ItemData(pdef.ListIndex)
            factura.findByClientePeriodoPrev dbapp
            If factura.autoID = 0 Then
                factu.Enabled = True
                anul.Enabled = False
                impri.Enabled = False
                Me.cmdImprimirDigital.Enabled = False
                Me.cmdSendMail.Enabled = False
                desc.Enabled = True
                desc.Text = Format(0, "#,###,##0.00")
                llenar usuarioDesconectado
            Else
                If factura.periodoIdfin = 0 Or IsNull(factura.periodoIdfin) Or factura.periodoIdfin < pdef.ItemData(pdef.ListIndex) Then
                    factu.Enabled = True
                    anul.Enabled = False
                    impri.Enabled = False
                    Me.cmdImprimirDigital.Enabled = False
                    Me.cmdSendMail.Enabled = False
                    desc.Enabled = True
                    desc.Text = Format(0, "#,###,##0.00")
                    llenar usuarioDesconectado
                Else
                    factu.Enabled = False
                    anul.Enabled = True
                    impri.Enabled = True
                    Me.cmdImprimirDigital.Enabled = True
                    Me.cmdSendMail.Enabled = True
                    desc.Enabled = False
                    If factura.pagada Or factura.cancelada Then anul.Enabled = False
                    traer usuarioDesconectado
                End If
            End If
        Else
            If factura.anulada Then
                factu.Enabled = True
                anul.Enabled = False
                impri.Enabled = False
                Me.cmdImprimirDigital.Enabled = False
                Me.cmdSendMail.Enabled = False
                desc.Enabled = True
                desc.Text = Format(0, "#,###,##0.00")
                llenar (usuarioDesconectado)
            Else
                factu.Enabled = False
                anul.Enabled = True
                impri.Enabled = True
                Me.cmdImprimirDigital.Enabled = True
                Me.cmdSendMail.Enabled = True
                desc.Enabled = False
                If factura.pagada Or factura.cancelada Then anul.Enabled = False
                traer usuarioDesconectado
            End If
        End If
    Else
        factu.Enabled = True
        anul.Enabled = False
        impri.Enabled = False
        Me.cmdImprimirDigital.Enabled = False
        Me.cmdSendMail.Enabled = False
        desc.Enabled = True
        desc.Text = Format(0, "#,###,##0.00")
        llenar usuarioDesconectado
    End If
    
    loading = False
    
End Sub

Public Sub llenar(ByVal usuarioDesconectado As Boolean)
Dim conc As String
Dim servicio As String
Dim segmento As String
Dim critu, critc, crith As String
Dim cobro As Integer
Dim situacionIVA As Integer
Dim categoriaID As Integer
Dim serv As Integer
Dim cantidadPeriodos As Integer
Dim i, li As Integer
Dim importe As Currency
Dim cant As Currency
Dim metrosConsumo As Currency
Dim factorCobro As Currency
Dim cat, ct As Currency
Dim listar As Boolean
Dim subt, ivct, ivat, rnit
Dim tasa As Double
Dim ted, tatr As Currency
Dim tasaMayor, tasaMenor, ides As Currency
Dim fec As Date

Dim operador As New clsMyAOperador
Dim factura As New clsMyAFactura
Dim medidor As New clsMyAMedidor
Dim periodo As New clsRESTPeriodo
Dim objMLec As New clsMyALectura
Dim objMRub As New clsMyARubro
Dim novedad As New clsMyANovedad
Dim objMRan As New clsMyARango

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    cantidadPeriodos = 1
    desc.Enabled = True
    prefijoId = 1
    facturaId = 1
    
    operador.findLast dbapp
    If Not IsNull(operador.puntoVta) Then prefijoId = operador.puntoVta
    If Not IsNull(operador.nroComprob) Then facturaId = operador.nroComprob
    
    If factura.collectionAny(dbapp).Count > 0 Then
        factura.findLastLast dbapp
        If prefijoId <= factura.puntoVta Then
            prefijoId = factura.puntoVta
            factura.puntoVta = prefijoId
            factura.findLast dbapp
            If facturaId < factura.nroComprob + 1 Then facturaId = factura.nroComprob + 1
        End If
    End If
    efec.Text = Date
    tasaMayor = 75
    tasaMenor = 25
    If CDate(efec.Text) - opini > 365 Then
        tasaMayor = 50
        tasaMenor = 50
    End If
    If CDate(efec.Text) - opini > 730 Then
        tasaMayor = 25
        tasaMenor = 75
    End If
    If CDate(efec.Text) - opini > 1095 Then
        tasaMayor = 0
        tasaMenor = 100
    End If
    nrofac.Caption = Right("0000" & prefijoId, 4) & "-" & Right("00000000" & facturaId, 8)
    subt = 0
    ivct = 0
    ivat = 0
    rnit = 0
    Set cliente = clienteRep.findLastByClienteId(cliente.clienteId)
    cobro = cliente.cobro
    dmed.Enabled = True
    If cobro = 3 Then dmed.Enabled = False
    situacionIVA = cliente.situacionIVA
    categoriaID = cliente.categoria
    serv = cliente.servicio
    factorCobro = 1
    If serv < opserv Then factorCobro = 0.5
    
    medidor.clienteId = cliente.clienteId
    medidor.findColocadoByClienteID dbapp
    nmed.Caption = 0
    If medidor.autoID > 0 Then nmed.Caption = medidor.medidorID
    critu = "IDMedidor = '" & Trim(nmed.Caption) & "'"
    If Not usuarioDesconectado Then
        periodo.periodoId = pdef.ItemData(pdef.ListIndex)
        periodo.findByPrimaryKey
        fec = periodo.fechaInicio
        objMLec.medidorID = Trim(nmed.Caption)
        objMLec.periodoId = pdef.ItemData(pdef.ListIndex)
        objMLec.findByPrimaryKey dbapp
        
        medidor.medidorID = Trim(nmed.Caption)
        medidor.findByMedidorID dbapp
        If objMLec.autoID = 0 Then
            fact.Caption = ""
            eact.Caption = 0
            If medidor.autoID > 0 Then eact.Caption = medidor.estadoInicio
        Else
            fact.Caption = objMLec.fechaLectura
            eact.Caption = medidor.estadoInicio
            If medidor.fechaColocacion <= fec Then eact.Caption = objMLec.estado
        End If
        
        objMLec.medidorID = Trim(nmed.Caption)
        objMLec.periodoId = pdef.ItemData(pdef.ListIndex)
        objMLec.findByMedidorIDPrev dbapp
        If objMLec.autoID = 0 Then
            fant.Caption = ""
            eant.Caption = 0
            If medidor.autoID > 0 Then eant.Caption = medidor.estadoInicio
        Else
            periodo.periodoId = objMLec.periodoId
            periodo.findByPrimaryKey
            fec = periodo.fechaFin
            fant.Caption = objMLec.fechaLectura
            eant.Caption = medidor.estadoInicio
            If medidor.fechaColocacion <= fec Then eant.Caption = objMLec.estado
        End If
        creg.Caption = CDbl(eact.Caption) - CDbl(eant.Caption)
    Else
        fact.Caption = ""
        fant.Caption = ""
        eact.Caption = ""
        eant.Caption = ""
        creg.Caption = 0
    End If

    rfac.Rows = 1
    rubroID = 0
    lrub = 0
    For Each objMRub In objMRub.collectionSinRepeticion(dbapp)
        If rubroID < objMRub.rubroID Then
            rubroID = objMRub.rubroID
            objMRub.findLast dbapp
            listar = False
            If usuarioDesconectado Then
                If cobro >= objMRub.cobro And objMRub.desconectado Then
                    listar = True
                    importe = objMRub.precioUnitario * factorCobro
                    servicio = ""
                    segmento = ""
                    
                    novedad.clienteId = cliente.clienteId
                    novedad.periodoId = pdef.ItemData(pdef.ListIndex)
                    novedad.rubroID = rubroID
                    novedad.findRango dbapp
                    If novedad.autoID = 0 Then
    'Analiza los rangos aplicados al consumo
                        If objMRub.rangoID > 0 Then
                            objMRan.categoria = categoriaID
                            objMRan.rangoID = objMRub.rangoID
                            objMRan.findLast dbapp
                            If objMRan.autoID = 0 Then
                                listar = False
                            Else
                                importe = objMRan.tarifa * factorCobro
                                If CDbl(creg.Caption) > objMRan.limiteSuperior Then metrosConsumo = objMRan.limiteSuperior - objMRan.limiteInferior
                                If CDbl(creg.Caption) <= objMRan.limiteSuperior Then metrosConsumo = CDbl(creg.Caption) - objMRan.limiteInferior
                                If CDbl(creg.Caption) <= objMRan.limiteInferior Then
                                    metrosConsumo = 0
                                    listar = False
                                End If
                                If objMRan.limiteSuperior > 99998 Then
                                    segmento = " ( más de " & objMRan.limiteInferior & " m³)"
                                Else
                                    segmento = " (" & objMRan.limiteInferior & "-" & objMRan.limiteSuperior & " m³)"
                                End If
                                servicio = " Servicio"
                                Select Case serv
                                    Case 1
                                        servicio = servicio & " Agua"
                                    Case 2
                                        servicio = servicio & " Cloaca"
                                    Case 3
                                        servicio = servicio & " Agua y Cloaca"
                                End Select
                                servicio = servicio & " Sistema Medido"
                            End If
                        End If
                    Else
                        listar = False
                    End If
                End If
            Else
                If cobro >= objMRub.cobro And (objMRub.comun Or (objMRub.comunSocio And Val(cliente.numeroSocio) > 0)) Then
                    listar = True
                    importe = objMRub.precioUnitario * factorCobro
                    servicio = ""
                    segmento = ""
                    
                    novedad.clienteId = cliente.clienteId
                    novedad.periodoId = pdef.ItemData(pdef.ListIndex)
                    novedad.rubroID = rubroID
                    novedad.findRango dbapp
                    If novedad.autoID = 0 Then
    'Analiza los rangos aplicados al consumo
                        If objMRub.rangoID > 0 Then
                            objMRan.categoria = categoriaID
                            objMRan.rangoID = objMRub.rangoID
                            objMRan.findLast dbapp
                            If objMRan.autoID = 0 Then
                                listar = False
                            Else
                                importe = objMRan.tarifa * factorCobro
                                If CDbl(creg.Caption) > objMRan.limiteSuperior Then metrosConsumo = objMRan.limiteSuperior - objMRan.limiteInferior
                                If CDbl(creg.Caption) <= objMRan.limiteSuperior Then metrosConsumo = CDbl(creg.Caption) - objMRan.limiteInferior
                                If CDbl(creg.Caption) <= objMRan.limiteInferior Then
                                    metrosConsumo = 0
                                    listar = False
                                End If
                                If objMRan.limiteSuperior > 99998 Then
                                    segmento = " ( más de " & objMRan.limiteInferior & " m³)"
                                Else
                                    segmento = " (" & objMRan.limiteInferior & "-" & objMRan.limiteSuperior & " m³)"
                                End If
                                servicio = " Servicio"
                                Select Case serv
                                    Case 1
                                        servicio = servicio & " Agua"
                                    Case 2
                                        servicio = servicio & " Cloaca"
                                    Case 3
                                        servicio = servicio & " Agua y Cloaca"
                                End Select
                                servicio = servicio & " Sistema Medido"
                            End If
                        End If
                    Else
                        listar = False
                    End If
                End If
            End If
        
            If cobro = 3 And objMRub.cobro = 1 Then listar = False
            
            If listar Then
                lrub = lrub + 1
                rfac.Rows = lrub + 1
                rfac.row = lrub
                rfac.col = 0
                rfac.Text = Right("00" & objMRub.rubroID, 2)
                rfac.col = 1
                conc = objMRub.concepto & servicio & segmento
                If cobro = 2 And objMRub.cobro = 2 Then conc = conc & " (" & tasaMayor & "%)"
                If cobro = 2 And objMRub.cobro = 1 Then conc = conc & " (" & tasaMenor & "%)"
                rfac.Text = conc
                rfac.col = 2
                cant = cantidadPeriodos
                If metrosConsumo > 0 Then
                    cant = metrosConsumo
                    If cobro = 2 And objMRub.cobro = 1 Then cant = metrosConsumo * tasaMenor / 100
                End If
                rfac.Text = cant
                rfac.col = 3
    'Controla si el rubro corresponde a un cargo fijo en transicion
                If objMRub.cobro = 2 And cobro < 3 Then importe = importe * tasaMayor / 100
                If objMRub.cobro = 1 And cobro = 2 And objMRub.rangoID = 0 Then importe = importe * tasaMenor / 100
                rfac.Text = Format(importe, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(importe * cant, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + importe * cant * aIVA
                        Case 2:
                            ivat = ivat + importe * cant * aIVA
                            rnit = rnit + importe * cant * aRNI
                        Case 3 To 5:
                            ivct = ivct + importe * cant * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + importe * cant
                metrosConsumo = 0
            End If
        End If
    Next
    
    If Not usuarioDesconectado Then
'Novedades únicas
        For Each novedad In novedad.collectionUnicasByClienteID(cliente.clienteId, pdef.ItemData(pdef.ListIndex), dbapp)
            rubroID = novedad.rubroID
            ct = 0
            For i = 0 To lrub
                rfac.col = 0
                rfac.row = i
                If Val(rfac.Text) = rubroID Then
                    li = rfac.row
                    rfac.col = 2
                    ct = ct + CDbl(rfac.Text)
                End If
            Next i
            If ct = 0 Then
                lrub = lrub + 1
                rfac.Rows = lrub + 1
                rfac.row = lrub
                rfac.col = 0
                rfac.Text = Right("00" & novedad.rubroID, 2)
                objMRub.rubroID = novedad.rubroID
                objMRub.findLast dbapp
                rfac.col = 1
                rfac.Text = objMRub.concepto
                rfac.col = 2
                cant = novedad.cantidad
                If cant = 0 Then cant = novedad.porcentaje
                rfac.Text = cant
                rfac.col = 3
                rfac.Text = Format(objMRub.precioUnitario, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(objMRub.precioUnitario * cant, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                        Case 2:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                            rnit = rnit + objMRub.precioUnitario * cant * aRNI
                        Case 3 To 5:
                            ivct = ivct + objMRub.precioUnitario * cant * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + objMRub.precioUnitario * cant
            Else
                rfac.row = li
                objMRub.rubroID = novedad.rubroID
                objMRub.findLast dbapp
                rfac.col = 2
                cant = novedad.cantidad
                If cant = 0 Then cant = novedad.porcentaje
                cat = CDbl(rfac.Text) + cant
                rfac.Text = cat
                rfac.col = 3
                rfac.Text = Format(objMRub.precioUnitario, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(objMRub.precioUnitario * cat, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                        Case 2:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                            rnit = rnit + objMRub.precioUnitario * cant * aRNI
                        Case 3 To 5:
                            ivct = ivct + objMRub.precioUnitario * cant * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + objMRub.precioUnitario * cant
            End If
        Next
'Novedades indefinidas
        For Each novedad In novedad.collectionIndefinidasByClienteID(cliente.clienteId, pdef.ItemData(pdef.ListIndex), dbapp)
            rubroID = novedad.rubroID
            ct = 0
            For i = 0 To lrub
                rfac.col = 0
                rfac.row = i
                If Val(rfac.Text) = rubroID Then
                    li = rfac.row
                    rfac.col = 2
                    ct = ct + CDbl(rfac.Text)
                End If
            Next i
            If ct = 0 Then
                lrub = lrub + 1
                rfac.Rows = lrub + 1
                rfac.row = lrub
                rfac.col = 0
                rfac.Text = Right("00" & novedad.rubroID, 2)
                objMRub.rubroID = novedad.rubroID
                objMRub.findLast dbapp
                rfac.col = 1
                rfac.Text = objMRub.concepto
                rfac.col = 2
                cant = novedad.cantidad
                If cant = 0 Then cant = novedad.porcentaje
                rfac.Text = cant
                rfac.col = 3
                rfac.Text = Format(objMRub.precioUnitario, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(objMRub.precioUnitario * cant, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                        Case 2:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                            rnit = rnit + objMRub.precioUnitario * cant * aRNI
                        Case 3 To 5:
                            ivct = ivct + objMRub.precioUnitario * cant * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + objMRub.precioUnitario * cant
            Else
                rfac.row = li
                objMRub.rubroID = novedad.rubroID
                objMRub.findLast dbapp
                rfac.col = 2
                cant = novedad.cantidad
                If cant = 0 Then cant = novedad.porcentaje
                cat = CDbl(rfac.Text) + cant
                rfac.Text = cat
                rfac.col = 3
                rfac.Text = Format(objMRub.precioUnitario, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(objMRub.precioUnitario * cat, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                        Case 2:
                            ivat = ivat + objMRub.precioUnitario * cant * aIVA
                            rnit = rnit + objMRub.precioUnitario * cant * aRNI
                        Case 3 To 5:
                            ivct = ivct + objMRub.precioUnitario * cant * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + objMRub.precioUnitario * cant
            End If
        Next
'Novedades en veces
        For Each novedad In novedad.collectionVecesByClienteID(cliente.clienteId, pdef.ItemData(pdef.ListIndex), dbapp)
            rubroID = novedad.rubroID
            ct = 0
            For i = 0 To lrub
                rfac.col = 0
                rfac.row = i
                If Val(rfac.Text) = rubroID Then
                    li = rfac.row
                    rfac.col = 2
                    ct = ct + CDbl(rfac.Text)
                End If
            Next i
            If ct = 0 Then
                lrub = lrub + 1
                rfac.Rows = lrub + 1
                rfac.row = lrub
                rfac.col = 0
                rfac.Text = Right("00" & novedad.rubroID, 2)
                objMRub.rubroID = novedad.rubroID
                objMRub.findLast dbapp
                rfac.col = 1
                rfac.Text = objMRub.concepto
                rfac.col = 2
                cant = novedad.cantidad
                If cant = 0 Then cant = novedad.porcentaje
                rfac.Text = cant
                rfac.col = 3
                rfac.Text = Format(novedad.importe / novedad.veces, "#,###,##0.00")
                rfac.col = 4
                rfac.Text = Format(novedad.importe * cant / novedad.veces, "#,###,##0.00")
                rfac.col = 5
                If objMRub.IVA Then
                    rfac.Text = "Si"
                    Select Case situacionIVA
                        Case 1, 6:
                            ivat = ivat + novedad.importe * cant / novedad.veces * aIVA
                        Case 2:
                            ivat = ivat + novedad.importe * cant / novedad.veces * aIVA
                            rnit = rnit + novedad.importe * cant / novedad.veces * aRNI
                        Case 3 To 5:
                            ivct = ivct + novedad.importe * cant / novedad.veces * aCF
                    End Select
                Else
                    rfac.Text = "No"
                End If
                subt = subt + novedad.importe * cant / novedad.veces
            Else
                If ct < novedad.veces - novedad.vecesCobradas Then
                    rfac.row = li
                    objMRub.rubroID = novedad.rubroID
                    objMRub.findLast dbapp
                    rfac.col = 2
                    cant = novedad.cantidad
                    If cant = 0 Then cant = novedad.porcentaje
                    rfac.Text = cant * (ct + 1)
                    rfac.col = 4
                    rfac.Text = Format(novedad.importe * cant * (ct + 1) / novedad.veces, "#,###,##0.00")
                    rfac.col = 5
                    If objMRub.IVA Then
                        rfac.Text = "Si"
                        Select Case situacionIVA
                            Case 1, 6:
                                ivat = ivat + novedad.importe * cant / novedad.veces * aIVA
                            Case 2:
                                ivat = ivat + novedad.importe * cant / novedad.veces * aIVA
                                rnit = rnit + novedad.importe * cant / novedad.veces * aRNI
                            Case 3 To 5:
                                ivct = ivct + novedad.importe * cant / novedad.veces * aCF
                        End Select
                    Else
                        rfac.Text = "No"
                    End If
                    subt = subt + novedad.importe * cant / novedad.veces
                End If
            End If
        Next
    End If
    tatr = 0
    For Each factura In factura.collectionParaInteresByClienteID(cliente.clienteId, dbapp)
        periodo.periodoId = factura.periodoId
        periodo.findByPrimaryKey
        If periodo.fechaSegundo < factura.fechapago Then tatr = tatr + interes(factura.total, periodo.tasa, periodo.fechaSegundo, factura.fechapago)
    Next
    If tatr > 0 Then
        rubroID = 0
        lrub = lrub + 1
        rfac.Rows = lrub + 1
        rfac.row = lrub
        rfac.col = 0
        rfac.Text = "00"
        rfac.col = 1
        rfac.Text = "Intereses por mora de Facturas pagadas fuera de término"
        rfac.col = 2
        cant = 1
        rfac.Text = cant
        rfac.col = 3
        rfac.Text = Format(tatr, "#,###,##0.00")
        rfac.col = 4
        rfac.Text = Format(tatr, "#,###,##0.00")
        rfac.col = 5
        rfac.Text = "Si"
        Select Case situacionIVA
            Case 1, 6:
                ivat = ivat + tatr * aIVA
            Case 2:
                ivat = ivat + tatr * aIVA
                rnit = rnit + tatr * aRNI
            Case 3 To 5:
                ivct = ivct + tatr * aCF
        End Select
        subt = subt + tatr
    End If
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey
    tasa = periodo.tasa
    ides = CDbl(desc.Text)
'    subt = subt - ides / (1 + aIVA)
    subt = subt - ides
'    Select Case situacionIVA
'        Case 1, 2, 6:
'            ivat = ivat - ides * (1 - 1 / (1 + aIVA))
'        Case 3 To 5:
'            ivct = ivct - ides * (1 - 1 / (1 + aIVA))
'    End Select
    tsub.Caption = Format(subt, "#,###,##0.00")
    tcfl.Caption = Format(ivct, "#,###,##0.00")
    tiva.Caption = Format(ivat, "#,###,##0.00")
    trni.Caption = Format(rnit, "#,###,##0.00")
    tpri.Caption = Format(subt + ivct + ivat + rnit, "#,###,##0.00")
    tseg.Caption = Format(subt + ivct + ivat + rnit + interes(subt + ivct + ivat + rnit, tasa, periodo.fechaPrimero, periodo.fechaSegundo), "#,###,##0.00")
    factu.Enabled = True
    If subt + ivct + ivat + rnit < 0.01 Then factu.Enabled = False

End Sub

Public Sub traer(usuarioDesconectado As Boolean)
Dim conc, part As String
Dim siva, cobr As Integer
Dim impo, cant, metr As Currency
Dim subt, ivct, ivat, rnit As Currency
Dim tasa As Double
Dim ted, ides As Currency

Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim medidor As New clsMyAMedidor
Dim objMLec As New clsMyALectura
Dim detalle As New clsMyADetalle

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    factura.clienteId = cliente.clienteId
    factura.periodoId = pdef.ItemData(pdef.ListIndex)
    factura.findByClientePeriodo dbapp, IIf(cliente.cobro = 1, False, True)
    periodo.periodoId = factura.periodoId
    periodo.findByPrimaryKey
    pdef.Text = periodo.comboText
    prefijoId = factura.puntoVta
    facturaId = factura.nroComprob
    siva = factura.situacionIVA
    tasa = factura.tasa
    ides = factura.descuento
    desc.Text = Format(factura.descuento, "#,###,##0.00")
    desc.Enabled = False
    nrofac.Caption = Right("0000" & prefijoId, 4) & "-" & Right("00000000" & facturaId, 8)
    efec.Text = factura.fecha
    subt = 0
    ivct = 0
    ivat = 0
    rnit = 0
    If Not IsNull(factura.ivacf) Then ivct = factura.ivacf
    If Not IsNull(factura.ivari) Then ivat = factura.ivari
    If Not IsNull(factura.ivarn) Then rnit = factura.ivarn
    cobr = cliente.cobro
    dmed.Enabled = True
    If cobr = 3 Then dmed.Enabled = False
    medidor.clienteId = cliente.clienteId
    medidor.findColocadoByClienteID dbapp
    nmed.Caption = 0
    If medidor.autoID > 0 Then nmed.Caption = medidor.medidorID
    If Not usuarioDesconectado Then
        objMLec.medidorID = Trim(nmed.Caption)
        objMLec.periodoId = pdef.ItemData(pdef.ListIndex)
        objMLec.findByPrimaryKey dbapp
        fact.Caption = ""
        eact.Caption = 0
        If objMLec.autoID > 0 Then
            fact.Caption = objMLec.fechaLectura
            eact.Caption = objMLec.estado
        End If
        objMLec.medidorID = Trim(nmed.Caption)
        objMLec.periodoId = pdef.ItemData(pdef.ListIndex)
        objMLec.findByMedidorIDPrev dbapp
        If objMLec.autoID = 0 Then
            fant.Caption = ""
            medidor.medidorID = Trim(nmed.Caption)
            medidor.findByMedidorID dbapp
            eant.Caption = 0
            If medidor.autoID > 0 Then eant.Caption = medidor.estadoInicio
        Else
            fant.Caption = objMLec.fechaLectura
            eant.Caption = objMLec.estado
        End If
        creg.Caption = CDbl(eact.Caption) - CDbl(eant.Caption)
    Else
        creg.Caption = 0
        fact.Caption = ""
        fant.Caption = ""
        eact.Caption = ""
        eant.Caption = ""
    End If
    lrub = 0
    For Each detalle In detalle.collectionByLiquidacion(prefijoId, facturaId, dbapp)
        lrub = lrub + 1
        rfac.Rows = lrub + 1
        rfac.row = lrub
        rfac.col = 0
        rfac.Text = Right("00" & detalle.rubroID, 2)
        rfac.col = 1
        rfac.Text = detalle.concepto
        rfac.col = 2
        cant = detalle.cantidad
        rfac.Text = cant
        rfac.col = 3
        impo = detalle.precioUnitario
        rfac.Text = Format(impo, "#,###,##0.00")
        rfac.col = 4
        rfac.Text = Format(impo * cant, "#,###,##0.00")
        rfac.col = 5
        If detalle.IVA Then
            rfac.Text = "Si"
        Else
            rfac.Text = "No"
        End If
        subt = subt + impo * cant
    Next
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey
    subt = subt - ides
    tsub.Caption = Format(subt, "#,###,##0.00")
    tcfl.Caption = Format(ivct, "#,###,##0.00")
    tiva.Caption = Format(ivat, "#,###,##0.00")
    trni.Caption = Format(rnit, "#,###,##0.00")
    tpri.Caption = Format(subt + ivct + ivat + rnit, "#,###,##0.00")
    tseg.Caption = Format(subt + ivct + ivat + rnit + interes(subt + ivct + ivat + rnit, tasa, periodo.fechaPrimero, periodo.fechaSegundo), "#,###,##0.00")

End Sub

Private Sub anul_Click()
Dim factura As New clsMyAFactura
Dim detalle As New clsMyADetalle
Dim novedad As New clsMyANovedad

On Error Resume Next

    prefijoId = Val(Left(Trim(nrofac.Caption), 4))
    facturaId = Val(Right(Trim(nrofac.Caption), 8))
    factura.puntoVta = prefijoId
    factura.nroComprob = facturaId
    factura.findByPrimaryKey dbapp
    If factura.autoID = 0 Then
        MsgBox "Error en el Número de Factura . . ."
        Exit Sub
    End If
    
    If MsgBox("Está SEGURO ?", vbYesNo, "Eliminación de FACTURA") = vbNo Then Exit Sub
    
    factura.anulada = 1
    factura.uid = "admin"
    factura.update dbapp
    
    For Each detalle In detalle.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
        novedad.clienteId = factura.clienteId
        novedad.rubroID = detalle.rubroID
        novedad.periodoId = factura.periodoId
        novedad.findVeces dbapp, False
        If novedad.autoID > 0 Then
            novedad.vecesCobradas = novedad.vecesCobradas - detalle.cantidad
            novedad.update dbapp
        End If
    Next
    
    ' Elimina las novedades de ajuste
    novedad.clienteId = factura.clienteId
    novedad.periodoId = factura.periodoId
    novedad.rubroID = 70
    novedad.delete dbapp
    
    novedad.clienteId = factura.clienteId
    novedad.periodoId = factura.periodoId + 1
    novedad.rubroID = 71
    novedad.delete dbapp
    
    ' Desmarca las facturas que cancelaba la factura anulada
    For Each factura In factura.collectionInteresByLiquidacion(prefijoId, facturaId, dbapp)
        factura.puntoVtaInteres = Null
        factura.nroComprobInteres = Null
        factura.uid = "admin"
        factura.update dbapp
    Next
    
    anul.Enabled = False
    impri.Enabled = False
    Me.cmdImprimirDigital.Enabled = False
    Me.cmdSendMail.Enabled = False
    factu.Enabled = True
    buscar_fac

End Sub

Private Sub cmdTest_Click()
Dim liquidacion As New clsMODFactura
Dim liquidacionrep As clsREPFactura

Dim pagofacil_service As New clsCtlPagoFacil

Dim code1 As String
Dim code2 As String

    Set liquidacionrep = New clsREPFactura
    'code1 = modI2of5.i2of5(pagofacil_service.codigopf(liquidacionrep.findByPrimaryKey(prefijoId, facturaId)))
    code2 = pagofacil_service.codigoI2of5(pagofacil_service.codigopf(liquidacionrep.findByPrimaryKey(prefijoId, facturaId)))
    Set liquidacionrep = Nothing
    
End Sub

Private Sub cmdImprimirDigital_Click()
Dim liquidacion_service As New clsCtlLiquidacion

    liquidacion_service.printLiquidacion Me.hwnd, prefijoId, facturaId, dbapp, Me.consumo, Me.crpLiquidacion, , True
    
End Sub

Private Sub cmdSendMail_Click()
Dim clienteDato As clsMODClienteDato

Dim clienteDatoRep As New clsREPClienteDato

Dim liquidacionService As New clsCtlLiquidacion

    Me.cmdSendMail.Enabled = False

    Set clienteDato = clienteDatoRep.findByClienteId(cliente.clienteId)
    
    clienteDato.email = Trim(Me.txtMail.Text)
    If IsNull(clienteDato.clienteId) Then
        clienteDato.clienteId = cliente.clienteId
        Set clienteDato = clienteDatoRep.add(clienteDato)
    Else
        Set clienteDato = clienteDatoRep.update(clienteDato, cliente.clienteId)
    End If
    
    MsgBox liquidacionService.sendLiquidacion(prefijoId, facturaId, dbapp)
    
    Me.cmdSendMail.Enabled = True
    
End Sub

Private Sub desc_GotFocus()
    
    desc.SelStart = 0
    desc.SelLength = Len(desc.Text)

End Sub

Private Sub desc_KeyPress(KeyAscii As Integer)

On Error Resume Next
    
    If KeyAscii = 13 Then factu.SetFocus

End Sub

Private Sub desc_LostFocus()
    
    If Not IsNumeric(desc.Text) Then desc.Text = 0
    desc.Text = Format(CDbl(desc.Text), "#,###,##0.00")
    llenar (usuarioDesconectado)

End Sub

Private Sub efec_GotFocus()
    
    efec.SelStart = 0
    efec.SelLength = Len(efec.Text)

End Sub

Private Sub efec_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then buscar_fac

End Sub

Private Sub efec_LostFocus()
    
    If Not IsDate(efec.Text) Then MsgBox "La Fecha de EMISION no es válida"

End Sub

Private Sub impri_Click()
Dim liquidacion_service As New clsCtlLiquidacion

    liquidacion_service.printLiquidacion Me.hwnd, prefijoId, facturaId, dbapp, Me.consumo, Me.crpLiquidacion
    
End Sub

Private Sub factu_Click()
Dim fila As Integer
Dim prr As Integer

Dim suspfactura As New clsMyASuspFactura
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura
Dim detalle As New clsMyADetalle
Dim novedad As New clsMyANovedad

Dim clienteRep As New clsREPCliente

Dim ctlFac As New clsCtlFactura
Dim pagofacil_service As New clsCtlPagoFacil
Dim liquidacion_service As New clsCtlLiquidacion

On Error Resume Next
    
    If Not IsDate(efec.Text) Then
        MsgBox "La Fecha de EMISION no es válida"
        Exit Sub
    End If
    suspfactura.periodoIDInicio = pdef.ItemData(pdef.ListIndex)
    suspfactura.findSuspendidaByClienteID dbapp
    If suspfactura.autoID > 0 Then
        If IsNull(suspfactura.periodoIdfin) Or suspfactura.periodoIdfin >= pdef.ItemData(pdef.ListIndex) Then
            MsgBox "Este Cliente tiene la Facturación SUSPENDIDA . . ."
            Exit Sub
        End If
    End If
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey
    
    ' Graba factura nueva
    factura.puntoVta = prefijoId
    factura.nroComprob = facturaId
    factura.fecha = CDate(efec.Text)
    factura.clienteId = cliente.clienteId
    factura.periodoId = pdef.ItemData(pdef.ListIndex)
    factura.periodoIdfin = 0
    factura.tasa = periodo.tasa
    factura.situacionIVA = cliente.situacionIVA
    factura.anulada = 0
    factura.total = tpri.Caption
    factura.ivacf = tcfl.Caption
    factura.ivari = tiva.Caption
    factura.ivarn = trni.Caption
    If Not IsNumeric(desc.Text) Then desc.Text = Format(0, "#,###,##0.00")
    factura.descuento = desc.Text
    factura.uid = "admin"
    factura.pfcodigo = pagofacil_service.codigopf(liquidacion_service.oldFactura2newFactura(factura))
    factura.add dbapp
    
    For fila = 1 To rfac.Rows - 1
        rfac.row = fila
        detalle.puntoVta = prefijoId
        detalle.nroComprob = facturaId
        rfac.col = 0
        detalle.rubroID = rfac.Text
        rfac.col = 1
        detalle.concepto = Left(rfac.Text, 80)
        rfac.col = 2
        detalle.cantidad = rfac.Text
        ' Actualiza novedades nuevas
        novedad.clienteId = cliente.clienteId
        novedad.rubroID = rfac.Text
        novedad.periodoId = pdef.ItemData(pdef.ListIndex)
        novedad.findVeces dbapp
        If novedad.autoID > 0 Then
            novedad.vecesCobradas = novedad.vecesCobradas + Val(rfac.Text)
            novedad.update dbapp
        End If
        rfac.col = 3
        detalle.precioUnitario = rfac.Text
        rfac.col = 5
        If Trim(rfac.Text) = "Si" Then
            detalle.IVA = 1
        Else
            detalle.IVA = 0
        End If
        detalle.uid = "admin"
        detalle.add dbapp
    Next fila
    
    For Each factura In factura.collectionParaInteresByClienteID(cliente.clienteId, dbapp)
        periodo.periodoId = factura.periodoId
        periodo.findByPrimaryKey
        If periodo.fechaSegundo < factura.fechapago Then
            factura.puntoVtaInteres = prefijoId
            factura.nroComprobInteres = facturaId
            factura.uid = "admin"
            factura.update dbapp
        End If
    Next
    
    factu.Enabled = False
    impri.Enabled = True
    Me.cmdImprimirDigital.Enabled = True
    Me.cmdSendMail.Enabled = True
    anul.Enabled = True
    
    impri_Click
    
    buscar_fac

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim crit As String

Dim alicuota As New clsMyAAlicuota
Dim objMRub As New clsMyARubro
Dim periodo As New clsRESTPeriodo

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    aCF = 0.21
    aIVA = 0.27
    aRNI = 0.135
    alicuota.findLast dbapp
    If Not IsNull(alicuota.ivacf) Then aCF = alicuota.ivacf
    If Not IsNull(alicuota.IVA) Then aIVA = alicuota.IVA
    If Not IsNull(alicuota.rni) Then aRNI = alicuota.rni
    efec.Text = Date
    If objMRub.collectionAny(dbapp).Count = 0 Then
        MsgBox "No hay Rubros definidos"
        Unload Me
        Exit Sub
    End If
    
    periodo.fillCombo pdef
    
    If cliente.clienteId > 0 Then buscar_fac
    
    efec.SetFocus

End Sub

Private Sub Form_Load()
Dim alicuota As New clsMyAAlicuota

    alicuota.findLast dbapp
    
    rfac.row = 0
    rfac.col = 0
    rfac.ColWidth(0) = 400
    rfac.Text = "Rub"
    rfac.col = 1
    rfac.ColWidth(1) = 4600
    rfac.Text = "Concepto"
    rfac.col = 2
    rfac.ColWidth(2) = 750
    rfac.Text = "Cantidad"
    rfac.col = 3
    rfac.ColWidth(3) = 1100
    rfac.Text = "Prec. Unitario"
    rfac.col = 4
    rfac.ColWidth(4) = 1100
    rfac.Text = "Imp. Parciales"
    rfac.col = 5
    rfac.ColWidth(5) = 400
    rfac.Text = "iva"
    
    loading = False

End Sub

Private Sub pdef_Click()

    If loading Then Exit Sub
    
    buscar_fac
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteDato As clsMODClienteDato

Dim clienteRep As New clsREPCliente
Dim clienteDatoRep As New clsREPClienteDato

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    mcli.Caption = cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & " " & cliente.inmuebleLocalidad
    
    Set clienteDato = clienteDatoRep.findByClienteId(cliente.clienteId)
    Me.txtMail.Text = clienteDato.email
    
    buscar_fac
    
End Sub

Private Sub txtMail_GotFocus()

    marcarseleccion Me.txtMail
    
End Sub
