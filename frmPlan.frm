VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuotas"
   ClientHeight    =   5955
   ClientLeft      =   1860
   ClientTop       =   2130
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8295
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton apago 
      Caption         =   "A&nular Pago Cuota"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "Elimina la fecha de Pago de la última cuota"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox iactu 
      Caption         =   "I&ncluir Factura Actual"
      Height          =   255
      Left            =   2400
      TabIndex        =   43
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton anulp 
      Caption         =   "&Anular Plan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      ToolTipText     =   "Anula el último PLAN realizado"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton impl 
      Caption         =   "&Imprimir Plan"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      ToolTipText     =   "Permite imprimir las CUOTAS del Plan"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cpag 
      Caption         =   "&Pago de Cuotas"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Permite la Carga del PAGO de Cuotas del PLAN"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton vecp 
      Caption         =   "&Ver Plan"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      ToolTipText     =   "Muestra el PLAN de CUOTAS del Cliente Activo"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton hacp 
      Caption         =   "&Hacer Plan"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "Permite realizar un PLAN de CUOTAS de la Deuda"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame vpla 
      Caption         =   "Plan de Cuotas"
      Height          =   2895
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Width           =   7815
      Begin VB.TextBox pper 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid rdet 
         Height          =   1575
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.CommandButton pconf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         ToolTipText     =   "Graba el PLAN de CUOTAS"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox pfpr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox ptem 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox pcan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "NUEVO Plan"
         Height          =   195
         Index           =   21
         Left            =   6000
         TabIndex        =   57
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label pdpl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   56
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Período (días)"
         Height          =   195
         Index           =   13
         Left            =   6000
         TabIndex        =   36
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de las Cuotas"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha 1er Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interés Mensual"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de Cuotas"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame vimp 
      Caption         =   "Impresión de Cuotas"
      Height          =   1575
      Left            =   240
      TabIndex        =   37
      Top             =   2880
      Width           =   7815
      Begin VB.ComboBox cimh 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox iresu 
         Caption         =   "Imprimir Resumen"
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox icuot 
         Caption         =   "Imprimir Cuotas"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox efec 
         Height          =   285
         Left            =   4080
         TabIndex        =   39
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cuim 
         Caption         =   "I&mprimir Cuota/s"
         Height          =   375
         Left            =   6000
         TabIndex        =   40
         ToolTipText     =   "Imprime la/s CUOTA/S seleccionadas"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cimp 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Width           =   1575
      End
      Begin Crystal.CrystalReport crpCuota 
         Left            =   6000
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "C:cuota.rpt"
         Destination     =   1
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cuota (Hasta)"
         Height          =   195
         Index           =   22
         Left            =   2160
         TabIndex        =   61
         Top             =   840
         Width           =   975
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Plan"
         Height          =   195
         Index           =   19
         Left            =   4080
         TabIndex        =   53
         Top             =   240
         Width           =   315
      End
      Begin VB.Label idpl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   52
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Emisión"
         Height          =   195
         Index           =   15
         Left            =   4080
         TabIndex        =   42
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cuota (Desde)"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   1020
      End
   End
   Begin VB.Frame vapg 
      Caption         =   "Anular Pago de Cuotas"
      Height          =   975
      Left            =   240
      TabIndex        =   44
      Top             =   2880
      Width           =   7815
      Begin VB.CommandButton uanul 
         Caption         =   "An&ular Pago"
         Height          =   375
         Left            =   6000
         TabIndex        =   48
         ToolTipText     =   "Anula el PAGO de una Cuota"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label ufpg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cuota"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   420
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Index           =   17
         Left            =   2160
         TabIndex        =   50
         Top             =   240
         Width           =   525
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de PAGO"
         Height          =   195
         Index           =   16
         Left            =   4080
         TabIndex        =   49
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label ucpg 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label uimpo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame vpag 
      Caption         =   "Pago de Cuotas"
      Height          =   975
      Left            =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   7815
      Begin VB.TextBox pfpa 
         Height          =   285
         Left            =   4080
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton gconf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         ToolTipText     =   "Imputa el PAGO de la CUOTA"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label pimp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label pncu 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de PAGO"
         Height          =   195
         Index           =   12
         Left            =   4080
         TabIndex        =   33
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Index           =   11
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   525
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cuota"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame vver 
      Caption         =   "Plan de Cuotas"
      Height          =   2895
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid vdet 
         Height          =   1575
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Plan"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   55
         Top             =   240
         Width           =   315
      End
      Begin VB.Label vdpl 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label vptm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label vpca 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de Cuotas"
         Height          =   195
         Index           =   10
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interés Mensual"
         Height          =   195
         Index           =   9
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de las Cuotas"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1500
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   62
      Top             =   240
      Width           =   480
   End
   Begin VB.Label ideu 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Importe de la DEUDA"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   1530
   End
End
Attribute VB_Name = "frmPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsin As Integer

Private cliente As New clsMODCliente

Private Function calculateDeuda() As Currency
Dim total As Currency
Dim total_factura As Currency
Dim interes_factura As Currency

Dim planID As Integer

Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim ncredito As New clsMyANCredito
Dim imputado As New clsMyAImputado
Dim recibo As New clsMyARecibo
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

Dim periodos As Collection

On Error Resume Next

    total = 0
    hacp.Enabled = False
    
    Set periodos = periodo.collectionAll
    
    For Each factura In factura.collectionDeudaByClienteId(cliente.clienteId, dbapp)
        Set periodo = periodos("k." & factura.periodoId)
        If periodo.fechaSegundo < Date Or iactu.value = 1 Then
            total_factura = factura.total
            For Each imputado In imputado.collectionByComprobante(1, factura.puntoVta, factura.nroComprob, dbapp, factura.clienteId)
                recibo.serieId = imputado.serieId
                recibo.numero = imputado.numeroID
                recibo.findByPrimaryKey dbapp
                If recibo.total <> 0 Then total_factura = total_factura - recibo.total
            Next
            For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
                total_factura = total_factura - ncredito.total
            Next
            interes_factura = interes(total_factura, factura.tasa, periodo.fechaPrimero, Date)
            total = total + total_factura + interes_factura
        End If
    Next
    vecp.Enabled = False
    cpag.Enabled = False
    impl.Enabled = False
    anulp.Enabled = False
    apago.Enabled = False
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If deuda.autoID > 0 Then
        planID = deuda.planID
        For Each deuda In deuda.collectionByClienteId(cliente.clienteId, dbapp)
            If deuda.pagado = 0 And deuda.cancelada = 0 Then total = total + deuda.deuda
        Next
        vecp.Enabled = True
        impl.Enabled = True
        anulp.Enabled = True
        apago.Enabled = True
        If cuota.collectionPendienteByPlanID(cliente.clienteId, planID, dbapp).Count > 0 Then cpag.Enabled = True
    End If
    If total > 0 Then hacp.Enabled = True
    
    calculateDeuda = total

End Function

Public Sub llenar_cuotas()
Dim total As Currency
Dim cuota As Currency
Dim tasa As Double
Dim fpri As Date
Dim ct As Integer
Dim cancu As Integer
Dim periodo As Integer

On Error Resume Next

    total = ideu.Caption
    cancu = Val(pcan.Text)
    tasa = CDbl(ptem.Text) / 100
    periodo = Val(pper.Text)
    rdet.Rows = 1
    If periodo = 0 Then Exit Sub
    If Not IsDate(pfpr.Text) Then Exit Sub
    fpri = CDate(pfpr.Text)
    If tasa = 0 Then
        cuota = total / cancu
    Else
        cuota = (total + interes(total, tasa, Date, fpri + (cancu * periodo))) / cancu
    End If
    For ct = 1 To cancu
        rdet.Rows = ct + 1
        rdet.row = ct
        rdet.col = 0
        rdet.CellAlignment = flexAlignCenterCenter
        rdet.Text = ct
        rdet.col = 1
        rdet.CellAlignment = flexAlignCenterCenter
        rdet.Text = CDate(pfpr.Text) + (ct - 1) * periodo
        rdet.col = 2
        rdet.CellAlignment = flexAlignRightCenter
        rdet.Text = Format(cuota, "#,###,##0.00")
    Next ct

End Sub

Private Sub anulp_Click()
Dim planID As Integer

Dim clienteId As Long

Dim deuda As New clsMyADeuda
Dim imputado As New clsMyAImputado
Dim cuota As New clsMyACuota
Dim factura As New clsMyAFactura

On Error Resume Next
    
    clienteId = cliente.clienteId
    deuda.clienteId = clienteId
    deuda.findLast dbapp
    If deuda.uid = "" Then
        MsgBox "No hay Plan de Pagos para anular . . ."
        Exit Sub
    End If
    If deuda.planID = 0 Then
        MsgBox "No puede anular Deuda Anterior desde acá . . ."
        Exit Sub
    End If
    planID = deuda.planID
    imputado.clienteId = deuda.clienteId
    imputado.tipoId = 2
    imputado.compSerieID = planID
    imputado.findByPlanIDClienteID dbapp
    If imputado.uid <> "" Then
        MsgBox "Imposible Anular, tiene Recibos Imputados . . ."
        Exit Sub
    End If
    cuota.clienteId = clienteId
    cuota.planID = planID
    cuota.findPagadoByPlanID dbapp
    If cuota.uid <> "" Then
        MsgBox "Imposible Anular, hay Cuotas Pagadas . . ."
        Exit Sub
    End If
    deuda.clienteId = clienteId
    deuda.planID = planID
    deuda.delete dbapp
    For Each cuota In cuota.collectionPendienteByPlanID(clienteId, planID, dbapp)
        cuota.delete dbapp
    Next
    For Each cuota In cuota.collectionByPlanIDCancela(clienteId, planID, dbapp)
        cuota.fechapago = Null
        cuota.cancelada = 0
        cuota.planIdcancela = 0
        cuota.uid = "admin"
        cuota.update dbapp
    Next
    For Each deuda In deuda.collectionByPlanIDCancela(clienteId, planID, dbapp)
        deuda.cancelada = 0
        deuda.planIdcancela = 0
        deuda.uid = "admin"
        deuda.update dbapp
    Next
    For Each factura In factura.collectionByPlanIDCancela(clienteId, planID, dbapp)
        factura.cancelada = 0
        factura.planIdcancela = 0
        factura.uid = "admin"
        factura.update dbapp
    Next
    
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")
    frmPlan.Height = rsin

End Sub

Private Sub apago_Click()
Dim crit As String

Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota
Dim imputado As New clsMyAImputado

On Error Resume Next
    
    vapg.Visible = True
    vpag.Visible = False
    vver.Visible = False
    vpla.Visible = False
    vimp.Visible = False
    frmPlan.Height = rsin + 1200
    If deuda.collectionByClienteId(cliente.clienteId, dbapp).Count = 0 Then frmPlan.Height = rsin
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    
    If cuota.collectionPagadoByPlanID(cliente.clienteId, deuda.planID, dbapp).Count = 0 Then
        frmPlan.Height = rsin
        Exit Sub
    End If
    
    cuota.clienteId = cliente.clienteId
    cuota.planID = deuda.planID
    cuota.findPagadoByPlanID dbapp
    If cuota.uid = "" Then
        MsgBox "No hay más cuotas Pagadas de este Plan . . ."
        frmPlan.Height = rsin
        Exit Sub
    End If
    
    imputado.clienteId = cliente.clienteId
    imputado.compSerieID = cuota.planID
    imputado.compNumeroID = cuota.cuotaID
    imputado.findImputadoByCuotaID dbapp
    If imputado.autoID > 0 Then
        MsgBox "No puede Anular, hay Recibos Imputados . . ."
        frmPlan.Height = rsin
        Exit Sub
    End If
    ucpg.Caption = cuota.planID & "-" & Right("00000" & cuota.cuotaID, 5)
    uimpo.Caption = Format(cuota.importe, "#,###,##0.00")
    ufpg.Caption = cuota.fechapago
    
End Sub

Private Sub cimp_Click()
Dim ik As Integer

On Error Resume Next
    
    If cimp.ListIndex = 0 Then
        cimh.Enabled = False
        Exit Sub
    Else
        cimh.Enabled = True
        cimh.Clear
        For ik = cimp.ListIndex To cimp.ListCount - 1
            cimh.AddItem cimp.List(ik)
            cimh.ItemData(cimh.NewIndex) = cimp.ItemData(ik)
        Next ik
    End If
    cimh.ListIndex = 0

End Sub

Private Sub cpag_Click()
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

On Error Resume Next
    
    vpag.Visible = True
    vver.Visible = False
    vpla.Visible = False
    vimp.Visible = False
    vapg.Visible = False
    frmPlan.Height = rsin + 1200
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If cuota.collectionPendienteByPlanID(deuda.clienteId, deuda.planID, dbapp).Count = 0 Then
        MsgBox "No hay más Cuotas adeudadas . . ."
        frmPlan.Height = rsin
        Exit Sub
    End If
    cuota.clienteId = deuda.clienteId
    cuota.planID = deuda.planID
    cuota.findPendienteByPlanID dbapp
    pncu.Caption = cuota.planID & "-" & Right("00000" & cuota.cuotaID, 5)
    pimp.Caption = Format(cuota.importe, "#,###,##0.00")
    pfpa.Text = Date

End Sub

Private Sub cuim_Click()
Dim mens As String
Dim cate As String
Dim medi As String
Dim cuit As String

Dim cant As Integer
Dim desde As Integer
Dim hasta As Integer
Dim np As Integer

Dim ini As Currency
Dim total As Currency

Dim operador As New clsMyAOperador

Dim impresionService As New clsCtlImpresion
Dim liquidacion_service As New clsCtlLiquidacion

On Error Resume Next

    If iresu.value = 1 Then
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
        impresionService.printReport Me.crpCuota, "rptPlan", dbapp.stringConnection, , _
            Array(Array("cliente_id", cliente.clienteId), Array("plan_id", Val(Me.idpl.Caption))), , , _
            Array(Array("nomope", operador.razonSocial), _
            Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
            Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
            Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
            Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
            Array("titulo", "Resumen Facturas y Cuotas Incluídas en el Plan de Cuotas"))
            
    End If
    If icuot.value = 1 Then
        If Not IsDate(efec.Text) Then
            MsgBox "La Fecha de EMISION no es válida"
            Exit Sub
        End If
        If cimp.ListIndex = 0 Then
            desde = 1
            hasta = cimp.ItemData(cimp.ListCount - 1)
        Else
            desde = cimp.ItemData(cimp.ListIndex)
            hasta = cimh.ItemData(cimh.ListIndex)
        End If

'Imprime la/s cuotas
        For cant = desde To hasta
            liquidacion_service.printCuota cliente.clienteId, Val(Me.idpl.Caption), cant, efec.Text, dbapp, Me.crpCuota
        Next cant
    End If

End Sub

Private Sub efec_GotFocus()
    
    efec.SelStart = 0
    efec.SelLength = Len(efec.Text)

End Sub

Private Sub efec_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cuim.SetFocus

End Sub

Private Sub efec_LostFocus()
    
    If Not IsDate(efec.Text) Then MsgBox "La Fecha de EMISION no es válida"

End Sub

Private Sub fin_Click()
    
    frmPlan.Height = rsin
    
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False
    
    Unload Me

End Sub

Private Sub Form_Activate()

On Error Resume Next
    
    frmPlan.Height = rsin
    vpla.Visible = False
    If cliente.clienteId = 0 Then Exit Sub
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")
    frmPlan.Height = rsin
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False

End Sub

Private Sub Form_Load()
    
    rdet.row = 0
    rdet.col = 0
    rdet.ColWidth(0) = 1650
    rdet.Text = "Cuota"
    rdet.col = 1
    rdet.ColWidth(1) = 1650
    rdet.Text = "Vencimiento"
    rdet.col = 2
    rdet.ColWidth(2) = 1650
    rdet.Text = "Importe"
    vdet.row = 0
    vdet.col = 0
    vdet.ColWidth(0) = 1450
    vdet.Text = "Cuota"
    vdet.col = 1
    vdet.ColWidth(1) = 1450
    vdet.Text = "Vencimiento"
    vdet.col = 2
    vdet.ColWidth(2) = 1450
    vdet.Text = "Importe"
    vdet.col = 3
    vdet.ColWidth(3) = 650
    vdet.Text = "Pagada"
    rsin = 3250
    frmPlan.Height = rsin
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False

End Sub

Private Sub gconf_Click()
Dim inte As Currency
Dim ulti As Currency

Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

On Error Resume Next
    
    If Not IsDate(pfpa.Text) Then
        MsgBox "La Fecha de PAGO no es válida"
        Exit Sub
    End If
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    cuota.clienteId = deuda.clienteId
    cuota.planID = deuda.planID
    cuota.findPendienteByPlanID dbapp
    deuda.deuda = deuda.deuda - cuota.importe
    deuda.cuotasPagadas = cuota.cuotaID
    If Abs(deuda.deuda) < 0.02 Then deuda.pagado = 1
    deuda.uid = "admin"
    deuda.update dbapp
    cuota.fechapago = CDate(pfpa.Text)
    cuota.ddl = "admin"
    cuota.update dbapp
    If cuota.fechapago > cuota.fechaVencimiento Then
        inte = interes(cuota.importe, deuda.tasa, cuota.fechaVencimiento, cuota.fechapago)
        cuota.clienteId = deuda.clienteId
        cuota.planID = deuda.planID
        cuota.cuotaID = deuda.cuotas
        cuota.findByPrimaryKey dbapp
        ulti = cuota.fechaVencimiento
        cuota.clienteId = deuda.clienteId
        cuota.planID = deuda.planID
        cuota.cuotaID = deuda.cuotas + 1
        cuota.findByPrimaryKey dbapp
        If cuota.autoID = 0 Then
            cuota.clienteId = deuda.clienteId
            cuota.planID = deuda.planID
            cuota.cuotaID = deuda.cuotas + 1
            cuota.fechaVencimiento = ulti + 30
            cuota.importe = inte
            cuota.uid = "admin"
            cuota.add dbapp
        Else
            cuota.importe = cuota.importe + inte
            cuota.uid = "admin"
            cuota.update dbapp
        End If
        
        deuda.deuda = deuda.deuda + inte
        deuda.uid = "admin"
        deuda.update dbapp
    End If
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")
    vpag.Visible = False
    frmPlan.Height = rsin

End Sub

Private Sub hacp_Click()
Dim planID As Integer

Dim deuda As New clsMyADeuda

On Error Resume Next
    
    vpla.Visible = True
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False
    frmPlan.Height = rsin + 3100
    pcan.Text = 1
    pfpr.Text = Date + 15
    ptem.Text = Format(0, "#,###,##0.00")
    pper.Text = 0
    planID = 1
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If deuda.autoID > 0 Then planID = deuda.planID + 1
    pdpl.Caption = planID
    llenar_cuotas
    pcan.SetFocus

End Sub

Private Sub iactu_Click()
    
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")
    If vpla.Visible Then llenar_cuotas

End Sub

Private Sub icuot_Click()

On Error Resume Next
    
    If icuot.value = 1 Then
        cimp.Enabled = True
        If cimp.ListIndex Then cimh.Enabled = True
    Else
        cimp.Enabled = False
        If cimh.Enabled Then cimh.Enabled = False
    End If

End Sub

Private Sub impl_Click()
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

On Error Resume Next
    
    vimp.Visible = True
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vapg.Visible = False
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    cimp.Clear
    cimp.AddItem "Todas"
    idpl.Caption = deuda.planID
    For Each cuota In cuota.collectionByPlanID(deuda.clienteId, deuda.planID, dbapp)
        cimp.AddItem cuota.cuotaID & "a. cuota"
        cimp.ItemData(cimp.NewIndex) = cuota.cuotaID
    Next
    cimp.ListIndex = 0
    cimh.Enabled = False
    frmPlan.Height = rsin + 1700
    efec.Text = Date
    cimp.SetFocus

End Sub

Private Sub pcan_GotFocus()
    
    pcan.SelStart = 0
    pcan.SelLength = Len(pcan.Text)

End Sub

Private Sub pcan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ptem.SetFocus

End Sub

Private Sub pcan_LostFocus()
    
    If Val(pcan.Text) < 1 Then pcan.Text = 1
    llenar_cuotas

End Sub

Private Sub pconf_Click()
Dim planID, ct As Integer

Dim importeCuota As Currency

Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

Dim periodos As Collection

On Error Resume Next
    
    If Not IsDate(pfpr.Text) Then
        MsgBox "La Fecha del 1er VENCIMIENTO no es válida"
        pfpr.SetFocus
        Exit Sub
    End If
    If CDbl(ptem.Text) = 0 Then
        MsgBox "La Tasa debe ser mayor que cero"
        ptem.SetFocus
        Exit Sub
    End If
    If Val(pper.Text) = 0 Then
        MsgBox "El Período debe ser mayor que cero"
        pper.SetFocus
        Exit Sub
    End If
    planID = 1
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If deuda.autoID > 0 Then planID = deuda.planID + 1
    Set periodos = periodo.collectionAll
    'MySQL
    For Each factura In factura.collectionDeudaByClienteId(cliente.clienteId, dbapp)
        Set periodo = periodos("k." & factura.periodoId)
        If periodo.fechaSegundo < Date Or iactu.value = 1 Then
            factura.cancelada = 1
            factura.planIdcancela = planID
            factura.fechapago = Date
            factura.uid = "admin"
            factura.update dbapp
        End If
    Next

    'MySQL
    For Each deuda In deuda.collectionPendienteByClienteID(cliente.clienteId, dbapp)
        For Each cuota In cuota.collectionByPlanID(deuda.clienteId, deuda.planID, dbapp)
            cuota.fechapago = Date
            cuota.cancelada = True
            cuota.planIdcancela = planID
            cuota.uid = "admin"
            cuota.update dbapp
        Next
        deuda.cancelada = True
        deuda.planIdcancela = planID
        deuda.uid = "admin"
        deuda.update dbapp
    Next
    
    
    For ct = 1 To pcan.Text
        rdet.row = ct
        cuota.clienteId = cliente.clienteId
        cuota.planID = planID
        cuota.cuotaID = ct
        rdet.col = 1
        cuota.fechaVencimiento = CDate(rdet.Text)
        rdet.col = 2
        cuota.importe = rdet.Text
        importeCuota = rdet.Text
        cuota.uid = "admin"
        cuota.save dbapp
    Next ct
    deuda.clienteId = cliente.clienteId
    deuda.planID = planID
    deuda.deuda = importeCuota * pcan.Text
    deuda.cuotas = pcan.Text
    deuda.planIdcancela = 0
    If Len(Trim(ptem.Text)) = 0 Then ptem.Text = Format(0, "#,###,##0.00")
    deuda.tasa = CDbl(ptem.Text) / 100
    deuda.periodo = pper.Text
    deuda.uid = "admin"
    deuda.save dbapp
    frmPlan.Height = rsin
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")

End Sub

Private Sub pfpa_GotFocus()
    
    pfpa.SelStart = 0
    pfpa.SelLength = Len(pfpa.Text)

End Sub

Private Sub pfpa_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then gconf.SetFocus

End Sub

Private Sub pfpa_LostFocus()
    
    If Not IsDate(pfpa.Text) Then MsgBox "La Fecha de PAGO no es válida"

End Sub

Private Sub pfpr_GotFocus()
    
    pfpr.SelStart = 0
    pfpr.SelLength = Len(pfpr.Text)

End Sub

Private Sub pfpr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then pper.SetFocus

End Sub

Private Sub pfpr_LostFocus()

On Error Resume Next
    
    If Not IsDate(pfpr.Text) Then
        MsgBox "La Fecha del 1er VENCIMIENTO no es válida"
        Exit Sub
    End If
    llenar_cuotas

End Sub

Private Sub pper_GotFocus()
    
    pper.SelStart = 0
    pper.SelLength = Len(pper.Text)

End Sub

Private Sub pper_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then pconf.SetFocus

End Sub

Private Sub pper_LostFocus()
    
    llenar_cuotas

End Sub

Private Sub ptem_GotFocus()
    
    MsgBox "Recuerde que la Tasa no debe ser mayor que la" & Chr(13) & Chr(13) & "TASA ACTIVA del Banco Nación"
    ptem.SelStart = 0
    ptem.SelLength = Len(ptem.Text)

End Sub

Private Sub ptem_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then pfpr.SetFocus

End Sub

Private Sub ptem_LostFocus()
    
    If Not IsNumeric(ptem.Text) Then ptem.Text = 0
    ptem.Text = Format(CDbl(ptem.Text), "#,###,##0.00")

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")
    frmPlan.Height = rsin
    vpla.Visible = False
    vver.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False

End Sub

Private Sub uanul_Click()
Dim total As Currency

Dim ncuo As Integer

Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota
Dim imputado As New clsMyAImputado

On Error Resume Next
    
    frmPlan.Height = rsin
    vapg.Visible = False
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    If cuota.collectionPendienteByPlanID(deuda.clienteId, deuda.planID, dbapp).Count = 0 Then Exit Sub
    cuota.clienteId = deuda.clienteId
    cuota.planID = deuda.planID
    cuota.findPagadoByPlanID dbapp
    total = cuota.importe
    ncuo = cuota.cuotaID
    imputado.clienteId = cuota.clienteId
    imputado.serieId = cuota.planID
    imputado.numeroID = cuota.cuotaID
    imputado.findImputadoByCuotaID dbapp
    If imputado.autoID > 0 Then
        MsgBox "No puede anular Pago, Recibo IMPUTADO . . ."
        Exit Sub
    End If
    cuota.fechapago = Null
    cuota.uid = "admin"
    cuota.update dbapp
    If ncuo <= deuda.cuotas Then deuda.deuda = deuda.deuda + total
    deuda.pagado = 0
    deuda.cuotasPagadas = ncuo - 1
    deuda.update dbapp
    ideu.Caption = Format(calculateDeuda, "#,###,##0.00")

End Sub

Private Sub vecp_Click()
Dim planID As Integer
Dim ct As Integer

Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

On Error Resume Next
    
    vver.Visible = True
    vpla.Visible = False
    vpag.Visible = False
    vimp.Visible = False
    vapg.Visible = False
    frmPlan.Height = rsin + 3100
    deuda.clienteId = cliente.clienteId
    deuda.findLast dbapp
    planID = deuda.planID
    vpca.Caption = deuda.cuotas
    vptm.Caption = Format(deuda.tasa * 100, "#,###,##0.00")
    vdpl.Caption = planID
    ct = 1
    For Each cuota In cuota.collectionByPlanID(deuda.clienteId, deuda.planID, dbapp)
        vdet.Rows = ct + 1
        vdet.row = ct
        vdet.col = 0
        vdet.CellAlignment = flexAlignCenterCenter
        vdet.Text = cuota.cuotaID
        vdet.col = 1
        vdet.CellAlignment = flexAlignCenterCenter
        vdet.Text = cuota.fechaVencimiento
        vdet.col = 2
        vdet.Text = Format(cuota.importe, "#,###,##0.00")
        vdet.col = 3
        vdet.CellAlignment = flexAlignCenterCenter
        If IsNull(cuota.fechapago) Then
            vdet.Text = "No"
        Else
            vdet.Text = "Si"
        End If
        ct = ct + 1
    Next

End Sub
