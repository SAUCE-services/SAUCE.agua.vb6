VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos a Cuenta"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7995
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin Crystal.CrystalReport impr 
      Left            =   7440
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cimp 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton rimpr 
      Caption         =   "&Imprimir Recibo"
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      ToolTipText     =   "Imprime el RECIBO Actual"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame impua 
      Caption         =   "Imputar a :"
      Height          =   975
      Left            =   2160
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
      Begin VB.OptionButton icuo 
         Caption         =   "&Cuota"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton ifac 
         Caption         =   "&Factura"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame separ 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   7575
   End
   Begin VB.Frame treci 
      Caption         =   "Recibo"
      Height          =   975
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton rnue 
         Caption         =   "&Nuevo"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton rant 
         Caption         =   "An&terior"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.ComboBox fimp 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton impu 
      Caption         =   ">>"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Imputar el RECIBO a la FACTURA"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton dimp 
      Caption         =   "<<"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      ToolTipText     =   "Retirar la Imputación del RECIBO a la FACTURA"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox lrec 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton fin 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton anul 
      Caption         =   "&Anular Recibo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "Permite Anular el RECIBO Actual"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton greci 
      Caption         =   "&Generar Recibo"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      ToolTipText     =   "Genera el RECIBO"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox imprec 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox efec 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid raimp 
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid rimpu 
      Height          =   1095
      Left            =   4320
      TabIndex        =   15
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Recibos a IMPUTAR"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Factura"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   17
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Recibos IMPUTADOS"
      Height          =   195
      Index           =   5
      Left            =   4320
      TabIndex        =   16
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   10
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de Recibo"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label nrorec 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   8
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private serieRecibo As Integer
Private numeroRecibo As Long

Private cliente As New clsMODCliente

Private Sub anul_Click()
Dim recibo As New clsMyARecibo
Dim imputado As New clsMyAImputado

On Error Resume Next
    
    If imputado.collectionByRecibo(serieRecibo, numeroRecibo, dbapp).Count > 0 Then
        MsgBox "Este RECIBO está IMPUTADO. No se puede anular . . ."
        Exit Sub
    End If
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .findByPrimaryKey dbapp
        
        .anulado = True
        .update dbapp
    End With
    
    rant_Click

End Sub

Private Sub cimp_Click()
    
    llenar_imp

End Sub

Private Sub efec_GotFocus()
    
    efec.SelStart = 0
    efec.SelLength = Len(efec.Text)

End Sub

Private Sub fimp_Click()
    
    llenar_imp

End Sub

Private Sub fin_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    efec.Text = Date
    rnue.value = True
    
    llenar

End Sub

Private Sub Form_Load()
    
    raimp.ColWidth(0) = 1200
    rimpu.ColWidth(0) = 1200

End Sub

Private Sub greci_Click()
Dim recibo As New clsMyARecibo
Dim clienteRep As clsREPCliente

On Error Resume Next
    
    If Not IsDate(efec.Text) Then
        MsgBox "Fecha NO Válida . . ."
        efec.SetFocus
        Exit Sub
    End If
    If Len(Trim(imprec.Text)) = 0 Then
        MsgBox "Debe especificar un IMPORTE . . ."
        imprec.SetFocus
        Exit Sub
    End If
    If CDbl(imprec.Text) <= 0 Then
        MsgBox "El importe debe ser mayor que cero . . ."
        imprec.SetFocus
        Exit Sub
    End If
    Set clienteRep = New clsREPCliente
    If clienteRep.collectionActivos().Count = 0 Then
        MsgBox "No Existen CLIENTES . . ."
        Unload Me
        Exit Sub
    End If
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .fecha = CDate(efec.Text)
        .clienteId = cliente.clienteId
        .situacionIVA = cliente.situacionIVA
        .total = CDbl(imprec.Text)
        .uid = "admin"
        
        .save dbapp
    End With
    
    rnue_Click
    llenar

End Sub

Private Sub icuo_Click()
    
    llenar

End Sub

Private Sub ifac_Click()
    
    llenar

End Sub

Private Sub imprec_GotFocus()
    
    imprec.SelStart = 0
    imprec.SelLength = Len(imprec.Text)

End Sub

Private Sub imprec_LostFocus()
    
    If Not IsNumeric(imprec.Text) Then imprec.Text = 0
    imprec.Text = Format(CDbl(imprec.Text), "#,###,##0.00")

End Sub

Private Sub rimpr_Click()
Dim recibo As New clsMyARecibo

On Error Resume Next
    
    If rnue.value Then
        MsgBox "Debe seleccionar un Recibo . . ."
        Exit Sub
    End If
    serieRecibo = Val(Left(lrec.Text, 4))
    numeroRecibo = Val(Right(lrec.Text, 8))
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .findByPrimaryKey dbapp
        
        If .autoID = 0 Then
            MsgBox " El Recibo no Existe . . ."
            Exit Sub
        End If
        If Not .imputado Then
            MsgBox "No puede imprimirse, no está Imputado . . ."
            Exit Sub
        End If
    End With
    
    impri_rec

End Sub

Private Sub lrec_Click()
Dim recibo As New clsMyARecibo

On Error Resume Next
    
    serieRecibo = Val(Left(lrec.Text, 4))
    numeroRecibo = Val(Right(lrec.Text, 8))
    
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .findByPrimaryKey dbapp
        
        imprec.Text = Format(.total, "#,###,##0.00")
        efec.Text = .fecha
    End With

End Sub

Private Sub rant_Click()
Dim crit As String, ic As Integer

Dim recibo As New clsMyARecibo

On Error Resume Next
    
    nrorec.Visible = False
    efec.Enabled = False
    imprec.Enabled = False
    greci.Enabled = False
    anul.Enabled = True
    rimpr.Enabled = True
    lrec.Visible = True
    lrec.Clear
    
    If recibo.collectionAny(dbapp).Count = 0 Then
        MsgBox "No hay datos de Recibos . . ."
        rnue.value = True
        rnue_Click
        Exit Sub
    End If
    
    ic = 0
    For Each recibo In recibo.collectionByClienteID(cliente.clienteId, dbapp)
        With recibo
            If ic = 0 Then
                serieRecibo = .serieId
                numeroRecibo = .numero
            End If
            ic = ic + 1
            lrec.AddItem Right("0000" & .serieId, 4) & "-" & Right("00000000" & .numero, 8)
        End With
    Next
    
    If lrec.ListCount = 0 Then
        MsgBox "No hay RECIBOS para este cliente . . ."
        rnue.value = True
        rnue_Click
        Exit Sub
    End If
    
    lrec.ListIndex = 0
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .findByPrimaryKey dbapp
        
        imprec.Text = Format(.total, "#,###,##0.00")
        efec.Text = .fecha
    End With

End Sub

Private Sub rnue_Click()
Dim objMOpe As New clsMyAOperador
Dim recibo As New clsMyARecibo

On Error Resume Next
    
    efec.Enabled = True
    greci.Enabled = True
    anul.Enabled = False
    rimpr.Enabled = False
    lrec.Visible = False
    nrorec.Visible = True
    imprec.Enabled = True
    serieRecibo = 1
    numeroRecibo = 1
    
    With objMOpe
        .findLast dbapp
        If Not IsNull(.reciboSerie) Then serieRecibo = .reciboSerie
        If Not IsNull(.recibo) Then numeroRecibo = .recibo
    End With
    
    With recibo
        .findLastLast dbapp
        If serieRecibo < .serieId Then serieRecibo = .serieId
        .serieId = serieRecibo
        .findLast dbapp
        If .autoID > 0 Then If numeroRecibo <= .numero Then numeroRecibo = .numero + 1
    End With
    nrorec.Caption = Right("0000" & serieRecibo, 4) & "-" & Right("00000000" & numeroRecibo, 8)
    imprec.Text = Format(0, "#,###,##0.00")
    efec.Text = Date

End Sub

Private Sub dimp_Click()
Dim serieRecibo As Integer
Dim tipoComprobanteID As Integer
Dim serieComprobanteID As Integer
Dim clienteId As Integer

Dim numeroRecibo As Long
Dim numeroComprobante As Long

Dim total As Currency

Dim imputado As New clsMyAImputado
Dim recibo As New clsMyARecibo
Dim factura As New clsMyAFactura
Dim cuota As New clsMyACuota
Dim deuda As New clsMyADeuda

On Error Resume Next
    
    If rimpu.row < 0 Then
        MsgBox "Debe elegir el Recibo Imputado . . ."
        Exit Sub
    End If
    serieRecibo = Val(Left(rimpu.Text, 4))
    numeroRecibo = Val(Right(rimpu.Text, 8))
    clienteId = cliente.clienteId
    If ifac.value = True Then
        tipoComprobanteID = 1
        serieComprobanteID = Val(Left(fimp.Text, 4))
        numeroComprobante = Val(Right(fimp.Text, 8))
    Else
        tipoComprobanteID = 2
        serieComprobanteID = Val(Left(cimp.Text, 4))
        numeroComprobante = Val(Right(cimp.Text, 8))
    End If
    
    With imputado
        .serieId = serieRecibo
        .numeroID = numeroRecibo
        .tipoId = tipoComprobanteID
        .compSerieID = serieComprobanteID
        .compNumeroID = numeroComprobante
        
        .findByReciboComprobante dbapp
        
        .delete dbapp
    End With
    
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        .findByPrimaryKey dbapp
        
        If .autoID > 0 Then
            total = .total
            .imputado = 0
            .update dbapp
        End If
    End With
    
    If ifac.value Then
        factura.puntoVta = serieComprobanteID
        factura.nroComprob = numeroComprobante
        factura.findByPrimaryKey dbapp
        
        factura.pagada = False
        factura.fechapago = Null
        factura.puntoVtaInteres = Null
        factura.nroComprobInteres = Null
        factura.uid = "admin"
        factura.update dbapp
    Else
        With cuota
            .clienteId = clienteId
            .planID = serieComprobanteID
            .cuotaID = numeroComprobante
            .findByPrimaryKey dbapp
            
            .fechapago = Null
            .uid = "admin"
            .update dbapp
        End With
        
        With deuda
            .clienteId = clienteId
            .planID = serieComprobanteID
            .findByPrimaryKey dbapp
            
            .deuda = .deuda + total
            .pagado = False
            .cuotasPagadas = cuota.cuotaID
            .update dbapp
        End With
    End If
    
    llenar

End Sub

Private Sub impu_Click()
Dim serieRecibo As Integer
Dim serieComprobante As Integer
Dim tipocomprobante As Integer
Dim clienteId As Integer
Dim ultimaCuota As Integer

Dim numeroRecibo As Long
Dim numeroComprobante As Long

Dim total As Currency
Dim totalRecibos As Currency
Dim imrec As Currency

Dim crit, critp, critq As String, fec As Date

Dim imputado As New clsMyAImputado
Dim recibo As New clsMyARecibo
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim deuda As New clsMyADeuda
Dim cuota As New clsMyACuota

On Error Resume Next
    
    If Not IsDate(efec.Text) Then
        MsgBox "Fecha No Válida . . ."
        efec.SetFocus
        Exit Sub
    End If
    If raimp.row < 0 Then
        MsgBox "Debe elegir el Recibo a Imputar . . ."
        Exit Sub
    End If
    serieRecibo = Val(Left(raimp.Text, 4))
    numeroRecibo = Val(Right(raimp.Text, 8))
    clienteId = cliente.clienteId
    If ifac.value = True Then
        tipocomprobante = 1
        serieComprobante = Val(Left(fimp.Text, 4))
        numeroComprobante = Val(Right(fimp.Text, 8))
        If fimp.ListIndex < 0 Then
            MsgBox "No hay Factura para Imputar . . ."
            Exit Sub
        End If
    Else
        tipocomprobante = 2
        serieComprobante = Val(Left(cimp.Text, 4))
        numeroComprobante = Val(Right(cimp.Text, 8))
        If cimp.ListIndex < 0 Then
            MsgBox "No hay Cuota para Imputar . . ."
            Exit Sub
        End If
    End If
    
    With imputado
        .tipoId = tipocomprobante
        .compSerieID = serieComprobante
        .compNumeroID = numeroComprobante
        .serieId = serieRecibo
        .numeroID = numeroRecibo
        .clienteId = clienteId
        .fecha = efec.Text
        .uid = "admin"
        
        .save dbapp
    End With
    
    With recibo
        .serieId = serieRecibo
        .numero = numeroRecibo
        
        .findByPrimaryKey dbapp
        
        imrec = .total
        .imputado = True
        .update dbapp
    End With
    
    If tipocomprobante = 1 Then
        With factura
            .puntoVta = serieComprobante
            .nroComprob = numeroComprobante
            
            .findByPrimaryKey dbapp
            
            total = .total
        End With
        totalRecibos = 0
        
        fec = CDate("01/01/1980")
        For Each imputado In imputado.collectionByComprobante(1, serieComprobante, numeroComprobante, dbapp)
            With recibo
                .serieId = imputado.serieId
                .numero = imputado.numeroID
                .findByPrimaryKey dbapp
                totalRecibos = totalRecibos + .total
                If .fecha > fec Then fec = .fecha
            End With
        Next
        If totalRecibos >= total Then
            periodo.periodoId = factura.periodoId
            periodo.findByPrimaryKey
            
            factura.puntoVta = serieComprobante
            factura.nroComprob = numeroComprobante
            factura.findByPrimaryKey dbapp
        
            factura.pagada = True
            factura.fechapago = fec
            If fec <= periodo.fechaSegundo Then
                factura.puntoVtaInteres = 0
                factura.nroComprobInteres = 0
            Else
                factura.puntoVtaInteres = Null
                factura.nroComprobInteres = Null
            End If
            factura.uid = "admin"
            factura.update dbapp
        End If
    Else
        With cuota
            .clienteId = clienteId
            .planID = serieComprobante
            .cuotaID = numeroComprobante
            .findByPrimaryKey dbapp
            total = .importe
        End With
        totalRecibos = 0
        fec = CDate("01/01/1980")
        For Each imputado In imputado.collectionByComprobante(2, serieComprobante, numeroComprobante, dbapp, clienteId)
            With recibo
                .serieId = imputado.serieId
                .numero = imputado.numeroID
                .findByPrimaryKey dbapp
                totalRecibos = totalRecibos + .total
                If imputado.fecha > fec Then fec = imputado.fecha
            End With
        Next
        If totalRecibos >= total Then
            With cuota
                .clienteId = clienteId
                .planID = serieComprobante
                .cuotaID = numeroComprobante
                .findByPrimaryKey dbapp
                
                .fechapago = fec
                .uid = "admin"
                .update dbapp
            End With
        End If
        ultimaCuota = 0
        With cuota
            .clienteId = clienteId
            .planID = serieComprobante
            .findLastPagada dbapp
            
            If .autoID > 0 Then ultimaCuota = .cuotaID
        End With
        With deuda
            .clienteId = clienteId
            .planID = serieComprobante
            .findByPrimaryKey dbapp
            .deuda = .deuda - imrec
            .cuotasPagadas = ultimaCuota
            If Abs(.deuda) < 0.02 Then .pagado = True
            .uid = "admin"
            .update dbapp
        End With
    End If
    
    llenar

End Sub

Public Sub llenar()
Dim factura As New clsMyAFactura
Dim cuota As New clsMyACuota

On Error Resume Next
    
    If ifac.value = False And icuo.value = False Then ifac.value = True
    If ifac.value = True Then
        etiq(6).Caption = "Facturas"
        fimp.Visible = True
        cimp.Visible = False
    Else
        etiq(6).Caption = "Cuotas"
        fimp.Visible = False
        cimp.Visible = True
    End If
    
    fimp.Clear
    For Each factura In factura.collectionDeudaByClienteId(cliente.clienteId, dbapp)
        With factura
            fimp.AddItem Right("0000" & .puntoVta, 4) & "-" & Right("00000000" & .nroComprob, 8)
        End With
    Next
    cimp.Clear
    For Each cuota In cuota.collectionDeudaByClienteId(cliente.clienteId, dbapp)
        With cuota
            cimp.AddItem Right("0000" & .planID, 4) & "-" & Right("00000000" & .cuotaID, 8)
        End With
    Next
    If fimp.ListCount Then fimp.ListIndex = 0
    If cimp.ListCount Then cimp.ListIndex = 0
    
    llenar_aim
    llenar_imp

End Sub

Public Sub llenar_aim()
Dim ct As Integer

Dim recibo As New clsMyARecibo

On Error Resume Next
    
    raimp.Clear
    raimp.Rows = 0
    ct = 0
    
    If recibo.collectionAny(dbapp).Count = 0 Then Exit Sub
    
    For Each recibo In recibo.collectionPendienteByClienteID(cliente.clienteId, dbapp)
        With recibo
            ct = ct + 1
            raimp.Rows = ct
            raimp.col = 0
            raimp.row = ct - 1
            raimp.Text = Right("0000" & .serieId, 4) & "-" & Right("00000000" & .numero, 8)
        End With
    Next

End Sub

Public Sub llenar_imp()
Dim comprobanteSerie As Integer
Dim comprobanteTipo As Integer
Dim clienteId As Integer
Dim ct As Integer

Dim comprobanteNumero As Long

Dim crit As String

Dim imputado As New clsMyAImputado

On Error Resume Next
    
    clienteId = cliente.clienteId
    If ifac.value = True Then
        comprobanteTipo = 1
        comprobanteSerie = Val(Left(fimp.Text, 4))
        comprobanteNumero = Val(Right(fimp.Text, 8))
    Else
        comprobanteTipo = 2
        comprobanteSerie = Val(Left(cimp.Text, 4))
        comprobanteNumero = Val(Right(cimp.Text, 8))
    End If
    rimpu.Clear
    rimpu.Rows = 0
    ct = 0
    If imputado.collectionAny(dbapp).Count = 0 Then Exit Sub
    For Each imputado In imputado.collectionByComprobante(comprobanteTipo, comprobanteSerie, comprobanteNumero, dbapp, clienteId)
        With imputado
            ct = ct + 1
            rimpu.Rows = ct
            rimpu.col = 0
            rimpu.row = ct - 1
            rimpu.Text = Right("0000" & .serieId, 4) & "-" & Right("00000000" & .numeroID, 8)
        End With
    Next

End Sub

Public Sub impri_rec()
Dim crit, mens, cate, medi, doc As String
Dim cant, desde, hasta, np, serieRecibo As Integer, numeroRecibo As Long

On Error Resume Next
    
'    impr.ReportFileName = Confi.PathDB & "recibo.rpt"
'    impr.Destination = Impresora
'    If Not IsDate(efec.Text) Then
'        MsgBox "La Fecha de EMISION no es válida"
'        Exit Sub
'    End If
'    serieRecibo = Val(Left(lrec.Text, 4))
'    numeroRecibo = Val(Right(lrec.Text, 8))
'    crit = "IDCliente = " & ncon.Text
'    clientes.findLast crit
'    mens = ""
'    If Len(Trim(clientes!cuit)) > 0 Then mens = Left(clientes!cuit, 2) & "-" & Mid(clientes!cuit, 3, 8) & "-" & Right(clientes!cuit, 1)
'    Select Case clientes!sitIva
'        Case 1
'            mens = mens & " R.I."
'        Case 2
'            mens = mens & " R.N.I."
'        Case 3
'            mens = "C.Final"
'        Case 4
'            mens = mens & " iva Exento"
'        Case 5
'            mens = mens & " iva No Resp."
'        Case 6
'            mens = mens & " Monotributo"
'    End Select
'    Select Case clientes!categoria
'        Case 1
'            cate = "General"
'        Case 2
'            cate = "Especial"
'    End Select
'    medi = ""
'    If medidores.RecordCount Then
'        medidores.findLast crit
'        If Not medidores.NoMatch Then medi = medidores!IDMedidor
'    End If
'    operadores.MoveFirst
'    np = 31
'    Do While Mid(operadores!razons, np, 1) <> " "
'        np = np - 1
'    Loop
'    crit = "IDSRe = " & serieRecibo & " and IDRec = " & numeroRecibo
'    recibos.findLast crit
'    impr.Formulas(0) = "numcuo= '" & Right("0000" & serieRecibo, 4) & "-" & Right("00000000" & numeroRecibo, 8) & "'"
'    impr.Formulas(1) = "fecemi= '" & efec.Text & "'"
'    impr.Formulas(2) = "nomcli= '" & clientes!apellido & ", " & clientes!nombre & "'"
'    impr.Formulas(3) = "ubiinm= '" & clientes!icalle & " " & clientes!ipuerta & " " & clientes!ipiso & " " & clientes!idpto & " " & clientes!ilocalidad & " (" & clientes!icodpostal & ")'"
'    impr.Formulas(4) = "nomcat= '" & clientes!nomCat & "'"
'    impr.Formulas(5) = "domcli= '" & clientes!fcalle & " " & clientes!fpuerta & " " & clientes!fpiso & " " & clientes!fdpto & " " & clientes!flocalidad & " (" & clientes!fcodpostal & ")'"
'    impr.Formulas(19) = "numsoc= '" & clientes!nsocio & "'"
'    impr.Formulas(6) = "cuicli= '" & mens & "'"
'    impr.Formulas(7) = "numcli= '" & ncon.Text & "'"
'    impr.Formulas(8) = "catego= '" & cate & "'"
'    impr.Formulas(9) = "nummed= '" & medi & "'"
'    imputados.findFirst crit
'    If imputados!tipco = 1 Then
'        doc = "Factura " & Right("0000" & imputados!IDSCo, 4) & "-" & Right("00000000" & imputados!IDNCo, 8)
'    Else
'        doc = "Cuota " & imputados!IDSCo & "-" & Right("00000" & imputados!IDNCo, 5)
'    End If
'    impr.Formulas(10) = "docimp= '" & doc & "'"
'    impr.Formulas(11) = "impprp= '" & Format(recibos!total, "#,###,##0.00") & "'"
'    impr.Formulas(12) = "implet= 'SON PESOS: " & numlet(recibos!total * 100) & "'"
'    Select Case operadores!sitIva
'        Case 1:
'            mens = "Resp. Inscripto"
'        Case 2:
'            mens = "Resp. No Inscripto"
'        Case 3:
'            mens = "Cons. Final"
'        Case 4:
'            mens = "Exento"
'        Case 5:
'            mens = "No Responsable"
'        Case 6:
'            mens = "Resp. Monotributo"
'    End Select
'    impr.Formulas(13) = "nomope= '" & Mid(operadores!razons, 1, np) & "'"
'    impr.Formulas(14) = "nomop1= '" & Mid(operadores!razons, np + 1) & "'"
'    impr.Formulas(15) = "domope= '" & operadores!calle & " " & operadores!puerta & " " & operadores!piso & " " & operadores!dpto & " - " & operadores!localidad & "'"
'    impr.Formulas(16) = "locope= 'C.P. " & operadores!cpostal & " " & operadores!provincia & " - Tel: " & operadores!telef & "'"
'    impr.Formulas(17) = "opegr1= 'C.U.I.T. NRO: " & operadores!cuit & " ING. BRUTOS: " & operadores!ingBrutos & "'"
'    impr.Formulas(18) = "opegr2= 'I.V.A. " & mens & "   NRO. E.P.A.S. " & operadores!NEpas & "'"
'    impr.Action = 1
'    impr.SelectionFormula = ""

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    rnue.value = True
    rnue_Click
    llenar

End Sub



