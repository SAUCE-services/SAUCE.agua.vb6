VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmNCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de Crédito"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7980
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin Crystal.CrystalReport crpNota 
      Left            =   5520
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox fcli 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton nimpr 
      Caption         =   "&Imprimir N. Crédito"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      ToolTipText     =   "Imprime una Nota de Crédito Anterior"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton anul 
      Caption         =   "&Anular N. Crédito"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "ANULA la Nota de Crédito Consultada"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton gncr 
      Caption         =   "&Generar e Imprimir"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Genera e Imprime la Nota de Crédito"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox efec 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox impncr 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox lncr 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame tncr 
      Caption         =   "Nota de Crédito"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton nant 
         Caption         =   "An&terior"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Consulta una Nota de Crédito Anterior"
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton nnue 
         Caption         =   "&Nueva"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Realiza una Nota de Crédito Nueva"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Factura Asociada"
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   17
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   16
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   480
   End
   Begin VB.Label nroncr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de N. Crédito"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   12
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "frmNCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ncredito_serie As Integer
Private ncredito_numero As Long

Private cliente As New clsMODCliente

Public Sub llenar_fac()
Dim factura As New clsMyAFactura

Dim facturas As Collection

On Error Resume Next
    
    Set facturas = factura.collectionByClienteId(cliente.clienteId, dbapp, True)
    
    If facturas.Count = 0 Then
        MsgBox "Este cliente no tiene Facturas . . ."
        gncr.Enabled = False
        Exit Sub
    End If
    fcli.Clear
    For Each factura In facturas
        With factura
            fcli.AddItem Right("0000" & .puntoVta, 4) & "-" & Right("00000000" & .nroComprob, 8)
        End With
    Next
    fcli.ListIndex = 0
    gncr.Enabled = True

End Sub

Public Sub impri_cred(serieId As Integer, numero As Long)
Dim Mensaje As String
Dim cate As String
Dim medi As String
Dim doc As String

Dim np As Integer

Dim medidor As New clsMyAMedidor
Dim operador As New clsMyAOperador
Dim ncredito As New clsMyANCredito

Dim clienteRep As New clsREPCliente

Dim impresionService As New clsCtlImpresion

On Error Resume Next
    
    If Not IsDate(efec.Text) Then
        MsgBox "La Fecha de EMISION no es válida"
        Exit Sub
    End If
    
    Set cliente = clienteRep.findLastByClienteId(cliente.clienteId)
    Mensaje = ""
    If Len(Trim(cliente.cuit)) > 0 Then Mensaje = Left(cliente.cuit, 2) & "-" & Mid(cliente.cuit, 3, 8) & "-" & Right(cliente.cuit, 1)
    Select Case cliente.situacionIVA
        Case 1
            Mensaje = Mensaje & " R.I."
        Case 2
            Mensaje = Mensaje & " R.N.I."
        Case 3
            Mensaje = "C.Final"
        Case 4
            Mensaje = Mensaje & " iva Exento"
        Case 5
            Mensaje = Mensaje & " iva No Resp."
        Case 6
            Mensaje = Mensaje & " Monotributo"
    End Select
    Select Case cliente.categoria
        Case 1
            cate = "General"
        Case 2
            cate = "Especial"
    End Select
    medi = ""
    medidor.clienteId = cliente.clienteId
    medidor.findByClienteId dbapp
    If medidor.autoID > 0 Then medi = medidor.medidorID
    
    operador.findLast dbapp
    np = 31
    Do While Mid(operador.razonSocial, np, 1) <> " "
        np = np - 1
    Loop
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
        Case 6:
            Mensaje = "Resp. Monotributo"
    End Select
    
    ncredito.serieId = serieId
    ncredito.numero = numero
    ncredito.findByPrimaryKey dbapp
    doc = "Factura " & Right("0000" & ncredito.puntoVta, 4) & "-" & Right("00000000" & ncredito.nroComprob, 8)
    
    impresionService.printReport Me.crpNota, "rptNCredito", dbapp.stringConnection, , , , , Array( _
        Array("numcuo", Right("0000" & serieId, 4) & "-" & Right("00000000" & numero, 8)), _
        Array("fecemi", efec.Text), _
        Array("nomcli", cliente.apellidonombre), _
        Array("ubiinm", cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & " " & cliente.inmuebleLocalidad & " (" & cliente.inmuebleCodpostal & ")"), _
        Array("nomcat", ""), _
        Array("domcli", cliente.fiscalCalle & " " & cliente.fiscalPuerta & " " & cliente.fiscalPiso & " " & cliente.fiscalDpto & " " & cliente.fiscalLocalidad & " (" & cliente.fiscalCodpostal & ")"), _
        Array("cuicli", Mensaje), _
        Array("numcli", cliente.clienteId), _
        Array("catego", cate), _
        Array("nummed", medi), _
        Array("docimp", doc), _
        Array("impprp", Format(ncredito.total, "#,###,##0.00")), _
        Array("implet", "SON PESOS: " & num2letras(ncredito.total)), _
        Array("nomope", Mid(operador.razonSocial, 1, np)), _
        Array("nomop1", Mid(operador.razonSocial, np + 1)), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & operador.cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & Mensaje & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("numsoc", cliente.numeroSocio))

End Sub

Private Sub anul_Click()
Dim ncredito As New clsMyANCredito

On Error Resume Next
    
    ncredito.serieId = ncredito_serie
    ncredito.numero = ncredito_numero
    ncredito.findByPrimaryKey dbapp
    
    ncredito.anulado = 1
    ncredito.save dbapp
    
    nnue.value = True
    nnue_Click

End Sub

Private Sub efec_GotFocus()
    
    efec.SelStart = 0
    efec.SelLength = Len(efec.Text)

End Sub

Private Sub fin_Click()
    
    Unload Me
    
End Sub

Private Sub gncr_Click()
Dim factura_serie As Integer
Dim factura_numero As Long

Dim ncredito As New clsMyANCredito
Dim factura As New clsMyAFactura

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If Not IsDate(efec.Text) Then
        MsgBox "Fecha NO Válida . . ."
        efec.SetFocus
        Exit Sub
    End If
    If Len(Trim(impncr.Text)) = 0 Then
        MsgBox "Debe especificar un IMPORTE . . ."
        impncr.SetFocus
        Exit Sub
    End If
    If CDbl(impncr.Text) <= 0 Then
        MsgBox "El importe debe ser mayor que cero . . ."
        impncr.SetFocus
        Exit Sub
    End If
    
    If clienteRep.collectionActivos.Count = 0 Then
        MsgBox "No Existen CLIENTES . . ."
        Unload Me
        Exit Sub
    End If
    
    If fcli.ListCount = 0 Then
        MsgBox "Este cliente no tiene Facturas . . ."
        Exit Sub
    End If
    factura_serie = Val(Left(fcli.Text, 4))
    factura_numero = Val(Right(fcli.Text, 8))
    factura.puntoVta = factura_serie
    factura.nroComprob = factura_numero
    factura.findByPrimaryKey dbapp
    
    ncredito.serieId = ncredito_serie
    ncredito.numero = ncredito_numero
    ncredito.fecha = CDate(efec.Text)
    ncredito.clienteId = cliente.clienteId
    ncredito.situacionIVA = cliente.situacionIVA
    ncredito.total = CDbl(impncr.Text)
    ncredito.puntoVta = factura_serie
    ncredito.nroComprob = factura_numero
    ncredito.ivacf = CDbl(impncr.Text) * factura.ivacf / factura.total
    ncredito.ivari = CDbl(impncr.Text) * factura.ivari / factura.total
    ncredito.ivarn = CDbl(impncr.Text) * factura.ivarn / factura.total
    ncredito.uid = "admin"
    ncredito.save dbapp
    
    impri_cred ncredito_serie, ncredito_numero
    nnue_Click

End Sub

Private Sub impncr_GotFocus()
    
    impncr.SelStart = 0
    impncr.SelLength = Len(impncr.Text)

End Sub

Private Sub lncr_Click()
Dim ncredito As New clsMyANCredito

On Error Resume Next
    
    ncredito_serie = Val(Left(lncr.Text, 4))
    ncredito_numero = Val(Right(lncr.Text, 8))
    ncredito.serieId = ncredito_serie
    ncredito.numero = ncredito_numero
    
    ncredito.findByPrimaryKey dbapp
    
    impncr.Text = Format(ncredito.total, "#,###,##0.00")
    efec.Text = ncredito.fecha
    fcli.Enabled = False
    fcli.Text = Right("0000" & ncredito.puntoVta, 4) & "-" & Right("00000000" & ncredito.nroComprob, 8)

End Sub

Private Sub Form_Activate()
    
    efec.Text = Date

End Sub

Private Sub nant_Click()
Dim ic As Integer

Dim ncredito As New clsMyANCredito

Dim clienteRep As New clsREPCliente

Dim ncreditos As Collection

On Error Resume Next
    
    nroncr.Visible = False
    efec.Enabled = False
    impncr.Enabled = False
    gncr.Enabled = False
    anul.Enabled = True
    nimpr.Enabled = True
    lncr.Visible = True
    lncr.Clear
    
    Set ncreditos = ncredito.collectionActivasByClienteID(cliente.clienteId, dbapp)
    
    If ncreditos.Count = 0 Then
        Set ncreditos = Nothing
        
        MsgBox "No hay NOTAS DE CREDITO para este cliente . . ."
        nnue.value = True
        nnue_Click
        Exit Sub
    End If
    
    ic = 0
    For Each ncredito In ncreditos
        With ncredito
            If ic = 0 Then
                ncredito_serie = .serieId
                ncredito_numero = .numero
            End If
            ic = ic + 1
            lncr.AddItem Right("0000" & .serieId, 4) & "-" & Right("00000000" & .numero, 8)
        End With
    Next
    
    lncr.ListIndex = 0
    ncredito.serieId = ncredito_serie
    ncredito.numero = ncredito_numero
    ncredito.findByPrimaryKey dbapp
    
    impncr.Text = Format(ncredito.total, "#,###,##0.00")
    efec.Text = ncredito.fecha

End Sub

Private Sub nimpr_Click()

On Error Resume Next
    
    If lncr.ListIndex < -1 Then
        MsgBox "Debe seleccionar una Nota de Crédito . . ."
        lncr.SetFocus
        Exit Sub
    End If
    impri_cred Val(Left(lncr.Text, 4)), Val(Right(lncr.Text, 8))

End Sub

Private Sub nnue_Click()
Dim crit As String

Dim operador As New clsMyAOperador
Dim ncredito As New clsMyANCredito

On Error Resume Next
    
    efec.Enabled = True
    gncr.Enabled = True
    fcli.Enabled = True
    anul.Enabled = False
    nimpr.Enabled = False
    lncr.Visible = False
    nroncr.Visible = True
    impncr.Enabled = True
    ncredito_serie = 1
    ncredito_numero = 1
    
    operador.findLast dbapp
    If operador.uid <> "" Then
        If Not IsNull(operador.ncreditoSerie) Then ncredito_serie = operador.ncreditoSerie
        If Not IsNull(operador.ncredito) Then ncredito_numero = operador.ncredito
    End If
    
    If ncredito.collectionAny(dbapp).Count > 0 Then
        ncredito.findLastLast dbapp
        If ncredito_serie < ncredito.serieId Then ncredito_serie = ncredito.serieId
        
        ncredito.serieId = ncredito_serie
        ncredito.findLast dbapp
        
        If ncredito.autoID > 0 Then If ncredito_numero <= ncredito.numero Then ncredito_numero = ncredito.numero + 1
    End If
    nroncr.Caption = Right("0000" & ncredito_serie, 4) & "-" & Right("00000000" & ncredito_numero, 8)
    impncr.Text = Format(0, "#,###,##0.00")
    efec.Text = Date

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    nnue.value = True
    nnue_Click
    llenar_fac

End Sub


