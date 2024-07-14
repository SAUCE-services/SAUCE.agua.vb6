VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSuspServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cortes y Restricciones del Servicio"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   7545
   Begin VB.TextBox copia 
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin Crystal.CrystalReport impr 
      Left            =   2760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Facturación Agua\restriccion.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox ncli 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   5175
   End
   Begin VB.OptionButton corte 
      Caption         =   "Cortes"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton restri 
      Caption         =   "Restricciones"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Fin de la TAREA"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton impri 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Imprime el ACTA correspondiente al CLIENTE Activo"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Copias"
      Height          =   195
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label estado 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmSuspServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub corte_Click()
    
    If corte.value Then llenar_cor

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
    
    restri.value = True
    llenar_res

End Sub

Private Sub impri_Click()
Dim np As Integer
Dim ct As Integer
Dim num As Integer

Dim busq As String
Dim mens As String
Dim tipo As String
Dim cuit As String
Dim numero As String

Dim total As Currency

Dim suspension As New clsMyASuspension
Dim periodo As New clsRESTPeriodo
Dim operador As New clsMyAOperador
Dim cliente As New clsMODCliente
Dim factura As New clsMyAFactura
Dim deuda As New clsMyADeuda

Dim clienterep As clsREPCliente

Dim periodos As Collection

On Error Resume Next
    
    If ncli.ListIndex < 0 Then Exit Sub
    If Val(copia.Text) = 0 Then copia.Text = 1
    If restri.value Then
        impr.ReportFileName = App.path & "\rptRestriccion.rpt"
        tipo = "R"
        numero = "numres= "
    End If
    If corte.value Then
        impr.ReportFileName = App.path & "\rptCorte.rpt"
        tipo = "C"
        numero = "numcor= "
    End If
    
    impr.Destination = crptToPrinter
    num = 1
    suspension.tipo = tipo
    suspension.findLast dbapp
    num = suspension.numero + 1
    
    suspension.tipo = tipo
    suspension.clienteId = ncli.ItemData(ncli.ListIndex)
    suspension.fecha = Date
    suspension.findByClienteID dbapp
    If suspension.autoID > 0 Then num = suspension.numero
    
    periodo.findToday
    
    suspension.tipo = tipo
    suspension.numero = num
    suspension.fecha = Date
    suspension.clienteId = ncli.ItemData(ncli.ListIndex)
    suspension.periodoId = periodo.periodoId + 1
    suspension.uid = "admin"
    suspension.save dbapp
    
    operador.findLast dbapp
    
    Set clienterep = New clsREPCliente
    Set cliente = clienterep.findLastByClienteId(ncli.ItemData(ncli.ListIndex))
    Set clienterep = Nothing
    
    Set periodos = periodo.collectionAll
    
    cuit = Left(operador.cuit, 2) & "-" & Mid(operador.cuit, 3, 8) & "-" & Right(operador.cuit, 1)
    np = 31
    If np > Len(operador.razonSocial) Then
        np = Len(operador.razonSocial)
    Else
        Do While Mid(operador.razonSocial, np, 1) <> " "
            np = np - 1
        Loop
    End If
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
    impr.Formulas(0) = "nomope= '" & Mid(operador.razonSocial, 1, np) & "'"
    impr.Formulas(1) = "nomop1= '" & Mid(operador.razonSocial, np + 1) & "'"
    impr.Formulas(2) = "domope= '" & operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad & "'"
    impr.Formulas(3) = "locope= 'C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono & "'"
    impr.Formulas(4) = "opegr1= 'C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos & "'"
    impr.Formulas(5) = "opegr2= 'I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas & "'"
    impr.Formulas(42) = "resol= '" & operador.resolucion & "'"
    impr.Formulas(43) = "perso= '" & operador.personeria & "'"
    impr.Formulas(6) = "nomcli= '" & cliente.apellidonombre & "'"
    impr.Formulas(7) = "ubiinm= '" & cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & "'"
    impr.Formulas(8) = "nomcat= ''"
    impr.Formulas(9) = "domcli= '" & cliente.fiscalCalle & " " & cliente.fiscalPuerta & " " & cliente.fiscalPiso & " " & cliente.fiscalDpto & "'"
    impr.Formulas(10) = "numsoc= '" & cliente.numeroSocio & "'"
    impr.Formulas(39) = "fecemi= '" & Date & "'"
    impr.Formulas(40) = numero & "'" & num & "'"
    mens = ""
    If Len(Trim(cliente.cuit)) > 0 Then mens = Left(cliente.cuit, 2) & "-" & Mid(cliente.cuit, 3, 8) & "-" & Right(cliente.cuit, 1)
    Select Case cliente.situacionIVA
        Case 1
            mens = mens & " R.I."
        Case 2
            mens = mens & " R.N.I."
        Case 3
            mens = "C. Final"
        Case 4
            mens = mens & " IVA Exento"
        Case 5
            mens = mens & " IVA No Resp."
    End Select
    impr.Formulas(11) = "cuicli= '" & mens & "'"
    impr.Formulas(12) = "numcli= '" & ncli.ItemData(ncli.ListIndex) & "'"
    Select Case cliente.categoria
        Case 1
            mens = "General"
        Case 2
            mens = "Especial"
    End Select
    impr.Formulas(13) = "catego= '" & mens & "'"
    For ct = 1 To 5
        impr.Formulas(13 + ct) = "per(" & ct & ")= ''"
        impr.Formulas(18 + ct) = "fac(" & ct & ")= ''"
        impr.Formulas(23 + ct) = "ven(" & ct & ")= ''"
        impr.Formulas(28 + ct) = "imf(" & ct & ")= ''"
        impr.Formulas(33 + ct) = "iin(" & ct & ")= ''"
    Next ct
    
    total = 0
    ct = 1
    For Each factura In factura.collectionDeudaByClienteID(ncli.ItemData(ncli.ListIndex), dbapp)
        If ct < 6 Then
            Set periodo = periodos("k." & factura.periodoId)
            impr.Formulas(13 + ct) = "per(" & ct & ")= '" & periodo.descripcion & "'"
            impr.Formulas(18 + ct) = "fac(" & ct & ")= '" & Right("0000" & factura.puntoVta, 4) & "-" & Right("00000000" & factura.nroComprob, 8) & "'"
            impr.Formulas(23 + ct) = "ven(" & ct & ")= '" & periodo.fechaPrimero & "'"
            impr.Formulas(28 + ct) = "imf(" & ct & ")= '" & Format(factura.total, "#,###,##0.00") & "'"
            impr.Formulas(33 + ct) = "iin(" & ct & ")= '" & Format(factura.total + interes(factura.total, factura.tasa, factura.fecha, Date), "#,###,##0.00") & "'"
            total = total + factura.total + interes(factura.total, factura.tasa, factura.fecha, Date)
        End If
        ct = ct + 1
    Next
    For Each deuda In deuda.collectionDeudaByClienteID(ncli.ItemData(ncli.ListIndex), dbapp)
        If ct < 6 Then
            impr.Formulas(23 + ct) = "ven(" & ct & ")= '" & deuda.cuotas - deuda.cuotasPagadas & " cuota(s)'"
            impr.Formulas(33 + ct) = "iin(" & ct & ")= '" & Format(deuda.deuda, "#,###,##0.00") & "'"
            total = total + deuda.deuda
        End If
        ct = ct + 1
    Next
    impr.Formulas(41) = "total= '" & Format(total, "#,###,##0.00") & "'"
    For ct = 1 To Val(copia.Text)
        estado.Caption = "Imprimiendo Acta . . ."
        estado.Refresh
        impr.Action = 1
    Next ct
    impr.SelectionFormula = ""

End Sub

Private Sub llenar_res()
Dim periodo As New clsRESTPeriodo

Dim clires As New clsVMyACliRes

On Error GoTo agreg
    
    ncli.Clear
    
    If periodo.collectionAll.Count = 0 Then Exit Sub
    
    estado.Caption = "Cargando posibles RESTRICCIONES . . ."
    estado.Refresh
    
    For Each clires In clires.collectionAll(dbapp)
        ncli.Text = clires.apellido & ", " & clires.nombre
    Next
    If ncli.ListCount Then ncli.ListIndex = 0
final:
    estado.Caption = ""
    Exit Sub
agreg:
    If Err.Number = 383 Then
        ncli.AddItem clires.apellido & ", " & clires.nombre
        ncli.ItemData(ncli.NewIndex) = clires.clienteId
    Else
        MsgBox "Ocurrió un error que no puedo resolver" & Chr(13) & Chr(13) & "ERROR : " & Err.Number & " - Por favor contáctese con el Servicio Técnico"
        Resume final
    End If
    Resume Next

End Sub

Private Sub llenar_cor()
Dim periodo As New clsRESTPeriodo

Dim clicor As New clsVMyACliCor

On Error GoTo agreg
    
    ncli.Clear
    
    If periodo.collectionAll.Count = 0 Then Exit Sub
    
    estado.Caption = "Cargando posibles CORTES . . ."
    estado.Refresh
    
    For Each clicor In clicor.collectionAll(dbapp)
        ncli.Text = clicor.apellido & ", " & clicor.nombre
    Next
    If ncli.ListCount Then ncli.ListIndex = 0
final:
    estado.Caption = ""
    Exit Sub
agreg:
    If Err.Number = 383 Then
        ncli.AddItem clicor.apellido & ", " & clicor.nombre
        ncli.ItemData(ncli.NewIndex) = clicor.clienteId
    Else
        MsgBox "Ocurrió un error que no puedo resolver" & Chr(13) & Chr(13) & "ERROR : " & Err.Number & " - Por favor contáctese con el Servicio Técnico"
        Resume final
    End If
    Resume Next
    
End Sub

Private Sub restri_Click()
    
    If restri.value Then llenar_res

End Sub
