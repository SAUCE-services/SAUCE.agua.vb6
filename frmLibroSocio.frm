VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibroSocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Socios"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   7920
   Begin VB.TextBox txtAnho 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picConsumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   9960
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1095
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crpLibro 
      Left            =   6720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "frmLibroSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerar_Click()
Dim librosocio As clsMyALibroSocio
Dim cliente As clsMODCliente
Dim clienteDato As New clsMyAClienteDato
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura
Dim estado As New clsMyAEstado
Dim categoriasocio As New clsMyACategoriaSocio

Dim clienteRep As New clsREPCliente

Dim clientes As Collection
Dim periodos As New Collection
Dim estados As Collection
Dim categoriasocios As Collection

Dim change As Boolean

Dim contador As Integer

Dim fecha As Date

    Me.cmdGenerar.Enabled = False
    Me.MousePointer = 11
    
    Set estados = estado.collectionAll(dbapp)
    Set categoriasocios = categoriasocio.collectionAll(dbapp)
    
    For contador = 1 To 12
        Set periodo = New clsRESTPeriodo
        fecha = CDate("15/" & Format(contador, "00") & "/" & Val(Me.txtAnho.Text))
        periodo.findByFecha fecha
        If periodo.periodoId > 0 Then periodos.add periodo, "k." & contador
    Next

    Set clientes = clienteRep.collectionSociosActivos
    
    Me.prbProgreso.Min = 0
    Me.prbProgreso.Max = clientes.Count
    Me.prbProgreso.value = Me.prbProgreso.Min

    For Each cliente In clientes
        change = False
        clienteDato.clienteId = cliente.clienteId
        clienteDato.findByPrimaryKey dbapp
        
        If cliente.estadoID = 0 Then
            cliente.estadoID = 1
            change = True
        End If
        If cliente.categoriasocioID = 0 Then
            cliente.categoriasocioID = 1
            If (Date - cliente.fechaAlta) / 365.25 > 7 Then cliente.categoriasocioID = 2
            change = True
        End If
        
        If change Then
            Set cliente = clienteRep.save(cliente)
        End If
        
        Set estado = New clsMyAEstado
        If modCollection.collectionExistElement(estados, "k." & cliente.estadoID) Then Set estado = estados("k." & cliente.estadoID)
        Set categoriasocio = New clsMyACategoriaSocio
        If modCollection.collectionExistElement(categoriasocios, "k." & cliente.categoriasocioID) Then Set categoriasocio = categoriasocios("k." & cliente.categoriasocioID)
        
        DoEvents
        Me.prbProgreso.value = Me.prbProgreso.value + 1
        Me.prbProgreso.Refresh
        
        Me.stbEstado.SimpleText = "Procesando SOCIO: " & cliente.numeroSocio & " - " & cliente.apellidonombre
    
        Set librosocio = New clsMyALibroSocio
        
        librosocio.numeroSocio = Val(cliente.numeroSocio)
        librosocio.anho = Val(Me.txtAnho.Text)
        librosocio.nombreApellido = cliente.apellidonombre
        librosocio.domicilio = cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & " " & cliente.inmuebleLocalidad
        librosocio.documento = clienteDato.documento
        librosocio.estado = estado.nombre
        librosocio.edad = 0
        librosocio.categoria = categoriasocio.nombre
        librosocio.ingreso = cliente.fechaAlta
        
        For contador = 1 To 12
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & contador) Then Set periodo = periodos("k." & contador)
            
            If periodo.periodoId > 0 Then
                factura.clienteId = cliente.clienteId
                factura.periodoId = periodo.periodoId
                factura.findByClientePeriodo dbapp
                
                Select Case contador
                    Case 1:
                        librosocio.enero = factura.fechapago
                    Case 2:
                        librosocio.febrero = factura.fechapago
                    Case 3:
                        librosocio.marzo = factura.fechapago
                    Case 4:
                        librosocio.abril = factura.fechapago
                    Case 5:
                        librosocio.mayo = factura.fechapago
                    Case 6:
                        librosocio.junio = factura.fechapago
                    Case 7:
                        librosocio.julio = factura.fechapago
                    Case 8:
                        librosocio.agosto = factura.fechapago
                    Case 9:
                        librosocio.setiembre = factura.fechapago
                    Case 10:
                        librosocio.octubre = factura.fechapago
                    Case 11:
                        librosocio.noviembre = factura.fechapago
                    Case 12:
                        librosocio.diciembre = factura.fechapago
                End Select
            End If
        Next contador
        
        librosocio.save dbapp
    Next
    
    Set periodos = Nothing
    Set clientes = Nothing
    
    Me.MousePointer = 0
    Me.cmdGenerar.Enabled = True
    
    Me.stbEstado.SimpleText = ""

End Sub

Private Sub cmdImprimir_Click()
Dim impresionService As New clsCtlImpresion

Dim operador As New clsMyAOperador

Dim cuit As String
Dim mens As String
    
    Me.cmdImprimir.Enabled = False
    
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

    impresionService.printReport Me.crpLibro, "rptLibroSocio", dbapp.stringConnection, , _
        Array(Array("anho", Val(Me.txtAnho.Text))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas))
        
    Me.cmdImprimir.Enabled = True
        
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.txtAnho.Text = Year(Date)
    
End Sub

Private Sub txtAnho_GotFocus()

    marcarseleccion Me.txtAnho
    
End Sub
