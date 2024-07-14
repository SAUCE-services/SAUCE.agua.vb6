VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMedidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Medidores"
   ClientHeight    =   8370
   ClientLeft      =   2715
   ClientTop       =   2640
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10350
   Begin MSFlexGridLib.MSFlexGrid grdMedidor 
      Height          =   5655
      Left            =   480
      TabIndex        =   48
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   47
      ToolTipText     =   "Fin de la TAREA"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton elim 
      Caption         =   "E&liminar Medidor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      ToolTipText     =   "Elimina un MEDIDOR que no tiene movimientos"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton inic 
      Caption         =   "&Estado Medidor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      ToolTipText     =   "Poner el ESTADO ACTUAL del Medidor"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton retb 
      Caption         =   "&Retirar Medidor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      ToolTipText     =   "Desasociar el MEDIDOR actual de su CONEXION"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton colb 
      Caption         =   "&Colocar Medidor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      ToolTipText     =   "Asociar el MEDIDOR seleccionado con una CONEXION"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton agrb 
      Caption         =   "&Agregar Medidor"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      ToolTipText     =   "Cargar los DATOS de un MEDIDOR Nuevo"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton hisb 
      Caption         =   "&Ver Histórico"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      ToolTipText     =   "Ver la Evolución del MEDIDOR"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame mcol 
      Caption         =   "Colocar MEDIDOR"
      Height          =   2175
      Left            =   240
      TabIndex        =   24
      Top             =   6120
      Width           =   9855
      Begin VB.TextBox cest 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox cfec 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox mcon 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cconf 
         Caption         =   "Co&nfirmar"
         Height          =   375
         Left            =   7920
         TabIndex        =   13
         ToolTipText     =   "Confirma la Colocación del MEDIDOR"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox ccon 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Estado del Medidor"
         Height          =   195
         Index           =   15
         Left            =   2160
         TabIndex        =   46
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Colocación"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   28
         Top             =   840
         Width           =   555
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de Conexión"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de MEDIDOR"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label cnum 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame mini 
      Caption         =   "Estado del Medidor"
      Height          =   1095
      Left            =   240
      TabIndex        =   40
      Top             =   6120
      Width           =   9855
      Begin VB.CommandButton iconf 
         Caption         =   "C&onfirmar"
         Height          =   375
         Left            =   7920
         TabIndex        =   43
         ToolTipText     =   "Graba el Estado Inicial del MEDIDOR"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox mein 
         Height          =   285
         Left            =   2160
         TabIndex        =   42
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label nmed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de MEDIDOR"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Estado del Medidor"
         Height          =   195
         Index           =   13
         Left            =   2160
         TabIndex        =   44
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame magr 
      Caption         =   "Agregar MEDIDOR"
      Height          =   1095
      Left            =   240
      TabIndex        =   21
      Top             =   6120
      Width           =   9855
      Begin VB.TextBox eini 
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton acon 
         Caption         =   "C&onfirmar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         TabIndex        =   16
         ToolTipText     =   "Graba los DATOS del MEDIDOR Nuevo"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox anum 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Estado Inicial"
         Height          =   195
         Index           =   12
         Left            =   2160
         TabIndex        =   39
         Top             =   240
         Width           =   945
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de MEDIDOR"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame mhis 
      Caption         =   "Histórico de MEDIDORES"
      Height          =   2175
      Left            =   240
      TabIndex        =   18
      Top             =   6120
      Width           =   9855
      Begin VB.ListBox hdef 
         Height          =   840
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   9375
      End
      Begin VB.Label mnum 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de MEDIDOR"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame mret 
      Caption         =   "Retirar MEDIDOR"
      Height          =   2175
      Left            =   240
      TabIndex        =   30
      Top             =   6120
      Width           =   9855
      Begin VB.ComboBox rmot 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton rconf 
         Caption         =   "Co&nfirmar"
         Height          =   375
         Left            =   8040
         TabIndex        =   8
         ToolTipText     =   "Confirma el RETIRO del MEDIDOR"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox rfec 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Motivo de RETIRO"
         Height          =   195
         Index           =   11
         Left            =   2160
         TabIndex        =   38
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label rcli 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label rcon 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label rnum 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de MEDIDOR"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de Conexión"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   32
         Top             =   840
         Width           =   480
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de RETIRO"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Lista de MEDIDORES"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMedidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsin As Integer
Dim vChange As Boolean

Public Sub fillGrid()
Dim medidor As New clsMyAMedidor
Dim medidorB As New clsMyAMedidor
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim clientes As Collection

    Set clientes = clienteRep.collectionActivos

    Me.grdMedidor.Rows = 1
    Me.grdMedidor.Redraw = False
    For Each medidor In medidor.collectionAll(dbapp)
        Set cliente = New clsMODCliente
        If medidorB.medidorID <> medidor.medidorID Then
            If medidorB.medidorID <> "" Then
                If modCollection.collectionExistElement(clientes, "k." & medidorB.clienteId) Then Set cliente = clientes("k." & medidorB.clienteId)
                With medidorB
                    Me.grdMedidor.AddItem modGrid.array2itemGrid(Array(.medidorID, .clienteId, cliente.apellidonombre, .fechaColocacion, .fechaRetiro))
                End With
            End If
        End If
        Set medidorB = medidor
    Next
    If medidorB.medidorID <> "" Then
        With medidorB
            Me.grdMedidor.AddItem modGrid.array2itemGrid(Array(.medidorID, .clienteId, .fechaColocacion, .fechaRetiro))
        End With
    End If
    Me.grdMedidor.Redraw = True
    
End Sub

Private Sub acon_Click()
Dim medidor As New clsMyAMedidor

    If Len(Trim(anum.Text)) = 0 Then Exit Sub
    
    With medidor
        .medidorID = Trim(anum.Text)
        .findByMedidorID dbapp
        
        If .autoID > 0 Then
            MsgBox "Ya Existe . . ."
            Exit Sub
        End If
    End With
    
    ' Agrega medidor sistema nuevo
    With medidor
        .medidorID = Trim(anum.Text)
        .estadoInicio = Val(eini.Text)
        .fechaAlta = Now
        .uid = "admin"
        
        .add dbapp
    End With
    
    vChange = False
    
    fillGrid
    
    Me.Height = rsin
    mini.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mhis.Visible = False
    hisb.Enabled = False
    colb.Enabled = False
    retb.Enabled = False
    inic.Enabled = False

End Sub

Private Sub agrb_Click()
Dim res As Integer

    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    
    magr.Visible = True
    mhis.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mini.Visible = False
    anum.Text = ""
    eini.Text = 0
    Me.Height = rsin + 1100
    vChange = False
    anum.SetFocus

End Sub

Private Sub anum_Change()
    
    vChange = True

End Sub

Private Sub anum_GotFocus()

    marcarseleccion Me.anum

End Sub

Private Sub anum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then eini.SetFocus

End Sub

Private Sub anum_LostFocus()
    
    If Len(Trim(anum.Text)) > 0 Then acon.Enabled = True

End Sub

Private Sub ccon_Change()
    
    vChange = True

End Sub

Private Sub ccon_Click()

On Error Resume Next
    
    mcon.Text = ccon.ItemData(ccon.ListIndex)

End Sub

Private Sub cconf_Click()
Dim fech As Variant

Dim medidor As New clsMyAMedidor

On Error Resume Next
    
    If Not IsDate(cfec.Text) Then
        MsgBox "La Fecha de COLOCACION no es válida"
        Exit Sub
    End If
    
    ' Sistema nuevo
    medidor.medidorID = Trim(cnum.Caption)
    medidor.findByMedidorID dbapp
    fech = medidor.fechaAlta
    If medidor.clienteId > 0 Then If medidor.fechaAlta <> Now Then fech = Now
    
    Set medidor = New clsMyAMedidor
    medidor.medidorID = Trim(cnum.Caption)
    medidor.fechaAlta = fech
    medidor.clienteId = mcon.Text
    medidor.fechaColocacion = CDate(cfec.Text)
    medidor.estadoInicio = Val(cest.Text)
    medidor.uid = "admin"
    medidor.save dbapp
    
    vChange = False
    
    fillGrid
    
    mini.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mhis.Visible = False
    Me.Height = rsin
    hisb.Enabled = False
    colb.Enabled = False
    retb.Enabled = False
    inic.Enabled = False

End Sub

Private Sub cest_Change()
    
    vChange = True

End Sub

Private Sub cest_GotFocus()
    
    marcarseleccion Me.cest

End Sub

Private Sub cfec_Change()
    
    vChange = True

End Sub

Private Sub cfec_GotFocus()

    marcarseleccion Me.cfec

End Sub

Private Sub cfec_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cconf.SetFocus

End Sub

Private Sub cfec_LostFocus()
    
    If Not IsDate(cfec.Text) Then MsgBox "La Fecha de COLOCACION no es válida"

End Sub

Private Sub cmdSalir_Click()

    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    mini.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mhis.Visible = False
    Me.Height = rsin
    
    Unload Me

End Sub

Private Sub colb_Click()
Dim ct As Integer
Dim agr As Boolean

Dim cliente As clsMODCliente
Dim medidor As New clsMyAMedidor

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    
    If grdMedidor.row < 0 Then Exit Sub
    
    If clienteRep.collectionActivos.Count = 0 Then
        MsgBox "No hay conexiones disponibles"
        vChange = False
        Exit Sub
    End If
    
    grdMedidor.col = 0
    cnum.Caption = grdMedidor.Text
    
    medidor.medidorID = Trim(cnum.Caption)
    medidor.findByMedidorID dbapp
    If Not IsNull(medidor.fechaColocacion) And IsNull(medidor.fechaRetiro) Then
        MsgBox "No puede asignar el mismo me a dos conexiones"
        Exit Sub
    End If
    mhis.Visible = False
    magr.Visible = False
    mret.Visible = False
    mcol.Visible = True
    cest.Text = medidor.estadoInicio
    Me.Height = rsin + 2300
    ct = 0
    ccon.Clear
    
    For Each cliente In clienteRep.collectionActivosMedibles
        medidor.clienteId = cliente.clienteId
        medidor.findByClienteID dbapp
        agr = False
        If medidor.medidorID = "" Then
            agr = True
        Else
            If Not IsNull(medidor.fechaRetiro) Then agr = True
        End If
        If agr Then
            ccon.AddItem cliente.apellidonombre
            ccon.ItemData(ccon.NewIndex) = cliente.clienteId
            If ct = 0 Then
                mcon.Text = cliente.clienteId
                ct = ct + 1
            End If
        End If
    Next
    If ccon.ListCount < 1 Then
        mcol.Visible = False
        Me.Height = rsin
        MsgBox "No hay clientes disponibles"
        vChange = False
        hisb.Enabled = False
        colb.Enabled = False
        retb.Enabled = False
        inic.Enabled = False
        Exit Sub
    End If
    ccon.ListIndex = 0
    cfec.Text = Date
    vChange = False
    mcon.SetFocus

End Sub

Private Sub eini_Change()
    
    vChange = True

End Sub

Private Sub eini_GotFocus()

    marcarseleccion Me.eini

End Sub

Private Sub eini_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then acon.SetFocus

End Sub

Private Sub elim_Click()
Dim lectura As New clsMyALectura
Dim medidor As New clsMyAMedidor

On Error Resume Next
    
    grdMedidor.col = 0
    lectura.medidorID = Trim(grdMedidor.Text)
    lectura.findLast dbapp
    If lectura.periodoId <> 0 Then
        MsgBox "No puede ELIMINAR este medidor"
        Exit Sub
    End If
    
    medidor.medidorID = Trim(grdMedidor.Text)
    medidor.findByMedidorID dbapp
    If medidor.autoID > 0 Then
        If MsgBox("Está seguro que desea eliminar este medidor ?" & Chr(13) & Chr(13) & "Medidor : " & medidor.medidorID, vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            For Each medidor In medidor.collectionByMedidorID(Trim(grdMedidor.Text), dbapp)
                medidor.delete dbapp
            Next
        End If
    End If
    
    fillGrid
    
End Sub

Private Sub Form_Activate()

    fillGrid
    Me.Height = rsin
    mhis.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mini.Visible = False

End Sub

Private Sub Form_Load()

    modGrid.makeGrid Me.grdMedidor, Array(Array("Medidor", 1300), Array("Conexion", 800), Array("Cliente", 3000), Array("Colocación", 1000), Array("Retiro", 1000)), 0, 1, flexSelectionFree
    rsin = Me.cmdSalir.Top + Me.cmdSalir.Height + 500
    rmot.Clear
    rmot.AddItem "Rotura"
    rmot.AddItem "Obsolescencia"
    rmot.AddItem "Planta de Prueba"
    rmot.AddItem "Mal Funcionamiento"
    rmot.ListIndex = 0
    
End Sub

Private Sub grdMedidor_SelChange()

    grdMedidor_Click
    
End Sub

Private Sub hisb_Click()
Dim moti As String

Dim medidor As New clsMyAMedidor

On Error Resume Next
    
    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    If grdMedidor.row < 0 Then Exit Sub
    mhis.Visible = True
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mini.Visible = False
    Me.Height = rsin + 2200
    grdMedidor.col = 0
    mnum.Caption = grdMedidor.Text
    hdef.Clear
    For Each medidor In medidor.collectionByMedidorID(Trim(mnum.Caption), dbapp)
        moti = " - Motivo : "
        Select Case medidor.motivoRetiro
            Case 1
                moti = moti & "Rotura"
            Case 2
                moti = moti & "Obsolescencia"
            Case 3
                moti = moti & "Prueba"
            Case 4
                moti = moti & "Mal Funcionamiento"
            Case Else
                moti = ""
        End Select
        hdef.AddItem "Conexion : " & Right("0000" & medidor.clienteId, 4) & " - F.Coloc : " & Left(medidor.fechaColocacion, 10) & " - F.Retiro : " & Left(medidor.fechaRetiro, 10) & moti
    Next

End Sub

Private Sub iconf_Click()
Dim medidor As New clsMyAMedidor

    With medidor
        .medidorID = Trim(nmed.Caption)
        .findByMedidorID dbapp
        
        .estadoInicio = Val(mein.Text)
    
        vChange = False
        
        .uid = "admin"
        
        .save dbapp
    End With
    Me.Height = rsin
    mini.Visible = False

End Sub

Private Sub inic_Click()
Dim medidor As New clsMyAMedidor

    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    
    If grdMedidor.row < 0 Then Exit Sub
    
    mhis.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mini.Visible = True
    Me.Height = rsin + 1100
    grdMedidor.col = 0
    nmed.Caption = grdMedidor.Text
    
    medidor.medidorID = nmed.Caption
    medidor.findByMedidorID dbapp
    mein.Text = medidor.estadoInicio
    If medidor.clienteId = 0 Or Not IsNull(medidor.fechaRetiro) Then
        iconf.Enabled = True
        mein.Enabled = True
        mein.SetFocus
    Else
        iconf.Enabled = False
        mein.Enabled = False
    End If
    vChange = False

End Sub

Private Sub mcon_Change()
    
    vChange = True

End Sub

Private Sub mcon_GotFocus()

    marcarseleccion mcon
    
End Sub

Private Sub mcon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cfec.SetFocus

End Sub

Private Sub mcon_LostFocus()
Dim cli As Long
Dim crit As String
Dim vue As Integer

Dim medidor As New clsMyAMedidor
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    medidor.clienteId = Val(mcon.Text)
    medidor.findByClienteID dbapp, False
    If IsNull(medidor.fechaRetiro) And medidor.medidorID <> "" Then
        MsgBox "No puede asignar dos medidores a la misma conexion a la vez"
        mcon.Text = Val(Left(ccon.Text, 6))
        Exit Sub
    End If
    Set cliente = clienteRep.findLastByClienteID(Val(Me.mcon.Text))
    If IsNull(cliente.uniqueId) Then
        mcon.Text = 0
        Exit Sub
    End If
    If Not IsNull(cliente.fechaBaja) Then
        MsgBox "Cliente Dado de Baja . . ."
        mcon.Text = 0
        Exit Sub
    End If
    mcon.Text = cliente.clienteId
    cli = Val(mcon.Text)
    ccon.Text = cliente.apellidonombre
    vue = 0
    Do While ccon.ItemData(ccon.ListIndex) <> cli And vue < 5000
        ccon.ListIndex = ccon.ListIndex + 1
        vue = vue + 1
    Loop
    If Err.Number = 383 Then
        MsgBox "Este usuario ya tiene medidor o no tiene servicio medido . . ."
        mcon.Text = ccon.ItemData(ccon.ListIndex)
    End If

End Sub

Private Sub grdMedidor_Click()
Dim nmed As String

Dim medidor As New clsMyAMedidor

On Error Resume Next
    
    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    Me.Height = rsin
    mhis.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mini.Visible = False
    grdMedidor.col = 0
    nmed = grdMedidor.Text
    medidor.medidorID = Trim(nmed)
    medidor.findByMedidorID dbapp
    hisb.Enabled = True
    inic.Enabled = True
    colb.Enabled = True
    elim.Enabled = True
    retb.Enabled = False
    If Not IsNull(medidor.fechaColocacion) And IsNull(medidor.fechaRetiro) Then
        colb.Enabled = False
        retb.Enabled = True
    End If

End Sub

Private Sub mein_Change()
    
    vChange = True

End Sub

Private Sub mein_GotFocus()

    marcarseleccion Me.mein

End Sub

Private Sub mein_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then iconf.SetFocus

End Sub

Private Sub rconf_Click()
Dim medidor As New clsMyAMedidor

On Error Resume Next
    
    If Not IsDate(rfec.Text) Then
        MsgBox "La Fecha de RETIRO no es válida"
        Exit Sub
    End If
    
    medidor.medidorID = Trim(rnum.Caption)
    medidor.findByMedidorID dbapp
    medidor.fechaRetiro = CDate(rfec.Text)
    medidor.motivoRetiro = rmot.ListIndex + 1
    medidor.uid = "admin"
    medidor.update dbapp
    
    vChange = False
    
    fillGrid
    
    mini.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = False
    mhis.Visible = False
    Me.Height = rsin
    hisb.Enabled = False
    colb.Enabled = False
    retb.Enabled = False
    inic.Enabled = False

End Sub

Private Sub retb_Click()
Dim medidor As New clsMyAMedidor
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If vChange Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vChange = False
        Else
            Exit Sub
        End If
    End If
    If grdMedidor.row < 0 Then Exit Sub
    grdMedidor.col = 0
    rnum.Caption = grdMedidor.Text
    medidor.medidorID = Trim(rnum.Caption)
    medidor.findByMedidorID dbapp
    If IsNull(medidor.fechaColocacion) Then
        MsgBox "No puede retirar un medidor que no ha sido colocado"
        Exit Sub
    End If
    mhis.Visible = False
    magr.Visible = False
    mcol.Visible = False
    mret.Visible = True
    mini.Visible = False
    Me.Height = rsin + 2300
    rcon.Caption = medidor.clienteId
    Set cliente = clienteRep.findLastByClienteID(Val(Me.rcon.Caption))
    rcli.Caption = cliente.comboText
    rfec.Text = Date
    vChange = False
    rfec.SetFocus

End Sub

Private Sub rfec_Change()
    
    vChange = True

End Sub

Private Sub rfec_GotFocus()

    marcarseleccion rfec

End Sub

Private Sub rfec_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then rconf.SetFocus

End Sub

Private Sub rfec_LostFocus()
    
    If Not IsDate(rfec.Text) Then MsgBox "La Fecha de RETIRO no es válida"

End Sub

Private Sub rmot_Change()
    
    vChange = True

End Sub
