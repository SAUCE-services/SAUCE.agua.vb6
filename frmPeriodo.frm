VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Períodos de Facturación"
   ClientHeight    =   4950
   ClientLeft      =   2760
   ClientTop       =   3030
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6150
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Fin de la TAREA"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton elim 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Elimina el PERIODO de Facturacion seleccionado"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton mper 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Modifica el PERIODO de Facturación seleccionado"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox pdef 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton aper 
      Caption         =   "&Agregar Período"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "Permite definir un PERIODO de Facturación Nuevo"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame per 
      Caption         =   "Agregar Período"
      Height          =   2775
      Left            =   240
      TabIndex        =   23
      Top             =   2040
      Width           =   5655
      Begin MSComCtl2.DTPicker dtpSegundo_add 
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpPrimer_add 
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpFin_add 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpInicio_add 
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin VB.TextBox aley 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   5175
      End
      Begin VB.ComboBox perd 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton agre 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   "Graba los DATOS del Nuevo PERIODO"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox tasa 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox desc 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Leyenda de la Factura"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Período de la Tasa"
         Height          =   195
         Index           =   13
         Left            =   2040
         TabIndex        =   39
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de Interés (%)"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Segundo Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   28
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Primer Vencimiento"
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   27
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   840
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Finalización"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame vper 
      Caption         =   "Modificar Período"
      Height          =   2775
      Left            =   240
      TabIndex        =   31
      Top             =   2040
      Width           =   5655
      Begin MSComCtl2.DTPicker dtpSegundo_upd 
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpPrimer_upd 
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpFin_upd 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin MSComCtl2.DTPicker dtpInicio_upd 
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   43976
      End
      Begin VB.TextBox mley 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   5175
      End
      Begin VB.ComboBox mprd 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox mdes 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox mtas 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton modi 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "Graba las modificaciones del PERIODO"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Leyenda de la Factura"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   41
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Período de la Tasa"
         Height          =   195
         Index           =   14
         Left            =   2040
         TabIndex        =   38
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Finalización"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   840
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Primer Vencimiento"
         Height          =   195
         Index           =   9
         Left            =   2040
         TabIndex        =   34
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Segundo Vencimiento"
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   33
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de Interés (%)"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   1365
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Períodos DEFINIDOS"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "frmPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ultima_fecha As Date
Private form_height As Integer
Private modify As Boolean

Private Sub llenar()
Dim periodo As New clsRESTPeriodo

On Error Resume Next

    pdef.Clear
    ultima_fecha = Date
    
    For Each periodo In periodo.collectionAll
        pdef.AddItem periodo.periodoId & " - " & periodo.descripcion & "  (" & periodo.fechainicio & " - " & periodo.fechafin & ")"
        pdef.ItemData(pdef.NewIndex) = periodo.periodoId
        If ultima_fecha = Date Then ultima_fecha = periodo.fechafin + 1
        If ultima_fecha < periodo.fechafin + 1 Then ultima_fecha = periodo.fechafin + 1
    Next

End Sub

Private Sub agre_Click()
Dim nuevo_periodoId As Integer
Dim tasa As Currency

Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    If Len(Trim(desc.Text)) < 1 Then
        MsgBox "Debe poner la DESCRIPCION"
        Exit Sub
    End If
    nuevo_periodoId = 1
    If periodo.collectionAll.Count > 0 Then
        periodo.findLast
        nuevo_periodoId = periodo.periodoId + 1
    End If
    If Me.dtpInicio_add.value < ultima_fecha And periodo.collectionAll.Count > 0 Then
        MsgBox "La fecha de inicio corresponde a otro período", vbOKOnly, "Advertencia"
        Exit Sub
    End If
    If Me.dtpFin_add.value <= Me.dtpInicio_add.value Then
        MsgBox "La fecha de finalización debe ser posterior a la de inicio", vbOKOnly, "Advertencia"
        Exit Sub
    End If
    If Me.dtpPrimer_add.value <= Me.dtpFin_add.value Then
        MsgBox "El primer vencimiento debe ser posterior a la fecha de finalizacion", vbOKOnly, "Advertencia"
        Exit Sub
    End If
    If Me.dtpSegundo_add.value <= Me.dtpPrimer_add.value Then
        MsgBox "El segundo vencimiento debe ser posterior al primero", vbOKOnly, "Advertencia"
        Exit Sub
    End If
        
    periodo.periodoId = nuevo_periodoId
    periodo.fechainicio = Me.dtpInicio_add.value
    periodo.fechafin = Me.dtpFin_add.value
    periodo.descripcion = Left(desc.Text, 20)
    periodo.fechaprimero = Me.dtpPrimer_add.value
    periodo.fechasegundo = Me.dtpSegundo_add.value
    tasa = CCur(Me.tasa.Text)
    If perd.ListIndex > 0 Then
        Select Case perd.ListIndex
            Case 1
                tasa = (1 + tasa) ^ (1 / 12) - 1
            Case 2
                tasa = (1 + tasa) ^ 30 - 1
        End Select
    End If
    periodo.tasa = tasa / 100
    periodo.leyenda = aley.Text & "  "
    periodo.uid = "admin"
    
    periodo.save
    
    modify = False
    
    llenar
    
    frmPeriodo.height = form_height

End Sub

Private Sub aley_Change()
    
    modify = True

End Sub

Private Sub aley_GotFocus()
    
    aley.SelStart = 0
    aley.SelLength = Len(aley.Text)

End Sub

Private Sub aley_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then agre.SetFocus

End Sub

Private Sub aper_Click()
Dim res As Integer

On Error Resume Next
    
    If modify Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            modify = False
        Else
            Exit Sub
        End If
    End If
    frmPeriodo.height = form_height + 3050
    per.Visible = True
    vper.Visible = False
    desc.Text = ""
    Me.dtpInicio_add.value = Date
    Me.dtpFin_add.value = Date
    Me.dtpPrimer_add.value = Date
    Me.dtpSegundo_add.value = Date
    tasa.Text = ""
    aley.Text = ""
    modify = False
    desc.SetFocus

End Sub

Private Sub desc_Change()
    
    modify = True

End Sub

Private Sub desc_GotFocus()
    
    desc.SelStart = 0
    desc.SelLength = Len(desc.Text)

End Sub

Private Sub desc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Me.dtpInicio_add.SetFocus

End Sub

Private Sub dtpFin_add_Change()

    modify = True

End Sub

Private Sub dtpFin_add_GotFocus()
Dim plazo As Integer

Dim operador As New clsMyAOperador

On Error Resume Next
    
    operador.findLast dbapp
    Select Case operador.periodoFactura
        Case 1:
            plazo = 60
        Case 2:
            plazo = 30
    End Select
    Me.dtpFin_add.value = Me.dtpInicio_add.value + plazo
    
End Sub

Private Sub dtpFin_add_LostFocus()
    
    If Me.dtpFin_add.value <= Me.dtpInicio_add.value Then MsgBox "La fecha de finalización debe ser mayor a la fecha de inicio", vbOKOnly, "Advertencia"

End Sub

Private Sub dtpFin_upd_Change()

    modify = True
    
End Sub

Private Sub dtpInicio_add_Change()

    modify = True
    
End Sub

Private Sub dtpInicio_add_GotFocus()

On Error Resume Next
    
    Me.dtpInicio_add.value = ultima_fecha

End Sub

Private Sub dtpInicio_upd_Change()

    modify = True
    
End Sub

Private Sub dtpPrimer_add_Change()

    modify = True
    
End Sub

Private Sub dtpPrimer_add_GotFocus()

    Me.dtpPrimer_add.value = Me.dtpFin_add.value + 20
    
End Sub

Private Sub dtpPrimer_upd_Change()

    modify = True
    
End Sub

Private Sub dtpSegundo_add_Change()

    modify = True
    
End Sub

Private Sub dtpSegundo_add_GotFocus()

    Me.dtpSegundo_add.value = Me.dtpPrimer_add.value + 15
    
End Sub

Private Sub dtpSegundo_upd_Change()

    modify = True
    
End Sub

Private Sub elim_Click()
Dim crit As String
Dim res As Integer

Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim lectura As New clsMyALectura
Dim novedad As New clsMyANovedad

On Error Resume Next
    
    If modify Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            modify = False
        Else
            Exit Sub
        End If
    End If
    If pdef.ListIndex < 0 Then Exit Sub
    
    periodo.findLast
    If factura.collectionByPeriodoID(periodo.periodoId, dbapp).Count > 0 Then
        MsgBox "Período utilizado por las facturas"
        Exit Sub
    End If
    If lectura.collectionByPeriodoID(periodo.periodoId, dbapp).Count > 0 Then
        MsgBox "Período utilizado en las lecturas del medidor"
        Exit Sub
    End If
    If novedad.collectionByPeriodoID(periodo.periodoId, dbapp).Count > 0 Then
        MsgBox "Período utilizado en las novedades"
        Exit Sub
    End If
    
    periodo.delete
    
    llenar

End Sub

Private Sub fin_Click()
Dim res As Integer

On Error Resume Next
    
    If modify Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            modify = False
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub

Private Sub Form_Activate()
    
    vper.Visible = False
    per.Visible = False
    frmPeriodo.height = form_height
    llenar
    If pdef.ListCount > 0 Then pdef.ListIndex = 0
    pdef_Click

End Sub

Private Sub Form_Load()

On Error Resume Next

    Me.dtpInicio_add.value = Date
    Me.dtpInicio_upd.value = Date
    Me.dtpFin_add.value = Date
    Me.dtpFin_upd.value = Date
    Me.dtpPrimer_add.value = Date
    Me.dtpPrimer_upd.value = Date
    Me.dtpSegundo_add.value = Date
    Me.dtpSegundo_upd.value = Date
    
    perd.Clear
    perd.AddItem "Mensual"
    perd.AddItem "Año"
    perd.AddItem "Día"
    perd.ListIndex = 0
    mprd.Clear
    mprd.AddItem "Mensual"
    mprd.AddItem "Año"
    mprd.AddItem "Día"
    mprd.ListIndex = 0
    form_height = 2350

End Sub

Private Sub mdes_Change()
    
    modify = True

End Sub

Private Sub mdes_GotFocus()
    
    mdes.SelStart = 0
    mdes.SelLength = Len(mdes.Text)

End Sub

Private Sub mdes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Me.dtpInicio_upd.SetFocus

End Sub

Private Sub mley_Change()
    
    modify = True

End Sub

Private Sub mley_GotFocus()
    
    mley.SelStart = 0
    mley.SelLength = Len(mley.Text)

End Sub

Private Sub mley_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then modi.SetFocus

End Sub

Private Sub modi_Click()
Dim crit As String
Dim tasa As Currency
Dim i As Integer

Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    If factura.collectionByPeriodoID(pdef.ItemData(pdef.ListIndex), dbapp).Count > 0 Then
        If modify Then
            MsgBox "Este PERIODO ya fue FACTURADO"
            modify = False
            Exit Sub
        End If
    End If
    If Len(Trim(mdes.Text)) < 1 Then
        MsgBox "Debe poner la DESCRIPCION"
        Exit Sub
    End If
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.fechainicio = Me.dtpInicio_upd.value
    periodo.fechafin = Me.dtpFin_upd.value
    periodo.descripcion = Left(mdes.Text, 20)
    periodo.fechaprimero = Me.dtpPrimer_upd.value
    periodo.fechasegundo = Me.dtpSegundo_upd.value
    tasa = CDbl(mtas.Text)
    If mprd.ListIndex > 0 Then
        Select Case mprd.ListIndex
            Case 1
                tasa = (1 + tasa) ^ (1 / 12) - 1
            Case 2
                tasa = (1 + tasa) ^ 30 - 1
        End Select
    End If
    periodo.tasa = tasa / 100
    periodo.leyenda = mley.Text & " "
    periodo.uid = "admin"
    
    periodo.save
    
    modify = False
    llenar
    frmPeriodo.height = form_height

End Sub

Private Sub mper_Click()
Dim crit As String
Dim res As Integer

Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    If modify Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            modify = False
        Else
            Exit Sub
        End If
    End If
    If pdef.ListIndex < 0 Then Exit Sub
    frmPeriodo.height = form_height + 3050
    per.Visible = False
    vper.Visible = True
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey

    mdes.Text = periodo.descripcion
    Me.dtpInicio_upd.value = periodo.fechainicio
    Me.dtpFin_upd.value = periodo.fechafin
    Me.dtpPrimer_upd.value = periodo.fechaprimero
    Me.dtpSegundo_upd.value = periodo.fechasegundo
    mtas.Text = Format(periodo.tasa * 100, "#,###,##0.00")
    If IsNull(periodo.leyenda) Then
        mley.Text = " "
    Else
        mley.Text = periodo.leyenda
    End If
    
    mprd.ListIndex = 0
    modify = False
    mdes.SetFocus

End Sub

Private Sub mprd_Change()
    
    modify = True

End Sub

Private Sub mtas_Change()
    
    modify = True

End Sub

Private Sub mtas_GotFocus()
    
    MsgBox "Recuerde que la Tasa no debe ser mayor que la" & Chr(13) & Chr(13) & "TASA ACTIVA del Banco Nación"
    mtas.SelStart = 0
    mtas.SelLength = Len(mtas.Text)

End Sub

Private Sub mtas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mprd.SetFocus

End Sub

Private Sub mtas_LostFocus()
    
    If Not IsNumeric(mtas.Text) Then mtas.Text = 0
    mtas.Text = Format(CDbl(mtas.Text), "#,###,##0.00")

End Sub

Private Sub pdef_Click()
Dim res As Integer

On Error Resume Next
    
    If modify Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            modify = False
        Else
            Exit Sub
        End If
    End If
    frmPeriodo.height = form_height
    vper.Visible = False
    per.Visible = False

End Sub

Private Sub perd_Change()
    
    modify = True

End Sub

Private Sub tasa_Change()
    
    modify = True

End Sub

Private Sub tasa_GotFocus()
    
    MsgBox "Recuerde que la Tasa no debe ser mayor que la" & Chr(13) & Chr(13) & "TASA ACTIVA del Banco Nación"
    tasa.SelStart = 0
    tasa.SelLength = Len(tasa.Text)

End Sub

Private Sub tasa_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then perd.SetFocus

End Sub

Private Sub tasa_LostFocus()
    
    If Not IsNumeric(tasa.Text) Then tasa.Text = 0
    tasa.Text = Format(CDbl(tasa.Text), "#,###,##0.00")

End Sub
