VERSION 5.00
Begin VB.Form frmRubro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros de Facturación"
   ClientHeight    =   5625
   ClientLeft      =   2445
   ClientTop       =   2520
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7695
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Fin de la TAREA"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton elim 
      Caption         =   "&Eliminar Rubro"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Permite la Eliminación del RUBRO seleccionado"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton agre 
      Caption         =   "&Agregar Rubro"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Permite agregar un RUBRO nuevo a la Lista"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton modi 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Permite modificar los DATOS del RUBRO seleccionado"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton hist 
      Caption         =   "&Ver Histórico"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      ToolTipText     =   "Muestra la evolución del RUBRO seleccionado"
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox rdef 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Frame vmod 
      Caption         =   "Modificación de Rubro"
      Height          =   3015
      Left            =   240
      TabIndex        =   29
      Top             =   2400
      Width           =   7215
      Begin VB.Frame mcomu 
         Height          =   1335
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   4095
         Begin VB.OptionButton mtde 
            Caption         =   "Todos los Desconectados"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   41
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox mcli 
            Caption         =   "Común"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton mtcl 
            Caption         =   "Todos los Clientes"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton mtso 
            Caption         =   "Todos los Socios"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.ComboBox cobr 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton conf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   5400
         TabIndex        =   25
         ToolTipText     =   "Graba las modificaciones del RUBRO activo"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox mcon 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox IVA 
         Caption         =   "Aplicar IVA"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ran 
         Caption         =   "Asoc. a un RANGO"
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox puni 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox rcom 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cobro por"
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   37
         Top             =   240
         Width           =   690
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   690
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Costo del Cargo"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Rangos DEFINIDOS"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.Frame vagr 
      Caption         =   "Agregar RUBRO"
      Height          =   3015
      Left            =   240
      TabIndex        =   33
      Top             =   2400
      Width           =   7215
      Begin VB.Frame acomu 
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   4095
         Begin VB.OptionButton atde 
            Caption         =   "Todos los Desconectados"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   42
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton atso 
            Caption         =   "Todos los Socios"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton atcl 
            Caption         =   "Todos los Clientes"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox acli 
            Caption         =   "Común"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox acob 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox apre 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox aran 
         Caption         =   "Asoc. a un RANGO"
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox acon 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox acom 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CheckBox aIVA 
         Caption         =   "Aplicar IVA"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton aconf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         ToolTipText     =   "Graba los DATOS del nuevo RUBRO"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Cobro por"
         Height          =   195
         Index           =   8
         Left            =   5400
         TabIndex        =   38
         Top             =   240
         Width           =   690
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Rangos DEFINIDOS"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Costo del Cargo"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame vhis 
      Caption         =   "Histórico"
      Height          =   2055
      Left            =   240
      TabIndex        =   27
      Top             =   2400
      Width           =   7215
      Begin VB.ListBox rhis 
         Height          =   1425
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Rubros DEFINIDOS"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsin As Integer
Private vari As Boolean

Public Sub fillRangos(pCombo As ComboBox)
Dim general As String
Dim rang As Integer
Dim enc As Boolean

Dim rango As New clsMyARango

On Error Resume Next

    rang = 1
    pCombo.Clear
    Do
        enc = False
        With rango
            .categoria = 1
            .rangoID = rang
            
            .findLast dbapp
            
            If .autoID > 0 Then
                general = Right("000" & .limiteInferior, 3) & " - " & Right("000" & .limiteSuperior, 3) & " m³  $ " & Format(.tarifa, "#,###,##0.00")
                enc = True
            Else
                general = Space(32)
            End If
        End With
        With rango
            .categoria = 2
            .rangoID = rang
            
            .findLast dbapp
        
            If .autoID > 0 Then
                general = general & "                     | " & Right("000" & .limiteInferior, 3) & " - " & Right("000" & .limiteSuperior, 3) & " m³  $ " & Format(.tarifa, "#,###,##0.00")
                enc = True
            End If
        End With
        If Len(Trim(general)) Then pCombo.AddItem general
        rang = rang + 1
    Loop While enc

End Sub

Public Sub llenar()
Dim nombre As String

Dim rubro As New clsMyARubro
Dim rango As New clsMyARango

On Error Resume Next

    rdef.Clear
    For Each rubro In rubro.collectionSinRepeticion(dbapp)
        With rubro
            nombre = Right("00" & .rubroID, 2) & " - " & .concepto
            If .rangoID > 0 Then
                nombre = nombre & " Sistema Medido "
                rango.categoria = 1
                rango.rangoID = .rangoID
                rango.findLast dbapp
                If rango.autoID > 0 Then nombre = nombre & "(" & rango.limiteInferior & "-" & rango.limiteSuperior & " m³ (G))"
                rango.categoria = 2
                rango.rangoID = .rangoID
                rango.findLast dbapp
                If rango.autoID > 0 Then nombre = nombre & "(" & rango.limiteInferior & "-" & rango.limiteSuperior & " m³ (E))"
            End If
        End With
        If Len(Trim(nombre)) Then
            rdef.AddItem nombre
            rdef.ItemData(rdef.NewIndex) = rubro.rubroID
        End If
    Next

End Sub

Private Sub acli_Click()
    
    If acli.Value = 1 Then
        atcl.Enabled = True
        atcl.Value = True
        atso.Enabled = True
        atde.Enabled = True
    Else
        atcl.Value = False
        atcl.Enabled = False
        atso.Value = False
        atso.Enabled = False
        atde.Enabled = False
    End If
    vari = True

End Sub

Private Sub acob_Change()
    
    vari = True

End Sub

Private Sub acom_Change()
    
    vari = True

End Sub

Private Sub acon_Change()
    
    vari = True

End Sub

Private Sub acon_GotFocus()
    
    acon.SelStart = 0
    acon.SelLength = Len(acon.Text)

End Sub

Private Sub acon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aIVA.SetFocus

End Sub

Private Sub acon_LostFocus()
    
    etiq(5) = "Costo de " & acon.Text

End Sub

Private Sub aconf_Click()
Dim rubro As New clsMyARubro

On Error Resume Next
    
    If aran.Value = 1 And acom.ListIndex < 0 Then
        MsgBox "Debe seleccionar un rango . . ."
        Exit Sub
    End If
    If Not IsNumeric(apre.Text) Then apre.Text = Format(0, "#,###,##0.00")
    If aran.Value = 0 And CDbl(apre.Text) = 0 Then
        MsgBox "El COSTO debe ser mayor que CERO . . ."
        apre.SetFocus
    End If
    If Len(Trim(acon.Text)) < 1 Then
        MsgBox "Debe ingresar un Concepto . . ."
        acon.SetFocus
        Exit Sub
    End If
    
    With rubro
        .rubroID = rdef.ListCount + 1
        .fecha = Date
        .concepto = Left(acon.Text, 80)
        If aIVA.Value = 0 Then
            .IVA = 0
        Else
            .IVA = 1
        End If
        .comun = atcl.Value
        .comunSocio = atso.Value
        .desconectado = atde.Value
        If aran.Value = 1 Then
            .rangoID = acom.ListIndex + 1
            .precioUnitario = 0
        Else
            .rangoID = 0
            .precioUnitario = CDbl(apre.Text)
        End If
        .cobro = acob.ListIndex
        .uid = "admin"
        
        .add dbapp
    End With
    
    vari = False
    
    llenar
    
    vagr.Visible = False
    vmod.Visible = False
    vhis.Visible = False
    frmRubro.Height = rsin

End Sub

Private Sub agre_Click()
Dim res As Integer

On Error Resume Next
    
    If vari Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    vhis.Visible = False
    vmod.Visible = False
    vagr.Visible = True
    frmRubro.Height = rsin + 3200
    acon.Text = ""
    aIVA.Value = 1
    aran.Value = 0
    acli.Value = 0
    apre.Visible = True
    apre.Text = ""
    acom.Visible = False
    etiq(6).Visible = False
    acob.ListIndex = 0
    vari = False
    acon.SetFocus

End Sub

Private Sub aIVA_Click()
    
    vari = True

End Sub

Private Sub apre_Change()
    
    vari = True

End Sub

Private Sub apre_GotFocus()
    
    apre.SelStart = 0
    apre.SelLength = Len(apre.Text)

End Sub

Private Sub apre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aran.SetFocus

End Sub

Private Sub apre_LostFocus()
    
    If Not IsNumeric(apre.Text) Then apre.Text = 0
    apre.Text = Format(CDbl(apre.Text), "#,###,##0.00")

End Sub

Private Sub aran_Click()

    vari = True
    If aran.Value = 1 Then
        etiq(6).Visible = True
        etiq(5).Visible = False
        acom.Visible = True
        apre.Visible = False
        acob.ListIndex = 1
        acli.Value = 1
        atcl.Value = True
        
        fillRangos acom
    Else
        etiq(6).Visible = False
        etiq(5).Visible = True
        acom.Visible = False
        apre.Visible = True
    End If

End Sub

Private Sub atcl_Click()
    
    vari = True

End Sub

Private Sub atde_Click()
    
    vari = True

End Sub

Private Sub atso_Click()
    
    vari = True

End Sub

Private Sub cobr_Change()
    
    vari = True

End Sub

Private Sub conf_Click()
Dim rubro As New clsMyARubro

On Error Resume Next
    
    If ran.Value = 1 And rcom.ListIndex < 0 Then
        MsgBox "Debe seleccionar un rango"
        Exit Sub
    End If
    
    If Len(puni.Text) = 0 Then puni.Text = Format(0, "#,###,##0.00")
    
    With rubro
        .rubroID = rdef.ListIndex + 1
        .fecha = Date
        .concepto = Left(mcon.Text, 80)
        If IVA.Value = 0 Then
            .IVA = False
        Else
            .IVA = True
        End If
        .comun = mtcl.Value
        .comunSocio = mtso.Value
        .desconectado = mtde.Value
        If ran.Value = 1 Then
            .rangoID = rcom.ListIndex + 1
            .precioUnitario = 0
        Else
            .rangoID = 0
            .precioUnitario = CDbl(puni.Text)
        End If
        .cobro = cobr.ListIndex
        .uid = "admin"
        .save dbapp
    End With
    
    vari = False
    
    llenar
    
    vmod.Visible = False
    vhis.Visible = False
    vagr.Visible = False
    frmRubro.Height = rsin

End Sub

Private Sub elim_Click()
Dim rubroID As Integer
Dim crub As Integer
Dim res As Integer

Dim rubro As New clsMyARubro
Dim detalle As New clsMyADetalle
Dim novedad As New clsMyANovedad

On Error Resume Next
    
    If vari Then
        res = MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA")
        If res = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    
    If rdef.ListIndex < 0 Then Exit Sub
    
    rubroID = rdef.ListIndex + 1
    
    crub = rubro.collectionByRubroID(rubroID, dbapp).Count
    If crub = 1 And rdef.ListIndex < rdef.ListCount - 1 Then Exit Sub
    
    If crub = 1 And detalle.collectionByRubroID(rubroID, dbapp).Count > 0 Then
        MsgBox "Rubro utilizado por la Facturacion"
        Exit Sub
    End If
    
    If crub = 1 And novedad.collectionByRubroID(rubroID, dbapp).Count > 0 Then
        MsgBox "Rubro utilizado por las Novedades"
        Exit Sub
    End If
    
    With rubro
        .rubroID = rubroID
        .findLast dbapp
        
        .delete dbapp
    End With
    
    llenar
    
    frmRubro.Height = rsin
    vmod.Visible = False
    vagr.Visible = False
    vhis.Visible = False

End Sub

Private Sub fin_Click()

    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim serv As String

Dim objMOpe As New clsMyAOperador

On Error Resume Next
    
    objMOpe.findLast dbapp
    etiq(2) = "Costo del Cargo Servicio"
    etiq(5) = "Costo del Cargo Servicio"
    Select Case objMOpe.servicio
        Case 1:
            serv = " Agua"
        Case 2:
            serv = " Cloaca"
        Case 3:
            serv = " Agua y Cloaca"
    End Select
    etiq(2) = etiq(2) & serv
    etiq(5) = etiq(5) & serv
    llenar
    frmRubro.Height = rsin

End Sub

Private Sub Form_Load()

On Error Resume Next
    
    cobr.Clear
    cobr.AddItem "Ninguno"
    cobr.AddItem "Sistema Medido"
    cobr.AddItem "Cuota Fija"
    acob.Clear
    acob.AddItem "Ninguno"
    acob.AddItem "Sistema Medido"
    acob.AddItem "Cuota Fija"
    rsin = 2750

End Sub

Private Sub hist_Click()
Dim rubroID As Integer
Dim precio As String

Dim rubro As New clsMyARubro
Dim rango As New clsMyARango

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    
    If rdef.ListIndex < 0 Then Exit Sub
    
    rubroID = rdef.ListIndex + 1
    frmRubro.Height = rsin + 2300
    vhis.Visible = True
    vmod.Visible = False
    vagr.Visible = False
    rhis.Clear
    
    For Each rubro In rubro.collectionByRubroID(rubroID, dbapp)
        If rubro.rangoID > 0 Then
            rango.rangoID = rubro.rangoID
            rango.findByFechaSinCategoria rubro.fecha, dbapp
            precio = rango.tarifa
        Else
            precio = rubro.precioUnitario
        End If
        
        rhis.AddItem "Fecha " & rubro.fecha & "   " & rubro.concepto & "   Tarifa : $ " & Format(precio, "#,###,##0.00")
    Next

End Sub

Private Sub iva_Click()
    
    vari = True

End Sub

Private Sub mcli_Click()
    
    If mcli.Value = 1 Then
        mtcl.Enabled = True
        mtcl.Value = True
        mtso.Enabled = True
        mtde.Enabled = True
    Else
        mtcl.Value = False
        mtcl.Enabled = False
        mtso.Value = False
        mtso.Enabled = False
        mtde.Value = False
        mtde.Enabled = False
    End If
    vari = True

End Sub

Private Sub mcon_Change()
    
    vari = True

End Sub

Private Sub mcon_GotFocus()
    
    mcon.SelStart = 0
    mcon.SelLength = Len(mcon.Text)

End Sub

Private Sub mcon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then IVA.SetFocus

End Sub

Private Sub mcon_LostFocus()
    
    etiq(2).Caption = "Costo de " & mcon.Text

End Sub

Private Sub modi_Click()
Dim rubroID As Integer

Dim rubro As New clsMyARubro

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    If rdef.ListIndex < 0 Then Exit Sub
    frmRubro.Height = rsin + 3200
    vmod.Visible = True
    vhis.Visible = False
    vagr.Visible = False
    rubroID = rdef.ListIndex + 1
    
    With rubro
        .rubroID = rubroID
        .findLast dbapp
        If .rangoID > 0 Then
            ran.Value = 1
        Else
            ran.Value = 0
        End If
        mcon.Text = .concepto
        IVA.Value = .IVA
        If .comun Or .comunSocio Or .desconectado Then
            mcli.Value = 1
        Else
            mcli.Value = 0
        End If
        mtcl.Value = .comun
        mtso.Value = .comunSocio
        mtde.Value = .desconectado
        If ran.Value = 1 Then
            etiq(3).Visible = True
            etiq(2).Visible = False
            rcom.Visible = True
            puni.Visible = False
            fillRangos rcom
            rcom.ListIndex = .rangoID - 1
        Else
            etiq(3).Visible = False
            etiq(2).Visible = True
            rcom.Visible = False
            puni.Visible = True
            puni.Text = Format(.precioUnitario, "#,###,##0.00")
        End If
        cobr.ListIndex = .cobro
    End With
    vari = False
    mcon.SetFocus

End Sub

Private Sub mtcl_Click()
    
    vari = True

End Sub

Private Sub mtde_Click()
    
    vari = True

End Sub

Private Sub mtso_Click()
    
    vari = True

End Sub

Private Sub puni_Change()
    
    vari = True

End Sub

Private Sub puni_GotFocus()
    
    puni.SelStart = 0
    puni.SelLength = Len(puni.Text)

End Sub

Private Sub puni_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ran.SetFocus

End Sub

Private Sub puni_LostFocus()
    
    If Not IsNumeric(puni.Text) Then puni.Text = 0
    puni.Text = Format(CDbl(puni.Text), "#,###,##0.00")

End Sub

Private Sub ran_Click()
Dim rubro As New clsMyARubro

On Error Resume Next

    With rubro
        .rubroID = Me.rdef.ItemData(Me.rdef.ListIndex)
        .findLast dbapp
    End With
    
    If ran.Value = 1 Then
        etiq(3).Visible = True
        etiq(2).Visible = False
        rcom.Visible = True
        puni.Visible = False
        cobr.ListIndex = 1
        mcli.Value = 1
        mtcl.Value = True
        fillRangos rcom
        rcom.ListIndex = rubro.rangoID - 1
    Else
        etiq(3).Visible = False
        etiq(2).Visible = True
        rcom.Visible = False
        puni.Visible = True
    End If
    vari = True

End Sub

Private Sub rcom_Change()
    
    vari = True

End Sub

Private Sub rdef_Click()

    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    vhis.Visible = False
    vmod.Visible = False
    frmRubro.Height = rsin

End Sub
