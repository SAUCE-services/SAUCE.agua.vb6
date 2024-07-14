VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSuspension 
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
Attribute VB_Name = "frmSuspension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub corte_Click()
    If corte.Value Then llenar_cor
End Sub
Private Sub fin_Click()
    suspension.Hide
End Sub
Private Sub Form_Activate()
    restri.Value = True
    llenar_res
End Sub
Private Sub impri_Click()
Dim np, ct, num As Integer
Dim busq, mens, crit, critp, critq, tipo, cuit, numero As String
Dim total As Double
'On Error Resume Next
    If ncli.ListIndex < 0 Then Exit Sub
    impr.Destination = Impresora
    If Val(copia.Text) = 0 Then copia.Text = 1
    If restri.Value Then
        impr.ReportFileName = Confi.PathDB & "restric.rpt"
        critq = "Tipo = 'R'"
        tipo = "R"
        numero = "numres= "
    End If
    If corte.Value Then
        impr.ReportFileName = Confi.PathDB & "corte.rpt"
        critq = "Tipo = 'C'"
        tipo = "C"
        numero = "numcor= "
    End If
    num = 1
    suspensiones.findLast critq
    If Not suspensiones.NoMatch Then num = suspensiones!num + 1
    busq = "tipo = '" & tipo & "' and IDCliente = " & ncli.ItemData(ncli.ListIndex) & " and fecha = datevalue('" & Date & "')"
    suspensiones.findLast busq
    If suspensiones.NoMatch Then
        suspensiones.AddNew
        suspensiones!tipo = tipo
        suspensiones!num = num
        suspensiones!fecha = Date
        suspensiones!IDCliente = ncli.ItemData(ncli.ListIndex)
        If Not periodos.BOF Then
            periodos.MoveLast
            Do While periodos!finicio > Date
                periodos.MovePrevious
            Loop
            suspensiones!IDPeriodo = periodos!IDPeriodo + 1
        End If
        suspensiones!login = usuario
        suspensiones!fmov = Now
        suspensiones.update
    Else
        num = suspensiones!num
    End If
    operadores.MoveFirst
    cuit = Left(operadores!cuit, 2) & "-" & Mid(operadores!cuit, 3, 8) & "-" & Right(operadores!cuit, 1)
    np = 31
    If np > Len(operadores!razons) Then
        np = Len(operadores!razons)
    Else
        Do While Mid(operadores!razons, np, 1) <> " "
            np = np - 1
        Loop
    End If
    Select Case operadores!sitIva
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
    impr.Formulas(0) = "nomope= '" & Mid(operadores!razons, 1, np) & "'"
    impr.Formulas(1) = "nomop1= '" & Mid(operadores!razons, np + 1) & "'"
    impr.Formulas(2) = "domope= '" & operadores!calle & " " & operadores!puerta & " " & operadores!piso & " " & operadores!dpto & " - " & operadores!localidad & "'"
    impr.Formulas(3) = "locope= 'C.P. " & operadores!cpostal & " " & operadores!provincia & " - Tel: " & operadores!telef & "'"
    impr.Formulas(4) = "opegr1= 'C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operadores!ingBrutos & "'"
    impr.Formulas(5) = "opegr2= 'I.V.A. " & mens & "   NRO. E.P.A.S. " & operadores!NEpas & "'"
    impr.Formulas(42) = "resol= '" & operadores!reso & "'"
    impr.Formulas(43) = "perso= '" & operadores!perso & "'"
    crit = "IDCliente = " & ncli.ItemData(ncli.ListIndex)
    clientes.findLast crit
    impr.Formulas(6) = "nomcli= '" & clientes!apellido & ", " & clientes!nombre & "'"
    impr.Formulas(7) = "ubiinm= '" & clientes!icalle & " " & clientes!ipuerta & " " & clientes!ipiso & " " & clientes!idpto & "'"
    impr.Formulas(8) = "nomcat= '" & clientes!nomCat & "'"
    impr.Formulas(9) = "domcli= '" & clientes!fcalle & " " & clientes!fpuerta & " " & clientes!fpiso & " " & clientes!fdpto & "'"
    impr.Formulas(10) = "numsoc= '" & clientes!nsocio & "'"
    impr.Formulas(39) = "fecemi= '" & Date & "'"
    impr.Formulas(40) = numero & "'" & num & "'"
    mens = ""
    If Len(Trim(clientes!cuit)) > 0 Then mens = Left(clientes!cuit, 2) & "-" & Mid(clientes!cuit, 3, 8) & "-" & Right(clientes!cuit, 1)
    Select Case clientes!sitIva
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
    Select Case clientes!categoria
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
    crit = "IDCliente = " & ncli.ItemData(ncli.ListIndex) & " and Pagada = False and Anulada = False"
    facturas.findLast crit
    ct = 1
    Do While ct < 6 And Not facturas.NoMatch
        critp = "IDPeriodo = " & facturas!IDPeriodo
        periodos.findLast critp
        impr.Formulas(13 + ct) = "per(" & ct & ")= '" & periodos!descrip & "'"
        impr.Formulas(18 + ct) = "fac(" & ct & ")= '" & Right("0000" & facturas!IDSuc, 4) & "-" & Right("00000000" & facturas!IDFac, 8) & "'"
        impr.Formulas(23 + ct) = "ven(" & ct & ")= '" & periodos!fprimer & "'"
        impr.Formulas(28 + ct) = "imf(" & ct & ")= '" & Format(facturas!total, "#,###,##0.00") & "'"
        impr.Formulas(33 + ct) = "iin(" & ct & ")= '" & Format(facturas!total + interes(facturas!total, facturas!tasa, facturas!fecha, Date), "#,###,##0.00") & "'"
        total = total + facturas!total + interes(facturas!total, facturas!tasa, facturas!fecha, Date)
        facturas.FindPrevious crit
        ct = ct + 1
    Loop
    crit = "IDCliente = " & ncli.ItemData(ncli.ListIndex) & " and Pagado = False"
    deudas.findLast crit
    Do While ct < 6 And Not deudas.NoMatch
        impr.Formulas(23 + ct) = "ven(" & ct & ")= '" & deudas!cuotas - deudas!cpagada & " cuota(s)'"
        impr.Formulas(33 + ct) = "iin(" & ct & ")= '" & Format(deudas!deuda, "#,###,##0.00") & "'"
        total = total + deudas!deuda
        deudas.FindPrevious crit
        ct = ct + 1
    Loop
    impr.Formulas(41) = "total= '" & Format(total, "#,###,##0.00") & "'"
    For ct = 1 To Val(copia.Text)
        estado.Caption = "Imprimiendo Acta . . ."
        estado.Refresh
        impr.Action = 1
    Next ct
    impr.SelectionFormula = ""
End Sub
Private Sub llenar_res()
Dim consul As String
On Error GoTo agreg
    ncli.Clear
    If periodos.RecordCount = 0 Then Exit Sub
    estado.Caption = "Cargando posibles RESTRICCIONES . . ."
    estado.Refresh
    consul = "SELECT * FROM CliRes"
    Set auxiliar = db.OpenRecordset(consul, dbOpenSnapshot)
    If auxiliar.RecordCount Then
        auxiliar.MoveFirst
        Do While Not auxiliar.EOF
            ncli.Text = auxiliar!apellido & ", " & auxiliar!nombre
            auxiliar.MoveNext
        Loop
    End If
    If ncli.ListCount Then ncli.ListIndex = 0
final:
    estado.Caption = ""
    Exit Sub
agreg:
    If Err.Number = 383 Then
        ncli.AddItem auxiliar!apellido & ", " & auxiliar!nombre
        ncli.ItemData(ncli.NewIndex) = auxiliar!IDCliente
    Else
        MsgBox "Ocurrió un error que no puedo resolver" & Chr(13) & Chr(13) & "ERROR : " & Err.Number & " - Por favor contáctese con el Servicio Técnico"
        Resume final
    End If
    Resume Next
End Sub
Private Sub llenar_cor()
Dim consul As String
On Error GoTo agreg
    ncli.Clear
    If periodos.RecordCount = 0 Then Exit Sub
    estado.Caption = "Cargando posibles CORTES . . ."
    estado.Refresh
    consul = "SELECT * FROM CliCor"
    Set auxiliar = db.OpenRecordset(consul, dbOpenSnapshot)
    If Not auxiliar.BOF Then
        auxiliar.MoveFirst
        Do While Not auxiliar.EOF
            ncli.Text = auxiliar!apellido & ", " & auxiliar!nombre
            auxiliar.MoveNext
        Loop
    End If
    If ncli.ListCount Then ncli.ListIndex = 0
final:
    estado.Caption = ""
    Exit Sub
agreg:
    If Err.Number = 383 Then
        ncli.AddItem auxiliar!apellido & ", " & auxiliar!nombre
        ncli.ItemData(ncli.NewIndex) = auxiliar!IDCliente
    Else
        MsgBox "Ocurrió un error que no puedo resolver" & Chr(13) & Chr(13) & "ERROR : " & Err.Number & " - Por favor contáctese con el Servicio Técnico"
        Resume final
    End If
    Resume Next
End Sub
Private Sub restri_Click()
    If restri.Value Then llenar_res
End Sub
