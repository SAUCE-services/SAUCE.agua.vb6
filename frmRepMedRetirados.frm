VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepMedRetirados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medidores Retirados"
   ClientHeight    =   810
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4095
   Begin Crystal.CrystalReport crpFacturas 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmRepMedRetirados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim listado As New clsMyAListado
Dim operador As New clsMyAOperador
Dim medlist As New clsVMyAMedList

Dim cuit As String
Dim mens As String

Dim impresionService As New clsCtlImpresion

Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer
Dim c4 As Integer

    Me.MousePointer = 11
    
    listado.truncate dbapp
    
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
    
    c1 = 0
    c2 = 0
    c3 = 0
    c4 = 0
    For Each medlist In medlist.collectionAll(dbapp)
        Set listado = New clsMyAListado
        listado.c6 = medlist.medidorID
        listado.c1 = medlist.fechaRetiro
        Select Case medlist.motivoRetiro
            Case 1:
                listado.c2 = "X"
                c1 = c1 + 1
            Case 2:
                listado.c3 = "X"
                c2 = c2 + 1
            Case 3:
                listado.c4 = "X"
                c3 = c3 + 1
            Case 4:
                listado.c5 = "X"
                c4 = c4 + 1
        End Select
        listado.add dbapp
    Next
    
    impresionService.printReport Me.crpFacturas, "rptMedidor", dbapp.stringConnection, , , , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Medidores Retirados"), _
        Array("totrot", c1), _
        Array("totobs", c2), _
        Array("totpru", c3), _
        Array("totmal", c4))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

