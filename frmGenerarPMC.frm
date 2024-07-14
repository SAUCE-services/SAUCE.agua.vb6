VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGenerarPMC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Facturas Pago Mis Cuentas"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   11790
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdGenerar 
      Cancel          =   -1  'True
      Caption         =   "Generar"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   99352577
      CurrentDate     =   41086
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   10560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   7455
   End
   Begin VB.CommandButton cmdArchivo 
      Height          =   255
      Left            =   11280
      Picture         =   "frmGenerarPMC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar Archivo"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   99352577
      CurrentDate     =   41086
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Archivo a Generar"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmGenerarPMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()
Dim filename As String
Dim path As String
    
    Me.dlgArchivo.DialogTitle = "Buscar Carpeta"
    Me.dlgArchivo.ShowOpen
    
    filename = modConv.parseFilename(Me.dlgArchivo.filename, path)
    filename = "FAC0462." & Format(Date, "ddmmyy")
    
    Me.txtArchivo.Text = path & filename

End Sub

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.value = Me.dtpDesde.value
    
End Sub

Private Sub cmdGenerar_Click()
Dim request As MSXML2.ServerXMLHTTP

Dim file_stream As ADODB.Stream

Dim url As String

    If Trim(Me.txtArchivo.Text) = "" Then
        MsgBox "ERROR: Falta ARCHIVO"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    Set request = New MSXML2.ServerXMLHTTP
    
    url = modUrls.url_agua & "pmc/generate/pago_mis_cuentas/" & modConv.date2datetimeIso(Me.dtpDesde.value) & "/" & modConv.date2datetimeIso(Me.dtpHasta.value)

    request.setTimeouts 200000, 200000, 200000, 200000
    request.Open "GET", url, False
    request.send
    
    If request.Status = 200 Then
        Set file_stream = New ADODB.Stream
        file_stream.Open
        file_stream.Type = adTypeBinary
        
        file_stream.Write request.responseBody
        file_stream.Position = 0
        
        file_stream.SaveToFile Me.txtArchivo.Text, adSaveCreateOverWrite
        file_stream.Close
        
        Set file_stream = Nothing
    End If
    
    Set request = Nothing
    
    Me.MousePointer = 0
    
    MsgBox "Generación TERMINADA"
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.dtpDesde.value = Date
    Me.dtpHasta.value = Date
    
End Sub
