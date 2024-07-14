VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7260
   ClientLeft      =   2310
   ClientTop       =   1455
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10215
   Begin Crystal.CrystalReport impr 
      Left            =   5520
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton dbaja 
      Caption         =   "&Baja de Cliente"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      ToolTipText     =   "Permite la BAJA de la CONEXION"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton modif 
      Caption         =   "&Modificar Datos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      ToolTipText     =   "Modifica los DATOS de la CONEXION actual"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton agreg 
      Caption         =   "&Agregar Conexión"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Agrega los datos del CLIENTE para una CONEXION nueva"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton histo 
      Caption         =   "&Ver Histórico"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      ToolTipText     =   "Muestra la evolución cronológica de la CONEXION"
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox ncli 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.TextBox ncon 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame vagr 
      Caption         =   "Agregar CONEXION"
      Height          =   4215
      Left            =   240
      TabIndex        =   36
      Top             =   2880
      Width           =   9735
      Begin VB.CommandButton aconf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   7680
         TabIndex        =   33
         ToolTipText     =   "Graba los DATOS de la nueva CONEXION"
         Top             =   360
         Width           =   1575
      End
      Begin TabDlg.SSTab adatos 
         Height          =   3015
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5318
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmCliente.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "etiq(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "etiq(4)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "etiq(5)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "etiq(7)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "etiq(8)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "etiq(9)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "etiq(10)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "etiq(50)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "etiq(62)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "etiq(66)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "etiq(67)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "etiq(68)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "etiq(69)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "etiq(30)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "aapel"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "anomb"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "acuit"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "asdgi"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "acate"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "aserv"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "acobr"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "asoci"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtDocumentoA"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "dtpNacimientoA"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "cboDestinoA"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "cboCategoriaA"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "cboEstadoA"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtNombreCategoriaA"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).ControlCount=   28
         TabCaption(1)   =   "Ubicación del Inmueble"
         TabPicture(1)   =   "frmCliente.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "aior"
         Tab(1).Control(1)=   "airu"
         Tab(1).Control(2)=   "aizo"
         Tab(1).Control(3)=   "aico"
         Tab(1).Control(4)=   "aipr"
         Tab(1).Control(5)=   "ailo"
         Tab(1).Control(6)=   "aide"
         Tab(1).Control(7)=   "aipi"
         Tab(1).Control(8)=   "aipu"
         Tab(1).Control(9)=   "aica"
         Tab(1).Control(10)=   "etiq(57)"
         Tab(1).Control(11)=   "etiq(56)"
         Tab(1).Control(12)=   "etiq(55)"
         Tab(1).Control(13)=   "etiq(17)"
         Tab(1).Control(14)=   "etiq(16)"
         Tab(1).Control(15)=   "etiq(15)"
         Tab(1).Control(16)=   "etiq(14)"
         Tab(1).Control(17)=   "etiq(13)"
         Tab(1).Control(18)=   "etiq(12)"
         Tab(1).Control(19)=   "etiq(11)"
         Tab(1).ControlCount=   20
         TabCaption(2)   =   "Domicilio del Cliente"
         TabPicture(2)   =   "frmCliente.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "etiq(18)"
         Tab(2).Control(1)=   "etiq(19)"
         Tab(2).Control(2)=   "etiq(20)"
         Tab(2).Control(3)=   "etiq(21)"
         Tab(2).Control(4)=   "etiq(22)"
         Tab(2).Control(5)=   "etiq(23)"
         Tab(2).Control(6)=   "etiq(24)"
         Tab(2).Control(7)=   "etiq(63)"
         Tab(2).Control(8)=   "etiq(64)"
         Tab(2).Control(9)=   "etiq(65)"
         Tab(2).Control(10)=   "afca"
         Tab(2).Control(11)=   "afpu"
         Tab(2).Control(12)=   "afpi"
         Tab(2).Control(13)=   "afde"
         Tab(2).Control(14)=   "aflo"
         Tab(2).Control(15)=   "afpr"
         Tab(2).Control(16)=   "afco"
         Tab(2).Control(17)=   "txtCorreoA"
         Tab(2).Control(18)=   "txtCelularA"
         Tab(2).Control(19)=   "txtFijoA"
         Tab(2).ControlCount=   20
         Begin VB.TextBox txtNombreCategoriaA 
            Height          =   285
            Left            =   240
            TabIndex        =   161
            Top             =   2520
            Width           =   3375
         End
         Begin VB.ComboBox cboEstadoA 
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox cboCategoriaA 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox cboDestinoA 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   1320
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpNacimientoA 
            Height          =   375
            Left            =   7440
            TabIndex        =   143
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100663297
            CurrentDate     =   43135
         End
         Begin VB.TextBox txtFijoA 
            Height          =   285
            Left            =   -71160
            TabIndex        =   139
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtCelularA 
            Height          =   285
            Left            =   -69360
            TabIndex        =   138
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtCorreoA 
            Height          =   285
            Left            =   -71160
            TabIndex        =   137
            Top             =   2520
            Width           =   3375
         End
         Begin VB.TextBox txtDocumentoA 
            Height          =   285
            Left            =   5640
            TabIndex        =   135
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox aior 
            Height          =   285
            Left            =   -69360
            TabIndex        =   25
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox airu 
            Height          =   285
            Left            =   -71160
            TabIndex        =   24
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox aizo 
            Height          =   285
            Left            =   -72960
            TabIndex        =   23
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox asoci 
            Height          =   285
            Left            =   3840
            TabIndex        =   15
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox afco 
            Height          =   285
            Left            =   -74760
            TabIndex        =   32
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox afpr 
            Height          =   285
            Left            =   -72960
            TabIndex        =   31
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox aflo 
            Height          =   285
            Left            =   -74760
            TabIndex        =   30
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox afde 
            Height          =   285
            Left            =   -72960
            TabIndex        =   29
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox afpi 
            Height          =   285
            Left            =   -74760
            TabIndex        =   28
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox afpu 
            Height          =   285
            Left            =   -71160
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox afca 
            Height          =   285
            Left            =   -74760
            TabIndex        =   26
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox aico 
            Height          =   285
            Left            =   -74760
            TabIndex        =   22
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox aipr 
            Height          =   285
            Left            =   -72960
            TabIndex        =   21
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox ailo 
            Height          =   285
            Left            =   -74760
            TabIndex        =   20
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox aide 
            Height          =   285
            Left            =   -72960
            TabIndex        =   19
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox aipi 
            Height          =   285
            Left            =   -74760
            TabIndex        =   18
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox aipu 
            Height          =   285
            Left            =   -71160
            TabIndex        =   17
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox aica 
            Height          =   285
            Left            =   -74760
            TabIndex        =   16
            Top             =   720
            Width           =   3375
         End
         Begin VB.ComboBox acobr 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox aserv 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox acate 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox asdgi 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox acuit 
            Height          =   285
            Left            =   3840
            TabIndex        =   14
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox anomb 
            Height          =   285
            Left            =   3840
            TabIndex        =   9
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox aapel 
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Categoría"
            Height          =   195
            Index           =   30
            Left            =   240
            TabIndex        =   162
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   69
            Left            =   7440
            TabIndex        =   150
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Categoría Socio"
            Height          =   195
            Index           =   68
            Left            =   5640
            TabIndex        =   148
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Destino Servicio"
            Height          =   195
            Index           =   67
            Left            =   5640
            TabIndex        =   146
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Nacimiento"
            Height          =   195
            Index           =   66
            Left            =   7440
            TabIndex        =   144
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono Fijo"
            Height          =   195
            Index           =   65
            Left            =   -71160
            TabIndex        =   142
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono Celular"
            Height          =   195
            Index           =   64
            Left            =   -69360
            TabIndex        =   141
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Correo Electrónico"
            Height          =   195
            Index           =   63
            Left            =   -71160
            TabIndex        =   140
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   62
            Left            =   5640
            TabIndex        =   136
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Orden Físico"
            Height          =   195
            Index           =   57
            Left            =   -69360
            TabIndex        =   126
            Top             =   2280
            Width           =   915
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
            Height          =   195
            Index           =   56
            Left            =   -71160
            TabIndex        =   125
            Top             =   2280
            Width           =   345
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            Height          =   195
            Index           =   55
            Left            =   -72960
            TabIndex        =   124
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número de Socio"
            Height          =   195
            Index           =   50
            Left            =   3840
            TabIndex        =   119
            Top             =   2280
            Width           =   1230
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal"
            Height          =   195
            Index           =   24
            Left            =   -74760
            TabIndex        =   59
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            Height          =   195
            Index           =   23
            Left            =   -72960
            TabIndex        =   58
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Localidad"
            Height          =   195
            Index           =   22
            Left            =   -74760
            TabIndex        =   57
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   21
            Left            =   -72960
            TabIndex        =   56
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Piso"
            Height          =   195
            Index           =   20
            Left            =   -74760
            TabIndex        =   55
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   19
            Left            =   -71160
            TabIndex        =   54
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Index           =   18
            Left            =   -74760
            TabIndex        =   53
            Top             =   480
            Width           =   345
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal"
            Height          =   195
            Index           =   17
            Left            =   -74760
            TabIndex        =   52
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            Height          =   195
            Index           =   16
            Left            =   -72960
            TabIndex        =   51
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Localidad"
            Height          =   195
            Index           =   15
            Left            =   -74760
            TabIndex        =   50
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   14
            Left            =   -72960
            TabIndex        =   49
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Piso"
            Height          =   195
            Index           =   13
            Left            =   -74760
            TabIndex        =   48
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   12
            Left            =   -71160
            TabIndex        =   47
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Index           =   11
            Left            =   -74760
            TabIndex        =   46
            Top             =   480
            Width           =   345
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cobro"
            Height          =   195
            Index           =   10
            Left            =   3840
            TabIndex        =   45
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Servicio"
            Height          =   195
            Index           =   9
            Left            =   2040
            TabIndex        =   44
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Categoría del Cliente"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Situación AFIP"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   42
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T."
            Height          =   195
            Index           =   5
            Left            =   3840
            TabIndex        =   41
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   40
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Apellido"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   480
            Width           =   555
         End
      End
      Begin VB.TextBox acon 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de CONEXION"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   37
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame vmod 
      Caption         =   "Modificar DATOS"
      Height          =   4215
      Left            =   240
      TabIndex        =   60
      Top             =   2880
      Width           =   9735
      Begin VB.CommandButton mconf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   7680
         TabIndex        =   108
         ToolTipText     =   "Graba los DATOS cargados"
         Top             =   360
         Width           =   1575
      End
      Begin TabDlg.SSTab mdatos 
         Height          =   3015
         Left            =   240
         TabIndex        =   61
         Top             =   960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5318
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmCliente.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "etiq(33)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "etiq(32)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "etiq(31)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "etiq(29)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "etiq(28)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "etiq(27)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "etiq(26)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "etiq(51)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "etiq(58)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "etiq(70)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "etiq(71)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "etiq(72)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "etiq(73)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "etiq(6)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "dtpNacimientoM"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "mapel"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "mnomb"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "mcuit"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "msdgi"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "mcate"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "mserv"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "mcobr"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "msoci"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtDocumentoM"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "cboDestinoM"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "cboCategoriaM"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "cboEstadoM"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtNombreCategoriaM"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).ControlCount=   28
         TabCaption(1)   =   "Ubicación del Inmueble"
         TabPicture(1)   =   "frmCliente.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "mior"
         Tab(1).Control(1)=   "miru"
         Tab(1).Control(2)=   "mizo"
         Tab(1).Control(3)=   "mico"
         Tab(1).Control(4)=   "mipr"
         Tab(1).Control(5)=   "milo"
         Tab(1).Control(6)=   "mide"
         Tab(1).Control(7)=   "mipi"
         Tab(1).Control(8)=   "mipu"
         Tab(1).Control(9)=   "mica"
         Tab(1).Control(10)=   "etiq(54)"
         Tab(1).Control(11)=   "etiq(53)"
         Tab(1).Control(12)=   "etiq(52)"
         Tab(1).Control(13)=   "etiq(34)"
         Tab(1).Control(14)=   "etiq(35)"
         Tab(1).Control(15)=   "etiq(36)"
         Tab(1).Control(16)=   "etiq(37)"
         Tab(1).Control(17)=   "etiq(38)"
         Tab(1).Control(18)=   "etiq(39)"
         Tab(1).Control(19)=   "etiq(40)"
         Tab(1).ControlCount=   20
         TabCaption(2)   =   "Domicilio del Cliente"
         TabPicture(2)   =   "frmCliente.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtFijoM"
         Tab(2).Control(1)=   "txtCelularM"
         Tab(2).Control(2)=   "txtCorreoM"
         Tab(2).Control(3)=   "mfco"
         Tab(2).Control(4)=   "mfpr"
         Tab(2).Control(5)=   "mflo"
         Tab(2).Control(6)=   "mfde"
         Tab(2).Control(7)=   "mfpi"
         Tab(2).Control(8)=   "mfpu"
         Tab(2).Control(9)=   "mfca"
         Tab(2).Control(10)=   "etiq(61)"
         Tab(2).Control(11)=   "etiq(60)"
         Tab(2).Control(12)=   "etiq(59)"
         Tab(2).Control(13)=   "etiq(41)"
         Tab(2).Control(14)=   "etiq(42)"
         Tab(2).Control(15)=   "etiq(43)"
         Tab(2).Control(16)=   "etiq(44)"
         Tab(2).Control(17)=   "etiq(45)"
         Tab(2).Control(18)=   "etiq(46)"
         Tab(2).Control(19)=   "etiq(47)"
         Tab(2).ControlCount=   20
         Begin VB.TextBox txtNombreCategoriaM 
            Height          =   285
            Left            =   240
            TabIndex        =   159
            Top             =   2520
            Width           =   3375
         End
         Begin VB.ComboBox cboEstadoM 
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox cboCategoriaM 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox cboDestinoM 
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtFijoM 
            Height          =   285
            Left            =   -71160
            TabIndex        =   131
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtCelularM 
            Height          =   285
            Left            =   -69360
            TabIndex        =   130
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtCorreoM 
            Height          =   285
            Left            =   -71160
            TabIndex        =   129
            Top             =   2520
            Width           =   3375
         End
         Begin VB.TextBox txtDocumentoM 
            Height          =   285
            Left            =   5640
            TabIndex        =   127
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox mior 
            Height          =   285
            Left            =   -69360
            TabIndex        =   78
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox miru 
            Height          =   285
            Left            =   -71160
            TabIndex        =   77
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox mizo 
            Height          =   285
            Left            =   -72960
            TabIndex        =   76
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox msoci 
            Height          =   285
            Left            =   3840
            TabIndex        =   82
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox mcobr 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox mserv 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox mcate 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox msdgi 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox mcuit 
            Height          =   285
            Left            =   3840
            TabIndex        =   81
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox mnomb 
            Height          =   285
            Left            =   3840
            TabIndex        =   80
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox mapel 
            Height          =   285
            Left            =   240
            TabIndex        =   79
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox mico 
            Height          =   285
            Left            =   -74760
            TabIndex        =   75
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox mipr 
            Height          =   285
            Left            =   -72960
            TabIndex        =   74
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox milo 
            Height          =   285
            Left            =   -74760
            TabIndex        =   73
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox mide 
            Height          =   285
            Left            =   -72960
            TabIndex        =   72
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox mipi 
            Height          =   285
            Left            =   -74760
            TabIndex        =   71
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox mipu 
            Height          =   285
            Left            =   -71160
            TabIndex        =   70
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox mica 
            Height          =   285
            Left            =   -74760
            TabIndex        =   69
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox mfco 
            Height          =   285
            Left            =   -74760
            TabIndex        =   68
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox mfpr 
            Height          =   285
            Left            =   -72960
            TabIndex        =   67
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox mflo 
            Height          =   285
            Left            =   -74760
            TabIndex        =   66
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox mfde 
            Height          =   285
            Left            =   -72960
            TabIndex        =   65
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox mfpi 
            Height          =   285
            Left            =   -74760
            TabIndex        =   64
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox mfpu 
            Height          =   285
            Left            =   -71160
            TabIndex        =   63
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox mfca 
            Height          =   285
            Left            =   -74760
            TabIndex        =   62
            Top             =   720
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker dtpNacimientoM 
            Height          =   375
            Left            =   7440
            TabIndex        =   151
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100663297
            CurrentDate     =   43135
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Categoría"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   160
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   73
            Left            =   7440
            TabIndex        =   158
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Categoría Socio"
            Height          =   195
            Index           =   72
            Left            =   5640
            TabIndex        =   156
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Destino Servicio"
            Height          =   195
            Index           =   71
            Left            =   5640
            TabIndex        =   154
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Nacimiento"
            Height          =   195
            Index           =   70
            Left            =   7440
            TabIndex        =   152
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono Fijo"
            Height          =   195
            Index           =   61
            Left            =   -71160
            TabIndex        =   134
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono Celular"
            Height          =   195
            Index           =   60
            Left            =   -69360
            TabIndex        =   133
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Correo Electrónico"
            Height          =   195
            Index           =   59
            Left            =   -71160
            TabIndex        =   132
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   58
            Left            =   5640
            TabIndex        =   128
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Orden Físico"
            Height          =   195
            Index           =   54
            Left            =   -69360
            TabIndex        =   123
            Top             =   2280
            Width           =   915
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
            Height          =   195
            Index           =   53
            Left            =   -71160
            TabIndex        =   122
            Top             =   2280
            Width           =   345
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            Height          =   195
            Index           =   52
            Left            =   -72960
            TabIndex        =   121
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número de Socio"
            Height          =   195
            Index           =   51
            Left            =   3840
            TabIndex        =   120
            Top             =   2280
            Width           =   1230
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cobro"
            Height          =   195
            Index           =   26
            Left            =   3840
            TabIndex        =   107
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Servicio"
            Height          =   195
            Index           =   27
            Left            =   2040
            TabIndex        =   106
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Categoría del Cliente"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   105
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Situación AFIP"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   104
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T."
            Height          =   195
            Index           =   31
            Left            =   3840
            TabIndex        =   103
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   32
            Left            =   3840
            TabIndex        =   102
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Apellido"
            Height          =   195
            Index           =   33
            Left            =   240
            TabIndex        =   101
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal"
            Height          =   195
            Index           =   34
            Left            =   -74760
            TabIndex        =   100
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            Height          =   195
            Index           =   35
            Left            =   -72960
            TabIndex        =   99
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Localidad"
            Height          =   195
            Index           =   36
            Left            =   -74760
            TabIndex        =   98
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   37
            Left            =   -72960
            TabIndex        =   97
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Piso"
            Height          =   195
            Index           =   38
            Left            =   -74760
            TabIndex        =   96
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   39
            Left            =   -71160
            TabIndex        =   95
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Index           =   40
            Left            =   -74760
            TabIndex        =   94
            Top             =   480
            Width           =   345
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal"
            Height          =   195
            Index           =   41
            Left            =   -74760
            TabIndex        =   93
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            Height          =   195
            Index           =   42
            Left            =   -72960
            TabIndex        =   92
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Localidad"
            Height          =   195
            Index           =   43
            Left            =   -74760
            TabIndex        =   91
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   44
            Left            =   -72960
            TabIndex        =   90
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Piso"
            Height          =   195
            Index           =   45
            Left            =   -74760
            TabIndex        =   89
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   46
            Left            =   -71160
            TabIndex        =   88
            Top             =   480
            Width           =   555
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Index           =   47
            Left            =   -74760
            TabIndex        =   87
            Top             =   480
            Width           =   345
         End
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de CONEXION"
         Height          =   195
         Index           =   25
         Left            =   480
         TabIndex        =   110
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label mcon 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   480
         TabIndex        =   109
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame vhis 
      Caption         =   "Histórico de Conexiones"
      Height          =   3015
      Left            =   480
      TabIndex        =   115
      Top             =   2880
      Width           =   7455
      Begin VB.TextBox hcon 
         Height          =   285
         Left            =   240
         TabIndex        =   116
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox hcli 
         Height          =   1815
         Left            =   240
         TabIndex        =   118
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de CONEXION"
         Height          =   195
         Index           =   49
         Left            =   240
         TabIndex        =   117
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame vbaj 
      Caption         =   "Baja de Clientes"
      Height          =   855
      Left            =   480
      TabIndex        =   111
      Top             =   2880
      Width           =   7455
      Begin VB.CommandButton bajac 
         Caption         =   "&Dar de Baja"
         Height          =   375
         Left            =   5640
         TabIndex        =   114
         ToolTipText     =   "Confirma la BAJA de la CONEXION"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label afech 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   113
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Alta"
         Height          =   195
         Index           =   48
         Left            =   840
         TabIndex        =   112
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del CLIENTE"
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   35
      Top             =   360
      Width           =   1530
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Conexión"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   34
      Top             =   360
      Width           =   660
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsin As Integer
Private vari As Boolean

Public Sub llenar()
Dim clienteRep As clsREPCliente

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    frmCliente.Height = rsin
    vagr.Visible = False
    vmod.Visible = False
    vbaj.Visible = False
    vhis.Visible = False
    
    Set clienteRep = New clsREPCliente
    clienteRep.fillCombo Me.ncli
    Set clienteRep = Nothing
    
    If ncli.ListCount Then
        ncli.ListIndex = 0
        histo.Enabled = True
        modif.Enabled = True
        dbaja.Enabled = True
    Else
        histo.Enabled = False
        modif.Enabled = False
        dbaja.Enabled = False
    End If
    
End Sub

Public Sub llenar_h()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If clienteRep.collectionActivos.Count = 0 Then Exit Sub
    
    hcli.Clear
    For Each cliente In clienteRep.collectionByClienteID(Val(hcon.Text))
        hcli.AddItem "Cliente : " & cliente.apellidonombre & " - Fecha de Alta : " & cliente.fechaAlta & " - Fecha de Baja : " & cliente.fechaBaja
    Next

End Sub

Private Sub aapel_Change()
    
    vari = True

End Sub

Private Sub aapel_GotFocus()
    
    aapel.SelStart = 0
    aapel.SelLength = Len(aapel.Text)

End Sub

Private Sub aapel_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        anomb.SetFocus
    End If

End Sub

Private Sub acata_Change()
    
    vari = True

End Sub

Private Sub acobr_Change()
    
    vari = True

End Sub

Private Sub acon_Change()
    
    vari = True

End Sub

Private Sub acuit_Change()
    
    vari = True

End Sub

Private Sub acuit_LostFocus()
    
    If Len(Trim(acuit.Text)) = 0 Then
        MsgBox "Debe escribir el número de CUIT"
        acuit.SetFocus
    End If

End Sub

Private Sub afca_Change()
    
    vari = True

End Sub

Private Sub afco_Change()
    
    vari = True

End Sub

Private Sub afde_Change()
    
    vari = True

End Sub

Private Sub aflo_Change()
    
    vari = True

End Sub

Private Sub afpi_Change()
    
    vari = True

End Sub

Private Sub afpr_Change()
    
    vari = True

End Sub

Private Sub afpu_Change()
    
    vari = True

End Sub

Private Sub aica_Change()
    
    vari = True
    afca.Text = aica.Text

End Sub

Private Sub aico_Change()
    
    vari = True
    afco.Text = aico.Text

End Sub

Private Sub aide_Change()
    
    vari = True
    afde.Text = aide.Text

End Sub

Private Sub ailo_Change()
    
    vari = True
    aflo.Text = ailo.Text

End Sub

Private Sub aior_Change()
    
    vari = True

End Sub

Private Sub aior_GotFocus()
    
    aior.SelStart = 0
    aior.SelLength = Len(aior.Text)

End Sub

Private Sub aipi_Change()
    
    vari = True
    afpi.Text = aipi.Text

End Sub

Private Sub aipr_Change()
    
    vari = True
    afpr.Text = aipr.Text

End Sub

Private Sub aipu_Change()
    
    vari = True
    afpu.Text = aipu.Text

End Sub

Private Sub airu_Change()
    
    vari = True

End Sub

Private Sub airu_GotFocus()
    
    airu.SelStart = 0
    airu.SelLength = Len(airu.Text)

End Sub

Private Sub aizo_Change()
    
    vari = True

End Sub

Private Sub aizo_GotFocus()
    
    aizo.SelStart = 0
    aizo.SelLength = Len(aizo.Text)

End Sub

Private Sub anomb_Change()
    
    vari = True

End Sub

Private Sub asdgi_Change()
    
    vari = True

End Sub

Private Sub asdgi_Click()
    
    If asdgi.ListIndex = 2 Then
        acuit.Text = ""
        acuit.Visible = False
        etiq(5).Visible = False
    Else
        acuit.Visible = True
        etiq(5).Visible = True
    End If

End Sub

Private Sub aserv_Change()
    
    vari = True

End Sub

Private Sub asoci_Change()
    
    vari = True

End Sub

Private Sub asoci_GotFocus()
    
    asoci.SelStart = 0
    asoci.SelLength = Len(asoci.Text)

End Sub

Private Sub asoci_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        adatos.Tab = 1
        aica.SetFocus
    End If

End Sub

Private Sub acon_GotFocus()
    
    acon.SelStart = 0
    acon.SelLength = Len(acon.Text)

End Sub

Private Sub acon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aapel.SetFocus

End Sub

Private Sub acon_LostFocus()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If Val(acon.Text) = 0 Then
        aconf.Enabled = False
        Exit Sub
    End If
    aconf.Enabled = True
    
    Set cliente = clienteRep.findLastByClienteID(Val(acon.Text))
    
    If IsNull(cliente.uniqueId) Then
        aapel.SetFocus
        Exit Sub
    End If
    If Not IsNull(cliente.fechaBaja) Then
        aapel.SetFocus
        Exit Sub
    End If
    
    MsgBox "Esta CONEXION ya Existe"
    vagr.Visible = False
    vmod.Visible = False
    vbaj.Visible = False
    vhis.Visible = False
    acon.Text = ""
    vari = False
    frmCliente.Height = rsin

End Sub

Private Sub aconf_Click()
Dim cui As String

Dim ct As Integer

Dim cliente As New clsMODCliente
Dim clientedato As New clsMyAClienteDato

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If asdgi.ListIndex <> 2 Then
        If Len(Trim(acuit.Text)) = 0 Then
            MsgBox "Debe ingresar el número de CUIT"
            adatos.Tab = 0
            acuit.SetFocus
            Exit Sub
        End If
    End If
        
    cliente.clienteId = Val(acon.Text)
    cliente.fechaAlta = Date
    cliente.apellido = Left(aapel.Text, 25)
    cliente.nombre = Left(anomb.Text, 25)
    cui = ""
    For ct = 1 To Len(acuit.Text)
        If Mid(acuit.Text, ct, 1) <> "-" Then cui = cui & Mid(acuit.Text, ct, 1)
    Next ct
    cliente.cuit = Left(cui, 11)
    cliente.numeroSocio = Left(asoci.Text, 6)
    cliente.inmuebleCalle = Left(aica.Text, 20)
    cliente.inmueblePuerta = Left(aipu.Text, 5)
    cliente.inmueblePiso = Left(aipi.Text, 3)
    cliente.inmuebleDpto = Left(aide.Text, 4)
    cliente.inmuebleLocalidad = Left(ailo.Text, 30)
    cliente.inmuebleProvincia = Left(aipr.Text, 15)
    cliente.inmuebleCodpostal = Val(aico.Text)
    cliente.fiscalCalle = Left(afca.Text, 20)
    cliente.fiscalPuerta = Left(afpu.Text, 5)
    cliente.fiscalPiso = Left(afpi.Text, 3)
    cliente.fiscalDpto = Left(afde.Text, 4)
    cliente.fiscalLocalidad = Left(aflo.Text, 30)
    cliente.fiscalProvincia = Left(afpr.Text, 15)
    cliente.fiscalCodpostal = Val(afco.Text)
    cliente.nombrecategoria = Me.txtNombreCategoriaA.Text
    cliente.situacionIVA = asdgi.ListIndex + 1
    cliente.cobro = acobr.ListIndex + 1
    cliente.servicio = aserv.ListIndex + 1
    cliente.categoria = acate.ListIndex + 1
    cliente.fechaAlta = Date
    cliente.zona = Val(aizo.Text)
    cliente.ruta = Val(airu.Text)
    cliente.orden = Val(aior.Text)
    cliente.estadoID = Me.cboEstadoA.ItemData(Me.cboEstadoA.ListIndex)
    cliente.fechaNacimiento = Me.dtpNacimientoA.value
    cliente.categoriasocioID = Me.cboCategoriaA.ItemData(Me.cboCategoriaA.ListIndex)
    cliente.destinoID = Me.cboDestinoA.ItemData(Me.cboDestinoA.ListIndex)
    cliente.uid = "admin"
    Set cliente = clienteRep.save(cliente)
    
    clientedato.clienteId = cliente.clienteId
    clientedato.documento = Val(Me.txtDocumentoA.Text)
    clientedato.fijo = Trim(Me.txtFijoA.Text)
    clientedato.celular = Trim(Me.txtCelularA.Text)
    clientedato.email = Trim(Me.txtCorreoA.Text)
    
    clientedato.save dbapp
    
    vari = False
    
    llenar
    
    modif.Enabled = True
    histo.Enabled = True
    dbaja.Enabled = True
    agreg.SetFocus

End Sub

Private Sub acuit_GotFocus()
    
    acuit.SelStart = 0
    acuit.SelLength = Len(acuit.Text)

End Sub

Private Sub acuit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then asoci.SetFocus

End Sub

Private Sub afca_GotFocus()
    
    afca.SelStart = 0
    afca.SelLength = Len(afca.Text)

End Sub

Private Sub afca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then afpu.SetFocus

End Sub

Private Sub afco_GotFocus()
    
    afco.SelStart = 0
    afco.SelLength = Len(afco.Text)

End Sub

Private Sub afco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aconf.SetFocus

End Sub

Private Sub afde_GotFocus()
    
    afde.SelStart = 0
    afde.SelLength = Len(afde.Text)

End Sub

Private Sub afde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aflo.SetFocus

End Sub

Private Sub aflo_GotFocus()
    
    aflo.SelStart = 0
    aflo.SelLength = Len(aflo.Text)

End Sub

Private Sub aflo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then afpr.SetFocus

End Sub

Private Sub afpi_GotFocus()
    
    afpi.SelStart = 0
    afpi.SelLength = Len(afpi.Text)

End Sub

Private Sub afpi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then afde.SetFocus

End Sub

Private Sub afpr_GotFocus()
    
    afpr.SelStart = 0
    afpr.SelLength = Len(afpr.Text)

End Sub

Private Sub afpr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then afco.SetFocus

End Sub

Private Sub afpu_GotFocus()
    
    afpu.SelStart = 0
    afpu.SelLength = Len(afpu.Text)

End Sub

Private Sub afpu_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then afpi.SetFocus

End Sub

Private Sub agreg_Click()
Dim nv As Integer

Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    
    frmCliente.Height = rsin + 5000
    vagr.Visible = True
    vhis.Visible = False
    vmod.Visible = False
    vbaj.Visible = False
    
    aapel.Text = ""
    anomb.Text = ""
    acuit.Text = ""
    asoci.Text = ""
    aica.Text = ""
    aipu.Text = ""
    aipi.Text = ""
    aide.Text = ""
    ailo.Text = ""
    aipr.Text = "Mendoza"
    aico.Text = ""
    aizo.Text = ""
    airu.Text = ""
    aior.Text = ""
    afca.Text = ""
    afpu.Text = ""
    afpi.Text = ""
    afde.Text = ""
    aflo.Text = ""
    afpr.Text = "Mendoza"
    afco.Text = ""
    Me.txtDocumentoA.Text = ""
    Me.txtFijoA.Text = ""
    Me.txtCelularA.Text = ""
    Me.txtCorreoA.Text = ""
    Me.txtNombreCategoriaA.Text = ""
    asdgi.ListIndex = 2
    acuit.Visible = False
    acobr.ListIndex = 0
    aserv.ListIndex = 0
    acate.ListIndex = 0
    Me.cboEstadoA.ListIndex = 0
    Me.dtpNacimientoA.value = Date
    Me.cboCategoriaA.ListIndex = 0
    Me.cboDestinoA.ListIndex = 0
    adatos.Tab = 0
    vari = False
    
    nv = 0
    If clienteRep.collectionActivos.Count > 0 Then
        Set cliente = clienteRep.findLastLast
        nv = cliente.clienteId
    End If
    acon.Text = nv + 1
    acon.SetFocus

End Sub

Private Sub aica_GotFocus()
    
    aica.SelStart = 0
    aica.SelLength = Len(aica.Text)

End Sub

Private Sub aica_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aipu.SetFocus

End Sub

Private Sub aico_GotFocus()
    
    aico.SelStart = 0
    aico.SelLength = Len(aico.Text)

End Sub

Private Sub aico_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        adatos.Tab = 2
        afca.SetFocus
    End If

End Sub

Private Sub aide_GotFocus()
    
    aide.SelStart = 0
    aide.SelLength = Len(aide.Text)

End Sub

Private Sub aide_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ailo.SetFocus

End Sub

Private Sub ailo_GotFocus()
    
    ailo.SelStart = 0
    ailo.SelLength = Len(ailo.Text)

End Sub

Private Sub ailo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aipr.SetFocus

End Sub

Private Sub aipi_GotFocus()
    
    aipi.SelStart = 0
    aipi.SelLength = Len(aipi.Text)

End Sub

Private Sub aipi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aide.SetFocus

End Sub

Private Sub aipr_GotFocus()
    
    aipr.SelStart = 0
    aipr.SelLength = Len(aipr.Text)

End Sub

Private Sub aipr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aico.SetFocus

End Sub

Private Sub aipu_GotFocus()
    
    aipu.SelStart = 0
    aipu.SelLength = Len(aipu.Text)

End Sub

Private Sub aipu_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aipi.SetFocus

End Sub

Private Sub anomb_GotFocus()
    
    anomb.SelStart = 0
    anomb.SelLength = Len(anomb.Text)

End Sub

Private Sub anomb_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then acate.SetFocus

End Sub

Private Sub bajac_Click()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    cliente.fechaBaja = Date
    cliente.uid = "admin"
    Set cliente = clienteRep.save(cliente)
    
    llenar

End Sub

Private Sub cboCategoriaA_Change()

    vari = True

End Sub

Private Sub cboCategoriaM_Change()

    vari = True

End Sub

Private Sub cboDestinoA_Change()

    vari = True

End Sub

Private Sub cboDestinoM_Change()

    vari = True
    
End Sub

Private Sub cboEstadoA_Change()

    vari = True

End Sub

Private Sub cboEstadoM_Change()

    vari = True

End Sub

Private Sub dbaja_Click()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    
    frmCliente.Height = rsin + 1600
    
    vbaj.Visible = True
    vhis.Visible = False
    vagr.Visible = False
    vmod.Visible = False
        
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    afech.Caption = cliente.fechaAlta
    
    bajac.SetFocus

End Sub

Private Sub dtpNacimientoA_Change()

    vari = True

End Sub

Private Sub dtpNacimientoM_Change()

    vari = True

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim clienteRep As New clsREPCliente
    
    If clienteRep.collectionActivos.Count = 0 Then
        modif.Enabled = False
        histo.Enabled = False
        dbaja.Enabled = False
    End If
    
    llenar

End Sub

Private Sub Form_Load()
Dim estado As New clsMyAEstado
Dim destino As New clsMyADestinoServ
Dim categoria As New clsMyACategoriaSocio
    
    rsin = 3100
    mcate.Clear
    mcate.AddItem "General"
    mcate.AddItem "Especial"
    mcate.ListIndex = 0
    mserv.Clear
    mserv.AddItem "Agua"
    mserv.AddItem "Cloaca"
    mserv.AddItem "Agua y Cloaca"
    mserv.ListIndex = 0
    mcobr.Clear
    mcobr.AddItem "Servicio Medido"
    mcobr.AddItem "Cuota Fija en Trans."
    mcobr.AddItem "Cuota Fija"
    mcobr.ListIndex = 0
    msdgi.Clear
    msdgi.AddItem "Responsable Inscripto"
    msdgi.AddItem "Responsable No Inscripto"
    msdgi.AddItem "Consumidor Final"
    msdgi.AddItem "IVA Exento"
    msdgi.AddItem "IVA No Responsable"
    msdgi.AddItem "Responsable Monotributo"
    msdgi.ListIndex = 2
    estado.fillCombo Me.cboEstadoM, dbapp
    categoria.fillCombo Me.cboCategoriaM, dbapp
    destino.fillCombo Me.cboDestinoM, dbapp
    acate.Clear
    acate.AddItem "General"
    acate.AddItem "Especial"
    acate.ListIndex = 0
    aserv.Clear
    aserv.AddItem "Agua"
    aserv.AddItem "Cloaca"
    aserv.AddItem "Agua y Cloaca"
    aserv.ListIndex = 0
    acobr.Clear
    acobr.AddItem "Servicio Medido"
    acobr.AddItem "Cuota Fija en Trans."
    acobr.AddItem "Cuota Fija"
    acobr.ListIndex = 0
    asdgi.Clear
    asdgi.AddItem "Responsable Inscripto"
    asdgi.AddItem "Responsable No Inscripto"
    asdgi.AddItem "Consumidor Final"
    asdgi.AddItem "iva Exento"
    asdgi.AddItem "iva No Responsable"
    asdgi.AddItem "Responsable Monotributo"
    asdgi.ListIndex = 0
    estado.fillCombo Me.cboEstadoA, dbapp
    categoria.fillCombo Me.cboCategoriaA, dbapp
    destino.fillCombo Me.cboDestinoA, dbapp

End Sub

Private Sub hcon_GotFocus()
    
    hcon.SelStart = 0
    hcon.SelLength = Len(hcon.Text)

End Sub

Private Sub hcon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then hcon_LostFocus

End Sub

Private Sub hcon_LostFocus()
    
    llenar_h

End Sub

Private Sub histo_Click()
    
    vhis.Visible = True
    vbaj.Visible = False
    vagr.Visible = False
    vmod.Visible = False
    frmCliente.Height = rsin + 3800
    hcon.Text = ncon.Text
    llenar_h
    hcon.SetFocus

End Sub

Private Sub mapel_Change()
    
    vari = True

End Sub

Private Sub mapel_GotFocus()
    
    mapel.SelStart = 0
    mapel.SelLength = Len(mapel.Text)

End Sub

Private Sub mapel_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mnomb.SetFocus

End Sub

Private Sub mcate_Change()
    
    vari = True

End Sub

Private Sub mcobr_Change()
    
    vari = True

End Sub

Private Sub mcuit_Change()
    
    vari = True

End Sub

Private Sub mcuit_LostFocus()
    
    If Len(Trim(mcuit.Text)) = 0 Then
        MsgBox "Debe ingresar el número de CUIT"
        mcuit.SetFocus
    End If

End Sub

Private Sub mfca_Change()
    
    vari = True

End Sub

Private Sub mfco_Change()
    
    vari = True

End Sub

Private Sub mfde_Change()
    
    vari = True

End Sub

Private Sub mflo_Change()
    
    vari = True

End Sub

Private Sub mfpi_Change()
    
    vari = True

End Sub

Private Sub mfpr_Change()
    
    vari = True

End Sub

Private Sub mfpu_Change()
    
    vari = True

End Sub

Private Sub mica_Change()
    
    vari = True
    mfca.Text = mica.Text

End Sub

Private Sub mico_Change()
    
    vari = True
    mfco.Text = mico.Text

End Sub

Private Sub mide_Change()
    
    vari = True
    mfde.Text = mide.Text

End Sub

Private Sub milo_Change()
    
    vari = True
    mflo.Text = milo.Text

End Sub

Private Sub mior_Change()
    
    vari = True

End Sub

Private Sub mior_GotFocus()
    
    mior.SelStart = 0
    mior.SelLength = Len(mior.Text)

End Sub

Private Sub mipi_Change()
    
    vari = True
    mfpi.Text = mipi.Text

End Sub

Private Sub mipr_Change()
    
    vari = True
    mfpr.Text = mipr.Text

End Sub

Private Sub mipu_Change()
    
    vari = True
    mfpu.Text = mipu.Text

End Sub

Private Sub miru_Change()
    
    vari = True

End Sub

Private Sub miru_GotFocus()
    
    miru.SelStart = 0
    miru.SelLength = Len(miru.Text)

End Sub

Private Sub mizo_Change()
    
    vari = True

End Sub

Private Sub mizo_GotFocus()
    
    mizo.SelStart = 0
    mizo.SelLength = Len(mizo.Text)

End Sub

Private Sub mnomb_Change()
    
    vari = True

End Sub

Private Sub msdgi_Change()
    
    vari = True

End Sub

Private Sub msdgi_Click()
    
    If msdgi.ListIndex = 2 Then
        mcuit.Text = ""
        mcuit.Visible = False
        etiq(31).Visible = False
    Else
        mcuit.Visible = True
        etiq(31).Visible = True
    End If

End Sub

Private Sub mserv_Change()
    
    vari = True

End Sub

Private Sub msoci_Change()
    
    vari = True

End Sub

Private Sub msoci_GotFocus()
    
    msoci.SelStart = 0
    msoci.SelLength = Len(msoci.Text)

End Sub

Private Sub msoci_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        mdatos.Tab = 1
        mica.SetFocus
    End If

End Sub

Private Sub mconf_Click()
Dim cui As String
Dim ct As Integer

Dim cliente As clsMODCliente
Dim clientedato As New clsMyAClienteDato

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If msdgi.ListIndex <> 2 Then
        If Len(Trim(mcuit.Text)) = 0 Then
            MsgBox "Debe ingresar el número de CUIT"
            mdatos.Tab = 0
            mcuit.SetFocus
            Exit Sub
        End If
    End If
    
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    
    cliente.apellido = Left(mapel.Text, 25)
    cliente.nombre = Left(mnomb.Text, 25)
    cui = ""
    For ct = 1 To Len(mcuit.Text)
        If Mid(mcuit.Text, ct, 1) <> "-" Then cui = cui & Mid(mcuit.Text, ct, 1)
    Next ct
    cliente.cuit = Left(cui, 11)
    cliente.numeroSocio = Left(msoci.Text, 6)
    cliente.inmuebleCalle = Left(mica.Text, 20)
    cliente.inmueblePuerta = Left(mipu.Text, 5)
    cliente.inmueblePiso = Left(mipi.Text, 3)
    cliente.inmuebleDpto = Left(mide.Text, 4)
    cliente.inmuebleLocalidad = Left(milo.Text, 30)
    cliente.inmuebleProvincia = Left(mipr.Text, 15)
    cliente.inmuebleCodpostal = Val(mico.Text)
    cliente.fiscalCalle = Left(mfca.Text, 20)
    cliente.fiscalPuerta = Left(mfpu.Text, 5)
    cliente.fiscalPiso = Left(mfpi.Text, 3)
    cliente.fiscalDpto = Left(mfde.Text, 4)
    cliente.fiscalLocalidad = Left(mflo.Text, 30)
    cliente.fiscalProvincia = Left(mfpr.Text, 15)
    cliente.fiscalCodpostal = Val(mfco.Text)
    cliente.nombrecategoria = Me.txtNombreCategoriaM.Text
    cliente.situacionIVA = msdgi.ListIndex + 1
    cliente.cobro = mcobr.ListIndex + 1
    cliente.servicio = mserv.ListIndex + 1
    cliente.categoria = mcate.ListIndex + 1
    cliente.zona = Val(mizo.Text)
    cliente.ruta = Val(miru.Text)
    cliente.orden = Val(mior.Text)
    cliente.estadoID = Me.cboEstadoM.ItemData(Me.cboEstadoM.ListIndex)
    cliente.fechaNacimiento = Me.dtpNacimientoM.value
    cliente.categoriasocioID = Me.cboCategoriaM.ItemData(Me.cboCategoriaM.ListIndex)
    cliente.destinoID = Me.cboDestinoM.ItemData(Me.cboDestinoM.ListIndex)
    cliente.uid = "admin"
    Set cliente = clienteRep.update(cliente, cliente.uniqueId)

    With clientedato
        .clienteId = cliente.clienteId
        .documento = Val(Me.txtDocumentoM.Text)
        .fijo = Trim(Me.txtFijoM.Text)
        .celular = Trim(Me.txtCelularM.Text)
        .email = Trim(Me.txtCorreoM.Text)
        
        .save dbapp
    End With

    vari = False
    
    llenar
    
    agreg.SetFocus

End Sub

Private Sub mcuit_GotFocus()
    
    mcuit.SelStart = 0
    mcuit.SelLength = Len(mcuit.Text)

End Sub

Private Sub mcuit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then msoci.SetFocus

End Sub

Private Sub mfca_GotFocus()
    
    mfca.SelStart = 0
    mfca.SelLength = Len(mfca.Text)

End Sub

Private Sub mfca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mfpu.SetFocus

End Sub

Private Sub mfco_GotFocus()
    
    mfco.SelStart = 0
    mfco.SelLength = Len(mfco.Text)

End Sub

Private Sub mfco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mconf.SetFocus

End Sub

Private Sub mfde_GotFocus()
    
    mfde.SelStart = 0
    mfde.SelLength = Len(mfde.Text)

End Sub

Private Sub mfde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mflo.SetFocus

End Sub

Private Sub mflo_GotFocus()
    
    mflo.SelStart = 0
    mflo.SelLength = Len(mflo.Text)

End Sub

Private Sub mflo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mfpr.SetFocus

End Sub

Private Sub mfpi_GotFocus()
    
    mfpi.SelStart = 0
    mfpi.SelLength = Len(mfpi.Text)

End Sub

Private Sub mfpi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mfde.SetFocus

End Sub

Private Sub mfpr_GotFocus()
    
    mfpr.SelStart = 0
    mfpr.SelLength = Len(mfpr.Text)

End Sub

Private Sub mfpr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mfco.SetFocus

End Sub

Private Sub mfpu_GotFocus()
    
    mfpu.SelStart = 0
    mfpu.SelLength = Len(mfpu.Text)

End Sub

Private Sub mfpu_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mfpi.SetFocus

End Sub

Private Sub mica_GotFocus()
    
    mica.SelStart = 0
    mica.SelLength = Len(mica.Text)

End Sub

Private Sub mica_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mipu.SetFocus

End Sub

Private Sub mico_GotFocus()
    
    mico.SelStart = 0
    mico.SelLength = Len(mico.Text)

End Sub

Private Sub mico_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        mdatos.Tab = 2
        mfca.SetFocus
    End If

End Sub

Private Sub mide_GotFocus()
    
    mide.SelStart = 0
    mide.SelLength = Len(mide.Text)

End Sub

Private Sub mide_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then milo.SetFocus

End Sub

Private Sub milo_GotFocus()
    
    milo.SelStart = 0
    milo.SelLength = Len(milo.Text)

End Sub

Private Sub milo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mipr.SetFocus

End Sub

Private Sub mipi_GotFocus()
    
    mipi.SelStart = 0
    mipi.SelLength = Len(mipi.Text)

End Sub

Private Sub mipi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mide.SetFocus

End Sub

Private Sub mipr_GotFocus()
    
    mipr.SelStart = 0
    mipr.SelLength = Len(mipr.Text)

End Sub

Private Sub mipr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mico.SetFocus

End Sub

Private Sub mipu_GotFocus()
    
    mipu.SelStart = 0
    mipu.SelLength = Len(mipu.Text)

End Sub

Private Sub mipu_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mipi.SetFocus

End Sub

Private Sub mnomb_GotFocus()
    
    mnomb.SelStart = 0
    mnomb.SelLength = Len(mnomb.Text)

End Sub

Private Sub mnomb_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mcate.SetFocus

End Sub

Private Sub modif_Click()
Dim crit, cui As String
Dim ct As Integer

Dim cliente As clsMODCliente
Dim clientedato As New clsMyAClienteDato
Dim estado As New clsMyAEstado
Dim destino As New clsMyADestinoServ
Dim categoria As New clsMyACategoriaSocio

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    frmCliente.Height = rsin + 5000
    vmod.Visible = True
    vhis.Visible = False
    vagr.Visible = False
    vbaj.Visible = False
    mdatos.Tab = 0
    mcon.Caption = ncon.Text
    
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    If IsNull(cliente.apellido) Then
        mapel.Text = ""
    Else
        mapel.Text = cliente.apellido
    End If
    If IsNull(cliente.nombre) Then
        mnomb.Text = ""
    Else
        mnomb.Text = cliente.nombre
    End If
    If IsNull(cliente.cuit) Then
        mcuit.Text = ""
    Else
        cui = Mid(cliente.cuit, 1, 2) & "-" & Mid(cliente.cuit, 3, 8) & "-" & Right(cliente.cuit, 1)
        mcuit.Text = cui
    End If
    If IsNull(cliente.numeroSocio) Then
        msoci.Text = ""
    Else
        msoci.Text = cliente.numeroSocio
    End If
    If IsNull(cliente.inmuebleCalle) Then
        mica.Text = ""
    Else
        mica.Text = cliente.inmuebleCalle
    End If
    If IsNull(cliente.inmueblePuerta) Then
        mipu.Text = ""
    Else
        mipu.Text = cliente.inmueblePuerta
    End If
    If IsNull(cliente.inmueblePiso) Then
        mipi.Text = ""
    Else
        mipi.Text = cliente.inmueblePiso
    End If
    If IsNull(cliente.inmuebleDpto) Then
        mide.Text = ""
    Else
        mide.Text = cliente.inmuebleDpto
    End If
    If IsNull(cliente.inmuebleLocalidad) Then
        milo.Text = ""
    Else
        milo.Text = cliente.inmuebleLocalidad
    End If
    If IsNull(cliente.inmuebleProvincia) Then
        mipr.Text = ""
    Else
        mipr.Text = cliente.inmuebleProvincia
    End If
    If IsNull(cliente.inmuebleCodpostal) Then
        mico.Text = ""
    Else
        mico.Text = cliente.inmuebleCodpostal
    End If
    If IsNull(cliente.zona) Then
        mizo.Text = ""
    Else
        mizo.Text = cliente.zona
    End If
    If IsNull(cliente.ruta) Then
        miru.Text = ""
    Else
        miru.Text = cliente.ruta
    End If
    If IsNull(cliente.orden) Then
        mior.Text = ""
    Else
        mior.Text = cliente.orden
    End If
    If IsNull(cliente.fiscalCalle) Then
        mfca.Text = ""
    Else
        mfca.Text = cliente.fiscalCalle
    End If
    If IsNull(cliente.fiscalPuerta) Then
        mfpu.Text = ""
    Else
        mfpu.Text = cliente.fiscalPuerta
    End If
    If IsNull(cliente.fiscalPiso) Then
        mfpi.Text = ""
    Else
        mfpi.Text = cliente.fiscalPiso
    End If
    If IsNull(cliente.fiscalDpto) Then
        mfde.Text = ""
    Else
        mfde.Text = cliente.fiscalDpto
    End If
    If IsNull(cliente.fiscalLocalidad) Then
        mflo.Text = ""
    Else
        mflo.Text = cliente.fiscalLocalidad
    End If
    If IsNull(cliente.fiscalProvincia) Then
        mfpr.Text = ""
    Else
        mfpr.Text = cliente.fiscalProvincia
    End If
    If IsNull(cliente.fiscalCodpostal) Then
        mfco.Text = ""
    Else
        mfco.Text = cliente.fiscalCodpostal
    End If
    If IsNull(cliente.situacionIVA) Then
        msdgi.ListIndex = 2
    Else
        msdgi.ListIndex = cliente.situacionIVA - 1
    End If
    If IsNull(cliente.cobro) Then
        mcobr.ListIndex = 0
    Else
        mcobr.ListIndex = cliente.cobro - 1
    End If
    If IsNull(cliente.servicio) Then
        mserv.ListIndex = 0
    Else
        mserv.ListIndex = cliente.servicio - 1
    End If
    If IsNull(cliente.categoria) Then
        mcate.ListIndex = 0
    Else
        mcate.ListIndex = cliente.categoria - 1
    End If
    If IsNull(cliente.nombrecategoria) Then
        Me.txtNombreCategoriaM.Text = ""
    Else
        Me.txtNombreCategoriaM.Text = cliente.nombrecategoria
    End If
    Me.dtpNacimientoM.value = cliente.fechaNacimiento
    
    estado.estadoID = cliente.estadoID
    estado.findByPrimaryKey dbapp
    categoria.categoriasocioID = cliente.categoriasocioID
    categoria.findByPrimaryKey dbapp
    destino.destinoID = cliente.destinoID
    destino.findByPrimaryKey dbapp
    Me.cboEstadoM.Text = estado.comboText
    Me.cboCategoriaM.Text = categoria.comboText
    Me.cboDestinoM.Text = destino.comboText
    With clientedato
        .clienteId = cliente.clienteId
        .findByPrimaryKey dbapp
        
        Me.txtDocumentoM.Text = .documento
        Me.txtFijoM.Text = .fijo
        Me.txtCelularM.Text = .celular
        Me.txtCorreoM.Text = .email
    End With
    mdatos.Tab = 0
    vari = False
    mapel.SetFocus

End Sub

Private Sub ncon_GotFocus()

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    vagr.Visible = False
    vmod.Visible = False
    vbaj.Visible = False
    vhis.Visible = False
    frmCliente.Height = rsin
    ncon.SelStart = 0
    ncon.SelLength = Len(ncon.Text)

End Sub

Private Sub ncon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then histo.SetFocus

End Sub

Private Sub ncon_LostFocus()
Dim clienteId As Long

Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If clienteRep.collectionActivos.Count = 0 Then Exit Sub
    
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    If IsNull(cliente.uniqueId) Then
        Me.ncli.ListIndex = 0
        Exit Sub
    End If
    If Not IsNull(cliente.fechaBaja) Then
        MsgBox "Cliente Dado de Baja . . ."
        Exit Sub
    End If
    clienteId = Val(ncon.Text)
    ncli.Text = cliente.comboText
    
    Do While ncli.ItemData(ncli.ListIndex) <> clienteId
        ncli.ListIndex = ncli.ListIndex + 1
    Loop
    histo.Enabled = True
    modif.Enabled = True
    dbaja.Enabled = True

End Sub

Private Sub ncli_Click()

'On Error Resume Next
    
    ncon.Text = ncli.ItemData(ncli.ListIndex)

End Sub

Private Sub ncli_GotFocus()

On Error Resume Next
    
    If vari Then
        If MsgBox("Descarta los datos ingresados ?", vbYesNo + vbDefaultButton2, "ADVERTENCIA") = vbYes Then
            vari = False
        Else
            Exit Sub
        End If
    End If
    vagr.Visible = False
    vmod.Visible = False
    vbaj.Visible = False
    vhis.Visible = False
    frmCliente.Height = rsin
    
End Sub

Private Sub txtCelularA_GotFocus()

    marcarseleccion Me.txtCelularA
    
End Sub

Private Sub txtCelularM_GotFocus()

    marcarseleccion Me.txtCelularM
    
End Sub

Private Sub txtCorreoA_GotFocus()

    marcarseleccion Me.txtCorreoA
    
End Sub

Private Sub txtCorreoM_GotFocus()

    marcarseleccion Me.txtCorreoM
    
End Sub

Private Sub txtDocumentoA_GotFocus()

    marcarseleccion Me.txtDocumentoA
    
End Sub

Private Sub txtDocumentoM_GotFocus()

    marcarseleccion Me.txtDocumentoM
    
End Sub

Private Sub txtFijoA_GotFocus()

    marcarseleccion Me.txtFijoA
    
End Sub

Private Sub txtFijoM_GotFocus()

    marcarseleccion Me.txtFijoM
    
End Sub

Private Sub txtNombreCategoriaA_Change()

    vari = True
    
End Sub

Private Sub txtNombreCategoriaA_GotFocus()

    marcarseleccion Me.txtNombreCategoriaA
    
End Sub

Private Sub txtNombreCategoriaM_Change()

    vari = True
    
End Sub

Private Sub txtNombreCategoriaM_GotFocus()

    marcarseleccion Me.txtNombreCategoriaM
    
End Sub
