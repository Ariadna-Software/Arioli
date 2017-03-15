VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Varios"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePaletMovimImprimir 
      Height          =   4455
      Left            =   840
      TabIndex        =   112
      Top             =   1080
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txtPalot 
         Height          =   285
         Index           =   4
         Left            =   6000
         TabIndex        =   126
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtPalot 
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   124
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtPalot 
         Height          =   285
         Index           =   2
         Left            =   3600
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPalot 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtPalot 
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   118
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optMovPalot 
         Caption         =   "Documento carga"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   117
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optMovPalot 
         Caption         =   "Albarán"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   116
         Top             =   1200
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdImpresionMovPalet 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   114
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   13
         Left            =   6000
         TabIndex        =   113
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Matricula"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   127
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image ImgTransporte 
         Height          =   255
         Left            =   3360
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   125
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "DNI"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   123
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Conductor"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   121
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Destino"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   119
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Impresión movimiento palot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Index           =   12
         Left            =   240
         TabIndex        =   115
         Top             =   360
         Width           =   4515
      End
   End
   Begin VB.Frame FrameModificaKilosDeposito 
      Height          =   3735
      Left            =   2520
      TabIndex        =   105
      Top             =   960
      Width           =   4815
      Begin VB.CommandButton cmdKilosDeposito 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2160
         TabIndex        =   107
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1440
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   12
         Left            =   3480
         TabIndex        =   108
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   111
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   110
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Modificar kilos depósito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   11
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.Frame FramePlanning 
      Height          =   6855
      Left            =   120
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   79
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtArt 
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtArtD 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   240
            Width           =   4215
         End
         Begin VB.Image imgArticulo 
            Height          =   240
            Index           =   4
            Left            =   6000
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   9
         Left            =   6000
         TabIndex        =   76
         Top             =   6240
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4455
         Left            =   240
         TabIndex        =   77
         Top             =   1560
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7858
         SortKey         =   5
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vta"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Referencia/Proveedor"
            Object.Width           =   4480
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Uds"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "FechaOculta"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   78
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Planning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   7
         Left            =   2760
         TabIndex        =   75
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame FrameHomologa 
      Height          =   5775
      Left            =   120
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtHomologa 
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Text            =   "frmVarios.frx":0000
         Top             =   1440
         Width           =   5175
      End
      Begin VB.CommandButton cmdHomologacion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   102
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   103
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   840
         Picture         =   "frmVarios.frx":0006
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   104
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   10
         Left            =   480
         TabIndex        =   99
         Top             =   360
         Width           =   4845
      End
   End
   Begin VB.Frame FrameAccionesRealizadas 
      Height          =   5775
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.OptionButton optVarios 
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   88
         Top             =   5040
         Width           =   1455
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   87
         Top             =   5040
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdAccionesRealizada 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   86
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   4440
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   4680
         TabIndex        =   83
         Top             =   5280
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3255
         Index           =   0
         Left            =   240
         TabIndex        =   96
         Top             =   1560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5468
         EndProperty
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3255
         Index           =   1
         Left            =   3840
         TabIndex        =   97
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   95
         Top             =   5400
         Width           =   2985
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   4800
         Picture         =   "frmVarios.frx":0091
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   5160
         Picture         =   "frmVarios.frx":01DB
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   1200
         Picture         =   "frmVarios.frx":0325
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   1560
         Picture         =   "frmVarios.frx":046F
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   94
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Acciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   93
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Top             =   720
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   4200
         Picture         =   "frmVarios.frx":05B9
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   1080
         TabIndex        =   91
         Top             =   720
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1560
         Picture         =   "frmVarios.frx":0644
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   3720
         TabIndex        =   90
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado acciones realizadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   9
         Left            =   600
         TabIndex        =   89
         Top             =   120
         Width           =   4845
      End
   End
   Begin VB.Frame FrameDHArticulo 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdEliminarArticulos 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   3
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar artículos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameImpresionFacturasDirectas 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Impresión facturas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   8
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameListArticulos 
      Height          =   6855
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6855
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAccionListview 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmVarios.frx":06CF
         ToolTipText     =   "Quitar seleccion"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmVarios.frx":0819
         ToolTipText     =   "Seleccionar todos"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar artículos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameEstadisticasConsultas 
      Height          =   3855
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdListConsultaPedido 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   29
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   1
         Left            =   3720
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   0
         Left            =   1200
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   38
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   37
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   3
         Left            =   1200
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   35
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Estadísticas consultas artículo / cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   750
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   1200
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame FrameListRevision 
      Height          =   2655
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtTrabDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdGeneListaRevision 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   49
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   6
         Left            =   4920
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         Width           =   945
      End
      Begin VB.Image imgTrab 
         Height          =   255
         Left            =   1080
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Crear nueva lista revision"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   4
         Left            =   720
         TabIndex        =   45
         Top             =   360
         Width           =   4875
      End
   End
   Begin VB.Frame FrameMarcas 
      Height          =   6135
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdMarcas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   43
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   42
         Top             =   5640
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4695
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmVarios.frx":0963
         ToolTipText     =   "Seleccionar todos"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmVarios.frx":0AAD
         ToolTipText     =   "Quitar seleccion"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Seleccionar marcas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   4755
      End
   End
   Begin VB.Frame FramePasswd 
      Height          =   3255
      Left            =   1080
      TabIndex        =   62
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdCambiopass 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2160
         TabIndex        =   66
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   8
         Left            =   3360
         TabIndex        =   67
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comprobar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   3360
         TabIndex        =   73
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   2160
         TabIndex        =   72
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Cambiar password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   6
         Left            =   960
         TabIndex        =   68
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame FrameCloro 
      Height          =   2415
      Left            =   2040
      TabIndex        =   53
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   3600
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegCloro 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   56
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   7
         Left            =   4440
         TabIndex        =   57
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   4
         Left            =   3360
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   61
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   60
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Crear registro control cloro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   1080
         TabIndex        =   59
         Top             =   240
         Width           =   3675
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.-   Impresion de facturas directas (tipo 4tonda)
    ' 1.-   Eliminar articulos masiva
    ' 2.-   Estadisticas consultas (archivo-facturacion-pedidos-consulta precio/cliente

    ' 5.-   Seleccionar marcas para las TOs

    ' 6.-   Lista de revision
    ' 7.-   Registro CLORO
    ' 8.-   Cambiar password
    ' 9.-   Ver articulo desde planning produccion


    '10.-   Impresion acciones realizadas

    '11.-   Acciones homologadas

    '12.-  Modificar kilos deposito

    '13.- Movimiento palots
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmVh As frmFacVehiculos
Attribute frmVh.VB_VarHelpID = -1

Private Cad As String
Private SePuedeCerrar As Boolean   'Puede llevar DoEvents
Private PrimeraVez As Boolean


Dim Codigo As String
Dim Cadselect As String
Dim Cadparam As String



Private Sub cmdAccionesRealizada_Click()
Dim NumParam As Integer
    CadenaDesdeOtroForm = ""
    Cad = ""
    For NumRegElim = 1 To Me.ListView4(0).ListItems.Count
        If Me.ListView4(0).ListItems(NumRegElim).Checked = True Then
            CadenaDesdeOtroForm = "O"
            Exit For
        End If
    Next
    If CadenaDesdeOtroForm = "" Then Cad = "-Seleccione alguna accion"
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To Me.ListView4(1).ListItems.Count
        If Me.ListView4(1).ListItems(NumRegElim).Checked = True Then
            CadenaDesdeOtroForm = "O"
            Exit For
        End If
    Next
    If CadenaDesdeOtroForm = "" Then Cad = Cad & vbCrLf & "-Seleccione algun trabajador"
    If Cad <> "" Then
        Cad = "Falta campos: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    
    'Vamos p'alla
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    CargarDatosAcciones
    Label3(18).Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If NumRegElim > 0 Then
        'Crgamos el report
        
        '-----------------------------------------------
        With frmImprimir
            VariablesReportLog NumParam
            .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
            .OtrosParametros = Cadparam
            .NumeroParametros = NumParam
            .ConSubInforme = False
            
            If Me.optVarios(0).Value Then
                .Opcion = 2022
            Else
                .Opcion = 2023
            End If
            

            .Show vbModal
        End With
        
    End If
    
    

End Sub

Private Sub cmdAccionListview_Click()
Dim T1 As Single

    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
    Next
    
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Seleccione algun artículo para eliminar", vbInformation
    Else
        CadenaDesdeOtroForm = Len(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = "Va a eliminar " & CadenaDesdeOtroForm & " artículo(s).   ¿Continuar?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then CadenaDesdeOtroForm = ""
    End If
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    
    
    
    
    'AHora eliminamos
    'Y el log de acciones
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    
    
    '-----------------------------------------------------------------------------
    
    Screen.MousePointer = vbHourglass
    lblElim(1).Caption = ""
    For NumRegElim = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(NumRegElim).Checked Then
            T1 = Timer
            ListView1.ListItems(NumRegElim).EnsureVisible
            conn.BeginTrans
            If EliminarArticulo(ListView1.ListItems(NumRegElim).Text, lblElim(1)) Then
                LOG.Insertar 7, vUsu, ListView1.ListItems(NumRegElim).Text & " " & ListView1.ListItems(NumRegElim).SubItems(1)
                conn.CommitTrans
                'QUitamos del nodo
                ListView1.ListItems.Remove ListView1.ListItems(NumRegElim).Index
                T1 = 1.5 - (Timer - T1)
                If T1 > 0 Then Espera T1
                
            Else
                'NO se ha podido eliminar
                conn.RollbackTrans
                ListView1.ListItems(NumRegElim).Bold = True
                ListView1.ListItems(NumRegElim).ForeColor = vbRed
                ListView1.ListItems(NumRegElim).Checked = False
            End If
        End If
    Next
    lblElim(1).Caption = ""
    Screen.MousePointer = vbDefault
    Set LOG = Nothing
    If ListView1.ListItems.Count = 0 Then
        SePuedeCerrar = True
        Unload Me
    End If
End Sub

Private Sub cmdCambiopass_Click()
    Cad = ""
    For NumRegElim = 0 To 3
        If Me.txtPassword(NumRegElim).Text = "" Then
            Cad = NumRegElim
            Exit For
        End If
    Next
    If Cad <> "" Then
        MsgBox "Campos obligados", vbExclamation
        PonerFoco Me.txtPassword(NumRegElim)
        Exit Sub
    End If
        
    If Me.txtPassword(2).Text <> Me.txtPassword(3).Text Then
        Cad = "-No coincide los campos password - verificar"
        NumRegElim = 2
    End If

    If Me.txtPassword(1).Text <> vUsu.PasswdPROPIO Then
        Cad = Cad & vbCrLf & vbCrLf & "-Password actual incorrecto"
        NumRegElim = 1
    End If
    
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        PonerFoco txtPassword(NumRegElim)
        Exit Sub
    End If
    
    If MsgBox("Desea cambiar el password?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Cad = "000" & CStr(vUsu.Codigo)
    Cad = CStr(Val(Right(Cad, 3)))
    Cad = "UPDATE usuarios.usuarios set passwordpropio = " & DBSet(Me.txtPassword(2).Text, "T") & " WHERE codusu = " & Cad
    If Not EjecutaSQL(conAri, Cad, False) Then
        MsgBox "Error cambiando password", vbExclamation
    Else
        MsgBox "Password cambiado", vbInformation
        vUsu.PasswdPROPIO = Me.txtPassword(1)
        Unload Me
    End If
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    
    If Opcion = 0 Then
        'Esta haciendo cosas. Preguntar si cerrar
        If Not SePuedeCerrar Then
            If MsgBox("Seguro que desea finalizar el proceso?", vbQuestion + vbYesNo) = vbYes Then SePuedeCerrar = True
            Exit Sub
        End If
    
    ElseIf Opcion = 5 Or Opcion = 6 Or Opcion = 11 Or Opcion = 12 Then
        CadenaDesdeOtroForm = ""
    End If
    
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdEliminarArticulos_Click()
Dim SQL As String
Dim It As ListItem

    '
    lblElim(0).Caption = "Cargando datos"
    lblElim(0).Refresh
    
    'Eliminamos los datos de tmpnseries
    conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo
    
    
    'Cargamos tmpnseries con los articulos del desde hasta
    SQL = ""
    If Me.txtArt(0).Text <> "" Then SQL = SQL & " codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " codartic <=" & DBSet(txtArt(1).Text, "T")
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    SQL = " SELECT " & vUsu.Codigo & ",codartic,0,0 FROM sartic " & SQL
    SQL = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinealb`,`numlinea`) " & SQL
    conn.Execute SQL
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Eliminamos de tmpnseries los articulos que seguro estan en
    ' alba, fact....
    EliminandoArticulos_Paso1
    
    
    'Ya tengo los articulos. Vere cuales puedo borrar
    lblElim(0).Caption = "Obteniendo registros"
    lblElim(0).Refresh
    
    SQL = "Select tmpnseries.codartic,nomartic from tmpnseries,sartic where codusu = " & vUsu.Codigo
    SQL = SQL & " AND tmpnseries.codartic=sartic.codartic"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        lblElim(0).Caption = ""
        MsgBox "No existen registros", vbExclamation
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    'Ajustamos los tamaños para cargar el LISTVIEW
    CargaColumnas
    NumRegElim = (Screen.Width - FrameListArticulos.Width - 420) \ 2
    Me.Left = NumRegElim
    NumRegElim = (Screen.Height - FrameListArticulos.Height - 360) \ 2
    Me.Top = NumRegElim
    Me.FrameDHArticulo.visible = False
    PonerFrameVisible Me.FrameListArticulos
    Me.lblTitulo(1).Caption = "Eliminar artículos"
    DoEvents
    
    'Vamos cargando los registros
    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add()
        It.Text = miRsAux!codartic
        It.SubItems(1) = miRsAux!NomArtic
        It.Checked = True
        'Sig
        miRsAux.MoveNext
    Wend
End Sub

Private Sub cmdGeneListaRevision_Click()
    If txtFecha(2).Text = "" Then
        MsgBox "Ponga la fecha", vbExclamation
        Exit Sub
    End If
    If txtTrab(0).Text = "" Or txtTrabDesc(0).Text = "" Then
        MsgBox "Ponga el trabajador", vbExclamation
        Exit Sub
    End If
    If GenerarDatosListaRevisiones Then
        'Generamos datos
        
        Cad = " {srevisiones.fecha} = Date(" & Year(txtFecha(2).Text) & "," & Month(txtFecha(2).Text) & "," & Day(txtFecha(2).Text) & ")"
        With frmImprimir
            .FormulaSeleccion = Cad
            .OtrosParametros = ""
            .NumeroParametros = 0
    
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 2002
            .Titulo = "Listado revision"
            .NombreRPT = "morListaRevision.rpt"
            .ConSubInforme = False
            .Show vbModal
        End With
        CadenaDesdeOtroForm = DBSet(txtFecha(2).Text, "F")
        Unload Me
    End If
End Sub

Private Sub cmdHomologacion_Click()
    If Trim(Me.txtFecha(7).Text) = "" Or Trim(Me.txtHomologa.Text) = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = txtFecha(7).Text & "|" & Me.txtHomologa.Text & "|"
    Unload Me
End Sub

Private Sub cmdImpresionMovPalet_Click()
    Codigo = "vallPalot.rpt"
    Cadparam = ""
    NumRegElim = 0
    If Me.optMovPalot(1).Value Then
        Codigo = "vallAlbaDocControlPalot.rpt"
        'Updateo siempre
        Cad = "REPLACE INTO spalots(codigo,anyo,fecha,Destino,TransConductor,TransCondDNI,TransEmpresa,TransMatricula) VALUES ("
        Cad = Cad & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & Year(CDate(RecuperaValor(CadenaDesdeOtroForm, 2))) & ","
        Cad = Cad & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F")
        For NumRegElim = 0 To 4
            Cad = Cad & ", " & DBSet(Me.txtPalot(NumRegElim).Text, "T")
        Next
        Cad = Cad & ")"
        EjecutaSQL conAri, Cad
        NumRegElim = 1
        
    End If
    Cad = "{tmprutas.codusu} = " & vUsu.Codigo
    LlamaImprimirGral Cad, Cadparam, CInt(NumRegElim), Codigo, "PALOTS"
    Unload Me
End Sub

Private Sub cmdKilosDeposito_Click()
    CadenaDesdeOtroForm = ""
    If txtDecimal(0).Text <> "" Then
        Cad = "Seguro que desea ajustar la cantidad del depósito a:        " & txtDecimal(0).Text & " Kg. ?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
            CadenaDesdeOtroForm = txtDecimal(0).Text
            Unload Me
        End If
    End If
End Sub

Private Sub cmdListConsultaPedido_Click()
Dim Aux As String


    Cad = ""
    Aux = CadenaDesdeHastaBD(txtArt(2).Text, txtArt(3).Text, "codartic", "T")
    If Aux <> "" Then Cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHastaBD(txtFecha(0).Text, txtFecha(1).Text, "DiaHora", "FH")
    If Aux <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & Aux
    End If
        
    If Not HayRegParaInforme("sconsulta", Cad) Then Exit Sub
    
    
    'Para el informe
    Cad = ""
    Aux = CadenaDesdeHasta(txtArt(2).Text, txtArt(3).Text, "{sconsulta.codartic}", "T")
    If Aux <> "" Then Cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHasta(txtFecha(0).Text, txtFecha(1).Text, "{sconsulta.DiaHora}", "FH")
    If Aux <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & Aux
    End If
    
    
    
    
    With frmImprimir
        .FormulaSeleccion = Cad
        .OtrosParametros = ""
        .NumeroParametros = 0

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2002
        .Titulo = "Estadistica consultas pedidos"
        .NombreRPT = "rFacConsuPrecioArt.rpt"
        .ConSubInforme = False
        .Show vbModal
    End With
    
    
    
End Sub

Private Sub cmdMarcas_Click()
        Cad = ""
        For NumRegElim = 1 To Me.ListView2.ListItems.Count
            If Me.ListView2.ListItems(NumRegElim).Checked Then Cad = Cad & ", " & ListView2.ListItems(NumRegElim).Text
        Next
        If Cad = "" Then
            MsgBox "Seleccione alguna marca", vbExclamation
            Exit Sub
        End If
        CadenaDesdeOtroForm = "(" & Mid(Cad, 2) & ")"
        Unload Me
End Sub


Private Sub cmdRegCloro_Click()
Dim F As Date
    If Me.txtFecha(3).Text = "" Or txtFecha(4).Text = "" Then
        MsgBox "Ponga fecha desde hasta.", vbExclamation
        Exit Sub
    End If
    
    
        
    If CDate(Me.txtFecha(3).Text) > CDate(txtFecha(4).Text) Then
        MsgBox "Fecha incio mayor que fecha fin.", vbExclamation
        Exit Sub
    End If
    
    If CadenaDesdeOtroForm <> "" Then
        'No deberia ser """
        If CDate(Me.txtFecha(3).Text) < CDate(CadenaDesdeOtroForm) Then
            Cad = "Fecha incio menor que ultima fecha impresa." & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    NumRegElim = DateDiff("d", CDate(Me.txtFecha(3).Text), CDate(Me.txtFecha(4).Text))
    If NumRegElim > 18 Then
        Cad = "Deberias imprimir como mucho de quinciena en quincena" & vbCrLf & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        
    Cad = "DELETE from tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Cad
    
    F = CDate(txtFecha(3).Text)
    NumRegElim = 1
    Cad = "INSERT INTO tmpinformes(codusu ,codigo1 ,fecha1) VALUES "
    While F <= CDate(txtFecha(4).Text)
        Cad = Cad & "(" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(F, FormatoFecha) & "') ,"
        NumRegElim = NumRegElim + 1
        F = DateAdd("d", 1, F)
    Wend
    NumRegElim = Len(Cad)
    Cad = Mid(Cad, 1, NumRegElim - 2)  'quito la ulimta coma
    conn.Execute Cad
    
    Cad = "{tmpinformes.codusu} = " & vUsu.Codigo
    LlamaImprimirGral Cad, "", 0, "morRegCloro.rpt", "Registro cloro: " & txtFecha(3).Text & " - " & txtFecha(4).Text
    
    'Insertamos en registro cloro
    '----------------------------
    NumRegElim = vUsu.Codigo Mod 1000
    Cad = ComputerName
    Cad = "," & DBSet(Now, "F") & "," & NumRegElim & "," & DBSet(Cad, "T") & ")"
    Cad = DBSet(txtFecha(3).Text, "F") & "," & DBSet(txtFecha(4).Text, "F") & Cad
    Cad = "insert into `sregcloro` (`Fech1`,`Fech2`,`FechaCreacion`,`codusu`,`pc`) values (" & Cad
    conn.Execute Cad
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Select Case Opcion
        Case 0
            'Se pone a imprimir las facturas
            HacerImpresionFacturas
        Case 5
            CargarMarcas
        
        Case 8
            
            PonerFoco Me.txtPassword(1)
        Case 9
            CargarDatosArticuloPlanning
            
        Case 10
            CargaListAccionesRealizadas
        Case 11
            CargaHomologacion
        Case 12
            Label5.Caption = "Ajustará los kilos del deposito y los del lote asociado a los kilos introducidos" & vbCrLf & vbCrLf
            Label5.Caption = Label5.Caption & "No generará movimiento de regularizacion."
        Case 13
            PonerDatosTransportePalot
            
        End Select
    End If
End Sub

Private Sub CargarIconos()
Dim I As Image


    For Each I In Me.imgArticulo
         I.Picture = frmppal.imgListComun.ListImages(19).Picture
         I.ToolTipText = "Articulo"
    Next
    For Each I In Me.imgFecha
         I.Picture = frmppal.imgListComun.ListImages(23).Picture
         I.ToolTipText = "fecha"
    Next
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    PrimeraVez = True
    
    limpiar Me
    CargarIconos
    FrameListArticulos.visible = False
    FrameDHArticulo.visible = False
    Me.FrameImpresionFacturasDirectas.visible = False
    Me.FrameEstadisticasConsultas.visible = False
    FrameMarcas.visible = False
    FrameListRevision.visible = False
    FrameCloro.visible = False
    FramePasswd.visible = False
    FramePlanning.visible = False
    Me.FrameAccionesRealizadas.visible = False
    FrameHomologa.visible = False
    FrameModificaKilosDeposito.visible = False
    FramePaletMovimImprimir.visible = False
    
    Select Case Opcion
    Case 0
        PonerFrameVisible Me.FrameImpresionFacturasDirectas
    Case 1
        PonerFrameVisible FrameDHArticulo
    Case 2
        PonerFrameVisible Me.FrameEstadisticasConsultas
    Case 5
        PonerFrameVisible Me.FrameMarcas
    Case 6
        PonerFrameVisible FrameListRevision
        txtTrab(0).Text = PonerTrabajadorConectado(Cad)
        Me.txtTrabDesc(0).Text = Cad
        Cad = ""
    Case 7
        PonerFrameVisible FrameCloro
        PonerDatosCloro
        
    Case 8
        PonerFrameVisible FramePasswd
        
        Me.txtPassword(0).Text = vUsu.Nombre
    Case 9
        PonerFrameVisible FramePlanning
    Case 10
        PonerFrameVisible FrameAccionesRealizadas
        Label3(18).Caption = "" 'el indicador
        
    Case 11
        PonerFrameVisible FrameHomologa
        Me.lblTitulo(10).Caption = "Acción homologación"
        If CadenaDesdeOtroForm = "" Then
            Me.lblTitulo(10).Caption = Me.lblTitulo(10).Caption & " (Nueva)"
        Else
            Me.lblTitulo(10).Caption = Me.lblTitulo(10).Caption & " (Modificar)"
        End If
    
    Case 12
        PonerFrameVisible FrameModificaKilosDeposito
        
    Case 13
        PonerFrameVisible FramePaletMovimImprimir
        
        txtPalot(0).Text = DevuelveDesdeBD(conAri, "pobclien", "tmprutas", "codusu", CStr(vUsu.Codigo))
        Me.ImgTransporte.Picture = frmppal.imgListComun.ListImages(19).Picture
    End Select
    cmdCancelar(Opcion).Cancel = True
    SePuedeCerrar = True
End Sub



Private Sub PonerFrameVisible(Fr As Frame)
    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    Me.Height = Fr.Height + 480
    Me.Width = Fr.Width + 320
End Sub




Private Sub HacerImpresionFacturas()
Dim I As Integer
Dim Fin As Boolean
    SePuedeCerrar = False
    
    Me.lblImpr(0).Caption = "Leyendo datos"
    lblImpr(0).Refresh
    Espera 0.25
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select count(*) from scafac WHERE " & CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim = 0 Then Exit Sub
    
    CadenaDesdeOtroForm = "Select codtipom, numfactu, fecfactu, nomclien from scafac where " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " ORDER BY fecfactu,numfactu"
    
    miRsAux.Open CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    Fin = False
    While Not Fin
        I = I + 1
        Me.lblImpr(1).Caption = "Fac. " & Format(miRsAux!NumFactu, "00000") & " de " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & "     " & Mid(miRsAux!nomclien, 1, 20)
        lblImpr(1).Refresh
        Me.lblImpr(0).Caption = "Registro: " & I & "   de   " & NumRegElim
        lblImpr(0).Refresh
    
        'IMprimimos la factura
        ReImprimirDirectoFact " scafac.codtipom ='" & miRsAux!Codtipom & "' AND scafac.numfactu = " & miRsAux!NumFactu
    
        DoEvents
        If SePuedeCerrar Then
            Fin = True  'Han pulsado cancelar
        Else
            'Siguiente
            miRsAux.MoveNext
            Fin = miRsAux.EOF
        End If
        If I Mod 50 = 25 Then Me.Refresh
            
        
    Wend
    If miRsAux.EOF Then
        'Significa que ha acabado toda la impresion. Con lo cual
        'pongo CadenaDesdeOtroForm="" para que el form de reimpresion lo cierre
        CadenaDesdeOtroForm = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    SePuedeCerrar = True
    Unload Me  'Y cierro
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not SePuedeCerrar Then Cancel = 1
    
    
End Sub


Private Sub imgSel_Click(Index As Integer)

End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmVh_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub imgArticulo_Click(Index As Integer)
        Cad = ""
        Set frmA = New frmAlmArticulos
        frmA.DeConsulta = True
        frmA.DatosADevolverBusqueda2 = "@1@"
        frmA.Show vbModal
        Set frmA = Nothing
        If Cad <> "" Then
            Me.txtArt(Index).Text = RecuperaValor(Cad, 1)
            Me.txtArtD(Index).Text = RecuperaValor(Cad, 2)
        End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index <= 1 Then
        For NumRegElim = 1 To ListView1.ListItems.Count
            ListView1.ListItems(NumRegElim).Checked = Index = 1
        Next NumRegElim
    ElseIf Index < 4 Then
        For NumRegElim = 1 To ListView2.ListItems.Count
            ListView2.ListItems(NumRegElim).Checked = Index = 3
        Next NumRegElim
        
    Else
        Cad = "0"
        If Index > 5 Then Cad = "1"
            
        For NumRegElim = 1 To ListView4(Val(Cad)).ListItems.Count
            ListView4(Val(Cad)).ListItems(NumRegElim).Checked = (Index Mod 2) = 1
        Next NumRegElim
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then txtFecha(Index).Text = Cad
End Sub

Private Sub ImgTransporte_Click()
    Cad = ""
    Set frmVh = New frmFacVehiculos
    frmVh.DatosADevolverBusqueda = "0"
    frmVh.Show vbModal
    Set frmVh = Nothing
    If Cad <> "" Then
        Set miRsAux = New ADODB.Recordset
        Cad = "select * from svehiculos where codigo=" & RecuperaValor(Cad, 1)
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            'matricula  Empresa   Conductor   DNIConductor
            If Not IsNull(miRsAux!Empresa) Then txtPalot(3).Text = miRsAux!Empresa
            If Not IsNull(miRsAux!matricula) Then txtPalot(4).Text = miRsAux!matricula
            If Not IsNull(miRsAux!conductor) Then txtPalot(1).Text = miRsAux!conductor
            If Not IsNull(miRsAux!DNIConductor) Then txtPalot(2).Text = miRsAux!DNIConductor
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        Cad = ""
    End If
End Sub

Private Sub optMovPalot_Click(Index As Integer)
    For NumRegElim = 0 To 4
        Label7(NumRegElim).visible = optMovPalot(1).Value
        txtPalot(NumRegElim).visible = optMovPalot(1).Value
    Next
    Me.ImgTransporte.visible = optMovPalot(1).Value
End Sub

Private Sub txtArt_GotFocus(Index As Integer)
 PonerFoco txtArt(Index)
End Sub

Private Sub txtArt_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtArt_LostFocus(Index As Integer)
Dim C As String

    txtArt(Index).Text = Trim(txtArt(Index).Text)
    If txtArt(Index).Text = "" Then
        C = ""
    Else
        C = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtArt(Index).Text, "T")
        If C = "" Then
            'El articulo no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
            
        End If
    End If
    txtArtD(Index).Text = C
End Sub



Private Sub EliminandoArticulos_Paso1()
Dim C As String
Dim SQL As String
Dim Aux As String
Dim nt As Integer
Dim J As Byte

    If Me.txtArt(0).Text <> "" Then SQL = SQL & " codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " codartic <=" & DBSet(txtArt(1).Text, "T")
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    
     
    'El stock
    lblElim(0).Caption = "Almacenes"
    lblElim(0).Refresh
    C = "select codartic,sum(canstock) from salmac " & SQL & " group by codartic having sum(canstock) <> 0"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
         conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    For J = 0 To 2
        DevuelveTablasBorre J, C, Aux, nt
        For NumRegElim = 1 To nt
            
            lblElim(0).Caption = RecuperaValor(Aux, CInt(NumRegElim)) & "   -"
            If J = 0 Then
                lblElim(0).Caption = lblElim(0).Caption & "Clientes"
            ElseIf J = 1 Then
                lblElim(0).Caption = lblElim(0).Caption & "Prove"
            Else
                lblElim(0).Caption = lblElim(0).Caption & "Varios"
            End If
            lblElim(0).Refresh
            
            
            miRsAux.Open "Select codartic from " & RecuperaValor(C, CInt(NumRegElim)) & SQL & " GROUP BY codartic", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                 conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Me.Refresh
        Next NumRegElim
    Next J
    
End Sub


Private Sub CargaColumnas()
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case Opcion

    Case 1
        Me.ListView1.Checkboxes = True
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Código"
        clmX.Width = 2200
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descripción"
        clmX.Width = 3600
        
    End Select
    Me.FrameListArticulos.ZOrder 1  'QUe lo traiga al frente
End Sub


 

Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDecimal_LostFocus(Index As Integer)
    txtDecimal(Index).Text = Trim(txtDecimal(Index).Text)
    If txtDecimal(Index).Text <> "" Then
        If Not PonerFormatoDecimal(txtDecimal(Index), 3) Then
            txtDecimal(Index).Text = ""
        Else
            If ImporteFormateado(txtDecimal(Index).Text) < 0 Then
                MsgBox "Importe no puede ser negativo", vbExclamation
                txtDecimal(Index).Text = ""
            End If
        End If
        If txtDecimal(Index).Text = "" Then PonerFoco txtDecimal(Index)
    End If
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        Cad = txtFecha(Index).Text
        If Not EsFechaOK(Cad) Then
            MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        Else
            txtFecha(Index).Text = Cad
        End If
    End If
End Sub




Private Sub CargarMarcas()
Dim It As ListItem
    ListView2.ListItems.Clear
    Cad = "Select * from smarca order by nommarca"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add()
        It.Text = miRsAux!codmarca
        It.SubItems(1) = DBLet(miRsAux!nommarca, "T")
        Cad = "|" & miRsAux!codmarca & "|"
        If InStr(1, CadenaDesdeOtroForm, Cad) > 0 Then It.Checked = True
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
End Sub



Private Function GenerarDatosListaRevisiones() As Boolean
    
    GenerarDatosListaRevisiones = True
    Cad = "insert into `srevisiones` (`Fecha`,`RealizadoPor`,`Comentarios`,`puntuacion`) values ("
    Cad = Cad & DBSet(txtFecha(2).Text, "F") & "," & DBSet(Me.txtTrabDesc(0).Text, "T") & ",NULL,NULL)"
    If EjecutaSQL(conAri, Cad, True) Then
        Cad = "SELECT " & DBSet(txtFecha(2).Text, "F") & ",codigo,descripcion,orden,denominacion,NULL from"
        Cad = Cad & " srevarea,srevaspectos where codarea = codigo"
        Cad = "INSERT INTO   srevisionesl (fecha,codigo,descripcion,orden,denominacion,ok) " & Cad
        If Not EjecutaSQL(conAri, Cad, True) Then
            Cad = "DELETE FROM srevisiones WHERE fecha = " & DBSet(txtFecha(2).Text, "F")
            EjecutaSQL conAri, Cad, False
            GenerarDatosListaRevisiones = False
        End If
    Else
        GenerarDatosListaRevisiones = False
    End If
End Function



Private Sub txtPassword_GotFocus(Index As Integer)
    ConseguirFoco txtPassword(Index), 3
End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
     ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)
    Dim C As String
    
    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    If txtTrab(Index).Text = "" Then
        C = ""
    Else
        C = DevuelveDesdeBDNew(conAri, "straba", "nomtraba", "codtraba", txtTrab(Index).Text, "N")
        If C = "" Then
            'El articulo no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
            MsgBox "No existe trabajador", vbExclamation
            txtTrab(Index).Text = ""
            PonerFoco txtTrab(Index)
        End If
    End If
    Me.txtTrabDesc(Index).Text = C
End Sub


Private Sub PonerDatosCloro()
Dim F As Date
On Error GoTo EPonerDatosCloro
    
    Cad = DevuelveDesdeBD(conAri, "max(fech2)", "sregcloro", "1", 1)
    CadenaDesdeOtroForm = "01/01/1900"  'memorizo la fecha
    If Cad = "" Then
        Cad = DateAdd("d", -1, Now)
    Else
        CadenaDesdeOtroForm = Cad
    End If
    
    F = CDate(Cad)
    F = DateAdd("d", 1, F)
    txtFecha(3).Text = Format(F, "dd/mm/yyyy")
    F = DateAdd("d", 14, F)
    If Day(F) >= 28 Then
        NumRegElim = DiasMes(CByte(Month(F)), Year(F))
        F = CDate(NumRegElim & "/" & Month(F) & "/" & Year(F))
    End If
    txtFecha(4).Text = Format(F, "dd/mm/yyyy")
EPonerDatosCloro:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cloro"
    Set miRsAux = Nothing
End Sub




Private Sub CargarDatosArticuloPlanning()
Dim It
    On Error GoTo ecargarDatosArticuloPlanning
    Set miRsAux = New ADODB.Recordset
    
    Me.txtArtD(4).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 1)
    Me.txtArt(4).Text = CadenaDesdeOtroForm
    'El LW esta ordenado por una columna oculta. Hay pondremos las fechas en formato yyyymmdd
    'Lo primero en poner el stock y luego pedcli y pedpro
    
    Cad = "Select * from salmac where codartic='" & CadenaDesdeOtroForm & "' AND codalmac = 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    Set It = ListView3.ListItems.Add()
    It.Text = " "
    It.SubItems(1) = " "
    It.SubItems(2) = " 1"
    It.SubItems(3) = " S T O C K   *** "
    It.SubItems(4) = Format(miRsAux!CanStock, FormatoCantidad)
    It.SubItems(5) = "00000000"  'el primero

    miRsAux.Close
    
    
   
    '
    Cad = "select scaped.fecentre,scaped.numpedcl,sliped.cantidad,nomclien,sliped.nomartic from sarti1,sartic,scaped,sliped where scaped.numpedcl=sliped.numpedcl AND"
    Cad = Cad & " sliped.codartic=sartic.codartic and sarti1.codArtic = sartic.codArtic"
    Cad = Cad & " and codarti1='" & CadenaDesdeOtroForm & "'"
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    
    While Not miRsAux.EOF
        Set It = ListView3.ListItems.Add()
        It.Text = "Si"
        It.SubItems(1) = Format(miRsAux!FecEntre, "dd/mm/yyyy")
        
        It.SubItems(2) = miRsAux!numpedcl
        It.SubItems(3) = miRsAux!NomArtic
        It.SubItems(4) = Format(miRsAux!Cantidad, FormatoCantidad)
        It.SubItems(5) = Format(miRsAux!FecEntre, "yyyymmdd")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
   
    
    
    Cad = "select scappr.numpedpr,fecpedpr,nomprove,cantidad from scappr,slippr where scappr.numpedpr =slippr.numpedpr "
    Cad = Cad & " and codartic='" & CadenaDesdeOtroForm & "'"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListView3.ListItems.Add()
        It.Text = " "
        It.SubItems(1) = Format(miRsAux!fecpedpr, "dd/mm/yyyy")
        
        It.SubItems(2) = miRsAux!numpedpr
        It.SubItems(3) = miRsAux!nomprove
        It.SubItems(4) = Format(miRsAux!Cantidad, FormatoCantidad)
        It.SubItems(5) = Format(miRsAux!fecpedpr, "yyyymmdd")
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close



ecargarDatosArticuloPlanning:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
End Sub





Private Sub CargaListAccionesRealizadas()
Dim L As Collection
Dim I As Integer
Dim It As ListItem
    Set L = New Collection
    Set LOG = New cLOG
    If LOG.DevuelveAcciones(L) Then
        For NumRegElim = 1 To L.Count
            Cad = L.Item(NumRegElim)
            I = RecuperaValor(Cad, 1)
            'Acciones que no van a entrar en el log
'            If i = 10 Or i = 6 Or i = 5 Or i = 3 Or i = 2 Then
'                'acciones que no mostrare
'                '
'                 Stop
'            Else
                Set It = Me.ListView4(0).ListItems.Add(, "K" & I)
                It.Text = RecuperaValor(Cad, 2)
'            End If
        Next
    End If
    Set LOG = Nothing
        
        
    'Los usuarios
    Cad = "select distinct(usuario) from slog order by 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux.Fields(0) <> "root" Then
            Set It = Me.ListView4(1).ListItems.Add()
            It.Text = miRsAux.Fields(0)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub






Private Sub CargarDatosAcciones()
Dim I As Integer
Dim miSQL As String
Dim OtrosDatos As String


    On Error GoTo ECargarDatosAcciones
    NumRegElim = 0
    
    Label3(18).Caption = "Inicio proceso"
    Label3(18).Refresh
    
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    
    
    miSQL = ""
    For NumRegElim = 1 To Me.ListView4(0).ListItems.Count
        If Me.ListView4(0).ListItems(NumRegElim).Checked = True Then miSQL = miSQL & ", " & Mid(ListView4(0).ListItems(NumRegElim).Key, 2)
    Next
    Codigo = " accion IN (" & Mid(miSQL, 2) & ")"
    
    
    miSQL = ""
    For NumRegElim = 1 To Me.ListView4(1).ListItems.Count
        If Me.ListView4(1).ListItems(NumRegElim).Checked = True Then miSQL = miSQL & ", " & DBSet((ListView4(1).ListItems(NumRegElim).Text), "T")
    Next
    Codigo = Codigo & " AND usuario IN (" & Mid(miSQL, 2) & ")"
    miSQL = Codigo
    If Me.txtFecha(5).Text <> "" Then miSQL = miSQL & " AND slog.fecha >= '" & Format(Me.txtFecha(5).Text, FormatoFecha) & " 00:00:00'"
    If Me.txtFecha(6).Text <> "" Then miSQL = miSQL & " AND slog.fecha <= '" & Format(Me.txtFecha(6).Text, FormatoFecha) & " 23:59:59'"
    

    miSQL = "Select * from slog WHERE " & miSQL
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    miSQL = ""
    Codigo = "INSERT INTO tmpinformes( codusu,codigo1,campo1,nombre2,nombre3,fecha1, obser) VALUES "
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Label3(18).Caption = NumRegElim
        Label3(18).Refresh
        'tmpinformes codusu,codigo1,campo1,nombre2, obser
        miSQL = miSQL & ", (" & vUsu.Codigo & "," & NumRegElim & "," & miRsAux!Accion & ","
        miSQL = miSQL & DBSet(miRsAux!Usuario, "T") & ",'" & Format(miRsAux!Fecha, "hh:mm:ss")
        miSQL = miSQL & "'," & DBSet(miRsAux!Fecha, "F") & ","
        If Val(miRsAux!Accion) = 14 Then
                'CAMBIO DE PRECIO. Voy a añadir el codartic codartic
            I = InStr(1, miRsAux!Descripcion, "culo ")
            If I > 0 Then
                Cadselect = Mid(miRsAux!Descripcion, 1, I + 4)
                Cadparam = Mid(miRsAux!Descripcion, I + 5)
                I = InStr(1, Cadparam, vbCrLf)
                If I > 0 Then
                    OtrosDatos = Mid(Cadparam, 1, I - 1)
                    Cadselect = Cadselect & OtrosDatos 'voy concatenando para luego devolver los textos
                    OtrosDatos = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", OtrosDatos, "T")
                    If OtrosDatos <> "" Then Cadselect = Cadselect & vbCrLf & " [" & OtrosDatos & "]"
                    Cadselect = Cadselect & Mid(Cadparam, I)
                Else
                   ' Stop
                End If
            Else
                'Stop
            End If
        Else
            '
            I = 0 'para que ponga ahi bajo la descripcion
        End If
        If I = 0 Then Cadselect = miRsAux!Descripcion
        miSQL = miSQL & DBSet(Cadselect, "T") & ")"
        
        If (NumRegElim Mod 100) = 0 Then
            miSQL = Mid(miSQL, 2)
            miSQL = Codigo & miSQL
            conn.Execute miSQL
            miSQL = ""
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        miSQL = Mid(miSQL, 2)
        miSQL = Codigo & miSQL
        conn.Execute miSQL
        miSQL = ""
    End If
    OtrosDatos = ""
    
    
    'Ya tengo todos metidos, updateare el nombre del login
    Label3(18).Caption = "Usuarios"
    Label3(18).Refresh
        
    miSQL = "Select nombre2 from tmpinformes where codusu = " & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Label3(18).Caption = miRsAux.Fields(0)
        Label3(18).Refresh
        
        miSQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "login", miRsAux.Fields(0), "T")
        If miSQL <> "" Then
            miSQL = "UPDATE tmpinformes SET nombre2 = " & DBSet(miSQL, "T")
            miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND nombre2 = " & DBSet(miRsAux.Fields(0), "T")
            conn.Execute miSQL
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    For I = 1 To Me.ListView4(0).ListItems.Count
        If ListView4(0).ListItems(I).Checked Then
            Label3(18).Caption = ListView4(0).ListItems(I).Text
            Label3(18).Refresh
            
            
            
            
                miSQL = "UPDATE tmpinformes SET nombre1 = " & DBSet(ListView4(0).ListItems(I).Text, "T")
                miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND campo1 = " & Mid(ListView4(0).ListItems(I).Key, 2)
                conn.Execute miSQL
            
            
        End If
    Next
    
    
    
    If NumRegElim = 0 Then MsgBox "No existen datos con esos valores", vbExclamation
    
    
    Exit Sub
ECargarDatosAcciones:
    MuestraError Err.Number, Err.Description
    NumRegElim = 0
End Sub


Private Sub VariablesReportLog(ByRef NumParam As Integer)
Dim Cuantos As Integer
    
        
    
    'Desde hasta 1
    '------------------
    Cad = ""
    Cuantos = 0
    Cadparam = ""
    
    For NumRegElim = 1 To Me.ListView4(1).ListItems.Count
        If Me.ListView4(1).ListItems(NumRegElim).Checked = True Then Cuantos = Cuantos + 1
    Next
    For NumRegElim = 1 To Me.ListView4(1).ListItems.Count
        If Me.ListView4(1).ListItems(NumRegElim).Checked = True Then
            Codigo = ListView4(1).ListItems(NumRegElim).Text
            If Cuantos > 4 Then
                'SON MUCHOS. Ponemos el login
                
            Else
                 Codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "login", Codigo, "T")
                 If Codigo = "" Then Codigo = ListView4(1).ListItems(NumRegElim).Text
            End If
            Cad = Cad & ", " & Codigo
        End If
    Next
    Cad = "Trab: " & Mid(Cad, 2)
    
    Codigo = ""
    If Me.txtFecha(5).Text <> "" Then Codigo = "desde " & Me.txtFecha(5).Text
    If Me.txtFecha(6).Text <> "" Then Codigo = Codigo & " hasta " & txtFecha(6).Text
    If Codigo <> "" Then Codigo = "Fecha: " & Codigo
    Cad = Trim(Codigo & "   " & Cad)
    
    Cad = "pdh1=""" & Cad & """|"
    Cadparam = Cadparam & Cad
    NumParam = NumParam + 1
        
    'Desde hasta 1
    '------------------
    Cad = ""
    Cuantos = 0
    For NumRegElim = 1 To Me.ListView4(0).ListItems.Count
        If Me.ListView4(0).ListItems(NumRegElim).Checked = True Then Cuantos = Cuantos + 1
    Next
    For NumRegElim = 1 To Me.ListView4(0).ListItems.Count
        If Me.ListView4(0).ListItems(NumRegElim).Checked = True Then
            
            If Cuantos > 6 Then
                'SON MUCHOS. Ponemos los codigos
                Codigo = Mid(ListView4(0).ListItems(NumRegElim).Key, 2)
            Else
                Codigo = ListView4(0).ListItems(NumRegElim).Text
            End If
            Cad = Cad & ", " & Codigo
        End If
    Next
    Cad = "Acc: " & Mid(Cad, 2)
    Cad = "pdh2=""" & Cad & """|"
    Cadparam = Cadparam & Cad
    NumParam = NumParam + 1
        
    
    
    Cad = "pEmpresa=""" & vEmpresa.nomempre & """|"
    Cadparam = Cadparam & Cad
    NumParam = NumParam + 1
    
End Sub




Private Sub CargaHomologacion()

    If CadenaDesdeOtroForm = "" Then
        Me.txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
        Me.txtHomologa.Text = ""
    Else
        Codigo = "N"
        Me.txtFecha(7).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.txtHomologa.Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    End If
    CadenaDesdeOtroForm = ""
End Sub



Private Sub PonerDatosTransportePalot()
    'SQL = "spalots(codigo,anyo,fecha,TransEmpresa,TransMatricula,TransConductor,TransCondDNI,Destino)"
    Set miRsAux = New ADODB.Recordset
    Cad = "Select codigo,anyo,fecha,TransEmpresa,TransMatricula,TransConductor,TransCondDNI,Destino from spalots"
    Cad = Cad & " WHERE codigo =" & RecuperaValor(CadenaDesdeOtroForm, 1)
    Cad = Cad & " AND anyo =" & Year(CDate(RecuperaValor(CadenaDesdeOtroForm, 2)))
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            If Not IsNull(miRsAux!TransEmpresa) Then txtPalot(3).Text = miRsAux!TransEmpresa
            If Not IsNull(miRsAux!TransMatricula) Then txtPalot(4).Text = miRsAux!TransMatricula
            If Not IsNull(miRsAux!TransConductor) Then txtPalot(1).Text = miRsAux!TransConductor
            If Not IsNull(miRsAux!TransCondDNI) Then txtPalot(2).Text = miRsAux!TransCondDNI
            If Not IsNull(miRsAux!Destino) Then txtPalot(0).Text = miRsAux!Destino
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub
