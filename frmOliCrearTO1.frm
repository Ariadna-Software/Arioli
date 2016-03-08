VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOliCrearTO1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso generación TOs"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAVAB 
      Height          =   6855
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   9375
      Begin VB.TextBox txtObserva 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdGenAVAB 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7920
         TabIndex        =   62
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenAVAB 
         Caption         =   "Gen. AVAB"
         Height          =   375
         Index           =   0
         Left            =   6600
         TabIndex        =   61
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   4080
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   375
         Index           =   2
         Left            =   6720
         Picture         =   "frmOliCrearTO1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Añadir cliente"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   375
         Index           =   3
         Left            =   7200
         Picture         =   "frmOliCrearTO1.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Quitar cliente seleccionado"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   12
         Left            =   1320
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   13
         Left            =   2040
         TabIndex        =   54
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   15
         Left            =   4800
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   17
         Left            =   8400
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   14
         Left            =   2760
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   16
         Left            =   5880
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cboTarifaAVAB 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   5880
         Width           =   3615
      End
      Begin MSComctlLib.ListView lw2 
         Height          =   2775
         Left            =   1320
         TabIndex        =   66
         Top             =   2400
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Obs."
         Height          =   195
         Index           =   4
         Left            =   5880
         TabIndex        =   79
         Top             =   720
         Width           =   435
      End
      Begin VB.Image imgObserva 
         Height          =   240
         Index           =   1
         Left            =   6360
         Picture         =   "frmOliCrearTO1.frx":0B14
         ToolTipText     =   "Buscar cliente"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Exportar Tarifa-Oferta al   AVAB "
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
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
         Index           =   15
         Left            =   360
         TabIndex        =   75
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Index           =   14
         Left            =   360
         TabIndex        =   74
         Top             =   720
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Index           =   3
         X1              =   1320
         X2              =   9000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmOliCrearTO1.frx":0C16
         ToolTipText     =   "Buscar cliente"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   3840
         Picture         =   "frmOliCrearTO1.frx":0D18
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   73
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   72
         Top             =   720
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1800
         Picture         =   "frmOliCrearTO1.frx":0DA3
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes %"
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
         Index           =   13
         Left            =   360
         TabIndex        =   71
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes "
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
         Index           =   12
         Left            =   4080
         TabIndex        =   70
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Margen"
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
         Index           =   11
         Left            =   7560
         TabIndex        =   69
         Top             =   1320
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Index           =   2
         X1              =   1320
         X2              =   9000
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Index           =   10
         Left            =   360
         TabIndex        =   68
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccione una opcion si va a generar las tarifas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   67
         Top             =   6000
         Width           =   3825
      End
   End
   Begin VB.Frame FrameCopiarTarifa 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox cboTarif2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtTar 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   1080
         Width           =   6015
      End
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   42
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "copiar"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtTar 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   11
         Left            =   7320
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   10
         Left            =   5400
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   9
         Left            =   4320
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   8
         Left            =   2520
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1200
         Picture         =   "frmOliCrearTO1.frx":0E2E
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa destino"
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
         Left            =   0
         TabIndex        =   49
         Top             =   2760
         Width           =   1170
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa origen"
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
         Left            =   0
         TabIndex        =   48
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Margen"
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
         Left            =   6480
         TabIndex        =   46
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes "
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
         Left            =   3600
         TabIndex        =   45
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes %"
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
         Left            =   0
         TabIndex        =   44
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Copiar tarifa"
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameFinal 
      Height          =   6495
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   9375
      Begin VB.TextBox txtObserva 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6600
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cboTarifa 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   5880
         Width           =   3615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   3
         Left            =   8400
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   375
         Index           =   1
         Left            =   7200
         Picture         =   "frmOliCrearTO1.frx":1830
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Quitar cliente seleccionado"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   375
         Index           =   0
         Left            =   6720
         Picture         =   "frmOliCrearTO1.frx":1DBA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Añadir cliente"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin MSComctlLib.ListView lwCli 
         Height          =   2775
         Left            =   1320
         TabIndex        =   21
         Top             =   2400
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Obs."
         Height          =   195
         Index           =   5
         Left            =   5880
         TabIndex        =   80
         Top             =   720
         Width           =   435
      End
      Begin VB.Image imgObserva 
         Height          =   240
         Index           =   0
         Left            =   6360
         Picture         =   "frmOliCrearTO1.frx":2344
         ToolTipText     =   "Buscar cliente"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccione una opcion si va a generar las tarifas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   31
         Top             =   6000
         Width           =   3825
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Left            =   360
         TabIndex        =   29
         Top             =   5520
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Index           =   1
         X1              =   1320
         X2              =   9000
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Margen"
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
         Left            =   7560
         TabIndex        =   28
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes "
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
         Left            =   4080
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Costes %"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   810
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "frmOliCrearTO1.frx":2446
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   24
         Top             =   720
         Width           =   210
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3720
         Picture         =   "frmOliCrearTO1.frx":24D1
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmOliCrearTO1.frx":255C
         ToolTipText     =   "Buscar cliente"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Index           =   0
         X1              =   1320
         X2              =   9000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
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
         Index           =   24
         Left            =   360
         TabIndex        =   20
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Generación Tarifa-Oferta"
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
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Modificar"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   12
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   13
      Top             =   7080
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   6495
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "PVP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Calculado"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   8040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5160
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar precio kilo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar articulo kilo"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Obtener articulos precios nuevos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir datos tarifa-oferta"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar tarifa-ofertas"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmOliCrearTO1.frx":265E
      Top             =   7140
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmOliCrearTO1.frx":27A8
      Top             =   7140
      Width           =   240
   End
End
Attribute VB_Name = "frmOliCrearTO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public vOpcion As Byte
    '0  .- Generacion TOs/Tarifas
    '1  .- Copiar tarifa
    '2  .- Exportar TO al AVAB
    
    
Public SegundoParametro As Long
    'Para vopcion=2  - Tarifa ORIGEN
    
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBu As frmBuscaGrid
Attribute frmBu.VB_VarHelpID = -1
Private WithEvents frmO As frmOliTarObservaciones
Attribute frmO.VB_VarHelpID = -1

Dim SQL As String



Private Sub Command1_Click()
    Me.Adodc1.Recordset.Update
End Sub


Private Sub cmdCancelaEdicion_Click()

    PonerModo False

End Sub







Private Sub cmdAccion_Click(Index As Integer)
    SQL = "¿Desea volver al paso anterior?"
    If Index = 0 Then
            
            If Me.lw1.visible Then
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                'Estamos en la fase 1. Con los articulos cargados pendientes de seleccionar. Si le
                'da a cancelar vuelve al punto 0. es decir , modificando KILOS
                Me.cmdAccion(0).visible = False
                PonerFase2 0
                
        
            Else
                    If Me.FrameFinal.visible Then
                        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                        'Esta en el ultimo trozo. Es decir, metiendo datos de  clientes
                        'Si le da a cancelar vuelve a la se
                        PonerFase2 1
                        
                
                    Else
                        'Modificar cantidades
                        PonerModo False
                    End If
            End If
    Else
        'index=1
        'modificar
        Updatear True
    End If
End Sub



Private Sub ClienteMorales(Index As Integer)
Dim It As ListItem

    
    If Index = 0 Then
        'Añadir
        If Me.txtCliente(0).Text = "" Then
            MsgBox "Seleccione un cliente", vbExclamation
            Exit Sub
        End If
        If lwCli.ListItems.Count > 0 Then
            For NumRegElim = 1 To lwCli.ListItems.Count
                If lwCli.ListItems(NumRegElim).Text = txtCliente(0).Text Then
                    MsgBox "Ya ha sido insertado", vbExclamation
                    txtCliente(0).Text = ""
                    txtDescClie(0).Text = ""
                End If
            Next
        End If
        
        'LO insertamos
        If txtCliente(0).Text <> "" Then
            Set It = lwCli.ListItems.Add()
            It.Text = txtCliente(0).Text
            It.SubItems(1) = Me.txtDescClie(0).Text
            txtCliente(0).Text = ""
            txtDescClie(0).Text = ""
        End If
        PonerFoco txtCliente(0)
    Else
        If lwCli.ListItems.Count = 0 Then Exit Sub
        
        If lwCli.SelectedItem Is Nothing Then Exit Sub
        
        SQL = "Desea quitar de la lista de clientes a: " & vbCrLf
        SQL = SQL & lwCli.SelectedItem.Text & " - " & lwCli.SelectedItem.SubItems(1)
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then lwCli.ListItems.Remove lwCli.SelectedItem.Index
    End If
End Sub



Private Sub ClienteAVAB(Index As Integer)
Dim It As ListItem

    
    If Index = 0 Then
        'Añadir
        If Me.txtCliente(1).Text = "" Then
            MsgBox "Seleccione un cliente", vbExclamation
            Exit Sub
        End If
        If lw2.ListItems.Count > 0 Then
            For NumRegElim = 1 To lw2.ListItems.Count
                If lw2.ListItems(NumRegElim).Text = txtCliente(1).Text Then
                    MsgBox "Ya ha sido insertado", vbExclamation
                    txtCliente(1).Text = ""
                    txtDescClie(1).Text = ""
                End If
            Next
        End If
        
        'LO insertamos
        If txtCliente(1).Text <> "" Then
            Set It = lw2.ListItems.Add()
            It.Text = txtCliente(1).Text
            It.SubItems(1) = Me.txtDescClie(1).Text
            txtCliente(1).Text = ""
            txtDescClie(1).Text = ""
        End If
        PonerFoco txtCliente(1)
    Else
        If lw2.ListItems.Count = 0 Then Exit Sub
        
        If lw2.SelectedItem Is Nothing Then Exit Sub
        
        SQL = "Desea quitar de la lista de clientes a: " & vbCrLf
        SQL = SQL & lw2.SelectedItem.Text & " - " & lw2.SelectedItem.SubItems(1)
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then lw2.ListItems.Remove lw2.SelectedItem.Index
    End If
End Sub


Private Sub cmdCliente_Click(Index As Integer)
    If Index < 2 Then
        ClienteMorales Index
    Else
        ClienteAVAB Index - 2
    End If
End Sub

Private Sub cmdCopiar_Click(Index As Integer)
    If Index = 0 Then
    
        If Me.txtTar(0).Text = "" Then
            MsgBox "Escriba la tarifa origen", vbExclamation
            Exit Sub
        End If
        Set miRsAux = New ADODB.Recordset
        SQL = "Select * from olitarifaoferta where codigo= " & Me.txtTar(0).Text
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
        If miRsAux.EOF Then
            
            MsgBox "Error obteniedo tarifa", vbExclamation
            miRsAux.Close
            Exit Sub
        Else
            If Not ComprobacionTarifaCopiada Then
                miRsAux.Close
                Exit Sub
            End If
        End If
        
        'Ahora realizamos la creacion
        
        SQL = SugerirCodigoSiguienteStr("olitarifaoferta", "codigo", "codigo < 100000")
        NumRegElim = Val(SQL)   'Codig para la insercion
        
        
        'La cabecera
        
        If Not CopiarTarifas Then
            Set miRsAux = Nothing
            Exit Sub
        End If
        
        
    End If
    Unload Me
End Sub

Private Sub cmdGenAVAB_Click(Index As Integer)
Dim HacerTarifa As Boolean
Dim b As Boolean

    If Index = 0 Then
        'Acciones
        
        
        'Comprobaciones iniciales
        If txtFecha(2).Text = "" Or txtFecha(3).Text = "" Then
            MsgBox "Fechas obligadas", vbExclamation
            Exit Sub
        End If
        
        If CDate(txtFecha(2).Text) > CDate(txtFecha(3).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        End If
        
        
        
        If lw2.ListItems.Count = 0 Then
        
            If Me.cboTarifaAVAB.ListIndex < 0 Then
                MsgBox "Introduzca algun cliente o una tarifa", vbExclamation
                Exit Sub
            End If
            HacerTarifa = True
            SQL = "Continuar con la generacion de la TO en AVAB?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Else
        
            'Ha seleccionado algun cliente
            SQL = ""
            If cboTarifaAVAB.ListIndex >= 0 Then
                'Aunque tenga marcada una tarifa, lo que manda es el cliente
                SQL = "Tiene insertados clientes y marcada una tarifa. " & vbCrLf
            End If
            SQL = SQL & "Va a generarar un TO en AVAB para los clientes seleccionados. ¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            HacerTarifa = False
            
        End If
        
        
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        b = GenerarTO_AVAB(HacerTarifa)
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        
        
        
        
        
        
        
        
        
        
        
        
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    Me.Icon = frmppal.Icon
    
    
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 5
        .Buttons(3).Image = 21
        .Buttons(5).Image = 36
        .Buttons(7).Image = 37
        .Buttons(9).Image = 15  'Salir
    End With
    Text1.visible = False
    limpiar Me
    Me.FrameFinal.visible = False
    FrameAVAB.visible = False
    FrameCopiarTarifa.visible = False
    Me.Toolbar1.visible = False
    'Creamos datos tmp
    Select Case vOpcion
    Case 0

                PonerFase2 0
                PonerModo False
    
                Conn.Execute "DELETE FROM olitmpto where codusu = " & vUsu.Codigo
                'Metemos los datos nuevos
                SQL = "INSERT INTO olitmpTO SELECT " & vUsu.Codigo & " ,codartic,precioUC from sartic where factorconversion<>1 and factorconversion >0"
                Conn.Execute SQL
                CargaGrid True
                
                '
                Text1.Height = DataGrid1.RowHeight
                Text1.Width = DataGrid1.Columns(4).Width
                

                
                Me.Height = 8085
                Me.Width = 9795
                Me.Toolbar1.visible = True
                'Carga combo tarifa
                CargaTarifas_ Me.cboTarifa
    Case 1
    
        'Copiar tarifa
        Me.FrameFinal.visible = False
        FrameCopiarTarifa.visible = True

        Me.Height = 3975
        Me.Width = 8415
        CargaTarifas_ Me.cboTarif2
        
    Case 2
        'Exportar TO
        Me.FrameAVAB.visible = True
        Me.Height = FrameAVAB.Height + 360
        Me.Width = FrameAVAB.Width + 60
        CargaTarifas_ Me.cboTarifaAVAB
    End Select
    

End Sub


Private Sub CargaGrid(enlaza As Boolean)

    SQL = "Select codusu,olitmpto.codartic,nomartic,preciouc,precioKilo from olitmpTO,sartic WHERE olitmpto.codartic =sartic.codartic and codusu = " & vUsu.Codigo
    If Not enlaza Then SQL = SQL & " AND codusu =-1"
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, True
    
    'Las dos primeras columnas visible FALSe
    DataGrid1.Columns(0).visible = False
    
    DataGrid1.Columns(1).Width = 1450

    
    DataGrid1.Columns(2).Caption = "Descripcion"
    DataGrid1.Columns(2).Width = 3550

    
    DataGrid1.Columns(3).Caption = "Coste"
    DataGrid1.Columns(3).Width = 1350
    DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(3).NumberFormat = FormatoPrecio
    
    DataGrid1.Columns(4).Caption = "Calculo"
    DataGrid1.Columns(4).Width = 1350
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = FormatoPrecio
                
        
End Sub



Private Sub frmBu_Selecionado(CadenaDevuelta As String)
    If vOpcion = 2 Then
        SQL = CadenaDevuelta   'cliente del AVAB
    Else
        txtTar(0).Text = RecuperaValor(CadenaDevuelta, 1)
    End If
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmO_Observaciones(Valor As String)
    Me.txtObserva(CInt(Me.txtObserva(0).Tag)).Text = Valor
End Sub

Private Sub Image1_Click()
    
    Set frmBu = New frmBuscaGrid
   
            SQL = "codigo|olitarifaoferta|codigo|T||14·"  'TARIFA  fechaini fechafin

            SQL = SQL & "Inicio|olitarifaoferta|fechaini|F||15·" 'TARIFA
            SQL = SQL & "Fin|olitarifaoferta|fechafin|F||15·" 'TARIFA
            SQL = SQL & "Tarifa|olitarifaoferta|tarifa|N||15·" 'TARIFA
            SQL = SQL & "Nombre|starif|nomlista|T||37·"
        
        Set frmBu = New frmBuscaGrid
        frmBu.vCampos = SQL
        frmBu.vTabla = " olitarifaoferta inner join starif on olitarifaoferta.tarifa = starif.codlista"
        frmBu.vSQL = ""

    
        frmBu.vDevuelve = "0|"
        frmBu.vTitulo = "Tarifas - Ofertas"
        frmBu.vselElem = 0
        frmBu.vConexionGrid = conAri 'Conexión a BD: Ariges
        
        frmBu.Show vbModal
        Set frmBu = Nothing
        PonerFoco txtTar(0)
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
        For NumRegElim = 1 To lw1.ListItems.Count
            lw1.ListItems(NumRegElim).Checked = (Index = 1)
        Next
    Else
    
    End If
End Sub

Private Sub imgCliente_Click(Index As Integer)
    SQL = ""
    If Index = 0 Then
        Set frmCli = New frmFacClientes
        frmCli.DatosADevolverBusqueda = "0|1|"
        frmCli.Show vbModal
        
    Else
        'Clientes en AVAB
            Set frmBu = New frmBuscaGrid
   
            SQL = "codigo|sclien|codclien|N||14·"  '
            SQL = SQL & "Nombre|sclien|nomclien|T||60·" 'nombre
            
        
        Set frmBu = New frmBuscaGrid
        frmBu.vCampos = SQL
        frmBu.vTabla = "ariges" & EmprAVAB & ".sclien"
        frmBu.vSQL = ""

    
        frmBu.vDevuelve = "0|1|"
        frmBu.vTitulo = "Clientes AVAB"
        frmBu.vselElem = 1
        frmBu.vConexionGrid = conAri 'Conexión a BD: Ariges
        SQL = ""
        frmBu.Show vbModal
        
    End If
    If SQL <> "" Then
        Me.txtCliente(Index).Text = RecuperaValor(SQL, 1)
        Me.txtDescClie(Index).Text = RecuperaValor(SQL, 2)
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    SQL = ""
    frmC.Show vbModal
    If SQL <> "" Then txtFecha(Index).Text = SQL
    Set frmC = Nothing
End Sub

Private Sub imgObserva_Click(Index As Integer)
    Me.txtObserva(0).Tag = Index
    Set frmO = New frmOliTarObservaciones
    frmO.Text1 = Me.txtObserva(Index).Text
    frmO.Show vbModal
    Set frmO = Nothing
    
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Para habilitar las flechas
   ' KEYdownLineas KeyCode
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
        KEYpressGnral KeyAscii, 3, False

End Sub

Private Sub Text1_LostFocus()
Dim HacerTarifa As Boolean
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then Exit Sub
    If Not PonerFormatoDecimal(Text1, 2) Then
        Text1.Text = ""
        PonerFoco Text1
    End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim b As Boolean
Dim HacerTarifa As Boolean
    Select Case Button.Index
    Case 1
        If Text1.visible Then Exit Sub 'YA esta modificando
        PonerModo True
        PonerFoco Text1
    Case 2
        If Text1.visible Then
            MsgBox "Esta editando cantidades", vbExclamation
            Exit Sub
        End If
        Eliminar
    
    Case 3
        If Text1.visible Then
            MsgBox "Esta editando cantidades", vbExclamation
            Exit Sub
        End If
        
        CadenaDesdeOtroForm = ""
        frmListado2.Opcion = 16
        frmListado2.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            'OK. Ha seleccionado desde hasta
            'Veremos si hay alguna articulo en el select
            
            SQL = "Select sartic.codartic,sartic.nomartic,preciove,codunida,margecom  " & CadenaDesdeOtroForm
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                miRsAux.MoveNext
            Wend
            
            
            
            
            If NumRegElim > 0 Then
                pb1.Value = 0
                pb1.Max = NumRegElim
                lw1.ListItems.Clear
                PonerFase2 1
                pb1.visible = True
                DoEvents
                CargamosArticulosConNuevoPrecio
                pb1.visible = False
                'Ponemos visble el selecc ALL
                imgCheck(0).visible = True
                imgCheck(1).visible = True
                Me.cmdAccion(0).visible = True 'Por si quiere volver al paso anterior
                Me.Refresh
            Else
                'Ninugun datos
                MsgBox "Ningun datos devuelto con estos valores.", vbExclamation
                miRsAux.Close
                
            End If
            Set miRsAux = Nothing
        End If
    Case 5
        'Compruebo si hay alguno seleccionado
        SQL = ""
        For NumRegElim = 1 To lw1.ListItems.Count
            If lw1.ListItems(NumRegElim).Checked Then SQL = SQL & "X"
        Next NumRegElim
        If SQL = "" Then
            MsgBox "Seleccione algun articulo para continuar con el proceso de generación de tarifas ofertas", vbExclamation
            Exit Sub
        End If
        
        SQL = Len(SQL)
        SQL = "Ha selecccionado " & SQL & " artículo(s).     ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        PonerFase2 2
        PonerFoco Me.txtFecha(0)
        
        
    Case 7
        'Comprobaciones iniciales
        If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
            MsgBox "Fechas obligadas", vbExclamation
            Exit Sub
        End If
        
        If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        End If
        
        
        
        If lwCli.ListItems.Count = 0 Then
        
            If Me.cboTarifa.ListIndex < 0 Then
                MsgBox "Introduzca algun cliente o una tarifa", vbExclamation
                Exit Sub
            End If
            HacerTarifa = True
            SQL = "Continuar con la generacion de la " & SQL & "?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Else
        
            'Ha seleccionado algun cliente
            SQL = ""
            If cboTarifa.ListIndex >= 0 Then
                'Aunque tenga marcada una tarifa, lo que manda es el cliente
                SQL = "Tiene insertados clientes y marcada una tarifa. " & vbCrLf
            End If
            SQL = SQL & "Va a generarar un TO para los clientes seleccionados. ¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            HacerTarifa = False
            
        End If
        
        
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        b = GenerarTOs(HacerTarifa)
        pb1.visible = False
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        If b Then Unload Me  'Ha ido bien. Chapo
        
        
    Case 9
        Unload Me
    End Select
End Sub


Private Sub PonerModo(Modificando As Boolean)
    If Modificando Then PosicionarCampo
    Text1.visible = Modificando
    DataGrid1.Enabled = Not Modificando
    Me.cmdAccion(0).visible = Modificando
    Me.cmdAccion(1).visible = Modificando
End Sub

Private Sub PosicionarCampo()
        DeseleccionaGrid Me.DataGrid1
        Text1.Left = DataGrid1.Columns(4).Left + 150
        Text1.Top = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
        Text1.Text = DataGrid1.Columns(4).Text
End Sub

Private Sub Updatear(SeguirUpdateando As Boolean)
    'Lanzamos el UPDATE
    If Text1.Text = "" Then Exit Sub
    If Text1.Text <> DataGrid1.Columns(4).Text Then
        SQL = "UPDATE olitmpto set preciokilo=" & TransformaComasPuntos(ImporteFormateado(Text1.Text))
        SQL = SQL & " WHERE codartic = '" & DataGrid1.Columns(1).Text & "' AND codusu = " & vUsu.Codigo
        Conn.Execute SQL
        Espera 0.2

    End If
    'Actualimos grid
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    CargaGrid True
    
    'Si no es el ultimo ponemos a modificar
    If Adodc1.Recordset.RecordCount = NumRegElim Then
        'Estaba en le ultimo. No hago nada
        NumRegElim = 0
    
    End If
    If NumRegElim > 0 Then Adodc1.Recordset.Move NumRegElim
    If SeguirUpdateando Then
        PosicionarCampo
        PonerFoco Text1
    Else
        PonerModo False
    End If
End Sub

'Fases
'   0- modioifcando kilos
'   1- seleccionando articulos
'   2- Introduciendo datos para hacer el desde hasta
Private Sub PonerFase2(vOp As Byte)
    
    DataGrid1.visible = (vOp = 0)
    Me.Toolbar1.Buttons(1).Enabled = vOp = 0
    Me.Toolbar1.Buttons(2).Enabled = vOp = 0
    Me.Toolbar1.Buttons(3).Enabled = (vOp = 0)
    Me.Toolbar1.Buttons(5).Enabled = (vOp = 1)
    Me.Toolbar1.Buttons(7).Enabled = (vOp = 2)
    
    If vOp >= 1 Then Text1.visible = False
    
    
    If vOp <> 1 Then
        'Si no estamos en la fase de cargar los articlulos, NO se veran seguro los checks de selccionar todo
        imgCheck(0).visible = False
        imgCheck(1).visible = False
    End If
    lw1.visible = (vOp = 1)
    


    FrameFinal.visible = (vOp = 2)
    If vOp = 0 Then
        Me.cmdAccion(0).Caption = "Cancelar"
    Else
        Me.cmdAccion(0).Caption = "Atras"
    End If

End Sub





Private Sub CargamosArticulosConNuevoPrecio()
Dim cArt As Collection
Dim RsNuevoPrecioKilo As ADODB.Recordset
Dim Importe As Currency
Dim PVP As Currency
Dim TipoUdAnterior As Integer
Dim ValorUd As Currency
Dim PVPCalculado As Currency
Dim It As ListItem

        On Error GoTo ECargamosArticulosConNuevoPrecio
        
        Set cArt = New Collection
        
        'Precargo los costes derivados por el tipo de formato(unidad)
        Set RsNuevoPrecioKilo = New ADODB.Recordset
        SQL = "Select * from olitmpto where codusu = " & vUsu.Codigo
        RsNuevoPrecioKilo.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        
        miRsAux.MoveFirst 'Ya esta cargado el RS
        While Not miRsAux.EOF
            cArt.Add CStr(miRsAux.Fields(0)) & "|" & miRsAux!NomArtic & "|" & miRsAux!CodUnida & "|" & Format(miRsAux!preciove, FormatoPrecio) & "|" & DBLet(miRsAux!margecom, "N") & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        TipoUdAnterior = -1
        
        'Ya tengo todos los articulos para mostrarles como van a quedar los precios
        For NumRegElim = 1 To cArt.Count
            pb1.Value = pb1.Value + 1
            SQL = cArt(NumRegElim)
            SQL = RecuperaValor(SQL, 1)
            SQL = DevNombreSQL(SQL)
            
            SQL = "select cantidad,preciouc,codarti1,factorconversion from sarti1,sartic where sarti1.codarti1 =sartic.codartic and sarti1.codartic ='" & SQL & "'"
            
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            PVPCalculado = 0
            While Not miRsAux.EOF
                Importe = DBLet(miRsAux!FactorConversion, "N")  'del articulo de la linea
                If Importe <> 1 Then
                    RsNuevoPrecioKilo.Find "codartic = '" & DevNombreSQL(miRsAux!codarti1) & "'", , adSearchForward, 1
                    If Not RsNuevoPrecioKilo.EOF Then
                        'Precio de la tabla tmporal
                        PVP = RsNuevoPrecioKilo!preciokilo
                    Else
                        PVP = DBLet(miRsAux!precioUC, "N")
                    End If
                    PVP = PVP * Importe
                Else
                    PVP = DBLet(miRsAux!precioUC, "N")
                End If
                'PVP
                PVP = DBLet(miRsAux!Cantidad, "N") * PVP
                PVPCalculado = PVPCalculado + PVP
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            'Para evitar que haga un select por cada articul
            SQL = cArt(NumRegElim)
            SQL = RecuperaValor(SQL, 3)
            If Val(SQL) <> TipoUdAnterior Then
                'Añadimos los costes derivados por el tipo de formato(unidad)
                'select sum(importe) from sunilin where codunida =1              codigounidad
                TipoUdAnterior = SQL
                SQL = DevuelveDesdeBD(conAri, "sum(importe)", "sunilin", "codunida", SQL)
                If SQL = "" Then SQL = "0"
                ValorUd = CCur(SQL)
                
            End If
            PVPCalculado = PVPCalculado + ValorUd
            
            'Le sumo el margen.
            SQL = cArt(NumRegElim)
            SQL = RecuperaValor(SQL, 5)
            If SQL = "" Then SQL = "0"
            Importe = CCur(SQL)
            Importe = ((PVPCalculado * Importe) / 100)
            PVPCalculado = PVPCalculado + Importe
            
            'Ya tenemos los dos importes, el antiguo y el recalculado
            'Añadimos el listview
            
            
            Set It = lw1.ListItems.Add()
            It.Text = RecuperaValor(cArt(NumRegElim), 1)
            It.SubItems(1) = RecuperaValor(cArt(NumRegElim), 2)
            It.SubItems(2) = RecuperaValor(cArt(NumRegElim), 4) 'precio
            It.SubItems(3) = Format(PVPCalculado, FormatoPrecio)
            It.Checked = True
            If (pb1.Value Mod 100) = 1 Then DoEvents
        Next NumRegElim
        
        
        
    RsNuevoPrecioKilo.Close
ECargamosArticulosConNuevoPrecio:
    If Err.Number <> 0 Then MuestraError Err.Number, ""
    Set cArt = Nothing
    Set miRsAux = Nothing
    Set RsNuevoPrecioKilo = Nothing
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
Dim Descri As String
    
    Descri = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            txtCliente(Index) = ""
            PonerFoco txtCliente(Index)
        Else
            If Index = 1 Then
                'PARA EL AVABA
                Descri = DevuelveDesdeBD(conAri, "nomclien", "ariges" & EmprAVAB & ".sclien", "codclien", txtCliente(Index).Text, "N")
            Else
                Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            End If
            If Descri = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
                txtCliente(Index).Text = ""
                PonerFoco txtCliente(Index)
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Descri
    
    
    
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub



Private Function GenerarTOs(HacerTarifa As Boolean) As Boolean
Dim C As Long
Dim CL As Collection
Dim Insert As String
Dim i As Integer
Dim Bucle As Integer


    On Error GoTo EComprobacionClientes
    GenerarTOs = False
    'Comprobacion
    '-----------------------------------------------------------------------------
    '
    'Para cada cliente veremos si tiene una oferta entre esas fechas
    '
    'Guardaremos los articulos en un tmp para poder hacer joins en el select
    Conn.Execute "DELETE FROM tmpstockfec where codusu = " & vUsu.Codigo
    'inserto el lw
    SQL = ""
    C = 0
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then
            SQL = SQL & ", (0," & vUsu.Codigo & ",'" & DevNombreSQL(lw1.ListItems(NumRegElim).Text) & "')"
            C = C + 1 'Para saber cuantos hay
        End If
    Next
    SQL = Mid(SQL, 2) 'QUITO LA PRIMERA COMA
    SQL = "insert into `tmpstockfec` (`codalmac`,`codusu`,`codartic`) VALUES " & SQL
    Conn.Execute SQL
        
     
    If lwCli.ListItems.Count > 2 Then
        pb1.Value = 0
        pb1.Max = lwCli.ListItems.Count
        pb1.visible = True
    End If
    
    'NO COMPROBAR. Lo dijo RAMON
    If HacerTarifa Then
        If Not ComprobacionTarifas() Then Exit Function
    End If
    
    'Y ahora, finalmente insertamos
    
    If pb1.visible Then pb1.Value = 0
    DoEvents
    
    Insert = ">"
    If HacerTarifa Then Insert = "<"
    Insert = "codigo " & Insert & " 100000"
    SQL = SugerirCodigoSiguienteStr("olitarifaoferta", "codigo", Insert)
    NumRegElim = Val(SQL)   'Codig para la insercion
    If Not HacerTarifa Then
        If NumRegElim < 100000 Then NumRegElim = 100001
    End If
    Insert = ""
    
    
    'Cargo el RS con los valores de los articulos(sartic)
    SQL = "select sartic.codartic,LitrosUnidad,margecom from sartic,tmpstockfec where codusu=" & vUsu.Codigo
    SQL = SQL & " and tmpstockfec.codartic=sartic.codartic"
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Comprobacion 1
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & "X"
        miRsAux.MoveNext
    Wend
    miRsAux.MoveFirst
    
    If Len(SQL) <> C Then
        MsgBox "No coinciden los datos grabados con el total de articulos.", vbExclamation
        miRsAux.Close
        Exit Function
    End If
    
    'En el rs tengo los datos del articulo
    Set CL = New Collection
    
    CrearLineasTOS CL
    
    CadenaDesdeOtroForm = "insert into `olitarifaofertalin` (`codigo`,`codartic`,`pivu`,`pivl`,`coste1`,`coste2`,`coste3`,coste4,coste5,`margen`,`pfvu`,`pfvl`) values "
    
    
    If HacerTarifa Then
        Bucle = 1  'TARIFA. Solo lo hara una vez
    Else
        Bucle = lwCli.ListItems.Count
    End If
    
    For C = 1 To Bucle
        If pb1.visible Then pb1.Value = pb1.Value + 1
        
        
        'Cabecera
        
        SQL = "insert into `olitarifaoferta` (`codigo`,`codclien`,`fechaini`,`fechafin`,`aceptada`,`tarifa`,`observaciones`) values ("
        SQL = SQL & NumRegElim & ","
        If HacerTarifa Then
            SQL = SQL & "NULL"   'TARIFA. Cliente a NULL
        Else
            SQL = SQL & Val(lwCli.ListItems(C).Text)
        End If                                                                              'aceptada
        SQL = SQL & "," & DBSet(txtFecha(0).Text, "F") & "," & DBSet(txtFecha(1).Text, "F") & ",0,"
        If HacerTarifa Then
            SQL = SQL & cboTarifa.ItemData(cboTarifa.ListIndex)
        Else
            SQL = SQL & "NULL"   'TARIFA  a NULL
        End If
        
        'Nuevo Dic 2009
        'Observaciones
        SQL = SQL & "," & DBSet(txtObserva(0).Text, "T")
        
        SQL = SQL & ")"
        Conn.Execute SQL
        
        
        'lineas
        Insert = ""
        For i = 1 To CL.Count
            Insert = Insert & ", (" & NumRegElim & CL(i)
        Next
        
        Insert = Mid(Insert, 2)
        SQL = CadenaDesdeOtroForm & Insert
        Conn.Execute SQL
        
        'Insertamos las lineas 2
        InsertaEnLineas2 SQL
        
        NumRegElim = NumRegElim + 1
        Me.Refresh
        Espera 0.2
    Next
    
    
    
    
    
    
    GenerarTOs = True
    
    
    Exit Function
EComprobacionClientes:
    MuestraError Err.Number, Err.Description
    Set CL = Nothing
End Function

Private Sub InsertaEnLineas2(ByRef MaterPrimaNoExisteAVAB As String)
Dim RT As ADODB.Recordset
Dim SQL As String

    On Error GoTo E1
    'Inserto en olitarifaofertalin2 que llevara el incio d ela simulacion
    If vOpcion = 2 Then
        
        
           'Veo si hay arituclos K no existen en AVAB
           Set RT = New ADODB.Recordset
            SQL = "Select * from olitarifaofertalin2 where codigo = " & Me.SegundoParametro
            SQL = SQL & " AND not codartic in (select codartic from ariges" & EmprAVAB & ".sartic )"
            RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            
            While Not RT.EOF
                If InStr(1, MaterPrimaNoExisteAVAB, "," & RT!codArtic) = 0 Then
                    'No existe y ademas no estaba
                    MaterPrimaNoExisteAVAB = MaterPrimaNoExisteAVAB & "," & RT!codArtic
                
                End If
                RT.MoveNext
            Wend
            RT.Close
            Set RT = Nothing
        
        
        
        SQL = "INSERT IGNORE INTO ariges" & EmprAVAB & ".olitarifaofertalin2 (`codigo`,`codartic`,`costereal`,`costesimul`) "
        SQL = SQL & "select " & NumRegElim & ",`codartic`,`costereal`,`costesimul` from olitarifaofertalin2 where "
        SQL = SQL & " codigo = " & SegundoParametro
        Conn.Execute SQL
    
    Else
        SQL = "INSERT INTO `olitarifaofertalin2` (`codigo`,`codartic`,`costereal`,`costesimul`) "
        SQL = SQL & "select " & NumRegElim & ",olitmpto.codartic,preciouc,preciokilo from olitmpto,sartic where "
        SQL = SQL & "olitmpto.codartic=sartic.codartic AND codusu = " & vUsu.Codigo
        Conn.Execute SQL
    End If
    
   Exit Sub
E1:
        MuestraError Err.Number, Err.Description, "Insertando en lineas(2). Proceso continua"
End Sub

Private Function ComprobacionTarifas() As Boolean

    
    ComprobacionTarifas = False

    SQL = "Select * from olitarifaoferta where codclien is null "
    SQL = SQL & " AND tarifa = " & Me.cboTarifa.ItemData(cboTarifa.ListIndex) & " AND ("
    'Ninguno dentro del intervalo
    SQL = SQL & " ( fechaini>=" & DBSet(txtFecha(0).Text, "F") & " AND fechafin<=" & DBSet(txtFecha(1).Text, "F") & ")"
    'Fecha inicio NO dentro intervalo
    SQL = SQL & " OR ( fechaini<=" & DBSet(txtFecha(0).Text, "F") & " AND fechafin>=" & DBSet(txtFecha(0).Text, "F") & ")"
    'Fecha FIN no dentro intervalor
    SQL = SQL & " or ( fechaini<=" & DBSet(txtFecha(1).Text, "F") & " AND fechafin>=" & DBSet(txtFecha(1).Text, "F") & ")"
    SQL = SQL & ")"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & ", " & miRsAux!Codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL = "" Then
        'OK, no tiene ofertas el cliente
        ComprobacionTarifas = True
        Exit Function
    End If
    
    
    MsgBox "La nueva tarifa se solapa con la(s) tarifa(s): " & SQL, vbExclamation
     
    
'    'El cliente tiene ofertas. Tenemos que ver si algun aritculo de lo que se esta ofertando AHORA
'    'ya esta en una oferta para ese cliente
'    SQL = Mid(SQL, 2) 'le quito la coma
'    SQL = "select codigo,olitarifaofertalin.codartic from olitarifaofertalin,tmpstockfec where codusu=" & vUsu.Codigo & _
'        " and tmpstockfec.codartic=olitarifaofertalin.codartic and codigo in (" & SQL & ")"
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    SQL = ""
'    While Not miRsAux.EOF
'        'sql = sql & Format(miRsAux!Codigo, "00000") & "       " & miRsAux!codArtic & vbCrLf
'        SQL = SQL & Format(miRsAux!Codigo, "00000") & "    "
'
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'    If SQL = "" Then
'        ComprobacionTarifas = True
'    Else
'        SQL = String(40, "-") & vbCrLf & SQL
'        SQL = "Oferta           Articulo" & vbCrLf & SQL
'        SQL = "El cliente: " & lwCli.ListItems(Indice).Text & " - " & lwCli.ListItems(Indice).SubItems(1) & vbCrLf & _
'            "ya tiene ofertas en esas fechas para los articulos. " & vbCrLf & "(" & Trim(SQL) & ")"
'        MsgBox SQL, vbExclamation
'    End If

End Function


Private Sub CrearLineasTOS(ByRef Col As Collection)
Dim i As Integer
Dim Aux As Currency
Dim LitrosUd As Currency
Dim Margen As Currency
Dim C1(4) As Currency

    'Valores comunes
    C1(0) = ImporteFormateado(txtNumero(0).Text)
    C1(1) = ImporteFormateado(txtNumero(1).Text)
    C1(2) = ImporteFormateado(txtNumero(4).Text)
    C1(3) = ImporteFormateado(txtNumero(2).Text)
    C1(4) = ImporteFormateado(txtNumero(5).Text)

    'margen. Si ha puesto algo, ponemos ese, si no el del articulo
    If txtNumero(3).Text <> "" Then
        'Ha puesto algo
        Margen = ImporteFormateado(txtNumero(3).Text)
    End If

    For i = 1 To lw1.ListItems.Count
        If lw1.ListItems(i).Checked Then
            miRsAux.Find "codartic = '" & DevNombreSQL(lw1.ListItems(i).Text) & "'"
'            if mirsaux.EOF   no hago el control. Si da error que acabe
            SQL = ",'" & DevNombreSQL(miRsAux!codArtic) & "',"
            LitrosUd = DBLet(miRsAux!LitrosUnidad, "N")
            If LitrosUd = 0 Then LitrosUd = 1 'Para que no de error
            
            'Cogemos el del articulo
            If txtNumero(3).Text = "" Then Margen = DBLet(miRsAux!margecom, "N")
            
            'PRECIO VENTA UNITARIO
            Aux = ImporteFormateado(lw1.ListItems(i).SubItems(3))
            
            '`pivu`
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
            'pivl
            SQL = SQL & TransformaComasPuntos(CStr(Round2(Aux / LitrosUd, 4))) & ","
            
            '`coste1`,`coste2`,`coste3`,coste4,coste5,margen`
            SQL = SQL & DBSet(C1(0), "N") & "," & DBSet(C1(1), "N") & ","
            SQL = SQL & DBSet(C1(2), "N") & "," & DBSet(C1(3), "N") & ","
            SQL = SQL & DBSet(C1(4), "N") & "," & DBSet(Margen, "N") & ","
            
            
            Aux = CalculaImporteLineaTO(Aux, C1(0), C1(1), C1(2), C1(3), C1(4), Margen, LitrosUd)
            
            
            '`pfvu`,
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
            '`pfvl`
            
            'Marzo 2010. Redondeamos a 3 decimales
            Aux = Round2(Aux / LitrosUd, 3)   'antes ponia 4
            
            
            
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ")"
            Col.Add SQL
        End If
    Next
    
    
    
End Sub




Private Sub CrearLineasTOSAVAB(ByRef Col As Collection, ByRef ArticulosInexistentes2 As String)
Dim i As Integer
Dim Aux As Currency
Dim LitrosUd As Currency
Dim Margen As Currency
Dim C1(4) As Currency

    'Valores comunes
    C1(0) = ImporteFormateado(txtNumero(12).Text)
    C1(1) = ImporteFormateado(txtNumero(13).Text)
    C1(2) = ImporteFormateado(txtNumero(14).Text)
    C1(3) = ImporteFormateado(txtNumero(15).Text)
    C1(4) = ImporteFormateado(txtNumero(16).Text)

    'margen. Si ha puesto algo, ponemos ese, si no el del articulo
    If txtNumero(17).Text <> "" Then
        'Ha puesto algo
        Margen = ImporteFormateado(txtNumero(17).Text)
    End If

    While Not miRsAux.EOF
            If InStr(1, ArticulosInexistentes2, miRsAux!codArtic) = 0 Then
                
'            if mirsaux.EOF   no hago el control. Si da error que acabe
                SQL = ",'" & DevNombreSQL(miRsAux!codArtic) & "',"
                LitrosUd = DBLet(miRsAux!LitrosUnidad, "N")
                If LitrosUd = 0 Then LitrosUd = 1 'Para que no de error
                
                'Cogemos el del articulo
                If txtNumero(17).Text = "" Then Margen = DBLet(miRsAux!margecom, "N")
                
                'PRECIO VENTA UNITARIO
                'Precio venta unitario es el PVPFinal
                
                Aux = miRsAux!pfvu
                
                '`pivu`
                SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
                'pivl
                
                SQL = SQL & TransformaComasPuntos(CStr(miRsAux!pfvl)) & ","
                
                '`coste1`,`coste2`,`coste3`,coste4,coste5,margen`
                SQL = SQL & DBSet(C1(0), "N") & "," & DBSet(C1(1), "N") & ","
                SQL = SQL & DBSet(C1(2), "N") & "," & DBSet(C1(3), "N") & ","
                SQL = SQL & DBSet(C1(4), "N") & "," & DBSet(Margen, "N") & ","
                
                
                Aux = CalculaImporteLineaTO(Aux, C1(0), C1(1), C1(2), C1(3), C1(4), Margen, LitrosUd)
                
                
                '`pfvu`,
                SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
                '`pfvl`
                Aux = Round2(Aux / LitrosUd, 4)
                SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ")"
                Col.Add SQL
            End If
            miRsAux.MoveNext
        Wend
    
    
End Sub



Private Sub txtNumero_GotFocus(Index As Integer)
    ConseguirFoco txtNumero(Index), 3
End Sub


Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
Dim b As Boolean
    txtNumero(Index).Text = Trim(txtNumero(Index))
    If txtNumero(Index).Text = "" Then Exit Sub
    
    If Index = 2 Or Index = 5 Or Index = 9 Or Index = 10 Or Index = 15 Or Index = 16 Then
        b = PonerFormatoDecimal(txtNumero(Index), 2)
    Else
        '
        b = PonerFormatoDecimal(txtNumero(Index), 4)
    End If
    If Not b Then
        txtNumero(Index).Text = ""
        PonerFoco txtNumero(Index)
    End If
End Sub

Private Sub Eliminar()
    If Me.Adodc1.Recordset.EOF Then Exit Sub
    
    SQL = "¿Desea eliminar el componente " & Adodc1.Recordset!codArtic & " " & Adodc1.Recordset!NomArtic & "?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    SQL = "DELETE FROM olitmpto where codusu = " & vUsu.Codigo & " AND codartic = '" & Adodc1.Recordset!codArtic & "'"
    Conn.Execute SQL
    CargaGrid True
End Sub



Private Sub CargaTarifas_(ByRef C As ComboBox)
    On Error GoTo ECargaTarifas
    C.Clear
    SQL = "Select * from "
    If vOpcion = 2 Then SQL = SQL & "ariges" & EmprAVAB & "."
    SQL = SQL & "starif order by codlista"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        C.AddItem miRsAux!nomlista
        C.ItemData(C.NewIndex) = miRsAux!codlista
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
ECargaTarifas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub

Private Sub txtTar_GotFocus(Index As Integer)
    ConseguirFoco txtTar(Index), 3
End Sub

Private Sub txtTar_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTar_LostFocus(Index As Integer)
    If Index <> 0 Then Exit Sub
    txtTar(0).Text = Trim(txtTar(0).Text)
    If txtTar(0).Text = "" Then
        txtTar(1).Text = ""
    Else
        If Not PonerFormatoEntero(txtTar(0)) Then
            txtTar(0).Text = ""
            txtTar(1).Text = ""
            PonerFoco txtTar(0)
        Else
            Set miRsAux = New ADODB.Recordset
            SQL = "Select nomlista,fechaini,fechafin from olitarifaoferta inner join starif on olitarifaoferta.tarifa = starif.codlista where codigo = " & txtTar(0).Text
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                SQL = ""
                MsgBox "No existe oferta: " & txtTar(0).Text, vbExclamation
            Else
                SQL = miRsAux!nomlista & "    Fechas: " & Format(miRsAux!FechaIni, "dd/mm/yyyy") & " - " & Format(miRsAux!FechaFin, "dd/mm/yyyy")
            End If
            miRsAux.Close
            Set miRsAux = Nothing
            txtTar(1).Text = SQL
            If SQL = "" Then
                txtTar(0).Text = ""
                PonerFoco txtTar(0)
            End If
        End If
    End If
End Sub



Private Sub CrearLineasCopiaTOs()

    
End Sub


'mirsaux lleva la fecha de inicio etc etc
Private Function ComprobacionTarifaCopiada() As Boolean
Dim R As ADODB.Recordset


    
    ComprobacionTarifaCopiada = False
    Set R = New ADODB.Recordset
    SQL = "Select * from olitarifaoferta where codclien is null "
    SQL = SQL & " AND tarifa = " & Me.cboTarif2.ItemData(cboTarif2.ListIndex) & " AND ("
    'Ninguno dentro del intervalo
    SQL = SQL & " ( fechaini>=" & DBSet(miRsAux!FechaIni, "F") & " AND fechafin<=" & DBSet(miRsAux!FechaFin, "F") & ")"
    'Fecha inicio NO dentro intervalo
    SQL = SQL & " OR ( fechaini<=" & DBSet(miRsAux!FechaIni, "F") & " AND fechafin>=" & DBSet(miRsAux!FechaIni, "F") & ")"
    'Fecha FIN no dentro intervalor
    SQL = SQL & " or ( fechaini<=" & DBSet(miRsAux!FechaFin, "F") & " AND fechafin>=" & DBSet(miRsAux!FechaFin, "F") & ")"
    SQL = SQL & ")"
    R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not R.EOF
        SQL = SQL & ", " & R!Codigo
        R.MoveNext
    Wend
    R.Close
    Set R = Nothing
    If SQL = "" Then
        'OK, no tiene ofertas el cliente
        ComprobacionTarifaCopiada = True
        Exit Function
    End If
    SQL = Mid(SQL, 2)
    
    MsgBox "La nueva tarifa se solapa con la(s) tarifa(s): " & SQL, vbExclamation

End Function



Private Function CopiarTarifas() As Boolean
Dim Aux As Currency
Dim Margen As Currency
Dim LitrosUd As Currency
Dim C1(4) As Currency
Dim CADENA As String
    On Error GoTo ECopiarTarifas
    CopiarTarifas = False

        SQL = "insert into `olitarifaoferta` (`codigo`,`codclien`,`fechaini`,`fechafin`,`aceptada`,`tarifa`) values ("
        SQL = SQL & NumRegElim & ","
        
        SQL = SQL & "NULL"   'TARIFA. Cliente a NULL
                                                
        SQL = SQL & "," & DBSet(miRsAux!FechaIni, "F") & "," & DBSet(miRsAux!FechaFin, "F") & ",0,"
        SQL = SQL & cboTarif2.ItemData(cboTarif2.ListIndex)
        
        SQL = SQL & ")"
        Conn.Execute SQL

        miRsAux.Close
        
        
        'Valores comunes
        C1(0) = ImporteFormateado(txtNumero(6).Text)
        C1(1) = ImporteFormateado(txtNumero(7).Text)
        C1(2) = ImporteFormateado(txtNumero(8).Text)
        C1(3) = ImporteFormateado(txtNumero(9).Text)
        C1(4) = ImporteFormateado(txtNumero(10).Text)

        'margen. Si ha puesto algo, ponemos ese, si no el del articulo
        If txtNumero(11).Text <> "" Then
            'Ha puesto algo
            Margen = ImporteFormateado(txtNumero(11).Text)
        End If
            
        'Vamos con la slineas
        SQL = "select olitarifaofertalin.*,litrosunidad from olitarifaofertalin ,sartic where "
        SQL = SQL & " olitarifaofertalin.codartic=sartic.codartic AND codigo = " & txtTar(0).Text
        CADENA = ""
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
           
    


    
    
            
'            if mirsaux.EOF   no hago el control. Si da error que acabe
            SQL = ", (" & NumRegElim & ",'" & DevNombreSQL(miRsAux!codArtic) & "',"
            
            LitrosUd = DBLet(miRsAux!LitrosUnidad, "N")
            If LitrosUd = 0 Then LitrosUd = 1 'Para que no de error
            
            'Cogemos el del articulo
            If txtNumero(11).Text = "" Then Margen = DBLet(miRsAux!Margen, "N")
            
            'PRECIO VENTA UNITARIO
            Aux = miRsAux!PIVU
            
            '`pivu`
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
            'pivl
            SQL = SQL & TransformaComasPuntos(CStr(Round2(Aux / LitrosUd, 4))) & ","
            
            '`coste1`,`coste2`,`coste3`,coste4,coste5,margen`
            SQL = SQL & DBSet(C1(0), "N") & "," & DBSet(C1(1), "N") & ","
            SQL = SQL & DBSet(C1(2), "N") & "," & DBSet(C1(3), "N") & ","
            SQL = SQL & DBSet(C1(4), "N") & "," & DBSet(Margen, "N") & ","
            
            
            Aux = CalculaImporteLineaTO(Aux, C1(0), C1(1), C1(2), C1(3), C1(4), Margen, LitrosUd)
            
            
            '`pfvu`,
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ","
            '`pfvl`
            Aux = Round2(Aux / LitrosUd, 4)
            SQL = SQL & TransformaComasPuntos(CStr(Aux)) & ")"
            CADENA = CADENA & SQL
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        If CADENA <> "" Then
            SQL = "insert into `olitarifaofertalin` (`codigo`,`codartic`,`pivu`,`pivl`,`coste1`,`coste2`,`coste3`,coste4,coste5,`margen`,`pfvu`,`pfvl`) values "
            CADENA = Mid(CADENA, 2)
            Conn.Execute SQL & CADENA
            
            
            
            'Copiaos tb desde las materias primas
            SQL = "INSERT INTO olitarifaofertalin2 (`codigo`,`codartic`,`costereal`,`costesimul`) SELECT"
            SQL = SQL & " " & NumRegElim & ",`codartic`,`costereal`,`costesimul` FROM "
            SQL = SQL & " olitarifaofertalin2 where codigo =" & txtTar(0).Text
            Conn.Execute SQL
        End If
        CopiarTarifas = True
    Exit Function
ECopiarTarifas:
        MuestraError Err.Number, "Copiar Tarifas"
End Function








Private Function GenerarTO_AVAB(HacerTarifa As Boolean) As Boolean
Dim C As Long
Dim CL As Collection
Dim Insert As String
Dim i As Integer
Dim Bucle As Integer
Dim ArticulosInexistentes As String
Dim materiaPrimaNoexiste As String

    On Error GoTo ETOAVAB

    GenerarTO_AVAB = False

    
    
    
    
    'Veo si hay arituclos K no existen en AVAB
    SQL = "Select * from olitarifaofertalin where codigo = " & Me.SegundoParametro
    SQL = SQL & " AND not codartic in (select codartic from ariges" & EmprAVAB & ".sartic )"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    ArticulosInexistentes = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        ArticulosInexistentes = ArticulosInexistentes & miRsAux!codArtic & "|"
        SQL = SQL & miRsAux!codArtic & "  " & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If SQL <> "" Then
        SQL = "No existen articulos(" & NumRegElim & ") en AVAB" & vbCrLf & vbCrLf & SQL
        SQL = SQL & vbCrLf & "¿Coninuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
        ArticulosInexistentes = "|" & ArticulosInexistentes
    End If
    
    
    Insert = ">"
    If HacerTarifa Then Insert = "<"
    Insert = "codigo " & Insert & " 100000"
    SQL = SugerirCodigoSiguienteStr("ariges" & EmprAVAB & ".olitarifaoferta", "codigo", Insert)
    NumRegElim = Val(SQL)   'Codig para la insercion
    If Not HacerTarifa Then
        If NumRegElim < 100000 Then NumRegElim = 100001
    End If
    Insert = ""
    
    
    
    'Cargo el RS con los valores de los articulos(sartic)
    SQL = "select sartic.codartic,LitrosUnidad,margecom,olitarifaofertalin.* from sartic,olitarifaofertalin where codigo=" & Me.SegundoParametro
    SQL = SQL & " and olitarifaofertalin.codartic=sartic.codartic"
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Comprobacion 1
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & "X"
        miRsAux.MoveNext
    Wend
    miRsAux.MoveFirst
    
'    If Len(SQL) <> C Then
'        MsgBox "No coinciden los datos grabados con el total de articulos.", vbExclamation
'        miRsAux.Close
'        Exit Function
'    End If
    
    
    If Len(SQL) = 0 Then
        MsgBox "Sin datos para la TO.", vbExclamation
        miRsAux.Close
        Exit Function
    End If
    
    'En el rs tengo los datos del articulo
    Set CL = New Collection
    
    CrearLineasTOSAVAB CL, ArticulosInexistentes
    
    CadenaDesdeOtroForm = "insert into ariges" & EmprAVAB & ".olitarifaofertalin (`codigo`,`codartic`,`pivu`,`pivl`,`coste1`,`coste2`,`coste3`,coste4,coste5,`margen`,`pfvu`,`pfvl`) values "
    
    
    If HacerTarifa Then
        Bucle = 1  'TARIFA. Solo lo hara una vez
    Else
        Bucle = lw2.ListItems.Count
    End If
    
    materiaPrimaNoexiste = ""
    For C = 1 To Bucle
        If pb1.visible Then pb1.Value = pb1.Value + 1
        
        
        'Cabecera
        
        SQL = "insert into ariges" & EmprAVAB & ".olitarifaoferta (`codigo`,`codclien`,`fechaini`,`fechafin`,`aceptada`,`tarifa`,`observaciones`) values ("
        SQL = SQL & NumRegElim & ","
        If HacerTarifa Then
            SQL = SQL & "NULL"   'TARIFA. Cliente a NULL
        Else
            SQL = SQL & Val(lw2.ListItems(C).Text)
        End If                                                                              'aceptada
        SQL = SQL & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(txtFecha(3).Text, "F") & ",0,"
        If HacerTarifa Then
            SQL = SQL & cboTarifaAVAB.ItemData(cboTarifaAVAB.ListIndex)
        Else
            SQL = SQL & "NULL"   'TARIFA  a NULL
        End If
        SQL = SQL & "," & DBSet(Me.txtObserva(1).Text, "T")
        SQL = SQL & ")"
        Conn.Execute SQL
        
        
        'lineas
        Insert = ""
        For i = 1 To CL.Count
            Insert = Insert & ", (" & NumRegElim & CL(i)
        Next
        
        Insert = Mid(Insert, 2)
        SQL = CadenaDesdeOtroForm & Insert
        Conn.Execute SQL
        
        'Insertamos las lineas 2
        InsertaEnLineas2 materiaPrimaNoexiste
        
        NumRegElim = NumRegElim + 1
        Me.Refresh
        Espera 0.2
    Next
    
    If materiaPrimaNoexiste <> "" Then
        
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
        SQL = "select codartic,nomartic from sartic where codartic in (" & Mid(materiaPrimaNoexiste, 2) & ")"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            SQL = SQL & miRsAux!codArtic & "   " & miRsAux!NomArtic & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        SQL = "Materias primas no existen en AVAB" & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
    End If
    
    
    
    
    GenerarTO_AVAB = True
    
    
    Exit Function
ETOAVAB:
    MuestraError Err.Number, Err.Description
    Set CL = Nothing
End Function



