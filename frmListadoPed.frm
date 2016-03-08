VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmListadoPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12735
   Icon            =   "frmListadoPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameGenAlbaran 
      Height          =   5415
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   6195
      Begin VB.CheckBox chkAlbValorado 
         Caption         =   "Albarán valorado"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Frame FramepedidoFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         Height          =   615
         Left            =   840
         TabIndex        =   202
         Top             =   4200
         Width           =   4935
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   203
            Text            =   "Text5"
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   315
            MaxLength       =   6
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   0
            Left            =   0
            Picture         =   "frmListadoPed.frx":000C
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cta prevista cobro"
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
            Index           =   23
            Left            =   0
            TabIndex        =   204
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   25
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   16
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkImpAlbaran 
         Caption         =   "Imprimir Albaran"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   15
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   3360
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.CommandButton cmdAceptarGenAlb 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   21
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   32
         Left            =   840
         TabIndex        =   36
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmListadoPed.frx":010E
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Envío"
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
         Index           =   11
         Left            =   840
         TabIndex        =   35
         Top             =   3120
         Width           =   1110
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   840
         Picture         =   "frmListadoPed.frx":0199
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Material Preparado por"
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
         Index           =   10
         Left            =   840
         TabIndex        =   33
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   840
         Picture         =   "frmListadoPed.frx":029B
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a "
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
         Left            =   600
         TabIndex        =   31
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos: "
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
         Left            =   600
         TabIndex        =   30
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador de Albaran"
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
         Index           =   15
         Left            =   840
         TabIndex        =   29
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListadoPed.frx":039D
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FrameEstVentas 
      Height          =   3975
      Left            =   480
      TabIndex        =   185
      Top             =   120
      Width           =   7035
      Begin VB.CheckBox chkConsolidado 
         Caption         =   "Ver todos los almacenes"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   211
         Top             =   2520
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   189
         Top             =   1440
         Width           =   840
      End
      Begin VB.CommandButton cmdAceptarEstVentas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   191
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5280
         TabIndex        =   192
         Top             =   3120
         Width           =   975
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         TabIndex        =   186
         Top             =   1800
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   8
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   187
            Text            =   "Text5"
            Top             =   120
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1020
            MaxLength       =   6
            TabIndex        =   190
            Top             =   120
            Width           =   840
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   705
            Picture         =   "frmListadoPed.frx":049F
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Index           =   30
            Left            =   0
            TabIndex        =   188
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Ventas por meses"
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
         Left            =   480
         TabIndex        =   194
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Index           =   57
         Left            =   480
         TabIndex        =   193
         Top             =   1440
         Width           =   330
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePreFacMante 
      Height          =   6135
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   7275
      Begin VB.Frame Frame2 
         Height          =   1350
         Left            =   360
         TabIndex        =   95
         Top             =   915
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   52
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   114
            Text            =   "Text5"
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   76
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   77
            Top             =   945
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   47
            Left            =   2865
            Locked          =   -1  'True
            TabIndex        =   96
            Text            =   "Text5"
            Top             =   945
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   44
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Prev. Cobro"
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
            Left            =   120
            TabIndex        =   116
            Top             =   600
            Width           =   1350
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   31
            Left            =   1730
            Picture         =   "frmListadoPed.frx":05A1
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Operador"
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
            TabIndex        =   98
            Top             =   945
            Width           =   795
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   26
            Left            =   1730
            Picture         =   "frmListadoPed.frx":06A3
            Top             =   945
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   15
            Left            =   1730
            Picture         =   "frmListadoPed.frx":07A5
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Facturación"
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
            TabIndex        =   97
            Top             =   240
            Width           =   1530
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3120
         Left            =   360
         TabIndex        =   69
         Top             =   2280
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   50
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   91
            Text            =   "Text5"
            Top             =   2280
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   82
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   51
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "Text5"
            Top             =   2640
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   83
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   81
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   49
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "Text5"
            Top             =   1680
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   48
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   80
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   48
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   74
            Text            =   "Text5"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   46
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   79
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   46
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   72
            Text            =   "Text5"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   45
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   78
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   45
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   70
            Text            =   "Text5"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   29
            Left            =   1080
            Picture         =   "frmListadoPed.frx":0830
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Formas de Pago"
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
            Left            =   120
            TabIndex        =   94
            Top             =   2040
            Width           =   1350
         End
         Begin VB.Label Label7 
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
            Height          =   195
            Index           =   10
            Left            =   555
            TabIndex        =   93
            Top             =   2280
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   30
            Left            =   1080
            Picture         =   "frmListadoPed.frx":0932
            Top             =   2640
            Width           =   240
         End
         Begin VB.Label Label7 
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
            Height          =   195
            Index           =   11
            Left            =   555
            TabIndex        =   92
            Top             =   2640
            Width           =   420
         End
         Begin VB.Label Label7 
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
            Height          =   195
            Index           =   9
            Left            =   555
            TabIndex        =   89
            Top             =   1680
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   28
            Left            =   1080
            Picture         =   "frmListadoPed.frx":0A34
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label7 
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
            Height          =   195
            Index           =   8
            Left            =   555
            TabIndex        =   88
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            TabIndex        =   87
            Top             =   1080
            Width           =   585
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   27
            Left            =   1080
            Picture         =   "frmListadoPed.frx":0B36
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mes a facturar"
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
            TabIndex        =   73
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Contrato"
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
            TabIndex        =   71
            Top             =   240
            Width           =   1155
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   25
            Left            =   1380
            Picture         =   "frmListadoPed.frx":0C38
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5400
         TabIndex        =   85
         Top             =   5565
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPreFacMan 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   84
         Top             =   5565
         Width           =   975
      End
      Begin VB.Label lblFactMant 
         Caption         =   "Label5"
         Height          =   375
         Left            =   360
         TabIndex        =   201
         Top             =   5640
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Prefacturación Mantenimientos"
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
         Left            =   360
         TabIndex        =   68
         Top             =   480
         Width           =   6375
      End
   End
   Begin VB.Frame FrameFacturar 
      Height          =   7575
      Left            =   120
      TabIndex        =   65
      Top             =   0
      Width           =   7395
      Begin VB.Frame Frame15 
         Height          =   1215
         Left            =   360
         TabIndex        =   206
         Top             =   5040
         Width           =   6855
         Begin VB.TextBox txtCSB 
            Height          =   285
            Index           =   2
            Left            =   2280
            TabIndex        =   113
            Text            =   "Text1"
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtCSB 
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   112
            Text            =   "Text1"
            Top             =   540
            Width           =   4455
         End
         Begin VB.TextBox txtCSB 
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   111
            Text            =   "Text1"
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb4"
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
            Index           =   26
            Left            =   1200
            TabIndex        =   210
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb3"
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
            Index           =   25
            Left            =   1200
            TabIndex        =   209
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb2"
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
            Index           =   24
            Left            =   1200
            TabIndex        =   208
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tesoreria"
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
            Left            =   240
            TabIndex        =   207
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   360
         TabIndex        =   181
         Top             =   6240
         Visible         =   0   'False
         Width           =   4695
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   345
            Left            =   120
            TabIndex        =   182
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   184
            Top             =   350
            Width           =   4335
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   183
            Top             =   135
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3255
         Left            =   360
         TabIndex        =   122
         Top             =   1800
         Width           =   6855
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   42
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "Text5"
            Top             =   2520
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   109
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   43
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   135
            Text            =   "Text5"
            Top             =   2880
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   110
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   108
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   41
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "Text5"
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   107
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   40
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   130
            Text            =   "Text5"
            Top             =   1680
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   105
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   4980
            MaxLength       =   10
            TabIndex        =   106
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   103
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   4980
            MaxLength       =   10
            TabIndex        =   104
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3180
            MaxLength       =   10
            TabIndex        =   102
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   1920
            Picture         =   "frmListadoPed.frx":0D3A
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            TabIndex        =   139
            Top             =   2520
            Width           =   1005
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
            Height          =   195
            Index           =   48
            Left            =   1395
            TabIndex        =   138
            Top             =   2520
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   23
            Left            =   1920
            Picture         =   "frmListadoPed.frx":0E3C
            Top             =   2880
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
            Height          =   195
            Index           =   49
            Left            =   1395
            TabIndex        =   137
            Top             =   2880
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   1920
            Picture         =   "frmListadoPed.frx":0F3E
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   1920
            Picture         =   "frmListadoPed.frx":1040
            Top             =   2040
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
            Height          =   195
            Index           =   50
            Left            =   1395
            TabIndex        =   134
            Top             =   2040
            Width           =   420
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
            Height          =   195
            Index           =   51
            Left            =   1395
            TabIndex        =   133
            Top             =   1680
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            TabIndex        =   132
            Top             =   1680
            Width           =   585
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
            Height          =   195
            Index           =   37
            Left            =   4200
            TabIndex        =   129
            Top             =   1200
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   2280
            Picture         =   "frmListadoPed.frx":1142
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Albaran"
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
            TabIndex        =   128
            Top             =   1200
            Width           =   1200
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
            Height          =   195
            Index           =   46
            Left            =   1755
            TabIndex        =   127
            Top             =   1200
            Width           =   450
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   4680
            Picture         =   "frmListadoPed.frx":11CD
            Top             =   1215
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
            Height          =   195
            Index           =   36
            Left            =   4200
            TabIndex        =   126
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Albaran"
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
            Left            =   240
            TabIndex        =   125
            Top             =   720
            Width           =   900
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
            Height          =   195
            Index           =   45
            Left            =   1755
            TabIndex        =   124
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidad de la Facturación"
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
            Left            =   240
            TabIndex        =   123
            Top             =   240
            Width           =   2520
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   360
         TabIndex        =   118
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   5160
            MaxLength       =   10
            TabIndex        =   100
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   99
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   0
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   119
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   101
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura"
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
            Left            =   4200
            TabIndex        =   195
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de la Facturación"
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
            Left            =   240
            TabIndex        =   121
            Top             =   240
            Width           =   1980
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   2280
            Picture         =   "frmListadoPed.frx":1258
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Prevista Cobro"
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
            Left            =   240
            TabIndex        =   120
            Top             =   600
            Width           =   1620
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   24
            Left            =   1920
            Picture         =   "frmListadoPed.frx":12E3
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   115
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   6240
         TabIndex        =   117
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Albaranes"
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
         Left            =   360
         TabIndex        =   66
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   360
         TabIndex        =   205
         Top             =   3360
         Width           =   6615
      End
   End
   Begin VB.Frame FramePedxArtic 
      Height          =   5415
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Width           =   7515
      Begin VB.ComboBox cmbTipAlbaran 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListadoPed.frx":13E5
         Left            =   2040
         List            =   "frmListadoPed.frx":13F2
         Style           =   2  'Dropdown List
         TabIndex        =   199
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   480
         TabIndex        =   172
         Top             =   1920
         Width           =   6495
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   240
            TabIndex        =   178
            Top             =   1320
            Width           =   2655
            Begin VB.OptionButton OptResumen 
               Caption         =   "Resumen"
               Height          =   255
               Left            =   1320
               TabIndex        =   180
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OptDetalle 
               Caption         =   "Detalle"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   179
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   174
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   173
            Text            =   "Text5"
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   1
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1425
            Top             =   360
            Width           =   240
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
            Index           =   22
            Left            =   120
            TabIndex        =   177
            Top             =   120
            Width           =   945
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
            Height          =   195
            Index           =   21
            Left            =   480
            TabIndex        =   176
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   2
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1527
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
            Height          =   195
            Index           =   20
            Left            =   480
            TabIndex        =   175
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   480
         TabIndex        =   165
         Top             =   2880
         Width           =   6015
         Begin VB.Frame Frame11 
            Caption         =   " Ordenar por "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1215
            Left            =   120
            TabIndex        =   168
            Top             =   600
            Width           =   2175
            Begin VB.OptionButton OptOrdenVentas 
               Caption         =   "Volumen ventas"
               Height          =   255
               Left            =   120
               TabIndex        =   171
               Top             =   840
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptOrdenNomclien 
               Caption         =   "Nombre cliente"
               Height          =   375
               Left            =   120
               TabIndex        =   170
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton OptOrdenCodclien 
               Caption         =   "Cod. cliente"
               Height          =   255
               Left            =   120
               TabIndex        =   169
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3840
            MaxLength       =   15
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "€"
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
            Index           =   19
            Left            =   5640
            TabIndex        =   167
            Top             =   260
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar Clientes con ventas superiores a"
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
            Index           =   18
            Left            =   120
            TabIndex        =   166
            Top             =   240
            Width           =   3465
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         TabIndex        =   140
         Top             =   2640
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   7
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   21
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   142
            Text            =   "Text5"
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   20
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
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
            Height          =   195
            Index           =   12
            Left            =   480
            TabIndex        =   145
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   12
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1629
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
            Height          =   195
            Index           =   13
            Left            =   480
            TabIndex        =   144
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Index           =   16
            Left            =   120
            TabIndex        =   143
            Top             =   120
            Width           =   585
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   11
            Left            =   1080
            Picture         =   "frmListadoPed.frx":172B
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   480
         TabIndex        =   159
         Top             =   1920
         Width           =   6375
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   13
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   161
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   14
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   160
            Text            =   "Text5"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   4
            Left            =   960
            Picture         =   "frmListadoPed.frx":182D
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
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
            TabIndex        =   164
            Top             =   120
            Width           =   735
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
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   163
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   5
            Left            =   960
            Picture         =   "frmListadoPed.frx":192F
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
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   162
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   153
         Top             =   3120
         Width           =   6975
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   15
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   "Text5"
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   16
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   "Text5"
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   5
            Top             =   840
            Width           =   1215
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   6
            Left            =   960
            Picture         =   "frmListadoPed.frx":1A31
            Top             =   480
            Width           =   240
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
            Index           =   1
            Left            =   120
            TabIndex        =   158
            Top             =   240
            Width           =   660
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
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   157
            Top             =   480
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   7
            Left            =   960
            Picture         =   "frmListadoPed.frx":1B33
            Top             =   840
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
            Height          =   195
            Index           =   9
            Left            =   480
            TabIndex        =   156
            Top             =   840
            Width           =   420
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   12
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   12
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedxArtic 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   11
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo de albaranes:"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   200
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   3840
         Picture         =   "frmListadoPed.frx":1C35
         Top             =   1440
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
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   26
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
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
         Left            =   600
         TabIndex        =   25
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos por Artículo"
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
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   4815
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1440
         Picture         =   "frmListadoPed.frx":1CC0
         Top             =   1440
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
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   23
         Top             =   1440
         Width           =   420
      End
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   5775
      Left            =   240
      TabIndex        =   37
      Top             =   720
      Width           =   7035
      Begin VB.ComboBox cmbTipAlbaran 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListadoPed.frx":1D4B
         Left            =   1920
         List            =   "frmListadoPed.frx":1D58
         Style           =   2  'Dropdown List
         TabIndex        =   198
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkResumenForpa 
         Caption         =   "Resumen forma de pago"
         Height          =   195
         Left            =   3840
         TabIndex        =   196
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox chkSoloFacturar 
         Caption         =   "Solo Albaranes para facturar"
         Height          =   375
         Left            =   3840
         TabIndex        =   50
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tipo Informe"
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
         Height          =   735
         Left            =   480
         TabIndex        =   152
         Top             =   4100
         Width           =   3135
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   49
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   48
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   500
         TabIndex        =   146
         Top             =   2880
         Width           =   6135
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   47
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   33
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   148
            Text            =   "Text5"
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   46
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   32
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "Text5"
            Top             =   360
            Width           =   3615
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
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   151
            Top             =   720
            Width           =   420
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   150
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   19
            Left            =   960
            Picture         =   "frmListadoPed.frx":1D8B
            Top             =   740
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Index           =   38
            Left            =   0
            TabIndex        =   149
            Top             =   120
            Width           =   615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   18
            Left            =   960
            Picture         =   "frmListadoPed.frx":1E8D
            Top             =   380
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   40
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarPreFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   51
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   52
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   27
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   41
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   44
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   31
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   45
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   43
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   42
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo de albaranes:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   197
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
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
         Height          =   195
         Index           =   44
         Left            =   3120
         TabIndex        =   64
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListadoPed.frx":1F8F
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes"
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
         Left            =   360
         TabIndex        =   63
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   43
         Left            =   480
         TabIndex        =   62
         Top             =   1200
         Width           =   1200
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
         Height          =   195
         Index           =   42
         Left            =   920
         TabIndex        =   61
         Top             =   1440
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3600
         Picture         =   "frmListadoPed.frx":201A
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1440
         Picture         =   "frmListadoPed.frx":20A5
         Top             =   3260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formas de Pago"
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
         Index           =   41
         Left            =   480
         TabIndex        =   60
         Top             =   3000
         Width           =   1350
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
         Height          =   195
         Index           =   40
         Left            =   915
         TabIndex        =   59
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1440
         Picture         =   "frmListadoPed.frx":21A7
         Top             =   3620
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
         Height          =   195
         Index           =   39
         Left            =   915
         TabIndex        =   58
         Top             =   3600
         Width           =   420
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
         Height          =   195
         Index           =   35
         Left            =   920
         TabIndex        =   57
         Top             =   2520
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1440
         Picture         =   "frmListadoPed.frx":22A9
         Top             =   2520
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
         Height          =   195
         Index           =   34
         Left            =   920
         TabIndex        =   56
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   33
         Left            =   480
         TabIndex        =   55
         Top             =   1920
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1440
         Picture         =   "frmListadoPed.frx":23AB
         Top             =   2160
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListadoPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)
      
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmMtoCliente As frmFacClientes
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoAlmacen As frmAlmAlPropios
Attribute frmMtoAlmacen.VB_VarHelpID = -1
Private WithEvents frmMtoArticulo As frmAlmArticulos
Attribute frmMtoArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoFPago As frmFacFormasPago
Attribute frmMtoFPago.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
'Private WithEvents frmMtoTipCo As frmManTiposContrato


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String
Private NumParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim IndCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub chkAlbValorado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkImpAlbaran_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkImpAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptarEstVentas_Click()
'Estadistica Ventas por meses
Dim campo As String
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1

    
    'El campo AÑO es obligarotorio
    txtCodigo(53).Text = Trim(txtCodigo(53).Text)
    If txtCodigo(53).Text = "" Then
        MsgBox "Debe seleccionar una año para el informe.", vbInformation
        Exit Sub
    Else
        campo = "year({scafac.fecfactu})"
        cadFormula = campo & " = " & txtCodigo(53).Text
'        campo = campo & " = " & CInt(txtCodigo(53).Text) - 1
'        cadFormula = "(" & cadFormula & " OR " & campo & ")"
        
        'Parametro del año solicitado para el informe
        'Pasar el año solicitado como parametro
        campo = "Año: " & txtCodigo(53).Text
        If vUsu.TrabajadorB And chkConsolidado(0) = 0 Then campo = campo & "     Almacen " & vParamAplic.AlmacenB
        cadParam = cadParam & "pAnyo=""" & campo & """|"
        NumParam = NumParam + 1
    End If
    
    'Campo seleccion de un CLIENTE
    txtCodigo(8).Text = Trim(txtCodigo(8).Text)
    If txtCodigo(8).Text <> "" Then
        campo = "{scafac.codclien}"
        cadFormula = cadFormula & " AND (" & campo & " =" & txtCodigo(8).Text & ")"
        'Pasar el cliente solicitado como parametro
        cadParam = cadParam & "pDHCliente=""" & "Cliente: " & txtCodigo(8).Text & " - " & txtNombre(8).Text & """|"
    Else
        'Mostrar en el informe el total del Año Anterior
        campo = "year({scafac.fecfactu}) = " & CInt(txtCodigo(53).Text) - 1
        cadFormula = "(" & cadFormula & " OR " & campo & ")"
        
        cadParam = cadParam & "pDHCliente=""" & "Cliente: Todos" & """|"
    End If
    NumParam = NumParam + 1
    
        
    
    If vUsu.TrabajadorB Then
        If chkConsolidado(0).Value = 0 Then
            'SOLO QUIERE ver el almacen 2
            'Codigo = Codigo & Space(20) & "Alma*"
            campo = "(codtipom,numfactu,fecfactu) IN (select distinct codtipom,numfactu,fecfactu FROM slifac where codalmac=" & vParamAplic.AlmacenB & ")"
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        End If
    End If
    
    
    
    
    
    'Comprobar si hay registros para mostrar en el informe
    cadSelect = cadFormula
    If Not HayRegParaInforme("scafac", cadSelect) Then Exit Sub
    
    
    'Borro los datos temporales,por si acaso se hubiera quedado
    BorrarTempInformes
    
    'Generar la temporal con los totales por año, mes y cliente (tmpinformes)
    If Not TempVentasMeses_(cadSelect, txtCodigo(53).Text, chkConsolidado(0).Value = 0) Then
        'Borrar los registros generados por el usuario de la temporal
        BorrarTempInformes
        Exit Sub
    End If
    
    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    
    Titulo = "Ventas por meses"
'    If Me.OptTipoInf(0).Value = True Then
        nomRPT = "rFacVentasxMesGra.rpt"
'    Else
'        Exit Sub
'        nomRPT = "rFacVentasxMesTex.rpt"
'    End If
    conSubRPT = False
    
    LlamarImprimir
    
    'Borrar los registros generados por el usuario de la temporal
    BorrarTempInformes
End Sub



Private Sub cmdAceptarFac_Click()
'Facturacion de Albaranes
Dim campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean
    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCodigo(0).Text) = "" Then 'Banco propio
        MsgBox "El campo cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    If OpcionListado <> 222 Then 'Facturas Ventas (FACTURACION)
                                 '222: Facturas de Mostrador/Rectificativa
        'Desde/Hasta Nº ALBARAN
        '-------------------------
        If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
            campo = NomTabla & ".numalbar"
            Cad = ""
            If Not PonerDesdeHasta(campo, "N", 36, 37, Cad) Then Exit Sub
        End If
    
        'Desde/Hasta FECHA del ALBARAN
        '--------------------------------------------
        If txtCodigo(38).Text <> "" Or txtCodigo(39).Text <> "" Then
            'Para MySQL
            campo = "scaalb.fechaalb"
            Cad = CadenaDesdeHastaBD(txtCodigo(38).Text, txtCodigo(39).Text, campo, "F")
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            'Para Crystal Report
            campo = "{scaalb.fechaalb}"
            Cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 38, 39, Cad) Then Exit Sub
        End If
    
        'Cadena para seleccion D/H CLIENTE
        '----------------------------------------
        If txtCodigo(40).Text <> "" Or txtCodigo(41).Text <> "" Then
            campo = "scaalb.codclien"
            Cad = ""
            If Not PonerDesdeHasta(campo, "N", 40, 41, Cad) Then Exit Sub
        End If
    
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(42).Text <> "" Or txtCodigo(43).Text <> "" Then
            campo = "scaalb.codforpa"
            Cad = " "
            If Not PonerDesdeHasta(campo, "N", 42, 43, Cad) Then Exit Sub
        End If

    
        'Otros criterios de Seleccion
        '---------------------------------------------
        'Seleccionar de la Tabla de albaranes scaalb, solo los Albaranes que sean
        'del tipo:Ventas o Reparacion o Mantenimiento
    '    cad = " scaalb.codtipom='ALV' "
        Cad = " scaalb.codtipom='" & CodClien & "' " 'filtrar por tipo de albaran segun llamado de Alb.Ventas o Alb. Reparacion
        'Solo lo añadimos a CadSelect porque vamos a Facturar y no a sacar un listado
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
    
    
        'Seleccionar los Albanares de la Periodicidad indicada
        If txtCodigo(35).Text <> "" Then
            Cad = " sclien.periodof=" & txtCodigo(35).Text
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            cadFrom = " scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
        End If
        
    Else
        'Facturar UNA solo
        If MsgBox("Generar la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        'en la llamada reutilizamos las vbles codclien y NumCod para guardar tipomov y numalbar.
        cadFormula = "{scaalb.codtipom}='" & CodClien & "' AND scaalb.numalbar=" & NumCod
        cadSelect = cadFormula
    End If
    
    cadSQL = cadSelect
''''''''   'Seleccionar los Albaranes que tiene scaalb.factursn=1
''''''''   cad = " {scaalb.factursn=1} "
    
                                                                    'Septiembre 2009
    'Seleccionar los Albaranes que tiene scaalb.factursn=1     y TENGAN lineas
    Cad = " {scaalb.factursn=1} "
    Cad = Cad & " and (scaalb.codtipom,scaalb.numalbar) in (select codtipom,numalbar from slialb group by codtipom,numalbar)"

    
    If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
    AnyadirAFormula cadFormula, Cad
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    Cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " scaalb "
    Cad = Cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    If OpcionListado <> 222 Then
        'Seleccionar los Albaranes que tiene scaalb.factursn=0
        campo = " scaalb.factursn=0 "
        If Not AnyadirAFormula(cadSQL, campo) Then Exit Sub
        cadSQL = Cad & " WHERE " & cadSQL
        If RegistrosAListar(cadSQL) > 0 Then
            'Mostrar los Albaranes que no se van a Facturar
            cadSQL = Replace(cadSQL, "count(*)", "scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,scaalb.codclien,scaalb.nomclien")
            frmMensajes.OpcionMensaje = 12
            frmMensajes.cadWhere = cadSQL
            frmMensajes.Show vbModal
            If frmMensajes.vCampos = "0" Then Exit Sub
        End If
    End If
    
    Cad = Cad & " WHERE " & cadSelect
    'Pasar Albaranes a Facturas
    If InStr(Cad, "sclien") <> 0 Then 'hay JOIN con sclien
        Cad = Replace(Cad, "count(*)", "scaalb.*, sclien.periodof")
    Else
        Cad = Replace(Cad, "count(*)", "*")
    End If



    'Albarananes EN B
    If CodClien = "ALZ" Then
        If Not AbrirConexionConta(True) Then
            Cad = "Error MUY grave." & vbCrLf & "Error conectando con BD: " & vParamAplic.ContabilidadB
            MsgBox Cad, vbCritical
            End
            Exit Sub
        End If
        CambiamosConta = True
    End If



    '--- Mostrar Barra de PRogreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador/Rectificativa
                                 '52: Facturas de Venta
                                 'Facturas Reparacion
        
        Me.Height = Me.Height + 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
        Me.FrameProgress.visible = True
        Me.FrameProgress.Top = 6250
        Me.ProgressBar1.Left = 200
        Me.ProgressBar1.Value = 0
        Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
        
        
        'Si vamos a facturar albaranes "B" tenemos que cerrar la conexion CONTA y abrirla contra la
        'Segunda BD que nos indica
        
    End If
    

    '--- Traspasa Albaranes a Facturas
    If OpcionListado = 222 And Me.EstaRecupFact = True Then
        '#### Laura: 14/11/2006 Recuperar facturas ALZIRA
        'comprobar q se ha introducido el nº de factura
        If Trim(txtCodigo(4).Text) = "" Then
            MsgBox "Debe introducir el nº de factura"
            Exit Sub
        End If
        'comprobar q la factura esta en un rango de recuperacion
        If Not (4255 <= CLng(txtCodigo(4).Text) And CLng(txtCodigo(4).Text) <= 5220) Then
            MsgBox "El Nº de factura no esta en el rango de recuperación."
            Exit Sub
        End If
        'comprobar q no exista ya ese nº de factura en ariges
        campo = "SELECT COUNT(*) FROM scafac WHERE "
        campo = campo & "codtipom='FAV' and numfactu=" & DBSet(txtCodigo(4).Text, "N") & " and year(fecfactu)=" & Year(txtCodigo(34).Text) '" and fecfactu=" & DBSet(txtCodigo(34).Text, "F")
        If Not (RegistrosAListar(campo) > 0) Then
            'comprobar si existe la factura en contabilidad
            campo = ""
            campo = ObtenerLetraSerie("FAV")
            If campo = "" Then Exit Sub
            campo = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(campo, "T")
            campo = campo & " AND codfaccl=" & txtCodigo(4).Text & " AND anofaccl=" & Year(txtCodigo(34).Text)
            
            If Not (RegistrosAListar(campo, conConta) > 0) Then
                'no existe en contabilidad recuperamos la factura y ya esta (no insertamos en tesoreria)
                TraspasoAlbaranesFacturas_RecuperaFac Cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
            Else
                'si esiste
                'comprobar q el cliente es el mismo en la factura q vamos a recuperar
                'y en la factura de la conta
                If Not ComprobarCliente_RecuperarFac(cadSelect, txtCodigo(34).Text, txtCodigo(4).Text) Then Exit Sub
                'si existe en contabilidad recuperamos la factura y marcar como contabilizada
                TraspasoAlbaranesFacturas_RecuperaFac Cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
                
                
            End If
        Else
            MsgBox "Ya existe la Factura en Ariges", vbExclamation
        End If
        '####################
    Else
        'proceso normal
         'Fecha de la factura, Cta Prevista de Cobro
         Screen.MousePointer = vbHourglass
         
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        campo = "Albaran: " & CodClien & " : " & NumCod
        LOG.Insertar 2, vUsu, campo
        Set LOG = Nothing
        '-----------------------------------------------------------------------------

        campo = txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
        TraspasoAlbaranesFacturas Cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo

    End If
    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
       AbrirConexionConta False
    End If
    '--- Ocultar Barra de Progreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador
        Me.Height = Me.Height - 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
        Me.FrameProgress.visible = False
    Else
        'Cierro y salgo
        Unload Me
    End If
End Sub



'#### Laura 14/11/2006 Recuperar facturas ALZIRA
Private Function ComprobarCliente_RecuperarFac(cadSelAlb As String, fecFac As String, numFac As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim codMacta1 As String 'cliente factura ariges
Dim codMacta2 As String 'cliente factura conta
Dim letra As String

    On Error GoTo ErrCompCliente
    ComprobarCliente_RecuperarFac = False
    
    'codmacta del cliente del albaran a facturar en Ariges
    SQL = "select scaalb.codclien,sclien.codmacta"
    SQL = SQL & " from scaalb inner join sclien on scaalb.codclien=sclien.codclien "
    SQL = SQL & " Where " & cadSelAlb
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        codMacta1 = DBLet(RS!Codmacta, "T")
    
    End If
    Set RS = Nothing
    
    
    'codmacta en la contabilidad
    letra = ObtenerLetraSerie("FAV")
    SQL = "SELECT codmacta FROM cabfact "
    SQL = SQL & " WHERE numserie=" & DBSet(letra, "T") & " AND codfaccl=" & numFac & " AND anofaccl=" & Year(fecFac)
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        codMacta2 = DBLet(RS!Codmacta, "T")
    End If
    Set RS = Nothing
    
    If codMacta1 <> "" And codMacta2 <> "" Then
        If codMacta1 = codMacta2 Then
            ComprobarCliente_RecuperarFac = True
        Else
            ComprobarCliente_RecuperarFac = False
            MsgBox "La cuenta contable en la factura de Contabilidad no coincide con la del cliente del Albaran", vbExclamation
        End If
    Else
        ComprobarCliente_RecuperarFac = False
        MsgBox "No se ha podido leer la cuenta contable del cliente", vbExclamation
    End If
    
    Exit Function
    
ErrCompCliente:
    ComprobarCliente_RecuperarFac = False
    MuestraError Err.Number, "Comprobar cliente", Err.Description
End Function
'#####


Private Sub cmdAceptarGenAlb_Click()
'Solicitar datos para Generar Albaran a partir de un Pedido
Dim Cad As String

    'DAVID
    'Comprobar que me han puesto algun dato
    '-------------------------------------------------------------------
    Cad = ""
    If txtCodigo(17).Text = "" Or txtCodigo(18).Text = "" Or txtCodigo(19).Text = "" Or txtCodigo(25).Text = "" Then Cad = "M"
    If OpcionListado = 1000 Then
        If txtCodigo(5).Text = "" Then Cad = "M"
        If txtNombre(5).Text = "" Then Cad = "M"
    End If
    If txtNombre(17).Text = "" Or txtNombre(18).Text = "" Or txtNombre(19).Text = "" Then Cad = "M"
    
    If Cad <> "" Then
        MsgBox "Campos obligatorios ", vbExclamation
        Exit Sub
    End If
    
    
    
    Cad = txtCodigo(17).Text & "|"
    Cad = Cad & txtCodigo(18).Text & "|"
    Cad = Cad & txtCodigo(19).Text & "|"
    Cad = Cad & txtCodigo(25).Text & "|"
    Cad = Cad & Me.chkImpAlbaran.Value & "|"
    'mando el banco propio
    If OpcionListado = 1000 Then
        Cad = Cad & txtCodigo(5).Text & "||"                'DOS CAMPOS
    Else
        'Envio tb el albaran si sale valorado o NO
        Cad = Cad & "|" & chkAlbValorado.Value & "|"      'DOS CAMPOS
    End If
    
    
    
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub



Private Sub cmdAceptarPedxArtic_Click()
'41: Informe de Pedidos por Articulo
'44: Informe de Pedidos por Cliente
'49: Informe de Albaranes por Artículo
Dim campo As String
Dim Cad As String
Dim SQL As String
Dim cadFormula2 As String
Dim cadSelect2 As String
Dim cadSelect3 As String
Dim Indice As Integer


    InicializarVbles
    cadFormula2 = ""
    cadSelect2 = ""
    cadSelect3 = ""
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1


    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Fechas de Pedido/Albaran/Factura
    '--------------------------------------------
    'Desde/Hasta FECHA
    'para el informe 227 fecha requerida
    If OpcionListado = 227 Then
        If txtCodigo(11).Text = "" Or txtCodigo(12).Text = "" Then
            MsgBox "Los campos D/H fecha factura deben tener valor.", vbInformation
            Exit Sub
        End If
        
        If DateDiff("d", txtCodigo(11).Text, txtCodigo(12).Text) > 365 Then
            MsgBox "El intervalo de fechas no puede ser superior a un año.", vbInformation
            Exit Sub
        End If
    End If
    
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        If OpcionListado = 227 Or OpcionListado = 228 Then
            campo = "{" & NomTabla & ".fecfactu}"
        ElseIf OpcionListado = 49 Then
            campo = "{" & NomTabla & ".fechaalb}"
        Else
            campo = "{" & NomTabla & ".fecpedcl}"
        End If
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 11, 12, Cad) Then Exit Sub
        cadSelect = CadenaDesdeHastaBD(txtCodigo(11).Text, txtCodigo(12).Text, campo, "F")
        
        'Guardamos el periodo para calcular las ventas
        If OpcionListado = 227 Then
            cadFormula2 = cadFormula
            cadSelect2 = cadSelect
            'obtenemos el periodo anterior de ventas
            Cad = "": SQL = ""
            If txtCodigo(11).Text <> "" Then Cad = Day(txtCodigo(11).Text) & "/" & Month(txtCodigo(11).Text) & "/" & Year(txtCodigo(11).Text) - 1
            If txtCodigo(12).Text <> "" Then SQL = Day(txtCodigo(12).Text) & "/" & Month(txtCodigo(12).Text) & "/" & Year(txtCodigo(12).Text) - 1
            cadSelect3 = CadenaDesdeHastaBD(Cad, SQL, campo, "F")
        
        ElseIf OpcionListado = 41 Or OpcionListado = 42 Then '42:Disponibilidad Stock
        'pasar D/H fecha como parametro para enlazar con la cabecera de pedidos proveedor
        'que esta como subinforme y que seleccione el mismo rango de fecha que
        'para la cabecera de pedidos de cliente
            If txtCodigo(11).Text <> "" Then
                Cad = "pFechaD=" & "Date(" & Year(txtCodigo(11).Text) & ", " & Month(txtCodigo(11).Text) & ", " & Day(txtCodigo(11).Text) & ")"
            Else
                Cad = "pFechaD=" & "Date(1900,01,01)"
            End If
            cadParam = cadParam & Cad & "|"
            NumParam = NumParam + 1
            If txtCodigo(12).Text <> "" Then
                Cad = "pFechaH=" & "Date(" & Year(txtCodigo(12).Text) & ", " & Month(txtCodigo(12).Text) & ", " & Day(txtCodigo(12).Text) & ")"
            Else
                Cad = "pFechaH=" & "Date(9999,12,31)"
            End If
            cadParam = cadParam & Cad & "|"
            NumParam = NumParam + 1
        End If
    End If
    
    'Cadena para seleccion ALMACEN
    '--------------------------------------------
    If Me.Frame9.visible Then
        If txtCodigo(13).Text <> "" Or txtCodigo(14).Text <> "" Then
            campo = "{" & NomTablaLin & ".codalmac}"
            'Parametro Desde/Hasta Almacen
            Cad = "pDHAlmacen=""Almacen: "
            If Not PonerDesdeHasta(campo, "N", 13, 14, Cad) Then Exit Sub
        End If
    End If
    
    
    'Cadena para seleccion ARTICULO
    '--------------------------------------------
    If Me.Frame8.visible Then
        If txtCodigo(15).Text <> "" Or txtCodigo(16).Text <> "" Then
            campo = "{" & NomTablaLin & ".codartic}"
            'Parametro Desde/Hasta Articulo
            Cad = "pDHArticulo=""Artículo: "
             If Not PonerDesdeHasta(campo, "T", 15, 16, Cad) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If Me.Frame5.visible Then
        If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
            campo = "{" & NomTabla & ".codclien}"
            'Parametro Desde/Hasta Cliente
            Cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 20, 21, Cad) Then Exit Sub
        End If
    End If
    
    
    'Cadena para seleccion TRABAJADOR
    '--------------------------------------------
    If Me.Frame12.visible Then
        If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
            campo = "{scafac1.codtraba}"
            'Parametro Desde/Hasta Trabajador
            Cad = "pDHTrabajador=""Trabajador: "
            If Not PonerDesdeHasta(campo, "N", 2, 3, Cad) Then Exit Sub
        End If
    End If
    
    
    
    '227: Listado Ventas por cliente
    'Importe ventas superior a ....
    If Me.Frame10.visible Then
        Cad = DBSet(txtCodigo(1).Text, "N")
        cadParam = cadParam & "pImporte=" & Cad & "|"
        NumParam = NumParam + 1
            
        If txtCodigo(1).Text <> "" Then
            'seleccionar solo los clientes que el total de la BaseImp supere esa cantidad
            If cadSelect <> "" Then SQL = cadSelect2 & " AND "
            Cad = ObtenerClientes(cadSelect, Cad)
            cadSelect = SQL & Cad
'            If cadSelect3 <> "" Then cadSelect3 = cadSelect3 & " AND "
'            cadSelect3 = cadSelect3 & cad
            If cadFormula2 <> "" Then cadFormula2 = cadFormula2 & " AND "
            cadFormula = cadFormula2 & Cad
        End If
    End If
    
    
    If OpcionListado = 49 Then
        campo = ".numalbar"
'        cad = "{" & NomTabla & ".codtipom}='ALV'"
'        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
'        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        '-- Ahora en este informe hay mas posibilidades de selección [SERVICIOS]
        If vParamAplic.Servicios Then
            Indice = cmbTipAlbaran(1).ListIndex
            If Indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case Indice
                    Case 0 ' solo ventas
                        Cad = "{" & NomTabla & ".codtipom}='ALV'"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas)"
                    Case 1 ' solo servicios
                        Cad = "{" & NomTabla & ".codtipom}='ALS'"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Servicios)"
                    Case 2 ' ventas y servicios
                        Cad = " ({" & NomTabla & ".codtipom}='ALV'" & _
                                " OR {" & NomTabla & ".codtipom}='ALS')"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas y servicios)"
                End Select
            End If
        Else
            Cad = "{" & NomTabla & ".codtipom}='ALV'"
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            Titulo = "Albaranes por artículo (Ventas)"
        End If
        'Pasar nombre el título del informe
        cadParam = cadParam & "|pTitulo=""" & Titulo & """|"
        NumParam = NumParam + 1
    Else
        campo = ".numpedcl"
    End If

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 227 Then
        Cad = NomTabla
        Titulo = "Ventas por Cliente"
        nomRPT = "rFacVentasxClien.rpt"
        conSubRPT = False
    ElseIf OpcionListado = 228 Then
        Cad = NomTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.fecfactu=scafac1.fecfactu and scafac.numfactu=scafac1.numfactu"
        Titulo = "Ventas por Trabajador"
        If Me.OptDetalle(2).Value = True Then 'Inf. Detalle
            nomRPT = "rFacVentasxTrabaDet.rpt"
            conSubRPT = True
        ElseIf Me.OptResumen.Value = True Then 'Inf. Resum
            nomRPT = "rFacVentasxTrabaRes.rpt"
            conSubRPT = False
        End If
    Else
       Cad = NomTabla & " INNER JOIN " & NomTablaLin
       Cad = Cad & " ON " & NomTabla & campo & "=" & NomTablaLin & campo
       If OpcionListado = 49 Then _
       Cad = Cad & " AND " & NomTabla & ".codtipom=" & NomTablaLin & ".codtipom "
       
    End If
    If OpcionListado = 44 Then conSubRPT = True
    If Not HayRegParaInforme(Cad, cadSelect) Then Exit Sub
    
    
    If OpcionListado = 227 Then
        BorrarTempInformes
        
        'Pasar los datos a la tabla temporal tmpInformes y luego mostrar
        'el informe de esta tabla
        cadSelect2 = Replace(cadSelect2, "{", "")
        cadSelect2 = Replace(cadSelect2, "}", "")
        
        cadSelect3 = Replace(cadSelect3, "{", "")
        cadSelect3 = Replace(cadSelect3, "}", "")
        If Not TempVentasClientes(cadSelect, cadSelect2, cadSelect3) Then Exit Sub
        
        'Añadir como parametros el total del periodo que devuelve en cadSelect2
        'y añadir el parametro del total periodo anterior q devuelve en cadSelect3
        cadParam = cadParam & "pTotal=" & cadSelect2 & "|"
        NumParam = NumParam + 1
        cadParam = cadParam & "pTotalAnt=" & cadSelect3 & "|"
        NumParam = NumParam + 1
        
        
        'Añadir el parametro para el orden del informe
        'Orden del Informe
        If Me.OptOrdenCodclien.Value Then
            Cad = "{tmpinformes.codigo1}"
            SQL = "Orden: Cod. cliente"
        ElseIf Me.OptOrdenNomclien.Value Then
            Cad = "{tmpinformes.nombre1}"
            SQL = "Orden: Nombre cliente"
        ElseIf Me.OptOrdenVentas.Value Then
            Cad = "{tmpinformes.importe5}"
            SQL = "Orden: Volumen ventas"
        End If
        cadParam = cadParam & "pOrden=" & Cad & "|"
        NumParam = NumParam + 1
        cadParam = cadParam & "pCadOrden=""" & SQL & """|"
        NumParam = NumParam + 1
        
        
        'no le pasamos formula de seleccion porque los datos ya estan en la temporal
        'solo el usuario que genero la informacion en la temporal
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
    End If
    
    
    LlamarImprimir
End Sub


Private Sub cmdAceptarPreFac_Click()
'Prevision de Facturacion de Albaranes
Dim campo As String, Cad As String
Dim B As Boolean
Dim Indice As Integer

    InicializarVbles
    B = (OpcionListado = 50)
    
    If (Not B) Or (B And CodClien = "ALV") Then
        'Pasar nombre de la Empresa como parametro
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        NumParam = NumParam + 1
    End If
    
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    If Trim(txtCodigo(26).Text) <> "" Or Trim(txtCodigo(27).Text) <> "" Then
        If B And CodClien <> "ALV" Then
            campo = "scaalb.fechaalb"
            Cad = "FECHA: "
            cadFormula = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
            cadParam = cadParam & AnyadirParametroDH(Cad, 26, 27) & """|"
        Else
            'Para MySQL
            campo = "scaalb.fechaalb"
            cadSelect = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
            'Para Crystal Report
            campo = "{scaalb.fechaalb}"
            Cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 26, 27, Cad) Then Exit Sub
        End If
    End If

    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
        If B And CodClien <> "ALV" Then
            campo = "scaalb.codclien"
            Cad = "CLIENTE: "
        Else
            campo = "{scaalb.codclien}"
            Cad = "pDHCliente=""Cliente: "
        End If
        If Not PonerDesdeHasta(campo, "N", 28, 29, Cad) Then Exit Sub
    End If

    If B Then 'opcionlistado=50
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(30).Text <> "" Or txtCodigo(31).Text <> "" Then
            If B And CodClien <> "ALV" Then
                campo = "scaalb.codforpa"
                Cad = "F. PAGO: "
            Else
                campo = "{scaalb.codforpa}"
                Cad = "pDHForpa=""Forma Pago: "
            End If
            If Not PonerDesdeHasta(campo, "N", 30, 31, Cad) Then Exit Sub
        End If
        
        'seleccionar los Albaranes de Venta/Repar/Mantenim.
        'seleccionamos tipo de movimiento segun albaran de venta o Reparacion (ALV,ALR)
        '-- Aqui es donde se seleccionaban los albaranes a mostrar en el informe, ahora
        '   como se pueden seleccionar diferentes combinaciones se modifica la carga de la
        '   selección (se queda en rem la antigua línea) [SERVICIOS]
        If vParamAplic.Servicios Then
            Indice = cmbTipAlbaran(0).ListIndex
            If Indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case Indice
                    Case 0 ' solo ventas
                        Cad = " {scaalb.codtipom}='ALV' "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                    Case 1 ' solo servicios
                        Cad = " {scaalb.codtipom}='ALS' "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                    Case 2 ' ventas y servicios
                        Cad = " ({scaalb.codtipom}='ALV'" & _
                                " OR {scaalb.codtipom}='ALS') "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                End Select
            End If
        Else
            Cad = " {scaalb.codtipom}='" & CodClien & "' "
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
        End If
        'Seleccionar los que esten marcados para facturar
        'Seleccionar solo aquellos que el campo scaalb.factursn=1
        If Me.chkSoloFacturar.Value = 1 Then
            Cad = " {scaalb.factursn}=1 "
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
        End If
    Else
        'Cadena para seleccion AGENTE
        '--------------------------------------------
        If txtCodigo(32).Text <> "" Or txtCodigo(33).Text <> "" Then
            campo = "{scaalb.codagent}"
            Cad = "pDHAgente="""
            If Not PonerDesdeHasta(campo, "N", 32, 33, Cad) Then Exit Sub
        End If
        
        'Seleccionar solo aquellos que tienen Nº de Pedido para comparar los Plazos de Entrega
        campo = " NOT ISNULL({scaalb.numpedcl}) "
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    
    If OpcionListado = 51 Then
        Titulo = "Incumplimiento Plazos de Entrega"
        nomRPT = "rFacIncumPlazos.rpt"
        
    ElseIf OpcionListado = 50 And CodClien = "ALV" Then
    
        If chkResumenForpa.Value = 1 Then
            'VAMOS A MOSTRAR LA HOJA RESUMEN DE FORMAS DE PAGO
            Conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
        
            If Me.OptDetalle(0).Value Then
                Titulo = "SELECT scaalb.codforpa, sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   ((scaalb scaalb INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien) INNER JOIN slialb slialb ON (scaalb.codtipom=slialb.codtipom) AND (scaalb.numalbar=slialb.numalbar)) INNER JOIN starif starif ON sclien.codtarif=starif.codlista"
            
            Else
                Titulo = "SELECT  codforpa ,sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   slialb slialb INNER JOIN scaalb scaalb ON (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
            End If
    
            If cadSelect <> "" Then Titulo = Titulo & " WHERE " & cadSelect
            
            Titulo = Titulo & " GROUP BY codforpa"
            Titulo = "INSERT INTO tmpinformes (codigo1,importe1,codusu) " & Titulo
            Conn.Execute Titulo
        End If
    
    
        Titulo = "Previsión Facturación Ventas"
        '-- Si estan activos los servicios hay diferentes posibilidades y el título
        '   las refleja, la variabele 'indice' lleva la información del combo seleccionado y
        '   ha sido cargada un poco más arriba [SERVICIOS]
        If vParamAplic.Servicios Then
            Select Case Indice
                Case 0
                    Titulo = "Previsión Facturación Ventas"
                Case 1
                    Titulo = "Previsión Facturación Servicios"
                Case 2
                    Titulo = "Previsión Facturación Ventas y Servicios"
            End Select
        End If
        conSubRPT = True
        If Me.OptDetalle(0).Value = True Then
            nomRPT = "rFacPrevFactDetalle.rpt"
        Else
            nomRPT = "rFacPrevFactResum.rpt"
        End If
        
        Cad = "pCodUsu=" & vUsu.Codigo & "|"
        cadParam = cadParam & Cad
        NumParam = NumParam + 1
        
        '-- Ahora el título depende de los tipos de albaranes seleccionados [SERVICIOS]
        Cad = "pTitulo=""" & Titulo & """|"
        cadParam = cadParam & Cad
        NumParam = NumParam + 1
        
        
        '--  Mostrara , o no, el subreport con el resumen por forma pago
        Cad = "pVerForpa=" & Abs(chkResumenForpa.Value) & "|"
        cadParam = cadParam & Cad
        NumParam = NumParam + 1
        
        
        On Error GoTo EPreFact
        Cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        Conn.Execute Cad
        
        'Insertar total bonificaciones por cliente,articulo en una TEMPORAL
        Cad = "SELECT " & vUsu.Codigo & " as codusu,  slialb.codartic,scaalb.codclien,sum(slialb.cantidad) as stock "
        Cad = Cad & "FROM (((scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
        Cad = Cad & " INNER JOIN sbonif ON slialb.codartic=sbonif.codartic ) "
        Cad = Cad & " INNER JOIN sclien ON scaalb.codclien=sclien.codclien) "
        Cad = Cad & " INNER JOIN starif ON sclien.codtarif=starif.codlista "
        Cad = Cad & "WHERE " & cadSelect
        Cad = Cad & " AND starif.bonifica=1 "
        Cad = Cad & " GROUP BY scaalb.codclien,slialb.codartic"
        
        Cad = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock) " & Cad
        Conn.Execute Cad
    End If
    
    If B And CodClien <> "ALV" Then 'OpcionListado = 50 'NO Imprime, mostrar resultado en pantalla
        frmMensajes.cadWhere = cadSelect
        frmMensajes.vCampos = cadParam
        frmMensajes.OpcionMensaje = 6 'Prefacturacion Albaranes
        frmMensajes.Show vbModal
    Else
        LlamarImprimir
    End If
    
    If OpcionListado = 50 And CodClien = "ALV" Then
        Cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        Conn.Execute Cad
    End If
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub


Private Sub cmdAceptarPreFacMan_Click()
'74: PreFacturar Mantenimientos
'75: Facturar Mantenimientos
Dim campo As String, Cad As String
Dim B As Boolean
Dim PreguntaHecha As Boolean

     InicializarVbles
     B = (OpcionListado = 74) 'Prefacturar (mostrar listado)
     
     'Introducir el mes que se va a facturar
     If txtCodigo(46).Text = "" Then
        MsgBox "Debe introducir el mes a Facturar.", vbInformation
        Exit Sub
    End If
     
    If Not B Then 'Vamos a facturar
        'si vamos a facturar comprobar que la fecha de factura tiene valor
        If txtCodigo(44).Text = "" Then
            MsgBox "El campo Fecha Factura debe tener valor.", vbInformation
            Exit Sub
        End If
        
        'si vamos a facturar debe haber una cta prev. de cobro
        If txtCodigo(52).Text = "" Then
            MsgBox "El campo Cta. Prev. de cobro debe tener valor.", vbInformation
            Exit Sub
        End If
        
        'si vamos a facturar comprobar que el cod. de operador tiene valor
        If txtCodigo(47).Text = "" Then
            MsgBox "El campo operador debe tener valor.", vbInformation
            Exit Sub
        End If
    End If
     
     
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1

     
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Then
        campo = "{scaman.codtipco}"
'        If Not PonerDesdeHasta(campo, "N", 48, 49, cad) Then Exit Sub
        Cad = campo & "= '" & txtCodigo(45).Text & "'"
        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
        
        'Parametro
        Cad = "pTipCo=""Tipo Contrato: "
        cadParam = cadParam & Cad & txtCodigo(45).Text & " - " & txtNombre(45).Text & """|"
        NumParam = NumParam + 1
    End If
     
     
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(48).Text <> "" Or txtCodigo(49).Text <> "" Then
        campo = "{scaman.codclien}"
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 48, 49, Cad) Then Exit Sub
    End If
    
    
    'Cadena para seleccion FORMA PAGO
    '--------------------------------------------
    If txtCodigo(50).Text <> "" Or txtCodigo(51).Text <> "" Then
        campo = "{scaman.codforpa}"
        Cad = "pDHForpa=""Forma Pago: "
        If Not PonerDesdeHasta(campo, "N", 50, 51, Cad) Then Exit Sub
    End If
        
    'MES A FACTURAR
    'Seleccionar solo aquellos que el campo del mes seleccionado sea no nulo
    '------------------------------------------------------------------------
    Cad = Format(txtCodigo(46).Text, "00")
    campo = "mes" & Cad & "act"
    Cad = "(NOT ISNULL({scaman." & campo & "})) and ({scaman." & campo & "}<>0)"
    If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
    'Parametro
    Cad = "pCampoMes={scaman." & campo & "}" & "|"
    cadParam = cadParam & Cad
    NumParam = NumParam + 1
    Cad = "pMes=""Mes a Facturar: " & UCase(txtNombre(46).Text) & """|"
    cadParam = cadParam & Cad
    NumParam = NumParam + 1
    
        
    cadSelect = cadFormula
    If Not HayRegParaInforme("scaman", cadSelect) Then Exit Sub
    
    
    'Aqui deberiamos comporbar si el periodo indicado YA esta facturado o no
    PreguntaHecha = False
    If Not B Then
        'FACTURACION
        Cad = CStr(EsFechaOKConta(CDate(txtCodigo(44).Text)))
        If Val(Cad) > 0 Then
            Cad = "Fecha factura incorrecta para la contabilidad. ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            PreguntaHecha = True
        End If
    End If
    
    
    'Comprobamos si hay matenimientos ya facturados
    Cad = "Select * from " & cadFormula
    If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
    Cad = "SELECT scaman.codclien,nomclien"
    Cad = Cad & " FROM scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien "
    Cad = Cad & " WHERE " & cadSelect
    'Que el ultimo mes de facturado sea mayor o igual  al que voy a facturar
    Cad = Cad & " AND ulmesfac >= " & txtCodigo(46).Text
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & "    .- " & miRsAux!nomClien & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Cad <> "" Then
        Cad = "Los siguientes mantenimientos ya estan facturados: " & vbCrLf & Cad & vbCrLf & vbCrLf
        Cad = Cad & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        PreguntaHecha = True
    End If
    If B Then 'OpcionListado = 74 'NO Imprime, mostrar resultado en pantalla
        Titulo = "Prefacturación Mantenimientos"
        nomRPT = "rManPrefacturar.rpt"
        LlamarImprimir
    Else
    
        
        If Not PreguntaHecha Then
            Cad = "¿Seguro que desea seguir con el proceso?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
            
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        Cad = "MANTENIMIENTOS: " & vbCrLf & cadSelect
        LOG.Insertar 2, vUsu, Cad
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        'Generar facturas de los mantenimientos seleccionados para facturar
        'cada mantenimiento genera una factura
        Cad = "SELECT scaman.codclien,scaman.coddirec,sdirec.nomdirec,nummante,fechaini,codtipco,codforpa,tipopago," & campo & " as importe "
        'David
        'Necesito el campo concefaccl y el tipopago(mensual...)
        Cad = Cad & ", concefac"
        Cad = Cad & " FROM scaman LEFT OUTER JOIN sdirec ON scaman.codclien=sdirec.codclien AND scaman.coddirec=sdirec.coddirec "
        Cad = Cad & " WHERE " & cadSelect
        
        lblFactMant.Caption = "Obteniendo datos"
        lblFactMant.Refresh
        'Pasamos la SQL que selecciona los mantenimientos a facturar y
        'le pasamos la fecha y operador de la factura.
        If TraspasoMtosAFacturas(Cad, cadSelect, txtCodigo(44).Text, txtCodigo(47).Text, txtCodigo(52).Text, txtCodigo(46).Text, lblFactMant) Then 'Fecha de la factura, Operador
            Unload Me
        End If
        lblFactMant.Caption = ""
    End If
End Sub



Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
     
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 41, 42, 44, 49, 227, 228 '41: Informe de Pedidos por Articulo
                        '42: Informe de Disponibilidad de Stocks
                        '44: Informe de Pedidos por Cliente
                        '49: Informe de Albaranes por Articulo
                        '227: Inf. estadistica Ventas por cliente
                PonerFoco txtCodigo(11)
            Case 43, 1000
                    '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                    '1000: Pedido a factura:  Piede ademas de los datos del albaran, la cta prevista
                    If txtCodigo(17).Text <> "" Then
                        PonerFoco txtCodigo(18)
                    Else
                        PonerFoco txtCodigo(17)
                    End If
            Case 50, 51 '50: Prevision de Facturacion Albaranes (NO IMPRIME LISTADO)
                        '51: Inf. Incumplimiento Plazos de Entrega
                PonerFoco txtCodigo(26)
            Case 52, 222  '52: Facturacion de Albaranes
                         '222: Factura de Mostrador
                PonerFoco txtCodigo(34)
            Case 74 '74: Previsión facturación Mantenimientos
                PonerFoco txtCodigo(45)
            Case 75 '75: Facturacion de Mantenimientos
                PonerFoco txtCodigo(44)
            Case 229 '229: Inf. estadistica ventas por meses
                PonerFoco txtCodigo(53)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmppal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FramePedxArtic.visible = False
    Me.FrameGenAlbaran.visible = False
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    Me.FramePreFacMante.visible = False
    Me.FrameEstVentas.visible = False
    
    CommitConexion
    
    NomTabla = "scaped"
    NomTablaLin = "sliped"
        
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
            
        Case 41, 42, 44, 49, 227, 228 '41: Informe de Pedidos por Articulo
                    '42: Informe de Disponibilidad de Stocks
                    '44: Informe de Pedidos por Cliente
                    '49: Informe de Albaranes por Articulo
                    '227: Inf. estadistica Ventas por cliente
            PonerFramePedxArticVisible True, H, W
            indFrame = 2 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            If vParamAplic.Servicios Then
                lblTipAlbaran(1).visible = True
                cmbTipAlbaran(1).visible = True
            Else
                lblTipAlbaran(1).visible = False
                cmbTipAlbaran(1).visible = False
            End If
            
            If OpcionListado = 49 Then 'Albaranes de Venta
                NomTabla = "scaalb"
                NomTablaLin = "slialb"
            ElseIf OpcionListado = 227 Or OpcionListado = 228 Then
                NomTabla = "scafac"
                NomTablaLin = "slifac"
                
                'poner por defecto las fechas del ejercicio contable
                Me.txtCodigo(11).Text = vEmpresa.FechaIni
                Me.txtCodigo(12).Text = vEmpresa.FechaFin
            End If
            
        Case 43, 1000
                '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                '1000:  Pedido a factura: pide la cta prevista de cobro
            
            W = 6515
            H = 5415
            
            PonerFrameVisible Me.FrameGenAlbaran, True, H, W
            txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
            indFrame = 3
            chkImpAlbaran.Caption = "Impimir "
            If OpcionListado = 1000 Then
                 Label4(32).Caption = "Fec. FACTURA"
                  Label3.Caption = "FACTURAR pedido"
                  chkImpAlbaran.Caption = chkImpAlbaran.Caption & "FACTURA"
            Else
                Label4(32).Caption = "Fecha albarán"
                chkImpAlbaran.Caption = chkImpAlbaran.Caption & "albaran"
                If NumCod = "REP" Then
                    Label3.Caption = "Pasar Reparación a Albaran"
                Else
                    Label3.Caption = "Pasar Pedido a Albaran"
                End If
            End If
            FramepedidoFactura.visible = OpcionListado = 1000
            chkAlbValorado.visible = OpcionListado <> 1000
            'Poner el trabajador conectado
            Me.txtCodigo(17).Text = PonerTrabajadorConectado(cadParam)
            Me.txtNombre(17).Text = cadParam
            cadParam = ""
            
        Case 50, 51 '50: Prevision Facturacion de Albaranes (NO IMPRIME LISTADO)
                    '51: Inf. Incumplimiento Plazos de Entrega
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            If vParamAplic.Servicios Then
                lblTipAlbaran(0).visible = True
                cmbTipAlbaran(0).visible = True
                lblTipAlbaran(0).Top = cmdAceptarPreFac.Top
                cmbTipAlbaran(0).Top = cmdAceptarPreFac.Top
            Else
                lblTipAlbaran(0).visible = False
                cmbTipAlbaran(0).visible = False
            End If
            chkResumenForpa.visible = OpcionListado = 50
        Case 52, 222
                    '52: Facturacion de Albaranes
                    '222: Factura de Mostrador
                    
            PonerFrameFacVisible True, H, W
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            NomTabla = "scaalb"
            NomTablaLin = "slialb"
            
            'Si es facturacion directa: 222 oculto el frame y muestro el albaran que voy a facturar
            Frame4.visible = (OpcionListado = 52)
            If OpcionListado = 52 Then
                Label10(10).Caption = ""
                Me.Frame15.Top = 5040
            Else
                Label10(10).Caption = "Albarán:     " & CodClien & "   " & NumCod
                Me.Frame15.Top = 1800
            End If
            
        Case 74, 75 '74: Prefacturación Mantenimientos
                    '75: Facturacion de Mantenimientos
            lblFactMant.Caption = ""
            PonerFramePreFacManteVisible True, H, W
            indFrame = 7 'solo para el boton cancelar
            
        Case 229 '229: Inf. estadistica ventas por mes
            Me.chkConsolidado(0).visible = vUsu.TrabajadorB
            H = 4000
            W = 7035
            PonerFrameVisible Me.FrameEstVentas, True, H, W
            indFrame = 8
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    If indFrame > 0 Then Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agente
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlmacen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Almacen
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFPago_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTipCo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Contrato del Mantenimiento
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 11, 12, 14, 15, 20, 21, 27, 28, 32 'Cod. CLIENTE
            Select Case Index
                Case 11, 12: IndCodigo = Index + 9
                Case 14, 15: IndCodigo = Index + 14
                Case 20, 21: IndCodigo = Index + 20
                Case 27, 28: IndCodigo = Index + 21
                Case 32: IndCodigo = 8
            End Select
            Set frmMtoCliente = New frmFacClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(IndCodigo).Text) Then txtCodigo(IndCodigo).Text = ""
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
            
        Case 4, 5 'Cod. ALMACEN
            If Index = 4 Then IndCodigo = 13
            If Index = 5 Then IndCodigo = 14
            Set frmMtoAlmacen = New frmAlmAlPropios
            frmMtoAlmacen.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(IndCodigo).Text) Then txtCodigo(IndCodigo).Text = ""
            frmMtoAlmacen.Show vbModal
            Set frmMtoAlmacen = Nothing
            
        Case 6, 7 'Cod. ARTICULO
            If Index = 6 Then
                IndCodigo = 15
            Else
                IndCodigo = 16
            End If
            Set frmMtoArticulo = New frmAlmArticulos
            frmMtoArticulo.DatosADevolverBusqueda2 = "@1@"
            frmMtoArticulo.Show vbModal
            Set frmMtoArticulo = Nothing
        
        Case 1, 2, 8, 9 'cod. TRABAJADOR
            Select Case Index
                Case 1, 2: IndCodigo = Index + 1
                Case 8, 9: IndCodigo = Index + 9
            End Select
            If Index = 8 And txtCodigo(17).Text <> "" Then Exit Sub
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 10 'Cod. Forma de Envio
            IndCodigo = 19
            Set frmMtoFEnvio = New frmFacFormasEnvio
            frmMtoFEnvio.DatosADevolverBusqueda = "0|1|"
            frmMtoFEnvio.Show vbModal
            Set frmMtoFEnvio = Nothing
            
        Case 16, 17, 22, 23, 29, 30 'Forma de PAGO
            Select Case Index
                Case 16, 17: IndCodigo = Index + 14
                Case 22, 23: IndCodigo = Index + 20
                Case 29, 30: IndCodigo = Index + 21
            End Select
            Set frmMtoFPago = New frmFacFormasPago
            frmMtoFPago.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(IndCodigo).Text) Then txtCodigo(IndCodigo).Text = ""
            frmMtoFPago.Show vbModal
            Set frmMtoFPago = Nothing
            
        Case 18, 19 'cod AGENTE
            IndCodigo = Index + 14
            Set frmMtoAgente = New frmFacAgentesCom
            frmMtoAgente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(IndCodigo).Text) Then txtCodigo(IndCodigo).Text = ""
            frmMtoAgente.Show vbModal
            Set frmMtoAgente = Nothing
            
        Case 0, 24, 31 'Bancos Propios
            IndCodigo = 0
            If Index = 31 Then
                IndCodigo = 52
            ElseIf Index = 0 Then IndCodigo = 5
            End If
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        Case 25 'Tipo CONTRATO
'            IndCodigo = 45
'            Set frmMtoTipCo = New frmManTiposContrato
'            frmMtoTipCo.DatosADevolverBusqueda = "0|1|"
'            frmMtoTipCo.Show vbModal
'            Set frmMtoTipCo = Nothing
    End Select
    PonerFoco txtCodigo(IndCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'Frame Pasar Pedido -> Albaran
            IndCodigo = 25
        Case 1 'framePedidos
            IndCodigo = 3 'Desde
        Case 2 'framePedidos
            IndCodigo = 4 'Hasta
        
        Case 6 'FramePedxArtic
            IndCodigo = 11 'Fecha Desde
        Case 7 'FramePedxArtic
            IndCodigo = 12 'Fecha Hasta
        Case 9 'FramePedCompras
            IndCodigo = 24 'Fecha Hasta
        Case 10 'FramePreFacturar
            IndCodigo = 26
        Case 11 'FramePreFacturar
            IndCodigo = 27
        Case 12 'Frame Factura
            IndCodigo = 38
        Case 13 'Frame Factura
            IndCodigo = 39
        Case 14 'FrameFactura
            IndCodigo = 34
   End Select
   
   PonerFormatoFecha txtCodigo(IndCodigo)
   If txtCodigo(IndCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(IndCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(IndCodigo)
End Sub







Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
    
    If Index = 17 Then
        If txtCodigo(17).Text <> "" Then PonerFoco txtCodigo(18)
    End If
    
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codCampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 1 'Importe (Decimal(12,2))
            PonerFormatoDecimal txtCodigo(Index), 1
            
        Case 0, 5, 52 'Bancos Propios
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                If txtCodigo(Index).Text <> "" And txtNombre(Index).Text <> "" Then
                    txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                Else
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        
        'FECHA Desde Hasta
        Case 11, 12, 25, 26, 27, 34, 38, 39, 44
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
                If Index = 34 Then _
                    txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            End If
           
'            'Fecha entrega para Pedido. Poner la semana
'            If Index = 26 Then txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
        
        Case 53 'AÑO
             PonerFormatoEntero txtCodigo(Index)
        
        Case 36, 37  'Nº de Pedido / Albaran
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            

        Case 35 'Periodicidad Facturacion
            PonerFormatoEntero txtCodigo(Index)

        Case 8, 20, 21, 28, 29, 40, 41, 48, 49 'Cod. CLIENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomclien"
                Tabla = "sclien"
                codCampo = "codclien"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, "Cliente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 13, 14 'ALMACEN
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomalmac"
                Tabla = "salmpr"
                codCampo = "codalmac"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, "Almacen")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
     
        Case 2, 3, 17, 18, 47 'Cod. Trabajador
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomtraba"
                Tabla = "straba"
                codCampo = "codtraba"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, "Trabajador")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 19 'Cod. Envio
            nomCampo = "nomenvio"
            Tabla = "senvio"
            codCampo = "codenvio"
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, "Forma de Envío")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
            
        Case 30, 31, 42, 43, 50, 51 'Cod. Formas de PAGO
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "Formas de Pago")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
        Case 32, 33 'AGENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sagent", "nomagent", "codagent", "Agente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 45 'TIPO CONTRATO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipco", "nomtipco", "codtipco", "Tipo Contrato", "T")
            
        Case 46 'MES a facturar
            If PonerFormatoEntero(txtCodigo(Index)) Then
                'Comprobar que el mes es correcto, valores entre 1-12
                devuelve = txtCodigo(Index).Text
                If (CByte(devuelve) >= 1) And (CByte(devuelve) <= 12) Then
                    txtNombre(Index).Text = UCase(MonthName(CLng(devuelve)))
                Else
                    MsgBox "El valor introducido no es un MES válido.(1-12).", vbInformation
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
            
        '##### Recuperar facturas ALZIRA
        Case 4 'nº factura
            PonerFocoBtn Me.cmdAceptarFac
        '#####
    End Select
End Sub



Private Sub PonerFramePedxArticVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Informe Pedidos por Articulo Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Ofertas

    H = 5415
    'If OpcionListado = 228 Then H = 5000
    W = 7515
    
        
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePedxArtic, visible, H, W
    
    If visible = True Then
        Me.Frame5.visible = (OpcionListado = 44) Or (OpcionListado = 227) 'D/H cliente
        'D/H Artículo
        Me.Frame8.visible = (OpcionListado <> 44) And (OpcionListado <> 227) And (OpcionListado <> 228)
        Me.Frame9.visible = (OpcionListado <> 227 And OpcionListado <> 228) 'D/H Almacen
        Me.Frame10.visible = (OpcionListado = 227)
        Me.Frame12.visible = (OpcionListado = 228)
        
        If OpcionListado = 44 Then 'Informe Pedido por cliente
            Me.Frame5.Top = 3120
            Me.Frame5.Left = 500
            Me.Label1.Caption = "Pedidos por Cliente"
        ElseIf OpcionListado = 227 Then 'Inf. Estadistica ventas x cliente
            Me.Frame5.Top = 1800
            Me.Frame5.Left = 500
            Me.Frame10.Top = 2800
            Me.Label1.Caption = "Ventas por Cliente"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4650
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        ElseIf OpcionListado = 228 Then 'Inf. Estadistica ventas x trabajador
            Me.Frame12.Top = 1900
            Me.Frame12.Left = 500
            Me.Label1.Caption = "Ventas por Trabajador"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4150
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        Else
            Me.Frame8.Top = 3120
            Me.Frame8.Left = 500
            If OpcionListado = 41 Then
                Me.Label1.Caption = "Pedidos por Artículo"
            ElseIf OpcionListado = 42 Then
                Me.Label1.Caption = "Disponibilidad de Stocks"
            ElseIf OpcionListado = 49 Then
                Me.Label1.Caption = "Albaranes por Artículo"
                Label4(4).Caption = "Fecha Albaran"
            End If
        End If
    End If
End Sub


Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim B As Boolean
Dim Cad As String

    H = 5600
    If OpcionListado = 51 Then 'Inf. Incum. plazos entrega
        H = 5300
        Me.cmdAceptarPreFac.Top = 4600
        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
    End If
    W = 7040
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        B = (OpcionListado = 50)
        Label4(41).visible = B
        Me.imgBuscarOfer(16).visible = B
        Me.imgBuscarOfer(17).visible = B
        Me.txtCodigo(30).visible = B
        Me.txtCodigo(31).visible = B
        Me.txtNombre(30).visible = B
        Me.txtNombre(31).visible = B
        Me.Frame6.visible = Not B
        Me.Frame6.Top = 2900
        Me.Frame6.Left = 460
        'solo albaranes a facturar
        Me.chkSoloFacturar.visible = B
        Me.chkSoloFacturar.Value = 1
        
        'Detalle o resumen
        Me.Frame7.visible = B And CodClien = "ALV"
        Me.OptDetalle(0).Value = True
        
        If Not B Then
            Me.Label9(0).Caption = "Incum. plazos entrega"
        Else 'Prevision Facturacion
            Select Case CodClien 'aqui guardamos el tipo de movimiento
                Case "ALV": Cad = "" ' antes "(Ventas)" [SERVICIOS]
                Case "ALR": Cad = "(Reparaciones)"
                Case "ALM": Cad = "(Mantenimientos)"
            End Select
            Me.Label9(0).Caption = "Previsión de facturación " & Cad
        End If
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 7100 + 180
    W = 7480
    
    If visible Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "ALV": Cad = "(Ventas)"
            Case "ALR": Cad = "(Reparaciones)"
            Case "ALM", "ART":
                If CodClien = "ALM" Then
                    Cad = "(Mostrador)"
                Else
                    Cad = "(Rectificativa)"
                End If
                'Me.Frame3.Top = 1200
                Me.Frame4.visible = False
                H = 4000
                Me.cmdAceptarFac.Top = 3200
                Me.cmdCancel(6).Top = Me.cmdAceptarFac.Top
            Case "ALS": Cad = "(Servicios)"
            Case "ALI": Cad = "(Internas)"
                
        End Select
        '#### Laura Recuperar facturas ALZIRA
        'nº de factura solo visible si estamos recuperando facturas
        Me.Label10(9).visible = Me.EstaRecupFact And OpcionListado = 222
        Me.txtCodigo(4).visible = Me.EstaRecupFact And OpcionListado = 222
        If Me.EstaRecupFact And OpcionListado = 222 Then txtCodigo(0).Text = "001"
        
        Me.Label10(0).Caption = "Facturación de Albaranes " & Cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub


Private Sub PonerFramePreFacManteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Mantenimientos Visible y Ajustado al Formulario, y visualiza los controles
Dim B As Boolean
Dim Cad As String

    
    If visible = True Then
        B = (OpcionListado = 74) 'prefacturar
        W = 7120
        If B Then 'prefacturar
            H = 5600
        Else 'facturar
            H = 6200
        End If
        Me.FramePreFacMante.Height = H
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        PonerFrameVisible Me.FramePreFacMante, visible, H, W
        
        If B Then 'prefacturar
            Me.Frame2.visible = False
            Me.Frame1.Top = Me.Frame1.Top - 800
            Me.Frame1.BorderStyle = 0
            Me.Label7(1).Caption = "Prefacturación Mantenimientos"
            Me.cmdAceptarPreFacMan.Top = Me.cmdAceptarPreFacMan.Top - 600
            Me.cmdCancel(7).Top = Me.cmdCancel(7).Top - 600
        Else 'facturar
            Me.Label7(1).Caption = "Facturación Mantenimientos"
            Me.txtCodigo(44).Text = Format(Now, "dd/mm/yyyy")
            Me.txtCodigo(47).Text = PonerTrabajadorConectado(Cad)
            Me.txtNombre(47).Text = Cad
        End If
    End If
End Sub



Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next

    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
            Cad = Cad & txtCodigo(indD).Text
            If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
            AnyadirParametroDH = Cad
            Exit Function
        End If
    End If
    
    If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = Cad
End Function


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            NumParam = NumParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    NumParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = NumParam
        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sartic", "nomartic", "codartic", "Articulo", "T")
            If txtNombre(Index).Text = "" And txtCodigo(Index) <> "" Then Cancel = True
    End Select
End Sub




Private Function ObtenerClientes(cadW As String, Importe As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    SQL = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    SQL = SQL & " From scafac "
    If cadW <> "" Then SQL = SQL & " where " & cadW
    SQL = SQL & " group by codclien "
    If Importe <> "" Then SQL = SQL & "having baseimp>" & Importe
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            SQL = SQL & RS!CodClien & ","
'        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    If SQL <> "" Then
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        SQL = "( {scafac.codclien} IN [" & SQL & "] )"
    End If
    ObtenerClientes = SQL
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function



Private Sub txtCSB_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



