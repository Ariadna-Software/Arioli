VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11235
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   240
      TabIndex        =   479
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   480
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
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
         Index           =   22
         Left            =   360
         TabIndex        =   481
         Top             =   840
         Width           =   5805
      End
   End
   Begin VB.Frame FramePteRecibir 
      Height          =   5205
      Left            =   480
      TabIndex        =   278
      Top             =   240
      Width           =   7035
      Begin VB.Frame Frame7 
         Caption         =   "Ordenar por"
         ForeColor       =   &H00000080&
         Height          =   940
         Left            =   600
         TabIndex        =   294
         Top             =   3960
         Width           =   2055
         Begin VB.OptionButton OptOrdenPed 
            Caption         =   "Nº Pedido"
            Height          =   255
            Left            =   240
            TabIndex        =   296
            Top             =   550
            Width           =   1215
         End
         Begin VB.OptionButton OptOrdenArt 
            Caption         =   "Artículo"
            Height          =   255
            Left            =   240
            TabIndex        =   295
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   288
         Top             =   2760
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   68
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   275
            Top             =   705
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   68
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   290
            Text            =   "Text5"
            Top             =   705
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   67
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   274
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   67
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   289
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label9 
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
            Index           =   15
            Left            =   600
            TabIndex        =   293
            Top             =   705
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   44
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":000C
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Index           =   14
            Left            =   600
            TabIndex        =   292
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   43
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":010E
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Index           =   13
            Left            =   240
            TabIndex        =   291
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   273
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   272
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   280
         Text            =   "Text5"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   270
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   279
         Text            =   "Text5"
         Top             =   1725
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   271
         Top             =   1725
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarPte 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   276
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5280
         TabIndex        =   277
         Top             =   4440
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":0210
         Top             =   2400
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
         Index           =   75
         Left            =   960
         TabIndex        =   287
         Top             =   2400
         Width           =   450
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
         Index           =   74
         Left            =   600
         TabIndex        =   286
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":029B
         Top             =   2400
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
         Index           =   72
         Left            =   3360
         TabIndex        =   285
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Material pendiente de recibir"
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
         Index           =   19
         Left            =   600
         TabIndex        =   284
         Top             =   360
         Width           =   4455
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   41
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":0326
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   600
         TabIndex        =   283
         Top             =   1035
         Width           =   885
      End
      Begin VB.Label Label9 
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
         Index           =   17
         Left            =   960
         TabIndex        =   282
         Top             =   1380
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   42
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":0428
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   16
         Left            =   960
         TabIndex        =   281
         Top             =   1725
         Width           =   420
      End
   End
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6855
      Left            =   120
      TabIndex        =   448
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   7920
         TabIndex        =   462
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   2475
         Index           =   1
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   461
         Text            =   "frmListadoOfer.frx":052A
         Top             =   3720
         Width           =   4335
      End
      Begin VB.ListBox ListTipoMov 
         Height          =   2310
         Index           =   1000
         Left            =   1200
         Style           =   1  'Checkbox
         TabIndex        =   456
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   450
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   111
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   475
         Text            =   "Text5"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   449
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   110
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   472
         Text            =   "Text5"
         Top             =   1185
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   0
         Left            =   5640
         TabIndex        =   460
         Text            =   "Text1"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Copia remitente"
         Height          =   255
         Left            =   5640
         TabIndex        =   459
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "administración"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   458
         Top             =   1320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "comercial"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   457
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   9000
         TabIndex        =   463
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   452
         Top             =   2295
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   451
         Top             =   2295
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   453
         Text            =   "wwwwwww"
         Top             =   3180
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   455
         Top             =   3180
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
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
         Left            =   5640
         TabIndex        =   478
         Top             =   3480
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         Index           =   20
         Left            =   240
         TabIndex        =   477
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Label9 
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
         Index           =   34
         Left            =   600
         TabIndex        =   476
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":0530
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   33
         Left            =   600
         TabIndex        =   474
         Top             =   1185
         Width           =   450
      End
      Begin VB.Label Label9 
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
         Index           =   32
         Left            =   240
         TabIndex        =   473
         Top             =   840
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":0632
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Index           =   19
         Left            =   5640
         TabIndex        =   471
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
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
         Left            =   5640
         TabIndex        =   470
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label14 
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
         Index           =   18
         Left            =   3120
         TabIndex        =   469
         Top             =   2340
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3600
         Picture         =   "frmListadoOfer.frx":0734
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":07BF
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         TabIndex        =   468
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label14 
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
         Index           =   17
         Left            =   600
         TabIndex        =   467
         Top             =   2340
         Width           =   450
      End
      Begin VB.Label Label14 
         Caption         =   "Envio em"
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
         Index           =   16
         Left            =   240
         TabIndex        =   466
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label14 
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
         Index           =   15
         Left            =   240
         TabIndex        =   465
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label14 
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
         Index           =   14
         Left            =   600
         TabIndex        =   464
         Top             =   3165
         Width           =   450
      End
      Begin VB.Label Label14 
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
         Index           =   13
         Left            =   3360
         TabIndex        =   454
         Top             =   3165
         Width           =   420
      End
   End
   Begin VB.Frame FrameGenAlbCom 
      Height          =   4455
      Left            =   240
      TabIndex        =   195
      Top             =   240
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   48
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   198
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   49
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   199
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAlbCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   200
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   201
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   840
         MaxLength       =   4
         TabIndex        =   197
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   196
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Albaran"
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
         Index           =   61
         Left            =   480
         TabIndex        =   215
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a Albaran"
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
         Left            =   600
         TabIndex        =   214
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alb."
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
         Index           =   60
         Left            =   480
         TabIndex        =   213
         Top             =   3000
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":084A
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Albaran de compra: "
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
         Index           =   59
         Left            =   600
         TabIndex        =   203
         Top             =   1200
         Width           =   5115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador del Albaran"
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
         Index           =   58
         Left            =   600
         TabIndex        =   202
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   27
         Left            =   600
         Picture         =   "frmListadoOfer.frx":08D5
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FramePedidos 
      Height          =   4455
      Left            =   600
      TabIndex        =   308
      Top             =   240
      Width           =   6075
      Begin VB.CheckBox chkPedidoValorado 
         Caption         =   "Valorado"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   313
         Top             =   3720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   310
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   75
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   312
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4080
         TabIndex        =   315
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   314
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   74
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   311
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   309
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
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
         Left            =   600
         TabIndex        =   322
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label12 
         Caption         =   "Imprimir otros Pedidos del Proveedor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   321
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   3480
         Picture         =   "frmListadoOfer.frx":09D7
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
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
         Left            =   840
         TabIndex        =   320
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label12 
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
         Left            =   600
         TabIndex        =   319
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Informe de Pedido Compras"
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
         Left            =   600
         TabIndex        =   318
         Top             =   360
         Width           =   4335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0A62
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
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
         Index           =   6
         Left            =   3000
         TabIndex        =   317
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
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
         Left            =   600
         TabIndex        =   316
         Top             =   1320
         Width           =   810
      End
   End
   Begin VB.Frame FrameEstVentasFam 
      Height          =   5805
      Left            =   240
      TabIndex        =   406
      Top             =   0
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   99
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   412
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   98
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   411
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   96
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   426
         Text            =   "Text5"
         Top             =   1020
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   409
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   97
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   425
         Text            =   "Text5"
         Top             =   1365
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   410
         Top             =   1365
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarEstVentas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   419
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   5520
         TabIndex        =   420
         Top             =   5160
         Width           =   975
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         TabIndex        =   407
         Top             =   2400
         Width           =   6495
         Begin VB.CheckBox chkDatosAlbaranes 
            Caption         =   "Datos albaranes"
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   416
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkDetallaArticulo 
            Caption         =   "Detalla articulo"
            Height          =   195
            Left            =   240
            TabIndex        =   415
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Frame FrameDetallaArticulo 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   975
            Left            =   240
            TabIndex        =   482
            Top             =   1560
            Visible         =   0   'False
            Width           =   6135
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   113
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   486
               Text            =   "Text5"
               Top             =   600
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Index           =   113
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   418
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   112
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   483
               Text            =   "Text5"
               Top             =   240
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Index           =   112
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   417
               Top             =   240
               Width           =   1095
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   59
               Left            =   840
               Picture         =   "frmListadoOfer.frx":0AED
               Top             =   600
               Width           =   240
            End
            Begin VB.Label Label9 
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
               Left            =   360
               TabIndex        =   487
               Top             =   600
               Width           =   420
            End
            Begin VB.Label Label9 
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
               Index           =   36
               Left            =   0
               TabIndex        =   485
               Top             =   0
               Width           =   660
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   58
               Left            =   840
               Picture         =   "frmListadoOfer.frx":0BEF
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label9 
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
               Index           =   35
               Left            =   360
               TabIndex        =   484
               Top             =   240
               Width           =   450
            End
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   101
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   414
            Top             =   705
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   101
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   421
            Text            =   "Text5"
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   100
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   413
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   100
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   408
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label9 
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
            Index           =   27
            Left            =   600
            TabIndex        =   424
            Top             =   705
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   55
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":0CF1
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Index           =   26
            Left            =   600
            TabIndex        =   423
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   54
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":0DF3
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
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
            Index           =   25
            Left            =   240
            TabIndex        =   422
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   3720
         Picture         =   "frmListadoOfer.frx":0EF5
         Top             =   2040
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
         Index           =   91
         Left            =   840
         TabIndex        =   433
         Top             =   2040
         Width           =   450
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
         Index           =   90
         Left            =   480
         TabIndex        =   432
         Top             =   1800
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   1335
         Picture         =   "frmListadoOfer.frx":0F80
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
         Index           =   89
         Left            =   3240
         TabIndex        =   431
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Ventas por Familia / Artículo"
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
         Index           =   31
         Left            =   600
         TabIndex        =   430
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   52
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":100B
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Left            =   480
         TabIndex        =   429
         Top             =   795
         Width           =   585
      End
      Begin VB.Label Label9 
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
         Index           =   29
         Left            =   840
         TabIndex        =   428
         Top             =   1020
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   53
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":110D
         Top             =   1365
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   28
         Left            =   840
         TabIndex        =   427
         Top             =   1365
         Width           =   420
      End
   End
   Begin VB.Frame FrameCompras 
      Height          =   5205
      Left            =   360
      TabIndex        =   378
      Top             =   480
      Width           =   7035
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Datos albaranes"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   387
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "Agrupar por"
         ForeColor       =   &H00000080&
         Height          =   945
         Left            =   360
         TabIndex        =   405
         Top             =   3880
         Width           =   2175
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   385
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia, Artículo"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   386
            Top             =   550
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   399
         Top             =   2640
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   94
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   401
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   94
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   383
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   95
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   400
            Text            =   "Text5"
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   95
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   384
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
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
            Index           =   20
            Left            =   240
            TabIndex        =   404
            Top             =   120
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   50
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":120F
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Index           =   12
            Left            =   600
            TabIndex        =   403
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   51
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":1311
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Left            =   600
            TabIndex        =   402
            Top             =   705
            Width           =   420
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5640
         TabIndex        =   389
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarCompras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4560
         TabIndex        =   388
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   91
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   380
         Top             =   1605
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   391
         Text            =   "Text5"
         Top             =   1605
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   90
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   379
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   390
         Text            =   "Text5"
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   92
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   381
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   93
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   382
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         Index           =   24
         Left            =   960
         TabIndex        =   398
         Top             =   1605
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   49
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1413
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   23
         Left            =   960
         TabIndex        =   397
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   600
         TabIndex        =   396
         Top             =   1035
         Width           =   885
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   48
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1515
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Compras por Familia/Artículo"
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
         Index           =   21
         Left            =   600
         TabIndex        =   395
         Top             =   360
         Width           =   4455
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
         Index           =   88
         Left            =   3360
         TabIndex        =   394
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":1617
         Top             =   2280
         Width           =   240
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
         Index           =   87
         Left            =   600
         TabIndex        =   393
         Top             =   2010
         Width           =   495
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
         Index           =   83
         Left            =   960
         TabIndex        =   392
         Top             =   2280
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":16A2
         Top             =   2280
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameRecordatorio 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   6915
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar  coste con:"
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
         Height          =   1695
         Left            =   4080
         TabIndex        =   76
         Top             =   4605
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   680
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1000
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1320
            Width           =   2055
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   15
         Left            =   720
         MaxLength       =   80
         TabIndex        =   37
         Top             =   5100
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   14
         Left            =   720
         MaxLength       =   80
         TabIndex        =   36
         Top             =   4800
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   34
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   33
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   3360
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   32
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame FrameTipoPapel2 
         Caption         =   "Tipo de Papel"
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
         Height          =   735
         Left            =   600
         TabIndex        =   41
         Top             =   5565
         Width           =   2775
         Begin VB.OptionButton OptPapelMembreteR 
            Caption         =   "Con Membrete"
            Height          =   255
            Left            =   1320
            TabIndex        =   49
            Top             =   320
            Width           =   1335
         End
         Begin VB.OptionButton OptPapelBlancoR 
            Caption         =   "Blanco"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   320
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   8
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   39
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdAcetarRecorda 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   38
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   35
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   7
         Left            =   1720
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1720
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lineas"
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
         Left            =   600
         TabIndex        =   75
         Top             =   4560
         Width           =   540
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
         Index           =   34
         Left            =   960
         TabIndex        =   74
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":172D
         Top             =   3720
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
         Index           =   33
         Left            =   960
         TabIndex        =   72
         Top             =   3360
         Width           =   450
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
         Index           =   32
         Left            =   600
         TabIndex        =   71
         Top             =   3120
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":182F
         Top             =   3360
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
         Index           =   31
         Left            =   960
         TabIndex        =   69
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1931
         Top             =   2770
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
         Index           =   30
         Left            =   960
         TabIndex        =   62
         Top             =   2400
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
         Index           =   29
         Left            =   600
         TabIndex        =   61
         Top             =   2160
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1A33
         Top             =   2410
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
         Index           =   28
         Left            =   3130
         TabIndex        =   57
         Top             =   1200
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
         Index           =   27
         Left            =   960
         TabIndex        =   56
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   3610
         Picture         =   "frmListadoOfer.frx":1B35
         Top             =   1800
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
         Index           =   26
         Left            =   960
         TabIndex        =   55
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Oferta"
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
         Index           =   25
         Left            =   600
         TabIndex        =   54
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Recordatorio de Ofertas"
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
         TabIndex        =   53
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
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
         Left            =   600
         TabIndex        =   52
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1BC0
         Top             =   4220
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1CC2
         Top             =   1800
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
         Index           =   22
         Left            =   3130
         TabIndex        =   51
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oferta"
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
         Index           =   21
         Left            =   600
         TabIndex        =   50
         Top             =   960
         Width           =   780
      End
   End
   Begin VB.Frame FrameEtiqProv 
      Height          =   5325
      Left            =   840
      TabIndex        =   246
      Top             =   360
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   62
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   208
         Top             =   3360
         Width           =   4335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   212
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEtiqProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   211
         Top             =   4560
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   360
         TabIndex        =   259
         Top             =   3645
         Width           =   6255
         Begin VB.Frame Frame3 
            Caption         =   "e-Mail"
            Height          =   780
            Left            =   1960
            TabIndex        =   262
            Top             =   560
            Width           =   1575
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   264
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   263
               Top             =   460
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   210
            Top             =   560
            Width           =   1575
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   63
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   260
            Text            =   "Text5"
            Top             =   105
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   209
            Top             =   105
            Width           =   855
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":1D4D
            Top             =   105
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
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
            TabIndex        =   261
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   60
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   255
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   206
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   61
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   254
         Text            =   "Text5"
         Top             =   2865
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   207
         Top             =   2865
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   205
         Top             =   1845
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   59
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   248
         Text            =   "Text5"
         Top             =   1845
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   204
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   247
         Text            =   "Text5"
         Top             =   1500
         Width           =   3735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A la atención de:"
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
         Left            =   600
         TabIndex        =   253
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPostal"
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
         Left            =   600
         TabIndex        =   258
         Top             =   2280
         Width           =   630
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   37
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1E4F
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   6
         Left            =   960
         TabIndex        =   257
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":1F51
         Top             =   2865
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   7
         Left            =   960
         TabIndex        =   256
         Top             =   2865
         Width           =   420
      End
      Begin VB.Label Label9 
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
         Index           =   4
         Left            =   960
         TabIndex        =   252
         Top             =   1845
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   36
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2053
         Top             =   1845
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Left            =   960
         TabIndex        =   251
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   600
         TabIndex        =   250
         Top             =   1155
         Width           =   885
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   35
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2155
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Proveedores"
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
         Left            =   600
         TabIndex        =   249
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameFacRectif 
      Height          =   4455
      Left            =   720
      TabIndex        =   297
      Top             =   480
      Width           =   5715
      Begin VB.TextBox txtCodigo 
         Height          =   645
         Index           =   87
         Left            =   600
         MaxLength       =   72
         MultiLine       =   -1  'True
         TabIndex        =   305
         Top             =   2760
         Width           =   4575
      End
      Begin VB.ComboBox cboTipomov 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListadoOfer.frx":2257
         Left            =   1865
         List            =   "frmListadoOfer.frx":2259
         Style           =   2  'Dropdown List
         TabIndex        =   302
         Top             =   1185
         Width           =   1875
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   3240
         TabIndex        =   307
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarFacRect 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2040
         TabIndex        =   306
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1865
         MaxLength       =   10
         TabIndex        =   304
         Top             =   2115
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   71
         Left            =   1865
         MaxLength       =   10
         TabIndex        =   303
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
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
         Index           =   82
         Left            =   600
         TabIndex        =   365
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
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
         Index           =   79
         Left            =   600
         TabIndex        =   301
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1500
         Picture         =   "frmListadoOfer.frx":225B
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         Index           =   77
         Left            =   600
         TabIndex        =   300
         Top             =   2115
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Factura a Rectificar"
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
         Index           =   5
         Left            =   480
         TabIndex        =   299
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   76
         Left            =   600
         TabIndex        =   298
         Top             =   1725
         Width           =   780
      End
   End
   Begin VB.Frame FramePasarHco 
      Height          =   4575
      Left            =   120
      TabIndex        =   216
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   52
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   227
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   219
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   51
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   222
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   218
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5400
         TabIndex        =   221
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   220
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   50
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   217
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   29
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":22E6
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Index           =   65
         Left            =   720
         TabIndex        =   228
         Top             =   2760
         Width           =   720
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   28
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":23E8
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
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
         Index           =   64
         Left            =   720
         TabIndex        =   226
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el histórico: "
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
         Index           =   63
         Left            =   600
         TabIndex        =   225
         Top             =   1200
         Width           =   4245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   2040
         Picture         =   "frmListadoOfer.frx":24EA
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Eliminación"
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
         Index           =   62
         Left            =   720
         TabIndex        =   224
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Albaran al histórico"
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
         Left            =   600
         TabIndex        =   223
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame FrameGenPedido 
      Height          =   4455
      Left            =   360
      TabIndex        =   100
      Top             =   120
      Width           =   6315
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   107
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   1820
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   64
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   68
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarGenPed 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   67
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   26
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   66
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   25
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   65
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Semana"
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
         Index           =   16
         Left            =   3600
         TabIndex        =   108
         Top             =   3000
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListadoOfer.frx":2575
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador de Pedido"
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
         TabIndex        =   106
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Pedido: "
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
         TabIndex        =   104
         Top             =   1200
         Width           =   4080
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1920
         Picture         =   "frmListadoOfer.frx":2677
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrega"
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
         Left            =   840
         TabIndex        =   103
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Oferta a Pedido"
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
         Left            =   600
         TabIndex        =   102
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   1920
         Picture         =   "frmListadoOfer.frx":2702
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
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
         Left            =   840
         TabIndex        =   101
         Top             =   2520
         Width           =   960
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   3735
      Left            =   0
      TabIndex        =   366
      Top             =   0
      Width           =   6315
      Begin VB.Frame FrameAgrupar 
         Caption         =   "Agrupar por"
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
         Height          =   1000
         Left            =   600
         TabIndex        =   377
         Top             =   2160
         Width           =   2055
         Begin VB.OptionButton optForpago 
            Caption         =   "Tipo de pago"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   370
            Top             =   620
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optForpago 
            Caption         =   "Forma de pago"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   369
            Top             =   320
            Width           =   1695
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   88
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   367
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarCierre 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   371
         Top             =   2785
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4440
         TabIndex        =   372
         Top             =   2785
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   89
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   368
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Index           =   3
         Left            =   3480
         TabIndex        =   376
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1480
         Picture         =   "frmListadoOfer.frx":278D
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Cierre de Caja"
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
         Left            =   600
         TabIndex        =   375
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label10 
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
         Left            =   600
         TabIndex        =   374
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
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
         Index           =   2
         Left            =   960
         TabIndex        =   373
         Top             =   1560
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   3960
         Picture         =   "frmListadoOfer.frx":2818
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame FrameEfectuadas 
      Height          =   4455
      Left            =   360
      TabIndex        =   81
      Top             =   240
      Width           =   6315
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Solo Ofertas Pendientes"
         Height          =   255
         Left            =   720
         TabIndex        =   91
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   46
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   45
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   44
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   48
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEfect 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   47
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   16
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   43
         Top             =   1560
         Width           =   1215
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
         Left            =   960
         TabIndex        =   90
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   7
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":28A3
         Top             =   2640
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
         Index           =   3
         Left            =   960
         TabIndex        =   89
         Top             =   2280
         Width           =   450
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
         Index           =   4
         Left            =   600
         TabIndex        =   88
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   6
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":29A5
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   3960
         Picture         =   "frmListadoOfer.frx":2AA7
         Top             =   1560
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
         Index           =   10
         Left            =   960
         TabIndex        =   87
         Top             =   1560
         Width           =   450
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
         Index           =   11
         Left            =   600
         TabIndex        =   86
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ofertas Efectuadas"
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
         TabIndex        =   85
         Top             =   600
         Width           =   3855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2B32
         Top             =   1560
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
         Index           =   13
         Left            =   3480
         TabIndex        =   84
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame FrameClienInactivos 
      Height          =   7005
      Left            =   0
      TabIndex        =   109
      Top             =   -120
      Width           =   10995
      Begin VB.Frame frameCliexFacturas 
         Caption         =   "Desde / hasta facturas"
         Height          =   4455
         Left            =   6480
         TabIndex        =   435
         Top             =   1080
         Width           =   4455
         Begin VB.ComboBox cboTipomov 
            Height          =   315
            Index           =   2
            ItemData        =   "frmListadoOfer.frx":2BBD
            Left            =   920
            List            =   "frmListadoOfer.frx":2BBF
            Style           =   2  'Dropdown List
            TabIndex        =   436
            Top             =   840
            Width           =   1875
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   104
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   439
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   102
            Left            =   960
            MaxLength       =   7
            TabIndex        =   437
            Text            =   "wwwwwww"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   103
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   438
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   105
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   440
            Top             =   3240
            Width           =   1080
         End
         Begin VB.Image imgClearCmbTipomov 
            Height          =   240
            Left            =   2880
            Picture         =   "frmListadoOfer.frx":2BC1
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
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
            Left            =   120
            TabIndex        =   447
            Top             =   600
            Width           =   1410
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   31
            Left            =   1635
            Picture         =   "frmListadoOfer.frx":314B
            Top             =   3000
            Width           =   240
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fact."
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
            TabIndex        =   446
            Top             =   2880
            Width           =   945
         End
         Begin VB.Label Label14 
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
            Index           =   11
            Left            =   240
            TabIndex        =   445
            Top             =   1680
            Width           =   885
         End
         Begin VB.Label Label14 
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
            Left            =   960
            TabIndex        =   444
            Top             =   1920
            Width           =   450
         End
         Begin VB.Label Label14 
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
            Left            =   2640
            TabIndex        =   443
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label14 
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
            Left            =   1080
            TabIndex        =   442
            Top             =   3030
            Width           =   450
         End
         Begin VB.Label Label14 
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
            Left            =   2640
            TabIndex        =   441
            Top             =   3023
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   32
            Left            =   3120
            Picture         =   "frmListadoOfer.frx":31D6
            Top             =   3000
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   6600
         TabIndex        =   265
         Top             =   1680
         Width           =   4215
         Begin VB.Frame Frame5 
            Caption         =   "e-Mail"
            Height          =   780
            Left            =   600
            TabIndex        =   124
            Top             =   1680
            Width           =   2000
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Comercial"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   269
               Top             =   460
               Width           =   1335
            End
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   268
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   64
            Left            =   180
            MaxLength       =   6
            TabIndex        =   122
            Top             =   860
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   64
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   266
            Text            =   "Text5"
            Top             =   860
            Width           =   3375
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   123
            Top             =   1395
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
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
            Left            =   180
            TabIndex        =   267
            Top             =   650
            Width           =   465
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   40
            Left            =   840
            Picture         =   "frmListadoOfer.frx":3261
            Top             =   580
            Width           =   240
         End
      End
      Begin VB.Frame FrameImpClien 
         Caption         =   "Imprimir clientes"
         ForeColor       =   &H00000080&
         Height          =   1050
         Left            =   600
         TabIndex        =   121
         Top             =   5760
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton OptCliTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   240
            TabIndex        =   243
            Top             =   735
            Width           =   1215
         End
         Begin VB.OptionButton OptCliSinMante 
            Caption         =   "Sin mantenimiento"
            Height          =   255
            Left            =   240
            TabIndex        =   242
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton OptCliConMante 
            Caption         =   "Con mantenimiento"
            Height          =   255
            Left            =   240
            TabIndex        =   241
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   480
         TabIndex        =   229
         Top             =   2900
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   57
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   244
            Text            =   "Text5"
            Top             =   2025
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   57
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   119
            Top             =   2025
            Width           =   855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   120
            Top             =   2385
            Width           =   4095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   56
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   118
            Top             =   1470
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   56
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   237
            Text            =   "Text5"
            Top             =   1470
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   117
            Top             =   1130
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   55
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   236
            Text            =   "Text5"
            Top             =   1130
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   116
            Top             =   580
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   54
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   231
            Text            =   "Text5"
            Top             =   580
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   115
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   53
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   230
            Text            =   "Text5"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   34
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3363
            Top             =   2025
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Situación"
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
            Index           =   73
            Left            =   120
            TabIndex        =   245
            Top             =   2025
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "A la atención de:"
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
            Index           =   71
            Left            =   120
            TabIndex        =   240
            Top             =   2385
            Width           =   1395
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
            Index           =   70
            Left            =   480
            TabIndex        =   239
            Top             =   1470
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   33
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3465
            Top             =   1470
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
            Index           =   69
            Left            =   480
            TabIndex        =   238
            Top             =   1130
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3567
            Top             =   1130
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CPostal"
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
            Index           =   68
            Left            =   120
            TabIndex        =   235
            Top             =   890
            Width           =   630
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
            Index           =   67
            Left            =   480
            TabIndex        =   234
            Top             =   580
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   31
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3669
            Top             =   580
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
            Index           =   66
            Left            =   480
            TabIndex        =   233
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Actividad"
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
            Left            =   120
            TabIndex        =   232
            Top             =   0
            Width           =   795
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   30
            Left            =   960
            Picture         =   "frmListadoOfer.frx":376B
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   32
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   127
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   31
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   114
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarClienInac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   125
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5280
         TabIndex        =   126
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "Text5"
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   110
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "Text5"
         Top             =   1600
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   111
         Top             =   1600
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   2200
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   112
         Top             =   2200
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "Text5"
         Top             =   2550
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   113
         Top             =   2550
         Width           =   855
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
         Left            =   3250
         TabIndex        =   141
         Top             =   3360
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   3720
         Picture         =   "frmListadoOfer.frx":386D
         Top             =   3375
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
         Index           =   43
         Left            =   960
         TabIndex        =   140
         Top             =   3360
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":38F8
         Top             =   3380
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inactividad"
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
         Left            =   600
         TabIndex        =   139
         Top             =   3120
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Clientes Inactivos"
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
         TabIndex        =   138
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":3983
         Top             =   1260
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
         Index           =   42
         Left            =   600
         TabIndex        =   137
         Top             =   1040
         Width           =   585
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
         Index           =   41
         Left            =   960
         TabIndex        =   136
         Top             =   1260
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":3A85
         Top             =   1600
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
         Index           =   40
         Left            =   960
         TabIndex        =   135
         Top             =   1600
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":3B87
         Top             =   2200
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
         Index           =   39
         Left            =   600
         TabIndex        =   134
         Top             =   1940
         Width           =   615
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
         Index           =   38
         Left            =   960
         TabIndex        =   133
         Top             =   2200
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":3C89
         Top             =   2550
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
         Index           =   37
         Left            =   960
         TabIndex        =   132
         Top             =   2550
         Width           =   420
      End
   End
   Begin VB.Frame FrameClientes 
      Height          =   6015
      Left            =   120
      TabIndex        =   142
      Top             =   120
      Width           =   8955
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   155
         Top             =   4695
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   156
         Top             =   5010
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   41
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   181
         Text            =   "Text5"
         Top             =   4695
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   42
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   180
         Text            =   "Text5"
         Top             =   5010
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   164
         Text            =   "Text5"
         Top             =   3270
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text5"
         Top             =   2955
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   152
         Top             =   3270
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   151
         Top             =   2955
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   7440
         TabIndex        =   158
         Top             =   5295
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6360
         TabIndex        =   157
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   147
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   148
         Top             =   1635
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   33
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   162
         Text            =   "Text5"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   34
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "Text5"
         Top             =   1635
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   149
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   150
         Top             =   2475
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   35
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "Text5"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "Text5"
         Top             =   2475
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   153
         Top             =   3795
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   154
         Top             =   4110
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   3795
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   40
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "Text5"
         Top             =   4110
         Width           =   3135
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":3D8B
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   2505
         Width           =   435
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":4095
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   1720
         Width           =   435
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   6480
         TabIndex        =   165
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   184
         Top             =   4695
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   183
         Top             =   5010
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
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
         Index           =   45
         Left            =   600
         TabIndex        =   182
         Top             =   4440
         Width           =   780
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   21
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":439F
         Top             =   4695
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   22
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":44A1
         Top             =   5025
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   18
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":45A3
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":46A5
         Top             =   2955
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
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
         Index           =   51
         Left            =   600
         TabIndex        =   179
         Top             =   2715
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   1080
         TabIndex        =   178
         Top             =   3270
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   1080
         TabIndex        =   177
         Top             =   2955
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   1080
         TabIndex        =   176
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   1080
         TabIndex        =   175
         Top             =   1635
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
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
         Index           =   49
         Left            =   600
         TabIndex        =   174
         Top             =   1080
         Width           =   795
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   13
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":47A7
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":48A9
         Top             =   1650
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Clientes"
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
         Left            =   600
         TabIndex        =   173
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   1080
         TabIndex        =   172
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   1080
         TabIndex        =   171
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   48
         Left            =   600
         TabIndex        =   170
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":49AB
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":4AAD
         Top             =   2505
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   1080
         TabIndex        =   169
         Top             =   3795
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   1080
         TabIndex        =   168
         Top             =   4110
         Width           =   420
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
         Index           =   47
         Left            =   600
         TabIndex        =   167
         Top             =   3540
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   19
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":4BAF
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   20
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":4CB1
         Top             =   4125
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
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
         Index           =   46
         Left            =   6480
         TabIndex        =   166
         Top             =   1200
         Width           =   1545
      End
   End
   Begin VB.Frame FrameFacReimprimir 
      Height          =   4455
      Left            =   240
      TabIndex        =   348
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chkFormatoTPV 
         Caption         =   "Formato factura TPV"
         Height          =   255
         Left            =   480
         TabIndex        =   434
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk_duplicado 
         Caption         =   "Duplicado"
         Height          =   375
         Left            =   480
         TabIndex        =   364
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   4365
         MaxLength       =   10
         TabIndex        =   353
         Top             =   2880
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   4080
         MaxLength       =   7
         TabIndex        =   351
         Top             =   2172
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   1400
         MaxLength       =   7
         TabIndex        =   350
         Text            =   "wwwwwww"
         Top             =   2172
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1685
         MaxLength       =   10
         TabIndex        =   352
         Top             =   2880
         Width           =   1080
      End
      Begin VB.CommandButton cmdAceptarReimpFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   354
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   355
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cboTipomov 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListadoOfer.frx":4DB3
         Left            =   1400
         List            =   "frmListadoOfer.frx":4DB5
         Style           =   2  'Dropdown List
         TabIndex        =   349
         Top             =   1425
         Width           =   1875
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   4080
         Picture         =   "frmListadoOfer.frx":4DB7
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   3600
         TabIndex        =   363
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label Label14 
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
         Left            =   840
         TabIndex        =   362
         Top             =   2895
         Width           =   450
      End
      Begin VB.Label Label14 
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
         Index           =   6
         Left            =   3600
         TabIndex        =   361
         Top             =   2145
         Width           =   420
      End
      Begin VB.Label Label14 
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
         Left            =   840
         TabIndex        =   360
         Top             =   2145
         Width           =   450
      End
      Begin VB.Label Label14 
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
         Index           =   4
         Left            =   480
         TabIndex        =   359
         Top             =   1867
         Width           =   885
      End
      Begin VB.Label Label14 
         Caption         =   "Reimprimir Facturas"
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
         Left            =   480
         TabIndex        =   358
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         Left            =   480
         TabIndex        =   357
         Top             =   2595
         Width           =   945
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1400
         Picture         =   "frmListadoOfer.frx":4E42
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
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
         Left            =   480
         TabIndex        =   356
         Top             =   1140
         Width           =   1410
      End
   End
   Begin VB.Frame FrameTraspasoHco 
      Height          =   5295
      Left            =   600
      TabIndex        =   92
      Top             =   360
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   43
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   44
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   190
         Text            =   "Text5"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text5"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   185
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   4140
         MaxLength       =   7
         TabIndex        =   27
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   22
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarTrasHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   58
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   59
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   23
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   25
         Top             =   3360
         Width           =   1215
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
         Index           =   56
         Left            =   600
         TabIndex        =   194
         Top             =   1200
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   23
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4ECD
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
         Index           =   55
         Left            =   960
         TabIndex        =   193
         Top             =   1440
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   24
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4FCF
         Top             =   1800
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
         Index           =   54
         Left            =   960
         TabIndex        =   192
         Top             =   1800
         Width           =   420
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
         Index           =   53
         Left            =   600
         TabIndex        =   189
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   25
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":50D1
         Top             =   2400
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
         Index           =   52
         Left            =   960
         TabIndex        =   188
         Top             =   2400
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   26
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":51D3
         Top             =   2760
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
         Left            =   960
         TabIndex        =   187
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oferta"
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
         Left            =   600
         TabIndex        =   99
         Top             =   3720
         Width           =   780
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
         TabIndex        =   98
         Top             =   3960
         Width           =   450
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
         Index           =   1
         Left            =   3360
         TabIndex        =   97
         Top             =   3960
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
         Index           =   9
         Left            =   3360
         TabIndex        =   96
         Top             =   3360
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":52D5
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Traspaso de Ofertas a Histórico"
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
         Left            =   600
         TabIndex        =   95
         Top             =   480
         Width           =   4695
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
         Index           =   8
         Left            =   600
         TabIndex        =   94
         Top             =   3120
         Width           =   495
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
         Left            =   960
         TabIndex        =   93
         Top             =   3360
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":5360
         Top             =   3360
         Width           =   240
      End
   End
   Begin VB.Frame FrameConfirmPed 
      Height          =   6255
      Left            =   480
      TabIndex        =   323
      Top             =   120
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   80
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   327
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   80
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   338
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   79
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   326
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   79
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   337
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Frame FrameTipoPapel3 
         Caption         =   "Tipo de Papel"
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
         Height          =   735
         Left            =   600
         TabIndex        =   334
         Top             =   4485
         Width           =   3375
         Begin VB.OptionButton OptPapelMembrete3 
            Caption         =   "Con Membrete"
            Height          =   255
            Left            =   1800
            TabIndex        =   336
            Top             =   320
            Width           =   1335
         End
         Begin VB.OptionButton OptPapelBlanco3 
            Caption         =   "Blanco"
            Height          =   195
            Left            =   240
            TabIndex        =   335
            Top             =   320
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   78
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   325
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5280
         TabIndex        =   332
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton cmdAcetarConfirm 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   331
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   81
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   328
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   81
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   333
         Text            =   "Text5"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   77
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   324
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   82
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   329
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox chkImpSaldo 
         Caption         =   "Imprimir saldo"
         Height          =   375
         Left            =   4680
         TabIndex        =   330
         Top             =   4680
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
         Index           =   86
         Left            =   960
         TabIndex        =   347
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   47
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":53EB
         Top             =   2640
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
         Index           =   85
         Left            =   960
         TabIndex        =   346
         Top             =   2280
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
         Index           =   84
         Left            =   600
         TabIndex        =   345
         Top             =   2040
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   46
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":54ED
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   3600
         Picture         =   "frmListadoOfer.frx":55EF
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label13 
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
         Index           =   2
         Left            =   960
         TabIndex        =   344
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label13 
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
         Index           =   1
         Left            =   600
         TabIndex        =   343
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label13 
         Caption         =   "Cartas Confirmación de Pedidos"
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
         Left            =   480
         TabIndex        =   342
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
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
         Index           =   81
         Left            =   600
         TabIndex        =   341
         Top             =   3360
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   45
         Left            =   1155
         Picture         =   "frmListadoOfer.frx":567A
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":577C
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
         Index           =   80
         Left            =   3120
         TabIndex        =   340
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Carta"
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
         Index           =   78
         Left            =   600
         TabIndex        =   339
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":5807
         Top             =   3720
         Width           =   240
      End
   End
   Begin VB.Frame FrameOfertas 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1660
         MaxLength       =   10
         TabIndex        =   6
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdAceptarOfer 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   9
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   4160
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Frame FrameTipoPapel 
         Caption         =   "Tipo de Papel"
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
         Height          =   855
         Left            =   600
         TabIndex        =   1
         Top             =   1720
         Width           =   3375
         Begin VB.OptionButton OptPapelBlanco 
            Caption         =   "Blanco"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptPapelMembrete 
            Caption         =   "Con Membrete"
            Height          =   255
            Left            =   1800
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oferta"
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
         Left            =   600
         TabIndex        =   17
         Top             =   1200
         Width           =   780
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
         Left            =   3360
         TabIndex        =   16
         Top             =   4320
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1340
         Picture         =   "frmListadoOfer.frx":5892
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListadoOfer.frx":591D
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
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
         Left            =   600
         TabIndex        =   15
         Top             =   2880
         Width           =   465
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
         Index           =   20
         Left            =   600
         TabIndex        =   13
         Top             =   3960
         Width           =   495
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
         Index           =   23
         Left            =   840
         TabIndex        =   12
         Top             =   4320
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":5A1F
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Imprimir otras Ofertas del Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Ofertas"
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
         TabIndex        =   14
         Top             =   360
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
    '316:  Exportar PDFs
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta/pedido a imprimir

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmFacClientes
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoProve As frmComProveedores
Attribute frmMtoProve.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoZona As frmFacZonas
Attribute frmMtoZona.VB_VarHelpID = -1
Private WithEvents frmMtoRuta As frmFacRutas
Attribute frmMtoRuta.VB_VarHelpID = -1
Private WithEvents frmMtoSitua As frmFacSituaciones
Attribute frmMtoSitua.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoArtic As frmAlmArticulos
Attribute frmMtoArtic.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'codigo postal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1



'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private Cadparam As String 'cadena con los parametros q se pasan a Crystal R.
Private NumParam As Byte
Private Cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------



Dim IndCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'Código para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub







Private Sub cboDeposito_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipomov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDatosAlbaranes_Click(Index As Integer)
    If Index = 0 Then
        Label4(90).Caption = "Fecha"
        If Me.chkDatosAlbaranes(0).Value = 1 Then Label4(90).Caption = Label4(90).Caption & " albaran"
    Else
        Label4(87).Caption = "Fecha"
        If Me.chkDatosAlbaranes(1).Value = 1 Then Label4(87).Caption = Label4(87).Caption & " albaran"
    End If
End Sub

Private Sub chkDatosAlbaranes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDetallaArticulo_Click()
    Me.FrameDetallaArticulo.visible = chkDetallaArticulo.Value = 1
End Sub

Private Sub chkDetallaArticulo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub chkImpSaldo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMail_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPedidoValorado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptarAlbCom_Click()
'Solicitar datos para Generar Albaran  a partir de Pedido de Compras
Dim cad As String
    txtCodigo(48).Text = Trim(txtCodigo(48).Text)
    If txtCodigo(48).Text = "" Then
        cad = "Deberia indicar el número de albarán.  ¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then
            PonerFoco txtCodigo(48)
            Exit Sub
        End If
    End If
    cad = txtCodigo(47).Text & "|"
    cad = cad & txtCodigo(48).Text & "|"
    cad = cad & txtCodigo(49).Text & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub cmdAceptarCierre_Click()
'Cierre de caja del TPV
Dim campo As String
Dim Devuelve As String


    InicializarVbles
    
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    Cadparam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    NumParam = NumParam + 1


    'comprobar que se ha introducido FECHA
    '---------------------------------------------------------
    If Trim(txtCodigo(88).Text) <> "" Or Trim(txtCodigo(89).Text) <> "" Then
        'Para Crystal Report
        campo = "{scafac.fecfactu}"
        Devuelve = "pDHFecha=""FECHA: " 'Parametro Desde/Hasta Fecha
        If Not PonerDesdeHasta(campo, "F", 88, 89, Devuelve) Then Exit Sub
    Else
        MsgBox "Debe introducir la fecha de cierre.    ", vbExclamation
        Exit Sub
    End If
    
    
    'Seleccionar solo las facturas que vienen de TICKET del TPV
    campo = "{scafac1.numventa}"
    campo = "(NOT ISNULL(" & campo & ")) and (" & campo & "<>0)"
    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
    If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
    
    
    'Seleccionar solo las facturas que los albaranes fueron generados en el TPV
    'para ello seleccionar que scafac1.codtipoa='ATI'
    campo = "{scafac1.codtipoa} = 'ATI'"
    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
    If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
    
    
    'ver si hay registros seleccionados para mostrar en el informe
    campo = "(scafac INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu)  INNER JOIN sforpa ON scafac.codforpa = sforpa.codforpa "
    If Not HayRegParaInforme(campo, Cadselect) Then Exit Sub
    
    Titulo = "Cierre de Caja"
    If Me.optForpago(0).Value = True Then
        'informe por Forma de Pago
        nomRPT = "rTPVcierreFP.rpt"
    Else
        'informe por Tipo de Forma de Pago
        nomRPT = "rTPVcierre.rpt"
    End If
    conSubRPT = True
    LlamarImprimir
     
End Sub

Private Sub cmdAceptarClien_Click()
'Listado de Clientes
Dim campo As String, Devuelve As String
Dim numOp As Byte

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    Cadparam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    NumParam = NumParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H ACTIVIDAD
    '--------------------------------------------
     If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{sclien.codactiv}"
        'Parametro Desde/Hasta Actividad
        Devuelve = "pDHActividad=""Actividad: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, Devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ZONA
    '--------------------------------------------
     If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{sclien.codzonas}"
        'Parametro Desde/Hasta Zona
        Devuelve = "pDHZona=""Zona: "
        If Not PonerDesdeHasta(campo, "N", 35, 36, Devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H RUTA
    '--------------------------------------------
     If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = "{sclien.codrutas}"
        'Parametro Desde/Hasta Ruta
        Devuelve = "pDHRuta=""Ruta: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, Devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        Devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 39, 40, Devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H SITUACION
    '--------------------------------------------
     If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = "{sclien.codsitua}"
        'Parametro Desde/Hasta Situacion
        Devuelve = "pDHSituacion=""Situación: "
        If Not PonerDesdeHasta(campo, "N", 41, 42, Devuelve) Then Exit Sub
    End If
    
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
    numOp = PonerGrupo(3, ListView1.ListItems(3).Text)
    numOp = PonerGrupo(4, ListView1.ListItems(4).Text)

    Cadselect = cadFormula
    If Not HayRegParaInforme("sclien", Cadselect) Then Exit Sub
    
    LlamarImprimir
End Sub


Private Sub cmdAceptarClienInac_Click()
'46: Informe de Clientes Inactivos
'47: Informe de Altas Nuevos Clientes
'90: Informe Etiquetas de clientes
Dim campo As String, Devuelve As String

    InicializarVbles
    
    If OpcionListado = 46 Then
        'Comprobar que se ha introdicido una FECHA de Inactividad
        If txtCodigo(31).Text = "" Then
            MsgBox "Debe introducir la Fecha de Inactividad para el informe.", vbInformation
            Exit Sub
        End If
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacClienInactivos.rpt"
        Titulo = "Clientes Inactivos"
        
    ElseIf OpcionListado = 48 Then
        'Comprobar si se ha introducido D/H FECHA Alta
        If txtCodigo(31).Text = "" And txtCodigo(32).Text = "" Then
            MsgBox "Debe introducir algún intervalo de Fechas de Alta.", vbInformation
            Exit Sub
        End If
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacClienAltas.rpt"
    End If
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    Cadparam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    NumParam = NumParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(27).Text <> "" Or txtCodigo(28).Text <> "" Then
        campo = "{sclien.codclien}"
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 27, 28, Devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        Devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 29, 30, Devuelve) Then Exit Sub
    End If
    
    
    
    If OpcionListado = 90 Or OpcionListado = 91 Then '90: Etiquetas de clientes
                                                     '91: Cartas a clientes
        'Cadena para seleccion D/H ACTIVIDAD
        '--------------------------------------------
         If txtCodigo(53).Text <> "" Or txtCodigo(54).Text <> "" Then
            campo = "{sclien.codactiv}"
            'Parametro Desde/Hasta Actividad
            Devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 53, 54, Devuelve) Then Exit Sub
        End If
                        
        'Cadena para seleccion D/H COD. POSTAL
        '--------------------------------------------
         If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta cod. Postal
            Devuelve = "pDHcpostal=""CPostal: "
            If Not PonerDesdeHasta(campo, "T", 55, 56, Devuelve) Then Exit Sub
        End If
        
        'Cadena para seleccion SITUACION
        '--------------------------------------------
        If txtCodigo(57).Text <> "" Then
            campo = "{sclien.codsitua}=" & txtCodigo(57).Text
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
            If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
        End If
        
        'Parametro a la Atencion de
        Cadparam = Cadparam & "pAtencion=""Att. " & txtCodigo(0).Text & """|"
        NumParam = NumParam + 1
        
        'seleccionamos todos los clientes por defecto,
        'pero si seleccionamos clientes con mantenimientos o sin mantenimientos
         'Comprobar si hay registros a Mostrar antes de abrir el Informe
        Cadselect = QuitarCaracterACadena(cadFormula, "{")
        Cadselect = QuitarCaracterACadena(Cadselect, "}")
        
        Devuelve = ""
        If Me.OptCliConMante Then
            Devuelve = ListaClientesMante(Cadselect)
            If Devuelve <> "" Then
                cadFormula = "{sclien.codclien} IN [" & Devuelve & "]"
                Cadselect = "sclien.codclien IN (" & Devuelve & ")"
            End If
        ElseIf Me.OptCliSinMante Then
            Devuelve = ListaClientesMante(Cadselect)
            If Devuelve <> "" Then
                campo = " NOT( {sclien.codclien}  IN [" & Devuelve & "])"
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                campo = " sclien.codclien NOT IN (" & Devuelve & ")"
                If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
            End If
        End If
        
        If OpcionListado = 90 Then
            
            Devuelve = ListaClientesDesdeHastaFactura()
            'Puede haber puesto desde hasta datos factura
            If Devuelve <> "" Then
                campo = " ( {sclien.codclien}  IN [" & Devuelve & "])"
                If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
            End If
        End If
        
        
        
        
        
        
        
        
        
        If OpcionListado = 90 Then 'Etiquetas
            'Nombre fichero .rpt a Imprimir
            nomRPT = "rFacClienEtiq.rpt"
            Titulo = "Etiquetas de Clientes"
            conSubRPT = False
        Else '91: CARTA/e-MAil
            'Parametro cod. carta
            Cadparam = "|pCodCarta= " & txtCodigo(64).Text & "|"
            NumParam = NumParam + 1
            
            'Nombre fichero .rpt a Imprimir
            nomRPT = "rFacClienCarta.rpt"
            Titulo = "Cartas a Clientes"
            conSubRPT = True
        End If
    Else
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        Cadselect = QuitarCaracterACadena(cadFormula, "{")
        Cadselect = QuitarCaracterACadena(Cadselect, "}")
    End If
    
    If OpcionListado = 46 Then
        'Seleccionar aquellos cliente que campo sclien.fechamov < fecha Inactividad
        If txtCodigo(31).Text <> "" Then
            campo = "sclien.fechamov"
            Devuelve = txtCodigo(31).Text
            Devuelve = "(isnull({sclien.fechamov}) or {" & campo & "} < Date(" & Year(Devuelve) & ", " & Month(Devuelve) & ", " & Day(Devuelve) & "))"
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
            Devuelve = "(" & campo & " < '" & Format(txtCodigo(31).Text, FormatoFecha) & "' OR isnull(sclien.fechamov))"
            If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Sub
            Devuelve = "pFechaMov=""" & txtCodigo(31).Text & """|"
            Cadparam = Cadparam & Devuelve
            NumParam = NumParam + 1
        End If
        
    ElseIf OpcionListado = 48 Then
        'Cadena para seleccion D/H FECHA
        '--------------------------------------------
        If txtCodigo(31).Text <> "" Or txtCodigo(32).Text <> "" Then
          'Para Crystal Report
            campo = "{sclien.fechaalt}"
            'Parametro Desde/Hasta Fecha
            Devuelve = "pDHFecha=""Fecha Alta: "
            If Not PonerDesdeHasta(campo, "F", 31, 32, Devuelve) Then Exit Sub
        End If
    End If
        
    If Not HayRegParaInforme("sclien", Cadselect) Then Exit Sub
    
    If OpcionListado = 90 Or OpcionListado = 91 Then 'OpcionListado = 90 'Etiquetas clientes
        Set frmMen = New frmMensajes
        frmMen.cadWhere = Cadselect
        frmMen.OpcionMensaje = 8 'Etiquetas clientes
        frmMen.Show vbModal
        Set frmMen = Nothing
        If Cadselect = "" Then Exit Sub
        
        If OpcionListado = 91 And Me.chkEMAIL(1).Value = 1 Then
            'Enviarlo por e-mail
            EnviarEMailMulti Cadselect, Titulo, "rFacClienCarta.rpt", "sclien" 'email para clientes
        Else
            LlamarImprimir
        End If
    Else
        LlamarImprimir
    End If
    
End Sub


Private Sub cmdAceptarCompras_Click()
'Listados de Compras
Dim campo As String
Dim cad As String
Dim Tabla As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{scafpc.codprove}"
        'Parametro Desde/Hasta Proveedor
        cad = "pDHProve=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 90, 91, cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
        'Para fam/articulo con albaranaes
        If OpcionListado = 311 And Me.chkDatosAlbaranes(1).Value = 1 Then
            campo = "{scafpa.fechaalb}"
        Else
            campo = "{scafpc.fecfactu}"
        End If
        cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 92, 93, cad) Then Exit Sub
    End If
    
    Tabla = "scafpc"
    If OpcionListado = 311 Then
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 94, 95, cad) Then Exit Sub
            
            
            If Me.chkDatosAlbaranes(1).Value = 0 Then
                Tabla = "( scafpc INNER JOIN slifpc ON scafpc.codprove=slifpc.codprove AND scafpc.numfactu=slifpc.numfactu "
                Tabla = Tabla & " AND scafpc.fecfactu=slifpc.fecfactu )"
                Tabla = Tabla & " INNER JOIN sartic ON slifpc.codartic=sartic.codartic "
        
        
            Else
                
            
            
            
            End If
        
        End If
    End If
        
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If OpcionListado = 312 Then
        'en esta opcion ver si hay albaranes
        Cadselect = Replace(Cadselect, Tabla, "scafpa")
        Cadselect = Replace(Cadselect, "fecfactu", "fechaalb")
        Tabla = "scafpa"
    End If
    
    'Para fam/articulo con albaranaes
    If OpcionListado = 311 And Me.chkDatosAlbaranes(1).Value = 1 Then
        'Es un contador de un UNION.
        'Hay que comprobar si hay reg en factuaras Y albaranes
        If Not ContadorDelUnion(False) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    
    Else
        If Not HayRegParaInforme(Tabla, Cadselect) Then
            If OpcionListado <> 312 Then Exit Sub
        
            If Not HayRegParaInforme("scaalp", Cadselect) Then Exit Sub
        End If
    End If
    
    If OpcionListado = 312 Then
    
    
        'insertar en la tmpInformes
        BorrarTempInformes
        
        'en esta opcion ver si hay albaranes
        cad = Replace(Cadselect, Tabla, "scaalp")
        cad = Replace(cad, "fecfactu", "fechaalb")
        
        'insertar los albaranes q cumple la seleccion
        If Not CargarTmpInformes_Compras_312("scaalp", cad) Then Exit Sub
        
        
        'insertar los albaranes de facturas q cumple la seleccion
        If Not CargarTmpInformes_Compras_312(Tabla, Cadselect) Then Exit Sub
        
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
    End If
    
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    conSubRPT = False
    If OpcionListado = 311 Then
        If Me.OptCompras(0).Value = True Then
            nomRPT = "rComEstProFam"
            Titulo = "Listado Compras por Familia"
            conSubRPT = True
        Else
            nomRPT = "rComEstProArt"
            Titulo = "Listado Compras por Artículo"
        End If
        
        If Me.chkDatosAlbaranes(1).Value = 1 Then
            nomRPT = nomRPT & "alb"
            Titulo = Titulo & " (con albaranes)"
            
            
            'Cambiamos los desde hasta
            'En la cadena selleccion cambiamos las tabla por
            cadFormula = Replace(cadFormula, "scafpa", "Command")
            cadFormula = Replace(cadFormula, "scafpc", "Command")
            cadFormula = Replace(cadFormula, "sartic", "Command")
            cadFormula = Replace(cadFormula, "slifpc", "Command")
            
            
            
        End If
        nomRPT = nomRPT & ".rpt"
        
        
    ElseIf OpcionListado = 310 Then
        nomRPT = "rComEstProImp.rpt"
        Titulo = "Listado Compras por Proveedor"
    Else '312: Albaranes x porveedor
        nomRPT = "rComEstProAlb.rpt"
        Titulo = "Listado albaranes por Proveedor"
    End If
    
    
    LlamarImprimir
    
    If OpcionListado = 312 Then BorrarTempInformes
End Sub

Private Sub cmdAceptarEfect_Click()
'Ofertas Efectuadas
Dim cad As String
Dim TotOfeA As String 'Nº total de Ofertas Aceptadas del Periodo( en cabecera y en historico)
Dim TotImpBA As String 'Importe Bruto total de Ofertas Aceptadas del Periodo (en cabecera e historico)
Dim TotOfeNA As String 'Nº total de Ofertas NO Aceptadas del Periodo( en cabecera y en historico)
Dim TotImpBNA As String 'Importe Bruto total de Ofertas NO Aceptadas del Periodo (en cabecera e historico)
    
    'Inicializar vbles
    InicializarVbles
    
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    
    '===================================================
    '============ PARAMETROS ===========================
    If OpcionListado = 34 Then
        'Imprimir solo las Ofertas Pendientes
        If Me.chkPendientes.Value = 1 Then
            cad = "True"
        Else
            cad = "False"
        End If
        Cadparam = Cadparam & "|pPtes=" & cad & "|"
        NumParam = NumParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacOfeEfectuadas.rpt"
        Titulo = "Ofertas Efectuadas"
    Else
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rAdmGastosTec.rpt"
        Titulo = "Gastos Técnicos"
    End If
    
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
     If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{scapre_union.fecofert}"
        Else
            Codigo = "{sgaste.fecgasto}"
        End If
        'Parametro Desde/Hasta FEcha
        cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 16, 17, cad) Then Exit Sub
    End If
    
    If OpcionListado = 34 Then
        If Me.chkPendientes.Value = 0 Then 'Se muestra resumen si SoloPEndientes=false
            Codigo = "scapre.fecofert"
            Cadselect = CadenaDesdeHastaBD(txtCodigo(16).Text, txtCodigo(17).Text, Codigo, "F")
            'Obtener total Nº Ofertas del Periodo seleccionado y
            'el total Importe Bruto de las Ofertas de Periodo seleccionado
            'y pasarlo al Informe como parametro
            If ObtenerTotalOferPeriodo(Cadselect, TotImpBA, TotImpBNA, TotOfeA, TotOfeNA) Then
                cad = "pTotPeriodoOfeA="""
                Cadparam = Cadparam & cad & TotOfeA & """|"
                cad = "pTotPeriodoOfeNA="""
                Cadparam = Cadparam & cad & TotOfeNA & """|"
                cad = "pTotPeriodoImpA="""
                Cadparam = Cadparam & cad & TotImpBA & """|"
                cad = "pTotPeriodoImpNA="""
                Cadparam = Cadparam & cad & TotImpBNA & """|"
                NumParam = NumParam + 4
            End If
        End If
    End If
    
    'Cadena para seleccion Desde y Hasta AGENTE
    '--------------------------------------------
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{scapre_union.codagent}"
        Else
            Codigo = "{sgaste.codtecni}"
        End If
        cad = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, cad) Then Exit Sub
    End If
        
    If Me.chkPendientes.visible And Me.chkPendientes.Value Then
        'Cadena para seleccion de Ofertas no Aceptadas
        Codigo = "{scapre_union.aceptado}=0"
        If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    End If

    '==============================================
    conSubRPT = False
    LlamarImprimir
End Sub


Private Sub cmdAceptarEstVentas_Click()
'Listados estadistica ventas por familia
'Listados de Compras
Dim campo As String
Dim cad As String
Dim Tabla As String


    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(96).Text <> "" Or txtCodigo(97).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 96, 97, cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    'MOdificacion  18 Novi 2008
    'Las estadisticas son sobre facturas.... Y ALBARANES!!!!
    'La fecha no se la puedo pasar porque en el union hacer referencia a dos campos
    'fecfactu(factura) y fechaalb (albaranes)
    'para ello hay un parametro en el informe
  
    If txtCodigo(98).Text <> "" Or txtCodigo(99).Text <> "" Then
        If Me.chkDatosAlbaranes(0).Value = 1 Then
            campo = "{scafac1.fechaalb}"
        Else
            campo = "{scafac.fecfactu}"
        End If
        cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 98, 99, cad) Then Exit Sub
        
        
        
            
    End If
    
    Tabla = "scafac"

    If OpcionListado = 230 Then
        campo = ""  'Para comprobar que alguno de los campos es distinto de ""
        
        
        '---------------   VENTAS x FAMILIA / ARITCULO
        If Me.chkDetallaArticulo.Value = 1 Then
            If txtCodigo(112).Text <> "" Or txtCodigo(112).Text <> "" Then
                campo = "{slifac.codArtic}"
                cad = "pDHFamilia=""Artículo: "
                If Not PonerDesdeHasta(campo, "T", 112, 113, cad) Then Exit Sub
            End If
        End If
    
    
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtCodigo(100).Text <> "" Or txtCodigo(101).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 100, 101, cad) Then Exit Sub
            
        End If
        
        
        'Si por algun campo de los de arriba es <>"" entonces tenemos que meter esto
        If campo <> "" Then
        
            If Me.chkDatosAlbaranes(0).Value = 0 Then
                'Sin albaranes
                Tabla = "( scafac INNER JOIN slifac ON scafac.codtipom=slifac.codtipom AND scafac.numfactu=slifac.numfactu "
                Tabla = Tabla & " AND scafac.fecfactu=slifac.fecfactu )"
                Tabla = Tabla & " INNER JOIN sartic ON slifac.codartic=sartic.codartic "
                
            End If
                
            
        End If
    End If
    
    
    'Para que el listado No mezcle las facturas del A o del B
    If OpcionListado = 231 Then
        campo = "{scafac.codtipom}"
        If vUsu.TrabajadorB Then
            campo = campo & " = "
        Else
            campo = campo & " <> "
        End If
        campo = campo & " 'FAZ'"
        
        If Cadselect <> "" Then Cadselect = Cadselect & " AND "
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        
        Cadselect = Cadselect & campo
        cadFormula = cadFormula & campo
        
    End If
    
    
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If Me.chkDatosAlbaranes(0).Value = 0 Then
        If Not HayRegParaInforme(Tabla, Cadselect) Then Exit Sub
    Else
        'Es un contador de un UNION
        If Not ContadorDelUnion(True) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    If OpcionListado = 230 Then
    
        If Me.chkDetallaArticulo.Value = 0 Then
            nomRPT = "rFacEstCliFam"
            Titulo = "Listado Ventas por Familia"
            conSubRPT = True
        Else
            nomRPT = "rFacEstCliFamArt"
            Titulo = "Listado ventas por familia/artículo"
            conSubRPT = False
        End If
        If Me.chkDatosAlbaranes(0).Value = 1 Then
            nomRPT = nomRPT & "Alb"
            Titulo = Titulo & "(Con albaranes)"
            
            'En la cadena seleccion cambiamos las tabla por
            'para el
            If Me.chkDetallaArticulo.Value = 1 Then
                cadFormula = Replace(cadFormula, "scafac1", "Command")
                cadFormula = Replace(cadFormula, "scafac", "Command")
                cadFormula = Replace(cadFormula, "sartic", "Command")
                cadFormula = Replace(cadFormula, "slifac", "Command")
            End If
            
        End If
        nomRPT = nomRPT & ".rpt"
    Else
    

        nomRPT = "rFacEstCliImp.rpt"
        Titulo = "Detalle Facturación Clientes"
        conSubRPT = False
    End If
    
    
    LlamarImprimir
    
End Sub

Private Function ContadorDelUnion(Compras As Boolean) As Boolean
Dim C As String

    'Con albaranes
    Codigo = Cadselect
    Codigo = QuitarCaracterACadena(Codigo, "{")
    Codigo = QuitarCaracterACadena(Codigo, "}")
    
    
    ContadorDelUnion = False
    If Compras Then
            C = "(SELECT count(*) FROM   (((`scafac1` `scafac1` INNER JOIN `scafac` `scafac` ON"
            C = C & " ((`scafac1`.`codtipom`=`scafac`.`codtipom`) AND (`scafac1`.`numfactu`=`scafac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`scafac`.`fecfactu`)) INNER JOIN `slifac` `slifac` ON"
            C = C & " ((((`scafac1`.`codtipom`=`slifac`.`codtipom`) AND (`scafac1`.`numfactu`=`slifac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`slifac`.`fecfactu`)) AND (`scafac1`.`numalbar`=`slifac`.`numalbar`))"
            C = C & " AND (`scafac1`.`codtipoa`=`slifac`.`codtipoa`)) INNER JOIN `sartic` `sartic`"
            C = C & " ON `slifac`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            
            If Codigo <> "" Then C = C & " WHERE " & Codigo
            C = C & ") + ("
            C = C & " SELECT count(*) from ((`slialb` INNER JOIN scaalb ON ((`slialb`.`codtipom`=`scaalb`.`codtipom`) AND"
            C = C & " (`slialb`.`numalbar`=`scaalb`.`numalbar`)))"
            C = C & " INNER JOIN `sartic` `sartic` ON `slialb`.`codartic`=`sartic`.`codartic`)"
            C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafac1", "scaalb")
                Codigo = Replace(Codigo, "scafac", "scaalb")
                Codigo = Replace(Codigo, "slifac", "slialb")
                
                C = C & " WHERE " & Codigo
            End If
            C = C & ")"
    
    Else
    
        'Ventas
        C = "(SELECT Count(*) from (`scafpc` `scafpc` INNER JOIN `scafpa` `scafpa`"
        C = C & " ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`=`scafpa`.`fecfactu`))"
        C = C & " AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
        C = C & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
        C = C & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
        C = C & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"
        If Codigo <> "" Then C = C & " WHERE " & Codigo
        C = C & ") + ("

        C = C & " SELECT count(*)"
        C = C & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
        C = C & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
        If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafpa", "scaalp")
                Codigo = Replace(Codigo, "scafpc", "scaalp")
                Codigo = Replace(Codigo, "slifac", "slialp")
                
                C = C & " WHERE " & Codigo
        End If
        C = C & ")"
    End If
    
    
    C = "Select " & C & " AS total"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then ContadorDelUnion = True
    End If
    miRsAux.Close
    Codigo = ""
End Function


Private Sub cmdAceptarEtiqProv_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String

    InicializarVbles
    
    'si es listado de CARTAS/eMAIL a proveedores comprobar que se ha seleccionado
    'una carta para imprimir
    If OpcionListado = 306 Then
        If txtCodigo(63).Text = "" Then
            MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
            Exit Sub
        End If
        
        'Parametro cod. carta
        Cadparam = "|pCodCarta= " & txtCodigo(63).Text & "|"
        NumParam = NumParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveCarta.rpt"
        Titulo = "Cartas a Proveedores"
        conSubRPT = True
        
    Else 'ETIQUETAS
        Cadparam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveEtiq.rpt"
        Titulo = "Etiquetas de Proveedores"
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtCodigo(58).Text <> "" Or txtCodigo(59).Text <> "" Then
        campo = "{sprove.codprove}"
        'Parametro Desde/Hasta Proveedor
        If Not PonerDesdeHasta(campo, "N", 58, 59, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H COD. POSTAL
    '--------------------------------------------
     If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
        campo = "{sprove.codpobla}"
        'Parametro Desde/Hasta cod. Postal
        If Not PonerDesdeHasta(campo, "T", 60, 61, "") Then Exit Sub
    End If
    
    '====================================================
        
        
    'Parametro a la Atencion de
    Cadparam = Cadparam & "pAtencion=""Att. " & txtCodigo(62).Text & """|"
    NumParam = NumParam + 1
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme("sprove", Cadselect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = Cadselect
    frmMen.OpcionMensaje = 9 'Etiquetas proveedores
    frmMen.Show vbModal
    Set frmMen = Nothing
    If Cadselect = "" Then Exit Sub
    
    If OpcionListado = 306 And Me.chkEMAIL(0).Value = 1 Then
        'Enviarlo por e-mail
        EnviarEMailMulti Cadselect, Titulo, "rComProveCarta.rpt", "sprove" 'email para proveedores
    Else
        LlamarImprimir
    End If
    
End Sub


Private Sub cmdAceptarFacRect_Click()
Dim cad As String
Dim TipoM As String * 3


    'Comprobar que se introdujo el motivo por el que se rectifica la factura
    If Trim(txtCodigo(87).Text) = "" Then
        MsgBox "Debe introducir el motivo de rectificación.", vbExclamation
        PonerFoco txtCodigo(87)
        Exit Sub
    End If


    TipoM = Mid(Me.cboTipoMov(0).List(Me.cboTipoMov(0).ListIndex), 1, 3)
    
    'comprobar que existe la factura en tabla "scafac"
    cad = "select count(*) from scafac where codtipom='" & TipoM & "' AND numfactu="
    cad = cad & txtCodigo(71).Text & " AND fecfactu=" & DBSet(txtCodigo(72).Text, "F")
    If RegistrosAListar(cad) = 0 Then
        cad = vbCrLf & String(40, "*") & vbCrLf
        cad = cad & vbCrLf & "No existe la factura que quiere rectificar" & vbCrLf & "¿Continuar?" & cad
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    

    'Llegado aqui pongo los datos
    'si existe devolver estos datos para recuperla en el formulario de Albaranes
    cad = TipoM & "|"
    cad = cad & txtCodigo(71).Text & "|"
    cad = cad & txtCodigo(72).Text & "|"
    cad = cad & QuitarCaracterEnter(txtCodigo(87).Text) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
    
End Sub


Private Sub cmdAceptarGenPed_Click()
'Solicitar datos para Generar Pedido a partir de una Oferta
Dim cad As String

    cad = txtCodigo(24).Text & "|"
    cad = cad & txtCodigo(25).Text & "|"
    cad = cad & txtCodigo(26).Text & "|"
    cad = cad & txtNombre(4).Text & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub cmdAceptarHco_Click()
'pedir datos para Pasar de Albaranes a historico
Dim cad As String

    'comprobar que todos los camos tienen valor
    If txtCodigo(50).Text = "" Or txtCodigo(51).Text = "" Or txtCodigo(52).Text = "" Then
        MsgBox "Debe rellenar todos los campos para pasar al histórico.", vbInformation
        Exit Sub
    End If

    'datos a devolver
    cad = txtCodigo(50).Text & "|"
    cad = cad & txtCodigo(51).Text & "|"
    cad = cad & txtCodigo(52).Text & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub cmdAceptarOfer_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim Devuelve As String, campo As String

    If txtCodigo(1).Text = "" Then 'And (txtCodigo(33).Text = "" Or txtCodigo(34).Text = "") Then
        MsgBox "Debe seleccionar una Oferta para Imprimir.", vbInformation
        PonerFoco txtCodigo(1)
        Exit Sub
    End If
    
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 31) Then
        indRPT = 5 '31: Informe de Ofertas
    ElseIf OpcionListado = 35 Then
        indRPT = 6 '35: Historico Informe de Ofertas
    End If
    conSubRPT = True
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomRPT) Then Exit Sub

    'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
    If Me.OptPapelBlanco.Value = True Then 'Blanco o con Membrete
        Devuelve = "True"
    Else
        Devuelve = "False"
    End If
    Cadparam = Cadparam & "pPapelB=" & Devuelve & "|"
    NumParam = NumParam + 1
                
    'Se pasa como parametro la carta a imprimir
    If Me.txtCodigo(2).Text <> "" Then
        Cadparam = Cadparam & "pCodCarta=" & CInt(Me.txtCodigo(2).Text) & "|"
    Else
        Cadparam = Cadparam & "pCodCarta=" & CInt(0) & "|"
    End If
    NumParam = NumParam + 1
    
    'Añadir el codigo de usuario como parametro para link con tabla Temporal en el Report
    Cadparam = Cadparam & "pCodUsu=" & vUsu.Codigo & "|"
    NumParam = NumParam + 1
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de OFERTA
    '--------------------------------------------
    If txtCodigo(1).Text <> "" Then
        Devuelve = "{" & NomTabla & ".numofert}=" & Val(txtCodigo(1).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        Cadselect = cadFormula
        
        'Si Imprimir Otras Ofertas del Cliente
        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
            campo = "{" & NomTabla & ".fecofert}"
            Devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
            If Devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & Devuelve & ")"
                Cadselect = "(" & Cadselect & " OR " & CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F") & ")"
            Else
                cadFormula = Devuelve
                Cadselect = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
            End If
            Devuelve = "{" & NomTabla & ".aceptado}=0"
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
            If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Sub
            
        ElseIf OpcionListado = 35 Then 'solo imprime la Oferta Seleccionada (si Historico filtrar x fecha)
            Devuelve = "{" & NomTabla & ".fecofert}=Date(" & Year(FecEntre) & ", " & Month(FecEntre) & ", " & Day(FecEntre) & ")"
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
            Devuelve = NomTabla & ".fecofert= '" & Format(FecEntre, FormatoFecha) & "'"
            AnyadirAFormula Cadselect, Devuelve
        End If
        'Filtrar solo las ofertas del cliente que las solicita
        If OpcionListado = 35 Then 'Historico
            Devuelve = DevuelveDesdeBDNew(conAri, NomTabla, "codclien", "numofert", txtCodigo(1).Text, "N", , "fecofert", FecEntre, "F")
        Else
            Devuelve = DevuelveDesdeBDNew(conAri, NomTabla, "codclien", "numofert", txtCodigo(1).Text, "N")
        End If
        CodClien = Devuelve
        If Devuelve <> "" Then
            campo = "{" & NomTabla & ".codclien}=" & Devuelve
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
            If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
        End If
        
    Else
'        'Comprobar si se imprimen varias Ofertas
'        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
'         'Cadena para seleccion Desde y Hasta Fecha
'         '--------------------------------------------
'            campo = "{" & NomTabla & ".fecofert}"
'            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'            If devuelve = "Error" Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, devuelve) Then
'                Exit Sub
'            Else
'                devuelve = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'                If devuelve = "Error" Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'                devuelve = "{" & NomTabla & ".aceptado=0}"
'                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'            End If
'        End If
    End If
   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente (0=NORMAL)
    Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CodClien, "N")
    If Devuelve <> "" Then
        Cadparam = Cadparam & "pTipoIVA=" & Devuelve & "|"
        NumParam = NumParam + 1
    End If

    'Cuando este cargada la tabla temporal añadir un parametro con la concatenacion de
    'Todas las ofertas que se van a imprimir
    PonerParamCadOferta Cadparam, NumParam, Cadselect
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme(NomTabla, Cadselect) Then Exit Sub
         
    LlamarImprimir
End Sub


Private Sub cmdAceptarPedCom_Click()
'55: Informe Pedido de Compras (a Proveedor)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim Devuelve As String, campo As String
Dim CodPed As String
Dim campo1 As String, campo2 As String, campo3 As String
    
    If txtCodigo(73).Text = "" Then 'Nº del Pedido
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        PonerFoco txtCodigo(73)
        Exit Sub
    Else
        NumCod = txtCodigo(73).Text
    End If
    
    If (OpcionListado = 239) And txtCodigo(76).Text = "" Then
        MsgBox "Debe seleccionar un Pedido y Fecha para Imprimir.", vbInformation
        PonerFoco txtCodigo(76)
        Exit Sub
    End If
    
    
    InicializarVbles
    conSubRPT = True
    
    '===================================================
    '============ PARAMETROS ===========================
    Select Case OpcionListado
        Case 38
            indRPT = 7 '7: Pedidos de Clientes
            Titulo = "Pedido de Ventas"
        Case 239
            indRPT = 8 '8: Pedidos de Clientes (Historico)
            Titulo = "Hist. Pedido de Venta"
        Case 55
            indRPT = 14 '14: Pedidos a Proveedores
            Titulo = "Pedidos de Compras"
        Case 56
            indRPT = 15
            Titulo = "Hist. Pedidos de Compras"
    End Select
    
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomRPT) Then Exit Sub
     
    If OpcionListado = 38 Or OpcionListado = 239 Then
        campo1 = "numpedcl"
        campo2 = "fecpedcl"
        campo3 = "codclien"
    Else
        campo1 = "numpedpr"
        campo2 = "fecpedpr"
        campo3 = "codprove"
    End If
        
        
    'VALORADO  Noviembre 2009
    NumParam = NumParam + 1
    Cadparam = Cadparam & "SinValorar= "
    If Me.chkPedidoValorado(0).Value = 1 Then
        Cadparam = Cadparam & "0"
    Else
        Cadparam = Cadparam & "1"
    End If
    Cadparam = Cadparam & "|"
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de PEDIDO
    '--------------------------------------------
    If NumCod <> "" Then
        Devuelve = "{" & NomTabla & "." & campo1 & "}=" & Val(NumCod)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        Cadselect = cadFormula
        
        If OpcionListado = 239 Then 'historico ( hay fecha)
            Devuelve = "{" & NomTabla & "." & campo2 & "}= Date(" & Year(txtCodigo(76).Text) & "," & Month(txtCodigo(76).Text) & "," & Day(txtCodigo(76).Text) & ")"
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
            Devuelve = NomTabla & "." & campo2 & "='" & Format(txtCodigo(76).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Sub
        End If
        
        'Seleccionar otros PEdidos entre esas FEchas
        If Not (txtCodigo(74).Text = "" And txtCodigo(75).Text = "") Then
            campo = "{" & NomTabla & "." & campo2 & "}"
            Devuelve = CadenaDesdeHasta(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            If Devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & Devuelve & ")"
                Cadselect = "((" & Cadselect & ") OR " & CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F") & ")"
            Else
                cadFormula = Devuelve
                Cadselect = CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            End If
        
            'Filtrar solo los Pedidos del CLIENTE/PROVEEDOR que las solicita
            If CodClien <> "" Then
                campo = "{" & NomTabla & "." & campo3 & "}=" & CodClien
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                If Not AnyadirAFormula(Cadselect, campo) Then Exit Sub
            End If
        End If
    Else
'        'Comprobar si se imprimen varios Pedidos
'        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
'         'Cadena para seleccion Desde y Hasta FECHA
'         '--------------------------------------------
'            campo = "{" & NomTabla & ".fecpedcl}"
'            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'            If devuelve = "Error" Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, devuelve) Then
'                Exit Sub
'            Else
'                devuelve = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'                If devuelve = "Error" Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'            End If
'        End If
    End If
    
    If OpcionListado = 38 Or OpcionListado = 239 Then
        'Añadimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CodClien, "N")
        If Devuelve <> "" Then
            If Devuelve = "3" Then Devuelve = "2" 'El intracom Lo trato como si fuera exento
        
            Cadparam = Cadparam & "pTipoIVA=" & Devuelve & "|"
            NumParam = NumParam + 1
        End If
    End If

    'comprobar que hay datos para mostrar en el Informe
    If Not HayRegParaInforme(NomTabla, Cadselect) Then Exit Sub
    
    LlamarImprimir
End Sub

Private Sub cmdAceptarPte_Click()
'LIstado Material Pendiente de recibir
Dim Codigo As String
Dim cad As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    'Pasar el ORDEN del informe como parametro
    If OpcionListado = 307 Then
        If Me.OptOrdenArt Then
            cad = "{slippr.codartic}"
        Else
            cad = "{scappr.numpedpr}"
        End If
        Cadparam = Cadparam & "pOrden=" & cad & "|"
        NumParam = NumParam + 1
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(65).Text <> "" Or txtCodigo(66).Text <> "" Then
        Codigo = "{scappr.codprove}"
        If OpcionListado = 308 Then Codigo = "{scaalp.codprove}"
        cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 65, 66, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(69).Text <> "" Or txtCodigo(70).Text <> "" Then
        Codigo = "{scappr.fecpedpr}"
        If OpcionListado = 308 Then Codigo = "{scaalp.fechaalb}"
        cad = "pDHFecha=""Fecha Ped.: "
        If OpcionListado = 308 Then cad = "pDHFecha=""Fecha Alb.: "
        If Not PonerDesdeHasta(Codigo, "F", 69, 70, cad) Then Exit Sub
    End If
    
    If OpcionListado = 307 Then '307: List. Materia pendiente de recibir
        'Cadena para seleccion D/H ARTICULO
        '--------------------------------------------
        If txtCodigo(67).Text <> "" Or txtCodigo(68).Text <> "" Then
            Codigo = "{slippr.codartic}"
            cad = "pDHArticulo=""Artículo: "
            If Not PonerDesdeHasta(Codigo, "T", 67, 68, cad) Then Exit Sub
        End If
    End If
    
    'Julio 2009
    'Si el usuario es de B salen los pendientes de facturar de B
    'SI no solo los de A
    If OpcionListado = 308 Then '308: List. Pendiente facturar
        cad = "{scaalp.presupuesto} = " & Abs(vUsu.TrabajadorB)
        AnyadirAFormula cadFormula, cad
        AnyadirAFormula Cadselect, cad
    End If
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If OpcionListado = 307 Then
        cad = "scappr INNER JOIN slippr ON scappr.numpedpr=slippr.numpedpr "
        Titulo = "Material Pendiente de recibir"
        nomRPT = "rComPteRecibir.rpt"
    Else
        cad = "scaalp INNER JOIN slialp ON scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Titulo = "Pendiente de Factura"
        nomRPT = "rComPteFactura.rpt"
    End If
    
    If Not HayRegParaInforme(cad, Cadselect) Then Exit Sub

    'Mostrar el Informe
    conSubRPT = False
    LlamarImprimir
End Sub


Private Sub cmdAceptarReimpFac_Click()
'Reimprimir Facturas ya contabilizadas
Dim TipoM As String * 3
'Dim TipoMh As String * 3
Dim Codigo As String
Dim b As Boolean
Dim TipoFactura As Byte

    InicializarVbles
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe
    
    'Desde/Hasta tipo movimiento
    '---------------------------------------------
    TipoM = Mid(Me.cboTipoMov(1).List(Me.cboTipoMov(1).ListIndex), 1, 3)
    If TipoM <> "" Then
        Codigo = "({scafac.codtipom}='" & TipoM & "') "
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
        Cadselect = cadFormula
'        If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    End If

    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtCodigo(83).Text <> "" Or txtCodigo(84).Text <> "" Then
        Codigo = "{scafac.numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 83, 84, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(85).Text <> "" Or txtCodigo(86).Text <> "" Then
        Codigo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "") Then Exit Sub
    End If
    
    If CBool(Me.chk_duplicado.Value) Then
        Cadparam = "pDuplicado=1|"
    Else
        Cadparam = "pDuplicado=0|"
    End If
    
    
    TipoFactura = 0
    Codigo = Mid(cboTipoMov(1).Text, 1, 3)
    If Codigo <> "" Then
        If Codigo = "FTI" Then
            TipoFactura = 1                        'Facturas ticket
        Else
            If Codigo = "FAZ" Then TipoFactura = 2 'FAacturas B
        End If
    End If
    
    
    ImprimirFacturas cadFormula, Cadparam, Cadselect, TipoFactura
    
End Sub

Private Sub cmdAceptarTrasHco_Click()
Dim Devuelve As String
Dim cad As String
'IMPRIME INFORME y DESPUES PREGUNTA SI TRASPASAR AL HISTORICO

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe y la SQL del Traspaso a Hco
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(43).Text <> "" Or txtCodigo(44).Text <> "" Then
        Codigo = "{scapre.codclien}"
        cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(Codigo, "N", 43, 44, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Or txtCodigo(46).Text <> "" Then
        Codigo = "{scapre.codagent}"
        cad = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(Codigo, "N", 45, 46, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(22).Text <> "" Or txtCodigo(23).Text <> "" Then
        Codigo = "{scapre.fecofert}"
        cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 22, 23, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta Nº OFERTA
    '---------------------------------------------
    If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
        Codigo = "{scapre.numofert}"
        cad = "pDHOferta=""Nº Oferta: "
        If Not PonerDesdeHasta(Codigo, "N", 20, 21, cad) Then Exit Sub
    End If
    
    'Seleccionar para estos criterios solo las Ofertas que no esten Aceptadas
    '------------------------------------------------------------------------
    Devuelve = " {scapre.aceptado} = 0 "
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub 'Para Crystal
    If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Sub 'Para MySQL
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme("scapre", Cadselect) Then Exit Sub

    'Mostrar el Informe
    LlamarImprimir
    
    'Preguntar si Traspasamos los Datos seleccionados al Histórico
    If MsgBox("¿Desea pasar estas Ofertas al Histórico?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        If TraspasoOfertaAHco(Cadselect) Then MsgBox "Traspaso de Ofertas a Histórico realizado correctamente. ", vbInformation
    End If
End Sub


Private Sub cmdAcetarConfirm_Click()
'Confirmacion de Pedidos
Dim Devuelve As String, campo As String

    If txtCodigo(81).Text = "" Then
        MsgBox "Debe seleccionar una carta para Imprimir la Confirmación de Pedidos.", vbInformation
        PonerFoco txtCodigo(81)
        Exit Sub
    End If
    
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    
    'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
    If Me.OptPapelBlanco3.Value = True Then
        campo = "True"
    Else
        campo = "False"
    End If
    Cadparam = Cadparam & "pPapelB=" & campo & "|"
    NumParam = NumParam + 1
                    
    'Si se impremen Saldos o no
    If Me.chkImpSaldo.Value = 1 Then
        campo = "True"
    Else
        campo = "False"
    End If
    Cadparam = Cadparam & "pImpSaldo=" & campo & "|"
    NumParam = NumParam + 1
    
                    
    'Se pasa como parametro la carta a imprimir
    If Me.txtCodigo(81).Text <> "" Then
        Cadparam = Cadparam & "pCodCarta=" & CInt(Me.txtCodigo(81).Text) & "|"
    Else
        Cadparam = Cadparam & "pCodCarta=" & CInt(0) & "|"
    End If
    NumParam = NumParam + 1
    
    'Añadir la fecha de la carta como parametro del informe
    Cadparam = Cadparam & "pFecha=""" & txtCodigo(82).Text & """|"
    NumParam = NumParam + 1
    
    'Añadir la poblacion de la empresa como parametro del informe
    Cadparam = Cadparam & "pPoblacion=""" & vParam.Poblacion & """|"
    NumParam = NumParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    nomRPT = "rFacPedConfirm.rpt"
    conSubRPT = True
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Fechas de Pedido
    '--------------------------------------------
    If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
        campo = "{" & NomTabla & ".fecpedcl}"
        If Not PonerDesdeHasta(campo, "F", 77, 78, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Clientes de Pedido
    '--------------------------------------------
    If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
        campo = "{" & NomTabla & ".codclien}"
        If Not PonerDesdeHasta(campo, "N", 79, 80, "") Then Exit Sub
    End If
       
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme(NomTabla, Cadselect) Then Exit Sub
       
    LlamarImprimir
End Sub


Private Sub cmdAcetarRecorda_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Devuelve As String
Dim bytPrecio As Byte 'Precio valoracion seleccionado
   
    'Comprobar que hay carta si vamos a imprimir un Recordatorio de Oferta
    If (OpcionListado = 32 And txtCodigo(13).Text = "") Then
        MsgBox "Debe seleccionar una carta para Imprimir el Recordatorio.", vbInformation
        PonerFoco txtCodigo(13)
        Exit Sub
    End If
    
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Pasar nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
        
    If OpcionListado = 32 Then
        indRPT = 5 'Añade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu) Then
            Exit Sub
        End If
    
        'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
        If Me.OptPapelBlancoR.Value = True Then 'Blanco o con Membrete
            Devuelve = "True"
        Else
            Devuelve = "False"
        End If
        Cadparam = Cadparam & "pPapelB=" & Devuelve & "|"
        NumParam = NumParam + 1
                    
        'Se pasa como parametro la carta a imprimir
        If Me.txtCodigo(13).Text <> "" Then
            Cadparam = Cadparam & "pCodCarta=" & CInt(Me.txtCodigo(13).Text) & "|"
        Else
            Cadparam = Cadparam & "pCodCarta=" & CInt(0) & "|"
        End If
        NumParam = NumParam + 1
        
        'Añadir las 2 lineas como parametros del informe
        If Me.txtCodigo(14).Text <> "" Then 'Linea A
            Cadparam = Cadparam & "pLineaA=""" & Me.txtCodigo(14).Text & """|"
            NumParam = NumParam + 1
        End If
        If Me.txtCodigo(15).Text <> "" Then 'Linea B
            Cadparam = Cadparam & "pLineaB=""" & Me.txtCodigo(15).Text & """|"
            NumParam = NumParam + 1
        End If
    
        'Añadir la poblacion de la empresa como parametro del informe
        Cadparam = Cadparam & "pPoblacion=""" & vParam.Poblacion & """|"
        NumParam = NumParam + 1
        nomRPT = "rFacOfeRecorda.rpt"
        
        'Nombre fichero .rpt a Imprimir
    Else
        
        indRPT = 33 'Añade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu) Then
            Exit Sub
        End If

        'nomRPT = "rFacOfeValoracion.rpt"
        nomRPT = nomDocu
        
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        Cadparam = Cadparam & "pPrecio=" & bytPrecio & "|"
        NumParam = NumParam + 1
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta CLIENTE
    '--------------------------------------------
    Codigo = "{scapre.codclien}"
    Devuelve = CadenaDesdeHasta(txtCodigo(9).Text, txtCodigo(10).Text, Codigo, "N", "Cliente")
    If Devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    'Cadena para seleccion Desde y Hasta Nº OFERTA
    '--------------------------------------------
    Codigo = "{scapre.numofert}"
    Devuelve = CadenaDesdeHasta(txtCodigo(5).Text, txtCodigo(6).Text, Codigo, "N", "Nº Oferta")
    If Devuelve = "Error" Then
        Exit Sub
    End If
    If Not AnyadirAFormula(cadFormula, Devuelve) Then
        Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    Codigo = "{scapre.fecofert}"
    Devuelve = CadenaDesdeHasta(txtCodigo(7).Text, txtCodigo(8).Text, Codigo, "F", "Fecha")
    If Devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    'Cadena para seleccion Desde y Hasta AGENTE
    '--------------------------------------------
    Codigo = "{scapre.codagent}"
    Devuelve = CadenaDesdeHasta(txtCodigo(11).Text, txtCodigo(12).Text, Codigo, "N", "Agente")
    If Devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    If OpcionListado = 32 Then
        'Cadena para seleccion de Ofertas no Aceptadas
        Codigo = "{scapre.aceptado}=0"
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    End If
    
    LlamarImprimir
End Sub


Private Sub cmdBajar_Click()
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdEnvioMail_Click()
Dim RS As ADODB.Recordset


    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    If OpcionListado = 315 Then
        If Text1(0).Text = "" Then
            MsgBox "Ponga el asunto", vbExclamation
            Exit Sub
        End If
    End If
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    
    'AHora pongo los tipo de facturas
    cadFormula = ""
    Cadselect = ""  'ME dira si estan todas o no
    For IndCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
        If Me.ListTipoMov(1000).Selected(IndCodigo) Then
            'Esta checkeado
            cadFormula = cadFormula & " OR scafac.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(IndCodigo), 1, 3)) & "'"
        Else
            Cadselect = "NO"
        End If
    Next IndCodigo
    
    If cadFormula = "" Then
        MsgBox "Seleccione algun tipo de factura", vbExclamation
        Exit Sub
    Else
        cadFormula = Mid(cadFormula, 4)
    End If
    If Cadselect = "" Then
        'Significa que estan todos. No tiene sentido poner que codtipo='fr or codtipo='FT  ESTAN TODAS
        cadFormula = " scafac.codtipom <> 'FTI'"
    End If
    'En notabla tendre

    NomTabla = "(" & cadFormula & ")"

    InicializarVbles
    cadFormula = ""
    Cadselect = ""
    If txtCodigo(110).Text <> "" Or txtCodigo(111).Text <> "" Then
        Codigo = "scafac.codclien"
        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
    End If
    
    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
        Codigo = "scafac.fecfactu"
        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
    End If
    
    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
        Codigo = "scafac.numfactu"
        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
    End If
        
        
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
    If Cadselect <> "" Then Cadselect = Cadselect & " AND "
    Cadselect = Cadselect & NomTabla
    Cadselect = " WHERE " & Cadselect
    
    Set RS = New ADODB.Recordset
    DoEvents
    
    
        
    'Ahora insertare en la tabla temporal tminformes las facturas que voy a generar pdf
    Codigo = "insert into tmpnlotes (codusu,numalbar,codprove,codartic,numlinea,fechaalb,codalmac,cantidad) "
    Codigo = Codigo & " values ( " & vUsu.Codigo & ",'"
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass

    
    
    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select codtipom,numfactu,codclien,fecfactu,totalfac from scafac  " & Cadselect
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY codtipom"
    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not RS.EOF
        NomTabla = RS!Codtipom & "'," & RS!CodClien & "," & RS!NumFactu & "," & CStr(RS!NumFactu Mod 32000) & ",'" & Format(RS!FecFactu, FormatoFecha)
        
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        NomTabla = NomTabla & "',12," & TransformaComasPuntos(CStr(DBLet(RS!TotalFac, "N"))) & ")"
        conn.Execute Codigo & NomTabla
        NumRegElim = NumRegElim + 1
        RS.MoveNext
    Wend
    RS.Close
    
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------------------------------
    '
    'Ahora cojemos las facturas que son FVA pero tienen numero terminal. COn el desde /hasta seleccionado
    'MIRAMOS en la tabla scafac1
    
    'Compruebo si tiene codclien
    NomTabla = "select scafac1.* from scafac1 ,scafac where scafac1.codtipom=scafac.codtipom and scafac1.numfactu=scafac.numfactu and scafac1.fecfactu =scafac.fecfactu"
    'NomTabla = "Select codtipom,numfactu,fecfactu from scafac1   " & cadSelect
    'El cad select LLEVA el where.  Se lo quito
    Cadselect = Mid(Cadselect, 7)
    NomTabla = NomTabla & " AND " & Cadselect & "  AND numtermi>=0  "
    
    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        NomTabla = "numalbar = '" & RS!Codtipom & "' AND fechaalb = '" & Format(RS!FecFactu, FormatoFecha) & "' AND numlinea = " & CStr(RS!NumFactu Mod 32000)
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        NomTabla = "UPDATE tmpnlotes SET codalmac = 18 WHERE codusu = " & vUsu.Codigo & " AND " & NomTabla
        conn.Execute NomTabla
    
    
        RS.MoveNext
    Wend
    RS.Close
    'Numero de registros
    NomTabla = NumRegElim
    
    'AHora ya tengo todos los datos de las facturas que voy  a imprimir
    'Entonces copruebo si para los clientes si tienen puesto el campo mail o no
    If OpcionListado = 315 Then
        If optEnvioMail(0).Value Then
            'Selecciona mail comercial
            Cadselect = "2"  'de maiclie2
        Else
            Cadselect = "1"  'de maiclie1
        End If
        Cadselect = "Select codclien,maiclie" & Cadselect
        Cadselect = Cadselect & " as email from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
        Cadselect = Cadselect & " group by codclien having email is null"
        RS.Open Cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not RS.EOF
            NumRegElim = NumRegElim + 1
            RS.MoveNext
        Wend
        RS.Close
        
        If NumRegElim > 0 Then
            If MsgBox("Tiene cliente sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                
            'Si no salimos borramos
            RS.Open Cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cadselect = "DELETE from tmpnlotes where codusu =" & vUsu.Codigo & " and codprove ="
            While Not RS.EOF
                conn.Execute Cadselect & RS!CodClien
                RS.MoveNext
            Wend
            RS.Close
        End If
    End If
    
    Cadselect = "Select count(*) from tmpnlotes where codusu =" & vUsu.Codigo
    RS.Open Cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then NumRegElim = DBLet(RS.Fields(0), "N")
        
    End If
    RS.Close
    
    If NumRegElim = 0 Then
        'NO hay datos para enviar
        
        Screen.MousePointer = vbDefault
        MsgBox "No hay datos para realizar el proceso", vbExclamation
        Exit Sub
    Else
        If OpcionListado = 315 Then
            Cadselect = "enviar por email"
        Else
            Cadselect = "exportar PDF"
        End If
            
        Cadselect = "Hay " & NumRegElim & " facturas para " & Cadselect & vbCrLf & "¿Continuar?"
        If MsgBox(Cadselect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
    End If
    If NumRegElim = 0 Then
        Set RS = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NomTabla = NumRegElim


        
    
    
        
    PonerTamnyosMail True
    frmppal.visible = False
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label4(22).Caption = "Preparando datos"
    Me.ProgressBar1.Max = CInt(NomTabla)
    Me.ProgressBar1.Value = 0
    
    
    
    NumRegElim = 0
    If GeneracionEnvioMail(RS) Then NumRegElim = 1
        
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
    
        If OpcionListado = 315 Then
    
            'Procederemos a enviarlos por mail
            If optEnvioMail(0).Value Then
                'Selecciona mail comercial
                Cadselect = "2"  'de maiclie2
            Else
                Cadselect = "1"  'de maiclie1
            End If
            Cadselect = "Select nomclien,maiclie" & Cadselect
            Cadselect = Cadselect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
    '        cadSelect = cadSelect & " group by codclien having email is null"
    
            
            frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(chkMail.Value) & "|" & Cadselect & "|"
            frmEMail.Opcion = 4 'Multienvio de facturacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBar1.visible = False
            Me.Refresh
            DoEvents
            Espera 1
            PrepararCarpetasEnvioMail
            Me.ProgressBar1.visible = True
            
        Else
            MsgBox "Proceso finalizado con exito. Archivos generados en : " & App.Path & "\Temp", vbInformation
        End If
    End If
    
    
    
    
    'Es para evitar la cantidad de pantallas abriendose y cerrandose
    Me.visible = False
    PonerTamnyosMail False
    Espera 1
    Unload Me
    frmppal.Show

    Screen.MousePointer = vbDefault
End Sub
        
        
        
Private Function GeneracionEnvioMail(ByRef RS As ADODB.Recordset) As Boolean

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    
    Cadselect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    RS.Open Cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not RS.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & RS!NumAlbar & " " & RS!codartic
        Label14(22).Refresh
        
        If CodClien <> RS!codAlmac Then   'If CodClien <> RS!codTipoM Then
            'OTRO TIPO DE DOCUMENTO
            
            '''''If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
            If Not PonerParamRPT(RS!codAlmac, Cadparam, NumParam, NumCod) Then
                Exit Function
            End If
            CodClien = RS!codAlmac
        End If
        cadFormula = "({scafac.codtipom}='" & RS!NumAlbar & "') "
        cadFormula = cadFormula & " AND ({scafac.numfactu}=" & RS!codartic & ") "
        cadFormula = cadFormula & " AND ({scafac.fecfactu}= Date(" & Year(RS!FechaAlb) & "," & Month(RS!FechaAlb) & "," & Day(RS!FechaAlb) & "))"


          
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = Cadparam
            .NumeroParametros = NumParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = NumCod
            .Opcion = 53
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        If (Me.ProgressBar1.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            Espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        If OpcionListado = 316 Then
            FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Format(RS!codProve, "000000") & "_" & RS!NumAlbar & Format(RS!codartic, "0000000") & ".pdf"
        Else
            FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codartic, "0000000") & ".pdf"
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 31, 35 '31: Informe Ofertas
                        '35: Informe Historico Ofertas
                PonerFoco txtCodigo(1)
            Case 32, 33 '32: Recordatorio de Oferta
                        '33: Informe Valoracion de Oferta
                PonerFoco txtCodigo(5)
            Case 34, 92 '34: Informe Ofertas Efectuadas
                        '92: Informe Gastos técnicos
                PonerFoco txtCodigo(16)
            Case 36 '36: Traspaso Ofertas a Historico
                PonerFoco txtCodigo(43)
            Case 37 '37: Generar Pedido de OFerta
                PonerFoco txtCodigo(24)
            Case 40 '40: Carta Confirmacion de Pedido
                PonerFoco txtCodigo(77)
            Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '48: Informe de Altas de Nuevos Clientes
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
                PonerFoco txtCodigo(27)
            Case 47 '47: Informe de Clientes
                PonerFoco txtCodigo(33)
            Case 38, 239, 55, 56 '55: Informe de Pedido de Compras (proveedor)
                PonerFoco txtCodigo(73)
            Case 57 '57: Pasar Pedido a Albaran de Compras(Proveedores)
                If Me.txtNombre(47).Text = "" Then
                    PonerFoco txtCodigo(47)
                Else
                    PonerFoco txtCodigo(48)
                End If
            Case 80, 81 '80: Pasar albaranes al historico (ventas clientes)
                            '81: Pasar pedidos al historico (ventas clientes)
                PonerFoco txtCodigo(50)
            
            Case 225 'Datos para Factura Rectificativa
                PonerFoco txtCodigo(71)
            Case 226 'Datos para Reimprimir Facturas
                PonerFocoCbo Me.cboTipoMov(1)
                
            Case 230 'Listado Ventas por Familia
                PonerFoco txtCodigo(96)
                
            Case 240 'Inf. Cierre caja TPV
                PonerFoco txtCodigo(88)
                
            Case 305, 306 '305: Listado Etiquetas proveedor
                          '306: Listado Cartas a proveedores
                PonerFoco txtCodigo(58)
            Case 307, 308 '307: List. Pendiente de Recibir (COMPRAS)
                          '308: List. Pendiente de Facturar (COMPRAS)
                PonerFoco txtCodigo(65)
                
            Case 310, 311, 312 'Listado Compras por Proveedor/Familia/Articulo
                                '312: Listado albaranes por proveedor
                PonerFoco txtCodigo(90)
            Case 315
              
                PonerFoco txtCodigo(110)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim Devuelve As String
    
    'Icono del formulario
    Me.Icon = frmppal.Icon

    PrimeraVez = True
    limpiar Me
    IndCodigo = 0
    NomTabla = ""

    'Ocultar todos los Frames de Formulario
    Me.FrameOfertas.visible = False
    Me.FrameRecordatorio.visible = False
    Me.FrameEfectuadas.visible = False
    Me.FrameTraspasoHco.visible = False
    Me.FrameGenPedido.visible = False
    Me.FrameClienInactivos.visible = False
    Me.FrameClientes.visible = False
    Me.FrameGenAlbCom.visible = False
    Me.FramePasarHco.visible = False
    Me.FrameEtiqProv.visible = False
    Me.FramePteRecibir.visible = False
    Me.FrameFacRectif.visible = False
    Me.FrameFacReimprimir.visible = False
    Me.FramePedidos.visible = False
    Me.FrameConfirmPed.visible = False
    Me.FrameCierreCaja.visible = False
    Me.FrameCompras.visible = False
    Me.FrameEstVentasFam.visible = False
    FrameEnvioFacMail.visible = False
    CommitConexion
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 31, 35 '31: Informe de Ofertas
                    '35: Informe Historico de Ofertas
            W = 6075
            H = 5655
            PonerFrameVisible Me.FrameOfertas, True, H, W
            Me.OptPapelBlanco.Value = True
            indFrame = 0
            If NumCod <> "" Then txtCodigo(1).Text = NumCod
            If OpcionListado = 35 Then Me.Label5.Caption = "Informe de Ofertas (Histórico)"
            
        Case 32, 33 '32: Recordatorio de Ofertas
                    '33:Informe Valoración de Ofertas
            PonerFrameRecordaVisible True, H, W
            indFrame = 1
            If CodClien <> "" Then
                txtCodigo(9).Text = CodClien
                txtCodigo(10).Text = CodClien
                Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", CodClien, "N")
                txtNombre(9).Text = Devuelve
                txtNombre(10).Text = Devuelve
            End If
            If NumCod <> "" Then
                txtCodigo(5).Text = NumCod
                txtCodigo(6).Text = NumCod
            End If
            
        Case 34, 92 '34: Informe Ofertas Efectuadas
                    '92: Informe Gastos Técnicos
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameEfectuadas, True, H, W
            If OpcionListado = 92 Then
                Label1.Caption = "Gastos Técnicos"
                Label4(4).Caption = "Técnico"
            End If
            Me.chkPendientes.visible = (OpcionListado = 34)
            indFrame = 2
            
        Case 36 '36: Traspaso a Historico (IMPRIME LISTADO Y PREGUNTA SI TRASPASO A HCO)
            W = 6815
            H = 5455
            PonerFrameVisible Me.FrameTraspasoHco, True, H, W
            indFrame = 3
            Me.Caption = "Ofertas"
            
        Case 37 '37: Pedir datos para pasar Oferta a Pedido (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenPedido, True, H, W
            indFrame = 4
            Me.Caption = "Generar Pedido"
            txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(26).Text = Format(FecEntre, "dd/mm/yyyy")
            txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
        
        
         Case 40 '40: Cartas Confirmacion de Pedidos
            W = 7035
            H = 6255
            PonerFrameVisible Me.FrameConfirmPed, True, H, W
            Me.OptPapelBlanco3.Value = True
            indFrame = 13 'solo para el boton cancelar
            txtCodigo(82).Text = Format(Now, "dd/mm/yy")
            NomTabla = "scaped"
            NomTablaLin = "sliped"
        
        Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
            PonerFrameClienInacVisible True, H, W
            indFrame = 5
            If OpcionListado = 90 Then
                CargarComboTipoMov 2
                FrameImpClien.visible = False
            End If
        Case 47 '47: Informe de Clientes
            W = 8960
            H = 6020
            PonerFrameVisible Me.FrameClientes, True, H, W
            CargarListViewOrden
            indFrame = 6
            
        Case 38, 239, 55, 56
                '38: Pedidos Venta
                '55: Informe de Pedido de Compras (Proveedor)
                '56: Informe de Hist. Pedido de Compras (Proveedor)
            PonerFramePedVisible H, W
            indFrame = 12
            If NumCod <> "" Then txtCodigo(73).Text = NumCod
            
            
            
        Case 57 '57: Pedir datos para pasar de Pedido a Albaran (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenAlbCom, True, H, W
            indFrame = 7
            Me.Caption = "Generar Albaran Compras"
            'Poner el trabajador conectado
            Me.txtCodigo(47).Text = PonerTrabajadorConectado(Devuelve)
            Me.txtNombre(47).Text = Devuelve
            Me.txtCodigo(49).Text = Format(Now, "dd/mm/yyyy")
        
            
        
        Case 80, 81 '80: pasar albaranes al historico (ventas)
                        '81: pasar pedidos al historico (ventas)
            H = 4575
            W = 6920
            PonerFrameVisible Me.FramePasarHco, True, H, W
            indFrame = 8
            Me.Caption = "Eliminar"
            Select Case OpcionListado
                Case 80, 82: Me.Label3(4).Caption = "Pasar Albaran al histórico"
                Case 81: Me.Label3(4).Caption = "Pasar Pedido al histórico"
            End Select
            Me.txtCodigo(50).Text = Format(Now, "dd/mm/yyyy")
            Me.txtCodigo(51).Text = PonerTrabajadorConectado(Devuelve)
            Me.txtNombre(51).Text = Devuelve
            
        Case 225 'Factura rectificativa
            H = 4420
            W = 5740
            PonerFrameVisible Me.FrameFacRectif, True, H, W
            indFrame = 11
            Me.Caption = "Facturas rectificativas"
            CargarComboTipoMov (0)
'            Me.cboTipomov(0).ListIndex = 2
            
        Case 226 'Reimprimir Factura
            H = 4455
            W = 6555
            PonerFrameVisible Me.FrameFacReimprimir, True, H, W
            indFrame = 14
            CargarComboTipoMov (1)
            
            
            cadFormula = DevuelveDesdeBDNew(conAri, "scryst", "nomcryst", "codcryst", "18", "N")
            Me.chkFormatoTPV.Value = 0
            If cadFormula = "" Then
                'NO SE HA ENCONTRADOR
                Me.chkFormatoTPV.Enabled = False
                cadFormula = "Formato NO encontrado"
            End If
            Me.chkFormatoTPV.Caption = cadFormula
            
'            CargarComboTipoMov (2)
            
        Case 230, 231 '230: Estadistica ventas por familia
                      '231: Detalle facturacion clientes
            indFrame = 17
            H = 5805
            If OpcionListado = 231 Then
                H = 4325
                Me.cmdAceptarEstVentas.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarEstVentas.Top
                Me.Label9(31).Caption = "Detalle Facturación Clientes"
            End If
            W = 7035
            Me.Frame12.visible = (OpcionListado = 230)
            PonerFrameVisible Me.FrameEstVentasFam, True, H, W
            
           
        
        Case 240 'Inf. cierre caja TPV
            H = 3800
            W = 6300
            PonerFrameVisible Me.FrameCierreCaja, True, H, W
            indFrame = 15
'            CargarComboTipoPago
'            Combo1.ListIndex = 0
            'Mostrar la fecha de hoy
            txtCodigo(88).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(89).Text = Format(Now, "dd/mm/yyyy")
            
        
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = 5325
            W = 7035
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (OpcionListado = 306)
            If (OpcionListado = 306) Then Me.Label9(1).Caption = "Cartas a Proveedores"
            
        Case 307, 308 '307: List. Material Pendiente de recibir (COMPRAS)
                      '308: List. Albaranes ptes de facturar (COMPRAS)
            indFrame = 10
            If OpcionListado = 307 Then
                Me.Label9(19).Caption = "Material pendiente de recibir"
                H = 5200
            Else
                Me.Label9(19).Caption = "Albaranes pendiente de factura"
                H = 4200
                Me.cmdAceptarPte.Top = 3500
                Me.cmdCancel(10).Top = Me.cmdAceptarPte.Top
            End If
            W = 7035
            PonerFrameVisible Me.FramePteRecibir, True, H, W
            Me.Frame6.visible = (OpcionListado = 307)
            Me.Frame7.visible = (OpcionListado = 307)
            
        Case 310, 311, 312 '310: Listado COMPRAS por proveedor
                            '312: Listado albaranes por proveedor
            indFrame = 16
            H = 5235
            If OpcionListado = 310 Or OpcionListado = 312 Then
                H = 4325
                Me.cmdAceptarCompras.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarCompras.Top
                If OpcionListado = 312 Then
                    Me.Label9(21).Caption = "Albaranes por Proveedor"
                Else
                    Me.Label9(21).Caption = "Compras por Proveedor"
                End If
                Me.Label4(87).Caption = "Fecha albaran"
            End If
            W = 7035
            
            PonerFrameVisible Me.FrameCompras, True, H, W
            Me.Frame8.visible = (OpcionListado = 311)
            Me.Frame9.visible = (OpcionListado = 311)
            chkDatosAlbaranes(1).visible = (OpcionListado = 311)
        Case 315, 316
            indFrame = 18
            
            PonerTamañosEmailPDF OpcionListado = 315
            
            H = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, H, W
            
            CargarComboTipoMov 1000
            
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    'Poner la tabla de Ofertas o la del Historico de Ofertas
    If NomTabla = "" Then
        If OpcionListado = 35 Then
            NomTabla = "schpre" 'Historico
        Else
            NomTabla = "scapre"
        End If
    End If
End Sub



Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de cod Postal
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If OpcionListado = 305 Or OpcionListado = 306 Then 'Proveedores
            cadFormula = "{sprove.codprove} IN [" & CadenaSeleccion & "]"
            Cadselect = "sprove.codprove IN (" & CadenaSeleccion & ")"
        Else 'clientes
            cadFormula = "{sclien.codclien} IN [" & CadenaSeleccion & "]"
            Cadselect = "sclien.codclien IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        Cadselect = ""
    End If
End Sub

Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Actividades
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArtic_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProve_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoRuta_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Rutas
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSitua_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoZona_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Zonas
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 39, 40, 45 'Cod. Carta
            Select Case Index
                Case 0: IndCodigo = 2
                Case 1: IndCodigo = 13
                Case 39: IndCodigo = 63
                Case 40: IndCodigo = 64
                Case 45: IndCodigo = 81
            End Select
            
            Set frmMtoCartasOfe = New frmFacCartasOferta
            frmMtoCartasOfe.DatosADevolverBusqueda = "0|1|"
            frmMtoCartasOfe.Show vbModal
            Set frmMtoCartasOfe = Nothing
            
        Case 2, 3, 9, 10, 23, 24, 46, 47, 52, 53, 56, 57 'Cod. CLIENTE
            Select Case Index
                Case 2, 3: IndCodigo = 7 + Index
                Case 9, 10: IndCodigo = 18 + Index
                Case 23, 24: IndCodigo = Index + 20
                Case 46, 47: IndCodigo = Index + 33
                Case 52, 53: IndCodigo = Index + 44
                Case 56, 57: IndCodigo = Index + 54
            End Select
            Set frmMtoCliente = New frmFacClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|1|"
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
            
        Case 4, 5, 6, 7, 11, 12, 19, 20, 25, 26 'Cod. AGENTE
            Select Case Index
                Case 4, 5: IndCodigo = 7 + Index
                Case 5: IndCodigo = 12
                Case 6, 7: IndCodigo = 12 + Index
                Case 11, 12: IndCodigo = 18 + Index
                Case 19, 20, 25, 26: IndCodigo = 20 + Index
            End Select
            If OpcionListado <> 92 Then
                Set frmMtoAgente = New frmFacAgentesCom
                frmMtoAgente.DatosADevolverBusqueda = "0|1|"
                frmMtoAgente.Show vbModal
                Set frmMtoAgente = Nothing
            ElseIf Index = 6 Or Index = 7 Then 'Gastos financieros (trabajador)
                Set frmMtoTraba = New frmAdmTrabajadores
                frmMtoTraba.DatosADevolverBusqueda = "0|1|"
                frmMtoTraba.Show vbModal
                Set frmMtoTraba = Nothing
            End If
            
        Case 8, 28 'cod. TRABAJADOR
            IndCodigo = 24
            If Index = 28 Then IndCodigo = 51
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 13, 14, 30, 31 'cod. ACTIVIDAD
            IndCodigo = 20 + Index
            If Index = 30 Or Index = 31 Then IndCodigo = Index + 23
            Set frmMtoActiv = New frmFacActividades
            frmMtoActiv.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(IndCodigo).Text) Then txtCodigo(IndCodigo).Text = ""
            frmMtoActiv.Show vbModal
            Set frmMtoActiv = Nothing
            
        Case 15, 16 'cod. ZONA
            IndCodigo = 20 + Index
            Set frmMtoZona = New frmFacZonas
            frmMtoZona.DatosADevolverBusqueda = "0|1|"
            frmMtoZona.Show vbModal
            Set frmMtoZona = Nothing
            
         Case 17, 18 'cod. RUTA
            IndCodigo = 20 + Index
            Set frmMtoRuta = New frmFacRutas
            frmMtoRuta.DatosADevolverBusqueda = "0|1|"
            frmMtoRuta.Show vbModal
            Set frmMtoRuta = Nothing
            
        Case 21, 22, 34 'cod. SITUACION
            IndCodigo = 20 + Index
            If Index = 34 Then IndCodigo = Index + 23
            Set frmMtoSitua = New frmFacSituaciones
            frmMtoSitua.DatosADevolverBusqueda = "0|1|"
            frmMtoSitua.Show vbModal
            Set frmMtoSitua = Nothing
            
        Case 29 'INCIDENCIAS
            IndCodigo = 52
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            txtCodigo(IndCodigo).Text = ""
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 32, 33, 37, 38 'Cod POSTAL
            IndCodigo = Index + 23
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0|1|"
            txtCodigo(IndCodigo).Text = ""
            frmCP.Show vbModal
            Set frmCP = Nothing
            
        Case 35, 36, 41, 42, 48, 49 'cod. PROVEEDOR
            Select Case Index
                Case 35, 36: IndCodigo = Index + 23
                Case 41, 42: IndCodigo = Index + 24
                Case 48, 49: IndCodigo = Index + 42
            End Select
'            If Index = 35 Or Index = 36 Then indCodigo = Index + 23
'            If Index = 41 Or Index = 42 Then indCodigo = Index + 24
'            If Index = 48 Or Index = 49 Then indCodigo = Index + 42
            Set frmMtoProve = New frmComProveedores
            frmMtoProve.DatosADevolverBusqueda = "0|1|"
            frmMtoProve.Show vbModal
            Set frmMtoProve = Nothing
            
        Case 43, 44, 58, 59 'cod. ARTICULO
            If Index <= 44 Then
                IndCodigo = Index + 24
            Else
                IndCodigo = Index + 54  'En listado de vetnas x familia articulo
            End If
            Set frmMtoArtic = New frmAlmArticulos
            frmMtoArtic.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
            frmMtoArtic.Show vbModal
            Set frmMtoArtic = Nothing
            
        Case 50, 51, 54, 55 'Cod. FAMILIA articulo
            Select Case Index
                Case 50, 51: IndCodigo = Index + 44
                Case 54, 55: IndCodigo = Index + 46
            End Select
            Set frmMtoFamilia = New frmAlmFamiliaArticulo
            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
    End Select
    PonerFoco txtCodigo(IndCodigo)
End Sub


Private Sub imgClearCmbTipomov_Click()
    cboTipoMov(2).ListIndex = -1
End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            IndCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            IndCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            IndCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            IndCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            IndCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            IndCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            IndCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            IndCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            IndCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            IndCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            IndCodigo = 49
        Case 14
            IndCodigo = 50
        Case 15, 16
            IndCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            IndCodigo = 72
        Case 18, 19 'Ped. Compras
            IndCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            IndCodigo = Index + 57
        Case 22: IndCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            IndCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            IndCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            IndCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            IndCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            IndCodigo = Index + 73
        Case 33, 34
            IndCodigo = Index + 75
   End Select
   
   
   PonerFormatoFecha txtCodigo(IndCodigo)
   If txtCodigo(IndCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(IndCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(IndCodigo)
End Sub












Private Sub ListTipoMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optEnvioMail_Click(Index As Integer)
    If Index > 1 Then PonerTamañosEmailPDF Index = 2
End Sub

Private Sub optEnvioMail_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    
    
    
        
    
End Sub

Private Sub optForpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 1 Then KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 3, 4, 7, 8, 16, 17, 22, 23, 25, 26, 31, 32, 49, 50, 69, 70, 72, 74, 75, 77, 78, 82, 85, 86, 88, 89, 92, 93, 98, 99, 104, 105, 108, 109
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
            'Fecha entrega para Pedido. Poner la semana
            If Index = 26 Then
                'Comprobar que fecha entrega es posterior a la del pedido
                If Not EsFechaIgualPosterior(txtCodigo(25).Text, txtCodigo(26).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                Else
                    txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
                End If
            End If
            
        Case 1, 5, 6, 20, 21, 71, 83, 84 'Nº de OFERTA/FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
        
        Case 2, 13, 63, 64, 81 'CARTA de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codCampo = "codcarta"
            nomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 9, 10, 27, 28, 43, 44, 79, 80, 96, 97, 110, 111 'Cod. CLIENTE
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            nomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"

        Case 11, 12, 18, 19, 29, 30, 39, 40, 45, 46 'Cod. AGENTE
            EsNomCod = True
            Formato = "0000"
            If OpcionListado = 92 Then 'Gastos tecnicos
                If Index = 18 Or Index = 19 Then
                    'cod agente / cod. trabajador
                    Tabla = "straba"
                    codCampo = "codtraba"
                    nomCampo = "nomtraba"
                    Titulo = "Trabajador"
                End If
            Else
                Tabla = "sagent"
                codCampo = "codagent"
                nomCampo = "nomagent"
                Titulo = "Agente"
            End If
        
        Case 24, 47, 51 'Cod. TRABAJADOR
            EsNomCod = True
            Tabla = "straba"
            codCampo = "codtraba"
            nomCampo = "nomtraba"
            Formato = "0000"
            Titulo = "Trabajador"
            
        Case 33, 34, 53, 54 'Cod ACTIVIDAD
            EsNomCod = True
            Tabla = "sactiv"
            codCampo = "codactiv"
            nomCampo = "nomactiv"
            Formato = "000"
            Titulo = "Actividad de Cliente"
            
        Case 35, 36 'cod ZONA
            EsNomCod = True
            Tabla = "szonas"
            codCampo = "codzonas"
            nomCampo = "nomzonas"
            Formato = "000"
            Titulo = "Zona de Cliente"
            
        Case 37, 38 'cod RUTA
            EsNomCod = True
            Tabla = "srutas"
            codCampo = "codrutas"
            nomCampo = "nomrutas"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
                        
        Case 41, 42, 57 'cod SITUACION
            EsNomCod = True
            Tabla = "ssitua"
            codCampo = "codsitua"
            nomCampo = "nomsitua"
            Formato = "00"
            Titulo = "Situación Especial"
            
        Case 52 'cod. Incidencias
            EsNomCod = True
            Tabla = "sincid"
            codCampo = "codincid"
            nomCampo = "nomincid"
            TipCampo = "T"
            Titulo = "Incidencias"
            
        Case 55, 56, 60, 61 'cod POSTAL
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scpostal", "provincia", "cpostal", "CPostal")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 58, 59, 65, 66, 90, 91 'Cod. PROVEEDOR
            EsNomCod = True
            Tabla = "sprove"
            codCampo = "codprove"
            nomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
            
        Case 67, 68, 112, 113 'cod. ARTICULO
            EsNomCod = True
            Tabla = "sartic"
            codCampo = "codartic"
            nomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            
        Case 73  'Nº de Pedido de Compras
            If txtCodigo(Index).Text = "" Then Exit Sub
            If OpcionListado = 55 Or OpcionListado = 56 Then
                nomCampo = "numpedpr"
                Titulo = "Proveedor"
            Else
                nomCampo = "numpedcl"
                Titulo = "Cliente"
            End If
            nomCampo = DevuelveDesdeBDNew(conAri, NomTabla, nomCampo, nomCampo, txtCodigo(Index).Text, "N")
            If nomCampo = "" Then
                MsgBox "No existe el Nº de Pedido de " & Titulo & ": " & txtCodigo(Index).Text, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 94, 95, 100, 101 'cod. FAMILIA articulos
            EsNomCod = True
            Tabla = "sfamia"
            codCampo = "codfamia"
            nomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        Case 107
            'Cuando pierda el foco, si estamos Exportando facturas, pasamos el foco al btn
            PonerFocoBtn cmdEnvioMail
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, Titulo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, nomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
End Sub




Private Sub PonerFrameRecordaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Ofertas Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Ofertas
Dim b As Boolean

    H = 7100
    W = 7100
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FrameRecordatorio, visible, H, W

    If visible = True Then
        '====================================
        Me.OptPapelBlancoR.Value = True

        b = (OpcionListado = 32) '32: Informe Recordatorio
                                 '33: Informe Valoracion Ofertas
        'Carta
        Me.Label4(24).visible = b
        Me.imgBuscarOfer(1).visible = b
        txtCodigo(13).visible = b
        txtNombre(13).visible = b
        'Lineas
        Me.Label4(0).visible = b
        txtCodigo(14).visible = b
        txtCodigo(15).visible = b
        'Pedir Tipo Papel (blanco o con membrete)
        Me.FrameTipoPapel2.visible = b

        'Frame Valorar coste con
        Me.FrameValorar.visible = Not b
        If Not b Then
            Me.FrameValorar.Top = 4520
            Me.FrameValorar.Left = 600
            Me.FrameRecordatorio.Width = 6800
            W = Me.FrameRecordatorio.Width
        End If

        'Poner el Titulo del Frame
        If b Then
            Me.Label7.Caption = "Recordatorio de Ofertas"
        Else
            Me.Label7.Caption = "Valoración de Ofertas"
        End If
    End If
End Sub

   
Private Sub PonerFrameClienInacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Clientes Inactivos Visible y Ajustado al Formulario, y visualiza los controles
'necesarios
Dim b As Boolean

    If OpcionListado = 90 Or OpcionListado = 91 Then
        H = 6980
        Me.cmdAceptarClienInac.Top = 5980
        Me.cmdCancel(5).Top = 5980
    Else
        H = 4460
        Me.cmdAceptarClienInac.Top = 3800
        Me.cmdCancel(5).Top = 3800
    End If
    Me.frameCliexFacturas.visible = OpcionListado = 90
    
    If OpcionListado = 90 Or OpcionListado = 91 Then
        W = 11000
    Else
        W = 6800
    End If
    
    PonerFrameVisible Me.FrameClienInactivos, visible, H, W

    If visible = True Then
        b = (OpcionListado = 48)
        'Mostrar D/H Fecha
        Label4(43).visible = b
        Label4(44).visible = b
        Me.imgFecha(12).visible = b
        Me.txtCodigo(32).visible = b
        
        If b Then
            Me.Label4(36).Caption = "Fecha Alta"
            Me.Label8.Caption = "Altas Nuevos Clientes"
        ElseIf OpcionListado = 90 Or OpcionListado = 91 Then
            Me.Frame1.visible = True
            Me.txtCodigo(31).visible = False
            Me.FrameImpClien.visible = True
            Me.OptCliTodos.Value = True
            If OpcionListado = 90 Then
                Me.Label8.Caption = "Etiquetas de Clientes"
                Me.FrameImpClien.Top = 5740
                Me.FrameImpClien.Left = 600
            Else
                Me.Label8.Caption = "Cartas a Clientes"
                Me.FrameImpClien.Left = 6800
                Me.FrameImpClien.Top = 4500
            End If
        End If
        Me.Frame4.visible = (OpcionListado = 91)
    End If
End Sub


Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function TraspasoOfertaAHco(cadWhere As String) As Boolean
'Realiza el traspaso de las ofertas seleccionadas por cadWhere
'Inserta en la tabla de Historico de ofertas (schpre, slhpre)
'Borra de las tablas de Ofertas (scapre, slipre)
Dim SQL As String
Dim Donde As String
Dim bol As Boolean

'Aqui empieza transaccion
    conn.BeginTrans
    On Error GoTo ETraspasoHco
    bol = ActualizarElTraspaso(Donde, cadWhere, "OFE")

ETraspasoHco:
        If Err.Number <> 0 Then
            SQL = "Traspaso Ofertas a Histórico." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            TraspasoOfertaAHco = True
        Else
            conn.RollbackTrans
            TraspasoOfertaAHco = False
        End If
End Function


Private Function ObtenerTotalOferPeriodo(cadWhere As String, TotImpA As String, TotImpNA As String, TotOfeA As String, TotOfeNA As String) As Boolean
'para INFORME DE OFERTAS EFECTUADAS
'TotImpA: suma del Importe bruto de todas las Ofertas Aceptadas del periodo seleccionado
'TotImpNA: suma del Importe bruto de todas las Ofertas NO Aceptadas del periodo
'TotOfeA: nº total de ofertas Aceptadas en el periodo
'TotOfeNA: nº total de Ofertas NO Aceptadas en el periodo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpBrutoLin As Currency
Dim ImpBrutoTotA As Currency
Dim ImpBrutoTotNA As Currency
Dim TotalOfeA As Integer
Dim TotalOfeNA As Integer
On Error GoTo ETotalPeriodo

    SQL = "SELECT scapre.numofert, scapre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal, (sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
    SQL = SQL & " FROM scapre INNER join slipre ON scapre.numofert=slipre.numofert "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " GROUP by scapre.numofert "
    SQL = SQL & " Union "
    SQL = SQL & " SELECT schpre.numofert, schpre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal,(sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
    SQL = SQL & " FROM schpre iNNER join slhpre ON schpre.numofert=slhpre.numofert "
    If cadWhere <> "" Then
'        cadWHERE = SustituirCadenas(cadWHERE, "scapre", "schpre")
        cadWhere = Replace(cadWhere, "scapre", "schpre")
        SQL = SQL & " WHERE " & cadWhere
    End If
    SQL = SQL & " GROUP by schpre.numofert "

    ImpBrutoTotA = 0
    ImpBrutoTotNA = 0
    TotalOfeA = 0
    TotalOfeNA = 0
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        ImpBrutoLin = RS!ImpTotal - RS!impdtopp - RS!impdtogn
        If RS!aceptado = 1 Then 'OFERTA ACEPTADA
            TotalOfeA = TotalOfeA + 1
            ImpBrutoTotA = ImpBrutoTotA + ImpBrutoLin
        Else 'OFERTA NO ACEPTADA
            TotalOfeNA = TotalOfeNA + 1
            ImpBrutoTotNA = ImpBrutoTotNA + ImpBrutoLin
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    TotImpA = Format(ImpBrutoTotA, "0.00")
    TotImpNA = Format(ImpBrutoTotNA, "0.00")
    TotOfeA = TotalOfeA
    TotOfeNA = TotalOfeNA
    ObtenerTotalOferPeriodo = True
    
ETotalPeriodo:
    If Err.Number <> 0 Then ObtenerTotalOferPeriodo = False
End Function


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameClientes
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Actividad, Zona, Ruta, Agente, Situación
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Campo", 1500

    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Actividad"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Zona"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Ruta"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Agente"
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    Cadselect = ""
    Cadparam = ""
    NumParam = 0
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim Devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    Devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If Devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(Cadselect, cad) Then Exit Function
    End If
    
    If Devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            Cadparam = Cadparam & AnyadirParametroDH(param, indD, indH) & """|"
            NumParam = NumParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub LlamarImprimir()
     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
'Pone por que campos se van a AGrupar los datos en el Informe de Crystal Report
'El informe tiene definido 4 formulas a las cuales ahora le asignamos un campo
'de la tabla segun el orden seleccionado para el agrupamiento
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Actividad"
            Cadparam = Cadparam & campo & "{sclien.codactiv}" & "|"
            If numGrupo = 1 Then
                Cadparam = Cadparam & nomCampo & " ""ACTIVIDAD:  "" & " & " totext({sclien.codactiv},""000"") & " & """  """ & " & {sactiv.nomactiv}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codactiv},""000"") & " & """ """ & " & {sactiv.nomactiv}" & "|"
                Cadparam = Cadparam & nomCampo & "{sactiv.nomactiv}" & "|"
                Cadparam = Cadparam & "pTitulo" & numGrupo & "=""Actividad""" & "|"
                NumParam = NumParam + 1
            End If
            NumParam = NumParam + 1
            
        Case "Zona"
            Cadparam = Cadparam & campo & "{sclien.codzonas}" & "|"
            If numGrupo = 1 Then
                Cadparam = Cadparam & nomCampo & " ""ZONA:  "" & " & " totext({sclien.codzonas},""000"") & " & """  """ & " & {szonas.nomzonas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codzonas},""000"") & " & """ """ & " & {szonas.nomzonas}" & "|"
                Cadparam = Cadparam & nomCampo & "{szonas.nomzonas}" & "|"
                Cadparam = Cadparam & "pTitulo" & numGrupo & "=""Zona""" & "|"
                NumParam = NumParam + 1
            End If
            NumParam = NumParam + 1
            
        Case "Ruta"
            Cadparam = Cadparam & campo & "{sclien.codrutas}" & "|"
            If numGrupo = 1 Then
                Cadparam = Cadparam & nomCampo & " ""RUTA:  "" & " & " totext({sclien.codrutas},""000"") & " & """  """ & " & {srutas.nomrutas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codrutas},""000"") & " & """ """ & " & {srutas.nomrutas}" & "|"
                Cadparam = Cadparam & nomCampo & "{srutas.nomrutas}" & "|"
                Cadparam = Cadparam & "pTitulo" & numGrupo & "=""Ruta""" & "|"
                NumParam = NumParam + 1
            End If
            NumParam = NumParam + 1
'            PonerGrupo = numGrupo
        Case "Agente"
            Cadparam = Cadparam & campo & "{sclien.codagent}" & "|"
            If numGrupo = 1 Then
                Cadparam = Cadparam & nomCampo & " ""AGENTE:  "" & " & " totext({sclien.codagent},""000000"") & " & """  """ & " & {sagent.nomagent}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codagent},""000000"") & " & """ """ & " & {sagent.nomagent}" & "|"
                Cadparam = Cadparam & nomCampo & "{sagent.nomagent}" & "|"
                Cadparam = Cadparam & "pTitulo" & numGrupo & "=""Agente""" & "|"
                NumParam = NumParam + 1
            End If
            NumParam = NumParam + 1
'        Case "Situacion"
    End Select
End Function


Private Function ListaClientesMante(cadWhere As String) As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim SQL As String, cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista

    cad = ""
    SQL = "SELECT sclien.codclien "
    SQL = SQL & " FROM sclien INNER JOIN scaman ON sclien.codclien=scaman.codclien "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not RS.EOF
        cad = cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    cad = Mid(cad, 1, Len(cad) - 1)
    ListaClientesMante = cad
ELista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Clientes con mantenimientos", Err.Description
End Function




Private Function ListaClientesDesdeHastaFactura() As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim SQL As String, cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista2

    'Monto el cad
    cad = ""
    If Me.cboTipoMov(2).ListIndex >= 0 Then
        'Tipo mov=
        cad = " AND codtipom = '" & Mid(Me.cboTipoMov(2).List(Me.cboTipoMov(2).ListIndex), 1, 3) & "'"
    End If
    If txtCodigo(102).Text <> "" Then cad = cad & " AND numfactu >= " & txtCodigo(102).Text
    If txtCodigo(103).Text <> "" Then cad = cad & " AND numfactu <= " & txtCodigo(103).Text
    If txtCodigo(104).Text <> "" Then cad = cad & " AND fecfactu >= '" & Format(txtCodigo(104).Text, FormatoFecha) & "'"
    If txtCodigo(105).Text <> "" Then cad = cad & " AND fecfactu <= '" & Format(txtCodigo(105).Text, FormatoFecha) & "'"
    If Len(cad) > 0 Then cad = Mid(cad, 5) 'QUITO EL PRIMER AND
    
    SQL = "SELECT DISTINCT(scafac.codclien) "
    SQL = SQL & " FROM scafac "
    If cad <> "" Then SQL = SQL & " WHERE " & cad


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not RS.EOF
        cad = cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    
    ListaClientesDesdeHastaFactura = cad
ELista2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Procedimiento: ListaClientesDesdeHastaFactura", Err.Description
End Function



Private Sub EnviarEMailMulti(cadWhere As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad1 As String, cad2 As String, Lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "sprove" Then
        'seleccionamos todos los proveedores a los que queremos enviar e-mail
        SQL = "SELECT codprove,nomprove,maiprov1,maiprov2 "
    ElseIf cadTabla = "sclien" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        SQL = "SELECT codclien,nomclien,maiclie1,maiclie2 "
    End If
    SQL = SQL & "FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    SQL = "CREATE TEMPORARY TABLE tmpMail ( "
    SQL = SQL & "codusu SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    SQL = SQL & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute SQL
    
    cont = 0
    Lista = ""
    
    While Not RS.EOF
    'para cada cliente/proveedor enviamos un e-mail
        cad1 = DBLet(RS.Fields(2), "T") 'e-mail administracion
        cad2 = DBLet(RS.Fields(3), "T") 'e-mail compras
        
        If cad1 = "" And cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              Lista = Lista & Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & vbCrLf
        ElseIf cad1 <> "" And cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If cadTabla = "sprove" Then
                If Me.OptMailCom(0).Value = True Then cad1 = cad2
            Else
                If Me.OptMailCom(1).Value = True Then cad1 = cad2
            End If
        Else 'alguno de los 2 tiene valor
            If cad2 <> "" Then cad1 = cad2  'e-mail para compras
        End If
        
        If cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            With frmImprimir
                .OtrosParametros = Cadparam
                .NumeroParametros = NumParam
                If cadTabla = "sprove" Then
                    SQL = "{sprove.codprove}=" & RS.Fields(0)
                    .Opcion = 306
                Else
                    SQL = "{sclien.codclien}=" & RS.Fields(0)
                    .Opcion = 91
                End If
                .FormulaSeleccion = SQL
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = cadTit
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(RS.Fields(0), "N") & "," & DBSet(RS.Fields(1), "T") & "," & DBSet(cad1, "T") & ")"
                    conn.Execute SQL
            
                    Me.Refresh
                    Espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    SQL = RS.Fields(0) & ".pdf"
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
                End If
            End With
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
      
    If cont > 0 Then
        Espera 0.4
        If cadTabla = "sprove" Then
            SQL = "Carta: " & txtNombre(63).Text & "|"
             SQL = SQL & "Att : " & txtCodigo(62).Text & "|"
        Else
            SQL = "Carta: " & txtNombre(64).Text & "|"
            SQL = SQL & "Att : " & txtCodigo(0).Text & "|"
        End If
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = SQL
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
        
        'Borrar la carpeta con temporales
        Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If Lista <> "" Then
        If cadTabla = "sprove" Then
            Lista = "Proveedores sin e-mail:" & vbCrLf & vbCrLf & Lista
        Else
            Lista = "Clientes sin e-mail:" & vbCrLf & vbCrLf & Lista
        End If
        MsgBox Lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub




Private Sub CargarComboTipoMov(Indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Byte

    On Error GoTo ECargaCombo

'    SQL = "select codtipom, nomtipom from stipom where tipodocu=2 " 'Documentos de Facturas
    '3 abril 2007.
    'Mostraba todas las facturas (movimientos que empizan por F, excepto las rectificativas
    'AHora tiene que mostrarlas todas
    'SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__') and (codtipom<>'FRT')"
    SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__')"  ' and (codtipom<>'FRT')"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    
    If Indice < 1000 Then
            'Son combos normales
         cboTipoMov(Indice).Clear
        
         While Not RS.EOF
             cboTipoMov(Indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
             cboTipoMov(Indice).ItemData(cboTipoMov(Indice).NewIndex) = I
             I = I + 1
             RS.MoveNext
         Wend
        
    
    Else
        
        ListTipoMov(Indice).Clear
        
        'LOS TIKCETS NO LOS ENVIO POR MAIL
        While Not RS.EOF
            If RS!Codtipom <> "FTI" Then
            
                ListTipoMov(Indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
                'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
                ListTipoMov(Indice).Selected((ListTipoMov(Indice).NewIndex)) = True
            End If
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'Pongo el dos para todos menos para la de etiquetas cliente
    If Indice < 1000 Then
        If Indice <> 2 Then Me.cboTipoMov(Indice).ListIndex = 2
    End If
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub PonerFramePedVisible(H As Integer, W As Integer)
'Frame de Pedidos de Venta y Compra
    W = 6075
    H = 4455
    PonerFrameVisible Me.FramePedidos, True, H, W
    chkPedidoValorado(0).Value = 1
    chkPedidoValorado(0).visible = False
    Select Case OpcionListado
        Case 38 'PEdidos venta
            Me.Label12(0).Caption = "Informe Pedidos ventas"
            NomTabla = "scaped"
            NomTablaLin = "sliped"
            chkPedidoValorado(0).visible = True
        Case 239 'Historico de Pedidos Venta
            Me.Label12(0).Caption = "Informe Hist. Pedidos ventas"
            NomTabla = "schped" 'Cabecera  Hco de Pedidos de clientes
            NomTablaLin = "slhped"
            If FecEntre <> "" Then txtCodigo(76).Text = FecEntre
        Case 55 'Cabecera de Pedidos de Compras (a proveedores)
            Me.Label12(0).Caption = "Informe Pedidos compras"
            NomTabla = "scappr"
            NomTablaLin = "slippr"
            chkPedidoValorado(0).visible = True
        Case 56 'Historico de Pedidos Compras
            Me.Label12(0).Caption = "Informe Hist. Pedidos compras"
            NomTabla = "schppr" 'Cabecera  Hco de Pedidos de Compras (a proveedores)
            NomTablaLin = "slhppr"
            If FecEntre <> "" Then txtCodigo(76).Text = FecEntre
    End Select
    
    
    'Ver Fecha Pedido (En Hist.)
    Label12(2).visible = (OpcionListado = 239) Or OpcionListado = 56
    txtCodigo(76).visible = (OpcionListado = 239) Or OpcionListado = 56
End Sub






 
Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameEnvioFacMail.Height
        Me.Width = Me.FrameEnvioFacMail.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    If peque Then
        FrameEnvioMail.Refresh
        FrameEnvioFacMail.visible = False
    End If
    DoEvents
    Me.Refresh
End Sub



Private Sub PonerTamañosEmailPDF(Email As Boolean)
    


    If Email Then
        Me.FrameEnvioFacMail.Width = 10215
        Label14(16).Caption = "Envio facturas por mail"
    Else
        Me.FrameEnvioFacMail.Width = 5555
        Label14(16).Caption = "Exportar PDF"
    End If
   
    cmdCancel(18).Left = FrameEnvioFacMail.Width - 1240
    cmdEnvioMail.Left = FrameEnvioFacMail.Width - 2275
    DoEvents
    Me.Refresh
End Sub
