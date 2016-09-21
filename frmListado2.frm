VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListado2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDeclaraAlmazara 
      Height          =   2895
      Left            =   2400
      TabIndex        =   524
      Top             =   4560
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdDeclaraAlmazara 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   528
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Declaración definitiva"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   527
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtNumeroEntero 
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   526
         Text            =   "Text3"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   525
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   37
         Left            =   3840
         TabIndex        =   529
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   98
         Left            =   360
         TabIndex        =   532
         Top             =   1920
         Width           =   4425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   960
         TabIndex        =   531
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Declaración mensual almazara"
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
         Index           =   30
         Left            =   120
         TabIndex        =   530
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame FrameAlbaranesVall 
      Height          =   3975
      Left            =   2160
      TabIndex        =   512
      Top             =   1320
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkVarios 
         Caption         =   "Agrupa transporte"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   515
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdListAlbVall 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   516
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   15
         Left            =   1200
         TabIndex        =   514
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   521
         Text            =   "Text5"
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   14
         Left            =   1200
         TabIndex        =   513
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   520
         Text            =   "Text5"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   4680
         TabIndex        =   517
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   97
         Left            =   360
         TabIndex        =   523
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   96
         Left            =   360
         TabIndex        =   522
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   15
         Left            =   840
         Picture         =   "frmListado2.frx":0000
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   14
         Left            =   840
         Picture         =   "frmListado2.frx":0102
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   519
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado albaranes oliva"
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
         Index           =   29
         Left            =   1200
         TabIndex        =   518
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame FramePalets 
      Height          =   3855
      Left            =   2040
      TabIndex        =   487
      Top             =   2040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdPalets 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   496
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Index           =   3
         Left            =   120
         TabIndex        =   502
         Top             =   1920
         Width           =   6615
         Begin VB.TextBox txtHora 
            Height          =   285
            Index           =   0
            Left            =   5400
            TabIndex        =   495
            Text            =   "Text3"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   42
            Left            =   3120
            TabIndex        =   494
            Text            =   "Text1"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtNumeroEntero 
            Height          =   285
            Index           =   3
            Left            =   960
            TabIndex        =   493
            Text            =   "Text3"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtDescArticulo 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   12
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   505
            Text            =   "Text5"
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtArticulo 
            Height          =   285
            Index           =   12
            Left            =   1440
            MaxLength       =   16
            TabIndex        =   492
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Hora"
            Height          =   195
            Index           =   95
            Left            =   4680
            TabIndex        =   511
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   94
            Left            =   2160
            TabIndex        =   510
            Top             =   840
            Width           =   675
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   42
            Left            =   2760
            Picture         =   "frmListado2.frx":0204
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Unidades"
            Height          =   195
            Index           =   93
            Left            =   120
            TabIndex        =   509
            Top             =   840
            Width           =   675
         End
         Begin VB.Image imgArticulo 
            Height          =   240
            Index           =   12
            Left            =   1080
            Picture         =   "frmListado2.frx":028F
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Articulo"
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
            TabIndex        =   506
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   501
         Top             =   1200
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtDescProve 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   13
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   507
            Text            =   "Text5"
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtCodProve 
            Height          =   285
            Index           =   13
            Left            =   1440
            TabIndex        =   491
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
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
            Index           =   17
            Left            =   120
            TabIndex        =   508
            Top             =   240
            Width           =   885
         End
         Begin VB.Image imgProveedor 
            Height          =   240
            Index           =   13
            Left            =   1080
            Picture         =   "frmListado2.frx":0391
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   500
         Top             =   1200
         Width           =   6615
         Begin VB.TextBox txtDescClie 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   8
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   503
            Text            =   "Text1"
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtCliente 
            Height          =   285
            Index           =   8
            Left            =   1440
            TabIndex        =   490
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgCliente 
            Height          =   240
            Index           =   8
            Left            =   1080
            Picture         =   "frmListado2.frx":0493
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblDpto 
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
            Left            =   120
            TabIndex        =   504
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   499
         Top             =   840
         Width           =   6615
         Begin VB.OptionButton optPalets 
            Caption         =   "Entrada"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   489
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton optPalets 
            Caption         =   "Salida"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   488
            Top             =   120
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   35
         Left            =   5400
         TabIndex        =   497
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Movimiento de palets"
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
         Index           =   28
         Left            =   1920
         TabIndex        =   498
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   4920
      TabIndex        =   486
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame FrameOliva 
      Height          =   2415
      Left            =   4920
      TabIndex        =   481
      Top             =   3960
      Width           =   5175
      Begin VB.CommandButton cmdGenerAlbOliva 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   484
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   34
         Left            =   3720
         TabIndex        =   482
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Generar albaranes  desde  la entrada de camiones"
         Height          =   315
         Index           =   92
         Left            =   360
         TabIndex        =   485
         Top             =   1200
         Width           =   4305
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar albaranes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   27
         Left            =   240
         TabIndex        =   483
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame FrameEntradaOliva 
      Height          =   3615
      Left            =   5760
      TabIndex        =   472
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtNumeroEntero 
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   475
         Text            =   "Text3"
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton optTransporte 
         Caption         =   "Albaranes"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   476
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optTransporte 
         Caption         =   "Etiquetas albaranes"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   474
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optTransporte 
         Caption         =   "Resumen carga"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   473
         Top             =   960
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmd4TondaAlbaranOliva 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1920
         TabIndex        =   477
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   33
         Left            =   3360
         TabIndex        =   478
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Vacias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   480
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Impresión entrada oliva"
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
         Index           =   26
         Left            =   360
         TabIndex        =   479
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameTransporte4Tonda 
      Height          =   2535
      Left            =   2160
      TabIndex        =   466
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdImpresion4Tonda 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   471
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optTransporte 
         Caption         =   "Documento de control"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   470
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optTransporte 
         Caption         =   "Alb. entrega y lista pesos"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   469
         Top             =   1080
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   32
         Left            =   3840
         TabIndex        =   467
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Impresión transporte"
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
         Index           =   25
         Left            =   1320
         TabIndex        =   468
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame FrameAmpliaProd 
      Height          =   2415
      Left            =   3240
      TabIndex        =   459
      Top             =   3120
      Width           =   5295
      Begin VB.CommandButton cmdLineaExtraProd 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   462
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   360
         MaxLength       =   48
         TabIndex        =   461
         Text            =   "Text2"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   360
         MaxLength       =   48
         TabIndex        =   460
         Text            =   "Text2"
         Top             =   480
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   4080
         TabIndex        =   463
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Linea extra 2 "
         Height          =   195
         Index           =   91
         Left            =   360
         TabIndex        =   465
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Linea extra1"
         Height          =   195
         Index           =   90
         Left            =   360
         TabIndex        =   464
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame FrameListadoTO 
      Height          =   7095
      Left            =   0
      TabIndex        =   264
      Top             =   0
      Width           =   6855
      Begin VB.Frame FrOrdenTO 
         Height          =   615
         Left            =   120
         TabIndex        =   456
         Top             =   5760
         Width           =   3855
         Begin VB.OptionButton optOrdenTO 
            Caption         =   "Nombre articulo"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   458
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optOrdenTO 
            Caption         =   "Codigo articulo"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   457
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkMatPrima 
         Caption         =   "Materias primas"
         Height          =   255
         Left            =   4200
         TabIndex        =   352
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Frame FrameTosTapa 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   316
         Top             =   480
         Width           =   6495
         Begin VB.TextBox txtCliente 
            Height          =   285
            Index           =   6
            Left            =   1560
            TabIndex        =   247
            Text            =   "Text1"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtDescClie 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   6
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   320
            Text            =   "Text1"
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtCliente 
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   246
            Text            =   "Text1"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtDescClie 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   317
            Text            =   "Text1"
            Top             =   360
            Width           =   3495
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
            Left            =   720
            TabIndex        =   321
            Top             =   840
            Width           =   420
         End
         Begin VB.Image imgCliente 
            Height          =   240
            Index           =   6
            Left            =   1320
            Picture         =   "frmListado2.frx":0595
            Top             =   840
            Width           =   240
         End
         Begin VB.Label lblDpto 
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
            Index           =   27
            Left            =   120
            TabIndex        =   319
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   62
            Left            =   720
            TabIndex        =   318
            Top             =   360
            Width           =   465
         End
         Begin VB.Image imgCliente 
            Height          =   240
            Index           =   5
            Left            =   1320
            Picture         =   "frmListado2.frx":0697
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.OptionButton optListadoTO 
         Caption         =   "Cliente (L)"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   259
         Top             =   5400
         Width           =   1335
      End
      Begin VB.OptionButton optListadoTO 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   258
         Top             =   5400
         Width           =   1095
      End
      Begin VB.OptionButton optListadoTO 
         Caption         =   "Articulo"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   257
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtNumeroEntero 
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   251
         Text            =   "Text1"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtNumeroEntero 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   250
         Text            =   "Text1"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton optListadoTO 
         Caption         =   "Codigo T."
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   256
         Top             =   5400
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdListadoTO 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   260
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   27
         Left            =   4680
         TabIndex        =   255
         Text            =   "Text1"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   26
         Left            =   1560
         TabIndex        =   254
         Text            =   "Text1"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   25
         Left            =   4680
         TabIndex        =   253
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   24
         Left            =   1560
         TabIndex        =   252
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   5400
         TabIndex        =   261
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   9
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   249
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   269
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   248
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   265
         Text            =   "Text5"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   67
         Left            =   3720
         TabIndex        =   304
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   66
         Left            =   600
         TabIndex        =   303
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Codigo TO"
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
         Left            =   240
         TabIndex        =   302
         Top             =   2880
         Width           =   840
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   4440
         Picture         =   "frmListado2.frx":0799
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   3840
         TabIndex        =   276
         Top             =   4800
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   1320
         Picture         =   "frmListado2.frx":0824
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   64
         Left            =   600
         TabIndex        =   275
         Top             =   4800
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha fin"
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
         Left            =   240
         TabIndex        =   274
         Top             =   4440
         Width           =   750
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   4440
         Picture         =   "frmListado2.frx":08AF
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   600
         TabIndex        =   273
         Top             =   4080
         Width           =   465
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
         Left            =   3840
         TabIndex        =   272
         Top             =   4080
         Width           =   420
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicio"
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
         Index           =   28
         Left            =   240
         TabIndex        =   271
         Top             =   3720
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1320
         Picture         =   "frmListado2.frx":093A
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   61
         Left            =   600
         TabIndex        =   270
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   9
         Left            =   1200
         Picture         =   "frmListado2.frx":09C5
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   60
         Left            =   600
         TabIndex        =   268
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   8
         Left            =   1200
         Picture         =   "frmListado2.frx":0AC7
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "L"
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
         Index           =   15
         Left            =   600
         TabIndex        =   267
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label lblDpto 
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
         Index           =   26
         Left            =   240
         TabIndex        =   266
         Top             =   1800
         Width           =   660
      End
   End
   Begin VB.Frame FrameResuProduccion 
      Height          =   2775
      Left            =   3360
      TabIndex        =   415
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkVarios 
         Caption         =   "Indicar palets"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   418
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdResumenProduccion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   419
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   39
         Left            =   3840
         TabIndex        =   417
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   38
         Left            =   1200
         TabIndex        =   416
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   3840
         TabIndex        =   420
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha producción"
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
         Index           =   40
         Left            =   240
         TabIndex        =   424
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   84
         Left            =   240
         TabIndex        =   423
         Top             =   1365
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   38
         Left            =   960
         Picture         =   "frmListado2.frx":0BC9
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   83
         Left            =   3000
         TabIndex        =   422
         Top             =   1365
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   39
         Left            =   3600
         Picture         =   "frmListado2.frx":0C54
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Resumen producción"
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
         Index           =   21
         Left            =   960
         TabIndex        =   421
         Top             =   360
         Width           =   3045
      End
   End
   Begin VB.Frame FrameCambioProveedor 
      Height          =   2535
      Left            =   3600
      TabIndex        =   425
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdCambioProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   427
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   429
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   12
         Left            =   1800
         TabIndex        =   426
         Text            =   "Text1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   28
         Left            =   5640
         TabIndex        =   428
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambio de proveedor en albaranes"
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
         Index           =   22
         Left            =   720
         TabIndex        =   431
         Top             =   360
         Width           =   5595
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListado2.frx":0CDF
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   360
         TabIndex        =   430
         Top             =   1200
         Width           =   885
      End
   End
   Begin VB.Frame FrLiqCambioPrecios 
      Height          =   5055
      Left            =   120
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdCambiarImporte 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   3600
         TabIndex        =   93
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   103
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   91
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   94
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   90
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   12
         Left            =   4560
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLiqu 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   106
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         TabIndex        =   105
         Top             =   3720
         Width           =   705
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmListado2.frx":0DE1
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   104
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmListado2.frx":0EE3
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   102
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
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
         Index           =   5
         Left            =   120
         TabIndex        =   101
         Top             =   1560
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":0FE5
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   100
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   4320
         Picture         =   "frmListado2.frx":10E7
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   97
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1680
         Picture         =   "frmListado2.frx":1172
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   960
         TabIndex        =   96
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         TabIndex        =   95
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambio precios albaranes proveedor"
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
         Index           =   6
         Left            =   240
         TabIndex        =   86
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame frameContabTickets 
      Height          =   3495
      Left            =   120
      TabIndex        =   188
      Top             =   0
      Width           =   6255
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   198
         Top             =   1440
         Width           =   6015
         Begin VB.TextBox txtTrab 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   203
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtDescTra 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   202
            Text            =   "Text1"
            Top             =   840
            Width           =   3255
         End
         Begin VB.OptionButton optTick 
            Caption         =   "Diario"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   200
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTick 
            Caption         =   "Mensual"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   199
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgTecnico 
            Height          =   240
            Index           =   2
            Left            =   1320
            Picture         =   "frmListado2.frx":11FD
            Top             =   840
            Width           =   240
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador: "
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
            TabIndex        =   204
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Agrupa por: "
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
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   201
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdContabTicket 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   3720
         TabIndex        =   191
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   21
         Left            =   4560
         TabIndex        =   190
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   189
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5040
         TabIndex        =   192
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   197
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         TabIndex        =   196
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   1080
         TabIndex        =   195
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   4320
         Picture         =   "frmListado2.frx":12FF
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1680
         Picture         =   "frmListado2.frx":138A
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   3840
         TabIndex        =   194
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "tickets agrupados"
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
         Index           =   10
         Left            =   600
         TabIndex        =   193
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrGeneraFactLiq 
      Height          =   6855
      Left            =   120
      TabIndex        =   107
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtDescForpa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text1"
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox txtForpa 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   119
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   118
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "Text1"
         Top             =   5160
         Width           =   3615
      End
      Begin VB.CheckBox chkFacturPorv 
         Caption         =   "Tesoreria"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   116
         Top             =   4245
         Width           =   1095
      End
      Begin VB.CheckBox chkFacturPorv 
         Caption         =   "Marcar Contabilizar"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   115
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtDescBancoPr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   158
         Text            =   "Text5"
         Top             =   4680
         Width           =   3615
      End
      Begin VB.TextBox txtBancoPr 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   17
         Left            =   1680
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   14
         Left            =   4680
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   112
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   121
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFacProv 
         Caption         =   "Generar"
         Height          =   375
         Left            =   3600
         TabIndex        =   120
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Image imgForPa 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":1415
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Forma pago"
         Height          =   195
         Index           =   38
         Left            =   240
         TabIndex        =   164
         Top             =   5640
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Operador"
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   162
         Top             =   5160
         Width           =   945
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":1517
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Banco propio"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   160
         Top             =   4680
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   35
         Left            =   240
         TabIndex        =   159
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":1619
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1320
         Picture         =   "frmListado2.frx":171B
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Datos facturación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   157
         Top             =   3840
         Width           =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   120
         X2              =   6240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   3720
         TabIndex        =   134
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   133
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. albaran"
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
         Left            =   120
         TabIndex        =   132
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar facturas Liq. proveedores"
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
         Index           =   7
         Left            =   240
         TabIndex        =   131
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         Left            =   120
         TabIndex        =   130
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   129
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   4440
         Picture         =   "frmListado2.frx":17A6
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   3720
         TabIndex        =   128
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1680
         Picture         =   "frmListado2.frx":1831
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   127
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   5
         Left            =   840
         Picture         =   "frmListado2.frx":18BC
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   10
         Left            =   120
         TabIndex        =   126
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   125
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmListado2.frx":19BE
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   124
         Top             =   6360
         Width           =   3375
      End
   End
   Begin VB.Frame FrameTO 
      Height          =   6015
      Left            =   120
      TabIndex        =   227
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdMarca 
         Caption         =   "-"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   233
         ToolTipText     =   "quitar marca"
         Top             =   3360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Caption         =   "+"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   232
         ToolTipText     =   "Añadir marca"
         Top             =   3360
         Width           =   255
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   305
         Top             =   3360
         Width           =   4455
      End
      Begin VB.TextBox txtDescFamilia 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   262
         Text            =   "Text1"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtFamilia 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   231
         Text            =   "Text1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDescFamilia 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   243
         Text            =   "Text1"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtFamilia 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   230
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdGeneraTO 
         Caption         =   "Continuar"
         Height          =   375
         Left            =   4200
         TabIndex        =   234
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   238
         Text            =   "Text5"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   7
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   229
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   235
         Text            =   "Text5"
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   6
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   228
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5400
         TabIndex        =   236
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   59
         Left            =   600
         TabIndex        =   263
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmListado2.frx":1AC0
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   58
         Left            =   600
         TabIndex        =   245
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Formatos"
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
         Left            =   120
         TabIndex        =   244
         Top             =   1920
         Width           =   810
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmListado2.frx":1BC2
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   24
         Left            =   120
         TabIndex        =   242
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   23
         Left            =   120
         TabIndex        =   241
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Proceso generación Tarifa-Oferta"
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
         Left            =   480
         TabIndex        =   240
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmListado2.frx":1CC4
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   480
         TabIndex        =   239
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado2.frx":1DC6
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   54
         Left            =   480
         TabIndex        =   237
         Top             =   960
         Width           =   465
      End
   End
   Begin VB.Frame FrameTraza 
      Height          =   4935
      Left            =   120
      TabIndex        =   205
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   210
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdTraza 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   215
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   224
         Text            =   "Text5"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   209
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   223
         Text            =   "Text5"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   221
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   212
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   219
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   211
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5160
         TabIndex        =   216
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   23
         Left            =   5160
         TabIndex        =   214
         Text            =   "Text1"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   22
         Left            =   2040
         TabIndex        =   213
         Text            =   "Text1"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   480
         TabIndex        =   226
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   480
         TabIndex        =   225
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado2.frx":1EC8
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado2.frx":1FCA
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   51
         Left            =   480
         TabIndex        =   222
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   11
         Left            =   1080
         Picture         =   "frmListado2.frx":20CC
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   480
         TabIndex        =   220
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   10
         Left            =   1080
         Picture         =   "frmListado2.frx":21CE
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   4920
         Picture         =   "frmListado2.frx":22D0
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   49
         Left            =   4320
         TabIndex        =   218
         Top             =   3645
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1800
         Picture         =   "frmListado2.frx":235B
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   1200
         TabIndex        =   217
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   208
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Albaran proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   207
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Trazabilidad albaranes"
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
         Index           =   11
         Left            =   240
         TabIndex        =   206
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameEcoEnves 
      Height          =   3015
      Left            =   120
      TabIndex        =   364
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton optEcoenves 
         Caption         =   "Facturas "
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   414
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optEcoenves 
         Caption         =   "Ecoembes"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   413
         Top             =   1920
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdEcoEnves 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   368
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4440
         TabIndex        =   369
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   35
         Left            =   4200
         TabIndex        =   367
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   34
         Left            =   1800
         TabIndex        =   366
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   360
         TabIndex        =   373
         Top             =   2400
         Width           =   2385
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   1560
         Picture         =   "frmListado2.frx":23E6
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   71
         Left            =   3240
         TabIndex        =   372
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   35
         Left            =   3960
         Picture         =   "frmListado2.frx":2471
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   840
         TabIndex        =   371
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   34
         Left            =   120
         TabIndex        =   370
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informes ECOEMBES"
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
         Index           =   19
         Left            =   120
         TabIndex        =   365
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.Frame FrameVentasMorales 
      Height          =   3255
      Left            =   360
      TabIndex        =   306
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   312
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   362
         Text            =   "Text1"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox chkConsolidado 
         Caption         =   "Ver todos los almacenes"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   354
         Top             =   2880
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4920
         TabIndex        =   315
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdVentasAceite 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3720
         TabIndex        =   314
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   29
         Left            =   4680
         TabIndex        =   311
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   28
         Left            =   2160
         TabIndex        =   307
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   7
         Left            =   600
         Picture         =   "frmListado2.frx":24FC
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Left            =   240
         TabIndex        =   363
         Top             =   1800
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   4440
         Picture         =   "frmListado2.frx":25FE
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   57
         Left            =   3720
         TabIndex        =   313
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         Index           =   31
         Left            =   120
         TabIndex        =   310
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   56
         Left            =   1200
         TabIndex        =   309
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   1920
         Picture         =   "frmListado2.frx":2689
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informes ventas Aceite"
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
         Index           =   16
         Left            =   480
         TabIndex        =   308
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrFacturaRecargas 
      Height          =   6015
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtBancoPr 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtDescBancoPr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   165
         Text            =   "Text5"
         Top             =   4560
         Width           =   4095
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   5400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   0
         Left            =   720
         MaxLength       =   16
         TabIndex        =   39
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtRecargaMov 
         Height          =   285
         Index           =   1
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   42
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdFacturaMov 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   41
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Banco propio"
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
         Left            =   120
         TabIndex        =   166
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":2714
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label lblIndicadorT 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   5040
         Width           =   3615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListado2.frx":2816
         Top             =   2175
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado2.frx":28A1
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   50
         Top             =   3600
         Width           =   660
      End
      Begin VB.Label lblDpto 
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
         Index           =   8
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   5040
         TabIndex        =   47
         Top             =   840
         Width           =   360
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":29A3
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   3000
         Picture         =   "frmListado2.frx":2AA5
         Top             =   1222
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   2520
         TabIndex        =   45
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmListado2.frx":2B30
         Top             =   1222
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   44
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha recarga"
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
         TabIndex        =   43
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Facturación  recargas moviles"
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
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameInvenACeite 
      Height          =   3375
      Left            =   4320
      TabIndex        =   355
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdInvenAceite 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2040
         TabIndex        =   361
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   33
         Left            =   2280
         TabIndex        =   359
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkConsolidado 
         Caption         =   "Ver todos los almacenes"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   358
         Top             =   2040
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   3360
         TabIndex        =   356
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1920
         Picture         =   "frmListado2.frx":2BBB
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inventario"
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
         TabIndex        =   360
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Inventario aceite"
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
         Index           =   18
         Left            =   1545
         TabIndex        =   357
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.Frame FrameLiqAgentes 
      Height          =   3615
      Left            =   240
      TabIndex        =   322
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkConsolidado 
         Caption         =   "Ver todos los almacenes"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   353
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ListBox List2 
         Height          =   1410
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   330
         Top             =   1440
         Width           =   5175
      End
      Begin VB.CommandButton cmdVentasAgentes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   329
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   31
         Left            =   4200
         TabIndex        =   326
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   30
         Left            =   1320
         TabIndex        =   324
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4320
         TabIndex        =   323
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Ventas por agentes"
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
         Index           =   17
         Left            =   600
         TabIndex        =   328
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   3360
         TabIndex        =   327
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   31
         Left            =   3960
         Picture         =   "frmListado2.frx":2C46
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   68
         Left            =   360
         TabIndex        =   325
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   1080
         Picture         =   "frmListado2.frx":2CD1
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameRecargaMov 
      Height          =   3375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtRecargaMov 
         Height          =   285
         Index           =   0
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cmbRecargaMov 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado2.frx":2D5C
         Left            =   3840
         List            =   "frmListado2.frx":2D69
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1980
         Width           =   975
      End
      Begin VB.ComboBox cmbRecargaMov 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado2.frx":2D7C
         Left            =   1800
         List            =   "frmListado2.frx":2D89
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1980
         Width           =   975
      End
      Begin VB.CommandButton cmdRecargasMov 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   24
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   4680
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   26
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   32
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobradas"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   31
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Facturadas"
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   30
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe recargas moviles"
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
         Index           =   3
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   28
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   4440
         Picture         =   "frmListado2.frx":2D9C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   25
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1560
         Picture         =   "frmListado2.frx":2E27
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrEstadisticasReparacionTecnico 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdEstadisticaReparacionTecnico 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmListado2.frx":2EB2
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Técnico"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   3960
         Picture         =   "frmListado2.frx":2FB4
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   9
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListado2.frx":303F
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Estadísticas reparación técnico"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrListadoReparaciones 
      Height          =   4335
      Left            =   120
      TabIndex        =   277
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   291
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   290
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   289
         Text            =   "Text1"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   288
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   287
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   286
         Text            =   "Text1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   285
         Text            =   "Text1"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   284
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   283
         Text            =   "Text1"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   282
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   281
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdReparaEfect 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   280
         Top             =   3720
         Width           =   975
      End
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha reparacion"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   279
         Top             =   3720
         Width           =   1815
      End
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha entrada"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   278
         Top             =   3720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":30CA
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListado2.frx":31CC
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Reparaciones efectuadas"
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
         Left            =   240
         TabIndex        =   301
         Top             =   120
         Width           =   5895
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":32CE
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":3359
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   0
         Left            =   240
         TabIndex        =   300
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListado2.frx":345B
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   840
         TabIndex        =   299
         Top             =   840
         Width           =   465
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
         Index           =   23
         Left            =   840
         TabIndex        =   298
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   297
         Top             =   1920
         Width           =   465
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
         Index           =   0
         Left            =   840
         TabIndex        =   296
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   295
         Top             =   3000
         Width           =   465
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
         Left            =   3600
         TabIndex        =   294
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblDpto 
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
         Index           =   1
         Left            =   240
         TabIndex        =   293
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblDpto 
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
         Left            =   240
         TabIndex        =   292
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmListado2.frx":355D
         Top             =   3000
         Width           =   240
      End
   End
   Begin VB.Frame FrameMultibase 
      Height          =   5175
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   5775
      Begin VB.ListBox lstMultibase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   960
         Width           =   5295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   14
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdMultibase 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblMultibase 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Revisar caracteres especiales"
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
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame FrImprimirFac 
      Height          =   4575
      Left            =   120
      TabIndex        =   135
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdImprimirFac 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3600
         TabIndex        =   142
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5040
         TabIndex        =   143
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   141
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   144
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   16
         Left            =   4560
         TabIndex        =   137
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   138
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   156
         Top             =   3720
         Width           =   6135
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   7
         Left            =   840
         Picture         =   "frmListado2.frx":35E8
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   155
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label4 
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
         Index           =   8
         Left            =   120
         TabIndex        =   154
         Top             =   2520
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmListado2.frx":36EA
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   153
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   4320
         Picture         =   "frmListado2.frx":37EC
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   3720
         TabIndex        =   152
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1680
         Picture         =   "frmListado2.frx":3877
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   151
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         TabIndex        =   150
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Imprimir facturas proveedores"
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
         Index           =   8
         Left            =   240
         TabIndex        =   149
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. factura"
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
         Left            =   120
         TabIndex        =   148
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   960
         TabIndex        =   147
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   3720
         TabIndex        =   146
         Top             =   2040
         Width           =   465
      End
   End
   Begin VB.Frame FrProveedorxVenta 
      Height          =   5895
      Left            =   120
      TabIndex        =   55
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   61
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   10
         Left            =   4560
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdVentaxProv 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   64
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5160
         TabIndex        =   65
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   60
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   84
         Top             =   4680
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado2.frx":3902
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   4
         Left            =   120
         TabIndex        =   82
         Top             =   3960
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   81
         Top             =   4320
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado2.frx":3A04
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":3B06
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   78
         Top             =   3525
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   77
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   76
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   75
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   3720
         TabIndex        =   74
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   4320
         Picture         =   "frmListado2.frx":3C08
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   73
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado2.frx":3C93
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmListado2.frx":3D95
         Top             =   2175
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmListado2.frx":3E97
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   11
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado venta x proveedor"
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
         Index           =   5
         Left            =   360
         TabIndex        =   68
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   66
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   960
         Picture         =   "frmListado2.frx":3F99
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameVtasVarias 
      Height          =   7335
      Left            =   8280
      TabIndex        =   374
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton optVtaGrup 
         Caption         =   "Marca / Modelo / Formato"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   412
         Top             =   5760
         Width           =   2535
      End
      Begin VB.OptionButton optVtaGrup 
         Caption         =   "Formato / Modelo / Marca"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   411
         Top             =   5760
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.CommandButton cmdVtasAgrupadas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   410
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   405
         Text            =   "Text1"
         Top             =   5160
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   404
         Text            =   "Text1"
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   402
         Text            =   "Text1"
         Top             =   4800
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   401
         Text            =   "Text1"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   399
         Text            =   "Text1"
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   398
         Text            =   "Text1"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   396
         Text            =   "Text1"
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   395
         Text            =   "Text1"
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   393
         Text            =   "Text1"
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   392
         Text            =   "Text1"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   390
         Text            =   "Text1"
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   389
         Text            =   "Text1"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   387
         Text            =   "Text1"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   386
         Text            =   "Text1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   384
         Text            =   "Text1"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   383
         Text            =   "Text1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   37
         Left            =   4200
         TabIndex        =   378
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   36
         Left            =   1440
         TabIndex        =   376
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   26
         Left            =   4560
         TabIndex        =   375
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
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
         Left            =   240
         TabIndex        =   409
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Formato"
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
         Left            =   240
         TabIndex        =   408
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   37
         Left            =   240
         TabIndex        =   407
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   82
         Left            =   600
         TabIndex        =   406
         Top             =   5160
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   7
         Left            =   1200
         Picture         =   "frmListado2.frx":4024
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   600
         TabIndex        =   403
         Top             =   4800
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   6
         Left            =   1200
         Picture         =   "frmListado2.frx":4126
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   600
         TabIndex        =   400
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   5
         Left            =   1200
         Picture         =   "frmListado2.frx":4228
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   600
         TabIndex        =   397
         Top             =   3840
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   4
         Left            =   1200
         Picture         =   "frmListado2.frx":432A
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   600
         TabIndex        =   394
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "frmListado2.frx":442C
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   600
         TabIndex        =   391
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmListado2.frx":452E
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   76
         Left            =   600
         TabIndex        =   388
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmListado2.frx":4630
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   75
         Left            =   600
         TabIndex        =   385
         Top             =   1800
         Width           =   465
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmListado2.frx":4732
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Categoría"
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
         Left            =   240
         TabIndex        =   382
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   35
         Left            =   240
         TabIndex        =   381
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe ventas agrupado"
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
         Index           =   20
         Left            =   240
         TabIndex        =   380
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   37
         Left            =   3840
         Picture         =   "frmListado2.frx":4834
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   74
         Left            =   3240
         TabIndex        =   379
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   73
         Left            =   480
         TabIndex        =   377
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   36
         Left            =   1200
         Picture         =   "frmListado2.frx":48BF
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameEnvioTarifa 
      Height          =   7815
      Left            =   0
      TabIndex        =   331
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdImprimirTOs 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   6840
         TabIndex        =   345
         Top             =   7320
         Width           =   1095
      End
      Begin VB.TextBox txtTarifa 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   334
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   8040
         TabIndex        =   333
         Top             =   7320
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6255
         Left            =   240
         TabIndex        =   332
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Clientes"
         TabPicture(0)   =   "frmListado2.frx":494A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "imgCheck(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "imgCheck(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "List3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Datos carta"
         TabPicture(1)   =   "frmListado2.frx":4966
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtFecha(32)"
         Tab(1).Control(1)=   "txtCarta(2)"
         Tab(1).Control(2)=   "chkCartaTO"
         Tab(1).Control(3)=   "cmdGuardar"
         Tab(1).Control(4)=   "txtCarta(4)"
         Tab(1).Control(5)=   "txtCarta(5)"
         Tab(1).Control(6)=   "txtCarta(3)"
         Tab(1).Control(7)=   "txtCarta(1)"
         Tab(1).Control(8)=   "txtCarta(0)"
         Tab(1).Control(9)=   "Label6(4)"
         Tab(1).Control(10)=   "Label6(3)"
         Tab(1).Control(11)=   "Label6(2)"
         Tab(1).Control(12)=   "Label6(1)"
         Tab(1).Control(13)=   "Label6(0)"
         Tab(1).ControlCount=   14
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   32
            Left            =   -73560
            TabIndex        =   350
            Text            =   "Text1"
            Top             =   5160
            Width           =   1215
         End
         Begin VB.TextBox txtCarta 
            Height          =   1125
            Index           =   2
            Left            =   -74640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   349
            Text            =   "frmListado2.frx":4982
            Top             =   2400
            Width           =   8175
         End
         Begin VB.CheckBox chkCartaTO 
            Caption         =   "Imprimir carta"
            Height          =   255
            Left            =   -74640
            TabIndex        =   348
            Top             =   5760
            Width           =   1575
         End
         Begin VB.CommandButton cmdGuardar 
            Height          =   375
            Left            =   -66960
            Picture         =   "frmListado2.frx":4988
            Style           =   1  'Graphical
            TabIndex        =   347
            ToolTipText     =   "Guardar datos carta en PC"
            Top             =   480
            Width           =   375
         End
         Begin VB.ListBox List3 
            Height          =   5010
            Left            =   720
            Style           =   1  'Checkbox
            TabIndex        =   344
            Top             =   600
            Width           =   7455
         End
         Begin VB.TextBox txtCarta 
            Height          =   285
            Index           =   4
            Left            =   -71160
            TabIndex        =   339
            Text            =   "Text1"
            Top             =   5160
            Width           =   4695
         End
         Begin VB.TextBox txtCarta 
            Height          =   285
            Index           =   5
            Left            =   -71160
            TabIndex        =   338
            Text            =   "Text1"
            Top             =   5640
            Width           =   4695
         End
         Begin VB.TextBox txtCarta 
            Height          =   1245
            Index           =   3
            Left            =   -74640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   337
            Text            =   "frmListado2.frx":4F12
            Top             =   3600
            Width           =   8175
         End
         Begin VB.TextBox txtCarta 
            Height          =   885
            Index           =   1
            Left            =   -74640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   336
            Text            =   "frmListado2.frx":4F18
            Top             =   1320
            Width           =   8175
         End
         Begin VB.TextBox txtCarta 
            Height          =   285
            Index           =   0
            Left            =   -74760
            TabIndex        =   335
            Text            =   "Text1"
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha carta"
            Height          =   255
            Index           =   4
            Left            =   -74640
            TabIndex        =   351
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   720
            Picture         =   "frmListado2.frx":4F1E
            ToolTipText     =   "Quitar seleccion"
            Top             =   5760
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   1080
            Picture         =   "frmListado2.frx":5068
            ToolTipText     =   "Seleccionar todos"
            Top             =   5760
            Width           =   240
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Firmado"
            Height          =   255
            Index           =   3
            Left            =   -72000
            TabIndex        =   343
            Top             =   5640
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Despedida"
            Height          =   255
            Index           =   2
            Left            =   -72240
            TabIndex        =   342
            Top             =   5160
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Parrafos"
            Height          =   255
            Index           =   1
            Left            =   -74640
            TabIndex        =   341
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "Saludos"
            Height          =   255
            Index           =   0
            Left            =   -74640
            TabIndex        =   340
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   346
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame FrameRecalPrStandard 
      Height          =   3855
      Left            =   1560
      TabIndex        =   441
      Top             =   3240
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdRecalPrSt 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   454
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton optRecalPrSt 
         Caption         =   "Producto venta"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   445
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton optRecalPrSt 
         Caption         =   "Materia auxiliar"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   444
         Top             =   2400
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CheckBox chkRecalPrSt 
         Caption         =   "Materia auxiliar desde precio standard"
         Height          =   255
         Left            =   480
         TabIndex        =   453
         Top             =   2880
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   11
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   443
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   452
         Text            =   "Text5"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   10
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   442
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   448
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   5160
         TabIndex        =   446
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   89
         Left            =   120
         TabIndex        =   455
         Top             =   3360
         Width           =   3465
      End
      Begin VB.Image ImgAyuda 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   11
         Left            =   960
         Picture         =   "frmListado2.frx":51B2
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   88
         Left            =   360
         TabIndex        =   451
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   87
         Left            =   360
         TabIndex        =   450
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   449
         Top             =   840
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   10
         Left            =   960
         Picture         =   "frmListado2.frx":52B4
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Recálculo precio standard"
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
         Index           =   24
         Left            =   1200
         TabIndex        =   447
         Top             =   240
         Width           =   4125
      End
   End
   Begin VB.Frame FrameAlbaProv 
      Height          =   4095
      Left            =   0
      TabIndex        =   167
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   178
         Text            =   "Text1"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   177
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   182
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   5
         Left            =   4920
         TabIndex        =   176
         Text            =   "Text1"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   173
         Text            =   "Text1"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   19
         Left            =   4920
         TabIndex        =   171
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   168
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlbaranProv 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4320
         TabIndex        =   179
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5400
         TabIndex        =   180
         Top             =   3480
         Width           =   975
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   9
         Left            =   840
         Picture         =   "frmListado2.frx":53B6
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   240
         TabIndex        =   187
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Imprimir albarán proveedor"
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
         Index           =   9
         Left            =   720
         TabIndex        =   185
         Top             =   120
         Width           =   5415
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListado2.frx":54B8
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   240
         TabIndex        =   184
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label4 
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
         Index           =   11
         Left            =   120
         TabIndex        =   183
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   3960
         TabIndex        =   181
         Top             =   1725
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. albaran"
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
         Left            =   120
         TabIndex        =   175
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   960
         TabIndex        =   174
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   40
         Left            =   3960
         TabIndex        =   172
         Top             =   885
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   4680
         Picture         =   "frmListado2.frx":55BA
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1680
         Picture         =   "frmListado2.frx":5645
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   960
         TabIndex        =   170
         Top             =   885
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         TabIndex        =   169
         Top             =   600
         Width           =   1185
      End
   End
   Begin VB.Frame FrameResumenProduccionMoixent 
      Height          =   2415
      Left            =   840
      TabIndex        =   432
      Top             =   5280
      Width           =   5055
      Begin VB.CommandButton cmdResprodMoixent 
         Caption         =   "MOIXENT"
         Height          =   375
         Left            =   2160
         TabIndex        =   440
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   41
         Left            =   3600
         TabIndex        =   439
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   40
         Left            =   1200
         TabIndex        =   435
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   29
         Left            =   3600
         TabIndex        =   433
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   40
         Left            =   3360
         Picture         =   "frmListado2.frx":56D0
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   86
         Left            =   2880
         TabIndex        =   438
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   32
         Left            =   960
         Picture         =   "frmListado2.frx":575B
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   85
         Left            =   240
         TabIndex        =   437
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   41
         Left            =   240
         TabIndex        =   436
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Resumen producción"
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
         Index           =   23
         Left            =   960
         TabIndex        =   434
         Top             =   360
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frmListado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public opcion As Integer
    '1      .- Listado reparaciones efectuadas
    '2      .- Reparaciones tecnico
    
    '3      .- Revision carcteres multibase
    '4      .- Listado recargas telefonia movil
    '5      .- Facturacion de recargas
    
    '6      .- Listado de TRAZA por codprove en ventas.   ENERO 2008
    
    '       LIQUIDACION PROVEEDORES. Socios tipo TERRASANA
    '7      .- Cambio precio articulos
    '8      .- Generar facturas
    '9      .- Imprimir facturas proveedores (socios)
    '10     .-   "      ALBARANES   "           "
    
    
    '13     .- Generacion y facturacion de tickets agrupados
    '14     .- Listado del punto anterior
    
    '15     .- Listado trazabilidad albaranes
    
    '16     .-  OLI.  Tarifa-Oferta
    '17     .-  Listado TOS  (tarifa-oferta)s
    
    '19     .-  Informe Ventas  MORALEs
    '20     .-  Liquidacion agentes
        
    '21     .- Envio tarifas
    
    '22     .- Inventario de ACEITE. A partir de un inventario, ver el aceite cuanto es
    
    '23     .- Ventas aceite, pero para un cliente determinado
    
    '24     .- ECOENVES
    
    
    '26     .- Ventas agrupado(x cetergoria ,model ,tipoar, marca)
    '27     .- Resumen diario nueva produccion
        
    '28     .-Cambio de proveedor en el albaran
    '29     .- Resumne produccion antiguo(para moixent)
    '30     .- Recalculo del precio standard
    
    '31     .- Lineas Extras produccion
    '32     .- Impresion transporte Coop 4tonda
    '33     .- Impresion entrada oliva
    '34     .- Generar albaranes entrada de oliva
    '35     .- Movimientos palet que no vayan por albaranes
    
    '36     .- Listado albaranes oliva (La vall)
    '37     .- declaración mensual de almazaras (la vall)
    
Private IndiceImg As Integer
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB2 As frmBuscaGrid
Attribute frmB2.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmPr As frmComProveedores
Attribute frmPr.VB_VarHelpID = -1
Private WithEvents frmBaPr As frmFacBancosPropios
Attribute frmBaPr.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmMar As frmAlmMarcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmFor As frmAlmFamiliaArticulo
Attribute frmFor.VB_VarHelpID = -1


Private PrimeraVez As Boolean




'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private Cadparam As String 'Cadena con los parametros para Crystal Report
Private NumParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private Cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------





'Variables comunes a todos os botones aceptar
Dim cadFrom As String
Dim campo As String, Devuelve As String
Dim Codigo  As String
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

Private cadImpresion As String  'Facturacion



Private Sub chkFacturPorv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
         KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmbRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmd4TondaAlbaranOliva_Click()
    If Me.optTransporte(2).Value Then
        
        miSQL = RecuperaValor(Codigo, 1)
        Devuelve = RecuperaValor(Codigo, 2)
        miSQL = "{vallentradacamion.entrada}=" & miSQL
        LlamaImprimirGral miSQL, "", 0, "vallEntradaOliva.rpt", "Entrada oliva: " & Devuelve
        
        
    ElseIf Me.optTransporte(3).Value Then
        If ImprimirEtiquetasAlbaranes Then
            miSQL = "{tmpinformes.codusu}=" & vUsu.Codigo
            LlamaImprimirGral miSQL, "", 0, "vallEtiqEntradaOliva.rpt", "Etiqueta entrada  " & Devuelve
        End If
    Else
        'Albaranes
        miSQL = RecuperaValor(Codigo, 1)
        Devuelve = RecuperaValor(Codigo, 2)
        miSQL = "{vallentradacamion.entrada}=" & miSQL
        LlamaImprimirGral miSQL, "", 0, "vallEntradaOlivaAlb.rpt", "Albaranes entrada  " & Devuelve
        
    End If
End Sub

Private Sub cmdAlbaranProv_Click()

    InicializarVbles
    
    'Albaran socio
    If Not PonerParamRPT(27, Cadparam, NumParam, cadNomRPT) Then Exit Sub
    
    Cadselect = "{sprove.tipprove}>=3"   'Estos proveedores son los REA o estimacion directa que luego
    cadFormula = "(" & Cadselect & ")"
    If txtFecha(18).Text <> "" Or txtFecha(19).Text <> "" Then
        campo = "{scaalp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 18, 19, Devuelve) Then Exit Sub
    End If
    
    If txtCodProve(8).Text <> "" Or txtCodProve(9).Text <> "" Then
        campo = "{scaalp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 8, 9, Devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(4).Text <> "" Or txtNumAlbar(5).Text <> "" Then
        campo = "{scaalp.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 4, 5, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    campo = "scaalp,sprove WHERE scaalp.codprove=sprove.codprove AND " & Cadselect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    LlamarImprimir2








    frmImprimir.opcion = 2010
    frmImprimir.Show vbModal
End Sub

Private Sub cmdCambiarImporte_Click()
Dim Fecha As Date
Dim vA As CArticulo
    Cadselect = ""
    'Comprobaciones
    If Me.txtImporte(0).Text = "" Then Cadselect = "      -Importe"
    If txtArticulo(3).Text = "" Or Me.txtDescArticulo(3).Text = "" Then Cadselect = Cadselect & vbCrLf & "     -Articulo"
    
    If Cadselect <> "" Then
        MsgBox "Campos obligatorios" & vbCrLf & Cadselect, vbExclamation
        Exit Sub
    End If
    
    
    InicializarVbles
    Devuelve = ""
    
    
    'Cadena obligada. Los proveedores , el tipo tiene que ser el 3: REA
    Cadselect = " {slialp.codprove}=  {sprove.codprove}  AND {sprove.tipprove}= 3"
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(11).Text <> "" Or txtFecha(12).Text <> "" Then
        campo = "{slialp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 11, 12, Devuelve) Then Exit Sub
    End If
    
    If txtCodProve(2).Text <> "" Or txtCodProve(3).Text <> "" Then
        campo = "{slialp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 2, 3, Devuelve) Then Exit Sub
    End If
    
    If Cadselect <> "" Then Cadselect = Cadselect & " AND "
    Cadselect = Cadselect & "  ({slialp.codartic} = '" & txtArticulo(3).Text & "')"
    
    'Vermos si hay registros
    
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    'Pongo el oreder por comodidad
    Cadselect = Cadselect & " ORDER BY fechaalb, slialp.codprove"
    cadFrom = "Select count(*) from slialp,sprove  WHERE " & Cadselect
    
    
    IndiceImg = NumRegistros(cadFrom)
    If IndiceImg = 0 Then
        MsgBox "No hay datos con estos valores", vbExclamation
        Exit Sub
    Else
        cadFrom = "Hay " & IndiceImg & " registro(s) para actualizar el precio" & vbCrLf & _
            "Desea continuar con la actualizacion de precios?"
        
        If MsgBox(cadFrom, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    End If
    
    cadFrom = "Select * from slialp,sprove WHERE " & Cadselect
    
    If Not BloqueoManual("LIQCMBPRE", "1") Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set vA = New CArticulo
    If vA.LeerDatos(CStr(txtArticulo(3).Text)) Then
         
        Set miRsAux = New ADODB.Recordset
        Cadselect = "Select ultfecco from sartic where codartic = '" & DevNombreSQL(txtArticulo(3).Text) & "'"
        miRsAux.Open Cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Fecha = CDate("01/01/1900")
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then Fecha = miRsAux.Fields(0)
        End If
        miRsAux.Close
        NumParam = 0   'Auqi tendre si ha cambiado la fecha o no
        
        ImpTeo = ImporteFormateado(txtImporte(0).Text)
        'Por si lo meto en una transaccion
        RealizarCambiosPreciosLiq Fecha
        
        'Si tengo que updatearl ultcompra
        If NumParam = 1 Then vA.ActualizarUltFechaCompra_ CStr(Fecha), txtImporte(0).Text
        
        
       Me.lblLiqu.Caption = ""
       MsgBox "Proceso finalizado", vbExclamation
           
           
       'Para que no vuelvan a anzar el proceso
       txtArticulo(3).Text = ""
       txtDescArticulo(3).Text = ""
    End If
    Set vA = Nothing
    Set miRsAux = Nothing
    DesBloqueoManual "LIQCMBPRE"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCambioProv_Click()
    If Me.txtCodProve(12).Text = "" Xor Me.txtDescProve(12).Text = "" Then
        MsgBox "Error en proveedor", vbExclamation
        PonerFoco txtCodProve(12)
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = txtCodProve(12).Text
    Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    If Index = 16 Or Index = 28 Or Index = 31 Then CadenaDesdeOtroForm = "" 'Fuerzo esto al cancelar
    Unload Me
End Sub

Private Sub cmdContabTicket_Click()
    If opcion = 13 Then
        ContabilizarTickets
    Else
        Screen.MousePointer = vbHourglass
        ListadoTicketsAgrupados
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ListadoTicketsAgrupados()

    'Meto el resume IVA en tmpnlotes
    Label5.Caption = "Obteniendo datos IVAs"
    Label5.Refresh
    Devuelve = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Devuelve
    
    Devuelve = " FROM    `sfactik` LEFT OUTER JOIN  `scafac` ON (`sfactik`.`numfactu`=`scafac`.`numfactu`) AND (`sfactik`.`fecfactu`=`scafac`.`fecfactu`) "
    cadFrom = ""
    If txtFecha(20).Text <> "" Then cadFrom = cadFrom & " AND fecfacFTG >='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
    If txtFecha(21).Text <> "" Then cadFrom = cadFrom & " AND fecfacFTG <='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
    If cadFrom <> "" Then Devuelve = Devuelve & " WHERE " & Mid(cadFrom, 5)
    Devuelve = Devuelve & " GROUP BY 1,2,3"
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 1
    cadTitulo = "insert into `tmpinformes` (`codusu`,`codigo1`,`nombre1`,`nombre2`,`importe1`,`importe2`,`importe3`) VALUES (" & vUsu.Codigo & ","
    For NumParam = 1 To 3
        cadFrom = ",porciva" & NumParam & " c1,sum(imporiv" & NumParam & ") c2,sum(baseimp" & NumParam & ") c3 "
        cadFrom = "SELECT numfacftg,fecfacftg" & cadFrom & Devuelve
        miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            ' importante. Formatear con 7 0's como minimo,para realizar el link en el informe
             cadFrom = NumRegElim & ",'" & Format(miRsAux!NumFacftg, "0000000") & "','" & miRsAux!FecFacftg & "'"
             'Los importes
             If Not IsNull(miRsAux!C1) Then
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!C1))
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!C2))
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!C3))
                conn.Execute cadTitulo & cadFrom & ")"
                NumRegElim = NumRegElim + 1
             End If
             miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next
    Set miRsAux = Nothing
    Me.Refresh
    Label5.Caption = ""
    InicializarVbles
    If Not PonerParamRPT(28, Cadparam, NumParam, cadNomRPT) Then Exit Sub
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    Devuelve = ""
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        campo = "{sfactik.fecfacftg}"
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHFecha=""Fecha " & Devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 20, 21, Devuelve) Then Exit Sub
    End If
    
    cadFrom = "sfactik"
    If Not HayRegParaInforme(cadFrom, Cadselect) Then Exit Sub
    
    
    conSubRPT = False
    LlamarImprimir2
    Screen.MousePointer = vbDefault
End Sub

Private Sub ContabilizarTickets()
    'IdTrabajador
    'Es importante para las tablas de analitica. Es el que pasa el CC
    If txtTrab(2).Text = "" Then
        MsgBox "Introduza el trabajador que realiza la contabilización", vbExclamation
        Exit Sub
    End If
    
    
    
    'La fecha HASTA sera la fecha de factura para los
    If Me.optTick(1).Value Then
        'MENSUAL
        If txtFecha(21).Text = "" Then
            MsgBox "Debe poner la fecha ""hasta"". Será la fecha de factura ", vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Compruebo si existe el tipo moviemiento
    campo = DevuelveDesdeBD(conAri, "nomtipom", "stipom", "codtipom", "FTG", "T")
    If campo = "" Then
        MsgBox "Falta definir el tipo de moviemiento: FTG", vbExclamation
        Exit Sub
    End If
    
    'Compruebo que no se ha quedado ningun FTG con anteriroridad
    campo = DevuelveDesdeBD(conAri, "numfactu", "scafac", "codtipom", "FTG", "T")
    If campo <> "" Then
        'EXISTE FTG sin haber sido borrado
        MsgBox "Existen FTG que no han sido borrados", vbExclamation
        Exit Sub
    End If
    
    
    
    
    If vEmpresa.TieneAnalitica Then
        cadFrom = ""  'cadena error
        
        
        'Falta configurar la forma de envio en empresa
        campo = DevuelveDesdeBD(conAri, "nomenvio", "senvio", "codenvio", vParamAplic.PorDefecto_Envio)
        If campo = "" Then cadFrom = "- Forma de envio en los parametros de la aplicacion" & vbCrLf
        
        'Comprobar que existen todos los centros de coste en los datos a facturar
        campo = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", txtTrab(2).Text)
        If campo = "" Then
            cadFrom = cadFrom & "- Trabajador sin CC asignado: " & txtTrab(2).Text & vbCrLf
        Else
            'Tiene CC puesto. Veremos que existe en la conta
            Devuelve = DevuelveDesdeBD(conConta, "nomccost", "cabccost", "codccost", campo, "T")
            If Devuelve = "" Then cadFrom = cadFrom & "- Centro de coste del trabajador NO existe." & campo
        End If
        
        If cadFrom <> "" Then
            MsgBox "Falta configurar." & vbCrLf & cadFrom, vbExclamation
            Exit Sub
        End If
    End If
    
    InicializarVbles
    
    
    
    
    'Obtengo el que sera el ultimo registro insertado hasta ahora.
    'No hace falta. TODO proceso debe eliminar las facturas FTG
    'campo = SugerirCodigoSiguienteStr("scafac", "numfactu", "codtipom=""FTG""")
    'NumRegElim = Val(campo)
    
    
    
    Cadselect = " codtipom='FTG'"
    
    campo = "scafac.fecfactu"
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 20, 21, Devuelve) Then Exit Sub
    End If
                    
    'Compruebo si hay facturas FTG que no han sido contabilizadas
    If HayRegParaInforme("scafac", Cadselect, True) Then
        'Existen registros anterior pendientes de contabilizar
        MsgBox "Existen facturas FTG que no han sido contabilizadas"
    End If
                    
    
    'Compruebo que no hay FTI inferiores a la fecha
    If txtFecha(20).Text <> "" Then
        cadNomRPT = "codtipom = 'FTI' and intconta=0 and fecfactu<'" & Format(txtFecha(20).Text, FormatoFecha) & "'"
        If HayRegParaInforme("scafac", cadNomRPT, True) Then
            MsgBox "Existen Tickets pendientes de contabilizar con fecha inferior a la solicitada", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    campo = DevuelveDesdeBD(conAri, "codclien", "spatpvg", "codigo", "1", "N")
    If campo = "" Then
        MsgBox "No se ha encotrnado el cliente ""varios""", vbExclamation
        Exit Sub
    End If
    NumRegElim = Val(campo)
    
    
    
    
    'Monto la select de las facturas FTI
    Cadselect = " intconta = 0 and codtipom='FTI'"
    campo = "scafac.fecfactu"
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 20, 21, Devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("scafac", Cadselect) Then Exit Sub
            
            
            
    'Si la contbilizacion es menusal, voy a ver si las fechas estan en el mismo mes
    'Si es mas de un mes NO dejo continuar
    If Me.optTick(1).Value Then
        Set miRsAux = New ADODB.Recordset
        miSQL = "Select distinct(fecfactu) from scafac WHERE " & Cadselect & " ORDER BY fecfactu"
        miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        NumParam = 0
        If Not miRsAux.EOF Then
            miSQL = Format(miRsAux!FecFactu, "mmyyyy")
            miRsAux.MoveLast
            campo = Format(miRsAux!FecFactu, "mmyyyy")
            If miSQL <> campo Then
                MsgBox "Las fechas de los tickets a contabilizar NO son del mismo mes. " & miSQL & " " & campo, vbExclamation
                NumParam = 1
            End If
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        If NumParam = 1 Then Exit Sub
    End If
    
    
    'Hay datos. Hago la pregunta
    campo = "Va a contabilizar los tickets agrupados. " & vbCrLf & "Se generará una factura "
    If Me.optTick(1).Value Then
        'Va a cojer un mes. Avisaremos que el periodo de facturacion es superior a un mes
        campo = campo & "con fecha: " & txtFecha(21).Text
    Else
        campo = campo & "por dia"
    End If
    campo = campo & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(campo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            
            
            
    'Si tiene registros hare la contabilizacion
    DesBloqueoManual ("GT")
    If Not BloqueoManual("GT", "1") Then
        MsgBox "Proceso inciado por otro usuario.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Label5.Caption = "Inicio proceso facturacion/contabilizacion"
    
    
    Set miRsAux = New ADODB.Recordset
    
    'En numregelim llevo el codclien de clivarios
    HacerFacturaTICKETS NumRegElim
    
    Set miRsAux = Nothing
    
    'Liberamos el bloqueo
    DesBloqueoManual ("GT")

    Espera 0.5

    
End Sub

Private Sub cmdDeclaraAlmazara_Click()

    If cboMes(0).ListIndex < 0 Then Exit Sub
    If Me.txtNumeroEntero(4).Text = "" Then Exit Sub
    If Val(Me.txtNumeroEntero(4).Text) > 3000 Then Exit Sub
    
    InicializarVbles
    Screen.MousePointer = vbHourglass
    
    If GenerarListadoAlmazara Then
    
    
    End If
    
    Screen.MousePointer = vbDefault
    
    
        
        
End Sub

Private Sub cmdEcoEnves_Click()
    'Vamos a ver el conjunto de albaranes para pasar
    InicializarVbles
    Devuelve = ""
    campo = ""
    If Me.optEcoenves(0).Value Then
        If txtFecha(34).Text <> "" Or txtFecha(35).Text <> "" Then
            campo = "slifac.fecfactu"
            If Not PonerDesdeHasta(campo, "F", 34, 35, Devuelve) Then Exit Sub
        End If
        
        If Not GenerarDatosEncoenves Then Exit Sub
    
         InicializarVbles
         
    Else
        'Listado facturas
        If txtFecha(34).Text <> "" Or txtFecha(35).Text <> "" Then
            campo = "{slifac.fecfactu}"
            If Not PonerDesdeHasta(campo, "F", 34, 35, Devuelve) Then Exit Sub
        End If
        
        
        
        
        Cadparam = Cadparam & "ElPuntoVerde=""" & vParamAplic.ArtReciclado & """|"
        NumParam = NumParam + 1
    End If
    Devuelve = AnyadirParametroDH("Fecha: ", txtFecha(34), txtFecha(35), Nothing, Nothing)
  

    
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    Cadparam = Cadparam & "pDHFecha= """ & Devuelve & """|"
    NumParam = 2
    




    If Me.optEcoenves(0).Value Then
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
        cadNomRPT = "rEcoenves.rpt"
        campo = "Ecoembes"
    Else
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & "({slifac.codartic} = '" & vParamAplic.ArtReciclado & "')"
        campo = "Ecoembes. List. fras."
    
    
        cadNomRPT = "rEcoEnvesFRA.rpt"
    End If
    
    conSubRPT = False
    LlamarImprimir2 campo
        
    
End Sub

Private Sub cmdEstadisticaReparacionTecnico_Click()

    If Me.txtTrab(0).Text = "" Then
        MsgBox "Seleccione un técnico", vbExclamation
        Exit Sub
    End If
    Cadselect = "schrep.codtrab2 = " & txtTrab(0).Text

    'Ya tenemos el tecnico. Miramos las fechas
    If txtFecha(2).Text <> "" Or txtFecha(3).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        campo = "schrep.fecrepar"
        If Not PonerDesdeHasta(campo, "F", 2, 3, Devuelve) Then Exit Sub
        'Aqui lo añadiremos a  cadparam
        
    End If
    
    
    
    
    Screen.MousePointer = vbHourglass
   
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    'Aqui iremos grabanod los datos.
    EstadisticaReparacionTecnico
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato a mostrar", vbExclamation
        Exit Sub
    End If
    
    
    'Llegados aqui imprimiremos los registros
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    Cadparam = Cadparam & "pDHFecha= "" Técnico: " & txtTrab(0).Text & " - " & Me.txtDescTra(0).Text & """|"
    NumParam = 2
    campo = ""
    If txtFecha(2).Text <> "" Then campo = "     Desde " & txtFecha(2).Text
    If txtFecha(3).Text <> "" Then campo = campo & "      Hasta " & txtFecha(3).Text
    If campo <> "" Then
        NumParam = 3
        campo = "pDHCliente= """ & Trim(campo) & """|"
        Cadparam = Cadparam & campo
    End If
    cadFormula = "{tmpnlotes.codusu}=" & vUsu.Codigo

    cadNomRPT = "rRepEstadisticaTec.rpt"
    conSubRPT = False
    LlamarImprimir2
    
End Sub

'------------------------------------------------------------------
'            F A C T U R A S     P R O V E E      S O C I O S
'------------------------------------------------------------------
Private Sub cmdFacProv_Click()
Dim Conjunto As Collection
Dim TipoM As CTiposMov
    'Comprobaciones iniciales
    Cadparam = ""
    If txtFecha(17).Text = "" Then Cadparam = Cadparam & "- fecha factura" & vbCrLf
    If txtBancoPr(0).Text = "" Then Cadparam = Cadparam & "- banco propio" & vbCrLf
    If txtForpa(0).Text = "" Then Cadparam = Cadparam & "- forma de pago" & vbCrLf
    If txtTrab(1).Text = "" Then Cadparam = Cadparam & "- trabajador" & vbCrLf

    Devuelve = ""
    If vParamAplic.PorReten > 0 Then Devuelve = "D"
    If vParamAplic.CtaReten = "" Xor Devuelve = "" Then Cadparam = Cadparam & vbCrLf & "- Falta configurar cta retencion -  % retencion en parametros"
    If Cadparam <> "" Then
        Cadparam = "Campos requeridos: " & vbCrLf & vbCrLf & Cadparam
        MsgBox Cadparam, vbExclamation
        Cadparam = ""
        Exit Sub
    End If
    
    'Tipo de moviemiento de facturas liqueidacion proveedores
    Set TipoM = New CTiposMov
    If Not TipoM.Leer("FLQ") Then  'tipo de movimiento FLQ
        MsgBox "No se puede continuar sin el tipo de moviemiento: FLQ", vbExclamation
        Exit Sub
    End If
    
    'Comprobaciones POSTERIORES ;)
    'Si la fecha esta en correctos
    'FALTA###
    
    
    
    'Cargo en ImpTeo el valor del porcentaje rea
    Devuelve = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA)
    If Devuelve = "" Then
        'ERROR con el tipo de IVA REA
        MsgBox "Tipo de IVA REA no configurado en parametros, o no existe", vbExclamation
        Exit Sub
    End If
    ImpTeo = CCur(Devuelve)
    
    
    
    
    'Vamos a ver el conjunto de albaranes para pasar
    InicializarVbles
    Devuelve = ""
    
    
    'Cadena obligada. Los proveedores , el tipo tiene que ser el 3: REA
    Cadselect = " {scaalp.codprove}=  {sprove.codprove}  AND {sprove.tipprove}= 3 "
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(13).Text <> "" Or txtFecha(14).Text <> "" Then
        campo = "{scaalp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 13, 14, Devuelve) Then Exit Sub
    End If
    
    If txtCodProve(4).Text <> "" Or txtCodProve(5).Text <> "" Then
        campo = "{scaalp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 4, 5, Devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(0).Text <> "" Or txtNumAlbar(1).Text <> "" Then
        campo = "{scaalp.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 0, 1, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    Cadselect = " scaalp,sprove WHERE " & Cadselect
    
    
    
    If Not HayRegParaInforme(Cadselect, "", True) Then
        MsgBox "No hay albaranes para facturar con estos valores", vbExclamation
        Exit Sub
    Else
        'llegado aqui preguntamos si desea continuar
        cadFrom = "Seguro que desea continuar con la generacion de las facturas de liquidación?"
        If MsgBox(cadFrom, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    'Monto el SQL para saber que albaranes facturo
    Screen.MousePointer = vbHourglass
    cadFrom = "Select sprove.codprove,albaranxfactura FROM " & Cadselect & " GROUP by 1,2 ORDER BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Conjunto = New Collection
    While Not miRsAux.EOF
        Conjunto.Add miRsAux!codProve & "|" & miRsAux!albaranxfactura & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'AHora vamos a ir facturando los diversos proveedores
    For IndiceImg = 1 To Conjunto.Count
        'Facturamos al proveedor
        FacturarProveedor CLng(RecuperaValor(Conjunto.Item(IndiceImg), 1)), Val(RecuperaValor(Conjunto.Item(IndiceImg), 2)) = 1, TipoM
    Next IndiceImg
    
    Label1.Caption = ""
    Set TipoM = Nothing
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Proceso finalizado", vbExclamation
End Sub


Private Sub FacturarProveedor(codProve As Long, albaranxfactura As Boolean, ByRef Ctip As CTiposMov)
Dim vFactu As CFacturaCom
Dim vProve As CProveedor
Dim cad As String
Dim RA As ADODB.Recordset
Dim ColFacturar As Collection
Dim J As Integer



    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(CStr(codProve)) Then
        Label1.Caption = "Error leyendo proveedor: " & codProve
        Me.Refresh
        DoEvents
        Espera 1
        Exit Sub
    End If
    
    
    Label1.Caption = "ALbaranes a facturar proveedor :        " & vProve.Nombre
    Label1.Refresh

    cad = "Select scaalp.numalbar,scaalp.fechaalb FROM " & Cadselect & " AND scaalp.codprove = " & codProve
    cad = cad & " ORDER BY scaalp.fechaalb,scaalp.numalbar"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = "codprove = " & codProve & " AND "
    Set ColFacturar = New Collection
    cadNomRPT = ""
    While Not miRsAux.EOF
        cad = "numalbar = '" & DevNombreSQL(miRsAux!NumAlbar) & "' AND fechaalb = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
        If albaranxfactura Then
            cad = cadFrom & cad
            ColFacturar.Add cad
        Else
            If cadNomRPT <> "" Then cadNomRPT = cadNomRPT & " OR "
            cadNomRPT = cadNomRPT & "(" & cad & ")"
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Not albaranxfactura Then
        cad = cadFrom & "(" & cadNomRPT & ")"
        ColFacturar.Add cad
    End If
    
    
    
   'AHORA YA TENGO EN Colfactuar el conjunto de labaraens y/o facturas
    For J = 1 To ColFacturar.Count
       'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = vProve.Codigo
        vFactu.NumFactu = Ctip.contador + 1
        vFactu.FecFactu = txtFecha(17).Text
        vFactu.FecRecep = txtFecha(17).Text
        vFactu.Trabajador = txtTrab(1).Text
        vFactu.BancoPr = txtBancoPr(0).Text
        
        vFactu.ForPago = txtForpa(0).Text
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0

        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan

        
        
    
        'Obtengo los totales mediante el cadselect
        cad = "Select sum(importel) FROM slialp WHERE " & ColFacturar.Item(J)
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            ImpTot = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        
        
        vFactu.BrutoFac = ImpTot
        vFactu.BaseIVA1 = ImpTot
        vFactu.TipoIVA1 = vParamAplic.IVA_REA
        vFactu.PorceIVA1 = ImpTeo
        ImpTot = Round2((ImpTot * ImpTeo) / 100, 2)
        vFactu.ImpIVA1 = ImpTot
        ImpTot = vFactu.BrutoFac + ImpTot  'Base + IVA
        
        'Retencion
        vFactu.TipoRet = 1
        vFactu.PorRet = vParamAplic.PorReten
        vFactu.ImpRet2 = Round2((ImpTot * vFactu.PorRet) / 100, 2)
            
        
        vFactu.TotalFac = vFactu.BrutoFac + vFactu.ImpIVA1 - vFactu.ImpRet2
        

         'El select
         cad = ColFacturar.Item(J)
         
         If Not vFactu.TraspasoAlbaranesAFactura(cad, (chkFacturPorv(1).Value = 1), (chkFacturPorv(0).Value = 1), True) Then
            'Para salir y finalizar el procesode facturacion de el proveedor
            cad = "Finalizacion de la facturacion para: " & vProve.Nombre & vbCrLf
            cad = cad & "Proceso: " & J & " / " & ColFacturar.Count & vbCrLf
            cad = cad & vbCrLf & "SQL: " & ColFacturar.Item(J)
            MsgBox cad, vbExclamation
            J = ColFacturar.Count + 1  'Para que se salga
        Else
            'incremento el contador de facturas
            Ctip.IncrementarContador Ctip.TipoMovimiento
        End If
'        Set vFactu = Nothing
'

    Next J

    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation

End Sub

Private Sub cmdFacturaMov_Click()
Dim AlbaranesGenerados As Collection
Dim MensajeError As String

    'Facuracion recargas moviles
    campo = ""
    If txtCliente(2).Text = "" Then campo = campo & " - Cliente" & vbCrLf
    If txtArticulo(0).Text = "" Then campo = campo & " - Articulo" & vbCrLf
    If txtBancoPr(1).Text = "" Or txtDescBancoPr(1).Text = "" Then campo = campo & " - Bancos propios" & vbCrLf
    
    
    If campo <> "" Then
        campo = "Campos requeridos : " & vbCrLf & campo
        MsgBox campo, vbExclamation
        Exit Sub
    End If
    
    'Alguna comprobacion mas
    
    If txtFecha(8).Text = "" Then
        MsgBox "Ponga la fecha de facturación", vbExclamation
        Exit Sub
    End If
    
    If vEmpresa.TieneAnalitica Then
        'Comprobar que existen todos los centros de coste en los datos a facturar
        'FALTA###
        
    End If
    
    InicializarVbles
    
    
    
    
    'Obtengo el que sera el ultimo registro insertado hasta ahora.
    campo = SugerirCodigoSiguienteStr("stelefonia", "id")
    NumRegElim = Val(campo)
    
    
    
    Cadselect = " id < " & NumRegElim & " AND Facturado = 0 "
    
    campo = "stelefonia.fecha"
    If txtFecha(6).Text <> "" Or txtFecha(7).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 6, 7, Devuelve) Then Exit Sub
    End If
    
    
    
            
                    
                    
                    
    'Compruebo si tiene registros
    If Not HayRegParaInforme("stelefonia", Cadselect) Then Exit Sub
            
            
            
    'Si tiene registros hare la contabilizacion
    
    DesBloqueoManual ("Telf")
    If Not BloqueoManual("Telf", "1") Then
        MsgBox "Proceso inciado por otro usuario.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblIndicadorT.Caption = "Inicio proceso facturacion"
    pb1.Value = 0
    
    Set AlbaranesGenerados = New Collection
    MensajeError = ""
    
    HacerFacturacionTelefonia AlbaranesGenerados, MensajeError
    
    If AlbaranesGenerados.Count > 0 Then
        
        If MensajeError <> "" Then
            campo = "Se generaron  " & AlbaranesGenerados.Count & " albaranes. "
            campo = campo & vbCrLf & vbCrLf & " ERROR GENERANDO ALBARANES" & vbCrLf & MensajeError
            MsgBox campo, vbInformation
        End If
        
        campo = ""
        For NumRegElim = 1 To AlbaranesGenerados.Count
            If campo <> "" Then campo = campo & ","
            campo = campo & AlbaranesGenerados.Item(NumRegElim)
        Next NumRegElim
        campo = "scaalb.codtipom = 'ALV' AND scaalb.numalbar IN (" & campo & ")"
        miSQL = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
        miSQL = miSQL & " WHERE " & campo
        TraspasoAlbaranesFacturas miSQL, campo, CStr(Now), txtBancoPr(1).Text, pb1, lblIndicadorT, True, "ALV", ""
        
    Else
        MensajeError = "No se ha generado ningun albaran" & MensajeError
        MsgBox MensajeError, vbExclamation
        
    End If
    'Liberamos el bloqueo
    DesBloqueoManual ("Telf")
    lblIndicadorT.Caption = "Proceso finalizado"
    Espera 0.3
End Sub

Private Sub cmdGenerAlbOliva_Click()
Dim b As Boolean
    
    miSQL = "¿Generar los albaranes?"
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    b = GenerarAlbaranesOliva
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If b Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
End Sub

Private Sub cmdGeneraTO_Click()

    'devolveremos los campos al formulario de OT
    '
    
    
    miSQL = ""
    'Las amrcas
    If List1.ListCount > 0 Then
        For NumRegElim = 0 To List1.ListCount - 1
            miSQL = miSQL & ", " & List1.ItemData(NumRegElim)
        Next
        miSQL = Mid(miSQL, 2)
        miSQL = " AND codmarca IN (" & miSQL & ")"
    End If
    'Voy a montar la select para el recalculo de precios
    miSQL = " conjunto = 1 " & miSQL
    If txtArticulo(5).Text <> "" Then miSQL = miSQL & " AND sartic.codartic >='" & DevNombreSQL(txtArticulo(5).Text) & "'"
    If txtArticulo(6).Text <> "" Then miSQL = miSQL & " AND sartic.codartic <='" & DevNombreSQL(txtArticulo(6).Text) & "'"
    
    If txtFamilia(0).Text <> "" Then miSQL = miSQL & " AND sartic.codfamia >= " & txtFamilia(0).Text
    If txtFamilia(1).Text <> "" Then miSQL = miSQL & " AND sartic.codfamia <= " & txtFamilia(1).Text
    
            
    
    
    'Meto el link con la tabla sarti
    
    miSQL = " from sarti1,sartic where sarti1.codartic =sartic.codartic and " & miSQL
    miSQL = miSQL & " AND codarti1 in (select codartic from olitmpto where codusu = " & vUsu.Codigo & ") group by sartic.codartic"
    CadenaDesdeOtroForm = miSQL
    'Salimos
    Unload Me
    
End Sub

Private Sub LeerEscribirDatosCarta(Leer As Boolean)
Dim NF As Integer

    miSQL = App.Path & "\CartaTO.dat"
    If vParamAplic.EsAVAB Then miSQL = App.Path & "\CartaTOAVAB.dat"


    On Error GoTo EGuar
    
    NF = FreeFile
    If Leer Then
        Devuelve = ""
        If Dir(miSQL, vbArchive) <> "" Then
            'No existe el archivo
            Open miSQL For Input As #NF
            While Not EOF(NF)
                Line Input #NF, miSQL
                Devuelve = Devuelve & miSQL & vbCrLf
            Wend
            Close #NF
        End If
        
        For NumRegElim = 0 To txtCarta.Count - 1
            txtCarta(NumRegElim).Text = RecuperaValor(Devuelve, NumRegElim + 1)
        Next
        
        
        
    Else
        Devuelve = ""
        For NumRegElim = 0 To txtCarta.Count - 1
            Devuelve = Devuelve & txtCarta(NumRegElim).Text & "|"
        Next
        
            Open miSQL For Output As #NF
            Print #NF, Devuelve
            Close #NF
            
    End If
    
    Exit Sub
EGuar:
    MuestraError Err.Number
End Sub




Private Sub cmdGuardar_Click()
    LeerEscribirDatosCarta False
End Sub

Private Sub cmdImpresion4Tonda_Click()
    
    InicializarVbles
    
    'If Not PonerParamRPT(26, Cadparam, NumParam, cadNomRPT) Then Exit Sub
    cadFrom = CadenaDesdeOtroForm
    Cadselect = "{scaalb.codtipom}='ALV' and {scaalb.numalbar} =" & CadenaDesdeOtroForm
    cadFormula = "(" & Cadselect & ")"
    
    
    'Montamos el select para los registros
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    campo = "scaalb WHERE " & Cadselect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    Cadparam = ""
    NumParam = 0
    conSubRPT = False
    
    If Me.optTransporte(1).Value Then
        cadTitulo = "Documento de control"
        cadNomRPT = "vallAlbaDocControl.rpt"
    Else
        cadTitulo = "Albaran de entrega y lista de precios"
        cadNomRPT = "vallAlbaEntrega.rpt"
    End If
    
    
    LlamarImprimir2 cadTitulo
    CadenaDesdeOtroForm = cadFrom 'por si quiere volver a imprimir
End Sub

Private Sub cmdImprimirFac_Click()
    'Impresion de las facturas de proveedores
    'es decir , para casos de cooperativas en las cuales el socio
    'es el que nos emite la factura a nosotros (ej TERRASANA)
    
    
    
    
    
    InicializarVbles
    
    If Not PonerParamRPT(26, Cadparam, NumParam, cadNomRPT) Then Exit Sub
    
    Cadselect = "{sprove.tipprove}=3"   'Estos proveedores son los REA que luego
    cadFormula = "(" & Cadselect & ")"                                    'les emitire SUS facturas
    If txtFecha(15).Text <> "" Or txtFecha(16).Text <> "" Then
        campo = "{scafpc.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 15, 16, Devuelve) Then Exit Sub
    End If
    
    If txtCodProve(6).Text <> "" Or txtCodProve(7).Text <> "" Then
        campo = "{scafpc.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 6, 7, Devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(2).Text <> "" Or txtNumAlbar(3).Text <> "" Then
        campo = "{scafpc.numfactu}"
        If Not PonerDesdeHasta(campo, "ALP", 2, 3, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    campo = "scafpc,sprove WHERE scafpc.codprove=sprove.codprove AND " & Cadselect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    LlamarImprimir2
End Sub

Private Function ImprimirOfertaBonito() As Boolean
    On Error GoTo EImprimirOfertaBonito
    ImprimirOfertaBonito = False
    
    miSQL = "DELETE FROM olitmpimprto WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    Devuelve = ""
    
    For NumRegElim = 0 To List3.ListCount - 1
        If List3.Selected(NumRegElim) Then
            Devuelve = "O"
            Exit For
        End If
    Next NumRegElim
    
    If Devuelve = "" Then
        MsgBox "Selecciona algun cliente", vbExclamation
        Exit Function
    End If
    
    Devuelve = "insert into `olitmpimprto` (`codusu`,`codTO`,fecha,"
    Devuelve = Devuelve & "asunto,`parrafo1`,`parrafo2`,`parrafo3`,"
    Devuelve = Devuelve & "`despedida`,`firmado`,`codclien` )"
    Devuelve = Devuelve & " VALUES (" & vUsu.Codigo & "," & Me.txtTarifa.Text
    If txtFecha(32).Text = "" Then
        miSQL = Format(Now, FormatoFecha)
    Else
        miSQL = Format(txtFecha(32).Text, FormatoFecha)
    End If
    Devuelve = Devuelve & ", '" & miSQL & "'"
    miSQL = ""
    If Me.chkCartaTO.Value = 0 Then
        'No lleva carta
        Devuelve = Devuelve & ",NULL,NULL,NULL,NULL,NULL,NULL"
        
    Else
        For NumRegElim = 0 To 5
            Devuelve = Devuelve & "," & DBSet(txtCarta(NumRegElim), "T")
        Next
    End If
    'PARA CAD cliente insertamos en la tabla
    For NumRegElim = 0 To List3.ListCount - 1
        If List3.Selected(NumRegElim) Then
            miSQL = Devuelve & "," & List3.ItemData(NumRegElim) & ")"
            conn.Execute miSQL
        End If
    Next
    
    ImprimirOfertaBonito = True
    Exit Function
EImprimirOfertaBonito:
        MuestraError Err.Number, "Impresion"
    
    
End Function

Private Sub InsertaItem()
    'Numregelim tiene el dato
    
End Sub

Private Sub cmdImprimirTOs_Click()
    InicializarVbles
    Screen.MousePointer = vbHourglass
    If ImprimirOfertaBonito Then
        cadNomRPT = ""
        Cadparam = "|ImprimeCarta= " & Abs(Me.chkCartaTO.Value) & "|"
        NumParam = 1
        cadFormula = "{olitmpimprto.codusu}=" & vUsu.Codigo
        LlamarImprimir2
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdInvenAceite_Click()
Dim R As ADODB.Recordset
Dim Cantidad As Currency
Dim vFamia As String


    If txtFecha(33).Text = "" Then
        MsgBox "Ponga la fecha de inventario", vbExclamation
        PonerFoco txtFecha(33)
        Exit Sub
    End If
    
    miSQL = "DELETE FROM tmpstockfec where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    InicializarVbles
    
    Set miRsAux = New ADODB.Recordset
    Codigo = "select codfamia from sartic where conjunto=1 group by 1"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    While Not miRsAux.EOF
        Codigo = Codigo & ", " & miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'cadSelect = " sartic.codartic = slifac.codartic and sartic.conjunto = 1 "
    vFamia = ""
    cadFormula = "" ' " {sartic.conjunto} = 1 "
    Cadselect = " sartic.codartic = salmac.codartic "
    If Codigo <> "" Then
        Codigo = Mid(Codigo, 2)
        vFamia = " and sartic.codfamia in (" & Codigo & ")"
        'Para la formula
        '        ({sartic.codfamia} in [1, 3, 6]) ejem
        cadFormula = " ({sartic.codfamia} in [" & Codigo & "])"
    End If
    Codigo = ""
    'Fecha inventario
    Codigo = Codigo & Space(20) & "Fecha inventario: " & txtFecha(33).Text
    campo = "{salmac.fechainv}"
    If Not PonerDesdeHasta(campo, "F", 33, 33, Devuelve) Then Exit Sub
    

    If vUsu.TrabajadorB Then
        If chkConsolidado(2).Value = 0 Then
            'SOLO QUIERE ver el almacen 2
            Codigo = Codigo & Space(20) & "Alma*"
            If Not AnyadirAFormula(Cadselect, "{salmac.codalmac}=" & vParamAplic.AlmacenB) Then Exit Sub
            cadFormula = cadFormula & " AND {salmac.codalmac}=" & vParamAplic.AlmacenB
        End If
    Else
        cadFormula = cadFormula & " AND {salmac.codalmac}=1"
        Cadselect = Cadselect & " AND {salmac.codalmac}=1"
    End If
    




    Codigo = Trim(Codigo)
    Cadparam = "pDH1= """ & Codigo & """|"
    NumParam = 1


    cadFrom = " salmac,sartic"
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    If Not HayRegParaInforme(cadFrom, Cadselect & vFamia) Then Exit Sub

    
    
    'El informe se divide en dos.
    '1.- slifac donde los articulos son articulos de ventas(tienen componentes)
    '2.- resumen de los componentes que son materias primas
    '     tanto los que se hayan facturado directamente (kilos)
    '     como los que entra por que es componente de uno de venta
    
    'Entonces, el primer trozo del informe sale directamente de las columnas. El segundo
    'saldra de una tabla temporal  tmpstockfec
    ' tendra usuario, codartic codalmac  -> 0 Vendidos directamente   1- Como compenente    stcok  la cantidad en cuestion
    
    
    'miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) select " & vUsu.Codigo & ",slifac.codartic,0,cantidad * factorconversion from slifac,sartic"
    miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) "
    '                                                       * factorconversion
    miSQL = miSQL & " select " & vUsu.Codigo & ",salmac.codartic,0,sum(stockinv)  from salmac,sartic"
    miSQL = miSQL & " where salmac.codartic=sartic.codartic and factorconversion<>1"
    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
    miSQL = miSQL & " group by salmac.codartic"
    conn.Execute miSQL    'Para el punto materias primas, los vendidos directamente
    
    
    'Cojeremos un cursor con todos las materias primas e iremos insertandolas en la tmpstock
    '-----------------
    miSQL = "select salmac.codartic,sum(stockinv) cantidad from salmac,sartic where salmac.codartic=sartic.codartic "
    miSQL = miSQL & " and conjunto=1 "
    'Las fechas
    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
    miSQL = miSQL & " group by salmac.codartic "
    
    Set R = New ADODB.Recordset
    Devuelve = "|"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
            'Para cada elemento facturado que tiene componentes, vere de sus componentes cual es el de mataria prima y calcular su cantidad
            miSQL = "select sarti1.codarti1,cantidad from sarti1,sartic where  sarti1.codarti1=sartic.codartic and factorconversion<>1"
            miSQL = miSQL & " AND sarti1.codartic =" & DBSet(miRsAux!codartic, "T")
            R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            'If Mid(miRsAux!codArtic, 1, 9) = "002700090" Then Stop
            
            
            
            Cantidad = DBLet(miRsAux!Cantidad, "N")

            
            'Para no tener que hacer un select para saber si ya ha sido insertado en tmpstock, utilizar
            'el string cadSelect para ir metiendo los ya insertados.
            While Not R.EOF
                'El articulo en cuestion
                miSQL = "|" & R!codarti1 & "|"
                Cantidad = Cantidad * R!Cantidad   'Esta es la cantidad nueva
                campo = TransformaComasPuntos(CStr(Cantidad))
                If InStr(1, Devuelve, miSQL) > 0 Then
                    'Ya esta insertado. Es un UPDATE
                    miSQL = "UPDATE tmpstockfec SET stock=stock + " & campo
                    miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " and codartic = " & DBSet(R!codarti1, "T")
                    miSQL = miSQL & " AND codalmac= 1"
                Else
                    Devuelve = Devuelve & R!codarti1 & "|"
                    miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock)  VALUES (" & vUsu.Codigo & "," & DBSet(R!codarti1, "T")
                    miSQL = miSQL & ",1," & campo & ")"
                    
                End If
                conn.Execute miSQL
                'No deberia haber mas (seria un coupage)
                R.MoveNext
            Wend
            R.Close
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set R = Nothing


    


    Cadparam = Cadparam & "pCodUsu= " & vUsu.Codigo & "|"
    'cadFormula = "{tmpnlotes.codusu}=" & vUsu.Codigo
    
    
    'Añadimos nombre empresa
    Cadparam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|" & Cadparam
    NumParam = NumParam + 1
    
    With frmImprimir
        .ConSubInforme = False
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .NombreRPT = DevuelveNombreReport(35) '"rInvenMor.rp"
        .Titulo = "Listado"
        .opcion = 2016
        .Show vbModal
    End With

End Sub

Private Sub cmdLineaExtraProd_Click()
    CadenaDesdeOtroForm = Text2(0).Text & "|" & Text2(1).Text & "|"
    Unload Me
End Sub

Private Sub cmdListadoTO_Click()
Dim C2 As String



    'Listado TO
    InicializarVbles
    Cadselect = " olitarifaoferta.codigo=olitarifaofertalin.codigo "
    Devuelve = ""
    C2 = ""
    
        
    If txtCliente(5).Text <> "" Or txtCliente(6).Text <> "" Then
        
        campo = "{olitarifaoferta.codclien}"
        If Not PonerDesdeHasta(campo, "CLI", 5, 6, Devuelve) Then Exit Sub
        C2 = AnyadirParametroDH("Cliente: ", txtCliente(5), txtCliente(6), txtDescClie(5), Me.txtDescClie(6))
    End If
    
    'Codigo
    If txtNumeroEntero(0).Text <> "" Or txtNumeroEntero(1).Text <> "" Then
        campo = "{olitarifaoferta.codigo}"
        'N_E: numero entero
        If Not PonerDesdeHasta(campo, "N_E", 0, 1, Devuelve) Then Exit Sub
        C2 = C2 & Space(20) & AnyadirParametroDH("Codigo TO: ", txtNumeroEntero(0), txtNumeroEntero(1), Nothing, Nothing)
    End If
    C2 = Trim(C2)
    Cadparam = "pDH1= """ & C2 & """|"
    NumParam = 1
    
    
    C2 = ""
    If txtFecha(24).Text <> "" Or txtFecha(25).Text <> "" Then
        C2 = C2 & Space(20) & AnyadirParametroDH("Fecha inicio: ", txtFecha(24), txtFecha(25), Nothing, Nothing)
        campo = "{olitarifaoferta.fechaini}"
        If Not PonerDesdeHasta(campo, "F", 24, 25, Devuelve) Then Exit Sub
    End If
    
    If txtFecha(26).Text <> "" Or txtFecha(27).Text <> "" Then
        C2 = C2 & Space(20) & AnyadirParametroDH("Fecha fin: ", txtFecha(26), txtFecha(27), Nothing, Nothing)
        campo = "{olitarifaoferta.fechafin}"
        If Not PonerDesdeHasta(campo, "F", 26, 27, Devuelve) Then Exit Sub
    End If

      
    If txtArticulo(8).Text <> "" Or txtArticulo(9).Text <> "" Then
        C2 = C2 & Space(20) & AnyadirParametroDH("Articulo: ", txtArticulo(8), txtArticulo(9), txtDescArticulo(8), txtDescArticulo(9))
        campo = "{olitarifaofertalin.codartic}"
        If Not PonerDesdeHasta(campo, "ART", 8, 9, Devuelve) Then Exit Sub
    End If
    
    'Parametro Desde/Hasta
    If C2 <> "" Then
        C2 = Trim(C2)
        Cadparam = Cadparam & "pDH2= """ & C2 & """|"
        NumParam = NumParam + 1
    End If
    
    
    If Me.FrameTosTapa.visible Then
        'Tarifa-oferta
        C2 = ">"
    Else
        'solo tarifa
        C2 = "<"
    End If
    C2 = C2 & " 100000"
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & " {olitarifaoferta.codigo} " & C2
    Cadselect = Cadselect & " AND olitarifaoferta.codigo " & C2
    'Llegados aqui tenemos el select
    
    
    
    'Si esta marcado el chk de mataria prima reemplazaremos en cadselect
    '  y cadformula olitarifaofertalin por olitarifaofertalin2
    If Me.chkMatPrima.Value = 1 Then
        cadFrom = " olitarifaoferta,olitarifaofertalin2"
        cadFormula = Replace(cadFormula, "olitarifaofertalin.", "olitarifaofertalin2.")
        Cadselect = Replace(Cadselect, "olitarifaofertalin.", "olitarifaofertalin2.")
    Else
        cadFrom = " olitarifaoferta,olitarifaofertalin"
    End If
    
    If Not HayRegParaInforme(cadFrom, Cadselect) Then Exit Sub
    
    
    Cadparam = Cadparam & "ElOrden=" & Abs(Me.optOrdenTO(1).Value) & "|"
    NumParam = NumParam + 1
    
    Cadparam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|" & Cadparam
    NumParam = NumParam + 1
    
    
    
    With frmImprimir
        .ConSubInforme = False
        .FormulaSeleccion = cadFormula
        If Me.chkMatPrima.Value = 0 Then
            If Me.optListadoTO(0).Value Then
                .NombreRPT = "rToCodigo"
            ElseIf Me.optListadoTO(1).Value Then
                .NombreRPT = "rToArticulo"
            ElseIf Me.optListadoTO(2).Value Then
                .NombreRPT = "rToCliente"
            Else
                .NombreRPT = "rToClienteLogo"
            End If
            'TARIFAS
            If Not Me.FrameTosTapa.visible Then .NombreRPT = .NombreRPT & "T"
            
        Else
            .NombreRPT = "rTomatePrima"
        End If
        .NombreRPT = .NombreRPT & ".rpt"
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .Titulo = "Listado"
        .opcion = 2016
        .Show vbModal
    End With
    
    
    
End Sub

Private Sub cmdListAlbVall_Click()
    
    InicializarVbles
    
    
'
    If txtCodProve(14).Text <> "" Or txtCodProve(15).Text <> "" Then
        Devuelve = "|pdh= ""Coop:" & " "
        campo = "{vallentradacamion.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 14, 15, Devuelve) Then Exit Sub
    End If
     
    If Not HayRegParaInforme("vallentradacamion", Cadselect) Then Exit Sub
    
    
    'Truco para que muestre un rpt con la misma opcion de listado (36)
    'El primer pipe llevamos el NOMrpt
    Codigo = "vallEntradaOlivaListadoSin.rpt"
    If Me.chkVarios(1).Value = 1 Then Codigo = "vallEntradaOlivaListado.rpt"
    Cadparam = Codigo & "|" & Cadparam
    
    'Visualizamos report
    LlamarImprimir2
End Sub

Private Sub cmdMarca_Click(Index As Integer)
    If Index = 0 Then
        AñadirMarcas
    Else
        If Me.List1.ListCount = 0 Then Exit Sub
        
        If List1.ListIndex < 0 Then
            MsgBox "Seleccione una marca para eliminar", vbExclamation
            Exit Sub
        End If
        
        If MsgBox("Quitar la marca?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        List1.RemoveItem List1.ListIndex
        
    End If
End Sub


'Private Sub InsertarMarca()
'Dim i As Integer
'    On Error Resume Next
'
'    For i = 0 To List1.ListCount - 1
'        If List1.ItemData(i) = txtMarca(0).Text Then
'            MsgBox "Ya esta insertada", vbExclamation
'            PonerFoco txtMarca(0)
'            Exit Sub
'        End If
'    Next
'
'
'    List1.AddItem Me.txtDescMarca(0).Text & " (" & Me.txtMarca(0).Text & ")"
'    List1.ItemData(List1.NewIndex) = Val(Me.txtMarca(0).Text)
'    Me.txtDescMarca(0).Text = ""
'    Me.txtMarca(0).Text = ""
'    PonerFoco txtMarca(0)
'    If Err.Number <> 0 Then Err.Clear
'End Sub


Private Sub cmdMultibase_Click()

    'Revision caracteres multibase
    NumParam = 0
    For NumRegElim = 1 To Me.lstMultibase.ListCount
        If Me.lstMultibase.Selected(CInt(NumRegElim - 1)) Then NumParam = NumParam + 1
    Next
    
    If NumParam = 0 Then
        MsgBox "Seleccion alguna tabla para cambiar", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Este proceso puede durar mucho tiempo." & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Me.Tag = ""
    Set miRsAux = New ADODB.Recordset
    For NumParam = 0 To Me.lstMultibase.ListCount - 1
        If Me.lstMultibase.Selected(CInt(NumParam)) Then HacerCambiosMultibase CInt(NumParam + 1)
    Next
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If Me.Tag <> "" Then
        Codigo = "Se han realizado los siguientes cambios:" & vbCrLf & vbCrLf & Me.Tag
        Me.Tag = ""
    Else
        Codigo = "Proceso finalizado. No se efectuaron cambios"
    End If
    MsgBox Codigo, vbInformation
End Sub

Private Sub cmdPalets_Click()
Dim cSt As cStock
Dim vT As CTiposMov
    miSQL = ""
    If Me.txtArticulo(12).Text = "" Then miSQL = miSQL & vbCrLf & "-Articulo"
    If Me.txtNumeroEntero(3).Text = "" Then miSQL = miSQL & vbCrLf & "-Cantidad"
    If Me.txtFecha(42).Text = "" Then miSQL = miSQL & vbCrLf & "-Fecha"
    If Me.txtHora(0).Text = "" Then miSQL = miSQL & vbCrLf & "-Hora"
    If optPalets(0).Value Then
        If Me.txtCliente(8).Text = "" Then miSQL = miSQL & vbCrLf & "-Cliente"
    Else
        If Me.txtCodProve(13).Text = "" Then miSQL = miSQL & vbCrLf & "-Proveedor"
    End If
    If miSQL <> "" Then
        MsgBox "Error en campos: " & vbCrLf & miSQL, vbExclamation
        Exit Sub
    End If
    
    
    'Pequeñas comprobaciones
    '................
    'Fecha >= inicio ejercicio
    miSQL = ""
    If Me.txtFecha(42).Text < vEmpresa.FechaIni Then miSQL = miSQL & vbCrLf & "-Fecha fuera ambito"
    cadTitulo = DevuelveDesdeBD(conAri, "tipartic", "sartic", "codartic", txtArticulo(12).Text, "T")
    If cadTitulo <> "31" Then miSQL = miSQL & vbCrLf & "-No es articulo de palets"
    If miSQL <> "" Then
        MsgBox "Error en campos: " & vbCrLf & miSQL, vbExclamation
        Exit Sub
    End If
    
    
    'OK. Vamos para alla
    Set cSt = New cStock
    Set vT = New CTiposMov
    vT.Leer "PAL"
    miSQL = Format(vT.ConseguirContador(vT.TipoMovimiento), "0000")
    cSt.Documento = miSQL
    cSt.Cantidad = CInt(txtNumeroEntero(3).Text)
    cSt.codAlmac = 1
    cSt.codartic = txtArticulo(12).Text
    cSt.DetaMov = vT.TipoMovimiento
    
    
    cSt.Fechamov = CDate(txtFecha(42).Text)
    cSt.HoraMov = cSt.Fechamov & " " & txtHora(0).Text
    cSt.Importe = 0
    cSt.LineaDocu = 1
    cadTitulo = cSt.codartic & "|"
    If Me.optPalets(0).Value Then
        cSt.tipoMov = "S"
        cSt.Trabajador = txtCliente(8).Text
        cadTitulo = cadTitulo & cSt.Trabajador & "||"
    Else
        cSt.tipoMov = "E"
        cSt.Trabajador = txtCodProve(13).Text
        cadTitulo = cadTitulo & "|" & cSt.Trabajador & "|"
    End If
    cSt.ActualizarStock False
    vT.IncrementarContador vT.TipoMovimiento
    'Articulo,codclien,codprove
    CadenaDesdeOtroForm = cadTitulo
    Unload Me
End Sub

Private Sub cmdRecalPrSt_Click()
    
    
    
    
    If optRecalPrSt(0).Value Then
        Codigo = "sarti1.codarti1"
        miSQL = "codarti1=sartic.codartic   and factorconversion=1 "
        campo = "sarti1,sartic"
        cadFrom = optRecalPrSt(0).Caption
        cadTitulo = "distinct(codarti1)"
    Else
        Codigo = "sartic.codartic"
        miSQL = "conjunto=1"
        campo = "sartic"
        cadFrom = optRecalPrSt(1).Caption
        cadTitulo = "*"
    End If
    
    
    If txtArticulo(10).Text <> "" Then miSQL = miSQL & " AND " & Codigo & " >='" & DevNombreSQL(txtArticulo(10).Text) & "'"
    If txtArticulo(11).Text <> "" Then miSQL = miSQL & " AND " & Codigo & "  <='" & DevNombreSQL(txtArticulo(11).Text) & "'"
    
    
    Set miRsAux = New ADODB.Recordset
    Codigo = "Select count(" & cadTitulo & ") from " & campo & " WHERE " & miSQL
    
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = "0"
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Codigo = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    
    If Val(Codigo) > 0 Then
    
        Codigo = "Va a actualizar el precio standard de " & Codigo & " articulos(s) de " & cadFrom
        Codigo = Codigo & vbCrLf & "¿Continuar?"
        If MsgBox(Codigo, vbQuestion + vbYesNoCancel) = vbYes Then
            Screen.MousePointer = vbHourglass
            RecalcularPrStandard
        
        End If
    Else
        MsgBox "Ningun datos", vbExclamation
    End If
        
    Set miRsAux = Nothing
    Label3(89).Caption = ""
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdRecargasMov_Click()


    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    Devuelve = ""
    campo = "{stelefonia.fecha}"
    If txtFecha(4).Text <> "" Or txtFecha(5).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHFecha=""Fecha " & Devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 4, 5, Devuelve) Then Exit Sub
    End If

    'Facturados
    Devuelve = ""
    Codigo = ""
    cadFrom = ""
    If Me.cmbRecargaMov(0).ListIndex > 0 Then
        If Me.cmbRecargaMov(0).ListIndex = 1 Then
            Codigo = Codigo & " Pendientes de facturar "
        Else
            Codigo = Codigo & "  Facturadas "
        End If
        campo = "({stelefonia.facturado} = " & cmbRecargaMov(0).ListIndex - 1 & ")"
        cadFrom = "facturado = " & cmbRecargaMov(0).ListIndex - 1
        Devuelve = campo
    End If
    
    'Cobrado
    If Me.cmbRecargaMov(1).ListIndex > 0 Then
        If Me.cmbRecargaMov(1).ListIndex = 1 Then
            Codigo = Codigo & "     Pendientes de cobro "
        Else
            Codigo = Codigo & "     Cobradas "
        End If
        campo = "({stelefonia.cobrado} = " & cmbRecargaMov(1).ListIndex - 1 & ")"
        
        If Devuelve <> "" Then
            Devuelve = Devuelve & " AND "
            cadFrom = cadFrom & " AND "
        End If
        cadFrom = cadFrom & "cobrado = " & cmbRecargaMov(1).ListIndex - 1
        Devuelve = Devuelve & campo
    End If
    
    
    'Tipo
    If txtRecargaMov(0).Text <> "" Then
        campo = "({stelefonia.tipo} = '" & txtRecargaMov(0).Text & "')"
        
        If Devuelve <> "" Then
            Devuelve = Devuelve & " AND "
            cadFrom = cadFrom & " AND "
        End If
        cadFrom = cadFrom & "tipo = """ & txtRecargaMov(0).Text & """"
        Devuelve = Devuelve & campo
    End If
    
    If Devuelve <> "" Then
        Cadparam = Cadparam & "pDHCliente= """ & Trim(Codigo) & """|"
        NumParam = NumParam + 1
        If Cadselect <> "" Then Cadselect = Cadselect & " AND "
        Cadselect = Cadselect & cadFrom
        AnyadirAFormula cadFormula, Devuelve
    End If
    
    
    
    
        
    
    
    If Not HayRegParaInforme("stelefonia", Cadselect) Then Exit Sub
    
    LlamarImprimir2

End Sub

Private Sub cmdReparaEfect_Click()
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    Codigo = "schrep"
    Devuelve = ""
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCliente(0).Text <> "" Or txtCliente(1).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 0, 1, Devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion DEPARTAMENTO
    '--------------------------------------------
    If txtDpto(0).Text <> "" Or txtDpto(1).Text <> "" Then
        campo = "{" & Codigo & ".coddirec}"
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHDpto=""Dpto: "
        If Not PonerDesdeHasta(campo, "DPT", 0, 1, Devuelve) Then Exit Sub
    End If
    
    
    'Este trozo lo hace siempre
    If Me.optReparaciones(0).Value Then
        campo = "entre"
    Else
        Devuelve = "reparación"
        campo = "repar"
    End If
    campo = "{" & Codigo & ".fec" & campo & "}"
    Cadparam = Cadparam & "pOrden=" & campo & "|"
    NumParam = NumParam + 1
    
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        Devuelve = "pDHFecha=""Fecha " & Devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 0, 1, Devuelve) Then Exit Sub
    End If
    
    cadFrom = "schrep"
    If Not HayRegParaInforme(cadFrom, Cadselect) Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Prepararo los datos
    Codigo = "DELETE from tmpnlotes where codusu = " & vUsu.Codigo
    conn.Execute Codigo
    CargaImporteRealReparaciones
    
    
    'MOSTRAMOS EL INFORME
    'Añadir el nombre de la Empresa como parametro
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & "(isnull({tmpnlotes.codusu}) or {tmpnlotes.codusu}=1000)"
    
    conSubRPT = False
    LlamarImprimir2
    Screen.MousePointer = vbDefault
End Sub
Private Sub LlamarImprimir2(Optional Titulo As String)
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .Titulo = Titulo
        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = 2000 + opcion   '2000 mas la opcion de entrada
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub





Private Sub cmdResprodMoixent_Click()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    If HacerResumenProduccionMoixent Then
          '===================================================
            '============ PARAMETROS ===========================
            'Añadir el nombre de la Empresa como parametro
            Cadparam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            Devuelve = ""
            If txtFecha(40).Text <> "" Then Devuelve = Devuelve & " desde " & txtFecha(40).Text
            If txtFecha(41).Text <> "" Then Devuelve = Devuelve & " hasta " & txtFecha(41).Text
            
            If Devuelve <> "" Then Devuelve = " Fechas: " & Devuelve
            
            Devuelve = "pDHFecha=""" & Devuelve & """|"
            Cadparam = Cadparam & Devuelve
            
            NumParam = 2
            
        

            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            campo = "Resumen produccion"
        
        
            cadNomRPT = "rResumProOLD.rpt"
    
            
            conSubRPT = False
            LlamarImprimir2 campo
    
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdResumenProduccion_Click()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    If HacerListadoResumeProduccion Then
        
    
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTraza_Click()
    Screen.MousePointer = vbHourglass
    HacerInformeTrazabilidad
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVentasAceite_Click()
Dim R As ADODB.Recordset
Dim Cantidad As Currency
Dim vFamia As String

    
    miSQL = "DELETE FROM tmpstockfec where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    InicializarVbles
    
    Set miRsAux = New ADODB.Recordset
    Codigo = "select codfamia from sartic where conjunto=1 group by 1"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    While Not miRsAux.EOF
        Codigo = Codigo & ", " & miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'cadSelect = " sartic.codartic = slifac.codartic and sartic.conjunto = 1 "
    vFamia = ""
    cadFormula = "" ' " {sartic.conjunto} = 1 "
    
    
    Cadselect = "  scafac1.codtipom=slifac.codtipom and scafac1.numfactu=slifac.numfactu "
    Cadselect = Cadselect & " AND scafac1.fecfactu=slifac.fecfactu  and scafac1.codtipoa=slifac.codtipoa"
    Cadselect = Cadselect & " AND sartic.codartic = slifac.codartic "
    
    'Para el cliente
    Cadselect = Cadselect & " AND scafac1.codTipoM = scafac.codTipoM And scafac1.NumFactu = scafac.NumFactu"
    Cadselect = Cadselect & " AND scafac1.numalbar=slifac.numalbar"  'Agosto 2012.  NO estaba
    Cadselect = Cadselect & " AND scafac1.fecfactu=scafac.fecfactu"
    
    
    If Codigo <> "" Then
        Codigo = Mid(Codigo, 2)
        vFamia = " and sartic.codfamia in (" & Codigo & ")"
        'Para la formula
        '        ({sartic.codfamia} in [1, 3, 6]) ejem
        cadFormula = " ({sartic.codfamia} in [" & Codigo & "])"
    End If
    Codigo = ""
    If txtFecha(28).Text <> "" Or txtFecha(29).Text <> "" Then
        Codigo = Codigo & Space(20) & AnyadirParametroDH("Fechas: ", txtFecha(28), txtFecha(29), Nothing, Nothing)
        campo = "{scafac1.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 28, 29, Devuelve) Then Exit Sub
    End If


    'CLIENTE
    If txtCliente(7).Text <> "" Then
        'Ha puesto un cliente
        Codigo = Trim(Codigo & "   Cliente: " & txtCliente(7).Text & " " & Me.txtDescClie(7).Text)
        Cadselect = Cadselect & " AND scafac.codclien = " & txtCliente(7).Text
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & " ({scafac.codclien} = " & txtCliente(7).Text & ")"
        
        
    End If
        
    If vUsu.TrabajadorB Then
        If chkConsolidado(1).Value = 0 Then
            'SOLO QUIERE ver el almacen 2
            Codigo = Codigo & Space(20) & "Alma*"
            If Not AnyadirAFormula(Cadselect, "{slifac.codalmac}=" & vParamAplic.AlmacenB) Then Exit Sub
            cadFormula = cadFormula & " AND {slifac.codalmac}=" & vParamAplic.AlmacenB
        End If
    Else
        'SOLO el UNO
            If Not AnyadirAFormula(Cadselect, "{slifac.codalmac}<>" & vParamAplic.AlmacenB) Then Exit Sub
            cadFormula = cadFormula & " AND ({slifac.codalmac}<>" & vParamAplic.AlmacenB & ")"
    End If
    




    Codigo = Trim(Codigo)
    Cadparam = "pDH1= """ & Codigo & """|"
    NumParam = 1


    cadFrom = " scafac1,slifac,sartic,scafac"
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    If Not HayRegParaInforme(cadFrom, Cadselect & vFamia) Then Exit Sub

    
    
    'El informe se divide en dos.
    '1.- slifac donde los articulos son articulos de ventas(tienen componentes)
    '2.- resumen de los componentes que son materias primas
    '     tanto los que se hayan facturado directamente (kilos)
    '     como los que entra por que es componente de uno de venta
    
    'Entonces, el primer trozo del informe sale directamente de las columnas. El segundo
    'saldra de una tabla temporal  tmpstockfec
    ' tendra usuario, codartic codalmac  -> 0 Vendidos directamente   1- Como compenente    stcok  la cantidad en cuestion
    
    
    'miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) select " & vUsu.Codigo & ",slifac.codartic,0,cantidad * factorconversion from slifac,sartic"
    miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) "
    '                                                       * factorconversion
    miSQL = miSQL & " select " & vUsu.Codigo & ",slifac.codartic,0,sum(cantidad)  "
    miSQL = miSQL & " from scafac1,slifac,sartic,scafac"
    miSQL = miSQL & " where slifac.codartic=sartic.codartic and factorconversion<>1"
    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
    miSQL = miSQL & " group by slifac.codartic"
    conn.Execute miSQL    'Para el punto materias primas, los vendidos directamente
    
    
    'Cojeremos un cursor con todos las materias primas e iremos insertandolas en la tmpstock
    '-----------------
    miSQL = "select slifac.codartic,sum(cantidad) cantidad from scafac1,slifac,sartic,scafac"
    miSQL = miSQL & " where slifac.codartic=sartic.codartic "
    miSQL = miSQL & " and conjunto=1 "
    'Las fechas
    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
    miSQL = miSQL & " group by slifac.codartic "
    
    Set R = New ADODB.Recordset
    Devuelve = "|"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
            'Para cada elemento facturado que tiene componentes, vere de sus componentes cual es el de mataria prima y calcular su cantidad
            miSQL = "select sarti1.codarti1,cantidad from sarti1,sartic where  sarti1.codarti1=sartic.codartic and factorconversion<>1"
            miSQL = miSQL & " AND sarti1.codartic =" & DBSet(miRsAux!codartic, "T")
            R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            'If Mid(miRsAux!codArtic, 1, 9) = "003700441513" Then Stop
            
            
            
            Cantidad = DBLet(miRsAux!Cantidad, "N")

            
            'Para no tener que hacer un select para saber si ya ha sido insertado en tmpstock, utilizar
            'el string cadSelect para ir metiendo los ya insertados.
            While Not R.EOF
                'El articulo en cuestion
                miSQL = "|" & R!codarti1 & "|"
                Cantidad = Cantidad * R!Cantidad   'Esta es la cantidad nueva
                campo = TransformaComasPuntos(CStr(Cantidad))
                If InStr(1, Devuelve, miSQL) > 0 Then
                    'Ya esta insertado. Es un UPDATE
                    miSQL = "UPDATE tmpstockfec SET stock=stock + " & campo
                    miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " and codartic = " & DBSet(R!codarti1, "T")
                    miSQL = miSQL & " AND codalmac= 1"
                Else
                    Devuelve = Devuelve & R!codarti1 & "|"
                    miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock)  VALUES (" & vUsu.Codigo & "," & DBSet(R!codarti1, "T")
                    miSQL = miSQL & ",1," & campo & ")"
                    
                End If
                conn.Execute miSQL
                'No deberia haber mas (seria un coupage)
                R.MoveNext
            Wend
            R.Close
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set R = Nothing


    


    Cadparam = Cadparam & "pCodUsu= " & vUsu.Codigo & "|"
    'cadFormula = "{tmpnlotes.codusu}=" & vUsu.Codigo
    
    
    'Añadimos nombre empresa
    Cadparam = "|pNomEmpre= """ & vParam.NombreEmpresa & """|" & Cadparam
    NumParam = NumParam + 1
    
    With frmImprimir
        .ConSubInforme = True
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .NombreRPT = "rventasmor2.rpt"  'Aunque pone
        .Titulo = "Listado"
        .opcion = 2016
        .Show vbModal
    End With

End Sub

Private Sub cmdVentasAgentes_Click()
    If txtFecha(30).Text = "" Or txtFecha(31).Text = "" Then
        MsgBox "Escriba fechas", vbExclamation
        Exit Sub
    End If

    InicializarVbles
    
    'Los agentes
    Devuelve = ""
    campo = ""
    For NumRegElim = 0 To List2.ListCount - 1
        If List2.Selected(NumRegElim) Then
            campo = campo & ", " & List2.ItemData(NumRegElim)
        Else
            Devuelve = "N" 'Hay alguno NO seleccionado
        End If
    Next
    If campo = "" Then
        MsgBox "Seleccione algun agente", vbExclamation
        Exit Sub
    End If
    
    If Devuelve <> "" Then
        'Hay que marcar algun agente
        campo = Mid(campo, 2)
        Cadselect = "( {scafac.codagent} IN (" & campo & "))"
        Codigo = "Agentes: " & campo
    Else
        Codigo = ""
    End If
    If txtFecha(30).Text <> "" Or txtFecha(31).Text <> "" Then
        Codigo = Codigo & Space(20) & AnyadirParametroDH("Fechas: ", txtFecha(30), txtFecha(31), Nothing, Nothing)
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 30, 31, Devuelve) Then Exit Sub
    End If
    
    
    If vUsu.TrabajadorB Then
        If chkConsolidado(0).Value = 0 Then
            'SOLO QUIERE ver el almacen 2
            Codigo = Codigo & Space(20) & "Alma*"
            If Not AnyadirAFormula(Cadselect, "{slifac.codalmac}=" & vParamAplic.AlmacenB) Then Exit Sub
        End If
    End If
    
    
    Codigo = Trim(Codigo)
    Cadparam = "|pDH1= """ & Codigo & """|"
    NumParam = 1
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    
    
    
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    cadFormula = "{olitmpventasagente.codusu} = " & vUsu.Codigo
    If ObtenerDatosVentasAgentes Then
        cadNomRPT = "rVentasAgentes.rpt"
        LlamarImprimir2
    
    End If
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing
    
End Sub

Private Sub cmdVentaxProv_Click()
Dim cad As String
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCliente(3).Text <> "" Or txtCliente(4).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 3, 4, cad) Then Exit Sub
    End If
   
    
    
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(9).Text <> "" Or txtFecha(10).Text <> "" Then
        campo = "{scafac.fecfactu}"
        cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 9, 10, cad) Then Exit Sub
    End If
    
    
    
    'Cadena para seleccion Desde y Hasta ARTICULO
    '--------------------------------------------
    If txtArticulo(1).Text <> "" Or txtArticulo(2).Text <> "" Then
        campo = "{slifac.codartic}"
        cad = "pDHDpto=""Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 1, 2, cad) Then Exit Sub
    End If
    
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '---------------------------------------------
    If txtCodProve(0).Text <> "" Or txtCodProve(1).Text <> "" Then
        campo = "{slifac.codprovex}"
        cad = "pDHPro=""Proveedor: "
        If Not PonerDesdeHasta(campo, "PRO", 0, 1, cad) Then Exit Sub
    End If

     
    
    'Pongo en campo el select
    Codigo = " scafac.codtipom=slifac.codtipom "
    Codigo = " scafac.fecfactu = slifac.fecfactu AND scafac.numfactu=slifac.numfactu AND " & Codigo
    cad = "scafac,slifac"
    If Cadselect <> "" Then Codigo = Codigo & " AND " & Cadselect
    campo = Codigo
    If Not HayRegParaInforme(cad, Codigo) Then Exit Sub
    
    
    cadNomRPT = "rvtaxcodprove.rpt"
    LlamarImprimir2
    

End Sub

Private Sub Command1_Click()
        
End Sub



Private Sub cmdVtasAgrupadas_Click()
  
       
    Screen.MousePointer = vbHourglass
    
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    Cadparam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    cadFormula = "{sartic.codmarca} > 0"
    
    
    
   
    
    
    
    'Cadena para seleccion D/H marca
    '--------------------------------------------
    campo = ""
    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
   
        'Parametro Desde/Hasta Cliente
        Devuelve = CadenaDesdeHasta(txtCodigo(0).Text, txtCodigo(1).Text, "{sartic.codfamia}", "N")
        If Devuelve <> "error" Then
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        End If
        campo = campo & "CATEGORIA: "
        If txtCodigo(0).Text <> "" Then campo = campo & " desde " & Trim(txtCodigo(0).Text & " " & Text1(0).Text)
        If txtCodigo(1).Text <> "" Then campo = campo & " hasta " & Trim(txtCodigo(1).Text & " " & Text1(1).Text)
    End If
   
    If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
   
        'Parametro Desde/Hasta Cliente
        Devuelve = CadenaDesdeHasta(txtCodigo(2).Text, txtCodigo(3).Text, "{sartic.codmarca}", "N")
        If Devuelve <> "error" Then
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        End If
        campo = Trim(campo & "         MARCA: ")
        If txtCodigo(2).Text <> "" Then campo = campo & " desde " & Trim(txtCodigo(2).Text & " " & Text1(2).Text)
        If txtCodigo(3).Text <> "" Then campo = campo & " hasta " & Trim(txtCodigo(3).Text & " " & Text1(3).Text)
    End If
    Cadparam = Cadparam & "dh1= """ & campo & """|"
    NumParam = NumParam + 1
    
    
    
    
        'Cadena para seleccion D/H marca
    '--------------------------------------------
    campo = ""
    If txtCodigo(4).Text <> "" Or txtCodigo(5).Text <> "" Then
   
        'Parametro Desde/Hasta Cliente
        Devuelve = CadenaDesdeHasta(txtCodigo(4).Text, txtCodigo(5).Text, "{sartic.codunida}", "N")
        If Devuelve <> "error" Then
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        End If
        campo = campo & "FORMATO: "
        If txtCodigo(4).Text <> "" Then campo = campo & " desde " & Trim(txtCodigo(4).Text & " " & Text1(4).Text)
        If txtCodigo(5).Text <> "" Then campo = campo & " hasta " & Trim(txtCodigo(5).Text & " " & Text1(5).Text)
    End If
   
    If txtCodigo(6).Text <> "" Or txtCodigo(7).Text <> "" Then
   
        'Parametro Desde/Hasta Cliente
        Devuelve = CadenaDesdeHasta(txtCodigo(6).Text, txtCodigo(7).Text, "{sartic.codtipar}", "T")
        If Devuelve <> "error" Then
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        End If
        campo = Trim(campo & "         MODELO: ")
        If txtCodigo(6).Text <> "" Then campo = campo & " desde " & Trim(txtCodigo(6).Text & " " & Text1(6).Text)
        If txtCodigo(7).Text <> "" Then campo = campo & " hasta " & Trim(txtCodigo(7).Text & " " & Text1(7).Text)
    End If
    Cadparam = Cadparam & "dh2= """ & campo & """|"
    NumParam = NumParam + 1
    
    
    
    
    
    
    
    Cadselect = cadFormula
    Cadselect = Replace(Cadselect, "{", "")
    Cadselect = Replace(Cadselect, "}", "")
     
     '====================================================
    '================= FORMULA ==========================
    campo = "       "
    
    If txtFecha(36).Text <> "" Or txtFecha(37).Text <> "" Then
        
        'Parametro Desde/Hasta Cliente
        Devuelve = CadenaDesdeHasta(txtFecha(36).Text, txtFecha(37).Text, "{slifac.fecfactu}", "F")
        If Devuelve <> "error" Then
            If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        End If
        
        
        campo = "Fechas: " & txtFecha(36).Text & " -- " & txtFecha(37).Text & campo
        
        If txtFecha(36).Text <> "" Then
            If Cadselect <> "" Then Cadselect = Cadselect & " AND "
            Cadselect = Cadselect & "slifac.fecfactu >='" & Format(txtFecha(36).Text, FormatoFecha) & "'"
        End If
        
        If txtFecha(37).Text <> "" Then
            If Cadselect <> "" Then Cadselect = Cadselect & " AND "
            Cadselect = Cadselect & "slifac.fecfactu <='" & Format(txtFecha(37).Text, FormatoFecha) & "'"
        End If
    End If
    Cadparam = Cadparam & "desdehastaFe= """ & campo & """|"
    NumParam = NumParam + 1
     
     
    
    'Pongo en campo el select
    If Cadselect <> "" Then Cadselect = " AND " & Cadselect
    Cadselect = " slifac.codartic=sartic.codartic " & Cadselect
    If HayRegParaInforme("sartic,slifac", Cadselect) Then
                
        If Me.optVtaGrup(0).Value Then
            cadNomRPT = "rMarModfamTip1.rpt"
        Else
             cadNomRPT = "rMarModfamTip2.rpt"
        End If
        LlamarImprimir2 "Ventas agrupadas ..."
    End If
    Screen.MousePointer = vbDefault
End Sub





Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case opcion
        Case 0
            PonerFoco txtCliente(0)
        Case 1
            PonerFoco txtTrab(0)
        Case 6 To 10
            'En ambos listados lo primero es una fecha
            If opcion = 6 Then
                NumParam = 9
            ElseIf opcion = 7 Then
                NumParam = 11
            ElseIf opcion = 8 Then
                NumParam = 13  'liquidacion factura sprov
                txtFecha(17).Text = Format(Now, "dd/mm/yyyy")
            Else
                NumParam = 8 + opcion 'impresion facturas  index:17 y 18
            End If
            PonerFoco txtFecha(CInt(NumParam))
        
        Case 13
            Cadparam = ""
            'Poner el nombre del trabajador que esta conectado
            Me.txtTrab(2).Text = PonerTrabajadorConectado(Cadparam)
            Me.txtDescTra(2).Text = Cadparam
        
        Case 16
            PonerFoco Me.txtArticulo(9)
            
        Case 17
            PonerFoco txtCliente(5)
            
        Case 19
            'Ventas morales
            PonerFoco txtFecha(28)
            
        Case 20
            PonerFoco txtFecha(30)
            
            
        Case 22
            PonerFoco txtFecha(33)
        
        Case 24
            PonerFoco txtFecha(34)
            
        Case 26
            PonerFoco txtFecha(36)
        Case 27
            PonerFoco txtFecha(38)
        Case 28
            PonerFoco txtCodProve(12)
        Case 29
            PonerFoco txtFecha(40)
        Case 35
           'PonerFocoChk Me.optPalets(0)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim IndiceCancel As Integer

    Me.Icon = frmppal.Icon
    PrimeraVez = True
    limpiar Me
    FrListadoReparaciones.visible = False
    FrEstadisticasReparacionTecnico.visible = False
    FrameMultibase.visible = False
    FrameRecargaMov.visible = False
    Me.FrFacturaRecargas.visible = False
    FrProveedorxVenta.visible = False
    FrLiqCambioPrecios.visible = False
    Me.FrGeneraFactLiq.visible = False
    Me.FrImprimirFac.visible = False
    FrameAlbaProv.visible = False
    frameContabTickets.visible = False
    Me.FrameTraza.visible = False
    FrameTO.visible = False
    Me.FrameListadoTO.visible = False
    FrameVentasMorales.visible = False
    FrameLiqAgentes.visible = False
    FrameEnvioTarifa.visible = False
    FrameInvenACeite.visible = False
    FrameEcoEnves.visible = False
    FrameVtasVarias.visible = False
    FrameResuProduccion.visible = False
    FrameCambioProveedor.visible = False
    FrameResumenProduccionMoixent.visible = False
    FrameRecalPrStandard.visible = False
    FrameAmpliaProd.visible = False
    FrameTransporte4Tonda.visible = False
    FrameEntradaOliva.visible = False
    FrameOliva.visible = False
    FramePalets.visible = False
    FrameAlbaranesVall.visible = False
    FrameDeclaraAlmazara.visible = False
    Caption = "Listado"
    IndiceCancel = opcion
    Select Case opcion
    Case 1
        'Listado reparaciones efectuadas
        PonerFrameVisible FrListadoReparaciones, H, W
        PonerLabelDptoDireccion Me.lblDpto(0)
        
        
        
    Case 2
        PonerFrameVisible Me.FrEstadisticasReparacionTecnico, H, W
        
        
    Case 3
        Caption = "MULTIBASE"
        PonerFrameVisible Me.FrameMultibase, H, W
        CargaListMultibase
        
    Case 4
        'Informe recarga movil
        PonerFrameVisible FrameRecargaMov, H, W
        Me.cmbRecargaMov(0).ListIndex = 0
        Me.cmbRecargaMov(1).ListIndex = 0
        
    Case 5
        'Facturacion recargas moviles
        Caption = "Facturación"
        PonerFrameVisible FrFacturaRecargas, H, W
        txtFecha(8).Text = Format(Now, "dd/mm/yyyy")
        lblIndicadorT.Caption = ""
        pb1.visible = False
        'Lo del articulo lo pongo visib
        txtArticulo(0).Text = vParamAplic.CodarticTfnia
        txtArticulo_LostFocus 0
        txtArticulo(0).visible = False
        Me.txtDescArticulo(0).visible = False
        Me.imgArticulo(0).visible = False
        Label4(2).visible = False
        
    Case 6
        'Ventas por codprove
        'TRAZA enero 2008
        PonerFrameVisible FrProveedorxVenta, H, W
        
    Case 7
        lblLiqu.Caption = ""
        PonerFrameVisible FrLiqCambioPrecios, H, W
    Case 8
        Label1.Caption = ""
        PonerFrameVisible FrGeneraFactLiq, H, W
    Case 9
        Label2.Caption = ""
        PonerFrameVisible FrImprimirFac, H, W
    Case 10
        PonerFrameVisible FrameAlbaProv, H, W
        
        
        
    Case 13, 14
        Caption = "Tickets agrupados"
        If opcion = 13 Then
            lblTitulo(10).Caption = "Facturar " & lblTitulo(10).Caption
            cmdContabTicket.Caption = "Contabilizar"
        Else
            lblTitulo(10).Caption = "Listados " & lblTitulo(10).Caption
            cmdContabTicket.Caption = "Aceptar"
        End If
        Me.FrameTapa.visible = opcion = 13
        PonerFrameVisible frameContabTickets, H, W
        IndiceCancel = 13
    Case 15
         
        PonerFrameVisible FrameTraza, H, W
    Case 16
        Caption = "Generación"
        PonerFrameVisible FrameTO, H, W
        
    Case 17
        PonerFrameVisible Me.FrameListadoTO, H, W
        'Si es de tarifas mostraresmo el desde hasta, si no , obviamente, no
        optListadoTO(3).visible = False
        optListadoTO(2).visible = NumRegElim = 0
        FrameTosTapa.visible = NumRegElim = 0
        lblTitulo(15).Caption = "Listado tarifa"
        If NumRegElim = 0 Then lblTitulo(15).Caption = lblTitulo(15).Caption & "-oferta"
        
        optOrdenTO(CInt(CadenaDesdeOtroForm)).Value = True
        
        
        
    Case 19
        'Ventas morales
        chkConsolidado(1).visible = vUsu.TrabajadorB
        PonerFrameVisible FrameVentasMorales, H, W
    Case 20
        chkConsolidado(0).visible = vUsu.TrabajadorB
    
        CargaAgentes
        PonerFrameVisible FrameLiqAgentes, H, W
    Case 21
        CargaClientes
        PonerFrameVisible FrameEnvioTarifa, H, W
        
    Case 22
        Caption = "Inventario"
        chkConsolidado(2).visible = vUsu.TrabajadorB
        PonerFrameVisible FrameInvenACeite, H, W
        
        
        
    Case 24
        Label3(72).Caption = ""
        PonerFrameVisible FrameEcoEnves, H, W
        
    Case 26
        PonerFrameVisible FrameVtasVarias, H, W
   
    Case 27
        PonerFrameVisible FrameResuProduccion, H, W
        txtFecha(38).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        
    Case 28
        'Cambio proveedor
        PonerFrameVisible FrameCambioProveedor, H, W
        
    Case 29
        cmdResprodMoixent.Caption = "&Aceptar"
        PonerFrameVisible FrameResumenProduccionMoixent, H, W
        
        
    Case 30
        'reclaculo precio stadnadar
        PonerFrameVisible FrameRecalPrStandard, H, W
        CargaImagenAyuda 0, Me.lblTitulo(24).Caption
    
    
    
    Case 31
        PonerFrameVisible FrameAmpliaProd, H, W
        Text2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Text2(1).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        CadenaDesdeOtroForm = ""
        
    Case 32
        PonerFrameVisible FrameTransporte4Tonda, H, W
    Case 33
        PonerFrameVisible FrameEntradaOliva, H, W
        Codigo = CadenaDesdeOtroForm  'Lo guardo en codigo pq despues se vacia
        
    Case 34
        PonerFrameVisible FrameOliva, H, W
        Caption = "Generar"
        'Entrada de camion para generar albaranes
        cadFormula = CStr(CadenaDesdeOtroForm)  'Codigo
        
        CadenaDesdeOtroForm = ""
    Case 35
        Caption = "PALETS"
        PonerFrameVisible FramePalets, H, W
        Frame2(0).BorderStyle = 0
        Frame2(1).BorderStyle = 0
        Frame2(2).BorderStyle = 0
        Frame2(3).BorderStyle = 0
        
        If CadenaDesdeOtroForm <> "" Then
            Me.txtArticulo(12).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.txtDescArticulo(12).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            CadenaDesdeOtroForm = ""
        End If
    Case 36
        PonerFrameVisible FrameAlbaranesVall, H, W
        
    Case 37
        PonerFrameVisible FrameDeclaraAlmazara, H, W
        Label3(98).Caption = ""
        CargaComboMes 0
        'Ponemos un mes mas del ultimo
   '     If vParamAplic.FechaActiva >= CDate("01/01/2016") Then
            campo = CStr(DateAdd("m", 1, vParamAplic.FechaActiva))
            NumParam = Month(CDate(campo))
            campo = Year(CDate(campo))
            cboMes(0).ListIndex = NumParam - 1
            Me.txtNumeroEntero(4).Text = campo
    '    End If
    End Select
    Me.Height = H + 150
    Me.Width = W
    Me.cmdCancel(IndiceCancel).Cancel = True
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    txtDpto(IndiceImg).Text = RecuperaValor(CadenaDevuelta, 1)
    txtDescDpto(IndiceImg) = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub frmBaPr_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtCliente(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescClie(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtForpa(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescForpa(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPr_DatoSeleccionado(CadenaSeleccion As String)
    txtCodProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 1)
    txtDescProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    txtTrab(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescTra(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgArticulo_Click(Index As Integer)
    IndiceImg = Index
    Set frmMtoArticulos = New frmAlmArticulos
    frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
End Sub

Private Sub ImgAyuda_Click(Index As Integer)
Dim C As String

    Select Case Index
    Case 0
    
        C = "Recalculará el precio de : " & vbCrLf
        C = C & "- Materia auxiliar:   Pondrá lo que tenga el coste(último precio compra)" & vbCrLf & vbCrLf
        C = C & "- Producto venta:   Recalcular el precio standard a partir del calculo de precio de venta. " & vbCrLf
        C = C & "    Una vez obtenido el precio , lo pondrá en la columna Pr. standard. " & vbCrLf
        C = C & "    La materia prima partirá siempre del  standard y la materia auxiliar  " & vbCrLf
        C = C & "    del precio st   o bien del precio coste(segun Check)  " & vbCrLf
    
    End Select
    C = ImgAyuda(Index).ToolTipText & vbCrLf & vbCrLf & C
    MsgBox C, vbInformation
End Sub

Private Sub imgBancoPr_Click(Index As Integer)
    IndiceImg = Index
    Set frmBaPr = New frmFacBancosPropios
    frmBaPr.DatosADevolverBusqueda = "1" 'Abrimos en Modo Busqueda
    frmBaPr.Show vbModal
    Set frmBaPr = Nothing
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For NumRegElim = 0 To List3.ListCount - 1
        List3.Selected(NumRegElim) = Index = 1
    Next NumRegElim
End Sub

Private Sub imgCliente_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    IndiceImg = Index
    Set frmCli = New frmFacClientes
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub imgCodigo_Click(Index As Integer)

    Set frmB2 = New frmBuscaGrid
    campo = "N"
    Select Case Index
        Case 0, 1
            Cadparam = "codfamia"
            Devuelve = "sfamia"
            miSQL = "nomfamia"
            IndiceImg = 36
        Case 2, 3
            Cadparam = "codmarca"
            Devuelve = "smarca"
            miSQL = "nommarca"
            IndiceImg = 37
        Case 4, 5
            Cadparam = "codunida"
            Devuelve = "sunida"
            miSQL = "nomunida"
            IndiceImg = 38
        Case 6, 7
            Cadparam = "codtipar"
            miSQL = "nomtipar"
            Devuelve = "stipar"
            campo = "T"
            IndiceImg = 39

        End Select
        'Cod.|sdirec|coddirec|N||20·"
        campo = "Codigo|" & Devuelve & "|" & Cadparam & "|" & campo & "||20·"
        campo = campo & "Descripcion|" & Devuelve & "|" & miSQL & "|T||60·"
    
        miSQL = ""
        frmB2.vTitulo = Me.lblDpto(IndiceImg).Caption
        frmB2.vCampos = campo
        frmB2.vCargaFrame = False
        frmB2.vDevuelve = "0|1|"
        frmB2.vselElem = 1
        frmB2.vConexionGrid = 1  'ODBC Ariges
        frmB2.vTabla = Devuelve
        frmB2.vSQL = ""
        frmB2.Show vbModal
        Set frmB = Nothing
        If miSQL <> "" Then
            txtCodigo(Index).Text = RecuperaValor(miSQL, 1)
            Text1(Index).Text = RecuperaValor(miSQL, 2)
            If Index = 7 Then
                PonFocoChk
            Else
                PonerFoco txtCodigo(Index + 1)
            End If
            miSQL = ""
        End If
        Codigo = ""
        campo = ""
        Cadparam = ""
End Sub

Private Sub PonFocoChk()
    On Error Resume Next
    Me.optVtaGrup(0).SetFocus
    Err.Clear
End Sub

Private Sub imgDpto_Click(Index As Integer)
    If Index < 2 Then
        'DPTO
        IndiceImg = Index
        If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
            'OK
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = Me.lblDpto(0).Caption & " " & txtCliente(0).Text & " - " & txtDescClie(0).Text
            campo = "Cod.|sdirec|coddirec|N||20·"
            campo = campo & "Desc.|sdirec|nomdirec|T||40·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Ariges
            frmB.vTabla = "sdirec"
            frmB.vSQL = "codclien = " & txtCliente(0).Text
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            
        Else
            MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
        End If
    End If
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    Set frmFor = New frmAlmFamiliaArticulo
    miSQL = ""
    frmFor.DatosADevolverBusqueda = "0|1|"
    frmFor.Show vbModal
    Set frmFor = Nothing
    If miSQL <> "" Then
        'Ha devuelvto datos
        Me.txtFamilia(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescFamilia(Index).Text = RecuperaValor(miSQL, 2)
        If Index = 0 Then
            PonerFoco txtFamilia(1)
        Else
          '  PonerFoco txtMarca(0)
        End If
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub


Private Sub imgForPa_Click(Index As Integer)
    IndiceImg = Index
    Set frmFP = New frmFacFormasPago
    frmFP.DatosADevolverBusqueda = "0|1|"
    frmFP.Show vbModal
    Set frmFP = Nothing
End Sub

Private Sub imgMarca_Click(Index As Integer)

        'Para el resto de casos
'        miSQL = ""
'        Set frmMar = New frmAlmMarcas
'        frmMar.DatosADevolverBusqueda = "0|1|"
'        frmMar.Show vbModal
'        Set frmMar = Nothing
'        If miSQL <> "" Then
'            'Ha devuelvto datos
'
'            Me.txtMarca(Index).Text = RecuperaValor(miSQL, 1)
'            Me.txtDescMarca(Index).Text = RecuperaValor(miSQL, 2)
'            If Index = 0 Then
'                PonerFoco txtMarca(1)
'            Else
'                PonerFocoBtn cmdGeneraTO
'            End If
'        End If

End Sub

Private Sub imgProveedor_Click(Index As Integer)
    IndiceImg = Index
    Set frmPr = New frmComProveedores
    frmPr.DatosADevolverBusqueda = "0|1|"
    frmPr.Show vbModal
    Set frmPr = Nothing
End Sub

Private Sub imgTecnico_Click(Index As Integer)
    IndiceImg = Index
    Set frmT = New frmAdmTrabajadores
    frmT.DatosADevolverBusqueda = "0|1|"
    frmT.Show vbModal
    Set frmT = Nothing
End Sub

Private Sub optListadoTO_Click(Index As Integer)
    Me.FrOrdenTO.visible = Index <> 1
End Sub

Private Sub optListadoTO_KeyPress(Index As Integer, KeyAscii As Integer)
    
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub optPalets_Click(Index As Integer)
    Me.Frame2(1).visible = Index = 0
    Me.Frame2(2).visible = Not Frame2(1).visible
    
End Sub

Private Sub optPalets_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optRecalPrSt_Click(Index As Integer)
    Me.chkRecalPrSt.visible = Me.optRecalPrSt(1).Value
End Sub

Private Sub optRecalPrSt_KeyPress(Index As Integer, KeyAscii As Integer)
 KEYpressGnral KeyAscii, 2, True
 
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    ConseguirFoco Text2(Index), 3
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Replace(Text2(Index).Text, "|", "-")
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    Codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If Codigo = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = Codigo
    Codigo = ""
    
End Sub



Private Sub txtBancoPr_GotFocus(Index As Integer)
    ConseguirFoco txtBancoPr(Index), 3
End Sub

Private Sub txtBancoPr_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtBancoPr_LostFocus(Index As Integer)
    txtBancoPr(Index).Text = Trim(txtBancoPr(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtBancoPr(Index).Text <> "" Then
        If IsNumeric(txtBancoPr(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", txtBancoPr(Index).Text, "N")
            If Codigo = "" Then miSQL = "El codigo no pertence a ningun banco propio"
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescBancoPr(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtBancoPr(Index).Text = ""
        PonerFoco txtBancoPr(Index)
    End If
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
            PonerFoco txtCliente(Index)
        Else
            Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Descri = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
                If Index = 8 Then txtCliente(Index).Text = ""
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Descri
    
    
    
End Sub


    

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim OK As Boolean
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Screen.MousePointer = vbHourglass
    Codigo = ""
    If txtCodigo(Index).Text <> "" Then
        If Index < 6 Then
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "Campo debe ser numérico", vbExclamation
                OK = False
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            Else
                OK = True
            End If
        Else
            OK = True
        End If
        If OK Then
            campo = "N"
            Select Case Index
            Case 0, 1
                Cadparam = "codfamia"
                Devuelve = "sfamia"
                miSQL = "nomfamia"
            Case 2, 3
                Cadparam = "codmarca"
                Devuelve = "smarca"
                miSQL = "nommarca"
                
            Case 4, 5
                Cadparam = "codunida"
                Devuelve = "sunida"
                miSQL = "nomunida"
            
            Case 6, 7
                Cadparam = "codtipar"
                miSQL = "nomtipar"
                Devuelve = "stipar"
                campo = "T"
    
    
            End Select
            Codigo = DevuelveDesdeBD(conAri, miSQL, Devuelve, Cadparam, txtCodigo(Index).Text, campo)
            If Codigo = "" Then MsgBox "No existe " & Cadparam & " (" & Devuelve & ")", vbInformation
                
            Cadparam = ""
            miSQL = ""
            Devuelve = ""
            campo = ""
        End If

    End If
    
    Text1(Index).Text = Codigo
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtCodProve_GotFocus(Index As Integer)
    ConseguirFoco txtCodProve(Index), 3
End Sub

Private Sub txtCodProve_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodProve_LostFocus(Index As Integer)
    txtCodProve(Index).Text = Trim(txtCodProve(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtCodProve(Index).Text <> "" Then
        If IsNumeric(txtCodProve(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtCodProve(Index).Text, "N")
            If Codigo = "" Then MsgBox "El codigo no pertence a ningun proveedor", vbExclamation
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescProve(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtCodProve(Index).Text = ""
        PonerFoco txtCodProve(Index)
    End If
End Sub




Private Sub txtFamilia_GotFocus(Index As Integer)
    ConseguirFoco txtFamilia(Index), 3
End Sub

Private Sub txtFamilia_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFamilia_LostFocus(Index As Integer)
    txtFamilia(Index).Text = Trim(txtFamilia(Index).Text)
    miSQL = ""
    If txtFamilia(Index).Text <> "" Then
        If IsNumeric(txtFamilia(Index).Text) Then
            miSQL = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamilia(Index).Text, "N")
            If miSQL = "" Then miSQL = "No pertence a ninguna familia"
        Else
            txtFamilia(Index).Text = ""
        End If
    End If
    Me.txtDescFamilia(Index).Text = miSQL
    
End Sub
Private Sub txtForpa_GotFocus(Index As Integer)
    ConseguirFoco txtForpa(Index), 3
End Sub

Private Sub txtForpa_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtForpa_LostFocus(Index As Integer)
    txtForpa(Index).Text = Trim(txtForpa(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtForpa(Index).Text <> "" Then
        If IsNumeric(txtForpa(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtForpa(Index).Text, "N")
            If Codigo = "" Then miSQL = "El codigo no pertence a ningun forma de pago"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescForpa(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtForpa(Index).Text = ""
        PonerFoco txtForpa(Index)
    End If
End Sub

Private Sub txtHora_GotFocus(Index As Integer)
    ConseguirFoco txtHora(Index), 3
End Sub

Private Sub txtHora_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtHora_LostFocus(Index As Integer)
    txtHora(Index).Text = Trim(txtHora(Index).Text)
    If txtHora(Index).Text = "" Then Exit Sub
    miSQL = txtHora(Index).Text
    If Not EsHoraOK(miSQL) Then
        txtHora(Index).Text = ""
    Else
        txtHora(Index).Text = miSQL
    End If
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    PonerFormatoDecimal txtImporte(Index), 2  '10,4  en formato decimal
End Sub




Private Sub txtMarca_GotFocus(Index As Integer)
'    ConseguirFoco txtMarca(Index), 3
End Sub

Private Sub txtMarca_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtMarca_LostFocus(Index As Integer)
'    txtMarca(Index).Text = Trim(txtMarca(Index).Text)
'    miSQL = ""
'    If txtMarca(Index).Text <> "" Then
'        If IsNumeric(txtMarca(Index).Text) Then
'            miSQL = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", txtMarca(Index).Text, "N")
'            If miSQL = "" Then
'                miSQL = "El código no pertence a ninguna marca"
'                If Index = 0 Then
'                    MsgBox miSQL, vbExclamation
'                    miSQL = ""
'                    PonerFoco txtMarca(Index)
'                End If
'            End If
'        Else
'            txtMarca(Index).Text = ""
'        End If
'    End If
'    Me.txtDescMarca(Index).Text = miSQL
'
End Sub

Private Sub txtNumAlbar_GotFocus(Index As Integer)
    ConseguirFoco txtNumAlbar(Index), 3
End Sub

Private Sub txtNumAlbar_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub txtNumeroEntero_GotFocus(Index As Integer)
    ConseguirFoco txtNumeroEntero(Index), 3
End Sub

Private Sub txtNumeroEntero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumeroEntero_LostFocus(Index As Integer)
    txtNumeroEntero(Index).Text = Trim(txtNumeroEntero(Index).Text)
    If txtNumeroEntero(Index).Text = "" Then Exit Sub
    
    If Not PonerFormatoEntero(txtNumeroEntero(Index)) Then
        txtNumeroEntero(Index).Text = ""
        PonerFoco txtNumeroEntero(Index)
    End If
End Sub

Private Sub txtRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)
    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtTrab(Index).Text <> "" Then
        If IsNumeric(txtTrab(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTrab(Index).Text, "N")
            If Codigo = "" Then miSQL = "El código no pertence a ningun trabajador"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescTra(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtTrab(Index).Text = ""
        PonerFoco txtTrab(Index)
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    ConseguirFoco txtDpto(Index), 3
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
Dim vC As CCliente
    'Si el cliente ES EL MISMO
    campo = ""
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If Index < 2 Then
        If txtDpto(Index).Text <> "" Then
             'Index=0 or 1.  Departamento sera puesto si, y solo si, el cliente es el mismo
             If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
                 'PERFECTO, el cliente existe y es el mismo
                 Set vC = New CCliente
                 vC.Codigo = txtCliente(0).Text
                 vC.DptoCliente txtDpto(Index).Text, campo
                 Set vC = Nothing
             Else
                 'Todavia no ha puesto el cliente
                 MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
                 txtDpto(Index).Text = ""
        
             End If
        End If
        Me.txtDescDpto(Index).Text = campo
    Else
    
    
    End If
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
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub

'Dado un FRAME lo pone a true y lo situa en x:120 y:0 y devuelve lo que debe medir el form
Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub


Private Sub PonerLabelDptoDireccion(L As Label)
    If vParamAplic.Departamento Then
        L.Caption = "Dpto."
    Else
        L.Caption = "Direc."
    End If
End Sub



Private Sub CargaComboMes(Indice As Integer)
    cboMes(Indice).Clear
    For NumParam = 1 To 12
        cboMes(Indice).AddItem MonthName(CLng(NumParam), False)
    Next
End Sub



Private Sub InicializarVbles()
    cadFormula = ""
    Cadselect = ""
    Cadparam = ""
    NumParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim Devuelve As String
Dim cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA
Dim TDes As TextBox
Dim THas As TextBox
Dim DesD As TextBox 'Descripcion DESDE
Dim DesH As TextBox '    "       HASTA

    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F"
        'Campos fecha
        Set TDes = txtFecha(indD)
        Set THas = txtFecha(indH)
        Subtipo = "F"
    Case "CLI"
        'Cliente
        Set TDes = txtCliente(indD)
        Set THas = txtCliente(indH)
        Set DesD = txtDescClie(indD)
        Set DesH = txtDescClie(indH)
        Subtipo = "N"
    Case "DPT"
        'DEpartamento
        Set TDes = txtDpto(indD)
        Set THas = txtDpto(indH)
        Set DesD = txtDescDpto(indD)
        Set DesH = txtDescDpto(indH)
        Subtipo = "N"
        
    Case "PRO"
        Set TDes = txtCodProve(indD)
        Set THas = txtCodProve(indH)
        Set DesD = txtDescProve(indD)
        Set DesH = txtDescProve(indH)
        Subtipo = "N"
 
    Case "ART"

        Set TDes = txtArticulo(indD)
        Set THas = txtArticulo(indH)
        Set DesD = txtDescArticulo(indD)
        Set DesH = txtDescArticulo(indH)
        Subtipo = "T"
 
 
    Case "ALP"
        'Numero albaran proveedores
         
        Set TDes = txtNumAlbar(indD)
        Set THas = txtNumAlbar(indH)
        Subtipo = "T"
        
    Case "N_E"
        'Numero entero
        Set TDes = Me.txtNumeroEntero(indD)
        Set THas = Me.txtNumeroEntero(indH)
        Subtipo = "N"
        
    End Select
    
    Devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If Devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Function
    
    If Subtipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(TDes.Text, THas.Text, campo, Tipo)
        If Not AnyadirAFormula(Cadselect, cad) Then Exit Function
    End If
    
    If Devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            Cadparam = Cadparam & AnyadirParametroDH(param, TDes, THas, DesD, DesH) & """|"
            NumParam = NumParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Function AnyadirParametroDH(cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
     If TextoDESDE.Text <> "" Then
        cad = cad & "desde " & TextoDESDE.Text
        If TD.Text <> "" Then cad = cad & " - " & TD.Text
    End If
    If TextoHasta.Text <> "" Then
        cad = cad & "  hasta " & TextoHasta.Text
        If TH <> "" Then cad = cad & " - " & TH.Text
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function

'Para las reparaciones. Carga el importe real y teorico.
Private Sub CargaImporteRealReparaciones()
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

    'A partir de la reparacion , mirare en los albaranes, y de los albaranes ver el coste real de la reparacion y el teorico
    Set miRsAux = New ADODB.Recordset
    
    'Meto el select para las
    If Cadselect <> "" Then
        Codigo = Replace(Cadselect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    

 
    
    'Montamos el select al reves
    Codigo = "s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & Codigo
    Codigo = "s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & Codigo
    Codigo = "s.codtipom=l.codtipom AND " & Codigo
    Codigo = "sartic.codartic = l.codartic AND " & Codigo
    Codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & Codigo
    'EL ORDEN
    Codigo = Codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    NumRegElim = 1
    miSQL = ""
    While Not miRsAux.EOF
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                'INSERTAMOS
                ImpTeo = Round(ImpTeo, 2) * 100
                miSQL = miSQL & NumRegElim & "," & CLng(ImpTeo) & "," & TransformaComasPuntos(CStr(ImpTot)) & ")"
                'EXEcuete
                'en codprove llevare el numero de albaran
                'en codartic llevare el importe total teorico
                'en cantidad                    TOTAL
                miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad) " & miSQL
                conn.Execute miSQL
                NumRegElim = NumRegElim + 1
            End If
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            miSQL = " VALUES (" & vUsu.Codigo & "," & miRsAux!numrepar & ",'" & Format(miRsAux!fecrepar, FormatoFecha) & "',0,0,"
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    If miSQL <> "" Then
        'El ultimo
        ImpTeo = Round(ImpTeo, 2) * 100
        miSQL = miSQL & NumRegElim & "," & CLng(ImpTeo) & "," & TransformaComasPuntos(CStr(ImpTot)) & ")"
        'EXEcuete
        'en codprove llevare el importe total teorico
        'en cantidad                    TOTAL
        miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad) " & miSQL
        conn.Execute miSQL
    End If


    'La fecha hasta la tengo en la txtfecha(1)
    'Ahora pondere, en una (Y SOLO una) de las lineas el importe del mantenimiento hasta la fecha
    ' Las demas a CERO. Con lo cual, en el report, la suma del campo dara ESE importe solo
    '

    
   
    If Cadselect <> "" Then
        Codigo = Replace(Cadselect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    

    miSQL = "select numrepar,fecrepar,tieneman,"
    miSQL = miSQL & " mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act"
    miSQL = miSQL & " from schrep h,sserie s left join scaman m  on s.nummante=m.nummante and s.codclien=m.codclien"
    miSQL = miSQL & " where h.numserie=s.numserie and s.codartic=h.codartic "
    If Codigo <> "" Then miSQL = miSQL & Codigo
    
    'EL ORDEN
    IndiceImg = 12
    If txtFecha(1).Text <> "" Then IndiceImg = Month(CDate(txtFecha(1).Text))
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        ImpTot = 0
        If miRsAux!TieneMan = 1 Then
            '--------------------------------------------------------------------
            'OK, TIENE MANTENIMIENTO
            'Ire recorriendo los importes desde mes01act hasta el mes hasta
            'Si la fecha es fin es nada, entonces hare tooodos
            For NumRegElim = 1 To IndiceImg
                If Not IsNull(miRsAux.Fields(NumRegElim + 2)) Then ImpTot = ImpTot + miRsAux.Fields(NumRegElim + 2)
            Next
        End If
        If ImpTot <> 0 Then
            'UPDATEAMOS LA tmp
            miSQL = "UPDATE tmpnlotes set nomartic=" & CLng(ImpTot * 100) & " WHERE codusu = " & vUsu.Codigo
            miSQL = miSQL & " AND codprove = " & miRsAux!numrepar & " AND fechaalb = '" & Format(miRsAux!fecrepar, FormatoFecha) & "' AND numalbar =0"
            conn.Execute miSQL
        End If
        '--------------
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
End Sub
    


Private Sub EstadisticaReparacionTecnico()
    'Preparamos las temporales
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    Codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    

    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
    
    'Montamos el select al reves
    'PARA LAS FACTURAS
    If Cadselect <> "" Then
        Codigo = Replace(Cadselect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    Codigo = " s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & Codigo
    Codigo = " s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & Codigo
    Codigo = " s.codtipom=l.codtipom AND " & Codigo
    Codigo = " sartic.codartic = l.codartic AND " & Codigo
    Codigo = " h.numserie=sserie.numserie AND h.codclien=sserie.codclien AND " & Codigo
    Codigo = " sclien.codclien = h.codclien AND " & Codigo
    Codigo = " where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & Codigo
    'Las tablas
    Codigo = " from schrep h,slifac l,scafac1 s,sclien , sserie,sartic" & Codigo
    Codigo = "select l.*,s.fechaalb,preciove,h.fecrepar,nomclien,tieneman,h.nomartic,h.numserie " & Codigo
    'EL ORDEN
    Codigo = Codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
    

    
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "0,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!nomclien) & "')|"
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    

    
    'AHORA HAGO EL INSERT PARA LOS ALBARANES QUE NO HAN SIDO FACTURADOS
    'PARA LOS ALBARANES
    If Cadselect <> "" Then
        Codigo = Replace(Cadselect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If

    miSQL = "select l.*,preciove,tieneman,h.fechaalb,h.numserie,nomclien,fecrepar"
    miSQL = miSQL & " from schrep h,scaalb c,slialb l,sartic a,sserie s "
    miSQL = miSQL & " WHERE h.codtipom=c.codtipom and h.numalbar=c.numalbar and h.fechaalb=c.fechaalb and"
    miSQL = miSQL & " l.numalbar=c.numalbar and l.codtipom=c.codtipom and l.codartic=a.codartic"
    miSQL = miSQL & " and h.numserie=s.numserie and h.codclien =s.codclien" & Codigo
    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "1,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!nomclien) & "')|"
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    


End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
'               M U L T I B A S E
'------------------------------------------------------------------
Private Sub CargaListMultibase()
    Me.lstMultibase.Clear
    miSQL = "Clientes|Proveedores|Trabajadores|Direcciones|"
    For NumParam = 1 To 4
        Me.lstMultibase.AddItem RecuperaValor(miSQL, CInt(NumParam))
    Next NumParam
    'Como organiza informacion
    '         tabla  clave    campos a cambiar(empieza con coma) tipodatos clave.
    'Clientes
    miSQL = "sclien:codclien:,nomclien,nomcomer ,domclien ,codpobla ,pobclien,perclie1,perclie2:N|"
    miSQL = miSQL & "sprove:codprove:,nomprove,nomcomer ,domprove ,codpobla ,pobprove ,perprov1 ,perprov2:N|"
    miSQL = miSQL & "straba:codtraba:,nomtraba,domtraba,codpobla,pobtraba:N|"
    miSQL = miSQL & "sdirec:codclien,coddirec:,nomdirec ,domdirec ,pobdirec ,prodirec ,perdirec:N,N|"
        
End Sub


Private Sub HacerCambiosMultibase(numlinea As Integer)
Dim TotalReg As Long
Dim i As Integer
Dim J As Integer
Dim Claves As Integer
Dim Campos As Integer
Dim Cambios As Long
Dim T1 As Single
'Reutilizacion de variables
'cadTitulo cadNomRPT  conSubRPT

    On Error GoTo EHacerCambiosMultibase
    campo = lstMultibase.List(numlinea - 1)
    lblMultibase.Caption = "Preparando datos: " & campo
    
    lblMultibase.Refresh

    cadFormula = RecuperaValor(miSQL, numlinea)
    cadFormula = Replace(cadFormula, ":", "|")
    cadFormula = cadFormula & "|"  'Le añado el pipe final
    'Primero el conteo
    Cadparam = "Select count(*) from " & RecuperaValor(cadFormula, 1)
    miRsAux.Open Cadparam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalReg = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then TotalReg = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    DoEvents
    If TotalReg = 0 Then
        lblMultibase.Caption = "Tabla vacia " & campo
        lblMultibase.Refresh
        Espera 1
    End If
    
    'Veamos cuantos campos hay que ver la conversion de campos, y las claves
    Cadparam = RecuperaValor(cadFormula, 2)
    Claves = 1
    Cambios = 0
    While Cadparam <> ""
        NumRegElim = InStr(1, Cadparam, ",")
        If NumRegElim = 0 Then
            Cadparam = ""
        Else
            Claves = Claves + 1
            Cadparam = Mid(Cadparam, NumRegElim + 1)
        End If
    Wend
    Cadparam = RecuperaValor(cadFormula, 3)
    Campos = 0 'aqui cero pq empieza con la coma
    While Cadparam <> ""
        NumRegElim = InStr(1, Cadparam, ",")
        If NumRegElim = 0 Then
            Cadparam = ""
        Else
            Campos = Campos + 1
            Cadparam = Mid(Cadparam, NumRegElim + 1)
        End If
    Wend
        

                            'claves                                 'campos cambiar
    Cadparam = "SELECT " & RecuperaValor(cadFormula, 2) & RecuperaValor(cadFormula, 3)
    Cadparam = Cadparam & " FROM " & RecuperaValor(cadFormula, 1)
    miRsAux.Open Cadparam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Cambios = 0
    T1 = Timer   'Para hacer doevents cada 3 segundos
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'Los labels
        lblMultibase.Caption = campo & " ( " & NumRegElim & " / " & TotalReg & " )"
        lblMultibase.Refresh
        If Timer - T1 > 3 Then
            DoEvents
            Me.Refresh
            Espera 0.2
            T1 = Timer
        End If
        
        Cadselect = "" 'LOS UPDATES
        For i = Claves To Campos
            If Not IsNull(miRsAux.Fields(i)) Then
                Cadparam = miRsAux.Fields(i)  'Cojo el valor del field
                cadNomRPT = RevisaCaracterMultibase(Cadparam)  'Obtengo la modificaicon por campos multibase
                If Cadparam <> cadNomRPT Then
                    'HAY que modificar ya que son disitintos el de laBD y el calculado por el modulo de multibase
                    Cadselect = Cadselect & ", " & miRsAux.Fields(i).Name & " = '" & DevNombreSQL(cadNomRPT) & "'"
                End If
            End If
        Next
        'SI cadselect <>"" entonces hay que ejecutar SQL
        If Cadselect <> "" Then
            'Los campos claves van del 0 a claves -1
            Cadparam = ""
            cadTitulo = RecuperaValor(cadFormula, 4) 'los tipos de datos
            cadTitulo = Replace(cadTitulo, ",", "|") & "|"
            For J = 0 To Claves - 1
                Cadparam = Cadparam & " AND " & miRsAux.Fields(J).Name & " = "
                Codigo = RecuperaValor(cadTitulo, J + 1)

                Select Case Codigo
                Case "F"
                    Cadparam = Cadparam & "'" & Format(miRsAux.Fields(i).Value, FormatoFecha) & "'"
                Case "T"
                    Cadparam = Cadparam & "'" & miRsAux.Fields(i).Value & "'"
                Case Else  'NUMERICO
                    Cadparam = Cadparam & miRsAux.Fields(J).Value
                End Select
            Next J
            
            
            'Acabas de montar el UPDATE
            cadTitulo = "UPDATE " & RecuperaValor(cadFormula, 1)
            Cadselect = Mid(Cadselect, 2)   'QUITO la coma
            Cadparam = Mid(Cadparam, 5)     'QUITO el primer AND
            cadTitulo = cadTitulo & " SET " & Cadselect & " WHERE " & Cadparam
            conn.Execute cadTitulo
            Cambios = Cambios + 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    lblMultibase.Caption = "FIN " & campo
    lblMultibase.Refresh
    If Cambios > 0 Then Me.Tag = Me.Tag & vbCrLf & "   .- " & campo & " : " & Cambios
    Exit Sub
EHacerCambiosMultibase:
    MuestraError Err.Number
End Sub
'       fin mULTIBASE
'------------------------------------------------------------------'------------------------------------------------------------------


'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion de recargas de telefonia
'
'------------------------------------------------------------------
'------------------------------------------------------------------



Private Sub HacerFacturacionTelefonia(vAlbaranes As Collection, MenError As String)
Dim RT As ADODB.Recordset
Dim b As Boolean
Dim NumAlb As String
Dim Almacen As Integer



    'El proceso sera el siguiente:
    'Voy a agrupar por dia (podria ser por mes),trabajador
    'Y para cada uno de los resultados del recodset voy a generar un albaran.
    'Me guardare los albaranes generados y despues los facturare.
    'Para ello
    campo = "Select codtraba,count(*) as cantidad,sum(importe)as total from stelefonia WHERE " & Cadselect & " GROUP by codtraba"
    
    Set RT = New ADODB.Recordset
    RT.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        
        Almacen = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", CStr(RT!CodTraba), "N")
        
        conn.BeginTrans
        
        'Obtener el contador de Albaran (ALV).
        b = ObtenerContadorAlbaran(NumAlb)
        
        If b Then
            'Actualizar los stocks de todos los articulos comprados
            'Insertar movimiento en smoval
            'B = InsertarMovAlmacen(NumAlb)  ¿ FALTA### ?
    
            'Insertar en las tablas de Albaranes: scaalb, slialb
            'en el campo scafac1.numalbar guardamos el nº de ticket
            If b Then b = InsertarAlbaran(NumAlb, CStr(RT!CodTraba), 1, RT!Cantidad, RT!Total, MenError)
        
        End If



       
        If Not b Then
            conn.RollbackTrans
            RT.Close
            Set RT = Nothing
            Exit Sub
        Else
            vAlbaranes.Add CStr(NumAlb)
            conn.CommitTrans
            
            'Le pongo a facturado en la telefonia
            miSQL = "UPDATE stelefonia SET facturado = 1 WHERE " & Cadselect & " AND codtraba = " & RT!CodTraba
            conn.Execute miSQL
        End If
    
    
        RT.MoveNext
    Wend
    RT.Close
    


End Sub


Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer("ALV") Then
        Do
            NumAlb = vTipoMov.ConseguirContador("ALV")
            vTipoMov.IncrementarContador ("ALV")
            miSQL = "select count(*) from scaalb where codtipom='ALV' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(miSQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function

Private Function InsertarAlbaran(NumAlb As String, CodTraba As String, CodAlmc As Integer, Cantidad As Currency, Importe As Currency, menErr As String) As Boolean
Dim b As Boolean
Dim vClien As CCliente
Dim SQL As String

    On Error GoTo EInsAlb



    'Cabecera de albaran
    '----------------------------------
    SQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    SQL = SQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa) "
                                                                    'Facturar   cliente
    SQL = SQL & " VALUES ('ALV'," & NumAlb & "," & DBSet(Now, "F") & ",1," & txtCliente(2).Text & ","
    
    'Obtenemos los datos del cliente
    Set vClien = New CCliente
    If vClien.Existe(txtCliente(2).Text) Then
        If vClien.LeerDatos(txtCliente(2).Text) Then
            SQL = SQL & DBSet(vClien.Nombre, "T", "N") & ", " & DBSet(vClien.Domicilio, "T", "N") & ","
            SQL = SQL & DBSet(vClien.CPostal, "T", "N") & ", " & DBSet(vClien.Poblacion, "T", "N") & "," & DBSet(vClien.Provincia, "T", "N") & ","
            SQL = SQL & DBSet(vClien.NIF, "T", "N") & "," & DBSet(vClien.TfnoClien, "T") & ","
            'coddirec,nomdirec,referenc a nulo
            SQL = SQL & "NULL,NULL,NULL,"
            
            SQL = SQL & CodTraba & "," & CodTraba & "," & CodTraba & "," 'trabajador
            '                              cod forpa
            SQL = SQL & vClien.Agente & ",1," & vClien.FEnvio & ",0,0," & vClien.TipoFactu & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'observaciones
            SQL = SQL & ValorNulo & "," & ValorNulo & "," 'datos oferta: aqui guardamos nº venta
            'En los campos de datos del pedido guardamos los datos del ticket
            'SQL = SQL & NumTicket & "," & DBSet(RSVenta!fecventa, "F") & "," & ValorNulo & "," & ValorNulo & ",1," & DBSet(RSVenta!NumTermi, "N") & "," & DBSet(RSVenta!NumVenta, "N", "S") & ")" 'esticket=1, terminal
            SQL = SQL & "NULL,NULL," & ValorNulo & "," & ValorNulo & ",0,NULL,NULL)"
            b = vClien.ActualizaUltFecMovim(Now)
        Else
            b = False
        End If
    End If
    Set vClien = Nothing
    
    
    If b Then
        'Insertar Cabecera
'    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
        conn.Execute SQL, , adCmdText
        
        'Lineas del albaran
        'Inserta en tabla "slialb" todas las lineas de venta
        SQL = "INSERT INTO slialb "
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, "
        SQL = SQL & "dtoline1, dtoline2, importel, origpre) VALUES ("
        SQL = SQL & "'ALV'," & DBSet(NumAlb, "N") & ",1," & CodAlmc & ",'" & DevNombreSQL(txtArticulo(0).Text) & "','" & DevNombreSQL(txtDescArticulo(0).Text)
        SQL = SQL & "',NULL," & Cantidad & "," & TransformaComasPuntos(CStr(Round(Importe / Cantidad, 4))) & ",0,0," & TransformaComasPuntos(CStr(Importe)) & ",'')"
        'SQL = SQL & " FROM sliven WHERE " & Replace(cadSel, "scaven", "sliven")
        conn.Execute SQL, , adCmdText
    End If


    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    If b Then cadImpresion = "{scaalb.codtipom}='ALV' and {scaalb.numalbar}=" & DBSet(NumAlb, "N")

EInsAlb:
    If Err.Number <> 0 Then
        menErr = "Insertando el Albaran: " & vbCrLf & Err.Description
        b = False
    End If
    InsertarAlbaran = b
End Function



'De momento no miro si tiene DTOs o no. Simplemente acltualizo precio y redondeo
'a dos decimales
Private Function RealizarCambiosPreciosLiq(ByRef FechaUltCompra As Date) As Boolean


    On Error GoTo ERealizarCambiosPreciosLiq
    RealizarCambiosPreciosLiq = False
    
    cadFormula = "UPDATE slialp Set precioar = " & TransformaComasPuntos(CStr(ImpTeo)) & " , importel = "
    cadTitulo = "UPDATE smoval SET impormov = "
    miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'Label
        Me.lblLiqu.Caption = miRsAux!NumAlbar & " - " & miRsAux!FechaAlb & " : " & miRsAux!codartic
        Me.lblLiqu.Refresh
        
        ImpTot = miRsAux!Cantidad * ImpTeo
        ImpTot = Round2(ImpTot, 2)
        Devuelve = TransformaComasPuntos(CStr(ImpTot)) & " WHERE numalbar = '" & DevNombreSQL(miRsAux!NumAlbar) & "'"
        Devuelve = Devuelve & " And fechaalb = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "' AND"
        Devuelve = Devuelve & " codprove = " & miRsAux!codProve
        Devuelve = Devuelve & " AND numlinea = " & miRsAux!numlinea
        Devuelve = cadFormula & Devuelve
        conn.Execute Devuelve
        
        'UPDATEO smoval
        Devuelve = cadTitulo & TransformaComasPuntos(CStr(ImpTot))
        Devuelve = Devuelve & " WHERE detamovi = 'ALC' AND fechamov = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
        Devuelve = Devuelve & " AND codigope = " & miRsAux!codProve & " AND document = '" & DevNombreSQL(miRsAux!NumAlbar) & "'"
        Devuelve = Devuelve & " AND codartic = '" & DevNombreSQL(miRsAux!codartic) & "' AND numlinea =" & miRsAux!numlinea
        conn.Execute Devuelve
        
        'Si el albaran es masyor que la utlima fecha de compra entonces
        If miRsAux!FechaAlb > FechaUltCompra Then
            NumParam = 1
            FechaUltCompra = miRsAux!FechaAlb
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close



    RealizarCambiosPreciosLiq = True
    Exit Function
ERealizarCambiosPreciosLiq:
    MuestraError Err.Number, Err.Description
End Function






'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion y contabilizacion de tickets
'       ========================================
'
'
'
'
'       Cuando esta la marca de contabilizar tickets agrupados, lo que haremos sera
'       a partir de los FTI crear las facturas agrupadas con el contador FTG "EN LA CONTABILIDAD"
'       en el ariges, en scafac, no creo ninguna factura
'       O bien una diaria o una mensual (dependera del parametro)
'
'
'       Insertaremos en una tabla los tckets que entran en cada factura
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub HacerFacturaTICKETS(ByVal ClienteVarios As Long)
Dim b As Boolean
  
    
        'Si va agrupado por fecha, o no
        Label5.Caption = "Obteniendo facturas"
        Label5.Refresh
        
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        miSQL = "Desde " & txtFecha(20).Text & " hasta " & txtFecha(21).Text & vbCrLf
        miSQL = miSQL & "Diario: " & CStr(Me.optTick(0).Value) & vbCrLf
        miSQL = miSQL & "Trabajador: " & txtTrab(2).Text & " " & Me.txtDescTra(2).Text
        LOG.Insertar 6, vUsu, miSQL
        miSQL = ""
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        If Me.optTick(1).Value Then
            'MENSUAL
            Devuelve = Format(txtFecha(21).Text, FormatoFecha)
            b = ObtenerDatosTickets(False, ClienteVarios)
        Else
            'Veo las fechas
            'Y para cada fecha
            miSQL = "Select fecfactu from scafac WHERE " & Cadselect & " GROUP by fecfactu ORDER BY fecfactu"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            campo = ""
            While Not miRsAux.EOF
                campo = campo & Format(miRsAux.Fields(0), FormatoFecha) & "|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'Ya tengo todas las fechas que voy a tratar
            While campo <> ""
                   NumParam = InStr(1, campo, "|")
                   If NumParam = 0 Then
                        campo = ""
                   Else
                        Devuelve = Mid(campo, 1, NumParam - 1)
                        campo = Mid(campo, NumParam + 1)
                    
                        Label5.Caption = "Obteniendo facturas. Fec: " & Devuelve
                        Label5.Refresh
                    
                    
                        'CONTABILIZAMOS LA FACTURA ESA
                        b = ObtenerDatosTickets(True, ClienteVarios)
                        'Se ha producido un error.Salgo aunaque queden fecs por tratar
                        If Not b Then campo = ""
                            
                   End If
            Wend
            
        End If
        Set miRsAux = Nothing
            
            
        If b Then
            'AHORA LANZAREMOS A CONTABILIZAR FACTURAS de frmlistado
            Label5.Caption = "Contablizando FTGs"
            Label5.Refresh
            AbrirListado 248   'Contabilizacion de facturas TICKET AGRUPADAS
            
            
            Label5.Caption = "Comprobando contabilizacion"
            Label5.Refresh
            DoEvents
            
            
            'Aqui viene la fiesta. Vere si hay facturas FTG con intconta=0
            'Significara que han dado error al entrar en la conta
            Set miRsAux = New ADODB.Recordset
            miSQL = "Select numfactu from scafac where codtipom='FTG' And intconta=0"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux!NumFactu) Then b = False
            End If
            miRsAux.Close
            Set miRsAux = Nothing
                       
            
        End If

            
        If Not b Then
            Screen.MousePointer = vbHourglass
            Label5.Caption = "Reestableciendo FTI. Paso 1"
            Label5.Refresh
            'HA IDO MAL
            'Vuelvo a poner los FTI que haya puesto a contabilizado, a 0
            
            
            'Dos pasos:
            'Primero ver que facturas FTG se han generado.
            'Las meto en la variable cadfrom
            
            Set miRsAux = New ADODB.Recordset
            miSQL = "Select numfactu,fecfactu from scafac where codtipom='FTG' And intconta=0"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cadFrom = ""
            While Not miRsAux.EOF
                cadFrom = cadFrom & " numfacftg =  " & miRsAux!NumFactu & " AND fecfacftg = '" & Format(miRsAux!FecFactu, FormatoFecha) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            'Segundo.
            'Para cad factura FTG generada veo que FTI asoacidos tiene y los updateo
            Label5.Caption = "Reestableciendo FTI. Paso 2"
            Label5.Refresh
            miSQL = "UPDATE scafac SET intconta=0 WHERE codtipom='FTI' AND numfactu ="
            While cadFrom <> ""
                NumParam = InStr(1, cadFrom, "|")
                If NumParam = 0 Then
                    cadFrom = ""
                Else
                    Devuelve = Mid(cadFrom, 1, NumParam - 1)
                    cadFrom = Mid(cadFrom, NumParam + 1)
                         
                    Devuelve = "Select numfactu,fecfactu FROM sfactik where " & Devuelve
                    miRsAux.Open Devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not miRsAux.EOF
                
                        Devuelve = miSQL & miRsAux!NumFactu & " AND fecfactu = '" & Format(miRsAux!FecFactu, FormatoFecha) & "'"
                        conn.Execute Devuelve
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
                End If
            Wend
       
            
                    
            Me.Refresh
            Espera 0.5
            Label5.Caption = "Eliminado asociaciones"
            Label5.Refresh
            
            'Si ha ido mal entonces borraremos tanto los FTG (proceso que se hace despues)
            'como en la tabla que asocia con los tickets
            ' REestablecer en contadores
            ' devuelve= MINIMO
            miSQL = "Select numfactu,fecfactu from scafac where codtipom='FTG' AND intconta=0"
            miSQL = miSQL & " GROUP BY numfactu,fecfactu ORDER BY numfactu"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Devuelve = ""
            While Not miRsAux.EOF
                If Devuelve = "" Then Devuelve = miRsAux!NumFactu & "|'" & Format(miRsAux!FecFactu, FormatoFecha) & "|"
                miSQL = "DELETE from sfactik WHERE numfacftg=" & miRsAux!NumFactu & " AND fecfacFTG='" & Format(miRsAux!FecFactu, FormatoFecha) & "'"
                conn.Execute miSQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            'Pong los contadores como estaban
            If Devuelve <> "" Then
                NumRegElim = Val(RecuperaValor(Devuelve, 1))
                'devuelve = RecuperaValor(devuelve, 1)
                miSQL = "UPDATE stipom SET contador = " & NumRegElim & " WHERE codtipom='FTG'"
                conn.Execute miSQL
            End If
        End If 'De si ha ido bien o mal
        
        'BORRAMOS todos los datos
        '------------------------------------------
        Label5.Caption = "Eliminando datos temporales de tablas scafac..."
        Label5.Refresh
        DoEvents
        
        miSQL = "DELETE  from slifac where codtipom='FTG'"
        conn.Execute miSQL
        
        

        'Habre metido una linea en scafac1
        miSQL = "DELETE  from scafac1 where codtipom='FTG'"
        conn.Execute miSQL

        
        miSQL = "DELETE  from scafac where codtipom='FTG'"
        conn.Execute miSQL
        
        
        
        'Si todo ha ido bien muestro un msg
        Label5.Caption = ""
        Label5.Refresh
        If b Then MsgBox "Proceso finalizado con éxito", vbInformation
        
        Screen.MousePointer = vbDefault
End Sub





Private Function ObtenerDatosTickets(Diario As Boolean, CodCliVarios As Long) As Boolean
Dim TiposIva As Byte
Dim vTipom As CTiposMov
Dim vCli As CCliente

        On Error GoTo EObteniendoDatosTickets


        ObtenerDatosTickets = False



        'En la tabla tmpspla pondre todos los importes por tp iva
        conn.Execute "DELETE from tmpinformes where codusu = " & vUsu.Codigo
        
        
        'Veo todos los importes y bases imponibles etc
        'Para no tener que hacer selects y demas me guardare que tipos de iva estoy tratatando
        '
        cadNomRPT = "|"
        TiposIva = 0
        For NumParam = 1 To 3
            miSQL = "SELECT codigiv" & NumParam & " tipodeiva,sum(baseimp" & NumParam & ") labase,sum(imporiv" & NumParam & ") importeiva FROM SCafac where "
            miSQL = miSQL & " intconta=0 and codtipom='FTI'"
            If Diario Then
                miSQL = miSQL & " AND fecfactu='" & Devuelve & "'"
            Else
                'MOdificacion 13 - Agosto - 2008
                'Si no pongo esto suma tooooodas las facturas FTI que no esten contabilizadas
                'Desde
                If txtFecha(20).Text <> "" Then miSQL = miSQL & " AND fecfactu>='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
                'El campo HASTA es obligado
                miSQL = miSQL & " AND fecfactu<='" & Format(txtFecha(21).Text, FormatoFecha) & "'"
            End If
            
            miSQL = miSQL & " group by 1 "
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
            
                If Not IsNull(miRsAux!tipodeiva) Then
                    ImpTot = DBLet(miRsAux!labase, "N")
                    ImpTeo = DBLet(miRsAux!ImporteIva, "N")
                    miSQL = "|" & miRsAux!tipodeiva & "|"
                    
                    If InStr(1, cadNomRPT, miSQL) > 0 Then
                        'YA LO HE INSERTADO. UPDATEO
                        miSQL = "UPDATE tmpinformes SET importe1=importe1 + " & TransformaComasPuntos(CStr(ImpTot))
                        miSQL = miSQL & " ,importe2=importe2 + " & TransformaComasPuntos(CStr(ImpTeo))
                        miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND codigo1 = " & miRsAux!tipodeiva
                    Else
                        miSQL = "INSERT INTO `tmpinformes` (`codusu`,`codigo1`,`importe1`,importe2) values (" & vUsu.Codigo & "," & miRsAux!tipodeiva
                        miSQL = miSQL & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & ")"
                        TiposIva = TiposIva + 1
                        cadNomRPT = cadNomRPT & miRsAux!tipodeiva & "|"
                    End If
                    conn.Execute miSQL
                
                End If
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Next NumParam
        
        If TiposIva > 3 Or cadNomRPT = "" Then
            'ERROR  ERROR ERROR
            'ERROR en los tipos de iva. Hay mas de 3 o no hay ninguno
            If cadNomRPT = "" Then TiposIva = 0
            cadNomRPT = "Error en los tipos de IVA." & vbCrLf & "Total IVAS: " & TiposIva & vbCrLf & " Fec: " & Devuelve
            MsgBox cadNomRPT, vbExclamation
            Exit Function
        End If
        
        'Ya tengo las bases ivas para las facturas
        'Ahora creo la FTG para poder utilizar las funciones ya realizadas
        Set vCli = New CCliente
        
        vCli.LeerDatos CStr(CodCliVarios)
        
             Set vTipom = New CTiposMov
             vTipom.Leer "FTG"
             vTipom.ConseguirContador vTipom.TipoMovimiento
             
             miSQL = "INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             miSQL = miSQL & "`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             miSQL = miSQL & "`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             miSQL = miSQL & "`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
             'LOS IVAS
             miSQL = miSQL & "`baseimp1`,`codigiv1`,`porciva1`,`imporiv1`,"
             miSQL = miSQL & "`baseimp2`,`codigiv2`,`porciva2`,`imporiv2`,"
             miSQL = miSQL & "`baseimp3`,`codigiv3`,`porciva3`,`imporiv3`)"
             
             'Cargo los ivas
             cadNomRPT = "Select codigo1,importe1,importe2 from tmpinformes where codusu = " & vUsu.Codigo
             miRsAux.Open cadNomRPT, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
             cadNomRPT = ""
             TiposIva = 0
             ImpTot = 0
             ImpTeo = 0
             While Not miRsAux.EOF
                 TiposIva = TiposIva + 1
                 Codigo = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", miRsAux!Codigo1)
                 cadFrom = "," & TransformaComasPuntos(CStr(miRsAux!Importe1)) & "," & miRsAux!Codigo1 & "," & TransformaComasPuntos(Codigo) & ","
                 cadFrom = cadFrom & TransformaComasPuntos(CStr(miRsAux!importe2))
                 
                 'Meto en el select
                 cadNomRPT = cadNomRPT & cadFrom
                 
                 'ImpTot
                 ImpTot = ImpTot + miRsAux!Importe1
                 ImpTeo = ImpTeo + miRsAux!importe2
                 miRsAux.MoveNext
             Wend
             miRsAux.Close
                 
                 
             'Si no tiene 3 tipos de ivas meter los null
             For NumParam = TiposIva + 1 To 3
                 cadNomRPT = cadNomRPT & ",NULL,NULL,NULL,NULL"
             Next
             
             
             'Ahora relleno los datos que faltan
             'INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             '`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             '`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             '`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
                         
             cadFrom = " VALUES ('" & vTipom.TipoMovimiento & "'," & vTipom.contador & ",'" & Devuelve & "'," & vCli.Codigo
             cadFrom = cadFrom & ",'" & vCli.Nombre & "','','0','','','0',NULL,NULL,NULL" '0: codpos y nif
             'Agente:
             cadFrom = cadFrom & "," & vCli.Agente & "," & vCli.ForPago & ",0,0,NULL,NULL,NULL,NULL,"
             'Bruto factra
             cadFrom = cadFrom & "" & TransformaComasPuntos(CStr(ImpTot)) & ",0,0,0," & TransformaComasPuntos(CStr(ImpTot + ImpTeo))
              
             miSQL = miSQL & cadFrom & cadNomRPT & ")"
             conn.Execute miSQL
             
            'Si lleva la analitica metere una linea en slifac1 que es desde donde,
            ' el proceso de contabilizacion cojera EL CODTRABA para obtener el CC
                
                miSQL = "insert into `scafac1` (`codtipom`,`numfactu`,`fecfactu`,codtipoa,numalbar,`codenvio`,`codtraba`,`codtrab1`,`codtrab2`)"
                miSQL = miSQL & " VALUES ('FTG'," & vTipom.contador & ",'" & Devuelve & "','DAV','8',"  'Pongo tipoa y numalbar a piñon
                miSQL = miSQL & vParamAplic.PorDefecto_Envio & "," & txtTrab(2).Text & "," & txtTrab(2).Text & "," & txtTrab(2).Text & ")"
                conn.Execute miSQL
            
            
            
            
            'Ahora, despues de crear la factura temporal FTG, insertare en la tabla
            'que lleva la relacion, numfactura, codticket
            miSQL = "INSERT INTO sfactik(`numfacFTG`,`fecfacFTG`,`numfactu`,`fecfactu`,`codtraba`)"
            miSQL = miSQL & " SELECT " & vTipom.contador & ",'" & Devuelve & "',numfactu,fecfactu," & txtTrab(2).Text & " FROM scafac where "
            miSQL = miSQL & Cadselect
            If Diario Then miSQL = miSQL & " AND fecfactu='" & Devuelve & "'"
            conn.Execute miSQL
             
            'Pongo la marca de contabilizado
            miSQL = "UPDATE scafac SET intconta = 1 WHERE " & Cadselect
            If Diario Then miSQL = miSQL & " AND fecfactu='" & Devuelve & "'"
            conn.Execute miSQL
             
            vTipom.IncrementarContador vTipom.TipoMovimiento
            ObtenerDatosTickets = True
            

    

EObteniendoDatosTickets:

    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & miSQL
    End If
    Set vCli = Nothing
    Set vTipom = Nothing
End Function




'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Informe de trazabilidad
'       ========================================
'
'
'
'
'       A partir del desde /hasta mostraremos el informe que tiene la asociacion
'       entre albaranes de compra / venta
'
'
'       Hay datos tanto en albaranes como en facturas, con lo cual insertare sobre tmp
'
'
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub HacerInformeTrazabilidad()

    
    
    InicializarVbles
    
    Cadparam = Cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    
    If txtFecha(22).Text <> "" Or txtFecha(23).Text <> "" Then
        campo = "{slcomven.fechaalbc}"
        Devuelve = "pDHFamilia=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 22, 23, Devuelve) Then Exit Sub
    End If
    
    If txtCodProve(10).Text <> "" Or txtCodProve(11).Text <> "" Then
        campo = "{slcomven.codprovec}"
        Devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "PRO", 10, 11, Devuelve) Then Exit Sub
    End If
     
    If txtArticulo(4).Text <> "" Or txtArticulo(5).Text <> "" Then
        campo = "{slcomven.codartic}"
        Devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 4, 5, Devuelve) Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    Cadselect = QuitarCaracterACadena(Cadselect, "{")
    Cadselect = QuitarCaracterACadena(Cadselect, "}")
    If Cadselect = "" Then Cadselect = " 1 = 1 "
    campo = "slcomven WHERE  " & Cadselect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    cadNomRPT = "rTraza.rpt"
    LlamarImprimir2
    
End Sub


Private Sub AñadirMarcas()
        '
        miSQL = ""
        For NumRegElim = 1 To Me.List1.ListCount
            miSQL = miSQL & List1.ItemData(NumRegElim - 1) & "|"
        Next
        If miSQL <> "" Then miSQL = "|" & miSQL
        CadenaDesdeOtroForm = miSQL
        frmVarios.opcion = 5
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            List1.Clear
            miSQL = "Select * from smarca where codmarca in " & CadenaDesdeOtroForm & " ORDER BY nommarca"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                List1.AddItem miRsAux!nommarca & "   (" & miRsAux!Codmarca & ")"
                List1.ItemData(List1.NewIndex) = miRsAux!Codmarca
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
End Sub



'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Ventas por agente
'       ========================================
'
'
'
'
'
'----------------------------------------------------------------------
'----------------------------------------------------------------------

Private Function ObtenerDatosVentasAgentes() As Boolean
Dim Litr As Currency

    On Error GoTo EObtenerDatosVentasAgentes

    ObtenerDatosVentasAgentes = False

    'Temporales
    miSQL = "DELETE FROM olitmpventasagente where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    
    
    
    'Cadselext traera los desde /hasta
    
    'Ventas productos conjunto=1  fechafac
    miSQL = "select scafac.codagent,slifac.codartic,cantidad,LitrosUnidad,comisionaceite,codclien,scafac.numfactu,importel,scafac.fecfactu  from scafac ,slifac,sartic,sagent where"
    miSQL = miSQL & " scafac.codTipoM = slifac.codTipoM And scafac.NumFactu = slifac.NumFactu"
    miSQL = miSQL & " And scafac.FecFactu = slifac.FecFactu and sagent.codagent=scafac.codagent"
    miSQL = miSQL & " and slifac.codartic =sartic.codartic and conjunto =1"
    If Cadselect <> "" Then miSQL = miSQL & " AND " & Cadselect
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'Por l,itros en formatos mayores y uds en menores
        If miRsAux!LitrosUnidad > 1 Then
            Litr = miRsAux!LitrosUnidad
        Else
            Litr = 1
        End If
        Litr = miRsAux!Cantidad * Litr
        
        miSQL = "insert into `olitmpventasagente` (`codusu`,`codigo`,`codartic`,`codagent`,`cantidad`,`importe`,cliente,factura,base,fechafac) values ( "
        miSQL = miSQL & vUsu.Codigo & "," & NumRegElim & ",'" & miRsAux!codartic & "'," & miRsAux!codagent & ","
        'Cantidad y importe
        miSQL = miSQL & TransformaComasPuntos(CStr(Litr)) & ","
        Litr = Round2(miRsAux!comisionaceite * Litr, 4)
        miSQL = miSQL & TransformaComasPuntos(CStr(Litr)) & ","
        'Cliente factura
        miSQL = miSQL & miRsAux!CodClien & ",'"
        miSQL = miSQL & Format(miRsAux!NumFactu, "000000") & "',"
        'importel
        miSQL = miSQL & TransformaComasPuntos(CStr(miRsAux!ImporteL)) & ",'"
        miSQL = miSQL & Format(miRsAux!FecFactu, FormatoFecha) & "')"
        conn.Execute miSQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ventas directas de materia prima
    miSQL = "select scafac.codagent,slifac.codartic,cantidad,LitrosUnidad,comisionaceite,codclien,scafac.numfactu,importel,scafac.fecfactu from scafac ,slifac,sartic,sagent where"
    miSQL = miSQL & " scafac.codTipoM = slifac.codTipoM And scafac.NumFactu = slifac.NumFactu"
    miSQL = miSQL & " And scafac.FecFactu = slifac.FecFactu and sagent.codagent=scafac.codagent"
    miSQL = miSQL & " and slifac.codartic =sartic.codartic and factorconversion<>1"
    If Cadselect <> "" Then miSQL = miSQL & " AND " & Cadselect
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Litr = miRsAux!Cantidad  'Seran kilos
        
        miSQL = "insert into `olitmpventasagente` (`codusu`,`codigo`,`codartic`,`codagent`,`cantidad`,`importe`,cliente,factura,base,fechafac) values ( "
        miSQL = miSQL & vUsu.Codigo & "," & NumRegElim & ",'" & miRsAux!codartic & "'," & miRsAux!codagent & ","
        'Cantidad y importe
        miSQL = miSQL & TransformaComasPuntos(CStr(Litr)) & ","
        Litr = Round2(miRsAux!comisionaceite * Litr, 4)
        miSQL = miSQL & TransformaComasPuntos(CStr(Litr)) & ","
        'Cliente factura
        miSQL = miSQL & miRsAux!CodClien & ",'"
        miSQL = miSQL & Format(miRsAux!NumFactu, "00") & "',"
        'importel
        miSQL = miSQL & TransformaComasPuntos(CStr(miRsAux!ImporteL)) & ",'"
        miSQL = miSQL & Format(miRsAux!FecFactu, FormatoFecha) & "')"
        conn.Execute miSQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If NumRegElim = 0 Then
        MsgBox "Ningun dato", vbExclamation
    Else
        ObtenerDatosVentasAgentes = True
    End If
    Exit Function
EObtenerDatosVentasAgentes:
    MuestraError Err.Number
End Function


Private Sub CargaAgentes()
Dim F As Date
    List2.Clear
    Set miRsAux = New ADODB.Recordset
    miSQL = "Select codagent,nomagent from sagent ORDER BY nomagent"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        List2.AddItem miRsAux!nomagent
        List2.ItemData(List2.NewIndex) = miRsAux!codagent
        List2.Selected(List2.NewIndex) = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    'Las fechas
    F = DateAdd("m", -1, Now)
    miSQL = "/" & Format(Month(F), "00") & "/" & Year(F)
    txtFecha(30).Text = "01" & miSQL
    NumRegElim = DiasMes(CByte(Month(F)), Year(F))
    txtFecha(31).Text = Format(NumRegElim, "00") & miSQL
End Sub



Private Sub CargaClientes()
    Set miRsAux = New ADODB.Recordset
    
    'CadenaDesdeOtroForm
    'Dira que tarifa / TO  es
    'Cargamos la tarifa oferta que coressponda
    List3.Clear
    miSQL = "Select * from olitarifaoferta where codigo = " & CadenaDesdeOtroForm
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Devuelve = "0" 'Por si acaso es tarifa
    Label7.Caption = "Tarifa"
    If Not miRsAux.EOF Then
        NumRegElim = Val(CadenaDesdeOtroForm)
        txtTarifa.Text = NumRegElim
        If NumRegElim > 100000 Then
            'Es una TO
            Label7.Caption = "Tarifa-Oferta"
            
        Else
           
            Devuelve = miRsAux!Tarifa
        End If
    End If
    miRsAux.Close
    SSTab1.TabVisible(0) = Devuelve <> "0"  'Las tarifas
    If NumRegElim = 0 Then
        MsgBox "Error buscando TO: " & CadenaDesdeOtroForm, vbExclamation
        Exit Sub
    End If
    'con lo cual si es una oferta cargaremos todos los clientes
    'si es una TO, solo la TO esa
    
    
    
    miSQL = "Select sclien.codclien,nomclien from sclien"
    If Val(CadenaDesdeOtroForm) > 100000 Then
        'ES UNA TO
        miSQL = miSQL & ",olitarifaoferta WHERE sclien.codclien=olitarifaoferta.codclien "
        miSQL = miSQL & " and olitarifaoferta.codigo = " & CadenaDesdeOtroForm
        
    Else
        miSQL = miSQL & " WHERE codtarif = " & Devuelve
    End If
    
    miSQL = miSQL & "   ORDER BY nomclien"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        List3.AddItem miRsAux!nomclien
        List3.ItemData(List3.NewIndex) = miRsAux!CodClien
        List3.Selected(List3.NewIndex) = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    txtFecha(32).Text = Format(Now, "dd/mm/yyyy")
    LeerEscribirDatosCarta True
End Sub






Private Function GenerarDatosEncoenves() As Boolean
Dim C As String
Dim RS As ADODB.Recordset
Dim Fin As Boolean

    On Error GoTo EgenerarDatosEncoenves

    GenerarDatosEncoenves = False

    'Temporales
    miSQL = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    'variables
    
    'Incio Rs
    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset

    
    'Ventas productos conjunto=1  fechafac
    miSQL = "select numfactu,fecfactu from slifac WHERE"
    miSQL = miSQL & " slifac.codartic = '" & vParamAplic.ArtReciclado & "'"
    If Cadselect <> "" Then miSQL = miSQL & " AND " & Cadselect
    miSQL = miSQL & " GROUP BY 1,2 order by 1,2   "
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    campo = ""
    While Not miRsAux.EOF
        'Fra
        Label3(72).Caption = miRsAux!NumFactu & " - " & miRsAux!FecFactu
        Label3(72).Refresh
        C = ""
        'Para cada factura
        Fin = False
        miSQL = "Select slifac.*,sunida.codunida,nomunida from slifac,sartic,sunida WHERE "
        miSQL = miSQL & " slifac.codartic=sartic.codartic AND  sartic.codunida=sunida.codunida"
        miSQL = miSQL & " AND numfactu =" & miRsAux!NumFactu & " AND fecfactu = " & DBSet(miRsAux!FecFactu, "F")
        miSQL = miSQL & " ORDER BY numalbar,numlinea"
        RS.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        While Not Fin
            If RS!codartic = vParamAplic.ArtReciclado Then
                ImpTot = RS!Cantidad
                If RS.BOF Then
                    C = C & "No tiene articulo anterior"
                Else
                    
                    RS.MovePrevious
                    If RS!Cantidad <> ImpTot Then C = C & "Cantidades distintas: " & RS!numlinea & vbCrLf

                    'Aquiinserto el posteriro
                    NumRegElim = NumRegElim + 1
                    
                    
                    campo = "insert into `tmpinformes` (`codusu`,`codigo1`,campo1,nombre1,importe1,importe2,importe3) "
                    campo = campo & " VALUES (" & vUsu.Codigo & "," & NumRegElim & "," & RS!CodUnida & ","
                    campo = campo & DBSet(RS!nomUnida, "T") & "," & DBSet(ImpTot, "N") & ","
                    
                    
                    'Vuelvo a poner el registro donde toca
                    RS.MoveNext 'lo dejo en reciclado
                    
                    campo = campo & DBSet(RS!precioar, "N") & "," & DBSet(RS!ImporteL, "N") & ")"
                    EjecutaSQL conAri, campo
                    
                    
                    
                End If
                
            End If
            RS.MoveNext
            If RS.EOF Then Fin = True
        Wend
        If C <> "" Then
            C = "Fra: " & miRsAux!NumFactu & " " & miRsAux!FecFactu & vbCrLf & vbCrLf & C
            MsgBox C, vbExclamation
        End If
        
        'Siguiente factura
        RS.Close
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
 
    If NumRegElim = 0 Then
        MsgBox "Ningun dato", vbExclamation
    Else
        If HayRegParaInforme("tmpinformes", "codusu = " & vUsu.Codigo) Then GenerarDatosEncoenves = True
    End If
    Label3(72).Caption = ""
    Exit Function
EgenerarDatosEncoenves:
    MuestraError Err.Number
    Label3(72).Caption = ""
End Function





Private Function HacerListadoResumeProduccion() As Boolean

 'Temporales
    miSQL = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    miSQL = "DELETE FROM tmptraza where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    
    'variables
    miSQL = "insert into tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,"
    miSQL = miSQL & " importe1,importe2,porcen1,fecha1,fecha2)"
    miSQL = miSQL & " select " & vUsu.Codigo & ",lotetraza,prodlin.codigo,prodlin.idlin,prodlin.codartic"
    miSQL = miSQL & " ,nomartic,prodtrazlin.cantprodu can2,prodtrazlin.Cajasprod caj2,lineaprod,"
    miSQL = miSQL & " fhinicio,feccaduca from prodlin,prodtrazlin,sartic Where"
    miSQL = miSQL & " prodlin.Codigo = prodtrazlin.Codigo And"
    miSQL = miSQL & " prodlin.idlin = prodtrazlin.idlin and prodlin.codArtic = sartic.codArtic"
    
    If txtFecha(38).Text <> "" Then miSQL = miSQL & " AND fhinicio >='" & Format(txtFecha(38).Text, FormatoFecha) & " 00:00:00'"
    If txtFecha(39).Text <> "" Then miSQL = miSQL & " AND fhinicio <='" & Format(txtFecha(39).Text, FormatoFecha) & " 23:59:59'"
    
    miSQL = miSQL & " ORDER BY prodlin.codigo,prodlin.idlin ,lotetraza"
    conn.Execute miSQL


    
    'Ventas productos conjunto=1  fechafac
    miSQL = "select codigo1 from tmpinformes WHERE"  'codigo1=LOTE
    miSQL = miSQL & " codusu = " & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Devuelve = ""
    While Not miRsAux.EOF
        'Fra
'        Label3(72).Caption = miRsAux!NumFactu & " - " & miRsAux!FecFactu
'        Label3(72).Refresh
        Devuelve = Devuelve & miRsAux!Codigo1 & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Devuelve = "" Then
        MsgBox "No se ha generado ningun dato", vbExclamation
        Exit Function
    End If
    
    Do
        NumRegElim = InStr(1, Devuelve, "|")
        If NumRegElim = 0 Then
            Devuelve = ""
        Else
            campo = Mid(Devuelve, 1, NumRegElim - 1)
            Devuelve = Mid(Devuelve, NumRegElim + 1)
            
            
            'Antes los leia de prodcajas
            'ahora de prodcajasprod
            miSQL = "Select idpalet,count(*) from prodcajasprod where  lotetraza= " & campo & " GROUP BY idpalet"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            ImpTot = 0
            'tmptraza(codusu,contador,artppal,cantidad)
            miSQL = ""
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                ImpTeo = DBLet(miRsAux.Fields(1), "N")
                miSQL = miSQL & ", (" & vUsu.Codigo & "," & campo & "," & DBSet(miRsAux!IdPalet, "N", "S") & "," & DBSet(ImpTeo, "N") & ")"
                ImpTot = ImpTot + ImpTeo
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If miSQL <> "" Then
                miSQL = Mid(miSQL, 2)
                miSQL = "INSERT INTO tmptraza(codusu,contador,artppal,cantidad) VALUES " & miSQL
                conn.Execute miSQL
                
                miSQL = "UPDATE tmpinformes set importe4= " & NumRegElim & ",importe3=" & DBSet(ImpTot, "N")
                miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND codigo1=" & campo
                conn.Execute miSQL
                
            End If
                
        End If
    Loop Until Devuelve = ""
    
    
    Cadparam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
   
    
    campo = ""
    miSQL = "0"  '    MismaFecha
    If txtFecha(38).Text <> "" Or txtFecha(39).Text <> "" Then
        If txtFecha(38).Text = txtFecha(39).Text Then
            'Misma fecha
            miSQL = "1"
            campo = "Fecha: " & txtFecha(38).Text
        Else
            If txtFecha(38).Text <> "" Then campo = "Desde " & txtFecha(38).Text
            If txtFecha(39).Text <> "" Then campo = campo & "    hasta " & txtFecha(39).Text
            campo = Trim(campo)
            campo = "Fechas: " & campo
        End If
    End If
    Cadparam = Cadparam & "pDHFecha= """ & campo & """|"
    Cadparam = Cadparam & "MismaFecha= " & miSQL & """|"
    Cadparam = Cadparam & "detalla= " & Abs(Me.chkVarios(0).Value) & "|"
    
    NumParam = 4
    cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo

    cadNomRPT = "ResumenProd.rpt"
    conSubRPT = True
    LlamarImprimir2 "Resumen produccion"
    
    
End Function



Private Function HacerResumenProduccionMoixent() As Boolean

    HacerResumenProduccionMoixent = False

    'Temporales
    miSQL = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
 
    
    'variables
    miSQL = "insert into tmpinformes(codusu,campo1,codigo1,campo2,fecha1,fecha2"
    miSQL = miSQL & " ,nombre1,nombre2,nombre3,importe1,importe2)"
    miSQL = miSQL & " select " & vUsu.Codigo & ",@rownum:=@rownum+1,sordprod.codigo, codalmac,feccreacion,fecproduccion,"
    miSQL = miSQL & " sliordpr.codartic,nomartic,numlote,cantidad,cantidad*LitrosUnidad"
    miSQL = miSQL & " FROM sordprod,sliordpr,sartic,(SELECT @rownum:=0) r WHERE sordprod.codigo=sliordpr.codigo and"
    miSQL = miSQL & " sliordpr.codartic=sartic.codartic "

    If txtFecha(40).Text <> "" Then miSQL = miSQL & " AND fecproduccion >='" & Format(txtFecha(40).Text, FormatoFecha) & "'"
    If txtFecha(41).Text <> "" Then miSQL = miSQL & " AND fecproduccion <='" & Format(txtFecha(41).Text, FormatoFecha) & "'"
    
    miSQL = miSQL & "  Order by sliordpr.codigo"
    conn.Execute miSQL

    
    If Not HayRegParaInforme("tmpinformes", "codusu=" & vUsu.Codigo, False) Then
        
    Else
        HacerResumenProduccionMoixent = True
    End If


End Function

Private Sub CargaImagenAyuda(Indice As Integer, tooltip As String)
    ImgAyuda(Indice).ToolTipText = tooltip
    ImgAyuda(Indice).Picture = frmppal.ImageListMAIL.ListImages(22).Picture
End Sub



Private Sub RecalcularPrStandard()
Dim IMporteFormato As Currency


    Label3(89).Caption = "inicio proceso"
    Label3(89).Refresh
    

            
            
            
            
            
            
            
    If optRecalPrSt(0).Value Then
        Codigo = "UPDATE sartic,sarti1 set preciost = preciouc WHERE "
        Codigo = Codigo & miSQL
        conn.Execute Codigo
        
    Else
        
        Codigo = "SELECT factorconversion,sartic.preciouc,preciost,cantidad,sarti1.codartic"
        Codigo = Codigo & " FROM sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic where sarti1.codartic IN"
        Codigo = Codigo & " (Select codartic from sartic WHERE " & miSQL & ")"
        Codigo = Codigo & " ORDER BY sarti1.codartic,sarti1.numlinea"
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Codigo = ""
        While Not miRsAux.EOF
            If Codigo <> miRsAux!codartic Then
                If Codigo <> "" Then conn.Execute "UPDATE sartic set preciost = " & DBSet(Round2(ImpTeo, 3), "N") & " WHERE codartic =" & DBSet(Codigo, "T")
                
                
                
                Codigo = "codunida in (Select codunida from sartic WHERE sartic.codartic ='" & miRsAux!codartic & "') AND 1"
                Codigo = DevuelveDesdeBD(conAri, "sum(importe)", "sunilin", Codigo, "1")
                If Codigo = "" Then
                    ImpTeo = 0
                Else
                    ImpTeo = CCur(Codigo)
                End If
                
                Codigo = miRsAux!codartic
                Label3(89).Caption = Codigo
                Label3(89).Refresh
                
            End If
            
            If IsNull(miRsAux!preciost) Then
                ImpTot = 0
            Else
                ImpTot = miRsAux!preciost 'SEIMPRE con el preciost
            End If
            If miRsAux!FactorConversion < 1 Then
                'El aceite
                ImpTot = ImpTot * miRsAux!FactorConversion
                
            Else
                
                If Me.chkRecalPrSt.Value = 0 Then ImpTot = DBLet(miRsAux!PrecioUC, "N") 'ha marcado materia auxiliar desde uc

            End If
            ImpTot = miRsAux!Cantidad * ImpTot
            
            'La suma
            ImpTeo = ImpTeo + ImpTot
        
            miRsAux.MoveNext
        Wend
        
        If Codigo <> "" Then conn.Execute "UPDATE sartic set preciost = " & DBSet(Round2(ImpTeo, 3), "N") & " WHERE codartic =" & DBSet(Codigo, "T")
        
    End If
    MsgBox "Proceso finalizado", vbInformation
    
End Sub





'Imprimir etiquetas palets CAVA
Private Function ImprimirAlbaranesEntrada() As Boolean
    ImprimirAlbaranesEntrada = False
    'NO puedes tocar codigo (la variable)

    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    
    
    Set miRsAux = New ADODB.Recordset
    miSQL = RecuperaValor(Codigo, 1)
    Devuelve = RecuperaValor(Codigo, 2)
    
    
    miSQL = ""
    cadFrom = ""
    NumParam = 0
    'Eitquetas vacias
    
    
    Devuelve = "Select * from vallentradacamionlineas where entrada =" & RecuperaValor(Codigo, 1)
    miRsAux.Open Devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        miSQL = miSQL & ", (" & vUsu.Codigo & ",1," & miRsAux!NumAlbar & ","
               
               
        Devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codartic, "T")
        miSQL = miSQL & DBSet(Devuelve, "T") & ","
        
        'Palets, palots...
        
        Devuelve = ""
        If DBLet(miRsAux!codarti1, "T") <> "" Then Devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codarti1, "T")
        
        If DBLet(miRsAux!codarti2, "T") <> "" Then
            If Len(Devuelve) > 20 Then Devuelve = Mid(Devuelve, 1, 20)
            MsgBox "Lleva mas de un tipo de envase", vbExclamation
            Devuelve = Devuelve & " *+1*"
        End If
        miSQL = miSQL & DBSet(Devuelve, "T") & ","
    
        Devuelve = CStr(Val(DBLet(miRsAux!udArti1, "N") + DBLet(miRsAux!udArti2, "N") + DBLet(miRsAux!udArti3, "N") + DBLet(miRsAux!udArti4, "N")))
        miSQL = miSQL & DBSet(Devuelve, "N") & "," & DBSet(miRsAux!pesoprod, "N") & "," & DBSet(cadFrom, "F") & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    miSQL = Mid(miSQL, 2)
    miSQL = "INSERT INTO tmpinformes(codusu,campo1,codigo1,nombre1,nombre2,campo2,importe1,fecha1) VALUES " & miSQL
    conn.Execute miSQL
    ImprimirAlbaranesEntrada = True
End Function




Private Function ImprimirEtiquetasAlbaranes() As Boolean
    ImprimirEtiquetasAlbaranes = False
    'NO puedes tocar codigo (la variable)

    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    
    
    Set miRsAux = New ADODB.Recordset
    miSQL = RecuperaValor(Codigo, 1)
    Devuelve = RecuperaValor(Codigo, 2)
    
    
    miSQL = ""
    cadFrom = ""
    NumParam = 0
    'Eitquetas vacias
    If Val(Me.txtNumeroEntero(2).Text) > 0 Then
        For NumParam = 1 To Val(Me.txtNumeroEntero(2).Text)
            'tmpinformes(codusu,campo1,codigo1,nombre1,nombre2,campo2,importe1,fecha1,porcen1)"
            miSQL = miSQL & ", (" & vUsu.Codigo & ",0," & 100000 + NumParam & ",null,null,null,null,null)"
        Next
    End If
    
    Devuelve = "Select * from vallentradacamion where entrada =" & RecuperaValor(Codigo, 1)
    miRsAux.Open Devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = miRsAux!FechaEntrada
    miRsAux.Close
    
    Devuelve = "Select * from vallentradacamionlineas where entrada =" & RecuperaValor(Codigo, 1)
    miRsAux.Open Devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        miSQL = miSQL & ", (" & vUsu.Codigo & ",1," & miRsAux!NumAlbar & ","
       
        Devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codartic, "T")
        miSQL = miSQL & DBSet(Devuelve, "T") & ","
        
        'Palets, palots...
        
        Devuelve = ""
        If DBLet(miRsAux!codarti1, "T") <> "" Then Devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codarti1, "T")
        
        If DBLet(miRsAux!codarti2, "T") <> "" Then
            If Len(Devuelve) > 20 Then Devuelve = Mid(Devuelve, 1, 20)
            MsgBox "Lleva mas de un tipo de envase", vbExclamation
            Devuelve = Devuelve & " *+1*"
        End If
        miSQL = miSQL & DBSet(Devuelve, "T") & ","
    
        Devuelve = CStr(Val(DBLet(miRsAux!udArti1, "N") + DBLet(miRsAux!udArti2, "N") + DBLet(miRsAux!udArti3, "N") + DBLet(miRsAux!udArti4, "N")))
        miSQL = miSQL & DBSet(Devuelve, "N") & "," & DBSet(miRsAux!pesoprod, "N") & "," & DBSet(cadFrom, "F") & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    miSQL = Mid(miSQL, 2)
    miSQL = "INSERT INTO tmpinformes(codusu,campo1,codigo1,nombre1,nombre2,campo2,importe1,fecha1) VALUES " & miSQL
    conn.Execute miSQL
    ImprimirEtiquetasAlbaranes = True
End Function




Private Function GenerarAlbaranesOliva() As Boolean
Dim RT As ADODB.Recordset
Dim cSt As cStock
Dim J As Integer
    'cadFormula:  Lleva el codigo de la entrada de camion a traspasar. NO tocar
    'Cadparam:    Proveedor
    On Error GoTo eGenerarAlbaranesOliva
    GenerarAlbaranesOliva = False
    
    'Un par de comprobaciones
    'No existe ningun albaran con ese numero en scaalp
    miSQL = "Select * from scaalp where numalbar IN (select numalbar from vallentradacamionlineas where entrada=" & cadFormula & ")"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & "-   " & miRsAux!NumAlbar
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        miSQL = "Existen albaranes de proveedor con ese numero: " & vbCrLf & Mid(miSQL, 2)
        MsgBox miSQL, vbExclamation
        Exit Function
    End If
    
    
    
    'YA.
    'GENERAMOS DATOS
    Set cSt = New cStock
    Set RT = New ADODB.Recordset
    miSQL = "Select vallentradacamion.*,nomprove,domprove,codpobla,pobprove,pobprove,proprove,nifprove,nifprove,telprov1,codforpa"
    miSQL = miSQL & " from vallentradacamion,sprove where vallentradacamion.codprove =sprove.codprove AND entrada =" & cadFormula
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'No puede ser EOF, ni lo controlo
    miSQL = "Select vallentradacamionlineas.* ,nomartic from vallentradacamionlineas,sartic where vallentradacamionlineas.codartic=sartic.codartic "
    miSQL = miSQL & " AND entrada =" & cadFormula & " ORDER BY numalbar"
    RT.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Movimientos articulos
    cSt.tipoMov = "E"
    cSt.Trabajador = miRsAux!codProve  'En smoval guardamos el Proveedor, aunque ponga trabajadr
    cSt.DetaMov = "ALC"
    cSt.Fechamov = miRsAux!FechaEntrada
    cSt.Importe = 0
    cSt.HoraMov = miRsAux!FechaEntrada & " " & Format(Now, "hh:mm:ss")
    
    While Not RT.EOF
        'Cabecera albaran
        miSQL = "INSERT INTO scaalp(numalbar,fechaalb,codprove,nomprove,domprove,codpobla,pobprove,proprove,"
        miSQL = miSQL & "nifprove,telprove,codforpa,codtraba,codtrab1,dtoppago,dtognral,observa1) VALUES ("
        miSQL = miSQL & DBSet(RT!NumAlbar, "T") & "," & DBSet(miRsAux!FechaEntrada, "F") & "," & miRsAux!codProve & ","
        'nomprove,domprove,codpobla
        miSQL = miSQL & DBSet(miRsAux!nomprove, "T") & "," & DBSet(miRsAux!domprove, "T") & ",'" & miRsAux!codpobla & "',"
        'pobprove,proprove,nifprove
        miSQL = miSQL & DBSet(miRsAux!pobprove, "T") & "," & DBSet(miRsAux!proprove, "T") & ",'" & miRsAux!nifProve & "',"
        'telprove,codforpa,tra1 y 2
        miSQL = miSQL & DBSet(miRsAux!telprov1, "T") & "," & DBSet(miRsAux!codforpa, "T") & "," & vUsu.CodigoTrabajador & "," & vUsu.CodigoTrabajador
        'dtoppago,dtognral,observa1
        miSQL = miSQL & ",0,0,'Entrada camion: " & cadFormula & "   Generado: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & "')"
        conn.Execute miSQL
        
        
        
        'La lineas
        cSt.codAlmac = RT!codAlmac
        cSt.Cantidad = RT!Neto
        cSt.codartic = RT!codartic
        cSt.Documento = RT!NumAlbar
        cSt.LineaDocu = 1
        cSt.ActualizarStock False
        
        
        
        
        'Insertamos en slialp
        miSQL = "INSERT INTO slialp(numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,"
        miSQL = miSQL & "cantidad,precioar,dtoline1,dtoline2,importel) VALUES ("
        miSQL = miSQL & DBSet(RT!NumAlbar, "T") & "," & DBSet(miRsAux!FechaEntrada, "F") & "," & miRsAux!codProve & ",1,"
        miSQL = miSQL & DBSet(RT!codartic, "T") & "," & RT!codAlmac & "," & DBSet(RT!NomArtic, "T") & ",'Entrada: " & RT!entrada & "  " & miRsAux!matricula & "',"
        miSQL = miSQL & DBSet(RT!Neto, "N") & ",0,0,0,0)"
        conn.Execute miSQL
            
        'Si lleva palets....
        For J = 1 To 4
            
            campo = "codarti" & J
            miSQL = DBLet(RT.Fields(campo), "T")
            If miSQL <> "" Then
                Codigo = "udArti" & J
                
                cSt.codAlmac = RT!codAlmac
                cSt.Cantidad = RT.Fields(Codigo)
                cSt.codartic = RT.Fields(campo)
                cSt.Documento = RT!NumAlbar
                cadTitulo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", cSt.codartic, "T")
                cSt.LineaDocu = J + 1
                cSt.ActualizarStock False
                
            
            
            
                'Insertamos en slialp
                miSQL = "INSERT INTO slialp(numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,"
                miSQL = miSQL & "cantidad,precioar,dtoline1,dtoline2,importel) VALUES ("
                miSQL = miSQL & DBSet(RT!NumAlbar, "T") & "," & DBSet(miRsAux!FechaEntrada, "F") & "," & miRsAux!codProve & "," & J + 1 & ","
                miSQL = miSQL & DBSet(cSt.codartic, "T") & "," & RT!codAlmac & "," & DBSet(cadTitulo, "T") & ",'Entrada: " & RT!entrada & "  " & miRsAux!matricula & "',"
                miSQL = miSQL & DBSet(cSt.Cantidad, "N") & ",0,0,0,0)"
                conn.Execute miSQL
            
            End If
        Next J
            
            
        
        '---------------------------
        RT.MoveNext
    Wend
    RT.Close
        
        
    miSQL = "UPDATE vallentradacamion set EntradaFinalizada=1 WHERE entrada=" & cadFormula
    conn.Execute miSQL
    

    GenerarAlbaranesOliva = True
eGenerarAlbaranesOliva:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RT = Nothing
    Set cSt = Nothing
End Function




'****************************************************************
'****************************************************************
'****************************************************************
'
'
'
'   Declaración mensual de almazara
'
'
'
'****************************************************************
'****************************************************************
'****************************************************************

Private Function GenerarListadoAlmazara() As Boolean
Dim FechaInicioCampaña As Date

    On Error GoTo eGenerarListadoAlmazara
    GenerarListadoAlmazara = False
    
    Label3(98).Caption = "Preparando datos"
    Label3(98).Refresh
    

    If cboMes(0).ListIndex >= 8 Then
        miSQL = "01/" & Format(cboMes(0).ListIndex + 1, "00") & "/" & Format(Me.txtNumeroEntero(4).Text, "0000")
    Else
        miSQL = "01/" & Format(cboMes(0).ListIndex + 1, "00") & "/" & Format(Val(Me.txtNumeroEntero(4).Text) - 1, "0000")
    End If
    FechaInicioCampaña = CDate(miSQL)
    
    
    conn.Execute "DELETE from tmpinformes where codusu =" & vUsu.Codigo
    

    'Aqui
    'Descomentar y empezar a programas
'miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) "
'    '                                                       * factorconversion
'    miSQL = miSQL & " select " & vUsu.Codigo & ",salmac.codartic,0,sum(stockinv)  from salmac,sartic"
'    miSQL = miSQL & " where salmac.codartic=sartic.codartic and factorconversion<>1"
'    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
'    miSQL = miSQL & " group by salmac.codartic"
'    conn.Execute miSQL    'Para el punto materias primas, los vendidos directamente
'
'
'    'Cojeremos un cursor con todos las materias primas e iremos insertandolas en la tmpstock
'    '-----------------
'    miSQL = "select salmac.codartic,sum(stockinv) cantidad from salmac,sartic where salmac.codartic=sartic.codartic "
'    miSQL = miSQL & " and conjunto=1 "
'    'Las fechas
'    If Cadselect <> "" Then miSQL = miSQL & " and " & Cadselect
'    miSQL = miSQL & " group by salmac.codartic "
'
'    Set R = New ADODB.Recordset
'    Devuelve = "|"
'    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'            'Para cada elemento facturado que tiene componentes, vere de sus componentes cual es el de mataria prima y calcular su cantidad
'            miSQL = "select sarti1.codarti1,cantidad from sarti1,sartic where  sarti1.codarti1=sartic.codartic and factorconversion<>1"
'            miSQL = miSQL & " AND sarti1.codartic =" & DBSet(miRsAux!codartic, "T")
'            R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'            'If Mid(miRsAux!codArtic, 1, 9) = "002700090" Then Stop
'
'
'
'            Cantidad = DBLet(miRsAux!Cantidad, "N")
'
'
'            'Para no tener que hacer un select para saber si ya ha sido insertado en tmpstock, utilizar
'            'el string cadSelect para ir metiendo los ya insertados.
'            While Not R.EOF
'                'El articulo en cuestion
'                miSQL = "|" & R!codarti1 & "|"
'                Cantidad = Cantidad * R!Cantidad   'Esta es la cantidad nueva
'                campo = TransformaComasPuntos(CStr(Cantidad))
'                If InStr(1, Devuelve, miSQL) > 0 Then
'                    'Ya esta insertado. Es un UPDATE
'                    miSQL = "UPDATE tmpstockfec SET stock=stock + " & campo
'                    miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " and codartic = " & DBSet(R!codarti1, "T")
'                    miSQL = miSQL & " AND codalmac= 1"
'                Else
'                    Devuelve = Devuelve & R!codarti1 & "|"
'                    miSQL = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock)  VALUES (" & vUsu.Codigo & "," & DBSet(R!codarti1, "T")
'                    miSQL = miSQL & ",1," & campo & ")"
'
'                End If
'                conn.Execute miSQL
'                'No deberia haber mas (seria un coupage)
'                R.MoveNext
'            Wend
'            R.Close
'            miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
'    Set R = Nothing
'
'
'
'
'
'
'    'Existencias iniciales del mes. Es decir existencias
'
'
'    'Calculamos las existencias iniciales del mes. Es decir
'

eGenerarListadoAlmazara:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Function
