VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10845
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameListado 
      Height          =   5535
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   6555
      Begin VB.Frame FrameHomologacion 
         Height          =   1335
         Left            =   600
         TabIndex        =   601
         Top             =   3480
         Width           =   5415
         Begin VB.CheckBox chkProvHomologado 
            Caption         =   "Mostrar acciones"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   604
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox cboMultiPorposito 
            Height          =   315
            Index           =   0
            ItemData        =   "frmListado.frx":000C
            Left            =   1440
            List            =   "frmListado.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   602
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkProvHomologado 
            Caption         =   "Mostrar observaciones"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   603
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Homologados"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   605
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame frameOrdenar 
         Caption         =   "Ordenar por"
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
         TabIndex        =   150
         Top             =   2640
         Width           =   3375
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripción"
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
            Caption         =   "Código"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   1
         Top             =   2040
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1605
         TabIndex        =   0
         Top             =   1560
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   4
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   5
         Top             =   5040
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado.frx":0044
         ToolTipText     =   "Buscar marca"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":0146
         ToolTipText     =   "Buscar marca"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   720
         TabIndex        =   16
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   14
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   13
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Marcas"
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
         TabIndex        =   15
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame FrameEstMargenes 
      Height          =   5295
      Left            =   120
      TabIndex        =   400
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   126
         Left            =   4200
         TabIndex        =   409
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   125
         Left            =   1800
         TabIndex        =   408
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Frame FrameValorar2 
         Caption         =   "Valorar Con:"
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
         Height          =   1215
         Left            =   360
         TabIndex        =   420
         Top             =   3720
         Width           =   2535
         Begin VB.OptionButton optPrecioMP2 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   423
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC2 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   422
            Top             =   525
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd2 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   421
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   6240
         TabIndex        =   411
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEst 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   410
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   90
         Left            =   1800
         TabIndex        =   406
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   91
         Left            =   1800
         TabIndex        =   407
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   416
         Text            =   "Text5"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   415
         Text            =   "Text5"
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   88
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   404
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   405
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   88
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   403
         Text            =   "Text5"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   89
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   402
         Text            =   "Text5"
         Top             =   1440
         Width           =   3615
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
         Index           =   92
         Left            =   480
         TabIndex        =   600
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   3840
         Picture         =   "frmListado.frx":0248
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   599
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1560
         Picture         =   "frmListado.frx":02D3
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   3360
         TabIndex        =   598
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   419
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   418
         Top             =   2520
         Width           =   420
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
         Index           =   54
         Left            =   480
         TabIndex        =   417
         Top             =   1920
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   69
         Left            =   1515
         Picture         =   "frmListado.frx":035E
         ToolTipText     =   "Buscar artículo"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   70
         Left            =   1515
         Picture         =   "frmListado.frx":0460
         ToolTipText     =   "Buscar artículo"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   960
         TabIndex        =   414
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   960
         TabIndex        =   413
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   53
         Left            =   480
         TabIndex        =   412
         Top             =   840
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   67
         Left            =   1515
         Picture         =   "frmListado.frx":0562
         ToolTipText     =   "Buscar familia"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   68
         Left            =   1515
         Picture         =   "frmListado.frx":0664
         ToolTipText     =   "buscar familia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Informe Margenes de Venta por Artículo"
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
         TabIndex        =   401
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   6855
      Left            =   360
      TabIndex        =   256
      Top             =   0
      Width           =   7395
      Begin VB.CheckBox chkPreciosProvee 
         Caption         =   "Componentes"
         Height          =   195
         Left            =   600
         TabIndex        =   597
         ToolTipText     =   "Solo referencias que son componentes"
         Top             =   5760
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.Frame FrameSituacionArticulo 
         Caption         =   "Situación artículo"
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
         Left            =   480
         TabIndex        =   591
         Top             =   5880
         Width           =   4455
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Caducado"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   594
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Bloqueado"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   593
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   592
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkMinimoCorreg 
         Caption         =   "No mostrar tarifas por encima de margen"
         Height          =   195
         Left            =   600
         TabIndex        =   536
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Imprimir Stocks"
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
         Height          =   615
         Left            =   480
         TabIndex        =   313
         Top             =   5880
         Width           =   4335
         Begin VB.OptionButton optPuntoPedido 
            Caption         =   "Punto de pedido"
            Height          =   255
            Left            =   2520
            TabIndex        =   273
            Top             =   280
            Width           =   1575
         End
         Begin VB.OptionButton optStockMin 
            Caption         =   "Mínimos"
            Height          =   255
            Left            =   1320
            TabIndex        =   272
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optStockMax 
            Caption         =   "Máximos"
            Height          =   255
            Left            =   120
            TabIndex        =   271
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":0766
         Left            =   600
         List            =   "frmListado.frx":0773
         Style           =   2  'Dropdown List
         TabIndex        =   275
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Frame FrameTapaINCORRECTO 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   519
         Top             =   840
         Width           =   4215
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   107
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   520
            Text            =   "Text5"
            Top             =   45
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   107
            Left            =   360
            MaxLength       =   4
            TabIndex        =   259
            Top             =   45
            Width           =   615
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   87
            Left            =   80
            Picture         =   "frmListado.frx":0792
            ToolTipText     =   "Buscar almacen"
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   5760
         TabIndex        =   357
         Top             =   840
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":0894
            Style           =   1  'Graphical
            TabIndex        =   359
            Top             =   1305
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":0B9E
            Style           =   1  'Graphical
            TabIndex        =   358
            Top             =   600
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   120
            TabIndex        =   360
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2355
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
            Index           =   31
            Left            =   120
            TabIndex        =   361
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   260
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   72
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   309
         Text            =   "Text5"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   69
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   297
         Text            =   "Text5"
         Top             =   4470
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   68
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   296
         Text            =   "Text5"
         Top             =   4150
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   268
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   68
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   267
         Top             =   4155
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   292
         Text            =   "Text5"
         Top             =   2590
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   64
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   291
         Text            =   "Text5"
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   264
         Top             =   2590
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   263
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   63
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   280
         Text            =   "Text5"
         Top             =   1750
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   62
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   279
         Text            =   "Text5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   71
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   278
         Text            =   "Text5"
         Top             =   5400
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   70
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   277
         Text            =   "Text5"
         Top             =   5080
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   262
         Top             =   1750
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   261
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   71
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   270
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   269
         Top             =   5080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarArtic 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   274
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   6240
         TabIndex        =   276
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   265
         Top             =   3190
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   266
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "Text5"
         Top             =   3190
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   67
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   257
         Text            =   "Text5"
         Top             =   3510
         Width           =   4575
      End
      Begin VB.ComboBox cmbProduccion2 
         Height          =   315
         ItemData        =   "frmListado.frx":0EA8
         Left            =   2280
         List            =   "frmListado.frx":0EB2
         Style           =   2  'Dropdown List
         TabIndex        =   589
         Top             =   6120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Verificar sobre"
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
         Left            =   2280
         TabIndex        =   590
         Top             =   5880
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   75
         Left            =   600
         TabIndex        =   521
         Top             =   5880
         Width           =   870
      End
      Begin VB.Label Label4 
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
         Index           =   39
         Left            =   600
         TabIndex        =   289
         Top             =   1200
         Width           =   600
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
         Index           =   36
         Left            =   600
         TabIndex        =   312
         Top             =   890
         Width           =   735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   18
         Left            =   1515
         Picture         =   "frmListado.frx":0EE3
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   26
         Left            =   1515
         Picture         =   "frmListado.frx":0FE5
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4485
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   25
         Left            =   1515
         Picture         =   "frmListado.frx":10E7
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4155
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Articulo"
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
         Left            =   600
         TabIndex        =   300
         Top             =   3900
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   960
         TabIndex        =   299
         Top             =   4470
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   298
         Top             =   4155
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   22
         Left            =   1515
         Picture         =   "frmListado.frx":11E9
         ToolTipText     =   "Buscar marca"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   21
         Left            =   1515
         Picture         =   "frmListado.frx":12EB
         ToolTipText     =   "Buscar marca"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   35
         Left            =   600
         TabIndex        =   295
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   960
         TabIndex        =   294
         Top             =   2595
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   960
         TabIndex        =   293
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Articulos"
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
         TabIndex        =   290
         Top             =   360
         Width           =   6735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   20
         Left            =   1515
         Picture         =   "frmListado.frx":13ED
         ToolTipText     =   "Buscar familia"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   19
         Left            =   1515
         Picture         =   "frmListado.frx":14EF
         ToolTipText     =   "Buscar familia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   960
         TabIndex        =   288
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   287
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   28
         Left            =   1515
         Picture         =   "frmListado.frx":15F1
         ToolTipText     =   "Buscar artículo"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   27
         Left            =   1515
         Picture         =   "frmListado.frx":16F3
         ToolTipText     =   "Buscar artículo"
         Top             =   5085
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
         Index           =   38
         Left            =   600
         TabIndex        =   286
         Top             =   4820
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   960
         TabIndex        =   285
         Top             =   5400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   960
         TabIndex        =   284
         Top             =   5085
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   960
         TabIndex        =   283
         Top             =   3195
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   960
         TabIndex        =   282
         Top             =   3510
         Width           =   420
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
         Index           =   37
         Left            =   600
         TabIndex        =   281
         Top             =   2950
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   23
         Left            =   1515
         Picture         =   "frmListado.frx":17F5
         ToolTipText     =   "Buscar proveedor"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   24
         Left            =   1515
         Picture         =   "frmListado.frx":18F7
         ToolTipText     =   "Buscar proveedor"
         Top             =   3540
         Width           =   240
      End
   End
   Begin VB.Frame FrameInventario 
      Height          =   6495
      Left            =   240
      TabIndex        =   68
      Top             =   0
      Width           =   7995
      Begin VB.CheckBox chkStockFechaAceite 
         Caption         =   "Listado stock de aceite"
         Height          =   255
         Left            =   4200
         TabIndex        =   596
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Frame FrameOpciones 
         Height          =   1695
         Left            =   4080
         TabIndex        =   362
         Top             =   3720
         Width           =   3015
         Begin VB.CheckBox chkValorado 
            Caption         =   "Valorado"
            Height          =   255
            Left            =   240
            TabIndex        =   366
            Top             =   1320
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkImprimeStock 
            Caption         =   "Imprimir Stock"
            Height          =   255
            Left            =   240
            TabIndex        =   365
            Top             =   960
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkSinStock 
            Caption         =   "Imprimir Artículos sin Stock"
            Height          =   255
            Left            =   240
            TabIndex        =   364
            Top             =   600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox chkSaltaPag 
            Caption         =   "Salta pág. en Familia"
            Height          =   255
            Left            =   240
            TabIndex        =   363
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
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
         Height          =   1575
         Left            =   600
         TabIndex        =   90
         Top             =   3840
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   22
         Left            =   4920
         TabIndex        =   52
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   53
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   4680
         Width           =   4215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   3960
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   19
         Left            =   1920
         TabIndex        =   50
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   49
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5880
         TabIndex        =   55
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   54
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   14
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   45
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   15
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   46
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   47
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   48
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   20
         Left            =   2440
         TabIndex        =   51
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   13
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Index           =   91
         Left            =   840
         TabIndex        =   595
         Top             =   6120
         Width           =   3525
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   4670
         Picture         =   "frmListado.frx":19F9
         Top             =   4440
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
         Left            =   4200
         TabIndex        =   96
         Top             =   4440
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
         Index           =   8
         Left            =   3720
         TabIndex        =   95
         Top             =   4440
         Width           =   450
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
         Index           =   7
         Left            =   600
         TabIndex        =   89
         Top             =   4680
         Width           =   945
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   17
         Left            =   1635
         Picture         =   "frmListado.frx":1A84
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   16
         Left            =   1635
         Picture         =   "frmListado.frx":1B86
         ToolTipText     =   "Buscar provedor"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   15
         Left            =   1635
         Picture         =   "frmListado.frx":1C88
         ToolTipText     =   "Buscar proveedor"
         Top             =   3600
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
         Left            =   600
         TabIndex        =   87
         Top             =   3360
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   1080
         TabIndex        =   86
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   1080
         TabIndex        =   85
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   1080
         TabIndex        =   82
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   1080
         TabIndex        =   81
         Top             =   2040
         Width           =   420
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
         Left            =   600
         TabIndex        =   79
         Top             =   1440
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1635
         Picture         =   "frmListado.frx":1D8A
         ToolTipText     =   "Buscar artículo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1635
         Picture         =   "frmListado.frx":1E8C
         ToolTipText     =   "Buscar artículo"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   1080
         TabIndex        =   78
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   1080
         TabIndex        =   77
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   3
         Left            =   600
         TabIndex        =   76
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1635
         Picture         =   "frmListado.frx":1F8E
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1635
         Picture         =   "frmListado.frx":2090
         ToolTipText     =   "Buscar familia"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inventario"
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
         TabIndex        =   75
         Top             =   4440
         Width           =   1440
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
         Index           =   1
         Left            =   600
         TabIndex        =   74
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1635
         Picture         =   "frmListado.frx":2192
         ToolTipText     =   "Buscar almacen"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2140
         Picture         =   "frmListado.frx":2294
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe Toma de Inventario Articulos"
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
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame FrameRepxDia 
      Height          =   5415
      Left            =   480
      TabIndex        =   173
      Top             =   480
      Width           =   6075
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   600
         TabIndex        =   329
         Top             =   3480
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   400
            Left            =   120
            TabIndex        =   331
            Top             =   640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Comprobaciones:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   332
            Top             =   135
            Width           =   4455
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   330
            Top             =   375
            Width           =   4575
         End
      End
      Begin VB.Frame FrameTipMov 
         BorderStyle     =   0  'None
         Caption         =   "Nº Factura"
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
         Height          =   990
         Left            =   360
         TabIndex        =   581
         Top             =   2560
         Width           =   4815
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   122
            Left            =   3555
            TabIndex        =   172
            Top             =   440
            Width           =   1040
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   121
            Left            =   2360
            TabIndex        =   171
            Top             =   440
            Width           =   1040
         End
         Begin VB.ComboBox cboTipMov 
            Height          =   315
            ItemData        =   "frmListado.frx":231F
            Left            =   110
            List            =   "frmListado.frx":2321
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   440
            Width           =   2060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura: "
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
            TabIndex        =   585
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Tip. Mov."
            Height          =   195
            Index           =   95
            Left            =   110
            TabIndex        =   584
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   94
            Left            =   3555
            TabIndex        =   583
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   5
            Left            =   2360
            TabIndex        =   582
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdAceptarRepxDia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   174
         Top             =   3600
         Width           =   975
      End
      Begin VB.Frame FrameContab 
         Caption         =   " Facturas "
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
         Height          =   620
         Left            =   480
         TabIndex        =   328
         Top             =   1080
         Width           =   4455
         Begin VB.OptionButton OptProve 
            Caption         =   "Proveedores"
            Height          =   255
            Left            =   2280
            TabIndex        =   165
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton OptClientes 
            Caption         =   "Clientes"
            Height          =   255
            Left            =   600
            TabIndex        =   163
            Top             =   250
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   324
         Top             =   1680
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   31
            Left            =   1200
            TabIndex        =   167
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   32
            Left            =   3660
            TabIndex        =   169
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   327
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   29
            Left            =   2840
            TabIndex        =   326
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reparación:"
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
            Left            =   360
            TabIndex        =   325
            Top             =   200
            Width           =   1665
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   915
            Picture         =   "frmListado.frx":2323
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3360
            Picture         =   "frmListado.frx":23AE
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   3840
         TabIndex        =   175
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones por Día"
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
         TabIndex        =   176
         Top             =   465
         Width           =   5055
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   6000
      TabIndex        =   537
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar PBMail 
         Height          =   375
         Left            =   360
         TabIndex        =   538
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
         TabIndex        =   539
         Top             =   840
         Width           =   5805
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFrecuencia 
      Height          =   3855
      Left            =   240
      TabIndex        =   470
      Top             =   2400
      Width           =   6015
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   99
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   481
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   99
         Left            =   1320
         TabIndex        =   480
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   101
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   479
         Text            =   "Text5"
         Top             =   2400
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   100
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   478
         Text            =   "Text5"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   101
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   477
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   100
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   476
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   98
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   475
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1320
         TabIndex        =   474
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdFrecuencias 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   473
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   96
         Left            =   4800
         TabIndex        =   472
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   77
         Left            =   1035
         Picture         =   "frmListado.frx":2439
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   480
         TabIndex        =   487
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   79
         Left            =   1035
         Picture         =   "frmListado.frx":253B
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   78
         Left            =   1035
         Picture         =   "frmListado.frx":263D
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Index           =   67
         Left            =   120
         TabIndex        =   486
         Top             =   1800
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   480
         TabIndex        =   485
         Top             =   2400
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   480
         TabIndex        =   484
         Top             =   2040
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   76
         Left            =   1035
         Picture         =   "frmListado.frx":273F
         ToolTipText     =   "Buscar cliente"
         Top             =   1080
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
         Index           =   66
         Left            =   120
         TabIndex        =   483
         Top             =   720
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
         Index           =   65
         Left            =   480
         TabIndex        =   482
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Datos de frecuencias  clientes"
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
         TabIndex        =   471
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FrameMovArtic 
      Height          =   5535
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   10635
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   1485
         TabIndex        =   25
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   1485
         TabIndex        =   26
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeselTodos 
         Height          =   435
         Left            =   9000
         Picture         =   "frmListado.frx":2841
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   740
         Width           =   585
      End
      Begin VB.CommandButton cmdSelTodos 
         Height          =   435
         Left            =   9720
         Picture         =   "frmListado.frx":2F2B
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   740
         Width           =   585
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   6960
         TabIndex        =   27
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   24
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   23
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   3600
         TabIndex        =   22
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   65
         Left            =   1200
         Picture         =   "frmListado.frx":3615
         ToolTipText     =   "Cliente"
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   66
         Left            =   1200
         Picture         =   "frmListado.frx":3717
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   600
         TabIndex        =   398
         Top             =   4560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   397
         Top             =   4920
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente/Proveedor"
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
         Left            =   360
         TabIndex        =   396
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Movimiento"
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
         Left            =   6960
         TabIndex        =   61
         Top             =   960
         Width           =   1755
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3315
         Picture         =   "frmListado.frx":3819
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListado.frx":38A4
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   34
         Left            =   1155
         Picture         =   "frmListado.frx":392F
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   33
         Left            =   1155
         Picture         =   "frmListado.frx":3A31
         ToolTipText     =   "Buscar almacen"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   11
         Left            =   360
         TabIndex        =   60
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   59
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   58
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
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
         Index           =   10
         Left            =   360
         TabIndex        =   57
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   56
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   43
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   32
         Left            =   1155
         Picture         =   "frmListado.frx":3B33
         ToolTipText     =   "Buscar familia"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   31
         Left            =   1155
         Picture         =   "frmListado.frx":3C35
         ToolTipText     =   "Buscar familia"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   9
         Left            =   360
         TabIndex        =   42
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   41
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   40
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   30
         Left            =   1155
         Picture         =   "frmListado.frx":3D37
         ToolTipText     =   "Buscar artículo"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   29
         Left            =   1155
         Picture         =   "frmListado.frx":3E39
         ToolTipText     =   "Buscar artículo"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Movimiento Artículos"
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
         TabIndex        =   38
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   600
         TabIndex        =   37
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   36
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Frame FrameRepxClien 
      Height          =   5415
      Left            =   240
      TabIndex        =   179
      Top             =   240
      Width           =   6795
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   3720
         TabIndex        =   321
         Top             =   3240
         Width           =   2415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   186
            Text            =   "1"
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "reparaciones"
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
            Left            =   1200
            TabIndex        =   323
            Top             =   420
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar equipos con más de:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   322
            Top             =   120
            Width           =   2070
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   34
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   192
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         TabIndex        =   181
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   35
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   190
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   183
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   182
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarRepxClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   187
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   4560
         TabIndex        =   188
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1920
         TabIndex        =   184
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1920
         TabIndex        =   185
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   33
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         TabIndex        =   180
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1640
         Picture         =   "frmListado.frx":3F3B
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1640
         Picture         =   "frmListado.frx":3FC6
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   7
         Left            =   1635
         Picture         =   "frmListado.frx":4051
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1080
         TabIndex        =   202
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   9
         Left            =   1635
         Picture         =   "frmListado.frx":4153
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   8
         Left            =   1635
         Picture         =   "frmListado.frx":4255
         ToolTipText     =   "buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         TabIndex        =   201
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   1080
         TabIndex        =   200
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   1080
         TabIndex        =   199
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones  por Cliente"
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
         TabIndex        =   198
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   1080
         TabIndex        =   197
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   1080
         TabIndex        =   196
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Rep."
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
         TabIndex        =   195
         Top             =   3120
         Width           =   915
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   6
         Left            =   1635
         Picture         =   "frmListado.frx":4357
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
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
         Index           =   18
         Left            =   600
         TabIndex        =   194
         Top             =   1080
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
         Index           =   16
         Left            =   1080
         TabIndex        =   193
         Top             =   1680
         Width           =   420
      End
   End
   Begin VB.Frame FrameRepNSerie 
      Height          =   5415
      Left            =   360
      TabIndex        =   151
      Top             =   0
      Width           =   6795
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         TabIndex        =   140
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1920
         TabIndex        =   145
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1920
         TabIndex        =   144
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4560
         TabIndex        =   147
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   146
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   142
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   143
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   40
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   153
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         TabIndex        =   141
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   152
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
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
         Left            =   1080
         TabIndex        =   168
         Top             =   1680
         Width           =   420
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
         Index           =   19
         Left            =   600
         TabIndex        =   166
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   49
         Left            =   1635
         Picture         =   "frmListado.frx":4459
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   54
         Left            =   1635
         Picture         =   "frmListado.frx":455B
         ToolTipText     =   "Buscar contrato"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   53
         Left            =   1635
         Picture         =   "frmListado.frx":465D
         ToolTipText     =   "Buscar  contrato"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Contrato"
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
         Left            =   600
         TabIndex        =   164
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   1080
         TabIndex        =   162
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   1080
         TabIndex        =   161
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Informe Nº Serie"
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
         TabIndex        =   160
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   159
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   27
         Left            =   1080
         TabIndex        =   158
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Left            =   600
         TabIndex        =   157
         Top             =   2040
         Width           =   930
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   51
         Left            =   1635
         Picture         =   "frmListado.frx":475F
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   52
         Left            =   1635
         Picture         =   "frmListado.frx":4861
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   1080
         TabIndex        =   156
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   50
         Left            =   1635
         Picture         =   "frmListado.frx":4963
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FrameDtosFM 
      Height          =   5415
      Left            =   480
      TabIndex        =   314
      Top             =   600
      Width           =   6915
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   351
         Top             =   840
         Width           =   6135
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   74
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   353
            Text            =   "Text5"
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   74
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   304
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   73
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   352
            Text            =   "Text5"
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   73
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   303
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   1
            Left            =   1275
            Picture         =   "frmListado.frx":4A65
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   61
            Left            =   720
            TabIndex        =   356
            Top             =   360
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   0
            Left            =   1275
            Picture         =   "frmListado.frx":4B67
            ToolTipText     =   "Buscar cliente"
            Top             =   360
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
            Index           =   44
            Left            =   240
            TabIndex        =   355
            Top             =   120
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
            Index           =   45
            Left            =   720
            TabIndex        =   354
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         TabIndex        =   345
         Top             =   2880
         Width           =   6135
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   307
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   308
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   77
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   347
            Text            =   "Text5"
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   78
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   346
            Text            =   "Text5"
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   66
            Left            =   720
            TabIndex        =   350
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   67
            Left            =   720
            TabIndex        =   349
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label4 
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
            Index           =   42
            Left            =   240
            TabIndex        =   348
            Top             =   120
            Width           =   525
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   4
            Left            =   1275
            Picture         =   "frmListado.frx":4C69
            ToolTipText     =   "Buscar marca"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   5
            Left            =   1275
            Picture         =   "frmListado.frx":4D6B
            ToolTipText     =   "Buscar marca"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   339
         Top             =   3720
         Width           =   6255
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   79
            Left            =   1560
            TabIndex        =   301
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   80
            Left            =   1560
            TabIndex        =   302
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   79
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   341
            Text            =   "Text5"
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   80
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   340
            Text            =   "Text5"
            Top             =   720
            Width           =   3615
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
            Index           =   46
            Left            =   240
            TabIndex        =   342
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   65
            Left            =   720
            TabIndex        =   344
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   64
            Left            =   720
            TabIndex        =   343
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   63
            Left            =   1275
            Picture         =   "frmListado.frx":4E6D
            ToolTipText     =   "Buscar proveedor"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   64
            Left            =   1275
            Picture         =   "frmListado.frx":4F6F
            ToolTipText     =   "Buscar proveedor"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5040
         TabIndex        =   311
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarDtosFM 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   310
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   305
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   306
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   75
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   316
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   76
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   315
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Listado Descuentos Familia/Marca"
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
         TabIndex        =   320
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   1080
         TabIndex        =   319
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   62
         Left            =   1080
         TabIndex        =   318
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   40
         Left            =   600
         TabIndex        =   317
         Top             =   2040
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   2
         Left            =   1635
         Picture         =   "frmListado.frx":5071
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   3
         Left            =   1635
         Picture         =   "frmListado.frx":5173
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
   End
   Begin VB.Frame FrameHcoMante 
      Height          =   3495
      Left            =   0
      TabIndex        =   541
      Top             =   -120
      Width           =   6495
      Begin VB.CommandButton cmdHcoMante 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   546
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   112
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   545
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   112
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   551
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   544
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   111
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1680
         TabIndex        =   543
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   99
         Left            =   5160
         TabIndex        =   548
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo baja"
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
         Left            =   240
         TabIndex        =   552
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   90
         Left            =   1395
         Picture         =   "frmListado.frx":5275
         ToolTipText     =   "Buscar motivo baja"
         Top             =   2280
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
         Index           =   80
         Left            =   240
         TabIndex        =   550
         Top             =   1560
         Width           =   945
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   89
         Left            =   1395
         Picture         =   "frmListado.frx":5377
         ToolTipText     =   "Buscar trabajador"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1395
         Picture         =   "frmListado.frx":5479
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha baja"
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
         Index           =   79
         Left            =   240
         TabIndex        =   547
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Paso a mantenimientos anulados"
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
         Left            =   240
         TabIndex        =   542
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame FrEliminarFacturas 
      Height          =   4215
      Left            =   120
      TabIndex        =   492
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdElimiaFacturas 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3840
         TabIndex        =   496
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cmbEliFac 
         Height          =   315
         ItemData        =   "frmListado.frx":5504
         Left            =   3360
         List            =   "frmListado.frx":5506
         Style           =   2  'Dropdown List
         TabIndex        =   495
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   97
         Left            =   5040
         TabIndex        =   493
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "lore ipsum lorem ipsum lorem ipsum"
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   518
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "lore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   517
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
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
         Height          =   315
         Index           =   83
         Left            =   120
         TabIndex        =   497
         Top             =   3600
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Eliminar facturas hasta: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   82
         Left            =   360
         TabIndex        =   494
         Top             =   3000
         Width           =   2370
      End
   End
   Begin VB.Frame FrameAlbaranesMarcaFacturar 
      Height          =   3735
      Left            =   0
      TabIndex        =   562
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   578
         Top             =   1680
         Width           =   6135
      End
      Begin VB.CommandButton cmdFactAlbaranes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   568
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   82
         Left            =   5160
         TabIndex        =   569
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   3960
         TabIndex        =   565
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1680
         TabIndex        =   564
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1680
         TabIndex        =   567
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   118
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   571
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1680
         TabIndex        =   566
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   117
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   570
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3600
         Picture         =   "frmListado.frx":5508
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
         Index           =   87
         Left            =   3000
         TabIndex        =   577
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   86
         Left            =   240
         TabIndex        =   576
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   93
         Left            =   720
         TabIndex        =   575
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1320
         Picture         =   "frmListado.frx":5593
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
         Index           =   85
         Left            =   840
         TabIndex        =   574
         Top             =   2520
         Width           =   420
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
         Left            =   360
         TabIndex        =   573
         Top             =   1920
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   94
         Left            =   1395
         Picture         =   "frmListado.frx":561E
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   92
         Left            =   840
         TabIndex        =   572
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   93
         Left            =   1395
         Picture         =   "frmListado.frx":5720
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Marcar facturar albaranes"
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
         Left            =   360
         TabIndex        =   563
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame FrameBultos 
      Height          =   6975
      Left            =   240
      TabIndex        =   452
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   6
         Left            =   1320
         TabIndex        =   460
         Text            =   "Text1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   5
         Left            =   2280
         TabIndex        =   459
         Text            =   "Text1"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   4
         Left            =   1320
         TabIndex        =   458
         Text            =   "Text1"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   3
         Left            =   1320
         TabIndex        =   457
         Text            =   "Text1"
         Top             =   2640
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   2
         Left            =   1320
         TabIndex        =   456
         Text            =   "Text1"
         Top             =   2160
         Width           =   5175
      End
      Begin VB.ComboBox cmbBulto 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   455
         Top             =   1620
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   462
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdEtiqBulto 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4200
         TabIndex        =   463
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   95
         Left            =   5400
         TabIndex        =   464
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtBultos 
         Height          =   1695
         Index           =   0
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   461
         Text            =   "frmListado.frx":5822
         Top             =   4200
         Width           =   5175
      End
      Begin VB.TextBox txtClie 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   454
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   465
         Text            =   "Text5"
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
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
         Index           =   71
         Left            =   240
         TabIndex        =   491
         Top             =   3663
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Población"
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
         Index           =   70
         Left            =   240
         TabIndex        =   490
         Top             =   2703
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Index           =   69
         Left            =   240
         TabIndex        =   489
         Top             =   3183
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Index           =   68
         Left            =   240
         TabIndex        =   488
         Top             =   2223
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Copias"
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
         Index           =   64
         Left            =   240
         TabIndex        =   469
         Top             =   6480
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Texto"
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
         Left            =   240
         TabIndex        =   468
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Index           =   62
         Left            =   240
         TabIndex        =   467
         Top             =   1680
         Width           =   780
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
         Index           =   61
         Left            =   240
         TabIndex        =   466
         Top             =   840
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   75
         Left            =   960
         Picture         =   "frmListado.frx":5828
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Etiquetas de bultos"
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
         TabIndex        =   453
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame FrameTarifas 
      Height          =   6375
      Left            =   480
      TabIndex        =   97
      Top             =   120
      Width           =   7635
      Begin VB.ComboBox cboDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":592A
         Left            =   3480
         List            =   "frmListado.frx":593D
         Style           =   2  'Dropdown List
         TabIndex        =   580
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CheckBox chkMostrarErrores 
         Caption         =   "Mostrar solo tarifas con error"
         Height          =   255
         Left            =   960
         TabIndex        =   399
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   99
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkSaltaPagTarif 
         Caption         =   "Salta pág. en Familia"
         Height          =   255
         Left            =   960
         TabIndex        =   115
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   26
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   25
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   4320
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   101
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   100
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   30
         Left            =   1920
         TabIndex        =   105
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   29
         Left            =   1920
         TabIndex        =   104
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarTarif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5280
         TabIndex        =   106
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   107
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1920
         TabIndex        =   102
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1920
         TabIndex        =   103
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   98
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   88
         Left            =   3480
         TabIndex        =   579
         Top             =   5160
         Width           =   870
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   56
         Left            =   1635
         Picture         =   "frmListado.frx":595E
         ToolTipText     =   "Buscar tarifa"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   1080
         TabIndex        =   148
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   58
         Left            =   1635
         Picture         =   "frmListado.frx":5A60
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   57
         Left            =   1635
         Picture         =   "frmListado.frx":5B62
         ToolTipText     =   "Buscar familia"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   15
         Left            =   600
         TabIndex        =   139
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   1080
         TabIndex        =   138
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   137
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   62
         Left            =   1635
         Picture         =   "frmListado.frx":5C64
         ToolTipText     =   "Buscar artículo"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   61
         Left            =   1635
         Picture         =   "frmListado.frx":5D66
         ToolTipText     =   "Buscar artículo"
         Top             =   4320
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
         Index           =   14
         Left            =   600
         TabIndex        =   136
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label lblTituloTarif 
         Caption         =   "Informe Precios y Descuentos"
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
         TabIndex        =   135
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   134
         Top             =   4680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   133
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   132
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   131
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   600
         TabIndex        =   130
         Top             =   3000
         Width           =   525
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   59
         Left            =   1635
         Picture         =   "frmListado.frx":5E68
         ToolTipText     =   "Buscar marca"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   60
         Left            =   1635
         Picture         =   "frmListado.frx":5F6A
         ToolTipText     =   "Buscar marca"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
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
         Left            =   600
         TabIndex        =   118
         Top             =   5160
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   55
         Left            =   1635
         Picture         =   "frmListado.frx":606C
         ToolTipText     =   "Buscar tarifa"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   11
         Left            =   600
         TabIndex        =   117
         Top             =   960
         Width           =   495
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
         Left            =   1080
         TabIndex        =   116
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame FrameMantenimientos 
      Height          =   6975
      Left            =   360
      TabIndex        =   203
      Top             =   0
      Width           =   6735
      Begin VB.Frame FrameManteAnu 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   553
         Top             =   4800
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   116
            Left            =   5040
            TabIndex        =   217
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   115
            Left            =   2400
            TabIndex        =   216
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   114
            Left            =   1800
            TabIndex        =   215
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   114
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   557
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   113
            Left            =   1800
            TabIndex        =   214
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   113
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   554
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   91
            Left            =   4200
            TabIndex        =   561
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   90
            Left            =   1560
            TabIndex        =   560
            Top             =   1080
            Width           =   465
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   4680
            Picture         =   "frmListado.frx":616E
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   2160
            Picture         =   "frmListado.frx":61F9
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha baja"
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
            Index           =   83
            Left            =   480
            TabIndex        =   559
            Top             =   1080
            Width           =   915
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   92
            Left            =   1515
            Picture         =   "frmListado.frx":6284
            ToolTipText     =   "Buscar motivo baja"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   89
            Left            =   960
            TabIndex        =   558
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   91
            Left            =   1515
            Picture         =   "frmListado.frx":6386
            ToolTipText     =   "Buscar motivo baja"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Motivo baja"
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
            Index           =   82
            Left            =   520
            TabIndex        =   556
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   88
            Left            =   960
            TabIndex        =   555
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   360
         TabIndex        =   531
         Top             =   840
         Width           =   6255
         Begin VB.CheckBox chkMante 
            Caption         =   "Copia remitente"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   540
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   109
            Left            =   1440
            TabIndex        =   207
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Comercial"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   534
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Administracion"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   533
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkMante 
            Caption         =   "Enviar e-mail"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   532
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha carta"
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
            Index           =   77
            Left            =   120
            TabIndex        =   535
            Top             =   720
            Width           =   990
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   109
            Left            =   1155
            Picture         =   "frmListado.frx":6488
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   480
         TabIndex        =   528
         Top             =   5040
         Width           =   5895
         Begin VB.ComboBox cboTipoList 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   529
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Listado"
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
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   530
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1920
         TabIndex        =   205
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1920
         TabIndex        =   206
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   235
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   51
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   234
         Text            =   "Text5"
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   52
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   233
         Text            =   "Text5"
         Top             =   4440
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   48
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   222
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1920
         TabIndex        =   209
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   50
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   221
         Text            =   "Text5"
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   49
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   220
         Text            =   "Text5"
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   211
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   210
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarMante 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   218
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5280
         TabIndex        =   219
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1920
         TabIndex        =   212
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1920
         TabIndex        =   213
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   204
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1920
         TabIndex        =   208
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   600
         TabIndex        =   333
         Top             =   4920
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1560
            TabIndex        =   336
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   3840
            TabIndex        =   335
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   3555
            Picture         =   "frmListado.frx":6513
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   8
            Left            =   1275
            Picture         =   "frmListado.frx":659E
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   44
            Left            =   720
            TabIndex        =   338
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   45
            Left            =   3000
            TabIndex        =   337
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Revisiones Efectuadas"
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   334
            Top             =   120
            Width           =   4335
         End
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
         Left            =   1080
         TabIndex        =   239
         Top             =   1560
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
         Index           =   27
         Left            =   600
         TabIndex        =   238
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   41
         Left            =   1635
         Picture         =   "frmListado.frx":6629
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1080
         TabIndex        =   237
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   42
         Left            =   1635
         Picture         =   "frmListado.frx":672B
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   44
         Left            =   1635
         Picture         =   "frmListado.frx":682D
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   1080
         TabIndex        =   232
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   46
         Left            =   1635
         Picture         =   "frmListado.frx":692F
         ToolTipText     =   "Buscar agente"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   45
         Left            =   1635
         Picture         =   "frmListado.frx":6A31
         ToolTipText     =   "Buscar agente"
         Top             =   3120
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
         Index           =   26
         Left            =   600
         TabIndex        =   231
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   41
         Left            =   1080
         TabIndex        =   230
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1080
         TabIndex        =   229
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Mantenimientos"
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
         TabIndex        =   228
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   1080
         TabIndex        =   227
         Top             =   4080
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   1080
         TabIndex        =   226
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   25
         Left            =   600
         TabIndex        =   225
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   47
         Left            =   1635
         Picture         =   "frmListado.frx":6B33
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   48
         Left            =   1635
         Picture         =   "frmListado.frx":6C35
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   43
         Left            =   1635
         Picture         =   "frmListado.frx":6D37
         ToolTipText     =   "Buscar cliente"
         Top             =   2160
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
         Index           =   24
         Left            =   600
         TabIndex        =   224
         Top             =   1920
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
         Index           =   23
         Left            =   1080
         TabIndex        =   223
         Top             =   2520
         Width           =   420
      End
   End
   Begin VB.Frame FrameListMant2 
      Height          =   4215
      Left            =   1080
      TabIndex        =   498
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir artículos"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   516
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton cmdManteTeorico 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   515
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   77
         Left            =   5040
         TabIndex        =   514
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1680
         TabIndex        =   511
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   105
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   510
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1680
         TabIndex        =   508
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   104
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   507
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1680
         TabIndex        =   504
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   103
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   503
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1680
         TabIndex        =   501
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   102
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   500
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label4 
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
         Index           =   73
         Left            =   240
         TabIndex        =   513
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   83
         Left            =   1395
         Picture         =   "frmListado.frx":6E39
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   87
         Left            =   840
         TabIndex        =   512
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   82
         Left            =   1395
         Picture         =   "frmListado.frx":6F3B
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   840
         TabIndex        =   509
         Top             =   2280
         Width           =   465
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
         Index           =   72
         Left            =   240
         TabIndex        =   506
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   81
         Left            =   1395
         Picture         =   "frmListado.frx":703D
         ToolTipText     =   "Buscar cliente"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   85
         Left            =   840
         TabIndex        =   505
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   80
         Left            =   1395
         Picture         =   "frmListado.frx":713F
         ToolTipText     =   "Buscar cliente"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   84
         Left            =   840
         TabIndex        =   502
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Informe teórico de mantenimientos"
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
         Left            =   240
         TabIndex        =   499
         Top             =   480
         Width           =   5100
      End
   End
   Begin VB.Frame FrameFichasMan 
      Height          =   5295
      Left            =   0
      TabIndex        =   240
      Top             =   0
      Width           =   7395
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir artículos"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   527
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   124
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   108
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   525
         Text            =   "Text5"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   123
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   106
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   522
         Text            =   "Text5"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   2520
         TabIndex        =   125
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   2520
         TabIndex        =   126
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   2520
         TabIndex        =   119
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   55
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   244
         Text            =   "Text5"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   2520
         TabIndex        =   122
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   2520
         TabIndex        =   121
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   6120
         TabIndex        =   129
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarFichas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   128
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   2520
         TabIndex        =   120
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   56
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   243
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   242
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   57
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   241
         Text            =   "Text5"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   5880
         TabIndex        =   127
         Top             =   3840
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
         Index           =   78
         Left            =   1680
         TabIndex        =   526
         Top             =   3240
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   88
         Left            =   2280
         Picture         =   "frmListado.frx":7241
         ToolTipText     =   "Buscar ruta"
         Top             =   3240
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
         Index           =   76
         Left            =   1680
         TabIndex        =   524
         Top             =   2925
         Width           =   450
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
         Index           =   74
         Left            =   360
         TabIndex        =   523
         Top             =   2880
         Width           =   405
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   86
         Left            =   2280
         Picture         =   "frmListado.frx":7343
         ToolTipText     =   "Buscar ruta"
         Top             =   2902
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   39
         Left            =   2280
         Picture         =   "frmListado.frx":7445
         ToolTipText     =   "Buscar contrato"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   1680
         TabIndex        =   255
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   1680
         TabIndex        =   254
         Top             =   4200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Contrato"
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
         Left            =   360
         TabIndex        =   253
         Top             =   3720
         Width           =   990
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   40
         Left            =   2280
         Picture         =   "frmListado.frx":7547
         ToolTipText     =   "Buscar contrato"
         Top             =   4200
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
         Index           =   34
         Left            =   1680
         TabIndex        =   252
         Top             =   1320
         Width           =   420
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
         Left            =   240
         TabIndex        =   251
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   35
         Left            =   2160
         Picture         =   "frmListado.frx":7649
         ToolTipText     =   "Buscar cliente"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   38
         Left            =   2235
         Picture         =   "frmListado.frx":774B
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   37
         Left            =   2235
         Picture         =   "frmListado.frx":784D
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   32
         Left            =   240
         TabIndex        =   250
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   1680
         TabIndex        =   249
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   1680
         TabIndex        =   248
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Informe Fichas de Mantenimientos"
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
         TabIndex        =   247
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   1680
         TabIndex        =   246
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   36
         Left            =   2160
         Picture         =   "frmListado.frx":794F
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ejercicio"
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
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   245
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Frame FrameListAvisosPtes 
      Height          =   4815
      Left            =   0
      TabIndex        =   385
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   380
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   97
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   450
         Text            =   "Text5"
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   379
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   96
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   447
         Text            =   "Text5"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox cboSituaAviso 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   381
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   82
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   375
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAviPtes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   382
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   4800
         TabIndex        =   383
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   83
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   376
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   84
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   387
         Text            =   "Text5"
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   377
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   85
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   386
         Text            =   "Text5"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   378
         Top             =   2280
         Width           =   615
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
         Index           =   60
         Left            =   960
         TabIndex        =   451
         Top             =   3480
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   97
         Left            =   1440
         Picture         =   "frmListado.frx":7A51
         ToolTipText     =   "Buscar tecnico"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   59
         Left            =   600
         TabIndex        =   449
         Top             =   2880
         Width           =   645
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
         Index           =   58
         Left            =   960
         TabIndex        =   448
         Top             =   3120
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   96
         Left            =   1440
         Picture         =   "frmListado.frx":7B53
         ToolTipText     =   "Buscar tecnico"
         Top             =   3120
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
         Index           =   52
         Left            =   600
         TabIndex        =   395
         Top             =   4200
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
         Index           =   51
         Left            =   3480
         TabIndex        =   394
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListado.frx":7C55
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Avisos de avería pendientes"
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
         Left            =   1080
         TabIndex        =   393
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha aviso"
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
         Index           =   50
         Left            =   600
         TabIndex        =   392
         Top             =   840
         Width           =   990
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
         Index           =   49
         Left            =   960
         TabIndex        =   391
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3960
         Picture         =   "frmListado.frx":7CE0
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   84
         Left            =   1440
         Picture         =   "frmListado.frx":7D6B
         ToolTipText     =   "Buscar ruta"
         Top             =   1920
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
         Index           =   48
         Left            =   600
         TabIndex        =   390
         Top             =   1680
         Width           =   405
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
         Index           =   47
         Left            =   960
         TabIndex        =   389
         Top             =   1920
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   85
         Left            =   1440
         Picture         =   "frmListado.frx":7E6D
         ToolTipText     =   "Buscar ruta"
         Top             =   2280
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
         Index           =   6
         Left            =   960
         TabIndex        =   388
         Top             =   2280
         Width           =   420
      End
   End
   Begin VB.Frame FrameEtiqEstanteria 
      Height          =   4935
      Left            =   0
      TabIndex        =   424
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   124
         Left            =   4140
         TabIndex        =   430
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   123
         Left            =   1800
         TabIndex        =   429
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox cboDecimal 
         Height          =   315
         ItemData        =   "frmListado.frx":7F6F
         Left            =   1440
         List            =   "frmListado.frx":7F82
         Style           =   2  'Dropdown List
         TabIndex        =   446
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkImprimeCodigoBarras 
         Caption         =   "Impime codigo barras"
         Height          =   255
         Left            =   2520
         TabIndex        =   431
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   95
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   437
         Text            =   "Text5"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   94
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   436
         Text            =   "Text5"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   95
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   426
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   94
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   425
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   93
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   435
         Text            =   "Text5"
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   92
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   433
         Text            =   "Text5"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   93
         Left            =   1800
         TabIndex        =   428
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   92
         Left            =   1800
         TabIndex        =   427
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdEtiqEstanteria 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   432
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   94
         Left            =   6240
         TabIndex        =   434
         Top             =   4080
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   3840
         Picture         =   "frmListado.frx":7F95
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1515
         Picture         =   "frmListado.frx":8020
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   96
         Left            =   3315
         TabIndex        =   588
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   587
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha ult. cambio precio P.V.P."
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
         Index           =   89
         Left            =   480
         TabIndex        =   586
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         TabIndex        =   445
         Top             =   4080
         Width           =   870
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Etiquetas estanterias"
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
         Left            =   480
         TabIndex        =   444
         Top             =   360
         Width           =   5895
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   74
         Left            =   1515
         Picture         =   "frmListado.frx":80AB
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   73
         Left            =   1515
         Picture         =   "frmListado.frx":81AD
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   56
         Left            =   480
         TabIndex        =   443
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   960
         TabIndex        =   442
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   960
         TabIndex        =   441
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   72
         Left            =   1515
         Picture         =   "frmListado.frx":82AF
         ToolTipText     =   "Buscar artículo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   71
         Left            =   1515
         Picture         =   "frmListado.frx":83B1
         ToolTipText     =   "Buscar artículo"
         Top             =   2400
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
         Index           =   55
         Left            =   480
         TabIndex        =   440
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   76
         Left            =   960
         TabIndex        =   439
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   438
         Top             =   2400
         Width           =   465
      End
   End
   Begin VB.Frame FrameRepSustNSerie 
      Height          =   3735
      Left            =   240
      TabIndex        =   367
      Top             =   0
      Width           =   5715
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   81
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   368
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdAceptarSustNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   369
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   3120
         TabIndex        =   370
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   384
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   374
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Introduce el nuevo Nº de Serie que va a sustituir al: "
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
         Left            =   360
         TabIndex        =   373
         Top             =   1000
         Width           =   3780
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Sustitución Nº de Serie"
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
         Left            =   360
         TabIndex        =   372
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   371
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.Frame FrameInfAlmacen 
      Height          =   3495
      Left            =   1560
      TabIndex        =   30
      Top             =   1080
      Width           =   5835
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   33
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Almacenes"
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
         TabIndex        =   32
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Traspaso"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   920
         Picture         =   "frmListado.frx":84B3
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3200
         Picture         =   "frmListado.frx":85B5
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionListado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Artículos
    ' 5 .- Listado de Familias de artículos
    
    ' 6 .- Listado de Artículos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    '11 .-
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoración Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecharEtiqBulto.rpt
    
    '110 .- Listado de Ubicaciones
    
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Envío
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoración de Ofertas
    '34 .- Informe de Ofertas Efectuadas
    '35 .- Informe Historico de Ofertas
    
    '36 .- Traspaso de Ofertas al Historico (NO IMPRIME INFORME)
    '37 .- Solicitar datos para pasar de Oferta a Pedido (NO IMPRIME INFORME)
    '38 .- Informe de Pedidos
    '239 .- Hco de Pedidos de venta (Historico)
    '39 .- Orden de Instalacion
    '40 .- Cartas Confirmacion de Pedidos
    
    '41 .- Informe de Pedidos por Articulo
    '42 .- Informe de Disponibilidad de Stocks
    '43 .- Generar Albaran desde Pedido (NO IMPRIME LISTADO)
    '44 .- Informe de Pedidos por Cliente
    '45 .- Informe de Albaran
    
    '46 .- Informe de Clientes Inactivos
    '47 .- Informe de Clientes
    '48 .- Informe de Altas de Nuevos Cliente
    '49 .- Informe de Albaranes por Articulo
    '50 .- Prevision de Facturacion de ALbaranes
    
    '51 .- Informe Incumplimiento Plazos de Entrega
    '52 .- Facturacion de Albaranes (NO IMPRIME LISTADO?)
    '53 .- Informe de Factura
    '54 .- Listado de Descuentos Familia/Marca
    
    '59 .- Informe de Factura ProForma
    '222 .- Informe de Factura Mostrador
    '223 .- Pedir datos para contabilizar facturas CLIENTES
    '224 .- Pedir datos para contabilizar facturas PROVEEDOR
    '225 .- Pedir datos para generar Facturas Rectificativas
    '226 .- Pedir datos para reimprimir Facturas
    '227 .- Informe estadistica Ventas por cliente
    '228 .- Informe estadistica Ventas por Trabajador
    '229 .- Informe estadistica Ventas por meses
    '230 .- Informe estadistica Ventas por familia
    '231 .- Informe detalle facturacion clientes
    
    '240 .- Informe Cierre de Caja del TPV
    
    '245 .- Informe control margenes tarifas
    '246 .- Informe Margen ventas por articulo
    '247 .- Corrección de errores y acutalizacion de tarifas
    
    
    'Abril 2008
    '248 .- Contabilizar facturas de tickets AGRUPADAS
    
    
    
    '==== Listados de COMPRAS ====
    '=============================
    '55 .- Informe de Pedido Proveedor
    '56 .- Inf. Historico Pedido Proveedor
    '57 .- Pasa Pedido a Albaran compras (NO IMPRIME LISTADO)
    '58 .- Listado de Proveedores
    
    
    '305 .- Listado Etiquetas de Proveedores
    '306 .- Listado Cartas a Proveedores
    '307 .- Listado Material pendiente de recibir
    '308 .- Listado Albaranes pendientes de facturar
    '309 .- Listado  Precios de Compra
    '310 .- Listado Compras por Proveedor
    '311 .- Listado Compras por Familia
    '312 .- Listado albaranes por proveedor
    
    
    '==== Listados de REPARACIONES ====
    '==================================
    '60 .- Informe de Numeros de Serie
    '61 .- Listado Motivos Pend. Rep.
    '62 .- Listado Resguardo Reparacion
    '63 .- Listado Reparaciones por Día
    '64 .- Listado Reparaciones por Cliente
    '65 .- Listado motivos baja equipos
    
    '406 .- Listado Frecuencia de reparaciones
    '407 .- Sustitución Nº de Serie
    '408 .- Informe Aviso de Averia
    '409 .- Listado Avisos de averia pendientes
    
    
    '==== Listados de ADMINISTRACION ====
    '====================================
    
    '501 .- Listado de Nominas y Gastos
    
    
    '==== Listados de MANTENIMIENTOS ====
    '==================================
    '70 .- Listado Mantenimiento
    '71 .- Listado Revisiones de Mantenimientos
    '72 .- Informe Fichas de Mantenimientos
    '73 .- Listado Altas de Mantenimientos
    '74 .- Prefacturación Mantenimientos
    '75 .- Facturación de Mantenimientos
    '76 .- IGUAL QUE EL 70 pero en ANULADOS
        
        
        
    '77 .- Informe teórico de mantenimientos
    '78 .- Cartas de renovacion
    '79 .- Etiquetas manteimiento
    
    
    '==== Listados OTROS ====
    '==================================
    
    '80 .- Pasar Albaranes Ventas al historico (NO IMPRIME)
    '81 .- Pasar Pedidos Ventas al historico (NO IMPRIME)
       
           
    '82 .- Marcar facturar albaranes
    '83 .- Borre avisos cerrados
       
    
       
       
    '90 .- Etiquetas de Clientes
    '91 .- Cartas a Clientes
    
    '92 .- Informe de Gastos Técnicos
    '93 .- Ticket del TPV
      
    '94 .- Etiquetas estanteria
    
    '95 .- Etiquetas de bultos
    '96 .- Frecuencias
    '97 .- Eliminar facturas
    '99 .- Traspaso a mantenimientos anulados
    
    
    '510 .- AVAB. Correcion de precios caomparando con morales
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoAlPropios As frmAlmAlPropios
Attribute frmMtoAlPropios.VB_VarHelpID = -1
Private WithEvents frmMtoUbica As frmAlmUbicaciones 'Ubicaciones de Almacen
Attribute frmMtoUbica.VB_VarHelpID = -1
Private WithEvents frmMtoMarcas As frmAlmMarcas
Attribute frmMtoMarcas.VB_VarHelpID = -1
Private WithEvents frmMtoTUnidad As frmAlmTipoUnidad
Attribute frmMtoTUnidad.VB_VarHelpID = -1
Private WithEvents frmMtoTArticulo As frmAlmTipoArticulo
Attribute frmMtoTArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoZonas As frmFacZonas
Attribute frmMtoZonas.VB_VarHelpID = -1
Private WithEvents frmMtoRutas As frmFacRutas
Attribute frmMtoRutas.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoTarifas As frmFacTarifas
Attribute frmMtoTarifas.VB_VarHelpID = -1
Private WithEvents frmMtoSituac As frmFacSituaciones
Attribute frmMtoSituac.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmComProveedores
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmFacClientes
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoMotivos As frmRepMotivosPend
Attribute frmMtoMotivos.VB_VarHelpID = -1
Private WithEvents frmMtoAgentes As frmFacAgentesCom
Attribute frmMtoAgentes.VB_VarHelpID = -1
'Private WithEvents frmMtoTiposCon As frmManTiposContrato
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private NumParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------





'Para ademas de insertarlas en la conta, que las contabilice (pase a hsaldos)
'es decir, en el momento que inserta en cabfact tb insertaremos en hlinapu, hacabapu, hsaldos y hsaldosanal (si procede)










Dim IndCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private AntiguaFormaInventariar As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cboSituaAviso_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cboTipMov_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoList_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkImprimeCodigoBarras_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub chkSitaucionArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmbBulto_Click()
    PonerCamposDireccionBultos cmbBulto.ListIndex
End Sub

Private Sub cmbBulto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbProduccion_Click()
'    If PrimeraVez Then Exit Sub
'    PonerLabelsArticulosFrameVisible cmbProduccion.ListIndex = 1
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
   InicializarVbles
   
   Select Case Index
   '========= Frame Listados =================================================
    Case 1 'Frame Listados
        If Me.Optcodigo.Value = True Then
            cadAux = Orden1
        Else
            cadAux = Orden2
        End If
        cadParam = "|pOrden=" & cadAux & "|"
        NumParam = 1
        
        'Añadir el parametro de Empresa
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        NumParam = NumParam + 1
        
        If Trim(txtCodigo(1).Text) <> "" Or Trim(txtCodigo(2).Text) <> "" Then
            'Cadena para seleccion Desde y Hasta
            If OpcionListado = 4 Or OpcionListado = 110 Then
                '4: Listado Tipos de Articulos, 110: List. Ubicaciones
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "T")
            Else
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "N")
            End If
            
            'Homologacion
            If OpcionListado = 58 Then
                If Me.cboMultiPorposito(0).ListIndex > 0 Then
                    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                    cadFormula = cadFormula & " {sprove.homologado} = " & cboMultiPorposito(0).ItemData(cboMultiPorposito(0).ListIndex)
                End If
            End If
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(1).Text <> "" Then cadAux = "Desde: " & txtCodigo(1).Text & " " & txtNombre(1).Text
                If txtCodigo(2).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(2).Text & " " & txtNombre(2).Text
                End If
                
                If OpcionListado = 58 And Me.cboMultiPorposito(0).ListIndex > 0 Then cadAux = cadAux & "     " & cboMultiPorposito(0).Text
                    
                
                
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
                NumParam = NumParam + 1
            End If
        Else
            'Por si acaso no pone desde hasta, y en proveedores, marca homologados
            If OpcionListado = 58 Then
                If Me.cboMultiPorposito(0).ListIndex > 0 Then
                    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                    cadFormula = cadFormula & " {sprove.homologado} = " & cboMultiPorposito(0).ItemData(cboMultiPorposito(0).ListIndex)
                
                    cadAux = cboMultiPorposito(0).Text
                    
                
                
                    cadParam = cadParam & "pDesde=""" & cadAux & """|"
                    NumParam = NumParam + 1
                    cadAux = ""
                 End If
            End If
        End If
        
        'JUNIO 2014
        'Si pone o no mostrar acciones homolgacion
        If OpcionListado = 58 Then
            cadParam = cadParam & "MostrarAcciones=" & Abs(chkProvHomologado(1).Value) & "|"
            NumParam = NumParam + 1
            
            cadParam = cadParam & "MostrarObserva=" & Abs(chkProvHomologado(0).Value) & "|"
            NumParam = NumParam + 1
        End If
        
        
    '========= Frame Informes Almacen ========================================
    Case 2 'Frame Informes Almacen
        If OpcionListado = 7 Then '7: Traspaso Almacen
            indRPT = 1
            cadAux = "scatra"
            cadTitulo = "Informe Traspaso Almacenes"
        ElseIf OpcionListado = 8 Then '8: Movimientos Almacen
            indRPT = 3
            cadAux = "scamov"
            cadTitulo = "Informe Movimientos Almacen"
        End If
        
        cadParam = "|"
        If Not PonerParamEmpresa(cadParam, NumParam) Then Exit Sub
        If PonerParamRPT(indRPT, cadParam, NumParam, cadNomRPT) Then
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(Codigo, "N", 3, 4, "") Then Exit Sub
            End If
        
            If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        End If
                       
                   
                   
    '========= Frame Listado Movimiento de Artículos ========================
    Case 3 'Frame Listado Movimiento de Artículos
        'Nombre fichero .rpt a Imprimir
        cadNomRPT = "rAlmMovim.rpt"
        
        If Not PonerFormulaYParametrosInf9() Then Exit Sub
        'comprobar que hay datos para mostrar en el Informe
        cadAux = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
        If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        conSubRPT = True
    
    '========= Frame de Inventario ==========================================
    Case 4 'Frame de Inventario
        If Not ValidarCamposInventario Then Exit Sub
        If OpcionListado = 19 Then
            cadNomRPT = "rAlmStocksFecha.rpt"
        Else
            'Nombre fichero .rpt a Imprimir
            If vParamAplic.InventarioxProv Then 'Se realiza inventario por Proveedor
                                                'Ordenar por: codprove, codfamia, codartic
                Select Case OpcionListado
                    Case 12: cadNomRPT = "rAlmInvenxProv.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInvenxProvDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenxProvValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracionxProv.rpt"  'Listado Valoracion Stocks (Por Proveedor)
                End Select
            Else 'Ordenar por Cod. Familia y no por Proveedor. Ordenar por: codfamia, codartic.
            
            
           
            
                Select Case OpcionListado
                    
                
                
                
                    Case 12: cadNomRPT = "rAlmInventario.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInventarioDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracion.rpt"  'Listado Valoracion Stocks)
                End Select
            End If
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        bol = PonerFormulaYParametrosInf12()
        If bol Then
                'Morales Nov 2009
                'Para poder hacer inventario con fecha anterior a ultimos movimientos
                
                If OpcionListado = 12 And Not AntiguaFormaInventariar Then InsertarDatosTemporalInventario
         End If
        
        Screen.MousePointer = vbDefault
        If Not bol Then Exit Sub
        
   End Select
    
       
   If OpcionListado = 14 Then 'Actualizar Inventario (NO IMPRIME INFORME)
        If Trim(txtCodigo(21).Text) <> "" Then
            'Quitar las llaves:{tabla.codigo} de la cadena consulta
            'para el FormulaSelection del informe Crystal Report y
            'Tendremos la clausula WHERE para insertar en la tabla:sinven
            cadAux = QuitarCaracterACadena(cadFormula, "{")
            cadFormula = QuitarCaracterACadena(cadAux, "}")
            If ActualizarInventario Then
                MsgBox "La Actualización de Inventario se ha realizado correctamente.", vbInformation
            End If
        Else
            MsgBox "El campo Trabajador debe tener valor", vbInformation
            PonerFoco txtCodigo(21)
            Exit Sub
        End If
        
   Else 'Listados
   
        If Not vUsu.TrabajadorB Then
            Dim B As Boolean
            If OpcionListado = 12 Or OpcionListado = 14 Or OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 19 Then
                'stock a una fecha
                B = (txtCodigo(13).Text = vParamAplic.AlmacenB)
                cadAux = "Almacen: " & txtCodigo(13).Text & " - " & txtNombre(13).Text
            ElseIf OpcionListado = 9 Then
                If txtCodigo(11).Text = "" And txtCodigo(12).Text = "" Then
                    B = True
                    cadAux = "Todos los almacenes"
                Else
                    B = Val(txtCodigo(11).Text) <= vParamAplic.AlmacenB And Val(txtCodigo(12).Text) >= vParamAplic.AlmacenB
                    cadAux = "Almacen: " & txtCodigo(11).Text & " :" & txtCodigo(12).Text
                End If
            End If
            If B Then
                
                cadAux = "Usuario: " & vUsu.Nombre & " (" & vUsu.Login & ")" & vbCrLf & cadAux
                cadAux = "No existen datos: " & vbCrLf & cadAux
                MsgBox cadAux, vbExclamation
                Exit Sub
            End If
            cadAux = ""
        End If
   
   
'        If OpcionListado = 19 Then cadFormula = ""
        If OpcionListado = 19 Then
            cadFormula = "({tmpstockfec.codusu} =" & vUsu.Codigo & ")"
            
            If Me.chkStockFechaAceite.Value = 1 Then cadNomRPT = "rAlmStocksFechaMor.rpt"
            
        End If
        'DAVID
        If OpcionListado = 12 Then
            If Not AntiguaFormaInventariar Then
                cadFormula = "({tmpTomInventario.codusu} =" & vUsu.Codigo & ")"
                cadNomRPT = "morInventario.rpt" 'Nuevo de morales
            End If
        End If
        
        If OpcionListado = 13 Then
            If Not AntiguaFormaInventariar Then cadNomRPT = "morAlmInventarioDif.rpt" 'Nuevo de morales
        End If
        LlamarImprimir

        'Realizar otras acciones segun el informe que llame
        Select Case OpcionListado
            Case 12 'Toma de Inventario
                If frmVisReport.EstaImpreso = True Then
                    frmVisReport.EstaImpreso = False 'para que vuelva a pedirlo
                    PrepararTomaInventario
                End If
            Case 7, 8 'Movimientos
                ActualizarImprimir
            Case 19
                DescargarDatosTMPStockFecha
        End Select
        
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub PrepararTomaInventario()
Dim cadAux As String

    On Error GoTo ETomaInv
    
    
    
    
    'La toma de invntario sera el ultimo del dia ultimo movimiento del dia
   
    If MsgBox("¿Impresión correcta para Actualizar Inventario?", vbQuestion + vbYesNo) = vbYes Then
        'Quitar las llaves:{tabla.codigo} de la cadena consulta
        'para el FormulaSelection del informe Crystal Report y
        'Tendremos la clausula WHERE para insertar en la tabla:sinven
'                cadAux = QuitarCaracterACadena(cadFormula, "{")
'                cadFormula = QuitarCaracterACadena(cadAux, "}")
       If CrearTmpInventario(cadSelect) Then
            If InsertarInventario Then
                MsgBox "Puede pasar a realizar la Entrada de Inventario Real", vbInformation
            End If
       End If
       cadAux = "DROP TABLE IF EXISTS tmpInven "
       Conn.Execute cadAux
    End If
    
ETomaInv:
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim campo As String
Dim devuelve As String
Dim opcion As Byte, numOp As Byte
Dim cadFrom As String





    InicializarVbles
    
    'Si es informe=18 de Stocks Maximos y Minimos comprobar
    'que se ha seleccionado un almacen
    Select Case OpcionListado
    Case 18
        'If OpcionListado = 18 Then
        If txtCodigo(72).Text = "" Then
            MsgBox "Se debe seleccionar un Almacen para el informe.", vbInformation
            Exit Sub
        End If
        cadNomRPT = "rAlmStocksMaxMin.rpt"
        cadFrom = " salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Case 247, 510
        
        
        txtCodigo(107).Text = ""
        txtNombre(107) = ""
    
    Case Else
        'El 6
        cadNomRPT = "rAlmListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
        cadFrom = " sartic"
        cadParam = ""
        For opcion = 0 To 2
            If Me.chkSitaucionArticulo(opcion).Value = 1 Then cadParam = cadParam & "O"
        Next
        If cadParam = "" Then
            MsgBox "Seleccione la situacion del articulo", vbExclamation
            Exit Sub
        End If
        opcion = 0
    End Select
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    'Empresa
    cadParam = cadParam & "pEmpresa=""" & vParam.NombreEmpresa & """|"
    NumParam = NumParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion  ALMACEN
    '--------------------------------------------
    If OpcionListado = 18 And txtCodigo(72).Text <> "" Then
        campo = "{salmac.codalmac}"
        cadFormula = campo & "= " & txtCodigo(72).Text
        
        
    Else
        'Es tarifa para la correccion
        If OpcionListado = 247 And txtCodigo(107).Text <> "" Then
            campo = "{slista.codlista}"
            cadFormula = campo & "= " & txtCodigo(107).Text
        End If
    End If
    
    
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
     If txtCodigo(62).Text <> "" Or txtCodigo(63).Text <> "" Then
        campo = "{sartic.codfamia}"
        'Parametro Desde/Hasta Familila
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 62, 63, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H MARCA
    '--------------------------------------------
    If txtCodigo(64).Text <> "" Or txtCodigo(65).Text <> "" Then
        campo = "{sartic.codmarca}"
        'Parametro Desde/Hasta Marca
        devuelve = "pDHMarca=""Marca: "
        If Not PonerDesdeHasta(campo, "N", 64, 65, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
        campo = "{sartic.codprove}"
        'Parametro Desde/Hasta Proveedor
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 66, 67, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO ARTICULO
    '--------------------------------------------
    If txtCodigo(68).Text <> "" Or txtCodigo(69).Text <> "" Then
        campo = "{sartic.codtipar}"
        'Parametro Desde/Hasta Tipo Articulo
        devuelve = "pDHTipoArt=""Tipo Articulo: "
        If Not PonerDesdeHasta(campo, "T", 68, 69, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    If txtCodigo(70).Text <> "" Or txtCodigo(71).Text <> "" Then
        campo = "{sartic.codartic}"
        'Parametro Desde/Hasta Articulo
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(campo, "T", 70, 71, devuelve) Then Exit Sub
    End If
    
    
    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    Select Case OpcionListado
    Case 6
    
        'Veos que articulos quiere mostrar en funcion de la situacion
        '---------------------------------
        ' si los de situacion NORMAL
        devuelve = ""
        If Me.chkSitaucionArticulo(0).Value = 1 Then
            'SI los BLOQUEADO
            If Me.chkSitaucionArticulo(1).Value = 1 Then
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                    'LOS QUIERE TODOS. NO PONGO NADA
                Else
                    'NO QUEIRE LOS CADUCADOS
                    devuelve = " < 2"
                End If
            Else
                'Los bloqueados NO
                '-----------------
                
                '       si los caducados
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                    devuelve = " <> 1"
                    
                Else
                '       los caducados tampoco, es decir solo los normales
                    devuelve = " = 0"
                End If
            End If
        Else
            'NO QUIERE LOS NORMALES
            If Me.chkSitaucionArticulo(1).Value = 1 Then
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                        devuelve = " > 0"
                Else
                        devuelve = " = 1" 'solo bloqueados
                End If
            Else
                'Es decir, NO QUIERE ni normal ni bloqueados, SOLO caducados
                devuelve = " = 2"
            End If

        End If
        If devuelve <> "" Then
            campo = "{sartic.codstatu} " & devuelve
            AnyadirAFormula cadFormula, campo
            devuelve = ""
        End If
        
    ''''If OpcionListado = 6 Then '6: Listado de Articulos
        numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
        If numOp <> 0 Then opcion = numOp
        numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
        If numOp <> 0 Then opcion = numOp
        numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
        If numOp <> 0 Then opcion = numOp
        numOp = PonerGrupo(4, ListView2.ListItems(4).Text)
        If numOp <> 0 Then opcion = numOp
        opcion = opcion - 1
    
        Select Case opcion
            Case 1 'El group2 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
            Case 2 'El Group3 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
            Case 3, 0 'El Group4 es el Proveedor
                      '0 'El Group1 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                NumParam = NumParam + 1
                
                If opcion = 0 Then
                    campo = "pTitulo3=""" & ListView2.ListItems(4).Text & """"
                    cadParam = cadParam & campo & "|"
                    NumParam = NumParam + 1
                End If
        End Select
       
        'Parametro Orden del Informe
        campo = "pOrden=" & opcion
        cadParam = cadParam & campo & "|"
        NumParam = NumParam + 1
        
    Case 18
    ''ElseIf OpcionListado = 18 Then
        'filtrar ademas por solo articulos con control de stock
        campo = "{sartic.ctrstock}=1"
        AnyadirAFormula cadFormula, campo
    
    
        'David.  Enero 2009
        'Los articulos cuya situacion NO este cadaducado, es decir, NORMAL y BLOQUEADO
        campo = "{sartic.codstatu}<2"
        AnyadirAFormula cadFormula, campo
    
        'Filtrar ademas por stock<stockMin o stock>stockMax
        campo = "{salmac.canstock}"
        If Me.optStockMax Then
            cadFormula = cadFormula & " AND (" & campo & "> {salmac.stockmax})"
        Else
            'David G 30/01/2007
            If optPuntoPedido.Value Then
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.puntoped})"
            Else
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.stockmin})"
            End If
        End If
    
        'Añadir el Parametro de Stocks Maximos o Minimos
        If Me.optStockMax.Value = True Then
            campo = "0"
        Else
            If optPuntoPedido.Value Then
                campo = "2"
            Else
                campo = "1"
            End If
        End If
        cadParam = cadParam & "pStockMax=" & campo & "|"
        NumParam = NumParam + 1
    Case 247, 510
        '                           510=AVAB
        'Correccion de importes
        '-------------------------------------------------------
        
        If BloqueoManual("CORRIGEPRECIOS", "1") Then
            
            
        
            'Mostrare el list
            cadSelect = QuitarCaracterACadena(cadFormula, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            
            
            If OpcionListado = 510 Then
                'Vamos a ver la opcion de buscar los articulos K son componentes
                If Me.chkPreciosProvee.Value = 1 Then
                    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
                    cadSelect = cadSelect & " sartic.factorconversion=1 and sartic.codartic in (Select distinct(codarti1) from sarti1)"
                End If
                
            End If
            
            
            
            frmMensajes.cadWhere = cadSelect
            
            If OpcionListado = 247 Then
                frmMensajes.OpcionMensaje = 20  'Siempre sera este. Solo articulos con componentes componentes
            Else
                frmMensajes.OpcionMensaje = 25  'Correcion AVAB desde morales
            End If
            frmMensajes.vCampos = txtCodigo(107).Text
            frmMensajes.cadWHERE2 = Trim(Me.cmbDecimales.Text)
            'Por no utilizar otra variable
            NumRegElim = 0
            If Me.chkMinimoCorreg.Value = 1 Then NumRegElim = 1
            frmMensajes.Show vbModal
       
        End If
        DesBloqueoManual ("CORRIGEPRECIOS")
        Exit Sub
    End Select
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
        
        
    
        
    LlamarImprimir
End Sub


Private Sub cmdAceptarAviPtes_Click()
'409: Listado Avisos averias pendientes
Dim Tabla As String
Dim campo As String, Cad As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    Tabla = "scaavi"
    cadTitulo = "Listado Avisos de averías Pendientes"
    cadNomRPT = "rRepAvisosPtes.rpt"
    conSubRPT = False
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H RUTA
    '----------------------------------
    If txtCodigo(84).Text <> "" Or txtCodigo(85).Text <> "" Then
        campo = "{sclien.codrutas}"
        Cad = "pDHRuta=""Rutas: "
        If Not PonerDesdeHasta(campo, "N", 84, 85, Cad) Then Exit Sub
    End If



    'Cadena para seleccion SITUACION
    '----------------------------------
    Cad = "pDHSitua=""Situación: "
    If Me.cboSituaAviso.ListIndex = -1 Or Me.cboSituaAviso.ListIndex = 0 Then
        Cad = Cad & "Todas" & """|"
    Else
        Cad = Cad & Me.cboSituaAviso.List(Me.cboSituaAviso.ListIndex) & """|"
        campo = "{" & Tabla & ".situacio}=" & Me.cboSituaAviso.ListIndex - 1
        
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    End If
    cadParam = cadParam & Cad
    NumParam = NumParam + 1


    'Cadena para seleccion D/H FECHA
    '----------------------------------
    If txtCodigo(82).Text <> "" Or txtCodigo(83).Text <> "" Then
        campo = "{scaavi.fechaavi}"
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 82, 83, Cad) Then Exit Sub
    End If


    'Cadena para seleccion D/H RUTA
    '----------------------------------
    If txtCodigo(96).Text <> "" Or txtCodigo(97).Text <> "" Then
        campo = "{scaavi.codtecni}"
        Cad = "pDHTecni=""Técnico: "
        If Not PonerDesdeHasta(campo, "N", 96, 97, Cad) Then Exit Sub
    End If



    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Tabla = Tabla & " INNER JOIN sclien ON " & Tabla & ".codclien=sclien.codclien"
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir
End Sub

Private Sub cmdAceptarDtosFM_Click()
'54: Listado de Descuentos Familia/Marca
'309: Listado precio compras
Dim campo As String, Cad As String
Dim Tabla As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
        
    If OpcionListado = 54 Then
        Tabla = "sdtofm"
        conSubRPT = True
    ElseIf OpcionListado = 309 Then
        Tabla = "slispr"
        cadTitulo = "Listado Precios de compra"
        cadNomRPT = "rComPrecios.rpt"
        conSubRPT = False
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H FAMILIA
    '----------------------------------
    If txtCodigo(75).Text <> "" Or txtCodigo(76).Text <> "" Then
        campo = "{" & Tabla & ".codfamia}"
        If OpcionListado = 309 Then campo = "{sartic.codfamia}"
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 75, 76, Cad) Then Exit Sub
    End If

    If OpcionListado = 54 Then
        'Cadena para seleccion D/H CLIENTE
        '--------------------------------------------
        If txtCodigo(73).Text <> "" Or txtCodigo(74).Text <> "" Then
            campo = "{sdtofm.codclien}"
            Cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 73, 74, Cad) Then Exit Sub
        End If
    
    
        'Cadena para seleccion D/H MARCA
        '--------------------------------------------
        If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
            campo = "{sdtofm.codmarca}"
            Cad = "pDHMarca=""Marca: "
            If Not PonerDesdeHasta(campo, "N", 77, 78, Cad) Then Exit Sub
        End If
    ElseIf OpcionListado = 309 Then
        'Cadena para seleccion D/H PROVEEDOR
        '--------------------------------------------
        If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
            campo = "{" & Tabla & ".codprove}"
            Cad = "pDHProveedor=""Proveedor: "
            If Not PonerDesdeHasta(campo, "N", 79, 80, Cad) Then Exit Sub
        End If
    End If
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 309 Then Tabla = Tabla & " INNER JOIN sartic ON " & Tabla & ".codartic=sartic.codartic"
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir
End Sub


Private Sub cmdAceptarEst_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim Tabla As String
Dim opcPrecio As String
Dim desPrecio As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    'Parametro Precio de Valoracion
    'elegir un Precio para realizar la valoracion
    '==================================================
    desPrecio = "Valoración coste: "
    If Me.optPrecioMP2.Value Then
        opcPrecio = "{slifac.preciomp}" 'precio medio ponderado
        desPrecio = desPrecio & "Precio medio ponderado"
    ElseIf Me.optPrecioUC2.Value Then
        opcPrecio = "{slifac.preciouc}" 'precio ultima compra
        desPrecio = desPrecio & "Precio última compra"
    ElseIf Me.optPrecioStd2.Value Then
        opcPrecio = "{slifac.preciost}" 'precio standard
        desPrecio = desPrecio & "Precio standard"
    End If
    cadParam = cadParam & "pCampo=" & opcPrecio & "|"
    cadParam = cadParam & "pDesCampo=""" & desPrecio & """|"
    NumParam = NumParam + 2
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H familia
    '--------------------------------------------
    If txtCodigo(88).Text <> "" Or txtCodigo(89).Text <> "" Then
        campo = "{sartic.codfamia}"
        param = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 88, 89, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{slifac.codartic}"
        param = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "T", 90, 91, param) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(125).Text <> "" Or txtCodigo(126).Text <> "" Then
        campo = "{slifac.fecfactu}"
        param = "pDHFecha=""Fecha:  "
        If Not PonerDesdeHasta(campo, "F", 125, 126, param) Then Exit Sub
    End If
    
    
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Tabla = " slifac INNER JOIN sartic ON slifac.codartic=sartic.codartic "
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    cadNomRPT = "rFacEstMargen.rpt"
    
    LlamarImprimir
     
End Sub

Private Sub cmdAceptarFichas_Click()
'Fichas de Mantenimientos
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim campo As String


    InicializarVbles
    
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    indRPT = 13
    If Not PonerParamRPT(indRPT, cadParam, NumParam, cadNomRPT) Then Exit Sub
    'Ejercicio
    cadParam = cadParam & "pEjercicio=""" & txtCodigo(61).Text & """|"
    NumParam = NumParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
    If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
        campo = "{scaman.codclien}"
        If Not PonerDesdeHasta(campo, "N", 55, 56, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Then
        campo = "{scaman.codtipco}"
        If Not PonerDesdeHasta(campo, "T", 57, 58, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Nº Mantenimiento
    '--------------------------------------------
    If txtCodigo(59).Text <> "" Or txtCodigo(60).Text <> "" Then
        campo = "{scaman.nummante}"
        If Not PonerDesdeHasta(campo, "T", 59, 60, "") Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H RUTA
    '--------------------------------------------
    If txtCodigo(106).Text <> "" Or txtCodigo(108).Text <> "" Then
        campo = "{sclien.codrutas}"
        If Not PonerDesdeHasta(campo, "N", 106, 108, "") Then Exit Sub
    End If
    
    
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    campo = "   (`ariges2`.`scaman` `scaman` INNER JOIN `ariges2`.`sclien` `sclien` ON `scaman`.`codclien`=`sclien`.`codclien`) INNER JOIN `ariges2`.`stipco` `stipco` ON `scaman`.`codtipco`=`stipco`.`codtipco`"
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    'Si detalla articulos o no
    cadParam = cadParam & "Detallar=" & Abs(Me.chkMante(1).Value) & "|"
    NumParam = NumParam + 1
    LlamarImprimir
End Sub


Private Sub cmdAceptarMante_Click()
'Listado de Mantenimientos
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String

    InicializarVbles
    cadFrom = ""
    
    Select Case OpcionListado
    Case 70, 76
        'comprobar que se ha seleccionado un Tipo de Informe
        If Me.cboTipoList.ListIndex = -1 Then Exit Sub
        'En funcion del valor seleccionado en Tipo Informe se abrira un listado diferente
        Select Case Me.cboTipoList.ListIndex
            Case 0 'Listado Equipos
                cadNomRPT = "rManListManEquipo"
            Case 1 'Listado Pagos
                cadNomRPT = "rManListManPago"
            Case 2 'Listado Importes Contrato
                cadNomRPT = "rManListManImporte"
        End Select
        
        cadTitulo = "Informe Mantenimientos"
        Codigo = "scaman"
        If OpcionListado = 76 Then
            'ANULADOS    rManListManImporteAnu.rpt
            cadTitulo = cadTitulo & " Anulados"
            Codigo = Codigo & "a"
            cadNomRPT = cadNomRPT & "Anu"
        End If
        cadNomRPT = cadNomRPT & ".RPT"
    Case 71
        cadNomRPT = "rManListRevisiones.rpt"
        Codigo = "scaman"
        cadTitulo = "Informe Revisiones"
    Case 78
    
        'PEqueña comprobacion.
        'Fecha obligatoria
        If txtCodigo(109).Text = "" Then
            MsgBox "Debe indicar la fecha", vbExclamation
            Exit Sub
        End If
    
    
        If Not PonerParamRPT(21, cadParam, NumParam, cadNomRPT) Then Exit Sub
        Codigo = "scaman"
    Case 79
        Codigo = "scaman"
    End Select
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
      
      
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
      
      
      
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion ZONA
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Or txtCodigo(46).Text <> "" Then
        campo = "{sclien.codzonas}"
'        'Parametro Desde/Hasta Zona
        devuelve = "pDHZona=""Zona: "
        If Not PonerDesdeHasta(campo, "N", 45, 46, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(47).Text <> "" Or txtCodigo(48).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 47, 48, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtCodigo(49).Text <> "" Or txtCodigo(50).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 49, 50, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(51).Text <> "" Or txtCodigo(52).Text <> "" Then
        campo = "{" & Codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 51, 52, devuelve) Then Exit Sub
    End If
    
    'Motivo de baja. Solo para anulados
    If OpcionListado = 76 Then
        If txtCodigo(115).Text <> "" Or txtCodigo(116).Text <> "" Then
            campo = "{scamana.fechabaj}"
            devuelve = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 115, 116, devuelve) Then Exit Sub
        End If
    
        If txtCodigo(113).Text <> "" Or txtCodigo(114).Text <> "" Then
            campo = "{" & Codigo & ".codincid}"
            'Parametro Desde/Hasta Cliente
            devuelve = "pDHMotivo=""Motivo anul.: "
            If Not PonerDesdeHasta(campo, "T", 113, 114, devuelve) Then Exit Sub
        End If
        
        
        
        
    End If
    
    'Cadena para seleccion FECHA
    '--------------------------------------------
    If OpcionListado = 71 Then
        If txtCodigo(53).Text = "" Or txtCodigo(54).Text = "" Then
            MsgBox "Los campos Fecha Desde/Hasta deben tener valor", vbInformation
            Exit Sub
        End If
        If txtCodigo(53).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(53).Text) & "," & Month(txtCodigo(53).Text) & "," & Day(txtCodigo(53).Text) & ")"
            'Parametro D/H Fecha
            If devuelve <> "" Then
                devuelve = "pDFecha=" & devuelve & "|"
                cadParam = cadParam & devuelve & """|"
                NumParam = NumParam + 1
            End If
        End If
        
        If txtCodigo(54).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(54).Text) & "," & Month(txtCodigo(54).Text) & "," & Day(txtCodigo(54).Text) & ")"
            If devuelve <> "" Then
                devuelve = "pHFecha=" & devuelve & "|"
                cadParam = cadParam & devuelve & """|"
                NumParam = NumParam + 1
            End If
        End If
    End If
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    'cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Esto lo hago siempre para gene temporales
    Conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
    If OpcionListado = 79 Then

        
        'Mostraremos los clientes para imprimirles etiquetas
        If cadSelect <> "" Then
            devuelve = " WHERE " & cadSelect
        Else
            devuelve = ""
        End If
        devuelve = "Select sclien.codclien,nomclien,nifclien FROM " & cadFrom & devuelve
        devuelve = devuelve & " group by 1"
        NumRegElim = 0
        frmMensajes.cadWhere = devuelve
        frmMensajes.OpcionMensaje = 17 'Etiquetas clientes mantenimientos
        frmMensajes.Show vbModal
        If NumRegElim = 0 Then Exit Sub
    
        cadFormula = "({tmpnlotes.codusu} =" & vUsu.Codigo & ")"
    End If
    devuelve = ""
    If OpcionListado = 78 Then
        If Me.chkMante(2).Value Then devuelve = "EMAIL"
    End If
    
    If devuelve = "" Then
        LlamarImprimir
    Else

        '------------------------------------------------------------
        'Envio por mail del desde hasta seleccionado
        'Comprobaremos los mail, que todos tienen

        
        
        
       
       
        DoEvents
        If Me.optMante(0).Value Then
            devuelve = "1"
        Else
            devuelve = "2"
        End If
        
        devuelve = "Select maiclie" & devuelve & " as el_mail,nomclien,scaman.* "
        devuelve = devuelve & " FROM  scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien"
        If cadSelect <> "" Then devuelve = devuelve & " AND " & cadSelect
        
        'INNER JOIN `ariges2`.`stipco` `stipco` ON `scaman`.`codtipco`=`stipco`.`codtipco`"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open devuelve, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
        devuelve = ""
        NumRegElim = 0
        While Not miRsAux.EOF
            If IsNull(miRsAux!el_mail) Then
                devuelve = devuelve & "    - " & miRsAux!nomClien & vbCrLf
            Else
                'INSERTAMOS
                NumRegElim = NumRegElim + 1
                Codigo = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
                Codigo = Codigo & vUsu.Codigo & ",1,'" & Format(txtCodigo(109).Text, FormatoFecha) & "'," & miRsAux!CodClien & ","
                Codigo = Codigo & NumRegElim & ",'" & miRsAux!numMante & "')"
                Conn.Execute Codigo
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        If NumRegElim = 0 Then
            MsgBox "No hay datos para poder enviar por email", vbExclamation
            Exit Sub
        End If
        
        
        If devuelve <> "" Then
            If Len(devuelve) > 500 Then devuelve = Mid(devuelve, 1, 500) & " ....."
            devuelve = "Clientes sin mail: " & vbCrLf & devuelve & "¿Continuar?"
            If MsgBox(devuelve, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        
        If Not PrepararCarpetasEnvioMail Then Exit Sub
            
        
        PonerTamnyosMail True
        frmppal.visible = False
        'Voy arriesgar.
        'Confio en que no envien por mail mas de 32000 facturas (un integer)
        Label4(22).Caption = "Preparando datos"
        Me.PBMail.Max = CInt(NumRegElim)
        Me.PBMail.Value = 0
        
        
        
        NumRegElim = 0
        If GeneracionEnvioMail() Then NumRegElim = 1
            
    
        'Si ha ido todo bien entonces numregelim=1
        If NumRegElim = 1 Then
            'Procederemos a enviarlos por mail
            If Me.optMante(0).Value Then
                '1
                cadSelect = "1"  'de maiclie2
            Else
                cadSelect = "2"  'de maiclie1
            End If
            cadSelect = "Select nomclien,maiclie" & cadSelect
            cadSelect = cadSelect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
        
            
            frmEMail.DatosEnvio = "Carta renovacion|Muchas gracias|" & Abs(chkMante(3).Value) & "|" & cadSelect & "|"
            frmEMail.opcion = 5 'Multienvio de renovacion
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
            
            
        End If
        
        
        
        
        'Es para evitar la cantidad de pantallas abriendose y cerrandose
        Me.visible = False
        PonerTamnyosMail False
        Espera 1
        Unload Me
        frmppal.Show
    
        Screen.MousePointer = vbDefault
    
    
    End If
    
    
End Sub





Private Sub cmdAceptarNSerie_Click()
Dim campo As String
Dim Cad As String

    If txtCodigo(37).Text = "" Or txtCodigo(38).Text = "" Then 'And (txtCodigo(33).Text = "" Or txtCodigo(34).Text = "") Then
        MsgBox "Debe seleccionar un cliente para Imprimir.", vbInformation
        PonerFoco txtCodigo(37)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadNomRPT = "rRepNumSerie.rpt"  'Informe Numeros de Serie Articulos
    cadTitulo = "Informe Num. Serie"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Del DEPARTAMENTO
    '--------------------------------------------
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = Codigo & ".coddirec}"
        'Parametro Desde/Hasta Direc/Dpto
        If vParamAplic.Departamento Then
            Cad = "pDHDirec=""Dpto.: "
        Else
            Cad = "pDHDirec=""Direc.: "
        End If
        If Not PonerDesdeHasta(campo, "N", 39, 40, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Nº CONTRATO
    '--------------------------------------------
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = Codigo & ".nummante}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHContrato=""Nº Mantenimiento: "
        If Not PonerDesdeHasta(campo, "T", 41, 42, Cad) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme("sserie", cadSelect) Then Exit Sub
    
    
    
    LlamarImprimir
    
End Sub


Private Sub cmdAceptarRepxClien_Click()
'Reparaciones por Cliente
Dim devuelve As String
Dim campo As String
Dim Tabla As String

    InicializarVbles
    
    If OpcionListado = 406 Then 'Frecuencia de reparaciones
        Tabla = "schrep"
    Else
        Tabla = "scarep"
    End If
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta CLIENTE
    '---------------------------------------------
    If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{" & Tabla & ".codclien}"
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta DIREC/DPTO
    '-----------------------------------------------
    If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{" & Tabla & ".coddirec}"
        If vParamAplic.Departamento Then
            devuelve = "pDHDpto=""Departamento: "
        Else
            devuelve = "pDHDpto=""Dirección: "
        End If
        If Not PonerDesdeHasta(campo, "N", 35, 36, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If Trim(txtCodigo(43).Text) <> "" Or Trim(txtCodigo(44).Text) <> "" Then
        campo = "{" & Tabla & ".fecentre}"
        If OpcionListado = 406 Then campo = "{" & Tabla & ".fecrepar}"
        devuelve = "pDHFecha=""Fecha Rep.: "
        If Not PonerDesdeHasta(campo, "F", 43, 44, devuelve) Then Exit Sub
    End If
    
   'Comprobar si hay registros a Mostrar antes de abrir el Informe
   If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    
    If OpcionListado <> 406 Then
        cadTitulo = "Reparaciones por Cliente"
        cadNomRPT = "rRepReparxClien.rpt"
        conSubRPT = True
    Else
        cadTitulo = "Frecuencia de Reparaciones"
        cadNomRPT = "rRepFrecuencia.rpt"
        conSubRPT = True
        
        'Nº de Reparaciones, Añadirlo como parametro
        '----------------------------------------------
        cadParam = cadParam & "pNumVeces=" & txtCodigo(0).Text & "|"
        NumParam = NumParam + 1
        
        On Error GoTo EFrecu
        'Insertar en la tabla temporal tmpInformes el total de reparaciones para cada
        'codartic, numserie para el criterio de seleccion introducid
        devuelve = "INSERT INTO tmpinformes(codusu,nombre1,nombre2,campo1) "
        devuelve = devuelve & "SELECT " & vUsu.Codigo & ", codartic,numserie,count(numserie) as campo1 from schrep "
        devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by codartic,numserie"
        Conn.Execute devuelve
        
        'Eliminamos de la tabla aquellos registros que no superen el nº de reparaciones introducido
        devuelve = "DELETE FROM tmpinformes where codusu=" & vUsu.Codigo & " and campo1<=" & txtCodigo(0).Text
        Conn.Execute devuelve
        
        'Volver a comprobar que hay registro a mostrar para ello miramos en la
        'tabla tmpInformes que supere el nº de reparaciones a mostrar
        cadSelect = "codusu=" & vUsu.Codigo
        If Not HayRegParaInforme("tmpinformes", cadSelect) Then
            BorrarTempInformes
            Exit Sub
        End If
    End If
    
    LlamarImprimir
    
    'Eliminar de la tabla temporal
    If OpcionListado = 406 Then BorrarTempInformes
    
EFrecu:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo nº de reparaciones.", Err.Description
End Sub


Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim RS As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String
Dim bOk As Boolean
Dim ConexionContaOk As Boolean
Dim CambiaConta As Boolean
Dim SeguirConLaContabilizacion As Boolean
    If OpcionListado = 223 Then
        If Me.OptClientes.Value Then
            If Me.cboTipMov.ListIndex <= 0 Then
                MsgBox "Seleccione el tipo de factura", vbExclamation
                Exit Sub
            End If
        End If
    End If


    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    Select Case OpcionListado
        Case 63
            Codigo = "{scarep.fecentre}"
            param = "pDHFecha=""Fecha Rep.: "
            NomTabla = "scarep"
            cadNomRPT = "rRepReparxDia.rpt"
            conSubRPT = True
            cadTitulo = "Reparaciones por día"
        Case 73
            'Añadir el parametro total Mantenim. si estamos en Informe de Altas
            devuelve = "SELECT DISTINCT COUNT(*) FROM scaman "
            Set RS = New ADODB.Recordset
            RS.Open devuelve, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                TotalMante = RS.Fields(0).Value
                cadParam = cadParam & "pTotalMante=" & TotalMante & "|"
                NumParam = NumParam + 1
            End If
            RS.Close
            Set RS = Nothing
            
            'Añadir el Total Mantenim. del Periodo anterior
            fecha1 = Day(txtCodigo(31).Text) & "/" & Month(txtCodigo(31).Text) & "/" & Year(txtCodigo(31).Text) - 1
            fecha2 = Day(txtCodigo(32).Text) & "/" & Month(txtCodigo(32).Text) & "/" & Year(txtCodigo(32).Text) - 1
            Codigo = "scaman.fechaini"
            devuelve = CadenaDesdeHastaBD(fecha1, fecha2, Codigo, "F")
            If devuelve <> "" And devuelve <> "Error" Then
                devuelve = "SELECT DISTINCT COUNT(*) FROM scaman WHERE " & devuelve
                Set RS = New ADODB.Recordset
                RS.Open devuelve, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    TotalMante = RS.Fields(0).Value
                    cadParam = cadParam & "pTotalAnte=" & TotalMante & "|"
                    NumParam = NumParam + 1
                End If
                RS.Close
                Set RS = Nothing
            End If
            
            '================= FORMULA =========================
            Codigo = "{scaman.fechaini}"
            param = "pDHFecha=""Fecha: "
            NomTabla = "scaman"
            cadNomRPT = "rManListAltas.rpt"
            cadTitulo = "Informe Altas Mantenimientos"
        
        Case 223
            param = ""
            If Me.OptClientes Then
                Codigo = "{scafac.fecfactu}"
                NomTabla = "scafac"
            Else
                Codigo = "{scafpc.fecrecep}"
                NomTabla = "scafpc"
            End If
    End Select
   
        
    '===================================================
    '================= FORMULA =========================
    
    '== Cadena para seleccion Desde y Hasta FECHA ==
    If OpcionListado = 223 Then
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        
        'fechaini del ejercicio de la conta
        If txtCodigo(31).Text = "" Then txtCodigo(31).Text = Orden1
     
        'fecha fin del ejercicio de la conta
        If txtCodigo(32).Text = "" Then txtCodigo(32).Text = Orden2
     
        'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
        'contabilidad par ello mirar en la BD de la Conta los parámetros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
    End If
    
    devuelve = CadenaDesdeHasta(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F", "Fecha Factura")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
        NumParam = NumParam + 1
    End If
    
    
    '## LAURA 20/06/2008
    '## Añadir frame de selec. factuar en contabilizar
    '- cadena para select en BDatos
    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    
    
    '== Cadena para seleccion Desde y Hasta NºFactura ==
    If OpcionListado = 223 Then
        '- comprobar: si nº factura tienen valor tipoMov tb
        If txtCodigo(121).Text <> "" Or txtCodigo(122).Text <> "" Then
            If Me.cboTipMov.ListIndex = -1 Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) = "" Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            '- añadir desde/hasta factura a cadena seleccion registros
            Codigo = "{scafac.numfactu}"
            devuelve = CadenaDesdeHasta(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N", "Nº Factura")
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            'Parametro D/H nº factura
            If devuelve <> "" And param <> "" Then
                cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
                NumParam = NumParam + 1
            End If
            ' añadir a la formula de bd
            devuelve = CadenaDesdeHastaBD(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N")
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
    
                
        '- añadir tipo movimiento a cadena seleccion
        If Me.cboTipMov.ListIndex >= 0 Then
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                Codigo = "{scafac.codtipom}"
                devuelve = Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3)
                devuelve = Codigo & "=" & DBSet(devuelve, "T")
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
            End If
        End If
    End If

    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    If OpcionListado = 223 Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & NomTabla & ".intconta=0 "
        
        'Nuevo 17 Abril 2009
        ' Contabilizar facturas en B
        If Not Me.OptClientes Then
            If vUsu.TrabajadorB Then
                devuelve = "1"
            Else
                devuelve = "0"
            End If
            cadSelect = cadSelect & " AND " & NomTabla & ".presupuesto = " & devuelve
        End If
        
        'Nuevo 7 Abril 08
        'Hay un parametro que permite contbilizar los tickets agrupados (NO uno a uno)
        'para ello, a partir de los FTI crearemos los FTG (tickets agrupados)
        'y los FTI NO se contabilizaran
        If Me.OptProve.Tag = "" Then
            'Contabilizacion NORMAL. Viene del MENU contabilizar
            'Comprueblo de agrupar tickets o no
            If vParamAplic.ContabilizarTicketAgrupados Then
                'Solo las de clientes
                If Me.OptClientes.Value Then cadSelect = cadSelect & " AND scafac.codtipom <> 'FTI'"
            End If
                
        Else
            'CONTABILZIACION DE LOS TICKETS AGRUPADOS
            'Añado el tipom al cad select
            cadSelect = cadSelect & " AND scafac.codtipom = 'FTG'"
        End If
    End If
    
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    If OpcionListado <> 223 Then
        LlamarImprimir
    Else
    
    
            
        
                
        If Me.OptProve.Tag = "" Then
            If Me.OptClientes.Value Then
                devuelve = "CLI"
            Else
                devuelve = "PRO"
            End If
        Else
            devuelve = "TIK"
        End If
        'Abril 2009.
        'Facturas proveedores van al B,  a la conta B
        'Fracli FAZ van al B tb
        
        CambiaConta = False
        ConexionContaOk = True
        If devuelve = "PRO" Then
            '------------------------------------
            '  Proveedores
            If vUsu.TrabajadorB Then
                If AbrirConexionConta(True) Then
                    CambiaConta = True
                    ConexionContaOk = True
                Else
                    ConexionContaOk = False
                End If
            End If
            
        ElseIf devuelve = "CLI" Then
            'CLIENTES para tipos de factura FAZ, es decir, el B
            If vUsu.TrabajadorB Then
                If AbrirConexionConta(True) Then
                    CambiaConta = True
                    ConexionContaOk = True
                Else
                    ConexionContaOk = False
                End If
            End If
        End If
        
        If ConexionContaOk Then
                        '------------------------------------------------------------------------------
                        '  LOG de acciones.                      5: Facturas compras
                        Set LOG = New cLOG
                        

                        
                    
                        devuelve = "Contabilizar facturas " & devuelve & ":" & vbCrLf & NomTabla & vbCrLf & cadSelect
                        If CambiaConta Then devuelve = devuelve & " Conta b."
                        LOG.Insertar 5, vUsu, devuelve
                        Set LOG = Nothing
                        '-----------------------------------------------------------------------------
                        
                        
                
                    
                        
                        bOk = ContabilizarFacturas(NomTabla, cadSelect)
                        TerminaBloquear
                         'Eliminar la tabla TMP
                        BorrarTMPFacturas
                        'Desbloqueamos ya no estamos contabilizando facturas
                        If Me.OptClientes.Value Then
                            DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
                        Else
                            DesBloqueoManual ("COMCON") 'COMpras CONtabilizar
                        End If
                        Me.FrameProgress.visible = False
                        If Me.FrameTipMov.visible Then
                            Me.FrameRepxDia.Height = 4400
                        Else
                            Me.FrameRepxDia.Height = 3500
                        End If
                        Me.Height = Me.FrameRepxDia.Height + 350
                        Me.Refresh
                        If bOk Then Unload Me
            End If
            If CambiaConta Then AbrirConexionConta False
                    
    
    End If
End Sub



Private Sub cmdAceptarSustNSerie_Click(Index As Integer)
'Sustitucion de un Nº de Serie que este en garantía por otro nº de serie.
Dim SQL As String
Dim RS As ADODB.Recordset

    txtCodigo(81).Text = Trim(txtCodigo(81).Text)
    
    If txtCodigo(81).Text <> "" Then
        'Comprobar que el nuevo nº de serie no existe ya
        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", txtCodigo(81).Text, "T", , "codartic", Me.CadTag, "T")
        If SQL <> "" Then
            MsgBox "Ya existe ese Nº de serie.", vbExclamation
            Exit Sub
        End If
        
        On Error GoTo ESustNSerie
        Conn.BeginTrans
        
        'Insertar un registro con ese nº de serie y todos los valores que tenga el
        'num serie que sustituye
        SQL = "SELECT codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2 FROM sserie "
        SQL = SQL & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If Not RS.EOF Then
            SQL = "(" & DBSet(txtCodigo(81).Text, "T") & ", " & DBSet(RS!codartic, "T", "N") & "," & DBSet(RS!codTipar, "T", "N") & ","
            SQL = SQL & DBSet(RS!CodClien, "N", "S") & "," & DBSet(RS!CodDirec, "N", "S") & "," & DBSet(RS!TieneMan, "N", "S") & ","
            SQL = SQL & DBSet(RS!numMante, "T", "S") & "," & DBSet(RS!ultrepar, "F", "S") & "," & DBSet(RS!fingaran, "F", "S") & ","
            SQL = SQL & DBSet(RS!codTipoM, "T", "S") & "," & DBSet(RS!NumFactu, "N", "S") & "," & DBSet(RS!FechaVta, "F", "S") & ","
            SQL = SQL & DBSet(RS!NumAlbar, "N", "S") & "," & DBSet(RS!numline1, "N", "S") & "," & DBSet(RS!CodProve, "N", "S") & ","
            SQL = SQL & DBSet(RS!numalbpr, "T", "S") & "," & DBSet(RS!fechacom, "F", "S") & "," & DBSet(RS!numline2, "N", "S") & ")"
        End If
        RS.Close
        Set RS = Nothing
        
        If SQL <> "" Then
            SQL = "INSERT INTO sserie (numserie,codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2) VALUES " & SQL
            Conn.Execute SQL
        
            'sustituir el campo numalbar del numserie viejo por 9999999
            'y poner en el campo "numsersu" en num. serie por el que se sustituye
            'limpiar campos del cliente
            SQL = "UPDATE sserie SET numalbar=9999999, numsersu=" & DBSet(txtCodigo(81).Text, "T")
            SQL = SQL & ", codclien=" & ValorNulo & ", coddirec=" & ValorNulo
            SQL = SQL & ", numfactu=" & ValorNulo
            SQL = SQL & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
            Conn.Execute SQL
        End If
    Else
        MsgBox "Debe introducir el Nº Serie por el que se sustituye.", vbInformation
        Exit Sub
    End If

ESustNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Sustitución Nº Serie.", Err.Description
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        Unload Me
    End If
End Sub



Private Sub cmdAceptarTarif_Click()
Dim cadFrom As String

    InicializarVbles
   
   '========= Frame de Tarifas y Descuentos ===============================
    'Nombre fichero .rpt a Imprimir
    'Ordenar por: codtarifa, codfamia, codmarca, codartic
    Select Case OpcionListado
        Case 28: cadNomRPT = "rFacTarifasAlm.rpt"  'Listado Tarifas Articulos
        Case 29: cadNomRPT = "rFacPromociones.rpt"  'Listado Promociones
        Case 30: cadNomRPT = "rFacPreciosEsp.rpt"
        Case 245: cadNomRPT = "rFacTarifasMargen.rpt"
    End Select
    
    If Not PonerFormulaYParametrosInf28() Then Exit Sub
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If cadFormula <> "" Or (OpcionListado = 245) Then
        cadFrom = Codigo & " INNER JOIN sartic ON " & Codigo & ".codartic=sartic.codartic "
    Else
        cadFrom = Codigo
    End If
    
    'seleccionar solo los que tienen margen con error
    If OpcionListado = 245 Then
        If Me.chkMostrarErrores Then
            AnyadirAFormula cadSelect, " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100,4)"
            AnyadirAFormula cadFormula, " {sartic.preciove} <> {sartic.preciouc} + round(({sartic.preciouc} * iif(IsNull({sartic.margecom}),0,{sartic.margecom}))/100,4)"
        End If
    End If
    
    
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    LlamarImprimir
End Sub


Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdDeselTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdElimiaFacturas_Click()
Dim B As Boolean

'Igual hay que quitarlo


    'Proceso de borre de facturas
    If cmbEliFac.ListIndex < 0 Then Exit Sub
    
    
    
    'Tablas que voy a tener que borrar
    'Para que no se queden datos
    cadTitulo = String(60, "*") & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " Se eliminarán los datos con fecha anterior a la solicitada de: " & vbCrLf
    cadTitulo = cadTitulo & " CLIENTES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes, ofertas, hco ofertas, pedidos, hco pedidos" & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "facturas, hco facturas, ventas tpv, reparaciones, hco reparaciones, produccion" & vbCrLf & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " PRVEEDORES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes,  pedidos, hco pedidos, facturas, hco facturas " & vbCrLf & vbCrLf & vbCrLf
    
    Codigo = cadTitulo & "El proceso es irreversible." & vbCrLf & vbCrLf & vbCrLf & "SEGURO QUE DESEA CONTINUAR?"
    
    'Reestablecer variables
    InicializarVbles
    cadTitulo = ""
    
    If MsgBox(Codigo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Codigo = InputBox("Password seguridad")
    Codigo = UCase(Codigo)
    If Codigo <> "ARIADNA" Then Exit Sub
    
    Label3(83).Caption = "Inicio del proceso del borre de facturas"
    Me.cmdElimiaFacturas.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    'Conn.BeginTrans
    B = BorrarFacturas
    'Conn.RollbackTrans
    'Volvemos a dejarlo todo como estaba
    Set miRsAux = Nothing
    Orden1 = ""
    Codigo = ""
    Label3(83).Caption = ""
    Me.cmdElimiaFacturas.Enabled = True
    Screen.MousePointer = vbDefault
    
    If B Then Unload Me
End Sub

Private Sub cmdEtiqBulto_Click()
    If Me.txtClie.Text = "" Then
        MsgBox "Ponga el cliente", vbExclamation
        Exit Sub
    End If
        
    If Val(txtBultos(1).Text) = 0 Then txtBultos(1).Text = "1"
    cadParam = "delete from tmpinformes where codusu =" & vUsu.Codigo
    Conn.Execute cadParam
       
    
    cadParam = "INSERT INTO tmpinformes(codusu   ,codigo1,campo1,nombre1) VALUES (" & vUsu.Codigo & "," & txtClie.Text & "," & vParam.Codigo
    cadParam = cadParam & ",'" & DevNombreSQL(txtNombre(10).Text) & "')"
    Conn.Execute cadParam
       
    'Como puede llevar saltos de linea
    Orden2 = SaltosDeLinea(txtBultos(0).Text)
    'Le pasare los datos
    cadParam = ""
    NumParam = 0
    If PonerParamRPT(19, cadParam, NumParam, cadNomRPT) Then
        Orden1 = "0"

        'Metemos los campos de direccion
        cadParam = cadParam & "Dom=""" & txtBultos(2).Text & """|"
        cadParam = cadParam & "Pob=""" & txtBultos(3).Text & """|"
        cadParam = cadParam & "Pro=""" & Trim(txtBultos(4).Text & "      " & txtBultos(5).Text) & """|"
        
        'AÑado la direccion que se ve
        cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
        cadParam = cadParam & "Texto= """ & Orden2 & """|"
        NumParam = NumParam + 2
        cadSelect = "codusu=" & vUsu.Codigo
        cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
        LlamarImprimir
    End If
        
End Sub

'INTENTARE METERLO DENTRO DE OTRO PROC
Private Sub cmdEtiqEstanteria_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim Tabla As String
Dim RS As ADODB.Recordset
Dim Li As Collection
Dim I As Integer

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    cadParam = cadParam & "|pImprimeBarras=""" & Abs(Me.chkImprimeCodigoBarras.Value) & """|"
    NumParam = NumParam + 1
    cadParam = cadParam & "|numerodecimales=" & Me.cboDecimal.List(cboDecimal.ListIndex) & "|"
    NumParam = NumParam + 1
    
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H familia
    '--------------------------------------------
    If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
        campo = "{sartic.codfamia}"
        param = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 94, 95, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
        campo = "{sartic.codartic}"
        param = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "T", 92, 93, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Fecha
    '--------------------------------------------
    If txtCodigo(123).Text <> "" Or txtCodigo(124).Text <> "" Then
        campo = "{sartic.ultfecpvp}"
        param = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 123, 124, param) Then Exit Sub
    End If
    
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Tabla = " sartic  "
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    
    
    
    'Borro tmptemporal
    Tabla = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    Conn.Execute Tabla
    
    'Añadire los tipos de IVA a esta tabla
    Tabla = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.Codigo & ",codigiva from sartic"
    If cadSelect <> "" Then Tabla = Tabla & " WHERE " & cadSelect
    Tabla = Tabla & " GROUP BY codigiva"
    Conn.Execute Tabla
    
    
    
    
    
    'AHora desde conta cargo los % de IVA desde la conta
    Set RS = New ADODB.Recordset
    Tabla = "Select * from tmpinformes where codusu =" & vUsu.Codigo
    RS.Open Tabla, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Li = New Collection
    While Not RS.EOF
        Li.Add Val(RS.Fields(1))
        RS.MoveNext
    Wend
    RS.Close
    
    'Abrimos los IVAS en conta
    Tabla = "Select codigiva,porceiva from tiposiva"
    RS.Open Tabla, ConnConta, adOpenKeyset, adLockOptimistic, adCmdText
    For I = 1 To Li.Count
        Tabla = "codigiva = " & Li.Item(I)
        RS.Find Tabla, , , 1
        If RS.EOF Then
            MsgBox "Tipo de IVA no encontrado en la contabilidad" & Tabla, vbExclamation
            RS.Close
            Exit Sub
        Else
            Tabla = "UPDATE tmpinformes SET porcen1 =" & TransformaComasPuntos(CStr(RS!PorceIVA))
            Tabla = Tabla & " WHERE codusu =" & vUsu.Codigo & " AND codigo1 = " & RS!codigiva
            Conn.Execute Tabla
        End If
    Next I
    RS.Close
    Set Li = Nothing
    
    
    'Borramos los datos de la tabla donde iran los articulos
    Tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    Conn.Execute Tabla
    I = Me.cboDecimal.List(cboDecimal.ListIndex)
    If I = 0 Then
        Tabla = "0"
    Else
        Tabla = "#,##0." & Mid("0000", 1, I)
    End If
    frmMensajes.cadWHERE2 = Tabla
    frmMensajes.cadWhere = cadSelect
    frmMensajes.OpcionMensaje = 15
    frmMensajes.Show vbModal
    
    'Si ha devuelto seleccionados
    Tabla = " tmpnseries   "
    cadFormula = " codusu =" & vUsu.Codigo
    
    If Not HayRegParaInforme(Tabla, cadFormula) Then Exit Sub
    
    cadFormula = "({tmpnseries.codusu} =" & vUsu.Codigo & ")"
    
    campo = ""
    If Not PonerParamRPT(23, cadParam, NumParam, campo) Then
        cadNomRPT = "rEtiqEsta.rpt"
    Else
        cadNomRPT = campo
    End If
    
    LlamarImprimir
    
    BorrarTempInformes
    
    'Borramos los datos de la tabla donde iran los articulos
    Tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    Conn.Execute Tabla
    
End Sub


Private Sub cmdFactAlbaranes_Click()
    Codigo = "¿Seguro que desea continuar?"
    If MsgBox(Codigo, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If HacerSQLListado82_83 Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
End Sub

Private Sub cmdFrecuencias_Click()
        'Le pasare los datos
    cadParam = ""
    NumParam = 0
    If PonerParamRPT(19, cadParam, NumParam, cadNomRPT) Then
        Orden1 = "0"
       ' If Me.optDirEnvio(1).Value Then Orden1 = "1"
        
        'AÑado la direccion que se ve
        cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
        cadParam = cadParam & "Texto= """ & Orden2 & """|"
        NumParam = NumParam + 2
        cadSelect = "codusu=" & vUsu.Codigo
        LlamarImprimir
    End If

End Sub


Private Sub cmdHcoMante_Click()
    Codigo = ""
    For IndCodigo = 110 To 112
        If txtCodigo(IndCodigo).Text = "" Then Codigo = Codigo & "M"
        If IndCodigo > 110 Then If txtNombre(IndCodigo).Text = "" Then Codigo = Codigo & "M"
    Next IndCodigo
    If Codigo <> "" Then
        MsgBox "Rellene correctamente todos los datos", vbExclamation
        Exit Sub
    End If
    'CUATRO CAMPOS. El primero de control
    CadenaDesdeOtroForm = "OK|" & txtCodigo(110).Text & "|" & txtNombre(111).Text & "|" & txtCodigo(112).Text & "|"
    Unload Me
End Sub

'===================================================
'===================================================
' Informe teorico mantenimientos
Private Sub cmdManteTeorico_Click()
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String

    InicializarVbles

    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1
    
    
    
        cadNomRPT = "rManListTeorico.rpt"
    
        
        cadTitulo = "Informe Mantenimientos"
        Codigo = "scaman"
    
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 102, 103, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(104).Text <> "" Or txtCodigo(105).Text <> "" Then
        campo = "{" & Codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 104, 105, devuelve) Then Exit Sub
    End If
       
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Si  detalla o no
    cadParam = cadParam & "Detallar=" & Abs(Me.chkMante(0).Value) & "|"
    NumParam = NumParam + 1

    
    LlamarImprimir
End Sub

Private Sub cmdSelTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = True
    Next I
End Sub


Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub




Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = -1
        
        Select Case OpcionListado
        Case 1, 2, 3, 4, 61, 20, 21, 22, 23, 24, 27, 58, 110
            '1:Listado de Marcas, 2:Almacenes Propios, 3:Tipos de Unidad
            '4:Tipos de Artículos, 6:Artículos
            '61:Motivos Pen. Rep
            '58:Proveedores, 110:Ubicaciones
             'PonerFoco txtCodigo(1)
             IndiceFoco = 1
        Case 6 '6: Informe de Articulos
            'PonerFoco txtCodigo(62)
            IndiceFoco = 62
        Case 7, 8 '7: Informe Traspaso Almacenes/Historico
                  '8: Informe Movimientos Almacen/Historico
            'PonerFoco txtCodigo(3)
            IndiceFoco = 3
        Case 9 'Informe Movimientos Artículos
            'PonerFoco txtCodigo(5)
            IndiceFoco = 5
        Case 12, 13, 14, 15, 16, 17, 19
                        '12: Listado Toma de Inventario Articulos
                        '13: Listado Diferencias de Inventario Articulos
                        '14: Actualizar Diferencias de Inventario (No IMPRIME INFORME)
                        '15: Listado Articulos Inactivos
                        '16: Listado Valoracion de Stocks Inventariados
                        '17: Listado Valoración Stocks
                        '19: Inf. Stocks a una Fecha
            'PonerFoco txtCodigo(13)
            IndiceFoco = 13
        Case 18      '18: Informe Stocks MAximos y Minimos
            'PonerFoco txtCodigo(72)
            IndiceFoco = 72
        Case 28, 29, 30 '28: Informe Tarifas de Articulos
                    '29: Informe Promociones
                    '30: Informe Precios Especiales
            'PonerFoco txtCodigo(23)
            IndiceFoco = 23
        Case 31, 73 '31: Informe Ofertas
                    '73: Listado Altas Mantenimientos
            'PonerFoco txtCodigo(31)
            IndiceFoco = 31
        Case 54 'Listado Descuentos Familia/ Marca
            'PonerFoco txtCodigo(73)
            IndiceFoco = 73
        Case 60 '60: Informe Reparacions - Nº Series
            'PonerFoco txtCodigo(37)
            IndiceFoco = 37
        Case 63, 73
            '63: Listado Reparaciones x día
            IndiceFoco = 31
        
        
        Case 223
            '223: Contabilizar facturas
            If Me.OptProve.Tag = "" Then
                'Contabilizacion normal clie/prov
                IndiceFoco = 31
            
            Else
                'TICKETS AGRUPADOS
                'Contabilizacion de facturas de tickets agrupadas. Lanzamos YA el proceso
                DoEvents
                cmdAceptarRepxDia_Click
                Me.Refresh
                Unload Me
                Exit Sub
            End If
        Case 246 '246: Informe margen ventas x articulo
            'PonerFoco txtCodigo(88)
            IndiceFoco = 88
        Case 64, 406 '64: Listado Reparaciones x Cliente
                     '406: List. Frecuencia de Reparaciones
            'PonerFoco txtCodigo(33)
            IndiceFoco = 33
        Case 70, 71, 76, 79 'Listado Mantenimientos
            'PonerFoco txtCodigo(45)
            IndiceFoco = 45
        Case 72 'Informe Fichas Mantenimientos
            'PonerFoco txtCodigo(55)
            IndiceFoco = 55
            
        Case 77
            'PonerFoco txtCodigo(102)
             IndiceFoco = 102
        Case 78
            'PonerFoco txtCodigo(109)
            IndiceFoco = 109
            
        Case 82, 83
            'Marca facturar a 1
            IndiceFoco = 119
            
        Case 309 '309:Listado precios de compra
            'PonerFoco txtCodigo(79)
            IndiceFoco = 79
        Case 407 'Sustitución Nº Serie
            'PonerFoco txtCodigo(81)
            IndiceFoco = 81
        Case 409 'List. Avisos de averias pendientes
            'PonerFoco txtCodigo(82)
            IndiceFoco = 82
        Case 95
            PonerFoco txtClie
            
        Case 99
            'PonerFoco txtCodigo(110)
            IndiceFoco = 110
        Case 247  'y Correccion de listados de precios tarias etc
             'PonerFoco txtCodigo(107)
             IndiceFoco = 107
             
        Case 510
            'AVAB
            IndiceFoco = 62
        End Select
        If IndiceFoco >= 0 Then PonerFoco txtCodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del formulario
    Me.Icon = frmppal.Icon

    PrimeraVez = True
    limpiar Me
    AntiguaFormaInventariar = False
    'Ocultar todos los Frames de Formulario
    frameListado.visible = False
    FrameInfAlmacen.visible = False
    FrameMovArtic.visible = False
    FrameInventario.visible = False
    FrameTarifas.visible = False
    FrameRepNSerie.visible = False
    FrameRepxDia.visible = False
    FrameRepxClien.visible = False
    FrameMantenimientos.visible = False
    Me.FrameFichasMan.visible = False
    FrameInfArticulos.visible = False
    FrameDtosFM.visible = False
    FrameRepSustNSerie.visible = False
    FrameListAvisosPtes.visible = False
    FrameEstMargenes.visible = False
    Me.FrameEtiqEstanteria.visible = False
    FrameBultos.visible = False
    Me.FrameFrecuencia.visible = False
    FrEliminarFacturas.visible = False
    FrameListMant2.visible = False
    FrameEnvioMail.visible = False
    FrameHcoMante.visible = False
    FrameAlbaranesMarcaFacturar.visible = False
    FrameHomologacion.visible = False
    CommitConexion
    
    
    
    cadTitulo = ""
    cadNomRPT = ""
    
    Select Case OpcionListado
        Case 1 To 19, 247, 510 'Listado de ALMACEN
            ListadosAlmacen H, W
        Case 100 To 199 'Listados de ALMACEN
            ListadosAlmacen H, W
        Case 20 To 30 'Listadod de FACTURACION
            ListadosFacturacion H, W
        Case 70 To 89 'Listados de MANTENIMIENTO
            ListadosMantenimiento H, W
        Case 245, 246 'Listados tarifas
            ListadosFacturacion H, W
        Case 300 To 390 'Listados de COMPRAS
            ListadosCompras H, W
        Case 407 To 490 'Listados de Reparaciones
            ListadosReparaciones H, W
    End Select
    
    
    Select Case OpcionListado
    
    'LISTADOS DE FACTURACION
    '-----------------------
        
    Case 54 '54: Listado Descuentos Familia/Marca
        H = 5450
        W = 6920
        PonerFrameVisible Me.FrameDtosFM, True, H, W
        Me.Frame4.visible = False
        indFrame = 6
        
    Case 58 '58: listado Proveedores
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado Proveedores"
        indFrame = 1
        Codigo = "{sprove.codprove}"
        Orden1 = "{sprove.codprove}"
        Orden2 = "{sprove.nomprove}"
        
        
        FrameHomologacion.visible = True
        Me.cboMultiPorposito(0).ListIndex = 0
        
    'LISTADOS DE REPARACIONES
    '-------------------------
    Case 60 '60: Informe Nº Series
        H = 5415
        W = 6675
        PonerFrameVisible Me.FrameRepNSerie, True, H, W
        indFrame = 6
        Codigo = "{sserie"
        
     Case 61, 65  'Listados de Motivos Pend. Rep.
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado de Motivos"
        indFrame = 1
        If OpcionListado = 61 Then
            Codigo = "{smotre.codmotre}"
            Orden1 = "{smotre.codmotre}"
            Orden2 = "{smotre.nommotre}"
        Else
            Codigo = "{smotba.codmotiv}"
            Orden1 = "{smotba.codmotiv}"
            Orden2 = "{smotba.desmotiv}"
        End If
        
    Case 63, 73, 223, 224, 248
                '63: Listado Reparaciones por Día
                '73: Listado Altas Mantenimientos
                '223,224,248  Contabi facturas
                
        PonerFrameRepxDiaVisible True, H, W
        indFrame = 7
        If Me.OptProve.Tag = "" Then
            txtCodigo(31).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(32).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        If OpcionListado = 223 Then
            Dim Cad As String
            
            'If vParamAplic.ContabilizarTicketAgrupados Then
            '    cad = "codtipom like 'FA%'"
            'Else
            '    cad = "(codtipom like 'FA%' or codtipom='FTI')"
            'End If
            
            If vUsu.TrabajadorB Then
                Cad = "codtipom = 'FAZ'"
            Else
                Cad = "codtipom <> 'FAZ'"
            End If
            Cad = Cad & " and not isnull(letraser) and trim(letraser)<>''"
            CargarCombo_TipMov Me.cboTipMov, "stipom", "codtipom", "nomtipom", Cad, True
        End If
        
    Case 64, 406 'Listado Reparaciones por Cliente
                 '406: Listado Frecuencia de reparaciones
        H = 5415
        W = 6850
        PonerFrameVisible Me.FrameRepxClien, True, H, W
        indFrame = 8
        txtCodigo(43).Text = Format(Now, "dd/mm/yyyy")
        txtCodigo(44).Text = Format(Now, "dd/mm/yyyy")
        cadTitulo = "Reparaciones por Cliente"
        conSubRPT = False
        Me.Frame1.visible = (OpcionListado = 406)
        If OpcionListado = 406 Then
             cadTitulo = "Frecuencia de Reparaciones"
             Me.lblTitulo(8).Caption = "Frecuencia de Reparaciones"
             Me.Label4(21).Caption = "Fecha Reparación:"
             txtCodigo(0).Text = "1"
        End If
        
        
        
    Case 82, 83
        
        'LIstado etiquetas estanterias
        H = Me.FrameAlbaranesMarcaFacturar.Height
        W = FrameAlbaranesMarcaFacturar.Width
        PonerFrameVisible Me.FrameAlbaranesMarcaFacturar, True, H, W
        indFrame = 82
        If OpcionListado = 82 Then
            cadTitulo = "Poner marca facturación"
            
        Else
            Label7(3).Caption = "Borre avisos cerrados"
        End If
        txtCodigo(117).visible = OpcionListado = 82
        txtCodigo(118).visible = OpcionListado = 82
        Frame7.visible = OpcionListado = 83
        conSubRPT = False
    Case 94
        'LIstado etiquetas estanterias
        H = Me.FrameEtiqEstanteria.Height
        W = FrameEtiqEstanteria.Width
        PonerFrameVisible Me.FrameEtiqEstanteria, True, H, W
        indFrame = 94
        cadTitulo = "Etiq. estanteria"
        conSubRPT = False
        cboDecimal.ListIndex = 4
        
    Case 95
        'LIstado etiquetas estanterias
        H = Me.FrameBultos.Height
        W = FrameBultos.Width
        PonerFrameVisible Me.FrameBultos, True, H, W
        indFrame = 95
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
'        If vParamAplic.Departamento Then
'            optDirEnvio(1).Caption = "Departamento"
'        Else
'            optDirEnvio(1).Caption = "Dirección"
'        End If
        LimpiarTextosBultos
        Me.cmbBulto.Clear
    Case 96
        
        H = Me.FrameFrecuencia.Height
        W = FrameFrecuencia.Width
        PonerFrameVisible Me.FrameFrecuencia, True, H, W
        indFrame = 96
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
        HabilitarTextoCliente False
        
    Case 97
        H = Me.FrEliminarFacturas.Height
        W = Me.FrEliminarFacturas.Width
        PonerFrameVisible FrEliminarFacturas, True, H, W
        indFrame = 97
        cadTitulo = "Eliminar facturas"
        conSubRPT = False
        'Textos
        '--------------------------------------------------------------------
        Label11(0).Caption = "Este proceso es irreversible." & vbCrLf & " No deberia haber nadie trabajando en esta empresa y " & vbCrLf & _
            "deberia hacer una copia de seguridad."
        
        Label11(1).Caption = ""
        CargaFechasPosibleEliminacion
        
    Case 99
        
        H = Me.FrameHcoMante.Height
        W = Me.FrameHcoMante.Width
        PonerFrameVisible FrameHcoMante, True, H, W
        indFrame = 99
        cadTitulo = "Pasar a mantenimientos anulados"
        conSubRPT = False
        txtCodigo(110).Text = Format(Now, "dd/mm/yyyy")

    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
End Sub



Private Sub Frame3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Actividades de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoAgentes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes Comerciales
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlPropios_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(32).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(32).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If IndCodigo > 0 Then
        txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        'EL 0 es para el listado de bultos
        Me.txtClie.Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtClie_LostFocus
        
    End If

End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMarcas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Artículos
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMotivos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Artículos
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoRutas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Rutas
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSituac_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTarifas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tarifas
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Artículo
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTiposCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Contrato
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTUnidad_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Unidad
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoUbica_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Ubicaciones de Almacen
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoZonas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Zonas
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgBuscar(1).Tag = Index
    IndCodigo = Index
    
    Select Case Index
    Case 1, 2 'FrameListado
        Select Case OpcionListado
            Case 1 'Listado de MARCAS
                AbrirFrmMarcas
                    
            Case 2 'Listado de ALMACENES Propios
                AbrirFrmAlmPropios
            
            Case 3  'Listado de Tipos de Unidad
                Set frmMtoTUnidad = New frmAlmTipoUnidad
                frmMtoTUnidad.DatosADevolverBusqueda = "0|1"
                frmMtoTUnidad.DeConsulta = True
                frmMtoTUnidad.Show vbModal
                Set frmMtoTUnidad = Nothing
            
            Case 4  'Listado de Tipos de Articulos
                AbrirFrmTipoArt

            Case 110 'Listado de Ubicaciones de Almacen
                Set frmMtoUbica = New frmAlmUbicaciones
                frmMtoUbica.DatosADevolverBusqueda = "0|1"
                frmMtoUbica.DeConsulta = True
                frmMtoUbica.Show vbModal
                Set frmMtoUbica = Nothing
        
            
            Case 20 'Listado de Actividades de Clientes
                AbrirFrmActividades
            
            Case 21 'Listado de Zonas de Clientes
                AbrirFrmZonas
            
            Case 22 'Listado de Rutas de Asistencia
                AbrirFrmRutas
                
'                Set frmMtoRutas = New frmFacRutas
'                frmMtoRutas.DatosADevolverBusqueda = "0|1"
'                frmMtoRutas.DeConsulta = True
'                frmMtoRutas.Show vbModal
'                Set frmMtoRutas = Nothing
            
            Case 23 'Listado de Formas de Envío
                Set frmMtoFEnvio = New frmFacFormasEnvio
                frmMtoFEnvio.DatosADevolverBusqueda = "0|1"
                frmMtoFEnvio.DeConsulta = True
                frmMtoFEnvio.Show vbModal
                Set frmMtoFEnvio = Nothing
            
            Case 24 'Listado de Tarifas Venta
                AbrirFrmTarifas
            
            Case 27 'Listado de Situaciones Especiales
                Set frmMtoSituac = New frmFacSituaciones
                frmMtoSituac.DatosADevolverBusqueda = "0|1"
                frmMtoSituac.DeConsulta = True
                frmMtoSituac.Show vbModal
                Set frmMtoSituac = Nothing
                
            Case 58
                'DAVID
                IndCodigo = Index
                Set frmMtoProveedor = New frmComProveedores
                frmMtoProveedor.DatosADevolverBusqueda = "0|1"
                frmMtoProveedor.Show vbModal
                Set frmMtoProveedor = Nothing
            Case 61 'Listado de Motivos Pend. Rep.
                Set frmMtoMotivos = New frmRepMotivosPend
                frmMtoMotivos.DatosADevolverBusqueda = "0|1"
                frmMtoMotivos.DeConsulta = True
                frmMtoMotivos.Show vbModal
                Set frmMtoMotivos = Nothing
        End Select
        
    Case 3, 4 'FrameInfAlmacen
            If OpcionListado = 7 Or OpcionListado = 8 Then
'            Case 7, 8 '7: Informe de Traspasos de Almacenes
                  '8: Informe de Movimientos de Almacen
                MandaBusquedaPrevia ""
            End If
    End Select
    
    PonerFoco Me.txtCodigo(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0, 1, 6, 7, 35, 36, 43, 44, 49, 50, 75, 76, 77, 80, 81, 93, 94 'cod. CLIENTE
            Select Case Index
                Case 0, 1: IndCodigo = Index + 73
                Case 6, 7: IndCodigo = Index + 27
                Case 35, 36: IndCodigo = Index + 20
                Case 43, 44: IndCodigo = Index + 4
                Case 49, 50: IndCodigo = Index - 12
                Case 75: IndCodigo = 0
                Case 76, 77, 80, 81: IndCodigo = Index + 22
                Case 93, 94: IndCodigo = Index + 24
            End Select
            AbrirFrmClientes
        
        Case 2, 3, 13, 14, 19, 20, 31, 32, 57, 58, 67, 68, 73, 74 'cod. FAMILIA
            Select Case Index
                Case 2, 3: IndCodigo = Index + 73
                Case 13, 14: IndCodigo = Index + 3
                Case 19, 20: IndCodigo = Index + 43
                Case 31, 32: IndCodigo = Index - 24
                Case 57, 58: IndCodigo = Index - 32
                Case 67, 68, 73, 74: IndCodigo = Index + 21
            End Select
            Set frmMtoFamilia = New frmAlmFamiliaArticulo
            frmMtoFamilia.DatosADevolverBusqueda = "0|1"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
            
            
        Case 90, 91, 92
            IndCodigo = 22 + Index
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 4, 5, 21, 22, 59, 60 'cod. MARCA
            Select Case Index
                Case 4, 5: IndCodigo = Index + 73
                Case 21, 22: IndCodigo = Index + 43
                Case 59, 60:  IndCodigo = Index - 32
            End Select
            AbrirFrmMarcas
            
        Case 8, 9, 51, 52 'cod. Direc/DPTO
'            Select Case Index
'                Case 8, 9:
'                Case 51, 52: indCodigo = Index - 12
'            End Select
        
        Case 10, 18, 33, 34 'cod. ALMACEN
            Select Case Index
                Case 10: IndCodigo = Index + 3
                Case 18: IndCodigo = Index + 54
                Case 33, 34: IndCodigo = Index - 22
            End Select
            AbrirFrmAlmPropios
            
        Case 11, 12, 27, 28, 29, 30, 61, 62, 69, 70, 71, 72 'cod. ARTICULO
            Select Case Index
                Case 11, 12: IndCodigo = Index + 3
                Case 27, 28: IndCodigo = Index + 43
                Case 29, 30: IndCodigo = Index - 24
                Case 61, 62: IndCodigo = Index - 32
                Case 69, 70, 71, 72: IndCodigo = Index + 21
            End Select
            Set frmMtoArticulos = New frmAlmArticulos
            frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
        Case 25, 26 'cod TIPO ARTICULO
            IndCodigo = Index + 43
            AbrirFrmTipoArt

        Case 55, 56
            IndCodigo = Index - 32
            If OpcionListado = 30 Then 'segun Informe mismo boton abre 2 distintas
               AbrirFrmClientes
            Else 'cod. TARIFA
                AbrirFrmTarifas
            End If
            
        Case 15, 16, 23, 24, 63, 64 'cod. PROVEEDOR
            Select Case Index
                Case 15, 16: IndCodigo = Index + 3
                Case 23, 24: IndCodigo = Index + 43
                Case 63, 64: IndCodigo = Index + 16
            End Select
            Set frmMtoProveedor = New frmComProveedores
            frmMtoProveedor.DatosADevolverBusqueda = "0|1"
            frmMtoProveedor.Show vbModal
            Set frmMtoProveedor = Nothing
            
        Case 41, 42, 86, 88 'cod. ZONA
            If Index <= 42 Then
                IndCodigo = Index + 4
            Else
                '86,88
                IndCodigo = Index + 20
            End If
            AbrirFrmZonas
            
        Case 17, 96, 97, 89 'cod. TRABAJADOR
            If Index = 89 Then
                IndCodigo = 111
            Else
                IndCodigo = 21
            End If
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 45, 46 'cod. AGENTE
            IndCodigo = Index + 4
            Set frmMtoAgentes = New frmFacAgentesCom
            frmMtoAgentes.DatosADevolverBusqueda = "0|1"
            frmMtoAgentes.Show vbModal
            Set frmMtoAgentes = Nothing
            
        Case 37, 38, 47, 48, 82, 83 'cod. TIPO CONTRATO (= nº mantenimiento)
            Select Case Index
                Case 37, 38: IndCodigo = Index + 20
                Case 47, 48: IndCodigo = Index + 4
                Case 82, 83: IndCodigo = Index + 22
            End Select
'            Set frmMtoTiposCon = New frmManTiposContrato
'            frmMtoTiposCon.DatosADevolverBusqueda = "0|1"
'            frmMtoTiposCon.Show vbModal
'            Set frmMtoTiposCon = Nothing
        
        Case 39, 40, 53, 54 'cod. Nº CONTRATO (= nº mantenimiento)

        
        Case 84, 85 'RUTA DEL CLIENTE
            IndCodigo = Index
            AbrirFrmRutas
        Case 87
            IndCodigo = 107
            AbrirFrmTarifas
    End Select
    PonerFoco txtCodigo(IndCodigo)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'frameMovArtic
            IndCodigo = 9
        Case 1 'frameMovArtic
            IndCodigo = 10
        Case 2 'frameInventario (indFrame=4)
            IndCodigo = 20
        Case 3 'frameInventario (indFrame=4)
            IndCodigo = 22
        Case 4 'frameReparacionesxDia (indFrame=7)
            IndCodigo = 31
        Case 5 'frameReparacionesxDia (indFrame=7)
            IndCodigo = 32
        Case 6 'frameReparacionesxClien (indFrame=8)
            IndCodigo = 43
        Case 7 'frameReparacionesxClien (indFrame=8)
            IndCodigo = 44
        Case 8 'frameMAntenimientos
            IndCodigo = 53
        Case 9 'frameMAntenimientos
            IndCodigo = 54
        Case 10 'FrameListAvisosPtes
            IndCodigo = 82
        Case 11 'FrameListAvisosPtes
            IndCodigo = 83
        Case 13, 14
            IndCodigo = Index + 102
        Case 15, 16
            IndCodigo = Index + 104
        Case 17, 18, 19, 20
            IndCodigo = Index + 106
        
        Case 109
            IndCodigo = 109
   End Select
   
   
   PonerFormatoFecha txtCodigo(IndCodigo)
   If txtCodigo(IndCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(IndCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(IndCodigo)
End Sub




Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptClientes_Click()
    If Me.OptClientes.Value = True Then
        Label2(2).Caption = "Fecha Factura: "
    End If
    
    Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub optDirEnvio_Click(Index As Integer)
    If Index = 0 Then
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 1)
    Else
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 2)
    End If
End Sub

Private Sub optDirEnvio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar(1)
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptProve_Click()
    If Me.OptProve.Value = True Then
        Label2(2).Caption = "Fecha Recepción: "
    End If
    
     Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
    
End Sub


Private Sub txtBultos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 0 Then KEYpress KeyAscii
End Sub

Private Sub txtClie_GotFocus()
    PonerFoco txtClie
    
End Sub

Private Sub txtClie_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtClie_LostFocus()
Dim Reestablecer As Boolean
Dim CliVario As Boolean
Dim RS As ADODB.Recordset
                Screen.MousePointer = vbHourglass
                txtClie.Text = Trim(txtClie.Text)
                Orden2 = ""
                CliVario = False
                If txtClie = "" Then
                    Reestablecer = True
                Else
                    If Not PonerFormatoEntero(txtClie) Then
                        Reestablecer = True
                    Else
                        cmbBulto.Clear
                        Set RS = New ADODB.Recordset
                        Codigo = "select nomclien,domclien,sclien.codpobla as cpos,sclien.pobclien,proclien,sdirec.*,clivario from sclien left join sdirec on sclien.codclien=sdirec.codclien "
                        Codigo = Codigo & " WHERE sclien.codclien =" & txtClie.Text
                        RS.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Orden1 = ""
                        
                        While Not RS.EOF
                            'Meto primero la direccion de la ficha
                            If Orden1 = "" Then
                                cmbBulto.AddItem "Ppal:  " & DBLet(RS.Fields(1), "T") & " - " & DBLet(RS.Fields(3), "T")
                                txtBultos(2).Tag = DBLet(RS.Fields(1), "T") & "|"
                                txtBultos(3).Tag = DBLet(RS.Fields(3), "T") & "|"
                                txtBultos(4).Tag = DBLet(RS.Fields(2), "T") & "|"
                                txtBultos(5).Tag = DBLet(RS.Fields(4), "T") & "|"
                                txtBultos(6).Tag = "|"
                                Orden1 = "T"
                                
                                Orden2 = RS!nomClien
                                CliVario = DBLet(RS!CliVario, "N") = 1
                            End If
                            'Las direcciones alternativas
                            If Not IsNull(RS!domdirec) Then
                                'TIENE DIRECCION ALTERNATIVA
                                txtBultos(2).Tag = txtBultos(2).Tag & DBLet(RS!domdirec, "T") & "|"
                                txtBultos(3).Tag = txtBultos(3).Tag & DBLet(RS!pobdirec, "T") & "|"
                                txtBultos(4).Tag = txtBultos(4).Tag & DBLet(RS!codpobla, "T") & "|"
                                txtBultos(5).Tag = txtBultos(5).Tag & DBLet(RS!prodirec, "T") & "|"
                                txtBultos(6).Tag = txtBultos(6).Tag & "|"
                                cmbBulto.AddItem "       " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                            End If
                            RS.MoveNext
                        Wend   '
                        If cmbBulto.ListCount > 0 Then
                            cmbBulto.ListIndex = 0
                            'PonerCamposDireccionBultos 0 'Lo hace el poner a 0 el list index
                        Else
                            Reestablecer = True
                        End If
                        RS.Close
                        Set RS = Nothing

                        
                    End If
                End If
                    'La direccion
                If Reestablecer Then
                    txtClie.Text = ""
                    'Hbilitamos o no
                    cmbBulto.Clear
                    LimpiarTextosBultos
                    txtNombre(10).Text = ""
                    CliVario = False
                Else
                    
                    txtNombre(10).Text = Orden2
                End If
                HabilitarTextoCliente CliVario
                
             Screen.MousePointer = vbDefault
    
End Sub

Private Sub HabilitarTextoCliente(Habilitar As Boolean)
    If Not Habilitar Then
        txtNombre(10).BackColor = &H80000018
    Else
        txtNombre(10).BackColor = &H80000005
    End If
    txtNombre(10).Locked = Not Habilitar
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub


Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Según de donde llamemos código de una tabla u otra
        Select Case OpcionListado
            Case 1 'Listado MARCAS
                EsNomCod = True
                Tabla = "smarca"
                codCampo = "codmarca"
                nomCampo = "nommarca"
                TipCampo = "N"
                Formato = "0000"
                Titulo = "Marca"
                
            Case 2 'Listado ALMACENES Propios
                EsNomCod = True
                Tabla = "salmpr"
                codCampo = "codalmac"
                nomCampo = "nomalmac"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Almacen Propio"
                
            Case 3 'Listado Tipos UNIDADES
                EsNomCod = True
                Tabla = "sunida"
                codCampo = "codunida"
                nomCampo = "nomunida"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Tipo Unidad"
                
            Case 4 'Listado Tipos Artículos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Artículo", "T")
    
            Case 110 'Listado Ubicaciones Almacen
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "subica", "nomubica", "codubica", "Ubicaciones Almacen", "T")
            
            
            Case 20 'Listado ACTIVIDADES de Clientes
                EsNomCod = True
                Tabla = "sactiv"
                codCampo = "codactiv"
                nomCampo = "nomactiv"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Actividad de Cliente"
            
            Case 21 'Listado ZONAS de Clientes
                EsNomCod = True
                Tabla = "szonas"
                codCampo = "codzonas"
                nomCampo = "nomzonas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Zona de Cliente"
            
            Case 22 'Listado RUTAS de Asistencia
                EsNomCod = True
                Tabla = "srutas"
                codCampo = "codrutas"
                nomCampo = "nomrutas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Ruta de Asistencia"
            
            Case 23 'Listado Formas de Envío
                EsNomCod = True
                Tabla = "senvio"
                codCampo = "codenvio"
                nomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Envío"
            
            Case 24 'Listado Tarifas Venta
                EsNomCod = True
                Tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            
            Case 27 'Listado SITUACIONES Especiales
                EsNomCod = True
                Tabla = "ssitua"
                codCampo = "codsitua"
                nomCampo = "nomsitua"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Situación Especial"
            
            Case 58 'Listado PROVEEDORES
                EsNomCod = True
                Tabla = "sprove"
                codCampo = "codprove"
                nomCampo = "nomprove"
                TipCampo = "N"
                Formato = "000000"
                Titulo = "Proveedor"
            
            Case 61 'Listado MOTIVOS Pend. Rep.
                EsNomCod = True
                Tabla = "smotre"
                codCampo = "codmotre"
                nomCampo = "nommotre"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Pend. Rep."
                
            Case 65 'Listados NOTIVOS baja equipos
                EsNomCod = True
                Tabla = "smotba"
                codCampo = "codmotiv"
                nomCampo = "desmotiv"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Baja equipos"
        End Select
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 0, 86, 87
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            End If
            
        Case 5, 6, 14, 15, 29, 30, 70, 71, 90, 91, 92, 93 'Cod. ARTICULO
            EsNomCod = True
            Tabla = "sartic"
            codCampo = "codartic"
            nomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 7, 8, 16, 17, 25, 26, 62, 63, 75, 76, 88, 89, 94, 95 'Cod. FAMILIA
            EsNomCod = True
            Tabla = "sfamia"
            codCampo = "codfamia"
            nomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        
        'FECHA Desde Hasta
        Case 9, 10, 20, 22, 31, 32, 43, 44, 53, 54, 82, 83, 109, 110, 115, 116, 119, 120, 123, 124, 125, 126
            If txtCodigo(Index).Text <> "" Then
                If Index = 22 And OpcionListado = 19 Then 'Este campo sera Hora y no Fecha
                    PonerFormatoHora txtCodigo(Index)
                Else
                    PonerFormatoFecha txtCodigo(Index)
                    If OpcionListado = 223 And txtCodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then PonerFoco txtCodigo(Index)
                    End If
                End If
            End If
            
        Case 11, 12, 13, 72 'ALMACENES Propios
            EsNomCod = True
            Tabla = "salmpr"
            codCampo = "codalmac"
            nomCampo = "nomalmac"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Almacen Propio"
            
        Case 18, 19, 66, 67, 79, 80 'PROVEEDOR
            EsNomCod = True
            Tabla = "sprove"
            codCampo = "codprove"
            nomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
        
        Case 21, 96, 97, 111 'Cod. Operario/Trabajador
            EsNomCod = True
            Tabla = "straba"
            codCampo = "codtraba"
            nomCampo = "nomtraba"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Trabajador"
        
        Case 23, 24, 107
            EsNomCod = True
            TipCampo = "N"
            If OpcionListado = 30 Then 'Precios Especiales
                Tabla = "sclien"
                codCampo = "codclien"
                nomCampo = "nomclien"
                Formato = "000000"
                Titulo = "Cliente"
            Else   'Tarifas Precios
                Tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            End If
        
        Case 27, 28, 64, 65, 77, 78 'MARCAS
            EsNomCod = True
            Tabla = "smarca"
            codCampo = "codmarca"
            nomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'Nº de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el código de Oferta: " & NumCod, vbInformation
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 32, 43 'Carta de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codCampo = "codcarta"
            nomCampo = "descarta"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Cartas para Ofertas"
            
        Case 37, 38, 33, 34, 47, 48, 55, 56, 73, 74, 98, 101, 102, 103, 117, 118 'Cod. CLIENTE
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            nomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
        Case 112, 113, 114
            EsNomCod = True
            Tabla = "sincid"
            codCampo = "codincid"
            nomCampo = "nomincid"
            TipCampo = "T"
            'Formato = "0000"
            Titulo = "Incidencias"
            
        Case 39, 40, 35, 36 'Direcc./Dpto del Cliente
            If txtCodigo(Index).Text = "" Then
                txtNombre(Index).Text = ""
                Exit Sub
            End If
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            'comprobar el departamento del cliente, cuando en el campo
            'Desde/Hasta se ha seleccionado un único cliente
            If Index = 39 Or Index = 40 Then
                If txtCodigo(37).Text <> txtCodigo(38).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            ElseIf Index = 35 Or Index = 36 Then
                If txtCodigo(33).Text <> txtCodigo(34).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            End If
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            codCampo = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", txtCodigo(Index - 2).Text, "N", , "coddirec", txtCodigo(Index).Text, "N")
            txtNombre(Index).Text = codCampo 'Nombre direc. o dpto
            If codCampo = "" Then 'No existe el dpto
                If vParamAplic.Departamento Then
                    codCampo = " el Departamento "
                Else
                    codCampo = " la Dirección "
                End If
                codCampo = "No existe" & codCampo & txtCodigo(Index).Text & " para el cliente: "
                codCampo = codCampo & txtCodigo(Index - 2).Text & " - " & txtNombre(Index - 2).Text
                MsgBox codCampo, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            End If
        
        Case 41, 42, 59, 60 'Nº Contrato
'            If txtCodigo(Index).Text <> "" Then
'                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            End If

        Case 45, 46, 106, 108 'ZONAS del Cliente
            EsNomCod = True
            Tabla = "szonas"
            codCampo = "codzonas"
            nomCampo = "nomzonas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Zonas de Clientes"
        
        Case 49, 50 'Cod. AGENTE
            EsNomCod = True
            Tabla = "sagent"
            codCampo = "codagent"
            nomCampo = "nomagent"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Agente"
            
        Case 51, 52, 57, 58, 104, 105 'Tipos Contratos/MAntenimientos
            EsNomCod = True
            Tabla = "stipco"
            codCampo = "codtipco"
            nomCampo = "nomtipco"
            TipCampo = "T"
            Titulo = "Tipos de Contratos"
            
        Case 61 'Año Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un Año", vbInformation
                Exit Sub
            End If
        
        Case 68, 69 'Tipos de Articulos
            EsNomCod = True
            Tabla = "stipar"
            codCampo = "codtipar"
            nomCampo = "nomtipar"
            TipCampo = "T"
            Titulo = "Tipo de Articulo"
            
        Case 84, 85 'RUTAS del cliente
            EsNomCod = True
            Tabla = "srutas"
            codCampo = "codrutas"
            nomCampo = "nomrutas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
            
        Case 121, 122 'Nº Factura
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
            End If
        End Select
    End If
    
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


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    Conexion = conAri    'Conexión a BD: Ariges
    Select Case OpcionListado
        Case 7 'Traspaso de Almacenes
            Cad = Cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
            Tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            Cad = Cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
            Tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            Cad = Cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
            Tabla = "sartic"
            Titulo = "Articulos"
    End Select
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case OpcionListado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(IndCodigo)
            Case 9, 12, 13, 14, 15, 16, 17 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(IndCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(IndCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(IndCodigo)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 5535
    W = 6555
    PonerFrameVisible Me.frameListado, visible, H, W

    If visible = True Then
        Me.Optcodigo.Value = True
    End If
End Sub


Private Sub PonerFrameInventarioVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Inventario Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Inventario
Dim VerOpcion As Boolean
       
    If visible = True Then
        H = 6400
        W = 7995
        VerOpcion = (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19)
        
        If VerOpcion Then
            H = 6900
            Me.cmdAceptar(4).Top = 6360
            Me.cmdCancel(4).Top = 6360
        ElseIf OpcionListado = 13 Then
            H = 6000
            Me.cmdAceptar(4).Top = 5200
            Me.cmdCancel(4).Top = Me.cmdAceptar(4).Top
        End If
        PonerFrameVisible Me.FrameInventario, visible, H, W

                
        '======================================
        'Valorar con Precios
        If VerOpcion Then
            Me.FrameValorar.visible = VerOpcion
            Me.FrameValorar.Left = 600
            If OpcionListado = 17 Then
                Me.FrameValorar.Top = 4500
            Else
                Me.FrameValorar.Top = 5000
            End If
            Me.chkSinStock.visible = VerOpcion
        End If
        
                 
                
        '====================================
        'Poner el Trabajador
        VerOpcion = (OpcionListado = 14)
        Me.Label4(7).visible = VerOpcion
        Me.imgBuscarG(17).visible = VerOpcion
        Me.txtCodigo(21).visible = VerOpcion
        Me.txtNombre(21).visible = VerOpcion
'        If VerOpcion Then txtCodigo(21).TabIndex = 47
        
        
        '======================================
        'Fecha Listados
        If OpcionListado = 15 Then '15: Listado Articulos Inactivos
            Me.Label4(5).Caption = "Fecha Inactividad"
        ElseIf OpcionListado = 19 Then
            Me.Label4(5).Caption = "Fecha Stock"
        Else
            Me.Label4(5).Caption = "Fecha Inventario"
        End If
        
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 19)
        Me.Label4(5).visible = VerOpcion  'campo fecha
        Me.imgFecha(2).visible = VerOpcion
        Me.txtCodigo(20).visible = VerOpcion
        'campo HAsta Fecha
        Me.Label4(8).visible = (OpcionListado = 16)
        'Si opcionlistado=19 este campo sera la hora
        Me.Label4(9).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        If OpcionListado = 19 Then
            Me.Label4(9).Caption = "Hora"
            Me.Label4(9).Left = 4250
            Me.txtCodigo(22).Left = 4700
        End If
        Me.imgFecha(3).visible = (OpcionListado = 16)
        Me.txtCodigo(22).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        If OpcionListado = 16 Then
            Me.Label4(8).Left = 2280
            Me.imgFecha(2).Left = 2820
            Me.txtCodigo(20).Left = 3120
            Me.Label4(9).Left = 4680
            Me.imgFecha(3).Left = 5160
            Me.txtCodigo(22).Left = 5430
'            txtCodigo(22).TabIndex = 48
        End If
        
        
        '====================================
        'Activar o no los check de Opcion:
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 13) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Or OpcionListado = 15
                    '12: Toma de Inventario
                    '13: Listado Diferencias stock
        
        Me.FrameOpciones.visible = VerOpcion
        Me.FrameOpciones.Top = 5000
        If OpcionListado = 13 Then
            Me.FrameOpciones.Top = 4500
            Me.FrameOpciones.BorderStyle = 0
        End If
        Me.FrameOpciones.Height = 1000

        Me.chkSaltaPag.visible = VerOpcion
        Me.chkValorado.visible = (OpcionListado = 16) Or (OpcionListado = 17)

        
        VerOpcion = (OpcionListado = 12)
        If VerOpcion Or OpcionListado = 13 Then Me.FrameOpciones.Left = 700
        Me.chkImprimeStock.visible = VerOpcion
        Me.chkImprimeStock.Top = 600
        If VerOpcion Then Me.txtCodigo(20).Text = Date
    End If
End Sub



Private Sub PonerFrameTarifasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Tarifas Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Tarifas
Dim VerOpcion As Boolean

    H = 6375
    If OpcionListado = 245 Then H = 5675
    W = 7635
    PonerFrameVisible Me.FrameTarifas, visible, H, W
    
    If visible = True Then
        '====================================
        '28: Tarifas Precios 29: Promociones
        VerOpcion = (OpcionListado = 28) Or (OpcionListado = 29)
        Me.chkSaltaPagTarif.visible = VerOpcion
        Me.Label4(12).visible = VerOpcion
        
        '====================================
        If OpcionListado = 30 Then Me.Label4(11).Caption = "Cliente"
        
        
        '245: Control margenes tarifas
        '==================================
        VerOpcion = (OpcionListado = 245)
        Me.chkMostrarErrores.visible = VerOpcion
        'Decimales
        Me.cboDecimales.visible = VerOpcion
        Label4(88).visible = VerOpcion
        If VerOpcion Then
            Me.chkMostrarErrores.Top = 4600
            Label4(88).Top = 4300
            cboDecimales.Top = 4600
            
            'no mostrar seleccion de marca D/H
            Me.Label4(13).visible = Not VerOpcion
            Me.Label3(13).visible = Not VerOpcion
            Me.Label3(14).visible = Not VerOpcion
            Me.imgBuscarG(59).visible = Not VerOpcion
            Me.imgBuscarG(60).visible = Not VerOpcion
            Me.txtCodigo(27).visible = Not VerOpcion
            Me.txtCodigo(28).visible = Not VerOpcion
            Me.txtNombre(27).visible = Not VerOpcion
            Me.txtNombre(28).visible = Not VerOpcion
            'subir seleccion Articulo D/H al sitio de la marca
            Me.Label4(14).Top = Me.Label4(13).Top
            Me.Label3(15).Top = Me.Label3(13).Top
            Me.Label3(16).Top = Me.Label3(14).Top
            Me.imgBuscarG(61).Top = Me.imgBuscarG(59).Top
            Me.imgBuscarG(62).Top = Me.imgBuscarG(60).Top
            Me.txtCodigo(29).Top = Me.txtCodigo(27).Top
            Me.txtCodigo(30).Top = Me.txtCodigo(28).Top
            Me.txtNombre(29).Top = Me.txtNombre(27).Top
            Me.txtNombre(30).Top = Me.txtNombre(28).Top
            Me.cmdAceptarTarif.Top = 4600
            Me.cmdCancel(indFrame).Top = Me.cmdAceptarTarif.Top
        End If
    End If
End Sub


Private Sub PonerFrameRepxDiaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de las Reparaciones x dia, de tabla: scarep
    

    If OpcionListado = 223 Or OpcionListado = 224 Then
        H = 4400
        W = 6100
    Else
        H = 3500
        W = 6000
    End If
    
    PonerFrameVisible Me.FrameRepxDia, visible, H, W
    
    If visible = True Then
        Me.Caption = "AriGes"
'        Me.FrameContab.Enabled = False
'        Me.OptClientes.Enabled = False
        Me.FrameContab.visible = (OpcionListado = 223 Or OpcionListado = 224 Or OpcionListado = 248)
        Me.FrameTipMov.visible = (OpcionListado = 223)
        Me.FrameProgress.visible = False
        
        '-- alto del boton aceptar y cancelar
        
        If OpcionListado = 223 Or OpcionListado = 224 Then
            Me.cmdAceptarRepxDia.Top = 3800
        Else
            Me.cmdAceptarRepxDia.Top = 2800
        End If
        Me.cmdCancel(7).Top = Me.cmdAceptarRepxDia.Top
        
        Select Case OpcionListado
            Case 63
                Me.lblTitulo(0).Caption = "Reparaciones por Día"
                Me.Label2(2).Caption = "Fecha Reparación:"
                Frame2.Top = 1350
            Case 73
                Me.lblTitulo(0).Caption = "Altas de Mantenimientos"
                Me.Label2(2).Caption = "Fecha Mantenimiento:"
                Frame2.Top = 1350
            Case 223, 224, 248 'Pedir datos para contabilizar facturas
                Me.lblTitulo(0).Caption = "Contabilizar Facturas"
                Me.Label2(2).Caption = "Fecha Factura:"
                Frame2.Top = 1680
                Me.FrameTipMov.Top = 2650
                
                
                Me.OptProve.Tag = ""
                If OpcionListado = 248 Then
                    Me.OptProve.Tag = "TIK"  'Son las de tickets agrupados
                    OpcionListado = 223
                End If
                If OpcionListado = 224 Then
                    Me.OptProve.Value = True
                    OpcionListado = 223
                End If
        End Select
    End If
End Sub


Private Sub PonerFrameManteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los Mantenimientos, de tabla: scaman
Dim B As Boolean
        'Opciones: 70,71,78,79,76
    H = 6975 '- 6375
    W = 6875
    PonerFrameVisible Me.FrameMantenimientos, visible, H, W

    If visible = True Then
        B = (OpcionListado = 70)
        
        Me.cboTipoList.visible = B 'List. Mantenimientos
        Me.Label1(4).visible = B
        
        
        
        'List Revisiones Mantenimientos
        Me.Frame3(1).visible = (OpcionListado = 70) Or (OpcionListado = 76)
        Me.Frame3(0).visible = (OpcionListado = 71)
        Me.Frame3(2).visible = (OpcionListado = 78)
        
        Select Case OpcionListado
        Case 70
                Me.Label7(0).Caption = "Informe de Mantenimientos"
        Case 71
                Me.Label7(0).Caption = "Informe Revisiones Mantenimientos"
               ' Me.Frame3.Top = 4800
                Me.txtCodigo(53).TabIndex = 211
                Me.txtCodigo(54).TabIndex = 212
                
        Case 76
                Me.Label7(0).Caption = "Inf.  Mantenimientos ANULADOS"
                
                
        Case 78
                'Cartas de renovacvion
                Me.Label7(0).Caption = "Cartas de renovacion"
                Me.txtCodigo(109).Text = Format(Now, "dd/mm/yyyy")
        Case 79
                'Etiquetas de mantenimientos
                Me.Label7(0).Caption = "Etiquetas de mantenimientos"
        End Select
    End If
End Sub


Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim B As Boolean



    'Hay una opcion mas que mostrara este frame. la 247. Correccion de de tarifas e importes en articulos
    FrameTapaINCORRECTO.visible = False
    chkMinimoCorreg.visible = False
    chkPreciosProvee.visible = (OpcionListado = 510)
    B = (OpcionListado = 6)
    If B Then
        Me.Label9.Caption = "Informe de Articulos"
       
        W = 8595
    Else
        If OpcionListado = 18 Then
            Me.Label9.Caption = "Informe Stocks Maximos y Minimos"
            Label4(36).Caption = "Almacén"
        Else
            'NUEVA OCPION:  247 y 510(AVAB)
            'Corregir tarifas y eso
            chkMinimoCorreg.visible = True
            If OpcionListado = 247 Then
                Me.Label9.Caption = "Verificación tarifas y P.V.P."
            Else
                Me.Label9.Caption = "Verificar precios (AVAB/Mor)."
            End If
            FrameTapaINCORRECTO.visible = True
            Label4(36).Caption = "Tarifa"
            cmbDecimales.ListIndex = 1
        End If
        W = 7395
       
    End If
    H = 6820
    
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = B
        Label4(36).visible = Not B

        Me.imgBuscarG(18).visible = Not B
        Me.txtCodigo(72).visible = Not B
        Me.txtNombre(72).visible = Not B
        
        'Trifa NO visible para 510
        Me.imgBuscarG(87).visible = OpcionListado <> 510
        Me.txtCodigo(107).visible = OpcionListado <> 510
        Me.txtNombre(107).visible = OpcionListado <> 510
        Label4(36).visible = OpcionListado <> 510
        
        
        'Visible Frame stocks Max Minimos si opcionlistado=18
        Me.optStockMax.Value = True
        Me.FrameStockMaxMin.visible = OpcionListado = 18
  
        FrameSituacionArticulo.visible = OpcionListado = 6
    
    
        'REajustes.
        'El articulo NO se muestra si la opcion es 247
        B = OpcionListado <> 247
        PonerLabelsArticulosFrameVisible B
        Label4(75).visible = Not B
        cmbDecimales.visible = Not B
        Label4(90).visible = Not B
        cmbDecimales.visible = Not B
    
    End If
End Sub


Private Sub PonerLabelsArticulosFrameVisible(Si As Boolean)
    Label4(38).visible = Si
    Label3(51).visible = Si
    imgBuscarG(27).visible = Si
    txtCodigo(70).visible = Si
    txtNombre(70).visible = Si
    Label3(54).visible = Si
    imgBuscarG(28).visible = Si
    txtCodigo(71).visible = Si
    txtNombre(71).visible = Si
    chkMinimoCorreg.visible = Not Si
    
End Sub


Private Sub CargarListView()
'Carga el List View del frame: frameMovimArtic
'con los parametros de la tabla: stipom (Tipos de Movimientos)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 800
    ListView1.ColumnHeaders.Add , , "Descripción", 2250
    
    SQL = "select * from stipom where muevesto=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value
        ItmX.Checked = True
        ItmX.SubItems(1) = RS.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    
'    'MARZO 2009
'    'Esta comentado, pq FALTA### ver si añadimos y eso
'    'Este lo añado pq el movimiento no tiene la marca de MUEVE stock
'    'Pero si ke lo mueve , es la produccion
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "PRO"
    ItmX.Checked = True
    ItmX.SubItems(1) = "PRODUCCION"
'
'
'
    Set RS = Nothing
End Sub



Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Marca"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim I As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    NumParam = 1
        
        
    'Añadiremos a la formula
    If vUsu.Nivel > 0 Then
        If Not vUsu.TrabajadorB Then
            cadFormula = "{smoval.codalmac} <> 2"
            cadSelect = "smoval.codalmac <> 2"
        End If
    End If
    '-- Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(5).Text <> "" Or txtCodigo(6).Text <> "" Then
        Codigo = "{smoval.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "T", 5, 6, devuelve) Then Exit Function
    End If
                    
    '-- Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(7).Text <> "" Or txtCodigo(8).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 7, 8, devuelve) Then Exit Function
    End If
        
    '-- Cadena para seleccion Desde y Hasta ALMACEN
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        Codigo = "{smoval.codalmac}"
        devuelve = "pDHAlmacen=""Almacen: "
        If Not PonerDesdeHasta(Codigo, "N", 11, 12, devuelve) Then Exit Function
    End If
    
    
    '-- Cadena para seleccion Desde y Hasta CLIENTE/PROVEEDOR
    If txtCodigo(86).Text <> "" Or txtCodigo(87).Text <> "" Then
        Codigo = "{smoval.codigope}"
        devuelve = "pDHOperario=""Cliente/Proveedor/Trab.: "
        If Not PonerDesdeHasta(Codigo, "N", 86, 87, devuelve) Then Exit Function
    End If
    
        
'    cadSelect = QuitarCaracterACadena(cadFormula, "{")
'    cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
    '=================================================
    '-- Cadena para seleccion Desde y Hasta FECHA
    If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
        Codigo = "{smoval.fechamov}"
        devuelve = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 9, 10, devuelve) Then Exit Function
    End If
        
    '-- seleccionar los articulos que tienen control de stock
    Codigo = "{sartic.ctrstock}=1"
    AnyadirAFormula cadFormula, Codigo
    AnyadirAFormula cadSelect, Codigo
        
        
    '-- Cadena de Seleccion TIPOS de MOVIMIENTOS
    Codigo = "{smoval.detamovi}"
    devuelve = ""
    'Si todos seleccionados no añadir la select
    todosMarcados = True
    I = 1
    While Not I > Me.ListView1.ListItems.Count And todosMarcados
        If Not Me.ListView1.ListItems(I).Checked Then todosMarcados = False
        I = I + 1
    Wend
    
    'si no estan todos seleccionados montar select de los seleccionados
    If Not todosMarcados Then
        Cad = ""
        devuelve = ""
        For I = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(I).Checked Then
                If Cad = "" Then
                    Cad = Me.ListView1.ListItems(I).Text
                Else
                    Cad = Cad & ", " & Me.ListView1.ListItems(I).Text
                End If
                If devuelve = "" Then
                    devuelve = Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                Else
                    devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                End If
            End If
        Next I

        If devuelve <> "" Then 'Hay algun movimiento marcado
            If cadFormula <> "" Then
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = cadSelect & " AND " & "(" & devuelve & ")"
                cadParam = cadParam
            Else
                cadFormula = "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = "(" & devuelve & ")"
            End If
            Cad = "pTiposMov=""Tipos Movimiento: " & Cad
            cadParam = cadParam & Cad & """|"
            NumParam = NumParam + 1
        Else 'Todos desmarcados
            Cad = ""
            For I = 1 To ListView1.ListItems.Count
                If Cad = "" Then
                    Cad = """" & ListView1.ListItems(I).Text & """"
                Else
                    Cad = Cad & ", """ & ListView1.ListItems(I).Text & """"
                End If
            Next I
            devuelve = Codigo & " NOT IN [" & Cad & "]"
            Cad = Codigo & " NOT IN (" & Cad & ")"
            Cad = QuitarCaracterACadena(Cad, "{")
            Cad = QuitarCaracterACadena(Cad, "}")
            If cadFormula = "" Then
                cadFormula = "(" & devuelve & ")"
                cadSelect = "(" & Cad & ")"
            Else
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
                cadSelect = cadSelect & " AND " & "(" & Cad & ")"
            End If
        End If
    End If
    
    
    If cadFormula = "" Then
        MsgBox "Introduzca algún criterio de selección para el Informe.", vbInformation
        Exit Function
    End If
    PonerFormulaYParametrosInf9 = True
    
End Function


Private Function PonerFormulaYParametrosInf12() As Boolean
Dim Cad As String, cadFrom As String
Dim devuelve As String
Dim ImprStock As String
Dim CodAux As String
Dim strValorado As String
Dim strSinStock As String
Dim bytPrecio As Byte

'    InicializarVbles
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    NumParam = NumParam + 1
    cadFrom = ""
    devuelve = ""
    PonerFormulaYParametrosInf12 = False
    
    '===================================================
    '================= FORMULA =========================
    
    Select Case OpcionListado
        Case 12, 15, 16, 17, 19
            CodAux = "{salmac."
            cadFrom = "  salmac "
'        Case 15 'Listado articulos inactivos
'            CodAux = "{salmac."
'            cadFrom = "  (salmac LEFT OUTER JOIN smoval ON salmac.codartic=smoval.codartic AND salmac.codalmac=smoval.codalmac) "
'            cadFrom = "salmac"
        Case 13, 14
            CodAux = "{sinven."
            cadFrom = " sinven "
    End Select
    
    'Cadena para seleccion De ALMACEN
    '-----------------------------------
    Codigo = CodAux & "codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = Codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        Cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        cadParam = cadParam & Cad & """|"
        NumParam = NumParam + 1
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = CodAux & "codartic}"
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(Codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codfamia}"
            Case Else: Codigo = "{sinven.codfamia}"
        End Select
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, Cad) Then Exit Function
    End If
    cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
    'Enero 2008
    'David
    cadFormula = cadFormula & " AND {sartic.ctrstock} = 1"
    
    'Enero 2009
    'David
    'Solo saldran los articulos que esten en situacion normal o bloqueados.
    'Los caducados NO salen
    cadFormula = cadFormula & " AND {sartic.codstatu} < 2"
    
    
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '----------------------------------------------
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codprove}"
            Case Else: Codigo = "{sinven.codprove}"
        End Select
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, Cad) Then Exit Function
    End If
    

    
    'Select para MySQL
    cadSelect = QuitarCaracterACadena(cadFormula, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    cadFrom = QuitarCaracterACadena(cadFrom, "{")
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If (OpcionListado = 16) Then
        If txtCodigo(20).Text <> "" Or txtCodigo(22).Text <> "" Then
            'codigo = "{salmac.codartic}"
            Codigo = CodAux & "fechainv}"
            devuelve = CadenaDesdeHasta(txtCodigo(20).Text, txtCodigo(22).Text, Codigo, "F")
    
            If devuelve = "Error" Then Exit Function
            
            If Not AnyadirAFormula(cadFormula, devuelve) Then
                Exit Function
            ElseIf devuelve <> "" Then
                Cad = "pDHFecha=""Fecha: "
                If txtCodigo(20).Text <> "" Then _
                    Cad = Cad & "desde " & txtCodigo(20).Text
                If txtCodigo(22).Text <> "" Then _
                    Cad = Cad & "  hasta " & txtCodigo(22).Text
                cadParam = cadParam & Cad & """|"
                NumParam = NumParam + 1
                'Para Comprobar si hay registros a Mostrar antes de abrir el Informe
                devuelve = "salmac.fechainv"
                devuelve = CadenaDesdeHastaBD(txtCodigo(20).Text, txtCodigo(22).Text, devuelve, "F")
                AnyadirAFormula cadSelect, devuelve
            Else
                'Si no hay fecha de inventario seleccionada coger solo
                'los articulos de los que se haya hecho inventario alguna vez
                devuelve = "not isnull({salmac.fechainv})"
                If Not AnyadirAFormula(cadFormula, devuelve) Then
                    Exit Function
                End If
                devuelve = "not isnull(salmac.fechainv)"
                AnyadirAFormula cadSelect, devuelve
            End If
        End If
    End If
    
    'Cadena de seleccion de FECHA de Inactividad
    '------------------------------------------------
    If OpcionListado = 15 Then '15: Listado de Articulos Inactivos
         If txtCodigo(20).Text <> "" Then _
            Cad = "pFechaInve=""" & txtCodigo(20).Text & """"
        
        'Poner en el parametro pListaArt la lista de Articulos que no tiene
        'un registro de movimiento en la smoval con fecha posterior a la
        'fecha de inactividad
        strValorado = ListaArtActivos(cadSelect, txtCodigo(20).Text)
        Cad = "pListaArtic=""" & strValorado & """|"
        cadParam = cadParam & Cad
        NumParam = NumParam + 1
        
        'Añadir a la formula de seleccion que no sea uno de la lista
        devuelve = " not (" & CodAux & "codartic} in {@pListaArtic})"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
        
        strValorado = QuitarCaracterACadena(strValorado, "[")
        strValorado = QuitarCaracterACadena(strValorado, "]")
        devuelve = " not (salmac.codartic in (" & strValorado & "))"
        AnyadirAFormula cadSelect, devuelve
    End If
    
    'Cadena de seleccion de FECHA de Stocks a una Fecha
    '--------------------------------------------------
     If OpcionListado = 19 Then
        If txtCodigo(20).Text <> "" Then
            Cad = txtCodigo(20).Text
            'Hora
            If txtCodigo(22).Text <> "" Then _
                Cad = Cad & "  " & txtCodigo(22).Text
                
            cadParam = cadParam & "pFechaStock=""" & Cad & """|"
            NumParam = NumParam + 1
        End If
     End If
     
    'Cadena para Seleccion de Articulos con Stock<>0
    '------------------------------------------------
    If OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 15 Then
        If Me.chkSinStock.Value = 0 Then
            If OpcionListado = 16 Then
                devuelve = "{salmac.stockinv}<>0"
            Else
                devuelve = CodAux & "canstock}<>0"
            End If
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
            
            devuelve = QuitarCaracterACadena(devuelve, "{")
            devuelve = QuitarCaracterACadena(devuelve, "}")
            devuelve = QuitarCaracterACadena(devuelve, "_1")
            AnyadirAFormula cadSelect, devuelve
        End If
    ElseIf OpcionListado = 19 Then
         If Me.chkSinStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSinStock=" & ImprStock & "|"
        NumParam = NumParam + 1
    End If
    
       
    '============================================
    '============= PARAMETROS ===================
    If OpcionListado = 12 Or OpcionListado = 15 Then
        '12: Toma de Inventario
        '15: Listado Articulos Inactivos
        cadParam = cadParam & "pFechaInve=""" & txtCodigo(20).Text & """|"
        NumParam = NumParam + 1
    End If
    
    If OpcionListado = 12 Then
        'Parámetro Imprime Stock (Si/No)
        If Me.chkImprimeStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pImprimeStock=" & ImprStock & "|"
        NumParam = NumParam + 1
        
'        'seleccionar para inventariar los articulos que no tienen control stock
'        devuelve = " {sartic.ctrstock} = 1 "
'        AnyadirAFormula cadFormula, devuelve
'        AnyadirAFormula cadSelect, devuelve
        'Laura 03/01/07
        If Not (InStr(cadFrom, "sartic") > 0) Then
            cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
        End If
    End If
    
    If OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 15 Or OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 19 Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPag.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & ImprStock & "|"
        NumParam = NumParam + 1
    End If
    
    If OpcionListado = 16 Or OpcionListado = 17 Then '16: Valoración de Stocks Inventariados
                                                     '17: Valoración Stocks
        'Parámetro Valorado
        If Me.chkValorado.Value Then
            strValorado = "True"
        Else
            strValorado = "False"
        End If
        cadParam = cadParam & "pValorado=" & strValorado & "|"
        NumParam = NumParam + 1
    End If
    
    If (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Then
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        NumParam = NumParam + 1
    End If
    '=====================================================================
    
       
    'comprobar si hay registros para mostrar en el Informe antes de Abrirlo
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Function
    
    If OpcionListado = 19 Then
        cadSelect = "Select count(*) FROM " & cadFrom & " WHERE " & cadSelect
        cadSelect = Replace(cadSelect, "count(*)", "*")
        DescargarDatosTMPStockFecha
        If Not CargarTMPStockFecha(cadSelect, txtCodigo(20).Text, txtCodigo(22).Text) Then Exit Function
        
        
        'Si es 19 (stock a una fecha)
        ' y esta nmarcada la opcion de aceites, borraremos updatearemos la cantidad por los litrosunidad
        If Me.chkStockFechaAceite.Value = 1 Then
            DoEvents
            'Para cada
            AjustesStocksFechaAceite
        
        End If
    End If
    
    PonerFormulaYParametrosInf12 = True
End Function



Private Function PonerFormulaYParametrosInf28() As Boolean
'Informes de Descuentos y Tarifas
Dim Cad As String
Dim cadCodigo As String

    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    NumParam = 1
    
    PonerFormulaYParametrosInf28 = False
    
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Desde y Hasta TARIFA o D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(23).Text <> "" Or txtCodigo(24).Text <> "" Then
        If OpcionListado = 30 Then 'Precios Especiales Cliente
            cadCodigo = Codigo & ".codclien}"
            Cad = "pDHCliente=""Cliente: "
        Else
            cadCodigo = Codigo & ".codlista}"
            Cad = "pDHTarifa=""Tarifa: "
        End If
        If Not PonerDesdeHasta(cadCodigo, "N", 23, 24, Cad) Then Exit Function
    End If
            
            
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(25).Text <> "" Or txtCodigo(26).Text <> "" Then
        cadCodigo = "{sartic.codfamia}"
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(cadCodigo, "N", 25, 26, Cad) Then Exit Function
    End If
    
    If OpcionListado <> 245 Then
        'Cadena para seleccion Desde y Hasta MARCA
        '--------------------------------------------
        If txtCodigo(27).Text <> "" Or txtCodigo(28).Text <> "" Then
            cadCodigo = "{sartic.codmarca}"
            Cad = "pDHMarca=""Marca: "
            If Not PonerDesdeHasta(cadCodigo, "N", 27, 28, Cad) Then Exit Function
        End If
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        cadCodigo = Codigo & ".codartic}"
        Cad = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(cadCodigo, "T", 29, 30, Cad) Then Exit Function
    End If
 
 
    '=====================================================================
    '====   PARAMETROS
    If (OpcionListado = 28 Or OpcionListado = 29) Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPagTarif.Value = 1 Then
            Cad = "True"
        Else
           Cad = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & Cad & "|"
        NumParam = NumParam + 1
    End If
       
    If OpcionListado = 245 Then
        'Parámetro mostrar solo tarifas con errores (Si/No)
        Cad = Abs(Val(Me.chkMostrarErrores.Value))
        cadParam = cadParam & "Suprimr=" & Cad & "|"
        NumParam = NumParam + 1
        'Decimales
        If cboDecimales.ListIndex < 0 Then
            MsgBox "Seleccione decimales", vbExclamation
            Exit Function
        End If
        Cad = (cboDecimales.ItemData(Me.cboDecimales.ListIndex))
        cadParam = cadParam & "Decimales=" & Cad & "|"
        NumParam = NumParam + 1
    End If
       
    PonerFormulaYParametrosInf28 = True
End Function


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function InsertarInventario() As Boolean
'Inserta en la Tabla:sinven los articulos seleccionados para realizar Inventario
'Inserta en la Tabla Hist.: shinve los datos que habia de inventario
'Además Actualiza la Tabla:salmac los campos:fechainv, horainve, statusin
Dim SQL As String, ADonde As String
Dim RS As ADODB.Recordset
Dim hora As Date
Dim CantidadI As Currency

On Error GoTo EInventario:
   
'   If CrearTmpInventario(cadSelect) Then
   

        'Aqui empieza transaccion
        Conn.BeginTrans
    
          
    
'        'Insertar en la tabla de Histórico: shinve
'        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
'        ADonde = "Insertando datos en Histórico. Tabla: shinve"
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & " SELECT salmac.codartic, salmac.codalmac, salmac.fechainv,salmac.horainve,salmac.stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'si no se ha inventariado antes no lo pasamos al historico
'        SQL = SQL & " AND not isnull(salmac.fechainv) "
'        Conn.Execute SQL
'
        
        
        
        'Insertar en la tabla de Histórico: shinve
        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
        ADonde = "Insertando datos en Histórico. Tabla: shinve"
        'Enero 2009 preciomp, precioma ,preciouc, preciost
        'añadimos los campos
        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc,movpost,preciomp, precioma ,preciouc, preciost) "
        SQL = SQL & " SELECT tmpInven.codartic,tmpInven.codalmac,tmpInven.fechainv,tmpInven.horainve,tmpInven.stockinv,tmpInven.movpost,preciomp, precioma ,preciouc, preciost "
        SQL = SQL & " FROM tmpInven LEFT JOIN salmac ON tmpInven.codartic=salmac.codartic "
        SQL = SQL & " AND  tmpInven.codalmac=salmac.codalmac "
        SQL = SQL & " WHERE not isnull(tmpInven.fechainv) "
        '--- Laura 03/01/2006
        SQL = SQL & " AND tmpInven.fechainv<>'0000-00-00' AND date(tmpInven.horainve)<>'0000-00-00' "
        '---
        Conn.Execute SQL

        'ANTES
        'hora = Format(txtCodigo(20).Text & " " & Time, "yyyy-mm-dd hh:mm:ss")
        
        'Ahora la hora sera el ultimo voimiento del dia
        hora = Format(txtCodigo(20).Text & " 23:59:59", "yyyy-mm-dd hh:mm:ss")
        
'        'Insertamos en la Tabla sinven
'        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
'        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
'        SQL = SQL & "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
'        Conn.Execute SQL


        
        'Insertamos en la Tabla sinven
        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc,movpost) "
        SQL = SQL & "SELECT codartic, codalmac, codfamia, codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc,movpost "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        SQL = SQL & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
        Conn.Execute SQL


        
        
'        SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove "
'        SQL = SQL & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
        
        'Enero 10.  Para meter los precios actuales AHORA
        'SQL = "SELECT codartic, codalmac, codfamia, codprove, movpost"
        'SQL = SQL & " FROM tmpInven"
        'ahora
        SQL = "SELECT tmpInven.codartic, tmpInven.codalmac, tmpInven.codfamia, tmpInven.codprove, tmpInven.movpost"
        SQL = SQL & ",preciomp, precioma, preciost, preciouc"
        SQL = SQL & " FROM tmpInven,sartic where tmpinven.codartic=sartic.codartic"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
            
        
        
    '        'Insertamos en la Tabla sinven
    '        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
    '        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
    '        SQL = SQL & " VALUES (" & DBSet(Rs.Fields(0).Value, "T") & ", " & Rs.Fields(1).Value & ", "
    '        SQL = SQL & Rs.Fields(2).Value & ", " & Rs.Fields(3).Value & ", '"
    '        'SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & hora & "', " & rs.Fields(2).Value & ")"
    '        SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', 0)"
    '        Conn.Execute SQL
            
            'Nuevo
            'Diciembre 2009
            ' Ponemos como cantidad inventariada la cantidad k hay en stock, con lo cual
            ' si no hay diferencias no tendra que teclear nada
            ADonde = "Actualizando datos sinven"
            SQL = "codartic=" & DBSet(RS.Fields(0).Value, "T") & " AND codalmac"
            SQL = DevuelveDesdeBD(conAri, "canstock", "salmac", SQL, CStr(RS.Fields(1)), "N")
            If SQL = "" Then SQL = "0"
            CantidadI = CCur(SQL) - DBLet(RS!movpost, "N")


            SQL = "UPDATE sinven SET existenc = " & TransformaComasPuntos(CStr(CantidadI))
            
            'ENERO 2010 preciomp precioma preciost preciouc
           ' SQL = SQL & "preciomp, precioma, preciost, preciouc)"
            SQL = SQL & ", preciomp = " & DBSet(RS!preciomp, "N")
            SQL = SQL & ", precioma = " & DBSet(RS!precioma, "N")
            SQL = SQL & ", preciost = " & DBSet(RS!preciost, "N")
            SQL = SQL & ", preciouc = " & DBSet(RS!precioUC, "N")
            
            SQL = SQL & " WHERE codartic=" & DBSet(RS.Fields(0).Value, "T") & " AND "
            SQL = SQL & "codalmac=" & RS.Fields(1).Value
            Conn.Execute SQL
            
            
            
            'Actualizamos la tabla salmac ponemos statusin=1 para indicar que se
            'esta realizando inventario y bloquear los articulos para que no se puedan
            'realizar movimientos, traspasos, etc.
            'Actualizamos la Tabla: salmac los campos: fechainv, horainve
            ADonde = "Actualizando datos en Articulos x Almacen"
            SQL = "UPDATE salmac SET fechainv='" & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', "
            SQL = SQL & " horainve='" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', " & "statusin=1 , stockinv=" & TransformaComasPuntos(CStr(CantidadI))
            'Enero 2010  guardr precios
            SQL = SQL & ", preciomp = " & DBSet(RS!preciomp, "N")
            SQL = SQL & ", precioma = " & DBSet(RS!precioma, "N")
            SQL = SQL & ", preciost = " & DBSet(RS!preciost, "N")
            SQL = SQL & ", preciouc = " & DBSet(RS!precioUC, "N")
            
            SQL = SQL & " WHERE codartic=" & DBSet(RS.Fields(0).Value, "T") & " AND "
            SQL = SQL & "codalmac=" & RS.Fields(1).Value
            Conn.Execute SQL
            RS.MoveNext
        Wend
    
        RS.Close
        Set RS = Nothing
'    Else
'        Exit Function
'    End If
    
EInventario:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
          SQL = "Insertando Datos de Inventario." & vbCrLf & "--------------------------------------" & vbCrLf
          SQL = SQL & ADonde
          MuestraError Err.Number, SQL, Err.Description
        Conn.RollbackTrans
        InsertarInventario = False
    Else
        InsertarInventario = True
        Conn.CommitTrans
    End If
End Function


Private Function CrearTmpInventario(cadFormula As String) As Boolean
Dim SQL As String
Dim B As Boolean

    On Error GoTo ECrearInv
    
    B = False
    SQL = "CREATE TEMPORARY TABLE tmpInven ( "
    SQL = SQL & "codartic varchar(16) NOT NULL default '0', "
    SQL = SQL & "codalmac smallint(3) unsigned NOT NULL default '0', "
    SQL = SQL & "codfamia smallint(4) unsigned NOT NULL default '0', "
    SQL = SQL & "codprove int(6) unsigned NOT NULL default '0', "
    SQL = SQL & "fechainv date NOT NULL default '0000-00-00', "
    SQL = SQL & "horainve datetime NOT NULL default '0000-00-00 00:00:00', "
    SQL = SQL & "stockinv decimal(12,2) NOT NULL default '0.00',"
    SQL = SQL & "movpost decimal(12,2) NOT NULL default '0.00'"
    SQL = SQL & ")"
    Conn.Execute SQL
    B = True
    
    
    'Seleccionar todos los registros que vamos a inventariar, insertarlos en la TMP
    'y trabajar con estos valores
    
    If AntiguaFormaInventariar Then
        SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove,salmac.fechainv,salmac.horainve,salmac.stockinv,0  "
        SQL = SQL & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        SQL = SQL & " WHERE " & cadFormula
        SQL = SQL & " AND sartic.ctrstock=1"
    
    
        'NUEVA FORMA INVENTARIAR
        
    
    
        SQL = " INSERT INTO tmpInven " & SQL
        
    Else
        'NUevao. Con movimientos posterior
        'SQL = " INSERT INTO tmpInven(codartic,codalmac fechainv horainve existenc, movpost)"
        SQL = " INSERT INTO tmpInven "
        SQL = SQL & " select tmptominventario.codartic,tmptominventario.codalmac,sartic.codfamia,sartic.codprove,"
        SQL = SQL & " salmac.fechainv,salmac.horainve,salmac.stockinv,movpost"
        SQL = SQL & " from tmptominventario,sartic,salmac"
        SQL = SQL & " where  tmptominventario.codartic=sartic.codartic and tmptominventario.codartic=salmac.codartic"
        SQL = SQL & " and tmptominventario.codalmac=salmac.codalmac  and sartic.ctrstock=1 and"
        SQL = SQL & " codusu= " & vUsu.Codigo
        
    End If
    Conn.Execute SQL
        
    
    
ECrearInv:
    If Err.Number <> 0 Then
        SQL = " DROP TABLE IF EXISTS tmpInven;"
        Conn.Execute SQL
        B = False
        'Err.Clear
        MuestraError Err.Number, "Crear temporal inventario.", Err.Description
    End If
    CrearTmpInventario = B
End Function






Private Function ActualizarInventario() As Boolean
'-----------------------------------------------------------------
'* Modifica en la Tabla: salmac los campos: cansotck, fechainv, horainve,statusin de los articulos seleccionados
'y les asigna los valores de los campos: existenc, fechainv, horainve, false de la tabla: sinven
'* Elimina de la Tabla: sinven los registros seleccinados para actualizar
'* Inserta Movimientos de Articulos en la Tabla: smoval
'-------------------------------------------------------------------
Dim SQL As String, ADonde As String
Dim RS As ADODB.Recordset
Dim DevStock As String
Dim CanStock2 As Currency, Diferencia As Currency
Dim Movimientos As Currency
Dim vTipoMov As CTiposMov
'Dim CodTipoMov As String * 3
Dim NumMovim As Long, numlinea As Long
Dim LetraSerie As String * 1
Dim CadValues As String
Dim bol As Boolean
    
    On Error Resume Next
    
    'Obtener Registros de la Tabla:sinven de los que se va a actualizar el Stock
    SQL = "SELECT sinven.* "
    
    'DAVID ENERO 2008
    'SQL = SQL & " FROM sinven "
    SQL = SQL & " FROM sinven  INNER JOIN sartic ON sinven.codartic=sartic.codartic"
    
    SQL = SQL & " WHERE " & cadFormula
    

    bol = True
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        bol = False
        ActualizarInventario = False
        MsgBox "No existen Registros en la Tabla: sinven para Actualizar Inventario.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    
    'Obtener el contador para los movimientos del Almacen que se esta inventariando
    'A cada registro de la tabla sinven se le asignará un numero de linea.
    '----------------------------------------------------------------------------
    Set vTipoMov = New CTiposMov
'    CodTipoMov = "REG"
    If vTipoMov.Leer("DFI") Then  'Se han cargado correctamente los valores de la clase
        'Obtener el contador para el codigo de Movimiento
        LetraSerie = vTipoMov.LetraSerie
        NumMovim = vTipoMov.ConseguirContador("DFI")
        numlinea = 1
        bol = True
    Else
        bol = False
    End If
    
    If Not bol Then
        Set vTipoMov = Nothing
        Exit Function
    End If
    
   
    On Error GoTo EActualizarInven:
    'Aqui empieza la transaccion
    Conn.BeginTrans
    
    While Not RS.EOF And bol 'Para cada registro de la tabla sinven
    
        'Introducir Movimiento de Entrada/Salida si hay diferencia entre el
        'Stock del Sistema y el Stock Real Inventariado.
        '------------------------------------------------------------------
        ADonde = "Introduciendo Movimiento de Entrada/Salida. Tabla: smoval."
        
        
        
                DevStock = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", RS!codartic, "T", , "codalmac", RS!codalmac, "N")
                If DevStock <> "" Then
        
                    CanStock2 = CCur(DevStock)
                    If AntiguaFormaInventariar Then
                        Movimientos = 0
                    Else
                        Movimientos = DBLet(RS!movpost, "N")
                    End If
                    Diferencia = RS!existenc - (CanStock2 - Movimientos)
                    If Diferencia <> 0 Then 'Insertar Movimiento de Entrada/Salida en Almacen
                        CadValues = DBSet(RS!codartic, "T") & ", " & RS!codalmac & ", '" & Format(RS!fechainv, "yyyy-mm-dd") & "', '"
                        CadValues = CadValues & Format(RS!horainve, "yyyy-mm-dd hh:mm:ss") & "', "
                        bol = InsertarMovimArticulos2(CadValues, RS!codartic, Diferencia, LetraSerie, NumMovim, numlinea)
                        numlinea = numlinea + 1
                    Else
                        bol = True
                    End If
                Else
                    bol = False
                End If
        

        
        
        'Actualizamos la Tabla: salmac
        '           salmac.canstock := existencia Real
        '           salmac.statusin := false (desbloqueamos los articulos )
        '---------------------------------------
        If bol Then
            ADonde = "Actualizando Stock de Articulos en Almacen. Tabla: salmac."
            Movimientos = CanStock2 + Diferencia
            SQL = "UPDATE salmac SET canstock=" & DBSet(Movimientos, "N") & ", statusin=0"
            SQL = SQL & " WHERE codartic=" & DBSet(RS!codartic, "T") & " AND codalmac=" & RS!codalmac
            Conn.Execute SQL
        End If

        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    If bol Then
'        'Pasamos la tabla de inventario real sinven al historico: shinve
'        'antes de eliminarla
'        ADonde = "Pasando registros de Inventario al Histórico: shinve."
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & "SELECT codartic,codalmac,fechainv,horainve,existenc "
'        SQL = SQL & " FROM sinven WHERE " & cadFormula
'        Conn.Execute SQL
    
        'Eliminamos los registros seleccionados de la Tabla: sinven
        '----------------------------------------------------------
        ADonde = "Eliminando registros de Inventario. Tabla: sinven."
       ' SQL = "DELETE FROM sinven "
  
        'DAVID ENERO 2008
        SQL = "DELETE sinven.* FROM sinven  INNER JOIN sartic ON"
        SQL = SQL & " sinven.codartic=sartic.codartic WHERE " & cadFormula
        Conn.Execute SQL
        
        
        'Incrementamos el contador para el Tipo De Movimiento
        '-----------------------------------------------------
        ADonde = "Actualizando el contador ."
        'bol = vTipoMov.IncrementarContador(
        vTipoMov.IncrementarContador ("DFI")
    End If
    Set vTipoMov = Nothing
        
EActualizarInven:
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
          SQL = "Actualizar Inventario." & vbCrLf & "----------------------------" & vbCrLf
          SQL = SQL & ADonde
          MuestraError Err.Number, SQL, Err.Description
          Conn.RollbackTrans
          ActualizarInventario = False
          Set vTipoMov = Nothing
    Else
        ActualizarInventario = True
        Conn.CommitTrans
    End If
End Function


Private Function InsertarMovimArticulos2(CadValues As String, codartic As String, Cantidad As Currency, LetraSerie As String, NumMovim As Long, numlinea As Long) As Boolean
Dim vImporte As Single, vPrecioVenta As String
Dim tipoMov As Byte
Dim SQL As String
On Error Resume Next
         
        'Obtener el precio de venta del articulo
         vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", codartic, "T")
        If vPrecioVenta <> "" Then
            vImporte = Abs(Cantidad * CSng(vPrecioVenta))
        Else
            vImporte = 0
        End If
        
        'Tipo de Movimiento de Almacen
        If Cantidad > 0 Then 'Insertar Movimiento de Entrada en Almacen
            tipoMov = 1
        ElseIf Cantidad < 0 Then 'Insertar Movimiento de Salida en Almacen
            tipoMov = 0
        End If
                                                                        
        SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                                                                             '      David 09
        SQL = SQL & " VALUES (" & CadValues & tipoMov & ", '" & "DFI" & "', " & DBSet(Abs(Cantidad), "N") & ", " & DBSet(vImporte, "N") & ", " & Val(txtCodigo(21).Text) & ", '"
        SQL = SQL & LetraSerie & "', " & NumMovim & ", " & numlinea & ")"
        Conn.Execute SQL
        
        If Err.Number <> 0 Then
             'Hay error , almacenamos y salimos
            InsertarMovimArticulos2 = False
        Else
            InsertarMovimArticulos2 = True
        End If
    
End Function


Private Function ValidarCamposInventario() As Boolean
'Comprobar que los campos requeridos tienen valor antes de abrir el listado
Dim B As Boolean

        B = True
        '- campo almacen debe tener valor
        If Trim(txtCodigo(13).Text) = "" Then
             MsgBox "El campo Almacen debe tener valor.", vbInformation
             PonerFoco txtCodigo(13)
             B = False
        End If
    
        '- fecha de inventario debe tener valor
        If B Then
            If (OpcionListado = 12 Or OpcionListado = 15 Or OpcionListado = 19) And Trim(txtCodigo(20).Text) = "" Then
                 MsgBox "El campo Fecha debe tener valor.", vbInformation
                 PonerFoco txtCodigo(20)
                 B = False
            End If
        End If
        
        'informe 19: stocks a una fecha
        'la fecha tiene que ser < a fecha hoy
        If OpcionListado = 19 And txtCodigo(20).Text <> "" Then
        
            'Esto estaba DEScomentado.
            'Hay que volverlo a conectar
        
'            If Not CDate(txtCodigo(20).Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
'                MsgBox "La fecha stock tiene que ser anterior a la fecha de hoy.", vbInformation
'                PonerFoco txtCodigo(20)
'                B = False
'            End If
        End If
        
        ValidarCamposInventario = B
End Function



Private Function ListaArtActivos(cadWhere As String, FechaIn As String) As String
    Dim RS As ADODB.Recordset
Dim SQL As String
Dim Lista As String
'Devuelve una cadena con la concatenacion de todos los articulos que
'no debe seleccionar ya que si tienen movimientos con fecha posterior
'a FechaIn.
'ej:    "[""00000004"", ""00000033""]"

    Lista = "["
    
    SQL = "SELECT distinct smoval.codartic from smoval "
    If InStr(cadWhere, "sartic") > 0 Then SQL = SQL & " INNER JOIN sartic ON smoval.codartic=sartic.codartic "
    SQL = SQL & " WHERE " & Replace(cadWhere, "salmac", "smoval")
    If cadWhere <> "" Then SQL = SQL & " AND "
    SQL = SQL & " fechamov>='" & Format(FechaIn, FormatoFecha) & "' "
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'        lista = lista & """" & RS.Fields(0).Value & """"
        Lista = Lista & DBSet(RS.Fields(0).Value, "T")
        RS.MoveNext
        If Not RS.EOF Then Lista = Lista & ", "
    Wend
    Lista = Lista & "]"
    ListaArtActivos = Lista
    RS.Close
    Set RS = Nothing
End Function



Private Sub ActualizarImprimir()
Dim I As Long
Dim Desde As Long, Hasta As Long
Dim SQL As String

    Select Case OpcionListado
    Case 7  'TRASPASO ALMACEN
        If frmVisReport.EstaImpreso = True Then
        'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
            If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
            If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
            For I = Desde To Hasta
                SQL = "UPDATE scatra SET situacio=1" 'Impreso
                SQL = SQL & " WHERE codtrasp=" & I
                Conn.Execute SQL
            Next I
        End If
        
    Case 8  'MOVIMIENTO ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For I = Desde To Hasta
                SQL = "UPDATE scamov SET situacio=1"
                SQL = SQL & " WHERE codmovim=" & I
                Conn.Execute SQL
           Next I
        End If
    End Select
End Sub


Private Sub CargarComboTipoList()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'1-Equipos, 2-Pagos, 3-Importes Contrato

    Me.cboTipoList.Clear
    cboTipoList.AddItem "Equipos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 1

    cboTipoList.AddItem "Pagos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 2

    cboTipoList.AddItem "Importes Contrato"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 3

End Sub




Private Sub CargarComboSituacion()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Abierta, 1-En Reparacion, 2-Pendiente, 3-Cerrado

    Me.cboSituaAviso.Clear
    
    cboSituaAviso.AddItem "-- Todas --"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 0
    
    cboSituaAviso.AddItem "Abierta"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 1

    cboSituaAviso.AddItem "En reparación"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 2
    
    cboSituaAviso.AddItem "Pendiente"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 3
    
    cboSituaAviso.AddItem "Cerrado"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 4

End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    NumParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
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


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = NumParam

        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            cadParam = cadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            NumParam = NumParam + 1
        Case "Marca"
            cadParam = cadParam & campo & "{sartic.codmarca}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
            End If
            NumParam = NumParam + 1
        Case "Proveedor"
            cadParam = cadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {sprove.nomprove}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {sprove.nomprove}" & "|"
            End If
            NumParam = NumParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            cadParam = cadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                cadParam = cadParam & nomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
            End If
            NumParam = NumParam + 1
    End Select

'Case "Familia"
'            cadParam = cadParam & "pGroup1=" & "{sartic.codfamia}" & "|"
'            cadParam = cadParam & "pGroup1Name= ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
'            numParam = numParam + 1
'            Select Case ListView2.ListItems(2).Text
'                Case "Marca"
'                    cadParam = cadParam & "pGroup2=" & "{sartic.codmarca}" & "|"
'                    cadParam = cadParam & "pGroup2Name= ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
'                    numParam = numParam + 1
'                    If ListView2.ListItems(3).Text = "Proveedor" Then
'                        Opcion = 1
'                    Else
'                        Opcion = 2
'                    End If
'                Case "Proveedor"
'                Case "Tipo Articulo"
'            End Select
End Function



Private Sub AbrirFrmActividades(Optional Indice As Integer)
    Set frmMtoActiv = New frmFacActividades
    frmMtoActiv.DatosADevolverBusqueda = "0|1|"
    frmMtoActiv.DeConsulta = True
    frmMtoActiv.Show vbModal
    Set frmMtoActiv = Nothing
End Sub



Private Sub AbrirFrmMarcas()
    Set frmMtoMarcas = New frmAlmMarcas
    frmMtoMarcas.DatosADevolverBusqueda = "0|1"
    frmMtoMarcas.DeConsulta = True
    frmMtoMarcas.Show vbModal
    Set frmMtoMarcas = Nothing
End Sub


Private Sub AbrirFrmAlmPropios()
    Set frmMtoAlPropios = New frmAlmAlPropios
    frmMtoAlPropios.DatosADevolverBusqueda = "0|1"
    frmMtoAlPropios.DeConsulta = True
    frmMtoAlPropios.Show vbModal
    Set frmMtoAlPropios = Nothing
End Sub


Private Sub AbrirFrmZonas()
    Set frmMtoZonas = New frmFacZonas
    frmMtoZonas.DatosADevolverBusqueda = "0|1"
    frmMtoZonas.DeConsulta = True
    frmMtoZonas.Show vbModal
    Set frmMtoZonas = Nothing
End Sub


Private Sub AbrirFrmRutas()
    Set frmMtoRutas = New frmFacRutas
    frmMtoRutas.DatosADevolverBusqueda = "0|1"
    frmMtoRutas.DeConsulta = True
    frmMtoRutas.Show vbModal
    Set frmMtoRutas = Nothing
End Sub


Private Sub AbrirFrmTarifas()
'tarifas venta
    Set frmMtoTarifas = New frmFacTarifas
    frmMtoTarifas.DatosADevolverBusqueda = "0|1"
    frmMtoTarifas.Show vbModal
    Set frmMtoTarifas = Nothing
End Sub


Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmMtoTArticulo = New frmAlmTipoArticulo
    frmMtoTArticulo.DatosADevolverBusqueda = "0|1"
    frmMtoTArticulo.DeConsulta = True
    frmMtoTArticulo.Show vbModal
    Set frmMtoTArticulo = Nothing
End Sub

Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmFacClientes
    frmMtoClientes.DatosADevolverBusqueda = "0|1"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
End Sub


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim RS As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set RS = New ADODB.Recordset
        RS.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RS.EOF Then
            FechaIni = DBLet(RS!FechaIni, "F")
            '## LAURA 19/06/2008
'            FechaFin = DBLet(RS!FechaFin, "F") + 365
'            FechaFin = DateAdd("d", 365, DBLet(RS!FechaFin, "F"))
            FechaFin = DateAdd("yyyy", 1, DBLet(RS!FechaFin, "F"))
            '##
            
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        RS.Close
        Set RS = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Function ContabilizarFacturas(cadTabla As String, cadWhere As String) As Boolean
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste2 As Byte


        '0.- Si devuelve la funcion el 0 habra CC sin confgurar en trabaja
        '1.- Todos los CC son el mismo
        '2.- Mas de un CC. Hay que agrupar

    ContabilizarFacturas = False

    If cadTabla = "scafac" Then
        SQL = "VENCON" 'contabilizar facturas de venta
    ElseIf cadTabla = "scafpc" Then
        SQL = "COMCON" 'contabilizar facturas de compra
    End If

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(31).Text = "" Then
        txtCodigo(31).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(32).Text = "" Then
        txtCodigo(32).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     
     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(32) Then Exit Function
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If cadTabla = "scafac" Then
        If Me.txtCodigo(31).Text = "" Then
            MsgBox "Fecha inicio incorrecta", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    'comprobar si existen en Ariges facturas anteriores al periodo solicitado
    'sin contabilizar.
    If Me.txtCodigo(31).Text <> "" Then 'anteriores a fechadesde
        SQL = "SELECT COUNT(*) FROM " & cadTabla
        If cadTabla = "scafac" Then
            SQL = SQL & " WHERE fecfactu <"
        ElseIf cadTabla = "scafpc" Then
            SQL = SQL & " WHERE fecrecep <"
        End If
        SQL = SQL & DBSet(txtCodigo(31), "F") & " AND intconta=0 "
        
        
        'Si contabiliza tickets agrupados
        If OptProve.Tag = "" Then
            If vParamAplic.ContabilizarTicketAgrupados Then SQL = SQL & " AND codtipom <>'FTI' "
        Else
            SQL = SQL & " AND scafac.codtipom  = 'FTG' "
        End If
        
        '## LAURA 20/06/2008
        If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
            SQL = SQL & " AND scafac.codtipom = " & DBSet(Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3), "T")
        End If
        
        
        If RegistrosAListar(SQL) > 0 Then
            If MsgBox("Hay Facturas anteriores sin contabilizar. " & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                Exit Function
            End If
        End If
    End If
    
    
'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100
        
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    B = CrearTMPFacturas(cadTabla, cadWhere)
    If Not B Then Exit Function
            
            
    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    If cadTabla = "scafac" Then
        SQL = SQL & ".codtipom=tmpFactu.codtipom AND "
    Else
        SQL = SQL & ".codprove=tmpFactu.codprove AND "
    End If
    SQL = SQL & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(SQL, cadWhere) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
            
    '---- Preparamos la pantalla de Contabilizar
    'Visualizar la barra de Progreso
    
    
    
    If Me.FrameTipMov.visible Then
        Me.FrameRepxDia.Height = 6100
        Me.FrameProgress.Top = 4400
    Else
        Me.FrameRepxDia.Height = 5100
        Me.FrameProgress.Top = 3350
    End If
    Me.Height = Me.FrameRepxDia.Height
    Me.FrameProgress.visible = True
    Me.Refresh
            
    Me.lblProgess(0).Caption = "Comprobaciones: "
    CargarProgres Me.ProgressBar1, 100
        
        
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariges
    '--------------------------------------------------------------------------
    IncrementarProgres Me.ProgressBar1, 10
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando letras de serie ..."
        B = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        SQL = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
        B = ComprobarNumFacturas_new(cadTabla, SQL)
    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    B = ComprobarCtaContable_new(cadTabla, 1)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    Else
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Compras en contabilidad ..."
    End If
    
    
    
    If vParamAplic.ContabilizacionMoixent Then
        B = ComprobarCtabasesMoixent(cadTabla = "scafac", 2)
    Else
        'Ventas moixent
        B = ComprobarCtaContable_new(cadTabla, 2)
    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    
    If Me.OptProve.Tag <> "" Then
        'TIKETS. Voy a comprobar las cuentas de la familia
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles tickets ..."
        Me.lblProgess(1).Refresh
        
        
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not B Then Exit Function
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    If vEmpresa.TieneAnalitica Then
       Me.lblProgess(1).Caption = "Comprobando Contabilidad Analítica ..."
       B = ComprobarCtaContable_new(cadTabla, 3)
       If Not B Then Exit Function
       
       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
       B = cadTabla = "scafac"
       CCoste2 = ComprobarCCoste2(cadWhere, B)
       If CCoste2 = 0 Then Exit Function 'Error comprobando CCs
       
    Else
        'No tiene analitica, NO agrupamos por codtraba
        CCoste2 = 0
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    
    If Me.OptProve.Tag <> "" Then
        Me.lblProgess(1).Caption = "Comprobando Ctas facmilias TICKETS ..."   'FTG
        B = ComprobarCtaContable_new(cadTabla, 4)
        If Not B Then Exit Function
    End If
    
    
    'Comprobamos, si es factura proveedore, que si el tipoprove = 3 (REA)
    'entonces tiene que existir el paremetro aplicacion codret
    If cadTabla = "scafpc" Then
        If vParamAplic.CtaReten = "" Then
            SQL = "SELECT COUNT(*) FROM scafpc,sprove WHERE scafpc.codprove = sprove.codprove and tipprove=3"
            If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
            If RegistrosAListar(SQL) > 0 Then
                MsgBox "Existen facturas SOCIOS proveedor con cta. retencion y no esta configurada", vbExclamation
                Exit Function
            End If
        
        
            'Neuvo 29Mayo 2008
            ' Cualquier factura puede llevar retencion. Necesito que la cuenta de retencion este configurada
            SQL = "SELECT COUNT(*) FROM scafpc  WHERE  tiporet=0 and impret<>0"
            If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
            If RegistrosAListar(SQL) > 0 Then
                MsgBox "Existen facturas proveedor con retencion y no esta configurada", vbExclamation
                Exit Function
            End If
         End If
        
    End If
    
    Me.lblProgess(1).Caption = "Fechas contabilizacion"
    Me.lblProgess(1).Refresh
    B = NuevasComprobacionesContabilizacion(cadTabla = "scafpc", cadWhere)
    If Not B Then Exit Function
    
    
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgess(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTabla & vbCrLf & cadWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)
    
    '---- Pasar las Facturas a la Contabilidad
    B = PasarFacturasAContab(cadTabla, CCoste2)
    
    
    
    '---- Mostrar ListView de posibles errores (si hay)
    If Not B Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        'Para la facturacion de TICKTS agrupada NO mostramos el mensaje de OK
        If Me.OptProve.Tag = "" Then
            If cadTabla = "scafac" Then MsgBox "El proceso ha finalizado correctamente.", vbInformation
        End If
    End If
    
    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If cadTabla <> "scafac" Then
        If NumRegistros("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
            InicializarVbles
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            NumParam = NumParam + 1
            
            cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
            NumParam = NumParam + 1
            cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
            cadNomRPT = "rContabPRO.rpt"
            conSubRPT = False
            cadTitulo = "Listado contabilizacion FRAPRO"
            
            LlamarImprimir
        End If
    End If
    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    ContabilizarFacturas = True
End Function

'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
Private Function PasarFacturasAContab(cadTabla As String, miCC As Byte) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim I As Integer
Dim NumFactu As Integer
Dim Codigo1 As String
Dim ContabilizacionAgrupadaTickets As Boolean
'ENERO 2009
Dim cContaFra As cContabilizarFacturas


    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    
    'Si escontailizacion de facturas de tickets agrupados
    ContabilizacionAgrupadaTickets = False
    If Me.OptProve.Tag <> "" Then ContabilizacionAgrupadaTickets = True
    
    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    If cadTabla = "scafac" Then
        Codigo1 = "codtipom"
    Else
        Codigo1 = "codprove"
    End If
    SQL = SQL & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    SQL = SQL & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        NumFactu = RS.Fields(0)
    Else
        NumFactu = 0
    End If
    RS.Close
    Set RS = Nothing


    'Enero 2009
    '------------------------------------------------------------
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        SQL = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        SQL = SQL & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        SQL = SQL & Space(50) & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    

    'Modificacion 20 Abril 2008
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If NumFactu > 0 Then
        CargarProgres Me.ProgressBar1, NumFactu
        
        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "
            
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        B = True
   
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not RS.EOF
        
            'Segun sea cli o pro
            If cadTabla = "scafac" Then
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "T") & " AND scafac.numfactu=" & RS!NumFactu
                SQL = SQL & " and scafac.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFactura(SQL, miCC, ContabilizacionAgrupadaTickets, cContaFra) = False And B Then B = False
            Else
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "N") & " and scafpc.numfactu=" & DBSet(RS!NumFactu, "T")
                SQL = SQL & " and scafpc.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFacturaProv(SQL, miCC, Orden2, cContaFra) = False And B Then B = False
            End If
            
            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(SQL, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----
            
            IncrementarProgres Me.ProgressBar1, 1
            Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & NumFactu & ")"
            Me.Refresh
            I = I + 1
            RS.MoveNext   'Siguiente factura
        Wend
        
        'Veremos si ha dado error la contabilizacion de factiras
        If cContaFra.TieneErrores Then cContaFra.MuestraErroresContabilizacion
        
        
        RS.Close
        Set RS = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then B = False
    Set cContaFra = Nothing
    If B Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function



Private Sub ListadosAlmacen(H As Integer, W As Integer)
    'LISTADOS DE ALMACENES
    '---------------------
    Label4(91).Caption = "" 'La del inventario
    chkStockFechaAceite.visible = False
    Select Case OpcionListado
        Case 1   'Listados de Marcas
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Marcas"
            indFrame = 1
            Codigo = "{smarca.codmarca}"
            Orden1 = "{smarca.codmarca}"
            Orden2 = "{smarca.nommarca}"
            cadTitulo = "Listado Marcas"
            cadNomRPT = "rAlmMarcas.rpt"
            conSubRPT = False
            
        Case 2   'Listado de Almacenes Propios
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Almacenes"
            indFrame = 1
            Codigo = "{salmpr.codalmac}"
            Orden1 = "{salmpr.codalmac}"
            Orden2 = "{salmpr.nomalmac}"
            cadTitulo = "Listado Almacenes Propios"
            cadNomRPT = "rAlmAPropios.rpt"
            conSubRPT = False
            
        Case 3   'Listado de Tipos de Unidad
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Unidad"
            indFrame = 1
            Codigo = "{sunida.codunida}"
            Orden1 = "{sunida.codunida}"
            Orden2 = "{sunida.nomunida}"
            cadTitulo = "Listado Tipos de Unidad"
            cadNomRPT = "rAlmTUnidad.rpt"
            conSubRPT = False
            
        Case 4   'Listado de Tipos de Artículos
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Artículos"
            indFrame = 1
            Codigo = "{stipar.codtipar}"
            Orden1 = "{stipar.codtipar}"
            Orden2 = "{stipar.nomtipar}"
            txtCodigo(1).Tag = CadTag
            txtCodigo(2).Tag = CadTag
            cadTitulo = "Listado Tipos de Artículos"
            cadNomRPT = "rAlmTArticulo.rpt"
            conSubRPT = False
            
        Case 6    'Listado de Artículo
            ponerFrameArticulosVisible True, H, W
            CargarListViewOrden
            Codigo = "{sartic"
            indFrame = 11
            cadTitulo = "Listado de Artículos"
            
            
        Case 110   'Listados Ubicaciones Almacen
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Ubicaciones Almacen"
            indFrame = 1
            Codigo = "{subica.codubica}"
            Orden1 = "{subica.codubica}"
            Orden2 = "{subica.nomubica}"
            cadTitulo = "Listado Ubicaciones Almacen"
            cadNomRPT = "rAlmUbica.rpt"
            conSubRPT = False
            
        Case 18, 247, 510 'Informe Stocks Maximos y Minimos   'OPCION: 247 es este tb
            ponerFrameArticulosVisible True, H, W
            Codigo = "{salmac"
            indFrame = 11
            'cmbProduccion.ListIndex = 0
            'cmbProduccion.visible = vParamAplic.Produccion
            'Pongo visible false la tarifa
            
            Label4(90).visible = False 'vParamAplic.Produccion
        Case 7, 8 '7: Informe de Traspasos de Almacen
                  '8: Informe de Movimientos de Almacen
            If OpcionListado = 7 Then
                Me.lblTitulo(2).Caption = "Informe Traspaso de Almacen"
                Me.Label2(1).Caption = "Nº Traspaso"
                Codigo = "{scatra.codtrasp}"
            Else
                Me.lblTitulo(2).Caption = "Informe Movimientos de Almacen"
                Me.Label2(1).Caption = "Nº Movimiento"
                Codigo = "{scamov.codmovim}"
            End If
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 9 'Informe Movimiento Artículos
            W = 10700
            H = 5775
            PonerFrameVisible Me.FrameMovArtic, True, H, W
            indFrame = 3
            Codigo = "{smoval.codartic}"
            cadTitulo = "Informe Movimientos Articulos"
            conSubRPT = True
            CargarListView
            
        Case 12 '12: Listado Toma de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.chkImprimeStock.visible = True
            Me.lbltituloInven.Caption = "Listado Toma de Inventario Articulos"
            cadTitulo = "Toma Inventario Articulos"
            'codigo = "{salmac.codalmac}"
            
        Case 13 '13: Listado Diferencias de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Diferencias de Inventario Articulos"
            'codigo = "{sinven.codalmac}"
            cadTitulo = "Diferencias Inventario Articulos"
            
        Case 14 '14: Actualizar Direfencias Inventario (NO IMPRIME INFORME)
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Actualizar Diferencias de Inventario de Articulos"
            Me.Caption = "Inventario de Articulos"
            
        Case 15 '15: Listado de Articulos Inactivos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Articulos Inactivos"
            cadTitulo = "Listado Articulos Inactivos"
    
        Case 16 '16 .- Listado Valoracion de Stocks Inventariados
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks Inventariados"
            cadTitulo = "Listado Valoración Stocks Inventariados"
            
        Case 17 '17 .- Listado Valoración Stocks
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks"
            cadTitulo = "Listado Valoración Stocks"
            
        Case 19 '19 .- Inf. Stocks a una Fecha
            chkStockFechaAceite.visible = True
            chkStockFechaAceite.Value = 1
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Informe Stocks a una Fecha"
            cadTitulo = "Stocks a una Fecha"
    End Select
End Sub



Private Sub ListadosFacturacion(H As Integer, W As Integer)
    Select Case OpcionListado
        Case 20    'Listado de Actividades de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Actividades de Clientes"
            indFrame = 1
            Codigo = "{sactiv.codactiv}"
            Orden1 = "{sactiv.codactiv}"
            Orden2 = "{sactiv.nomactiv}"
            cadTitulo = "Listado Actividades de Clientes"
            cadNomRPT = "rFacActividades.rpt"
            
        Case 21    'Listado de Zonas de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Zonas de Clientes"
            indFrame = 1
            Codigo = "{szonas.codzonas}"
            Orden1 = "{szonas.codzonas}"
            Orden2 = "{szonas.nomzonas}"
            cadTitulo = "Listado Zonas de Clientes"
            cadNomRPT = "rFacZonas.rpt"
        
        Case 22    'Listado de Rutas de Asistencia
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Rutas de Asistencia"
            indFrame = 1
            Codigo = "{srutas.codrutas}"
            Orden1 = "{srutas.codrutas}"
            Orden2 = "{srutas.nomrutas}"
            cadTitulo = "Listado Rutas de Asistencia"
            cadNomRPT = "rFacRutas.rpt"
            
        Case 23     'Listado de Tipos de Formas de Envío
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Formas de Envío"
            indFrame = 1
            Codigo = "{senvio.codenvio}"
            Orden1 = "{senvio.codenvio}"
            Orden2 = "{senvio.nomenvio}"
            cadTitulo = "Listado Formas de Envio"
            cadNomRPT = "rFacEnvio.rpt"
            
        Case 24    'Tarifas Venta
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tarifas Venta"
            indFrame = 1
            Codigo = "{starif.codlista}"
            Orden1 = "{starif.codlista}"
            Orden2 = "{starif.nomlista}"
            cadTitulo = "Listado Tarifas Venta"
            cadNomRPT = "rFacTarifasVen.rpt"
            
        Case 27     'Situaciones Especiales
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Situaciones Especiales"
            indFrame = 1
            Codigo = "{ssitua.codsitua}"
            Orden1 = "{ssitua.codsitua}"
            Orden2 = "{ssitua.nomsitua}"
            cadTitulo = "Listado Situaciones Especiales"
            cadNomRPT = "rFacSituaciones.rpt"
            
        Case 28    '28: Informe de Tarifas de Precios
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Tarifas de Artículos"
            Codigo = "{slista"
            indFrame = 5
            cadTitulo = "Listado Tarifas Articulos"
            
        Case 29  '29: Informe Promociones
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Promociones Tarifas"
            Codigo = "{spromo"
            indFrame = 5
            cadTitulo = "Listado Promociones de Tarifas"
            
        Case 30 '30: Informe Precios Especiales
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Precios Especiales Artículos"
            Codigo = "{sprees"
            indFrame = 5
            cadTitulo = "Listado Precios Especiales"
            
        Case 245, 247 '245: Informe control margenes tarifas
            indFrame = 5
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Control Margenes de Tarifas"
            Codigo = "{slista"
            cadTitulo = "Listado Control Margenes Tarifas"
            cboDecimales.ListIndex = 4
        Case 246 '246: Informe margen ventas x articulo
            indFrame = 15
            H = 5300
            W = 7820
            PonerFrameVisible Me.FrameEstMargenes, True, H, W
            cadTitulo = "Listado Margen ventas por artículo"
    End Select
End Sub


Private Sub ListadosMantenimiento(H As Integer, W As Integer)
'=============================================
'==== Listados de MANTENIMIENTOS

    Select Case OpcionListado
        Case 70, 71, 76, 78, 79 'Listado Mantenimientos
            FrameManteAnu.visible = False
            PonerFrameManteVisible True, H, W
            Select Case OpcionListado
            Case 70, 76
                CargarComboTipoList
                If OpcionListado = 70 Then
                    Me.cboTipoList.ListIndex = 0
                Else
                    FrameManteAnu.visible = True
                    Me.cboTipoList.ListIndex = 2
                End If
                cadTitulo = "Informe de Mantenimientos"
                conSubRPT = False
            Case 71
                    txtCodigo(53).Text = Format(Now, "dd/mm/yyyy")
                    txtCodigo(54).Text = Format(Now, "dd/mm/yyyy")
                    cadTitulo = "Informe Revisiones Mantenimientos"
                    conSubRPT = True
            Case 78
            
            Case 79
            
            End Select
            indFrame = 9
            
        Case 72 'Informe Fichas de Mantenimiento
            H = 5295
            W = 7395
            PonerFrameVisible Me.FrameFichasMan, True, H, W
            txtCodigo(61).Text = Year(Now) 'Ejercicio
            indFrame = 10
            cadTitulo = "Informe Fichas Mantenimientos"
            conSubRPT = True
        Case 77
            'Informe teorico
            H = FrameListMant2.Height
            W = FrameListMant2.Width
            PonerFrameVisible FrameListMant2, True, H, W
            indFrame = 77
    End Select
End Sub



Private Sub ListadosCompras(H As Integer, W As Integer)
'=============================================
'==== Listados de COMPRAS

    Select Case OpcionListado
        Case 309 '309: Listado precios de compra
            H = 4450
            W = 6920
            PonerFrameVisible Me.FrameDtosFM, True, H, W
            Me.Frame4.visible = True
            Me.Frame4.Top = 840
            Me.Frame5.visible = False
            Me.Frame6.visible = False
            Me.cmdAceptarDtosFM.Top = 3500
            Me.cmdCancel(12).Top = Me.cmdAceptarDtosFM.Top
            indFrame = 6
    End Select
End Sub



Private Sub ListadosReparaciones(H As Integer, W As Integer)
'=============================================
'==== Listados de REPARACIONES

    Select Case OpcionListado
        Case 407 'Sustitución Num. serie
            H = 3700
            W = 5720
            PonerFrameVisible Me.FrameRepSustNSerie, True, H, W
            Me.lblNumSerie(0).Caption = "Nº Serie:   " & NumCod
            Me.lblNumSerie(1).Caption = "Artículo:   " & Me.CadTag
            Me.Caption = "Numeros de Serie"
            indFrame = 13
            
        Case 409 '409: Listado de avisos de averia pendientes
            H = FrameListAvisosPtes.Height + 120
            W = FrameListAvisosPtes.Width + 120
            PonerFrameVisible Me.FrameListAvisosPtes, True, H, W
            CargarComboSituacion
            indFrame = 14
            Me.cboSituaAviso.ListIndex = 0
    End Select
End Sub




'---------------------------------------------------
'Para los bultos
Private Sub LimpiarTextosBultos()
Dim I As Integer
    For I = 2 To 6
        Me.txtBultos(I).Text = ""
        Me.txtBultos(I).Tag = ""
    Next I
End Sub



Private Sub PonerCamposDireccionBultos(Indice As Integer)
Dim I As Integer

    'El indice mara el listindex del combo, por lo tanto sera indice + 1
    For I = 2 To 6
        Me.txtBultos(I).Text = RecuperaValor(Me.txtBultos(I).Tag, Indice + 1)
    Next I
End Sub

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'   Borre de facturas
'
'
'   Borraremos las tablas de facturas , albaranes, hcos....
'
Private Sub CargaFechasPosibleEliminacion()
Dim F As Date
Dim F2 As Date
    Set miRsAux = New ADODB.Recordset
    cmbEliFac.Clear
    Codigo = "select min(fecfactu) from scafac"
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F2 = DateAdd("yyyy", -5, CDate("01/01/" & Year(Now)))

    Codigo = F2
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Codigo = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Codigo = "31/12/" & Year(CDate(Codigo))
    
    While CDate(Codigo) < F2
        
        cmbEliFac.AddItem "     " & Format(CDate(Codigo), "dd/mm/yyyy")
        Codigo = CStr(DateAdd("yyyy", 1, CDate(Codigo)))
    
    Wend
    If cmbEliFac.ListCount > 0 Then cmbEliFac.ListIndex = 0
End Sub

Private Function BorrarFacturas() As Boolean
Dim FechaBorre As Date



    On Error GoTo EBorraFac
    BorrarFacturas = False
    
    FechaBorre = CDate(Trim(Me.cmbEliFac.List(cmbEliFac.ListIndex)))
    
    'Compruebo si estaban todas las facturas contabilizadas
    '------------------------------------------------------
    Codigo = "Select count(*) from scafac where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
    
        
    'lo mismo para proeedores
    Codigo = "Select count(*) from scafpc where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas de proveedores sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
        
        
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 1, vUsu, "Borre facturas: " & Format(FechaBorre, "dd/mm/yyyy")
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '   Lo dicho. LAS TABLAS son las indcadas above (jeje arriba)
    '   La fecha la manda fecfactu
    Codigo = "slifac|scafac1|svenci|srecom|scafac|"
    For NumRegElim = 1 To 5
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla CLI: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        Me.Refresh
        DoEvents
        Conn.Execute Orden1
    Next NumRegElim
    
    '---------------------------------------------------------------------------------
    'Albarananes CLIENTES.
    '--
    Codigo = "scaalb|schalb|slialb|slhalb|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE codtipom = '"
        While Not miRsAux.EOF
            Conn.Execute Orden1 & miRsAux!codTipoM & "'  AND numalbar = " & miRsAux!NumAlbar
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las cabceeras
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'Pedidos CLIENTES.
    '--
    Codigo = "scaped|schped|sliped|slhped|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedcl = "
        While Not miRsAux.EOF
            Conn.Execute Orden1 & miRsAux!numpedcl
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Cabce
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'ofertas CLIENTES.
    '--
    Codigo = "scapre|schpre|slipre|slhpre|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numofert = "
        While Not miRsAux.EOF
            Conn.Execute Orden1 & miRsAux!NumOfert
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert <='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
        
    Next
    DoEvents
    
    
    Codigo = "scarep|schrep|slirep|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Reparaciones: " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar<='" & Format(FechaBorre, FormatoFecha) & "'"
        If NumRegElim = 1 Then
            'Lineas de reparacion solo hay en scarep
            'En shrep no hay lineas
            miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
            Orden1 = "DELETE FROM " & Orden1 & " WHERE numrepar = "
            While Not miRsAux.EOF
                Conn.Execute Orden1 & miRsAux!numrepar
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar <='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
        
    Next
    DoEvents
    
    
    'TPV
    Label3(83).Caption = "TPV"
    Label3(83).Refresh
    Orden1 = " WHERE  fecventa <='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute "DELETE FROM sliven " & Orden1
    Conn.Execute "DELETE FROM scaven " & Orden1

    
    'PRODUCCION
    Label3(83).Caption = "Produccion"
    Label3(83).Refresh
    Orden1 = "Select * from sordprod WHERE  feccreacion<='" & Format(FechaBorre, FormatoFecha) & "'"
    miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Orden1 = "DELETE FROM sliordpr WHERE codigo = "
    While Not miRsAux.EOF
        Conn.Execute Orden1 & miRsAux!Codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    Orden1 = "DELETE from sordprod WHERE  feccreacion <='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute Orden1

    Me.Refresh
    DoEvents
    
    '---------------------------------------------------------------------------------
    'Facturas proveedor
    '--
    Codigo = "slifpc|scafpa|scafpc|"
    For NumRegElim = 1 To 3
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla PRO: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
    Next NumRegElim
    
    
    
    
    Codigo = "slhalp|slialp|scaalp|schalp|"
    For NumRegElim = 1 To 4
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes prov: " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
        
    Next
    DoEvents
    
    
    
    
    '-----------------------------------------------
    'Pedidos proveedor
    '--
    Codigo = "scappr|schppr|slippr|slhppr|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedpr = "
        While Not miRsAux.EOF
            Conn.Execute Orden1 & miRsAux!numpedpr
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr <='" & Format(FechaBorre, FormatoFecha) & "'"
        Conn.Execute Orden1
        
    Next
    Me.Refresh
    DoEvents
    
    'slhmov slhtra
    Label3(83).Caption = "Hco movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhmov WHERE  fecmovim<='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute Orden1
    
    Label3(83).Caption = "Hco traspasos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhtra WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute Orden1
    
    
    'Ahora me cargo los movimientos en la smoval
    Label3(83).Caption = "Movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM smoval WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute Orden1
    
    'Inventario
    Label3(83).Caption = "Hco inventario"
    Label3(83).Refresh
    Orden1 = "DELETE FROM shinve WHERE  fechainv<='" & Format(FechaBorre, FormatoFecha) & "'"
    Conn.Execute Orden1
    
    
    BorrarFacturas = True
    Exit Function
EBorraFac:
    MuestraError Err.Number
End Function


'Envio -EMAIL

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameMantenimientos.Height
        Me.Width = Me.FrameMantenimientos.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    DoEvents
    Me.Refresh
End Sub





Private Function GeneracionEnvioMail() As Boolean
Dim m As CParamRpt

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    Set m = New CParamRpt
    If m.Leer(21) = 1 Then
        Set m = Nothing
        Exit Function
    End If
    
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    miRsAux.Open cadSelect, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not miRsAux.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Mantenimiento: " & miRsAux!codartic & " Cliente: " & miRsAux!CodProve
        Label14(22).Refresh
        
'
        cadFormula = "({scaman.nummante}='" & miRsAux!codartic & "') "
        cadFormula = cadFormula & " AND ({scaman.codclien}=" & miRsAux!CodProve & ") "


        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = NumParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = m.Documento
            .opcion = 78  'Carta renovacion manteniientos
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.PBMail.Value = Me.PBMail.Value + 1
        If (Me.PBMail.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            Espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Format(miRsAux!CodProve, "0000000") & ".pdf"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set m = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function



Private Function HacerSQLListado82_83() As Boolean
    
On Error GoTo EHacerSQLListado82_83
    
    
    HacerSQLListado82_83 = False
    InicializarVbles


    If OpcionListado = 82 Then
        'Hacer UPDATE de scaalb
        Codigo = "UPDATE scaalb set factursn = 1 "
        If NumCod <> "" Then cadSelect = " codtipom ='" & NumCod & "'"
        
        cadParam = "fechaalb"
        cadFormula = CadenaDesdeHastaBD(txtCodigo(117).Text, txtCodigo(118).Text, "codclien", "N")
        If cadFormula <> "" Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & cadFormula
        End If
        

    Else
        'Hacer borrar avisos
        Codigo = "DELETE FROM scaavi"
        cadSelect = " situacio = 3"
        cadParam = "fechaavi"
    End If
    
    cadFormula = CadenaDesdeHastaBD(txtCodigo(119).Text, txtCodigo(120).Text, cadParam, "F")
    If cadFormula <> "" Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadFormula
    End If
    
    If cadSelect <> "" Then cadSelect = " WHERE " & cadSelect
    Codigo = Codigo & cadSelect
    Conn.Execute Codigo
    
    If OpcionListado = 83 Then MsgBox "Proceso finalizado", vbExclamation
    
    HacerSQLListado82_83 = True
    Exit Function
EHacerSQLListado82_83:
    MuestraError Err.Number
End Function








Private Function InsertarDatosTemporalInventario() As Boolean
Dim C As String
Dim Aux As String


    On Error GoTo EInsertarDatosTemporalInventario
    Label4(91).Caption = "Datos 1"
    Label4(91).Refresh
    
    Conn.Execute "Delete from tmpTomInventario WHERE codusu = " & vUsu.Codigo
    
    Aux = "SELECT " & vUsu.Codigo & ",sartic.codfamia,salmac.codalmac, salmac.codartic ,sartic.codunida , salmac.canstock , 0.00"
    Aux = Aux & " FROM   (((salmac salmac INNER JOIN sartic sartic ON salmac.codartic=sartic.codartic)"
    Aux = Aux & " INNER JOIN sfamia sfamia ON sartic.codfamia=sfamia.codfamia) INNER JOIN sprove"
    Aux = Aux & " sprove ON sartic.codprove=sprove.codprove) INNER JOIN sunida sunida ON"
    Aux = Aux & " sartic.codunida=sunida.codunida"
    'WHERE
    Aux = Aux & " WHERE " & cadSelect
    'ORDER
    Aux = Aux & " ORDER BY salmac.codalmac, sartic.codfamia, sartic.nomartic"
    
    
    
    Aux = "insert into tmpTomInventario(`codusu`,`codfamia`,`codalmac`,`codartic`,`codunida`,`canstock`,`movpost`) " & Aux
    Conn.Execute Aux
    
    
    'OK. AHora veremos para cada articulo los movimientos posteriores a la fecha de inventario
    Set miRsAux = New ADODB.Recordset

    'Para cada articulo con movimientos posteriores a la fecha
    'sumo esos movimientos
    Label4(91).Caption = "Movimientos"
    Label4(91).Refresh
    
    Aux = "select codartic,sum(if(tipomovi=0,-cantidad,cantidad)) from smoval where "
    Aux = Aux & "  codalmac = " & txtCodigo(13).Text
    Aux = Aux & " AND fechamov > " & DBSet(txtCodigo(20).Text, "F") & " and codartic in"
    Aux = Aux & " (select codartic from tmpTomInventario WHERE codusu = " & vUsu.Codigo & ") group by codartic"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Label4(91).Caption = miRsAux!codartic
        Label4(91).Refresh
        
        Aux = "UPDATE tmpTomInventario SET movpost = " & TransformaComasPuntos(CStr(DBLet(miRsAux.Fields(1), "N")))
        Aux = Aux & " WHERE codartic = '" & miRsAux!codartic & "'"
        Aux = Aux & " AND codalmac = " & txtCodigo(13).Text
        Aux = Aux & " AND codusu  = " & vUsu.Codigo
        Conn.Execute Aux

    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
EInsertarDatosTemporalInventario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        InsertarDatosTemporalInventario = True
    End If
    Set miRsAux = Nothing
    Label4(91).Caption = ""
    Label4(91).Refresh
    
End Function


'Private Function MovimientosPosteriores(ByRef C As String) As Currency
'
'    MovimientosPosteriores = 0
'    C = "Select codartic,sum(if(detamovi=0,-cantidad,cantidad)) from smoval where codartic = '" & C & "'"
'    C = C & " AND codalmac=" & Me.txtCodigo(1).Text
'    C = C & " AND fechamov > " & DBSet(Me.txtCodigo(3).Text, "F")
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    while not
'End Function



Private Sub AjustesStocksFechaAceite()
Dim C As String
Dim miSql As String
Dim Cantidad As Currency
Dim campo As String
Dim R As ADODB.Recordset

    
    
    C = "DELETE from tmpstockfec2 where codusu = " & vUsu.Codigo
    Conn.Execute C
    
    
    Set R = New ADODB.Recordset
    C = "select codfamia from sartic where conjunto=1 or factorconversion<>1 group by 1"
    R.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    While Not R.EOF
        C = C & ", " & R.Fields(0)
        R.MoveNext
    Wend
    R.Close
    
    If C = "" Then Exit Sub
    C = Mid(C, 2)
    
    'QUITO LOS QU NO SON ACEITE

    C = "DELETE tmpstockfec.* from tmpstockfec,sartic where  tmpstockfec.codartic=sartic.codartic AND codusu = " & vUsu.Codigo & " AND NOT codfamia in (" & C & ")"
    Conn.Execute C



    'Si ha puesto que no muestre lo de stock =0 los borro tmb
    '
    
    If Me.chkSinStock.Value = 0 Then
        
        C = "DELETE  from tmpstockfec where  codusu = " & vUsu.Codigo & " AND stock=0"
        Conn.Execute C
    End If
    'Para los que quedan....
    'Ire viendo cuanto hay de lo suyo y cuanto de materia prima
    Set miRsAux = New ADODB.Recordset
    C = "select tmpstockfec.*,conjunto,factorconversion from tmpstockfec,sartic where  tmpstockfec.codartic=sartic.codartic AND codusu = " & vUsu.Codigo
    R.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = "|"
    While Not R.EOF
        
        
            If R!Conjunto > 0 Then
                ' LLEVA conjuntos
                miSql = "select sarti1.codarti1,cantidad from sarti1,sartic where  sarti1.codarti1=sartic.codartic and factorconversion<>1"
                miSql = miSql & " AND sarti1.codartic =" & DBSet(R!codartic, "T")
                miRsAux.Open miSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                
                'If Mid(miRsAux!codArtic, 1, 9) = "002700090" Then Stop
                
                
                
                Cantidad = DBLet(R!stock, "N")
        
                
                'Para no tener que hacer un select para saber si ya ha sido insertado en tmpstock, utilizar
                'el string cadSelect para ir metiendo los ya insertados.
                
                While Not miRsAux.EOF
                    'El articulo en cuestion
                    miSql = "|" & miRsAux!codarti1 & "|"
                    Cantidad = Cantidad * miRsAux!Cantidad   'Esta es la cantidad nueva
                    campo = TransformaComasPuntos(CStr(Cantidad))
                    If InStr(1, C, miSql) > 0 Then
                        'Ya esta insertado. Es un UPDATE
                        miSql = "UPDATE tmpstockfec2 SET stock=stock + " & campo
                        miSql = miSql & " WHERE codusu = " & vUsu.Codigo & " and codartic = " & DBSet(miRsAux!codarti1, "T")
                        miSql = miSql & " AND codalmac= 1"
                    Else
                        C = C & miRsAux!codarti1 & "|"
                        miSql = "INSERT INTO tmpstockfec2(codusu,codartic,codalmac,stock)  VALUES (" & vUsu.Codigo & "," & DBSet(miRsAux!codarti1, "T")
                        miSql = miSql & ",1," & campo & ")"
                        
                    End If
                    Conn.Execute miSql
                    'No deberia haber mas (seria un coupage)
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
           Else
              'NO ES CONJUNTO. Es una materia prima
              Cantidad = R!FactorConversion
              If Cantidad <> 1 Then
                    

                    Cantidad = R!stock    'Esta es la cantidad nueva
                    campo = TransformaComasPuntos(CStr(Cantidad))

       
                        miSql = "INSERT INTO tmpstockfec2(codusu,codartic,codalmac,stock)  VALUES (" & vUsu.Codigo & "," & DBSet(R!codartic, "T")
                        miSql = miSql & ",0," & campo & ")"
                        
                   
                    Conn.Execute miSql
                End If
           End If
           R.MoveNext
    Wend
    R.Close
    
End Sub




Private Function NuevasComprobacionesContabilizacion(Proveedores As Boolean, ByVal SQL As String) As Boolean
Dim RT As ADODB.Recordset
Dim C As String
Dim F As Date
Dim Fin As Boolean
Dim ComprobacionFechaMenor As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo ENuevasComprobacionesContabilizacion
    NuevasComprobacionesContabilizacion = False
    
    
    
    Set cControlFra = New CControlFacturaContab
        'Tenemos que comprobar la fecha factura
    Set RT = New ADODB.Recordset
    ComprobacionFechaMenor = False

    If Proveedores Then
        C = "select fecrecep from scafpc WHERE " & SQL
        C = C & " GROUP BY fecrecep ORDER BY fecrecep"
    Else
        C = "Select fecfactu from scafac WHERE " & SQL
        C = C & " GROUP BY fecfactu ORDER BY fecfactu"
    End If
    
    
    RT.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Fin = False
    While Not Fin
        F = RT.Fields(0)
        C = cControlFra.FechaCorrectaContabilizazion(ConnConta, F)
        If C <> "" Then
            Fin = True
        Else
            C = cControlFra.FechaCorrectaIVA(ConnConta, F)
            If C <> "" Then
                Fin = True
            Else
                If Proveedores Then
                    'Solo compruebo una vez
                    If Not ComprobacionFechaMenor Then
                        If cControlFra.FechaRecepMenorQueProveedor(ConnConta, F) Then
                            C = "Factura contabilizada con fecha de recepcion menor"
                            Fin = True
                        End If
                            
                        ComprobacionFechaMenor = True
                    End If
                End If
            End If
        End If
        RT.MoveNext
        If Not Fin Then Fin = RT.EOF
    Wend
    RT.Close
    
    If C <> "" Then
        C = C & "(" & F & ")"
        MsgBox C, vbExclamation
    Else
        NuevasComprobacionesContabilizacion = True
    End If
    
    
ENuevasComprobacionesContabilizacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Nueva Comprobacion Contabilizacion"
    Set RT = Nothing
    Set cControlFra = Nothing
End Function

