VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallEntradaOliva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entradas de olivas"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11475
   ClipControls    =   0   'False
   Icon            =   "frmVallEntradaOliva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Productos"
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
      Height          =   4815
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Width           =   11175
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   9840
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   9240
         MaxLength       =   12
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   8760
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   2880
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text2"
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text2"
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text2"
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   8760
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text2"
         Top             =   2880
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   8760
         MaxLength       =   12
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text2"
         Top             =   2400
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   1920
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   8760
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   1440
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   8760
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   4920
         MaxLength       =   12
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   600
         Width           =   3165
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   720
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmVallEntradaOliva.frx":000C
         Height          =   3975
         Left            =   360
         TabIndex        =   44
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VB.Label Label1 
         Caption         =   "Real"
         Height          =   195
         Index           =   10
         Left            =   10080
         TabIndex        =   74
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Rendimiento"
         Height          =   195
         Index           =   15
         Left            =   8760
         TabIndex        =   73
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "% Hoja"
         Height          =   255
         Index           =   14
         Left            =   7320
         TabIndex        =   72
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Total KG"
         Height          =   255
         Index           =   13
         Left            =   9480
         TabIndex        =   71
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Neto"
         Height          =   255
         Index           =   12
         Left            =   6240
         TabIndex        =   70
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tara envases"
         Height          =   195
         Index           =   11
         Left            =   4920
         TabIndex        =   69
         Top             =   3840
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   4800
         ToolTipText     =   "Buscar cliente"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   4800
         ToolTipText     =   "Buscar cliente"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   4800
         ToolTipText     =   "Buscar cliente"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   4800
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   4800
         ToolTipText     =   "Buscar cliente"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos"
         Height          =   195
         Index           =   9
         Left            =   10080
         TabIndex        =   61
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "Peso"
         Height          =   195
         Index           =   8
         Left            =   8880
         TabIndex        =   60
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Uds"
         Height          =   195
         Index           =   7
         Left            =   8160
         TabIndex        =   59
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto ticket"
         Height          =   195
         Index           =   6
         Left            =   3720
         TabIndex        =   58
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Envases"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   57
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Width           =   11175
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   720
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Tara|N|N|0||vallentradacamion|Tara|#,##0||"
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "Bruto|N|N|0||vallentradacamion|Bruto|#,##0||"
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   0
         Tag             =   "Codigo|N|N|0||vallentradacamion|entrada|0000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha|F|N|||vallentradacamion|FechaEntrada|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Proveedor|N|N|0||vallentradacamion|codprove|0000||"
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   720
         Width           =   3645
      End
      Begin VB.CheckBox chkCerrado 
         Caption         =   "Generación finalizada"
         Height          =   195
         Left            =   8520
         TabIndex        =   40
         Tag             =   "Fin|N|N|0||vallentradacamion|EntradaFinalizada|||"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Neto"
         Height          =   255
         Index           =   2
         Left            =   9240
         TabIndex        =   54
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Tara"
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   51
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Bruto"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   50
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3360
         Picture         =   "frmVallEntradaOliva.frx":0021
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha "
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar cliente"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Transporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1335
      Left            =   120
      TabIndex        =   38
      Top             =   1680
      Width           =   11175
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Empresa|N|S|||vallentradacamion|EmpresaTransporte|||"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   7200
         MaxLength       =   100
         TabIndex        =   7
         Tag             =   "Ve|T|S|||vallentradacamion|Conductor|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "Ve|T|S|||vallentradacamion|Matricula|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   12
         TabIndex        =   5
         Tag             =   "Ve|T|S|||vallentradacamion|TipoVehiculo|||"
         Text            =   "   "
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image imgEmpresasTransporte 
         Height          =   240
         Left            =   5400
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Conductor"
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Matrícula"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Vehiculo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   30
      Top             =   8130
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9555
      TabIndex        =   31
      Top             =   8130
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9555
      TabIndex        =   32
      Top             =   8130
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   360
      TabIndex        =   36
      Top             =   8040
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   180
         Width           =   2115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Asignar rendimiento"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar albaranes"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado albaranes"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   9720
         TabIndex        =   35
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   4560
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmVallEntradaOliva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmComProveedores  'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmVallTransOliva
Attribute frmT.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
Dim ModificaLineas As Byte
  
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean
Private Mc As CTiposMov




Private Sub cboEmpresa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCerrado_Click()
'    If Modo = 3 Or Modo = 4 Then 'Insertar o Modificar
'        If Me.chkPermiteDto.Value = 1 Then
'            Me.Text1(4).Text = ""
'            BloquearTxt Text1(4), True
'        Else
'            BloquearTxt Text1(4), False
'        End If
'    End If
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim SituarData As Integer

On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE entrada=" & Text1(3).Text & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonAlbaranes
                BotonAnyadirLinea True
            End If
        End If
        
    Case 4 'MODIFICAR
        If DatosOk Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData2
                 PonerModo 2
             End If
         End If
         
    Case 5
        If DatosOkLinea Then
           If InsertarModificarLinea Then
                
                If ModificaLineas = 1 Then Mc.IncrementarContador Mc.TipoMovimiento
                    
                
                DataGrid1.AllowAddNew = False
                SituarData = Data2.Recordset.AbsolutePosition
                If SituarData < 0 Then SituarData = 0
                CargaGrid True
                SituarDataPosicion Data2, CLng(SituarData), Me.lblIndicador.Caption
                        
                        
                b = True
                If ModificaLineas = 1 Then b = Not BotonAnyadirLinea(True)
                
                If b Then cmdCancelar_Click
                
           End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        Case 5
        
        
            ModificaLineas = 0
            TerminaBloquear
            Set Mc = Nothing
                
            DataGrid1.AllowAddNew = False
           
            PonerDatosForaGrid True
            LLamaLineas2 0
            kCampo = Data2.Recordset.AbsolutePosition
            If kCampo < 0 Then QueRegistro
            CargaGrid True
            SituarDataPosicion Data2, CLng(kCampo), Me.lblIndicador.Caption
            'If Not Data4.Recordset.EOF Then PonerCamposAlmacenes2
            
            PonerBotonCabecera True
            PonerFocoBtn Me.cmdRegresar
        
        
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub QueRegistro()
    On Error Resume Next
    kCampo = Data2.Recordset.RecordCount
    If kCampo < 0 Then kCampo = 1
    Err.Clear
End Sub


Private Sub cmdRegresar_Click()
    If Modo = 5 Then
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
        PonerModo 2
    End If
End Sub

Private Sub Data2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Pintar As Boolean
    
    
    If Not Data2.Recordset.EOF Then
        If Modo = 5 And ModificaLineas = 1 Then
            Pintar = False
        Else
            If Not PrimeraVez Then Pintar = True
        End If
    Else
        Pintar = False
    End If
    If Pintar Then
        PonerDatosForaGrid False
    Else
        PonerDatosForaGrid True
    End If
End Sub

'Private Sub cmdRegresar_Click()
''Este es el boton Cabeceraº
'Dim cad As String
'Dim Indicador As String
'
'    'Quitar lineas y volver a la cabecera
'    If Modo = 5 Then 'modo 5: Lineas Articulos x Almacen
'        DataGrid1.ClearFields
'        cad = "(codmovim=" & Val(Text1(0).Text) & ")"
'        If SituarData(Data1, cad, Indicador) Then
'            PonerModo 2
'            lblIndicador.Caption = Indicador
'            Me.Toolbar1.Buttons(9).Enabled = True
'            Me.Toolbar1.Buttons(10).Enabled = True
'        End If
'    End If
'End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgListComun.ListImages(19).Picture
    Next kCampo
    imgEmpresasTransporte.Picture = frmppal.imgListComun.ListImages(19).Picture
    
    'La toolbar
    btnPrimero = 18 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 10 'Mto Lineas Ofertas
        
        .Buttons(11).Image = 16 'Imprimir
        .Buttons(12).Image = 47 'Generar
        .Buttons(13).Image = 21 'Rendiminetos
        .Buttons(14).Image = 48 'Rendiminetos
        
        .Buttons(15).Image = 15 'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargarCombo_Tabla cboEmpresa, "vallempresatransoliva", "codEmpre", "NomEmpre"
        
    
    NombreTabla = "vallentradacamion" 'Tabla Precios Especiales de Articulos
    Ordenacion = " ORDER BY entrada"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE "
    
    
    CadenaConsulta = CadenaConsulta & " entrada = -1" 'No recupera datos
   
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
        PonerModo 0
        CargaGrid (Modo = 2)
        LLamaLineas2 0
   
    
    'Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim Inicio As Byte
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    DataGrid1.Columns(0).Caption = "Albaran"
    DataGrid1.Columns(0).Width = 1300
    DataGrid1.Columns(0).NumberFormat = "00000"
    DataGrid1.Columns(0).AllowSizing = False
    For Inicio = 1 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(Inicio).visible = False
    Next
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    'For i = 0 To DataGrid1.Columns.Count - 1
    '    DataGrid1.Columns(i).AllowSizing = False
    'Next i
    DataGrid1.Enabled = b
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Articulos
    CadenaConsulta = CadenaSeleccion
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim Indice As Byte
    Indice = 7
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Clientes
    Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim J As Integer
   
    
    If Index = 0 Then
        J = 0
        If Modo = 2 Or Modo = 0 Or Modo = 5 Then J = 1
         
    Else
        'lineas
        J = 1
        If Modo = 5 And ModificaLineas > 0 Then J = 0
        
    End If
    If J = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0  'Cod. Cliente
            Set frmC = New frmComProveedores
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            PonerFoco Text1(Index)
        Case 1 To 5 'Codigo Articulo
            CadenaConsulta = ""
            If Index = 1 Then
                J = 1
            Else
                J = 1 + (Index * 3)
            End If
            Set frmA = New frmAlmArticulos
            frmA.DatosADevolverBusqueda2 = "@1@" 'Abre en modo busqueda
            frmA.Show vbModal
            Set frmA = Nothing
            If CadenaConsulta <> "" Then
                Text3(J).Text = RecuperaValor(CadenaConsulta, 1)
                Text2(Index + 1).Text = ""
                Text3_LostFocus J
            End If
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgEmpresasTransporte_Click()
    If Modo = 2 Or Modo = 0 Then Exit Sub
    CadenaConsulta = ""
    Set frmT = New frmVallTransOliva
    frmT.DatosADevolverBusqueda = "0|"
    frmT.Show vbModal
    Set frmT = Nothing
    If CadenaConsulta <> "" Then
        If RecuperaValor(CadenaConsulta, 1) = 1 Then CargarCombo_Tabla cboEmpresa, "vallempresatransoliva", "codEmpre", "NomEmpre"
        CadenaConsulta = RecuperaValor(CadenaConsulta, 2)
        SituarCombo2 Me.cboEmpresa, Val(CadenaConsulta)
        CadenaConsulta = ""
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass

   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Indice = 7
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
Dim anc

    If Modo = 5 Then
        ModificaLineas = 2
        PonerDatosForaGrid False
        PonerBotonCabecera False
        
        anc = ObtenerAlto(DataGrid1, 20)
        LLamaLineas2 CSng(anc)
        PonerFoco Text3(1)

    Else
        'If BLOQUEADesdeFormulario(Me) Then
        BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then
        BotonAnyadirLinea False
    Else
        BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
     BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo, cadkey
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
        
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim campo As String
Dim Tabla As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    Select Case Index
        Case 0 'Codigo proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                If Text2(Index).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2
           
        Case 5, 6 'Pesos Dem momento entero
            If Text1(Index).Text <> "" Then
                 Text1(Index).Text = Replace(Text1(Index).Text, ".", "")
                If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
            End If
            Neto
        Case 7 'Fecha Cambio
            PonerFormatoFecha Text1(Index)
            
        Case 4
            Text1(Index).Text = UCase(Text1(Index))
    End Select
End Sub


Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), IIf(Modo = 5, 3, Modo)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
Dim CalcularTotales As Boolean
Dim cad As String
Dim i As Byte
Dim L As Integer

   ' If Modo <> 5 Then Exit Sub
    
    Text3(Index).Text = UCase(Trim(Text3(Index).Text))
    
    CalcularTotales = False
    Select Case Index
    Case 0
        'Num. albaran
        
    Case 1
        'Producto BASE (OLIVA)
        If Text3(Index).Text = "" Then
            cad = ""
        Else
            cad = "sartic.codfamia=sfamia.codfamia and codartic"
            CadenaConsulta = "tipfamia"
            cad = DevuelveDesdeBD(conAri, "nomartic", "sartic,sfamia ", cad, Text3(Index).Text, "T", CadenaConsulta)
            If cad = "" Then
                MsgBox "No existe el articulo: " & vbCrLf, vbExclamation
            Else
                If CadenaConsulta <> "30" Then
                    MsgBox "Producto NO es oliva", vbExclamation
                    cad = ""
                End If
            End If
        End If
        i = 2
        If cad = "" Then
            Text3(Index).Text = ""
            Text2(i).Text = ""
        Else
            Text2(i).Text = cad
        End If
        
    Case 7, 10, 13, 16
        'ENVASES
        
          'Producto BASE (OLIVA)
        If Text3(Index).Text = "" Then
            cad = ""
            CadenaConsulta = ""
            If Index > 7 Then PonerFoco Text3(2)
        Else
            cad = "sartic left join sarti4 on sartic.codartic=sarti4.codartic"
            CadenaConsulta = "pesobruto"
            cad = DevuelveDesdeBD(conAri, "nomartic", cad, "sartic.codartic", Text3(Index).Text, "T", CadenaConsulta)
            If cad = "" Then
                MsgBox "No existe el articulo: " & vbCrLf, vbExclamation
                CadenaConsulta = ""
            Else
               If CadenaConsulta <> "" Then CadenaConsulta = Format(CCur(CadenaConsulta), "#,##0")
            End If
        End If
        
        i = IIf(Index = 7, 3, IIf(Index = 10, 4, IIf(Index = 13, 5, 6)))
        If cad = "" Then Text3(Index).Text = ""
        Text2(i).Text = cad
                
                
        
        'I = IIf(Index = 7, 2, IIf(Index = 10, 3, IIf(Index = 13, 4, 5)))
        Text3(Index + 2).Text = CadenaConsulta
        
        CalcularLineaPesos i - 2
        
    Case 8, 9, 11, 12, 14, 15, 17, 18
        'UDS , peso (kg entero)
        If Text3(Index).Text <> "" Then
            Text3(Index).Text = Replace(Text3(Index).Text, ".", "")
            If Not PonerFormatoEntero(Text3(Index)) Then Text3(Index).Text = ""
        End If
        i = IIf(Index < 10, 1, IIf(Index < 13, 2, IIf(Index < 16, 3, 4)))
        CalcularLineaPesos CInt(i)
        CalcularTotales = True
        
    Case 2, 3, 4
        'PESOS , entero en KG
        If Text3(Index).Text <> "" Then
            Text3(Index).Text = Replace(Text3(Index).Text, ".", "")
            If Not PonerFormatoEntero(Text3(Index)) Then
                Text3(Index).Text = ""
            Else
                Text3(Index).Text = Format(Text3(Index).Text, "#,##0")
            End If
                
        End If
        CalculoSobreTotales
    Case 5, 19, 20
        '% Hoja y rendimiento
        If Text3(Index).Text <> "" Then
            If Not PonerFormatoDecimal(Text3(Index), 4) Then Text3(Index).Text = ""
        End If
        
        If Index = 5 Then CalculoSobreTotales
    End Select
    
End Sub
Private Sub CalcularLineaPesos(linea As Integer)
Dim Peso As Long

Dim cad As String
Dim K As Integer
    
    K = IIf(linea = 1, 8, IIf(linea = 2, 11, IIf(linea = 3, 14, 17)))
    If Me.Text3(K).Text = "" Or Me.Text3(K + 1) = "" Then
        cad = ""
    Else
        Peso = Val(Replace(Text3(K).Text, ".", ""))
        K = Val(Replace(Text3(K + 1).Text, ".", ""))
        Peso = Peso * K
        cad = Format(Peso, "#,##0")
    End If
    Text2(linea + 6) = cad
    
    Peso = 0
    cad = ""
    For K = 7 To 10
        If Text2(K).Text <> "" Then Peso = Peso + Val(Replace(Text2(K).Text, ".", ""))
    Next K
    If Peso <> 0 Then cad = Format(Peso, "#,##0")
    Text3(3).Text = cad
    
    CalculoSobreTotales
End Sub



Private Sub CalculoSobreTotales()
Dim J As Integer
Dim T2 As Currency
Dim T3 As Currency

    
    T2 = 0
    T3 = 0
    
    If Text3(2).Text <> "" Then T2 = ImporteFormateado(Text3(2).Text)
    If Text3(3).Text <> "" Then T3 = ImporteFormateado(Text3(3).Text)
    
    T2 = T2 - T3
    Me.Text3(4).Text = Format(T2, "#,##0")
    
    If Me.Text3(4).Text = "" Then
        Me.Text3(6).Text = Text3(4).Text
    Else
        J = Val(Replace(Text3(5).Text, ".", ""))
        J = Int(Round2((J * T2) / 100, 0))
        NumRegElim = T2 - J
        If NumRegElim = 0 Then
            Text3(6).Text = ""
        Else
            Text3(6).Text = Format(NumRegElim, "#,##0")
        End If
    End If
        
        
        
    
    
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        
        Case 9
            BotonAlbaranes
        
              
                
        Case 11, 12, 13
            If Modo <> 2 Then Exit Sub
            If Data2.Recordset.EOF Then Exit Sub
            
            Select Case Button.Index
            Case 11
                CadenaDesdeOtroForm = Data1.Recordset!entrada & "|" & Text1(0).Text & " " & Text2(0).Text & "|"
                frmListado2.Opcion = 33
                frmListado2.Show vbModal
                CadenaDesdeOtroForm = ""
            
            Case 12
                'Rendimientos
            
            
            
            Case 13
                If Modo <> 2 Then Exit Sub
                If Data2.Recordset.EOF Then Exit Sub
                If Val(Data1.Recordset!EntradaFinalizada) = 1 Then
                    MsgBox "Albaranes YA generados!!", vbExclamation
                    Exit Sub
                End If
                
                CadenaDesdeOtroForm = Data1.Recordset!entrada
                frmListado2.Opcion = 34
                frmListado2.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    'TerminaBloquear
                    CadenaConsulta = "Select * from " & NombreTabla & " WHERE entrada=" & Text1(3).Text & Ordenacion
                    Data1.Refresh
                    PosicionarData2
                    
                End If
            
            End Select
        Case 14
            'Imprimir albaranes gnerados con kilos por albaran
            If Modo <> 2 And Modo <> 0 Then Exit Sub
            frmListado2.Opcion = 36
            frmListado2.Show vbModal
            
        Case 15  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte


    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5


    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    
    
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearTxt Text1(3), Modo <> 1
    BloquearCmb Me.cboEmpresa, Modo = 0 Or Modo = 2
     
    BloquearChecks Me, IIf(Modo = 1, 1, 2)
    
    If Modo <> 5 Then cmdRegresar.visible = False
           
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    Me.imgBuscar(0).Enabled = b
    Me.imgFecha(0).Enabled = b
    chkCerrado.Enabled = Modo = 1
    
    Frame3.Enabled = Not (Modo = 3 Or Modo = 4)
    
    'Buscar cta en lineas
    b = False
    If Modo = 5 Then
        'Si no esta cerrado
        
        If ModificaLineas > 0 Then b = True
        
    
    End If
    
    For i = 1 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    b = False
    If Modo = 2 Then
        b = True
    Else
        If Modo = 5 And ModificaLineas = 0 Then b = True
    End If
    DataGrid1.Enabled = b
    
    
    
    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2) Or Modo = 5
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '===============================
    b = (Modo = 3 Or Modo = 4)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    
    'Buscar
    b = Modo >= 3
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkCerrado.Value = 0
    Me.cboEmpresa.ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub




Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "NumAlbar , codartic, codAlmac, bruto, tara, Neto, codartic, bruto, tara, Neto, rendimiento, PorcHoja,pesoprod,"
    SQL = SQL & " codarti1,udArti1,pesoArti1,codarti2,udArti2,pesoArti2,"
    SQL = SQL & " codarti3,udArti3,pesoArti3,codarti4,udArti4,pesoArti4,rdtoRea"
    
    SQL = "SELECT " & SQL
    SQL = SQL & " FROM vallentradacamionlineas"
    If enlaza Then
        SQL = SQL & " WHERE entrada=" & Data1.Recordset!entrada
    Else
        SQL = SQL & " WHERE entrada = -1"
    End If
    SQL = SQL & " ORDER BY numalbar "
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    Text1(3).Text = SugerirCodigoSiguienteStr(NombreTabla, "entrada")
    Text1(7).Text = Format(Now, "dd/mm/yyyy")
    SituarCombo2 Me.cboEmpresa, 0  'Cero es el valore por defecto
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    If Val(DBLet(Data1.Recordset!EntradaFinalizada, "N")) = 1 Then
        MsgBox "Albaranes generados", vbExclamation
        Exit Sub
    End If
    
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()

    If Modo = 5 Then
        BotonEliminarLinea
    Else
        BotonEliminar2
    End If

End Sub

Private Sub BotonEliminar2()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Val(DBLet(Data1.Recordset!EntradaFinalizada, "N")) = 1 Then
        MsgBox "Albaranes generados", vbExclamation
        Exit Sub
    End If
    
    SQL = "Entrada camión." & vbCrLf
    SQL = SQL & "--------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a eliminar la entrada de oliva:"
    SQL = SQL & vbCrLf & "Fecha : " & Text1(7).Text
    SQL = SQL & vbCrLf & "Proveedor : " & Text1(0).Text & " " & Text2(0).Text
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        'DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            'MsgBox Err.Number & " : " & Err.Description, vbExclamation
            MuestraError Err.Number, "Eliminar Precio Especial", Err.Description
            Data1.Recordset.CancelUpdate
        End If
End Sub



Private Sub BotonEliminarLinea()
    If Data2.Recordset.EOF Then Exit Sub
    CadenaConsulta = "Número: " & Data2.Recordset!NumAlbar & vbCrLf
    CadenaConsulta = CadenaConsulta & "Producto:   " & Text3(1).Text & " - " & Text2(2).Text & vbCrLf
    CadenaConsulta = CadenaConsulta & "Peso ticket:  " & Text3(2).Text
    CadenaConsulta = "Va a eliminar el albarán:" & vbCrLf & CadenaConsulta
    If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    CadenaConsulta = "DELETE FROM vallentradacamionlineas where entrada = " & Data1.Recordset!entrada
    CadenaConsulta = CadenaConsulta & " AND numalbar = " & Data2.Recordset!NumAlbar
    If EjecutaSQL(conAri, CadenaConsulta, True) Then
        'Intento devolver contador
        Set Mc = New CTiposMov
        Mc.Leer "PES"
        Mc.DevolverContador Mc.TipoMovimiento, CLng(Val(Data2.Recordset!NumAlbar))
        Set Mc = Nothing
    
        CargaGrid True
    End If

End Sub

Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
         
        SQL = " WHERE entrada=" & Val(Data1.Recordset!entrada)
        
        'Lineas
        conn.Execute "Delete  from vallentradacamionlineas " & SQL
        
        'Cabeceras
        conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
On Error Resume Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    
    If ImporteFormateado(Text1(6).Text) > ImporteFormateado(Text1(5).Text) Then
        MsgBox "La tara no puede ser mayor o igual que el bruto", vbExclamation
        Exit Function
    End If
    
    'Fecha activa.
    'Puesta por  para la VALL. Al resto sera 01/01/1900
    If CDate(Text1(7).Text) < vParamAplic.FechaActiva Then
        MsgBox "Periodo de produccion cerrado", vbExclamation
        Exit Function
    End If
    
    
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 10, "Cliente")
    cad = cad & "Nombre Cliente|sclien|nomclien|T||36·"
    cad = cad & ParaGrid(Text1(1), 15, "Cod. Artic")
    cad = cad & "Desc. Artic|sartic|nomartic|T||38·"
    
    Tabla = "(" & NombreTabla & " LEFT JOIN sclien ON " & NombreTabla & ".codclien=sclien.codclien" & ")"
    Tabla = Tabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic"
    
    Titulo = "Precios Especiales"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    'Poner el nombre del cod. cliente
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sprove", "nomprove")
    Neto

    
    
    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub BotonActualizar()
'Actualizar Precios Especiales
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún Precio Especial para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
   
    SQL = "Actualización Precios Especiales de Artículos." & vbCrLf
    SQL = SQL & "---------------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Actualizar el Precio Especial para:"
    SQL = SQL & vbCrLf & " Cod. Clien. :  " & CStr(Format(Data1.Recordset.Fields(0), "000000"))
    SQL = SQL & vbCrLf & " Cod. Artic. :  " & Data1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
        Exit Sub
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarPreEspecial Then
        SituarDataTrasEliminar Data1, NumRegElim
    End If
End Sub


Private Function ActualizarPreEspecial() As Boolean
'Actualiza los Precios Especiales insertando los precios actuales con la fecha de cambio en el hostórico
' y modificando el la tabla de precios especiales pasando los valores nuevos a ser los actuales.
Dim Donde As String
Dim SQL As String
Dim bol As Boolean
On Error GoTo EActualizarPreEspecial
    
   
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarElPrecio(Donde)

EActualizarPreEspecial:
        If Err.Number <> 0 Then
            SQL = "Actualizar Precio Especial." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
'            If OpcionActualizar = 1 Then
                MuestraError Err.Number, SQL, Err.Description
'            Else
'                SQL = Donde & " -> " & Err.Description
'                SQL = Mid(SQL, 1, 200)
'                InsertaError SQL
'            End If
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            ActualizarPreEspecial = True
        Else
            conn.RollbackTrans
            ActualizarPreEspecial = False
        End If
End Function


Private Function ActualizarElPrecio(ByRef ADonde As String) As Boolean

    ActualizarElPrecio = False
    
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Precios Especiales"
    If Not InsertarLineasHistorico Then Exit Function
'    IncrementarProgres 2
    
    
    'Modificamos en cabeceras de Tarifas
    ADonde = "Modificando datos en cabecera de Precios Especiales"
    If Not ModificarCabecera Then Exit Function
'    IncrementarProgres 2
    ActualizarElPrecio = True
End Function


Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim SQL As String

    On Error Resume Next

    SQL = "UPDATE " & NombreTabla & " SET precioac=precionu, precioa1=precion1, dtoespec=dtoespe1, fechanue=null, precionu=0, precion1=0"
    SQL = SQL & " WHERE codclien=" & Data1.Recordset!CodClien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
    conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
    Else
        ModificarCabecera = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim SQL As String
Dim NumF As String
On Error Resume Next

    'Obtenemos la siguiente numero de linea de tarifa
    SQL = "codclien=" & Data1.Recordset!CodClien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
    NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", SQL)

    SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec)"
    SQL = SQL & " VALUES (" & Data1.Recordset.Fields(0).Value & ", " & DBSet(Data1.Recordset.Fields(1).Value, "T") & ", "
    SQL = SQL & NumF & ", " & DBSet(Text1(4).Text, "F") & ", "
    SQL = SQL & DBSet(Data1.Recordset!precioac, "N") & ", " & DBSet(Data1.Recordset!precioa1, "N") & ", "
    SQL = SQL & DBSet(Data1.Recordset!dtoespec, "N") & ") "
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Sub BotonImprimir()
        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData2()
Dim vWhere As String

    vWhere = "entrada=" & Text1(3).Text
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.Refresh
    Data1.Recordset.Find vWhere
    PonerCampos
End Sub

Private Sub Neto()
Dim I1 As Currency
Dim i2 As Currency
    I1 = 0
    i2 = 0
    If Modo <> 1 Then
        If Me.Text1(5).Text <> "" Then I1 = ImporteFormateado(Text1(5).Text)
        If Me.Text1(6).Text <> "" Then i2 = ImporteFormateado(Text1(6).Text)
    End If
    I1 = I1 - i2
    If I1 = 0 Then
        Text2(1).Text = ""
    Else
        Text2(1).Text = Format(I1, "#,##0")
    End If
End Sub

Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas "
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub




'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'  Fora grid
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------


Private Sub BotonAlbaranes()

    If vUsu.Nivel > 0 Then Exit Sub
    If Modo <> 2 Then Exit Sub
    
    
    If Val(DBLet(Data1.Recordset!EntradaFinalizada, "N")) = 1 Then
        MsgBox "Albaranes generados", vbExclamation
        Exit Sub
    End If
        
    
    
    
    
    
    
    PonerModo (5)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    
    
  
    
    
    
    
    
    
    
End Sub


Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim i As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data2.Recordset Is Nothing) Then
            If Not Data2.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        For i = 2 To Text2.Count - 1
            Text2(i).Text = ""
        Next i
       
        
    Else
       
        
        Text3(1).Text = CStr(Data2.Recordset!codartic)
        Text3(7).Text = DBLet(Data2.Recordset!codarti1, "T")
        Text3(10).Text = DBLet(Data2.Recordset!codarti2, "T")
        Text3(13).Text = DBLet(Data2.Recordset!codarti3, "T")
        Text3(16).Text = DBLet(Data2.Recordset!codarti4, "T")
         
        Text3(2).Text = DBLet(Data2.Recordset!bruto, "T")
        Text3(3).Text = DBLet(Data2.Recordset!TARA, "T")
        Text3(4).Text = DBLet(Data2.Recordset!Neto, "T")
        Text3(5).Text = DBLet(Data2.Recordset!PorcHoja, "T")
        Text3(6).Text = DBLet(Data2.Recordset!pesoprod, "T")
        
        Text3(8).Text = DBLet(Data2.Recordset!udArti1, "T")
        Text3(9).Text = DBLet(Data2.Recordset!pesoArti1, "T")
        Text3(11).Text = DBLet(Data2.Recordset!udArti2, "T")
        Text3(12).Text = DBLet(Data2.Recordset!pesoArti2, "T")
        Text3(14).Text = DBLet(Data2.Recordset!udArti3, "T")
        Text3(15).Text = DBLet(Data2.Recordset!pesoArti3, "T")
        Text3(17).Text = DBLet(Data2.Recordset!codarti4, "T")
        Text3(18).Text = DBLet(Data2.Recordset!pesoArti4, "T")
        Text3(19).Text = DBLet(Data2.Recordset!rendimiento, "T")
        Text3(20).Text = DBLet(Data2.Recordset!rdtoRea, "T")
        
        
        
        
        For i = 1 To Text3.Count - 1
            Text3_LostFocus i
        Next i
        
    End If
End Sub

Private Sub LLamaLineas2(alto As Single)
Dim b As Boolean
Dim K As Integer
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    
        DeseleccionaGrid Me.DataGrid1
        If b Then
            Text3(0).Height = DataGrid1.RowHeight
            Text3(0).Left = DataGrid1.Columns(0).Left + 340
            Text3(0).Width = DataGrid1.Columns(0).Width
            Text3(0).Top = alto
            
        End If
        Text3(0).visible = b And ModificaLineas = 1
         
        If b Then
            BloquearTxt Text3(0), True
            For K = 1 To Text3.Count - 1
                If OpcionesText3(0, K) Then
                    BloquearTxt Text3(K), Not b
                Else
                    BloquearTxt Text3(K), True
                End If
            Next
        Else
            For K = 0 To Text3.Count - 1
                BloquearTxt Text3(K), True
            Next
        End If
        For K = 1 To 5
            Me.imgBuscar(K).Enabled = b
        Next
    
End Sub
'Para saber si un text 3 es editable, requerido...
Private Function OpcionesText3(Opcion As Byte, Indice As Integer) As Boolean
    
    If Opcion = 0 Then
        'Campos editables
        Select Case Indice
        Case 1, 2, 5, 7, 8, 10, 11, 13, 14, 16, 17
            OpcionesText3 = True
        Case Else
            OpcionesText3 = False
        End Select
    ElseIf Opcion = 1 Then
        'DAtosOK lin
        Select Case Indice
        Case 0, 2, 4, 6
            OpcionesText3 = True
        Case Else
            OpcionesText3 = False
        End Select
    Else
        'Campos texto
        Select Case Indice
        Case 1, 7, 10, 13, 16
            OpcionesText3 = True
        Case Else
            OpcionesText3 = False
        End Select
    End If
End Function

Private Function BotonAnyadirLinea(DesdeInsercion2 As Boolean) As Boolean
Dim anc As Single
Dim PrimeraLinea As Boolean

        BotonAnyadirLinea = False
        ModificaLineas = 1
        PrimeraLinea = Data2.Recordset.EOF
        NumRegElim = 0
        If Not PrimeraLinea Then
            CadenaConsulta = DevuelveDesdeBD(conAri, "sum(bruto)", "vallentradacamionlineas", "entrada", Data1.Recordset!entrada)
            anc = Data1.Recordset!bruto - Data1.Recordset!TARA
            'Anc son los kilos totales de carga
            'Le quitamos lo que suman los anterior  tenemos el disponible
            anc = anc - Val(CadenaConsulta)
            If anc <= 0 Then
                If Not DesdeInsercion2 Then
                    MsgBox "Carga excede de la capacidad total", vbExclamation
                Else
                    Exit Function
                End If
            Else
                NumRegElim = anc
            End If
        End If
        
        
        PonerDatosForaGrid True
        PonerBotonCabecera False
        AnyadirLinea DataGrid1, Data2
        anc = ObtenerAlto(DataGrid1, 20)
        LLamaLineas2 anc
        PonerFoco Text3(1)
        Text3(4).Text = "0"
        Set Mc = New CTiposMov
        Mc.Leer "PES"
        
        Text3(0).Text = Format(Mc.ConseguirContador(Mc.TipoMovimiento), "00000")
        If NumRegElim > 0 Then Text3(2).Text = Format(NumRegElim, "#,##0")
        
       ' If vUsu.Login = "root" Then BloquearTxt Text3(0), False
        BotonAnyadirLinea = True
End Function


Private Function DatosOkLinea() As Boolean
    
    DatosOkLinea = False
    CadenaDesdeOtroForm = ""
    
    
    
    If Text3(3).Text = "" Then Text3(3).Text = "0"
            
    
    For NumRegElim = 0 To Text3.Count - 1
        
        If OpcionesText3(1, CInt(NumRegElim)) Then
            If Text3(NumRegElim).Text = "" Then
                If NumRegElim = 0 And ModificaLineas = 2 Then
                    CadenaConsulta = "NO"  'Modificando NO valor el numalbar
                Else
                    CadenaConsulta = ""
                End If
                If CadenaConsulta = "" Then
                    MsgBox "Campo requerido", vbExclamation
                    PonerFoco Text3(NumRegElim)
                    Exit Function
                End If
            End If
        End If
    Next
    
    NumRegElim = Data1.Recordset!bruto - Data1.Recordset!TARA  'Disponible
    CadenaConsulta = DevuelveDesdeBD(conAri, "sum(bruto)", "vallentradacamionlineas", "entrada", Data1.Recordset!entrada)
    If ModificaLineas = 2 Then
        'Esta modificando la linea. Por lo tanto le resto la cantidad en data2
        CadenaConsulta = Val(CadenaConsulta) - DBLet(Data2.Recordset!pesoprod, "N")
    Else
        CadenaConsulta = "0"
    End If
    NumRegElim = NumRegElim - Val(CadenaConsulta)
    NumRegElim = Data1.Recordset!bruto - Data1.Recordset!TARA - ImporteFormateado(Text3(6).Text)
    
    
    'Si no ha indicado peso camion (ni tara) entonces no hace sumatorio pesos
    If Not (Data1.Recordset!bruto = 0 And Data1.Recordset!TARA = 0) Then
        If NumRegElim < 0 Then
            'MALLLL, execede
            MsgBox "Excede del peso maximo", vbExclamation
            Exit Function
        End If
    End If
    
    '7, 10, 13, 16
    For NumRegElim = 1 To 4
        kCampo = CInt(RecuperaValor("7|10|13|16|", CInt(NumRegElim)))
        If Text3(kCampo).Text <> "" Then
            CadenaConsulta = DevuelveDesdeBD(conAri, "tipartic", "sartic", "codartic", Text3(kCampo).Text, "T")
            If CadenaConsulta <> "31" Then
                MsgBox "Articulos no es de palets: " & Text3(kCampo).Text, "T", vbExclamation
                Exit Function
            End If
        End If
    Next
    
    
    
    
    'Sumatorios
    DatosOkLinea = True
End Function

Private Function ColumnasSQL() As String
    ColumnasSQL = "NumAlbar , codartic,  bruto, tara, Neto,PorcHoja,pesoprod,  "
    ColumnasSQL = ColumnasSQL & "codarti1,udArti1,pesoArti1,codarti2,udArti2,pesoArti2,"
    ColumnasSQL = ColumnasSQL & "codarti3,udArti3,pesoArti3,codarti4,udArti4,pesoArti4,rendimiento,rdtoRea"
End Function

Private Function InsertarModificarLinea() As Boolean
Dim SQL As String
    
    CadenaConsulta = ColumnasSQL & "|"
    CadenaConsulta = Replace(CadenaConsulta, ",", "|")
    SQL = ""

    If ModificaLineas = 1 Then
        
        For NumRegElim = 0 To Text3.Count - 1
            SQL = SQL & "," & RecuperaValor(CadenaConsulta, CInt(NumRegElim) + 1)
        Next
        SQL = "INSERT INTO vallentradacamionlineas (entrada,codalmac" & SQL & ") VALUES (" & Text1(3).Text & ",1"
        
        CadenaConsulta = ""
        For NumRegElim = 0 To Text3.Count - 1
            CadenaDesdeOtroForm = Text3(NumRegElim)
            If OpcionesText3(2, CInt(NumRegElim)) Then
               'texto
               CadenaDesdeOtroForm = DBSet(CadenaDesdeOtroForm, "T", IIf(OpcionesText3(1, CInt(NumRegElim)), "N", IIf(NumRegElim = 3, "N", "S")))
            Else
                'numero
                CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, ".", "")
                CadenaDesdeOtroForm = DBSet(CadenaDesdeOtroForm, "N", IIf(OpcionesText3(1, CInt(NumRegElim)), "N", IIf(NumRegElim = 3, "N", "S")))
            End If

            SQL = SQL & "," & CadenaDesdeOtroForm
            
        Next
        SQL = SQL & ")"
    Else
        
        
        
        SQL = ""
        For NumRegElim = 1 To Text3.Count - 1
            CadenaDesdeOtroForm = Text3(NumRegElim)
                
            If OpcionesText3(2, CInt(NumRegElim)) Then
               'texto
               CadenaDesdeOtroForm = DBSet(CadenaDesdeOtroForm, "T", IIf(OpcionesText3(1, CInt(NumRegElim)), "N", "S"))
            Else
                'numero
                CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, ".", "")
                CadenaDesdeOtroForm = DBSet(CadenaDesdeOtroForm, "N", IIf(OpcionesText3(1, CInt(NumRegElim)), "N", "S"))
            End If
            SQL = SQL & ", " & RecuperaValor(CadenaConsulta, CInt(NumRegElim) + 1) & " = " & CadenaDesdeOtroForm
            
        Next
        SQL = Mid(SQL, 2)
        SQL = "UPDATE vallentradacamionlineas SET " & SQL & " WHERE entrada =" & Data1.Recordset!entrada
        SQL = SQL & " AND numalbar = " & Data2.Recordset!NumAlbar
        
        
    End If
    
    InsertarModificarLinea = EjecutaSQL(conAri, SQL, True)
End Function
