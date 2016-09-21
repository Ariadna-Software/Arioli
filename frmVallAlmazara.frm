VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallAlmazara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro del control de proceso en almazara"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   17415
   Icon            =   "frmVallAlmazara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   60
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   59
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   58
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtaux2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   12360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "frmVallAlmazara.frx":000C
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   24
      Tag             =   "Lote"
      Text            =   "lot"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtaux2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   12480
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   4560
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3135
      Left            =   11880
      TabIndex        =   37
      Top             =   4440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   10
      TabIndex        =   17
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1080
      MaxLength       =   18
      TabIndex        =   18
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   19
      Tag             =   "Nombre Artículo"
      Text            =   "nomArtic"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   34
      ToolTipText     =   "Buscar almacen"
      Top             =   4380
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   17175
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   10920
         MaxLength       =   16
         TabIndex        =   63
         Tag             =   "Loteproducido|T|S|||vallalmazaraproceso|loteproducido||N|"
         Top             =   2760
         Width           =   2265
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   13560
         MaxLength       =   10
         TabIndex        =   61
         Tag             =   "Litros|N|S|||vallalmazaraproceso|litros|#,##0.00|N|"
         Top             =   2760
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   15360
         MaxLength       =   10
         TabIndex        =   56
         Tag             =   "Kilos|N|S|||vallalmazaraproceso|kilos|#,##0.00|N|"
         Top             =   2760
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Dep|N|S|0||vallalmazaraproceso|deposito|00|N|"
         Top             =   2760
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Hora fin|H|S|||vallalmazaraproceso|HoraFin|hh:nn:ss|N|"
         Top             =   2760
         Width           =   1545
      End
      Begin VB.ComboBox cboAlmazara 
         Height          =   315
         Index           =   3
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "Aspecto|N|S|||vallalmazaraproceso|CantidadSobreNada|||"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cboAlmazara 
         Height          =   315
         Index           =   2
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Aspecto|N|S|||vallalmazaraproceso|AspectoMasa|||"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   15360
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Temp. caldera |N|S|||vallalmazaraproceso|AguaDecanter|#0.00|N|"
         Top             =   360
         Width           =   945
      End
      Begin VB.ComboBox cboAlmazara 
         Height          =   315
         Index           =   1
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Aspecto|N|S|||vallalmazaraproceso|AspectoSalida|||"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   13560
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Fecha producción|T|S|||vallalmazaraproceso|numlote|||"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   12120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Dosis |N|S||100|vallalmazaraproceso|Dosis|#0.00|N|"
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Temp. caldera |N|S|||vallalmazaraproceso|TempMasa|#0.00|N|"
         Top             =   1080
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Temp. caldera |N|S|||vallalmazaraproceso|TempAceite|#0.00|N|"
         Top             =   1080
         Width           =   945
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   43
         Text            =   "nomArtic"
         Top             =   1080
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Trabajador almazara|N|N|||vallalmazaraproceso|codtrabaAlm|0000|N|"
         Top             =   1080
         Width           =   825
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   41
         Text            =   "nomArtic"
         Top             =   360
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Temp. caldera |N|S|||vallalmazaraproceso|TempCaldera|#0.00|N|"
         Top             =   1080
         Width           =   1065
      End
      Begin VB.ComboBox cboAlmazara 
         Height          =   315
         Index           =   0
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Tipo oliva|N|N|||vallalmazaraproceso|tipoOliva|||"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Operario recep|N|N|||vallalmazaraproceso|codtrabaRec|0000||"
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Height          =   1155
         Index           =   3
         Left            =   8880
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Tag             =   "Obs|T|S|||vallalmazaraproceso|Observa|||"
         Top             =   1080
         Width           =   7545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha creación|F|N|||vallalmazaraproceso|fecha|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº proc almazara|N|S|0||vallalmazaraproceso|id|00000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Lote producido"
         Height          =   255
         Index           =   18
         Left            =   10920
         TabIndex        =   64
         Top             =   2490
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Litros aceite"
         Height          =   255
         Index           =   19
         Left            =   13560
         TabIndex        =   62
         Top             =   2490
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos oliva"
         Height          =   255
         Index           =   17
         Left            =   15360
         TabIndex        =   57
         Top             =   2490
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   7320
         Top             =   2400
         Width           =   9615
      End
      Begin VB.Label Label1 
         Caption         =   "Deposito"
         Height          =   255
         Index           =   16
         Left            =   9720
         TabIndex        =   55
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Hora fin"
         Height          =   255
         Index           =   15
         Left            =   7800
         TabIndex        =   54
         Top             =   2490
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad aceite que sobrenada en la batidora"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   53
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Aspecto de la masa en la batidora"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   52
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agua decanter"
         Height          =   195
         Index           =   11
         Left            =   15240
         TabIndex        =   51
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Aspecto del aceite a la salida del decanter"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   50
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Index           =   9
         Left            =   13560
         TabIndex        =   49
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Talco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   11280
         TabIndex        =   48
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dosis"
         Height          =   195
         Index           =   2
         Left            =   12120
         TabIndex        =   47
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Temp. masa"
         Height          =   195
         Index           =   7
         Left            =   7440
         TabIndex        =   46
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Temp. aceite"
         Height          =   195
         Index           =   6
         Left            =   6240
         TabIndex        =   45
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Temp. caldera"
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   44
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Operario almazara"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   1350
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmVallAlmazara.frx":0012
         ToolTipText     =   "Buscar Nº Serie"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo oliva"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6360
         Picture         =   "frmVallAlmazara.frx":0114
         ToolTipText     =   "Buscar Nº Serie"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Operario recepción"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   36
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   14
         Left            =   1590
         TabIndex        =   32
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2520
         Picture         =   "frmVallAlmazara.frx":0216
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   31
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   0
      TabIndex        =   26
      Top             =   7935
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   16080
      TabIndex        =   23
      Top             =   7920
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   14760
      TabIndex        =   22
      Top             =   7920
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   8040
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas albaranes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas observaciones"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar proceso almazara"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7800
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2280
      Top             =   8040
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   16080
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVallAlmazara.frx":02A1
      Height          =   3120
      Left            =   120
      TabIndex        =   33
      Top             =   4440
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5503
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11880
      TabIndex        =   38
      Top             =   4080
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Albaranes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3960
      Width           =   3375
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
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnCambios 
         Caption         =   "Cambios"
         Shortcut        =   ^K
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
Attribute VB_Name = "frmVallAlmazara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1


 

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'   6.-  Modificar cantidad en componentes
'-------------------------------------------------------------------------


Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec



'SQL de la tabla principal del formulario
Private CadenaConsulta As String


Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1






Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange


    
    
Dim TrabajadorConectado As Integer



Private Sub cboAlmazara_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'================================================================================






Private Sub cmdAceptar_Click()
'Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR Cabecera Pedido
            
            If DatosOk Then InsertarCabecera
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                
                If ModificaDesdeFormulario(Me, 1) Then
                    
                    TerminaBloquear
                    PosicionarData
                    
                End If
            End If
            
         Case 5, 6 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Modo = 5 Then
                    If Data2.Recordset.EOF = True Then PrimeraLin = True
                Else
                    If data3.Recordset.EOF = True Then PrimeraLin = True
                End If
                
                If InsertarLinea Then
                    If Modo = 5 Then
                        CargaGrid2 DataGrid1, Data2
                    Else
                        CargaGrid3 True
                        Me.DataGrid2.Enabled = True
                    End If
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
            
                'Para saber si he cambiado cantidad
                If ModificarLinea Then
                    
                    TerminaBloquear
                    If Modo = 5 Then
                        CargaTxtAux False, False
                        CargaGrid2 DataGrid1, Data2
                        Me.DataGrid1.Enabled = True
                    Else
                        CargaTxtAux2 False, False
                        CargaGrid3 True
                        Me.DataGrid2.Enabled = True
                    End If
                    

                End If
                
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            
            If Modo = 5 Then
                CargaTxtAux False, True
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                DataGrid1.Enabled = True
            Else
                CargaTxtAux2 False, True
                DataGrid2.AllowAddNew = False
                If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
                DataGrid2.Enabled = True
            End If
            
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    
    'vamos a mostrar todos los albaranes que no esten ya
    
    'vallalmazaraprocesoalb
    
    MandaBusquedaPrevia ""
    
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            CargaTxtAux False, False
           
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            CargaGrid3 True
            Me.cmdRegresar.Cancel = True
        Case 6
            TerminaBloquear
            CargaTxtAux2 False, False
           
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid2.AllowAddNew = False
                If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid2.Enabled = True
            CargaGrid3 True
            Me.cmdRegresar.Cancel = True
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    'Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    'Text2(3).Text = NomTraba
    
    Me.cboAlmazara(1).ListIndex = 0
    Me.cboAlmazara(2).ListIndex = 0
    Me.cboAlmazara(3).ListIndex = 0
    
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineaAlb()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR "
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
   
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
End Sub

Private Sub BotonAnyadirLineaCambios()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR "
    
    AnyadirLinea DataGrid2, data3
    CargaTxtAux2 True, True
    

    
    PonerFoco txtaux2(0)
    Me.DataGrid2.Enabled = False
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
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

        

    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean

    If Not IsNull(Data1.Recordset!HoraFin) Then
        MsgBox "Proceso cerrado. No se puede modificar", vbExclamation
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        

End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Modo = 5 Then
    
        If Data2.Recordset.EOF Then Exit Sub
        CargaTxtAux True, False
        PonerFoco txtAux(0)
        Me.DataGrid1.Enabled = False
    Else
        If data3.Recordset.EOF Then Exit Sub
        CargaTxtAux2 True, False
        PonerFoco txtaux2(0)
        Me.DataGrid2.Enabled = False
    End If
    
    ModificaLineas = 2 'Modificar
    PonerBotonCabecera False
    Me.lblIndicador.Caption = "MODIFICAR"
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If Not IsNull(Data1.Recordset!HoraFin) Then
        MsgBox "Proceso cerrado. No se puede eliminar", vbExclamation
        Exit Sub
    End If

    Cad = "Almazara." & vbCrLf
    Cad = Cad & "----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el proceso:"
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea continuar? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Modo = 5 Then
        If Data2.Recordset.EOF Then Exit Sub
    Else
        If data3.Recordset.EOF Then Exit Sub
    End If
            
    ModificaLineas = 3 'Eliminar
    
    SQL = "Va a eliminar la linea:    " & vbCrLf
    If Modo = 5 Then
        SQL = SQL & vbCrLf & "Albaran:  " & Data2.Recordset!NumAlbar & " - " & Data2.Recordset!FechaAlb & " - " & Data2.Recordset!nomprove
    Else
        SQL = SQL & vbCrLf & "Obs:  " & data3.Recordset!observa
    End If
    SQL = SQL & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        If Modo = 5 Then
            NumRegElim = Data2.Recordset.AbsolutePosition
            SQL = "DELETE FROM vallalmazaraprocesoalb WHERE numalbar=" & DBSet(Data2.Recordset!NumAlbar, "T")
            SQL = SQL & " AND fechaalb=" & DBSet(Data2.Recordset!FechaAlb, "F")
            SQL = SQL & " AND codprove =" & Data2.Recordset!codProve
        Else
            SQL = "DELETE FROM vallalmazaraprocesocambios WHERE secuencial=" & data3.Recordset!secuencial
            
        End If
        SQL = SQL & " AND id =" & Data1.Recordset!ID
        
        conn.Execute SQL
        
        ModificaLineas = 0
        If Modo = 5 Then
            CargaGrid2 DataGrid1, Data2
            SituarDataPosicion Me.Data2, NumRegElim, SQL
    
        Else
            CargaGrid3 True
            SituarDataPosicion Me.data3, NumRegElim, SQL
        End If
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Or Modo = 6 Then 'modo 5: Mantenimientos Lineas
        PonerModo 2
        
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdCancelar.Cancel = True
        EsCabecera = True
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'cad = Data1.Recordset.Fields(0) & "|"
        'cad = cad & Data1.Recordset.Fields(1) & "|"
        Cad = Data1.Recordset.Fields(0)
        RaiseEvent DatoSeleccionado2(Cad)
        Unload Me
    End If
End Sub



Private Sub DataGrid1_DblClick()
        If Modo = 5 Then  'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            
            
        Else
            
        End If
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)


    On Error GoTo Error1

'    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
'
'    End If
'
    If Modo = 2 Or Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            'Poner descripcion de ampliacion lineas
            CargaGrid3 True
        Else
            
        End If
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub








Private Sub Form_Activate()
    If Me.Tag <> "" Then
        Me.Tag = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 23
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 37 'Cambiar cantidad componentes
        
        'Enero08
        .Buttons(12).Image = 21 'Cerrar orden produccion
        
       
        .Buttons(14).Image = 16 'Imprimir op
        .Buttons(15).Image = 40 'Imprimir con lotes
        .Buttons(16).Image = 16 'Imprimir MOIXENT
        
        .Buttons(20).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    CargaCombos
    LimpiarCampos   'Limpia los campos TextBox
   
   
    Ordenacion = PonerTrabajadorConectado(NombreTabla)
    If Ordenacion <> "" Then
        TrabajadorConectado = CInt(Ordenacion)
    Else
        TrabajadorConectado = -1
    End If
        

    NombreTabla = "vallalmazaraproceso"
    Ordenacion = " ORDER BY id "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    
    CadenaConsulta = "Select * from " & NombreTabla & " where id= "
    If DatosADevolverBusqueda2 = "" Then
        CadenaConsulta = CadenaConsulta & "-1"
        
    Else
        CadenaConsulta = CadenaConsulta & DatosADevolverBusqueda2
    End If
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
    Else
        If Data1.Recordset.EOF Then
            PonerModo 1
            Text1(0).BackColor = vbYellow
        Else
            Me.Tag = "P" 'Para que en el activate ponga los campos
            PonerModo 2
        End If
    End If

    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""


    Me.cboAlmazara(0).ListIndex = -1
    Me.cboAlmazara(1).ListIndex = -1
    Me.cboAlmazara(2).ListIndex = -1
    Me.cboAlmazara(3).ListIndex = -1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub






Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        Else
            
            Me.txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            Me.txtAux(2).Text = RecuperaValor(CadenaDevuelta, 3)
            Me.txtAux(3).Text = RecuperaValor(CadenaDevuelta, 4)
            
            CadenaConsulta = "numalbar=" & DBSet(txtAux(0).Text, "T") & " AND fechaalb=" & DBSet(txtAux(1).Text, "F") & " AND "
            CadenaConsulta = CadenaConsulta & "codprove=" & DBSet(txtAux(2).Text, "T") & " AND numlinea"
            CadenaConsulta = DevuelveDesdeBD(conAri, "concat(codartic,'|',nomartic,'|',cantidad,'|')", "slialp", CadenaConsulta, "1")
            If Len(CadenaConsulta) <= 3 Then
                MsgBox "Error obteniendo datos albaran", vbExclamation
            Else
                Me.txtAux(4).Text = RecuperaValor(CadenaConsulta, 1)
                Me.txtAux(5).Text = RecuperaValor(CadenaConsulta, 2)
                Me.txtAux(6).Text = RecuperaValor(CadenaConsulta, 3)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub









Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub







Private Sub frmPe_DatoSeleccionado2(CadenaSeleccion As String)
    Text1(4).Text = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo <= 3 Or Modo > 4 Then Exit Sub
    CadenaConsulta = ""
    Screen.MousePointer = vbHourglass
    Set frmT = New frmAdmTrabajadores
    frmT.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
    frmT.Show vbModal
    Set frmT = Nothing
    
    If CadenaConsulta <> "" Then
        Indice = IIf(Index = 0, 4, 6)
        Text1(Index).Text = RecuperaValor(CadenaConsulta, 1)
        Indice = IIf(Index = 0, 5, 0)
        Text2(Indice).Text = RecuperaValor(CadenaConsulta, 2)
        CadenaConsulta = ""
    End If
    
    Screen.MousePointer = vbDefault
    
    
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   If Index = 2 Then Index = 4 'para que lo ponga bien
   Indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
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


Private Sub mnCambios_Click()
     BotonMtoLineas False
End Sub

Private Sub mnEliminar_Click()
    If Modo >= 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub










Private Sub mnLineas_Click()
    'Si esta cerrado no dejo pasar
    If Modo <> 2 Then Exit Sub
    If Not IsNull(Data1.Recordset!HoraFin) Then
        MsgBox "Proceso cerrado", vbExclamation
        Exit Sub
    End If
    BotonMtoLineas True
End Sub


Private Sub mnModificar_Click()
        
    If Modo >= 5 Then 'Modificar lineas
         BotonModificarLinea
   
        
    Else
        'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo >= 5 Then 'Añadir lineas
        If Modo = 5 Then
            BotonAnyadirLineaAlb
        Else
            BotonAnyadirLineaCambios
        End If
    Else 'Añadir Cabecera de Pedidos
         BotonAnyadir
    End If
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim Devuelve As String
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)
            
            
    
        Case 4, 6 '
            Devuelve = ""
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "CodTraba", Text1(Index).Text)
                    If Devuelve = "" Then MsgBox "No existe trabajador " & Text1(Index).Text, vbExclamation
                End If
            End If
            Text2(IIf(Index = 4, 5, 0)).Text = Devuelve
            
        Case 2, 5, 7, 8, 10
             
             
        
            If Not PonerFormatoDecimal(Text1(Index), 3) Then Text1(Index).Text = ""
            
 
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
   
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, Devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera Then
        Cad = Cad & ParaGrid(Text1(0), 20, "Id")
        Cad = Cad & ParaGrid(Text1(1), 20, "Fecha")
        Cad = Cad & "Tipo|scaalp|if(tipooliva=1,'Arbol','Tierra') |N||45·"
        Tabla = NombreTabla
      
        Titulo = "Proceso almazara"
        Devuelve = "0|"

    Else
        Titulo = "Albaranes proveedor"
        Cad = Cad & "Nº Albaran " & "|scaalp|numalbar|N||15·"
        Cad = Cad & "Fecha " & "|scaalp|fechaalb|T||15·"
        Cad = Cad & "C.Prov" & "|scaalp|codprove|N|0000|10·"
        Cad = Cad & "Proveedor" & "|scaalp|nomprove|N||40·"
        Cad = Cad & "Producto" & "|slialp|codartic|N||20·"
        Tabla = "scaalp,slialp,sartic,sfamia"
        Devuelve = "0|1|2|3|"
        
        
    'SELECT distinct scaalp.numalbar, scaalp.fechaalb, scaalp.codprove, scaalp.nomprove ,slialp.codartic FROM scaalp,slialp,sartic,sfamia WHERE

    
    
        cadB = "scaalp.numalbar= slialp.numalbar and scaalp.fechaalb  = slialp.fechaalb and scaalp.codprove = slialp.codprove and slialp.codartic=sartic.codartic and"
        cadB = cadB & " sfamia.codfamia=sartic.codfamia and tipfamia=30 and"
        cadB = cadB & " not (scaalp.numalbar,scaalp.fechaalb,scaalp.codprove) in "
        cadB = cadB & " (select numalbar,fechaalb,codprove from vallalmazaraprocesoalb ) "


        
      
        
        
        
        
        
        
        
    End If
    
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = Devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        If Not EsCabecera Then frmB.Label1.FontSize = 11

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
          
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    If Text1(4).Text <> "" Then Text2(5).Text = DevuelveDesdeBD(conAri, "nomtraba", "straba", "CodTraba", Text1(4).Text)
    If Text1(6).Text <> "" Then Text2(0).Text = DevuelveDesdeBD(conAri, "nomtraba", "straba", "CodTraba", Text1(6).Text)
       
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg_ As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg_ = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg_ = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg_
        
        

    'Campo ID, fechafin y deposito siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    For NumReg_ = 11 To 15
        BloquearTxt Text1(NumReg_), b
    Next
    
    b = Modo = 0 Or Modo = 2 Or Modo >= 5
    For NumReg_ = 1 To 10
        If NumReg_ < 5 Then BloquearCmb cboAlmazara(NumReg_ - 1), b
        BloquearTxt Text1(NumReg_), b
    Next
    
    Me.imgFecha(I).Enabled = Not b
    
    
    'Si no es modo lineas Boquear los TxtAux
'
'    For I = 0 To txtAux.Count - 1
'        BloquearTxt txtAux(I), (Modo <> 5)
'    Next I
'
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    
    cmdCancelar.visible = b
    cmdAceptar.visible = b


    imgBuscar(0).visible = b


    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    

    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim b As Boolean
Dim Devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    

    'Si pone dosis de talco
    Devuelve = DevuelveDesdeBD(conAri, "artTalco", "vallparam", "1", "1")
    CadenaConsulta = ""
    If Devuelve = "" Then
        'NO esta configurado el talco
        If Text1(5).Text <> "" Or Text1(9).Text <> "" Then CadenaConsulta = "Talco sin configurar. No indique valores"
    Else
        If Text1(5).Text = "" Xor Text1(9).Text = "" Then
            CadenaConsulta = "Error en talco. Indique todos o ninguno de los dos valores"
        Else
            'Si ha puesto los valores
            If Text1(9).Text <> "" Then
                CadenaConsulta = "codartic =" & DBSet(Devuelve, "T") & " AND numlote"
                CadenaConsulta = DevuelveDesdeBD(conAri, "id", "spartidas", CadenaConsulta, Text1(9).Text, "T")
                If CadenaConsulta = "" Then
                   CadenaConsulta = "No existe el lote del talco: " & Text1(9).Text
                Else
                    'OK. PERFECTO TODO
                    CadenaConsulta = ""
                End If

            End If
        End If
    End If
    If CadenaConsulta <> "" Then
        MsgBox CadenaConsulta, vbExclamation
        CadenaConsulta = ""
        Exit Function
    End If
    
    
    b = True
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim Cad As String
Dim I As Integer
    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    If Modo = 5 Then
            'Albaranes
            For I = 0 To txtAux.Count - 1
                If txtAux(I).Text = "" And I <> 3 Then
                    MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                    b = False
                    PonerFoco txtAux(I)
                    Exit Function
                End If
            Next I
             
            'Ahora veremos dos cosas:
            '1.- Existe el albaran
            Cad = "numalbar=" & DBSet(txtAux(0).Text, "T") & " AND fechaalb =" & DBSet(txtAux(1).Text, "F")
            Cad = Cad & " AND codprove "
            
            Cad = DevuelveDesdeBD(conAri, "numalbar", "scaalp", Cad, txtAux(2).Text)
            If Cad = "" Then
                'No existe el albarán
                MsgBox " No existe el albaran", vbExclamation
                b = False
            End If
             
            If Not PonerFormatoDecimal(txtAux(6), 3) Then b = False
        
             
             
            'Vemos que no esta asignado en ningun otro parte molturacion
            Cad = "numalbar=" & DBSet(txtAux(0).Text, "T") & " AND fechaalb =" & DBSet(txtAux(1).Text, "F")
            If ModificaLineas = 2 Then Cad = Cad & " AND id<>" & Data1.Recordset!ID
            Cad = Cad & " AND codprove "
            Cad = DevuelveDesdeBD(conAri, "id", "vallalmazaraprocesoalb", Cad, txtAux(2).Text)
            If Cad <> "" Then
                'No existe el albarán
                MsgBox " Ya esta asignado el albaran. Codigo: " & Cad, vbExclamation
                b = False
            End If
             
             
            'Hay que ver si el albaran tiene lineas de OLIVA
            'Lineas de oliva, la familia tiene que ser OLIVA
            'Es decir, el tipfamia=30
            If b Then
                Cad = "slialp.codartic=sartic.codartic and sfamia.codfamia=sartic.codfamia and tipfamia=30 and numalbar=" & DBSet(txtAux(0).Text, "T") & " AND fechaalb =" & DBSet(txtAux(1).Text, "F")
                Cad = Cad & " AND slialp.codprove "
                
                Cad = DevuelveDesdeBD(conAri, "numalbar", "slialp,sartic,sfamia", Cad, txtAux(2).Text)
                If Cad = "" Then
                    MsgBox "Ningun articulo en el albaran es oliva", vbExclamation
                    b = False
                End If
            End If
    Else
        'Observaciones
        
        b = DatosOkLineaCompo
    End If
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLineaCompo() As Boolean
    DatosOkLineaCompo = False
    
   If txtaux2(0).Text = "" Or txtaux2(1).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Function
    End If
    
    If Not EsHoraOK(txtaux2(0).Text) Then
        MsgBox "Hora incorrecta", vbExclamation
        Exit Function
    End If
   
   
    
    DatosOkLineaCompo = True
End Function






Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim b As Boolean
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
            
        Case 10  'Lineas
            mnLineas_Click
            
        Case 11
            mnCambios_Click
            
        Case 12, 14, 15
            'cerrar(12) orden produccion
            '--------------------------------------------------------------------
            
            If Modo <> 2 Then Exit Sub
            
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una orden de almazara", vbExclamation
                Exit Sub
            End If
                  
            If Not IsNull(Data1.Recordset!HoraFin) Then
                MsgBox "Proceso cerrado. ", vbExclamation
                Exit Sub
            End If
                  
                  
            If Data2.Recordset.EOF Then
                MsgBox "Ningun albaran asignado", vbExclamation
                Exit Sub
            End If
            
            
            
            
            If Button.Index = 12 Then
                conn.Execute "DELETE FROM tmpnlotes WHERE codusu =" & vUsu.Codigo
                Espera 0.5
                CadenaDesdeOtroForm = "INSERT INTO tmpnlotes(codusu,numalbar,fechaalb,codprove,numlinea,codartic,nomartic,cantidad,numlotes) "
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " SELECT " & vUsu.Codigo
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",numalbar,fechaalb,codprove,numlinea,codartic,nomartic,"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " 0,cantidad numlote from slialp where numlinea=1"  'tab podriamos hacer linkar con sfamia y tipfamia=30
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND (numalbar,fechaalb,codprove) IN (select numalbar,fechaalb,codprove"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " FROM  vallalmazaraprocesoalb WHERE id =" & Data1.Recordset!ID & ")"
                conn.Execute CadenaDesdeOtroForm
                
                frmVallCierrAlmazara.ID = CLng(Data1.Recordset!ID)
                frmVallCierrAlmazara.Show vbModal
                
                If CadenaDesdeOtroForm <> "" Then
                    PosicionarData
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case 20    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()

End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String
Dim vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""


    If Not DatosOkLinea() Then Exit Function
    
    If Modo = 5 Then
        SQL = "INSERT INTO vallalmazaraprocesoalb(Id,numalbar,fechaalb,codprove,codartic,nomartic,kilos) VALUES (" & Data1.Recordset!ID & ","
        SQL = SQL & DBSet(txtAux(0).Text, "T") & "," & DBSet(txtAux(1).Text, "F") & "," & DBSet(txtAux(2).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(4).Text, "T") & "," & DBSet(txtAux(5).Text, "T") & "," & DBSet(txtAux(6).Text, "N") & ")"
    Else
        
        SQL = SugerirCodigoSiguienteStr("vallalmazaraprocesocambios", "secuencial", "id = " & Data1.Recordset!ID)
        SQL = "INSERT INTO vallalmazaraprocesocambios(Id,secuencial,FechaFin,observa) VALUES (" & Data1.Recordset!ID & "," & SQL & ",'" & Format(Data1.Recordset!Fecha, FormatoFecha)
        SQL = SQL & " " & txtaux2(0).Text & "'," & DBSet(txtaux2(1).Text, "T") & ")"
    End If
    If SQL <> "" Then
        conn.Execute SQL
        

        
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Produccion" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If Not DatosOkLinea() Then Exit Function
        'Creamos la sentencia SQL
    If Modo = 5 Then
        SQL = "UPDATE vallalmazaraprocesoalb SET numalbar = " & DBSet(txtAux(0).Text, "T") & ", fechaalb = " & DBSet(txtAux(1).Text, "F")
        SQL = SQL & " , codprove = " & DBSet(txtAux(2).Text, "N")
        SQL = SQL & " , codartic = " & DBSet(txtAux(4).Text, "T")
        SQL = SQL & " , nomartic = " & DBSet(txtAux(5).Text, "T")
        SQL = SQL & " , kilos = " & DBSet(txtAux(6).Text, "N")
        SQL = SQL & " WHERE numalbar = " & DBSet(Data2.Recordset!NumAlbar, "T") & " AND fechaalb = " & DBSet(Data2.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove = " & Data2.Recordset!codProve
    Else
        
        SQL = "UPDATE vallalmazaraprocesocambios SET FechaFin = '" & Format(Data1.Recordset!Fecha, FormatoFecha)
        SQL = SQL & " " & txtaux2(0).Text & "', observa = " & DBSet(txtaux2(1).Text, "T")
        SQL = SQL & " WHERE secuencial = " & data3.Recordset!secuencial
    End If
    SQL = SQL & " AND id = " & Data1.Recordset!ID
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        
        
        ModificarLinea = True
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    CargaGrid3 enlaza
    
    
    
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid3(enlaza As Boolean)
Dim SQL As String

    SQL = "id = -1"


    If enlaza Then

            SQL = " id = " & Data1.Recordset!ID
            

    End If

    SQL = "select FechaFin,observa,secuencial  from vallalmazaraprocesocambios where  " & SQL
    data3.ConnectionString = conn
    data3.RecordSource = SQL
    data3.Refresh
    DataGrid2.AllowRowSizing = False
    If DataGrid2.DataSource Is Nothing Then DataGrid2.ClearFields
        
    Set DataGrid2.DataSource = data3
    DataGrid2.RowHeight = 690
    
    DataGrid2.Columns(0).Caption = "Hora"
    DataGrid2.Columns(0).Width = 1300
    DataGrid2.Columns(0).NumberFormat = "hh:mm:ss"
    
    DataGrid2.Columns(1).Caption = "Observaciones"
    DataGrid2.Columns(1).Width = 3200
    DataGrid2.Columns(2).Width = 0
    
    
End Sub



Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte

    On Error GoTo ECargaGrid

    vData.Refresh

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(0).Caption = "Albarán"
                vDataGrid.Columns(0).Width = 900
                vDataGrid.Columns(2).NumberFormat = "000"
                
                vDataGrid.Columns(1).Caption = "Fecha"
                vDataGrid.Columns(1).Width = 1100

                vDataGrid.Columns(2).Caption = "Cod."
                vDataGrid.Columns(2).Width = 800
                vDataGrid.Columns(2).NumberFormat = "000000"
                vDataGrid.Columns(3).Caption = "Proveedor"
                vDataGrid.Columns(3).Width = 3050

             
                vDataGrid.Columns(4).Caption = "Refer."
                vDataGrid.Columns(4).Width = 1150
                vDataGrid.Columns(5).Caption = "Articulo"
                vDataGrid.Columns(5).Width = 2450
                vDataGrid.Columns(6).Caption = "Kilos"
                vDataGrid.Columns(6).Width = 1100
                vDataGrid.Columns(6).NumberFormat = FormatoCantidad
                vDataGrid.Columns(6).Alignment = dbgRight
    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible

    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), IIf(I > 3, True, False)
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I).Text
                BloquearTxt txtAux(I), IIf(I > 3, True, False)
            Next I
        End If
               

    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
      '  cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
      '  cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Alb
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(0).Width - 160
        
       
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Fecha
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(1).Width - 30
         'txtAux(1).BackColor = vbRed 'para ver si encja bien
        'Proveed
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 20
        txtAux(2).Width = DataGrid1.Columns(2).Width - 10
        'txtAux(2).BackColor = vbRed 'para ver si encja bien
        'Nomprove
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(3).Width - 10
        For I = 3 To 6
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I).Width - 10
            
        Next
      
      
      
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible

    End If
    BloquearTxt txtAux(3), True
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtaux2(I).Top = 290
            txtaux2(I).visible = visible
        Next I
       

    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            For I = 0 To txtaux2.Count - 1
                txtaux2(I).Text = ""
                BloquearTxt txtaux2(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtaux2.Count - 1
                txtaux2(I).Text = DataGrid2.Columns(I).Text
                txtaux2(I).Locked = False
            Next I
        End If
               

    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 10)
        
        For I = 0 To txtaux2.Count - 1
            txtaux2(I).Top = alto
            If I > 0 Then txtaux2(I).Height = DataGrid2.RowHeight
        Next I

        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Alb
        txtaux2(0).Left = DataGrid2.Left + 330
        txtaux2(0).Width = DataGrid2.Columns(0).Width - 10

      
        txtaux2(1).Left = txtaux2(0).Left + txtaux2(0).Width + 10
        txtaux2(1).Width = DataGrid2.Columns(1).Width - 10
 
      
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtaux2.Count - 1
            txtaux2(I).visible = visible
        Next I


    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub






Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub






Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    
        KEYpress KeyAscii
    
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Devuelve As String


    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
           
            'If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'fECHA
            Devuelve = txtAux(Index).Text
            If Not EsFechaOK(Devuelve) Then
                Devuelve = ""
                PonerFoco txtAux(Index)
            End If
            txtAux(Index).Text = Devuelve
            
        Case 2 'cODPROVE
            Devuelve = ""
             If PonerFormatoEntero(txtAux(Index)) Then
                Devuelve = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text)
                If Devuelve = "" Then
                    MsgBox "No existe el proveedor " & txtAux(Index).Text, vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
            txtAux(3).Text = Devuelve
      
    End Select
    
    If Index < 3 Then
        If txtAux(0).Text <> "" And txtAux(1).Text <> "" And txtAux(2).Text <> "" Then
            CadenaConsulta = "numalbar=" & DBSet(txtAux(0).Text, "T") & " AND fechaalb=" & DBSet(txtAux(1).Text, "F") & " AND "
            CadenaConsulta = CadenaConsulta & "codprove=" & DBSet(txtAux(2).Text, "T") & " AND numlinea"
            Devuelve = DevuelveDesdeBD(conAri, "concat(codartic,'|',nomartic,'|',cantidad,'|')", "slialp", CadenaConsulta, "1")
            If Len(Devuelve) <= 3 Then
                MsgBox "Error obteniendo datos albaran", vbExclamation
                Me.txtAux(4).Text = ""
                Me.txtAux(5).Text = ""
                Me.txtAux(6).Text = ""
            Else
                Me.txtAux(4).Text = RecuperaValor(Devuelve, 1)
                Me.txtAux(5).Text = RecuperaValor(Devuelve, 2)
                Me.txtAux(6).Text = RecuperaValor(Devuelve, 3)
            End If
        End If
    End If
End Sub


Private Sub BotonMtoLineas(Albaran As Boolean)
       If Not IsNull(Data1.Recordset!HoraFin) Then
            MsgBox "Proceso  cerrado", vbExclamation
        Else
            ModificaLineas = 0
            EsCabecera = False
            If Albaran Then
                TituloLinea = "alb."
                PonerModo 5
            Else
                TituloLinea = "cambios"
                PonerModo 6
            End If
            PonerBotonCabecera True
        End If
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean



    On Error GoTo FinEliminar

        conn.BeginTrans
        'Los lotes
        conn.Execute "DELETE FROM vallalmazaraprocesocambios where id =" & Text1(0).Text
        conn.Execute "Delete from vallalmazaraprocesoalb where id =" & Text1(0).Text
        conn.Execute "Delete from vallalmazaraproceso where id=" & Text1(0).Text
        b = True
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido" & vbCrLf, Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = Replace(ObtenerWhereCP, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String

    On Error Resume Next
    
    SQL = NombreTabla & ".id= " & Val(Text1(0).Text)
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "select numalbar,fechaalb,vallalmazaraprocesoalb.codprove ,nomprove,codartic,nomartic,kilos from vallalmazaraprocesoalb ,sprove "
    SQL = SQL & " where vallalmazaraprocesoalb.codprove=sprove.codprove and id = "
    If enlaza Then
        SQL = SQL & Data1.Recordset!ID
    Else
        SQL = SQL & " -1"
    End If
    SQL = SQL & " Order by id,numalbar"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo >= 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = (b Or Modo = 0)
        Me.mnModificar.Enabled = (b Or Modo = 0)
        'eliminar
        Toolbar1.Buttons(7).Enabled = (b Or Modo = 0)
        Me.mnEliminar.Enabled = (b Or Modo = 0)
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        Toolbar1.Buttons(11).Enabled = b
        Me.mnBuscar.Enabled = b
        Toolbar1.Buttons(12).Enabled = b
        
        

        
        
        
        
        
      
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub







    

Private Function PedidoConInstalaciones() As Boolean
'Comprobar si en las lineas del Pedido hay algun articulo que sea Instalacion
'Si no hay niguna linea que sea instalacion no se imprimira la Orden de Instalacion
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EInstalac

    PedidoConInstalaciones = False
    SQL = "SELECT sliped.codartic, sliped.numlinea,scaped.numpedcl, sfamia.instalac "
    SQL = SQL & " FROM ((sliped INNER JOIN scaped ON sliped.numpedcl=scaped.numpedcl) "
    SQL = SQL & " INNER JOIN sartic ON sliped.codartic=sartic.codartic) INNER JOIN "
    SQL = SQL & " sfamia ON sartic.codfamia=sfamia.codfamia "
    SQL = SQL & " WHERE scaped.numpedcl = " & Val(Text1(0).Text) & " And sfamia.instalac = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        PedidoConInstalaciones = False
    Else
        PedidoConInstalaciones = True
    End If
    RS.Close
    Set RS = Nothing
    
EInstalac:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar si hay Articulos que son Instalaciones.", Err.Description
End Function






Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String

    On Error GoTo EEliminarPed

     SQL = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
   ' conn.Execute "Delete from " & NomTablaLineas & sql

    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function








Private Sub InsertarCabecera()

    CadenaConsulta = ""
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "id", CadenaConsulta)

    
    If InsertarDesdeForm(Me) Then
    
           
    
            'Si tiene pedido traeremos las lineas del pedido
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE id = " & Text1(0).Text & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo In sertar Lineas
            BotonMtoLineas True
            BotonAnyadirLineaAlb
    
    Else
        CadenaConsulta = ""
    End If

End Sub






Private Sub PosicionaData2()
    On Error GoTo EPos
    Data2.Recordset.Find "Codartic= " & DBSet(txtAux(1).Text, "T")
    
    Exit Sub
EPos:
    MuestraError Err.Number, "Posicionando data2"
End Sub


                       

Private Sub CargaCombos()
    'Tipo de oliva
    cboAlmazara(0).Clear
    cboAlmazara(0).AddItem "Arbol"
    cboAlmazara(0).ItemData(cboAlmazara(0).NewIndex) = 0
    cboAlmazara(0).AddItem "Tierra"
    cboAlmazara(0).ItemData(cboAlmazara(0).NewIndex) = 1
    
    'Aspecto salida del decanter 0. Sindefinir 1Sucio 2Normal 3Limpio',
    cboAlmazara(1).Clear
    cboAlmazara(1).AddItem "" 'Sin definir
    cboAlmazara(1).ItemData(cboAlmazara(1).NewIndex) = 0
    cboAlmazara(1).AddItem "Sucio"
    cboAlmazara(1).ItemData(cboAlmazara(1).NewIndex) = 1
    cboAlmazara(1).AddItem "Normal"
    cboAlmazara(1).ItemData(cboAlmazara(1).NewIndex) = 2
    cboAlmazara(1).AddItem "Limpio"
    cboAlmazara(1).ItemData(cboAlmazara(1).NewIndex) = 3
    
    
    'Aspecto masa batidora 0Sindefinir 1Facil  2Dificil',
    cboAlmazara(2).Clear
    cboAlmazara(2).AddItem "" 'Sin definir
    cboAlmazara(2).ItemData(cboAlmazara(2).NewIndex) = 0
    cboAlmazara(2).AddItem "Fácil"
    cboAlmazara(2).ItemData(cboAlmazara(2).NewIndex) = 1
    cboAlmazara(2).AddItem "Difícil"
    cboAlmazara(2).ItemData(cboAlmazara(2).NewIndex) = 2
    
    
    ''Cant. que sobre nada en batidora 0Sindef 1Mucho 2Normal 3Poco 4Muypoco',
    cboAlmazara(3).Clear
    cboAlmazara(3).AddItem "" 'Sin definir
    cboAlmazara(3).ItemData(cboAlmazara(3).NewIndex) = 0
    cboAlmazara(3).AddItem "Mucho"
    cboAlmazara(3).ItemData(cboAlmazara(3).NewIndex) = 1
    cboAlmazara(3).AddItem "Normal"
    cboAlmazara(3).ItemData(cboAlmazara(3).NewIndex) = 2
    cboAlmazara(3).AddItem "Poco"
    cboAlmazara(3).ItemData(cboAlmazara(3).NewIndex) = 3
    cboAlmazara(3).AddItem "Muy poco"
    cboAlmazara(3).ItemData(cboAlmazara(3).NewIndex) = 4
    
    
    
End Sub

Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFoco txtaux2(Index), 3
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then KEYpress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
    If Not PerderFocoGnralLineas(txtaux2(Index), ModificaLineas) Then Exit Sub
    
    If txtaux2(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        PonerFormatoHora txtaux2(Index)
    End If
End Sub
