VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComFacturar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Compra Proveedores"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12000
   Icon            =   "frmComFacturar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComFacturar.frx":000C
   ScaleHeight     =   6975
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFactura 
      Height          =   4860
      Left            =   6840
      TabIndex        =   18
      Top             =   2000
      Width           =   5055
      Begin VB.CheckBox chkTipoRet 
         Caption         =   "Base + IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdIVA 
         Caption         =   "+"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   55
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   51
         Tag             =   "Impret|N|S|||scafac|impret|#,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   50
         Tag             =   "PorRet|N|S|0||scafac|PorRet|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   9
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   42
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1350
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   900
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   600
         MaxLength       =   5
         TabIndex        =   36
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3||N|"
         Text            =   "Text1 7"
         Top             =   2805
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   600
         MaxLength       =   5
         TabIndex        =   35
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2||N|"
         Text            =   "Text1 7"
         Top             =   2475
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   600
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1||N|"
         Text            =   "Text1 7"
         Top             =   2160
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   28
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2160
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1185
         MaxLength       =   5
         TabIndex        =   27
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2160
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   19
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2160
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2475
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1185
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2475
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   20
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2475
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2805
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   21
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2805
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   21
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2805
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   22
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4320
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Importe retencion"
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   53
         Top             =   3495
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "% Retencion"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   52
         Top             =   3495
         Width           =   1095
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4920
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   47
         Top             =   900
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   4920
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. gral."
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   46
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   45
         Top             =   570
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. ppago"
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   44
         Top             =   570
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4920
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   43
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   37
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   33
         Left            =   3480
         TabIndex        =   32
         Top             =   1950
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   31
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   3450
         TabIndex        =   30
         Top             =   4080
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   41
         Left            =   1185
         TabIndex        =   29
         Top             =   1950
         Width           =   495
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   1550
      Left            =   120
      TabIndex        =   8
      Top             =   385
      Width           =   11775
      Begin VB.CheckBox Check1 
         Caption         =   "Contabiliz."
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   48
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1240
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   1400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1000
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3830
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepci�n|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   400
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1000
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   400
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6915
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1000
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6915
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Operador|N|N|0|9999|scafpc|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   400
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   550
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   1000
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2005
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   400
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "N� Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   400
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4895
         Picture         =   "frmComFacturar.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3070
         Picture         =   "frmComFacturar.frx":0A99
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6600
         Picture         =   "frmComFacturar.frx":0B24
         ToolTipText     =   "Buscar banco propio"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmComFacturar.frx":0C26
         ToolTipText     =   "Buscar proveedor"
         Top             =   1030
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Recep."
         Height          =   255
         Index           =   3
         Left            =   3830
         TabIndex        =   16
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Prev. Pago"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   13
         Top             =   795
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   6600
         Picture         =   "frmComFacturar.frx":0D28
         ToolTipText     =   "Buscar trabajador"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Operador"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   12
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
         Height          =   255
         Index           =   29
         Left            =   2005
         TabIndex        =   10
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   9
         Top             =   200
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5640
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
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir Datos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Albaranes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame FrameList 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      TabIndex        =   58
      Top             =   2040
      Width           =   6735
      Begin MSComctlLib.ListView ListView1 
         Height          =   4770
         Left            =   0
         TabIndex        =   59
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8414
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmComFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmBanPr As frmFacBancosPropios 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1



Private Modo2 As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWhere As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------

Dim dtoGn As Currency
Dim dtoPP As Currency
Dim ForPa As Integer


Private vProve As CProveedor


                    
Private Sub chkTipoRet_Click()
    Text1_LostFocus 23  'Como si cambaira la retencion
End Sub

Private Sub cmdGenerar_Click(Index As Integer)
Dim b As Boolean
    If Index = 1 Then
        'QUITO EL PORCENTAJE
        Text1(23).Text = ""
        Text1(24).Text = ""
        CalcularDatosFactura
        'Le ha dado a cancelar
        PonerModo2 4
        
    Else
        'Aceptar
        If Text1(23).Text = "" Xor Text1(24).Text = "" Then
            MsgBox "Si pone porcentaje retencion debe poner importe(y viceversa)", vbExclamation
            Exit Sub
        End If
        
        
        'Abril 2009
        'Puede facturar en B
        If vUsu.TrabajadorB Then
            b = AbrirConexionConta(True)
        Else
            b = True
        End If
        
        If b Then GenerarFactura_
        
        'Reestablecemos la conexion con la conta normal
        If vUsu.TrabajadorB Then AbrirConexionConta (False)
        
    End If
End Sub

Private Sub cmdIVA_Click()
Dim Impor As Currency
    'Poner nuevo tipo de IVA
    Set frmB = New frmBuscaGrid
    CadenaDesdeOtroForm = ""
    frmB.vCampos = "C�digo|tiposiva|codigiva|N||20�Denominacion|tiposiva|nombriva|T||60�Porcentaje|tiposiva|porceiva|N||10�"
    frmB.vTabla = "tiposiva"
    frmB.vTitulo = "Tipos de IVA"
    frmB.vDevuelve = "0|2|"
    
    frmB.vselElem = 1
    frmB.vConexionGrid = conConta
    frmB.vCargaFrame = False
    frmB.Show vbModal
    Set frmB = Nothing
    If CadenaDesdeOtroForm <> "" Then
        
        Text1(10).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Impor = CCur(RecuperaValor(CadenaDesdeOtroForm, 2))
        Text1(13).Text = CStr(Impor)  '% iva
        NumRegElim = Impor * 100 'Para no decalrar mas variabnles
        Impor = ImporteFormateado(Text1(16).Text)
        Impor = Round2((Impor * NumRegElim / 10000), 2)
        Text1(19).Text = CStr(Impor)
        
        PonerFormatoEntero Text1(10)
        PonerFormatoDecimal Text1(13), 3
        PonerFormatoDecimal Text1(19), 3
        RecalculoDeImportes
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If VerAlbaranes Then RefrescarAlbaranes
    VerAlbaranes = False
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 18   'Pedir Datos
        .Buttons(2).Image = 43   'Ver albaranes
        .Buttons(3).Image = 26   'Generar FActura
        .Buttons(6).Image = 15   'Salir
    End With
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
   
    '## A mano
    NombreTabla = "scafpc" 'cabecera facturas compras a proveedor
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo2 0
'    Else
'        PonerModo 1
    End If
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "RECFAC"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDesdeOtroForm = CadenaDevuelta
End Sub

Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Proveedores
Dim Indice As Byte
    
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Proveedor
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom proveedor
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo2 = 2 Or Modo2 = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            Indice = 3
            
        Case 1 'Operador. Trabajador
            Indice = 4
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
       
       Case 2 'Bancos Propios
            Indice = 5
            Set frmBanPr = New frmFacBancosPropios
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
    End Select
    
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo2 = 2 Or Modo2 = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index + 1
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Cuando se selecciona un albaran de la lista
Dim I As Integer
Dim cad As String
Dim TipoFP As Integer 'Forma de pago
Dim TipoDtoPP As Currency 'descuento pronto pago
Dim tipoDtoGn As Currency 'descuento general

    Screen.MousePointer = vbHourglass
    
    Set ListView1.SelectedItem = Item
    
    'Inicializamos a cero
    TipoFP = 0
    TipoDtoPP = 0
    tipoDtoGn = 0
    
    'cuando seleccionamos un check vemos si lo podemos seleccionar
    'ya que si ya habia algun albaran selecionado tendremos que comprobar
    'que son de la misma forpa, dtoppago y dtognral.
    'si esto no se cumple no se pueden agrupar en la misma factura
    For I = 1 To ListView1.ListItems.Count
        If Item.Index <> I Then
            If ListView1.ListItems(I).Checked Then
                'ya habia otro albaran seleccionado
                TipoFP = ListView1.ListItems(I).SubItems(2)
                TipoDtoPP = CCur(ListView1.ListItems(I).SubItems(4))
                tipoDtoGn = CCur(ListView1.ListItems(I).SubItems(5))
                Exit For
            End If
        End If
    Next I
    
    If Not (TipoFP = 0 And TipoDtoPP = 0 And tipoDtoGn = 0) Then
    'si ya habia un albaran seleccionado, comprobar que es del mismo tipo
        If Item.SubItems(2) <> TipoFP Or Item.SubItems(4) <> TipoDtoPP Or Item.SubItems(5) <> tipoDtoGn Then
            MsgBox "Se debe seleccionar albaranes de la misma Forma de Pago y Descuentos", vbExclamation
            ListView1.SelectedItem.Checked = False
            Screen.MousePointer = vbDefault
            ListView1.SetFocus
            Exit Sub
        End If
    Else
        'Todo bien
        
    End If
    
    
    'MORALES. Comprobar lotaje
    If Not ComprobarNumerosLotes Then
            ListView1.SelectedItem.Checked = False
            Screen.MousePointer = vbDefault
            ListView1.SetFocus
            Exit Sub

    End If
    
    ' Calculamos los datos de factura
    If Not VerAlbaranes Then CalcularDatosFactura
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnGenerarFac_Click()
    BotonFacturar
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnVerAlbaran_Click()
    BotonVerAlbaranes
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo2
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
Dim Impor As Currency
Dim C As String

    If Modo2 <> 5 Then _
        If Not PerderFocoGnral(Text1(Index), Modo2) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                ' No debe existir el n�mero de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                End If
            End If
            
        Case 3 'Cod Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerDatosProveedor
                ' No debe existir el n�mero de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
                    DesBloqueoManual ("RECFAC")
                    If Not BloqueoManual("RECFAC", Text1(3).Text) Then
                        MsgBox "No se puede recepcionar factura de ese proveedor. Hay otro usuario recepcionando.", vbExclamation
                        BotonPedirDatos
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        CargarAlbaranes
                    End If
                    
                End If
                
            Else
                Text2(Index).Text = ""
            End If

        Case 4 'Cod Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index).Text = ""
            End If

        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
        Case 23, 24
            'SON EL IMPORTE DE RETENCION y el porcentaje
            Text1(Index).Text = Text1(Index).Text
            
            PonerFormatoDecimal Text1(Index), 3
                            
            If Index = 23 Then
                If Text1(23).Text <> "" Then
                    'Ha puesto porcentaje retencion
                    Impor = 0
                    For NumRegElim = 0 To 2
                        'Base imponible
                        If Text1(16 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(16 + NumRegElim).Text)
                        
                        'Si solo es sobre la BASE, esto no lo sumo
                        If Me.chkTipoRet.Value Then
                            If Text1(19 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(19 + NumRegElim).Text)
                        End If
                    Next NumRegElim
                    NumRegElim = ImporteFormateado(Text1(23).Text) * 100
                    Impor = Round2((Impor * NumRegElim / 10000), 2)
                    Text1(24).Text = Format(Impor, FormatoImporte)
                    RecalculoDeImportes
                End If
                
            End If
            
            
    End Select
End Sub


'RECALCULO DATOS FACTURA
'-----------------------------------------------------
Private Sub RecalculoDeImportes()
Dim Impor As Currency
        Impor = 0
        For NumRegElim = 0 To 2
            'Base imponible + iva
            If Text1(16 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(16 + NumRegElim).Text)
            If Text1(19 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(19 + NumRegElim).Text)
        Next NumRegElim
        'Memos la retencion
        If Text1(24).Text <> "" Then Impor = Impor - ImporteFormateado(Text1(24).Text)
        
        'TOTAL FACTURA
        Text1(22).Text = CStr(Impor)
        PonerFormatoDecimal Text1(22), 3
End Sub




'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'

'       MODO: 0 pidiendo datos encabezado


Private Sub PonerModo2(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo


    Modo2 = Kmodo
    
    
    'GEneral
    b = (Modo2 = 5)
    FrameFactura.Enabled = b   'Solo habilitado al final
    Toolbar1.Enabled = Not b
    'Antes. Para que no se quede en gris
    'ListView1.Enabled = Not b
    'para que no se quede el listview en gris
    FrameList.Enabled = Not b
    FrameIntro.Enabled = Not b
    
    cmdGenerar(0).visible = b
    cmdGenerar(1).visible = b
    'chkTipoRet.visible = b
    If Not b Then cmdIVA.visible = False
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo2 = 2)
        
                 
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo2
    
    'Importes siempre bloqueados
    For I = 6 To 22
        BloquearTxt Text1(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF    'Total factura
        
    
    
    If Modo2 = 4 Then
        For I = 0 To 4
            If I <> 2 Then
                Text1(I).Locked = False
                Text1(I).BackColor = vbWhite
            End If
        Next
    End If
        
    If Modo2 = 5 Then
        BloquearTxt Text1(23), False
        BloquearTxt Text1(24), False
        'Si el tipo de proveedor NO es REA
        'y solo tiene un tipo de IVA , podemos dejar que cambie el iva
        If Text1(11).Text = "" And Text1(12).Text = "" Then
            If vProve.TipoProv <> 2 Then cmdIVA.visible = True
        End If
    End If
    
    
    
    '---------------------------------------------
    b = (Modo2 <> 0 And Modo2 <> 2 And Modo2 <> 5)
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
                    
    Me.chkVistaPrevia.Enabled = (Modo2 <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo2, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim cad As String
Dim I As Byte

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For I = 0 To 5
        If Text1(I).Text = "" Then
            If Text1(I).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(I)) Then
                    cad = vtag.Nombre
                Else
                    cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                cad = "Campo"
                If I = 5 Then cad = "Cta. Prev. Pago"
            End If
            MsgBox cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerFoco Text1(I)
            Exit Function
        End If
    Next I
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepci�n debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    If Not EsFechaOKConta(CDate(Text1(2).Text)) Then Exit Function
    
    'comprobar que se han seleccionado lineas para facturar
    If cadWhere = "" Then
        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
        Exit Function
    End If
    
    
    ' No debe existir el n�mero de factura para el proveedor en hco
    If ExisteFacturaEnHco Then Exit Function
    
    
    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
    cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
    cad = cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    If RegistrosAListar(cad) > 1 Then
        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
        Exit Function
    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
    cad = "select distinct (codforpa) from scaalp "
    cad = cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = miRsAux.Fields(0)
    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    cad = "Select tipforpa from sforpa where codforpa=" & cad
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        I = 1
        cad = miRsAux.Fields(0)
        If Val(cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            I = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If I = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vProve.CuentaBan = "" Or vProve.DigControl = "" Or vProve.Sucursal = "" Or vProve.Banco = "" Then
            cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    �Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then I = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If I > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
             
        Case 2 'Ver Albaranes
            mnVerAlbaran_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

        Case 6    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo2, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String


    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
    InicializarListView
    
    'Como no habr� albaranes seleccionados vaciamos la cadwhere
    cadWhere = ""
    
    PonerModo2 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
    'poner trabajador conectado como operador
    Text1(4).Text = PonerTrabajadorConectado(Nombre)
    Text2(4).Text = Nombre
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub

Private Sub BotonVerAlbaranes()

    If Not SeleccionaRegistros Then Exit Sub
    
    VerAlbaranes = True
    
    frmComEntAlbaranes.cadSelAlbaranes = cadWhere
    frmComEntAlbaranes.Show vbModal
    frmComEntAlbaranes.cadSelAlbaranes = ""
End Sub
    


Private Sub CargarAlbaranes()
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
On Error GoTo ECargar

    ListView1.ListItems.Clear
    If VerAlbaranes = False Then cadWhere = ""
    
    'si no hay proveedor salir
    If Text1(3).Text = "" Then Exit Sub
    
    SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
    SQL = SQL & " sum(slialp.importel) as bruto "
    SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
    SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
    SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text
    SQL = SQL & " AND presupuesto = "
    'Abril 2009   El "B"
    If vUsu.TrabajadorB Then
        SQL = SQL & "1"
    Else
        SQL = SQL & "0"
    End If
    SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
    SQL = SQL & " ORDER BY scaalp.numalbar"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    InicializarListView
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = RS!NumAlbar
        ItmX.SubItems(1) = Format(RS!FechaAlb, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(RS!codforpa, "000")
        ItmX.SubItems(3) = RS!nomforpa
        ItmX.SubItems(4) = Format(RS!DtoPPago, "#0.00")
        ItmX.SubItems(5) = Format(RS!DtoGnral, "#0.00")
        ItmX.SubItems(6) = Format(RS!bruto, "#,###,#0.00") '(RAFA/ALZIRA) 12092006
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "N�Albaran", 1100
    ListView1.ColumnHeaders.Add , , "Fecha", 1100, 2
    ListView1.ColumnHeaders.Add , , "FPag", 550
    ListView1.ColumnHeaders.Add , , "Desc. FPago", 1450
    ListView1.ColumnHeaders.Add , , "DtoPP", 650, 2
    ListView1.ColumnHeaders.Add , , "DtoGr", 600, 2
    ListView1.ColumnHeaders.Add , , "Imp. Bruto", 1100, 1
End Sub



Private Sub CalcularDatosFactura()
Dim I As Integer
Dim SQL As String
Dim cadAux As String
Dim vFactu As CFacturaCom
Dim exento As Boolean
Dim CambioIVA As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 6 To 22
         Text1(I).Text = ""
    Next I

    cadAux = ""
    cadWhere = ""
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
        'para cada albaran seleccionado para la factura
            ForPa = ListView1.ListItems(I).SubItems(2)
            dtoPP = ListView1.ListItems(I).SubItems(4)
            dtoGn = ListView1.ListItems(I).SubItems(5)
            SQL = "(numalbar=" & DBSet(ListView1.ListItems(I).Text, "T") & " and "
            SQL = SQL & "fechaalb=" & DBSet(ListView1.ListItems(I).SubItems(1), "F") & ")"
            If cadAux = "" Then
                cadAux = SQL
            Else
                cadAux = cadAux & " OR " & SQL
            End If
        End If
    Next I
    
    If cadAux <> "" Then
    'se han seleccionado albaranes para facturar
    'Esta el la cadena WHERE de los albaranes seleccionados para obtener
    'el bruto de las lineas de los albaranes agrupadas por tipo de iva
        cadWhere = "slialp.codprove=" & Val(Text1(3).Text)
        cadWhere = cadWhere & " AND (" & cadAux & ")"
        
        
        
        
    Else
        Exit Sub
    End If
    
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWhere) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = dtoPP
    vFactu.DtoGnral = dtoGn
    
    
    exento = False
    If vUsu.TrabajadorB Then exento = True
    
    CambioIVA = False
    If Not exento Then
        If Text1(1).Text <> "" Then
            If CDate(Text1(1).Text) < CDate("01/09/2012") Then CambioIVA = True
        End If
    End If
    vFactu.FijarTipoIvaProveedor Val(Text1(3).Text)
    If vFactu.CalcularDatosFactura2(cadWhere, "scaalp", "slialp", CambioIVA, exento) Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        
        For I = 6 To 22
            FormateaCampo Text1(I)
        Next I
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For I = 11 To 20 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For I = 12 To 21 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
   
End Sub



Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van a�adiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWhere = "" Then Exit Function
    cadWhere = Replace(cadWhere, "slialp", "scaalp")
    
    SQL = "Select count(*) FROM scaalp"
    SQL = SQL & " WHERE " & cadWhere
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim cad As String

    Screen.MousePointer = vbHourglass
    
    
    
    cad = ""
    If Text1(3).Text = "" Then
        cad = "Falta proveedor"
    Else
        If Not IsNumeric(Text1(3).Text) Then cad = "Campo proveedor debe ser num�rico"
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
        
        
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(Text1(3).Text) Then
        Set vProve = Nothing
        Exit Sub
    End If
        
    
    
    If Not DatosOk Then Exit Sub
    
    PonerModo2 5
End Sub


Private Function GenerarFactura_() As Boolean
Dim vFactu As CFacturaCom

        On Error GoTo Error1
        GenerarFactura_ = False
        
        'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = Text1(3).Text
        vFactu.NumFactu = Text1(0).Text
        vFactu.Fecfactu = Text1(1).Text
        vFactu.FecRecep = Text1(2).Text
        vFactu.Trabajador = Text1(4).Text
        vFactu.BancoPr = Text1(5).Text
        vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
        vFactu.ForPago = ForPa
        vFactu.DtoPPago = dtoPP
        vFactu.DtoGnral = dtoGn
        vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
        vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
        vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
        vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
        vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
        vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
        vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
        vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
        vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
        vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
        vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
        vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
        vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
        vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
        vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
        
        'Sobre que calcual la retencion, si sobre el tota o sobre las bases(sun iva)
        If chkTipoRet.Value = 1 Then
            vFactu.TipoRet = 0
        Else
            vFactu.TipoRet = 1
        End If
        
        
        vFactu.PorRet = ImporteFormateado(Text1(23).Text)
        vFactu.ImpRet2 = ImporteFormateado(Text1(24).Text)
        
        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan
        vFactu.IBAN = vProve.IBAN
        
        'Si es trabajador de B factura en B
        vFactu.Facb = vUsu.TrabajadorB
            
        
        If vFactu.TraspasoAlbaranesAFactura(cadWhere, (Check1(0).Value = 1), (Check1(1).Value = 1), False) Then
            'Antes
            'BotonPedirDatos
            'AHora
            LimpiarCampos
            Me.ListView1.ListItems.Clear
            PonerModo2 0
        End If
        Set vFactu = Nothing
        Set vProve = Nothing
    GenerarFactura_ = True
 

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el n�mero de factura para el proveedor en hco
    cad = "SELECT count(*) FROM scafpc "
    cad = cad & " WHERE codprove=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(cad) > 0 Then
        MsgBox "Factura de proveedor ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function

Private Sub RefrescarAlbaranes()
Dim I As Integer
Dim SQL As String
Dim Itm As ListItem
Dim RS As ADODB.Recordset
    

    For I = 1 To ListView1.ListItems.Count
        SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
        SQL = SQL & " sum(slialp.importel) as bruto "
        SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
        SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text & " AND scaalp.numalbar=" & DBSet(ListView1.ListItems(I).Text, "T") & " AND scaalp.fechaalb=" & DBSet(ListView1.ListItems(I).SubItems(1), "F")
        SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
        SQL = SQL & " ORDER BY scaalp.numalbar"

        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        If Not RS.EOF Then 'Actualizamos los datos de este item en el list
            ListView1.ListItems(I).SubItems(2) = RS!codforpa
            ListView1.ListItems(I).SubItems(3) = RS!nomforpa
            ListView1.ListItems(I).SubItems(4) = RS!DtoPPago
            ListView1.ListItems(I).SubItems(5) = RS!DtoGnral
            ListView1.ListItems(I).SubItems(6) = RS!bruto

        End If
        
        If ListView1.ListItems(I).Checked Then 'comprobamos otra vez el chek y recalculamos factura
            Set Itm = ListView1.ListItems(I)
            ListView1_ItemCheck Itm
        End If

        RS.Close
        Set RS = Nothing
    Next I
    
    'recalcular el total de la factura
     For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            CalcularDatosFactura
            Exit For
        End If
     Next I
     
End Sub





Private Function PonerDatosProveedor() As String
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select nomprove,sprove.codbanpr,nombanpr from sprove ,sbanpr where sprove.codbanpr= sbanpr.codbanpr  and sprove.codprove =" & Text1(3).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Devolvemos el nombre del prove y fijamos la cadena del banco
    If miRsAux.EOF Then
        PonerDatosProveedor = ""
        Text1(5).Text = ""
        Text2(5).Text = ""
    Else
        PonerDatosProveedor = miRsAux!nomprove
        Text1(5).Text = miRsAux!codbanpr
        Text2(5).Text = miRsAux!nombanpr
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function

Private Function ComprobarNumerosLotes() As Boolean
Dim Aux As String
    
    ComprobarNumerosLotes = True
    If vParamAplic.EsAVAB Then Exit Function         'EN AVAB No comprobamos numeros de lote
        
    

    Aux = " Select slialp.codartic,slialp.nomartic,numlotes from slialp,sartic"
    Aux = Aux & " where slialp.codartic=sartic.codartic  and trazabilidad=1"
    Aux = Aux & " and numalbar=" & DBSet(ListView1.SelectedItem.Text, "T")
    Aux = Aux & " and fechaalb=" & DBSet(ListView1.SelectedItem.SubItems(1), "F")
    Aux = Aux & " and slialp.codprove=" & Text1(3).Text
    'and fechaalb = '2009-11-04' and slialp.codprove=163"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        If DBLet(miRsAux!numlotes, "T") = "" Then Aux = Aux & miRsAux!codArtic & "  " & miRsAux!NomArtic & vbCrLf
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    ComprobarNumerosLotes = True
    If Aux <> "" Then
        Aux = "Lineas que deberian ir con numero de lote(Trazabilidad)" & vbCrLf & String(40, "=") & vbCrLf & Aux
        Aux = Aux & "�Continuar?"
        If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then ComprobarNumerosLotes = False
    End If
End Function
'
'
'    cLot.tipoMov = 0  'Salida
'    cLot.Cantidad = Abs(Cantidad)
'    cLot.codalmac = cPar.codalmac
'    cLot.codartic = cPar.codartic
'    cLot.codarti2 = ArticuloProduccion
'    cLot.Numlote = cPar.Numlote
'
'    If Not cLot.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cPar.codartic
'
