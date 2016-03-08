VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmFacTPVEnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11910
   Icon            =   "frmFacTPVEnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   12
      Left            =   11280
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   40
      Tag             =   "P|T|S|||sliven|dto2|||"
      Text            =   "Dto"
      Top             =   5760
      Visible         =   0   'False
      Width           =   430
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   11
      Left            =   11280
      MaxLength       =   12
      TabIndex        =   39
      Tag             =   "PorcentajeIVA|N|N|||sliven|dto2|||"
      Text            =   "%IVA"
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   10
      Left            =   11040
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   38
      Tag             =   "precioarticulo|N|N|||sliven|dto2|||"
      Text            =   "dto2"
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   11040
      MaxLength       =   12
      TabIndex        =   37
      Tag             =   "Dto2|N|N|||sliven|dto2|||"
      Text            =   "dto2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   10680
      MaxLength       =   12
      TabIndex        =   36
      Tag             =   "Dto1|N|S|||sliven|codigiva||N|"
      Text            =   "dto1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   27
      Top             =   2700
      Width           =   11760
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   7425
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   240
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2280
         MaxLength       =   80
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   620
         Width           =   9015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   240
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label label1 
         Caption         =   "Dpto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   6440
         TabIndex        =   35
         Top             =   240
         Width           =   560
      End
      Begin VB.Image imgBuscar 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   7040
         ToolTipText     =   "Buscar artículo"
         Top             =   220
         Width           =   360
      End
      Begin VB.Label label1 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1100
         ToolTipText     =   "Buscar artículo"
         Top             =   220
         Width           =   360
      End
      Begin VB.Label label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   900
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1680
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10560
      TabIndex        =   26
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   11040
      MaxLength       =   12
      TabIndex        =   25
      Tag             =   "Tipo iva|N|N|0||sliven|codigiva|0|N|"
      Text            =   "tipoiva"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame FrameTot 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4080
      TabIndex        =   19
      Top             =   1080
      Width           =   7650
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   8
         Left            =   4880
         TabIndex        =   24
         Top             =   100
         Width           =   2660
      End
      Begin VB.Label label1 
         BackColor       =   &H00F5F5F5&
         Caption         =   "TOTAL LINEA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   100
         Width           =   4720
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   6
         Left            =   3720
         TabIndex        =   21
         Top             =   825
         Width           =   3825
      End
      Begin VB.Label label1 
         BackColor       =   &H00F5F5F5&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   825
         Width           =   2775
      End
   End
   Begin VB.Frame FrameCab 
      Height          =   1725
      Left            =   120
      TabIndex        =   14
      Top             =   975
      Width           =   3510
      Begin VB.ComboBox cboNumVenta 
         Height          =   315
         ItemData        =   "frmFacTPVEnt.frx":000C
         Left            =   1800
         List            =   "frmFacTPVEnt.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label label1 
         Caption         =   "Nº Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotVentas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   345
         Index           =   4
         Left            =   2280
         TabIndex        =   18
         Top             =   675
         Width           =   975
      End
      Begin VB.Label label1 
         Caption         =   "Ventas Abiertas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   675
         Width           =   2055
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label label1 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   840
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "Cod.EAN|T|S|||slitick|codigoea|||"
      Text            =   "Artic EAN"
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   10080
      MaxLength       =   12
      TabIndex        =   8
      Tag             =   "Precio art.|N|N|0|999999.0000|slitick|precioar|###,##0.0000|N|"
      Text            =   "Precio ar."
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "Linea|N|N|||slitick|numlinea|0|S|"
      Text            =   "line"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Cod. Artículo|T|N|||slitick|codartic|||"
      Text            =   "Artic Artic Artic5"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   6840
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Cantidad|N|N|||slitick|cantidad|#,###,###,##0.00||"
      Text            =   "1,234,567,891.25"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   8160
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "Precio|N|N|0|999999.0000|slitick|precioiv|###,##0.0000|N|"
      Text            =   "123,456.7879"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   9120
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Importe|N|N|||slitick|importel|#,###,###,##0.00|N|"
      Text            =   "Importe"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "Nombre artículo|T|N|||slitick|nomartic|||"
      Text            =   "nomArtic"
      Top             =   6600
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      ToolTipText     =   "Buscar artículo"
      Top             =   6600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6480
      Top             =   7680
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
      Height          =   630
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1111
      ButtonWidth     =   1535
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Venta  F4"
            Object.ToolTipText     =   "Nueva Venta"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Borrar  F6"
            Object.ToolTipText     =   "Eliminar Venta"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Total  F5"
            Object.ToolTipText     =   "Total Venta"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Termi."
            Object.ToolTipText     =   "Ventas de otros terminales"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar F2"
            Object.ToolTipText     =   "Buscar Artículo"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Linea  F7"
            Object.ToolTipText     =   "Eliminar linea"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Revisar F9"
            Object.ToolTipText     =   "Revisar ventas dia"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7920
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Left            =   10530
      TabIndex        =   10
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacTPVEnt.frx":0010
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   1650
      Left            =   3720
      Top             =   1050
      Width           =   7755
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nueva Venta"
         HelpContextID   =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar Venta"
         HelpContextID   =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnTotal 
         Caption         =   "&Total Venta"
         HelpContextID   =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnTraerVenta 
         Caption         =   "&Otros termi."
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineasBuscar 
         Caption         =   "&Buscar Artículo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnLineasElim 
         Caption         =   "Eliminar &Linea"
         HelpContextID   =   2
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnRevisarVentasDia 
         Caption         =   "Revisión ventas día"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnbarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacTPVEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NomrePC_conectado As String
                              

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
'Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Private WithEvents frmCli As frmFacClientes2   'Form clientes (busquedas)
Attribute frmCli.VB_VarHelpID = -1

'Pantalla para traer la venta de otro terminal
Private WithEvents frmTraerVen As frmFacTPVTraerVen
Attribute frmTraerVen.VB_VarHelpID = -1


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
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim PrimeraVez As Boolean


'Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

'Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
'WHERE para seleccionar una venta de otro terminal
Private CadSelVenta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Dim SQL As String
Dim i As Integer


Dim NumTermi As Byte

Dim CodTraba As String 'trabajador conectado
Dim NomTraba As String
Dim codAlmac As String 'almacen por defecto del trabajador



Private Sub cboNumVenta_Click()
   Modo = 2
   ModificaLineas = 0
   
   If PosicionarData Then
        CargaGrid Me.DataGrid1, Me.Data2, True
        
        If Me.Data1.Recordset.AbsolutePosition > 0 Then
            If Not Data1.Recordset.EOF Then
                'poner el total
                Me.label1(6).Caption = Format(DevuelveDesdeBDNew(conAri, "scaven", "imptotal", "numtermi", CStr(Data1.Recordset!NumTermi), "N", , "numventa", CStr(Data1.Recordset!NumVenta), "N", "fecventa", CStr(Data1.Recordset!fecventa), "F"), FormatoImporte)
                'poner el cliente
                Me.Text1(0).Text = Data1.Recordset!CodClien
                Me.Text1(0).Text = Format(Text1(0).Text, "000000")
                Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                'poner las observaciones de la venta
                Me.Text1(1).Text = DBLet(Data1.Recordset!observa1, "T")
                'poner direc./ departamento
                If Not IsNull(Data1.Recordset!CodDirec) Then
                    Text1(2).Text = DBLet(Data1.Recordset!CodDirec, "N")
                    Me.Text1(2).Text = Format(Text1(2).Text, "000")
                    Text2(2).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "coddirec", Text1(2).Text, "N", , "codclien", Text1(0).Text, "N")
                Else
                    Text1(2).Text = ""
                    Text2(2).Text = ""
                End If
'                Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "sdirec", "nomdirec", "coddirec", "cliente", "N")
            End If
                
            If vParamTPV.HayVisor Then EnviarVisorPuerto label1(7).Caption, label1(8).Caption, label1(5).Caption, label1(6).Caption
                
'                If Data2.Recordset.EditMode = adEditNone Then
'                    If Me.Data2.Recordset.AbsolutePosition <> Me.Data2.Recordset.RecordCount Then
                    If Data2.Recordset.RecordCount > DataGrid1.VisibleRows Then
                        Data2.Recordset.MoveLast
                        If DataGrid1.Row < 0 Then DataGrid1.Row = DataGrid1.VisibleRows - 1
                    End If
'                End If
                
                BotonAnyadirLinea
        End If
    End If
End Sub





Private Sub cmdAux_Click(Index As Integer)
'    Select Case Index
'        Case 1 'Busqueda de Cod. Artic
'            Set frmArt = New frmAlmArticulos
'            frmArt.DatosADevolverBusqueda = "1" 'Poner en Modo busqueda
'            frmArt.Show vbModal
'            Set frmArt = Nothing
'            txtAux_LostFocus (1)
'    End Select
'    PonerFoco txtAux(Index)
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Ventas (scaven)
' y los registros correspondientes de las tablas de lineas (sliven)
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    Cad = "Va a eliminar la venta:         " & vbCrLf & vbCrLf
    Cad = Cad & "     Nº venta:  " & Me.Data1.Recordset!NumVenta & vbCrLf
    Cad = Cad & "     Fecha:  " & Format(Me.Data1.Recordset!fecventa, "dd-mm-yyyy")
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
       
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition

        If Not Eliminar() Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            PosicionarDataTrasEliminar
            'Volvemos a cargar las ventas que quedan ahora
            PonerVentasAbiertas
            cargaComboVentas
            If Not Data1.Recordset.EOF Then
                PosicionarComboVentas (Me.Data1.Recordset!NumVenta)
            Else
                'Limpiar los totales
                ReiniciarVisor
                cmdCancelar_Click
                Text1(0).Text = ""
                Text2(0).Text = ""
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Venta.", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea de la venta. Tabla: sliven

    If Data2.Recordset.EOF Then Exit Sub
        
    If DBLet(Data2.Recordset!codArtic, "T") = "" Then Exit Sub
    SQL = "¿Seguro que desea eliminar la línea de venta?     " & vbCrLf
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    SQL = SQL & vbCrLf & "Importe:  " & Data2.Recordset!ImporteL
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
'            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
             
            If Not SituarDataTrasEliminar(Data2, NumRegElim, True) Then
                If Data2.Recordset.RecordCount = 0 Then
                'Elimine el ultimo registro
                    ReiniciarVisor
                    BotonAnyadirLinea
                End If
            End If
            If DataGrid1.Enabled Then DataGrid1.SetFocus
        End If
'        CancelaADODC
    End If

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar linea Venta.", Err.Description
End Sub



Private Sub cmdCancelar_Click()
    On Error Resume Next
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Data2.Recordset.EOF Then Me.Data2.Recordset.MoveLast
            If ModificaLineas = 1 Then
                ModificaLineas = 0
                Me.DataGrid1.Enabled = True
            End If
            PonerFocoGrid DataGrid1
'            If DataGrid1.Enabled Then DataGrid1.SetFocus
'            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'        Case 4 'Modificar
'            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    
    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Select Case KeyCode
'        Case 37 'Flecha izquierda

'        Case 38 'Flecha Arriba

        Case 40 'flecha Abajo
            'Pasar a la siguiente linea
            'si estamos en la ultima linea añadimos nueva linea
            If Me.Data2.Recordset.AbsolutePosition = Me.Data2.Recordset.RecordCount Then
                KeyCode = 0
                BotonAnyadirLinea
            End If
    End Select

    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then 'ESC
        cmdCancelar_Click
        Unload Me
    End If
End Sub



Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Data2.Recordset.EOF Or Me.Data2.Recordset.BOF Then
        Me.label1(7).Caption = ""
        Me.label1(8).Caption = ""
    Else
        'Actualizar el total de linea
        If (ModificaLineas = 0) Or (ModificaLineas = 1 And Data2.Recordset.AbsolutePosition = Data2.Recordset.RecordCount And Me.Data2.Recordset!NomArtic <> "") Then
            Me.label1(7).Caption = Mid(Me.Data2.Recordset!NomArtic, 1, 19)
            Me.label1(8).Caption = Format(Me.Data2.Recordset!ImporteL, FormatoImporte)
        End If
    End If
    If vParamTPV.HayVisor Then
        EnviarVisorPuerto label1(7).Caption, label1(8).Caption, "TOTAL", label1(6).Caption
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If vParamTPV Is Nothing Then Unload Me
    If CodTraba = "" Then Unload Me
    
    If PrimeraVez Then
        PrimeraVez = False
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    Me.imgBuscar(0).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(1).Picture = frmPpal.ImgListPpal.ListImages(17).Picture

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImageListTPV
        .Buttons(2).Image = 1   'Nueva Venta
        .Buttons(4).Image = 2   'Eliminar Venta
        .Buttons(6).Image = 3   'TOTALIZAR venta
        .Buttons(8).Image = 7   'Traer venta de otro terminal

        .Buttons(12).Image = 4   'buscar articulo
        .Buttons(14).Image = 5   'Eliminar linea venta
        
        .Buttons(17).Image = 8   'Revisar ventas dia
        
        .Buttons(21).Image = 6  'Salir
    End With

    PrimeraVez = True
    LimpiarCampos   'Limpia los campos TextBox

   
    'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName
    SQL = Me.NomrePC_conectado
    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
'    MsgBox "Nº de terminal para " & Me.NomrePC_conectado & ": " & SQL, vbInformation
    If Not IsNumeric(SQL) Then
        MsgBox "No se ha configurado el terminal de venta." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
        Exit Sub
    End If
    NumTermi = CInt(SQL)
   
    
    'Leemos los parametros del TPV
    Set vParamTPV = New CParamTPV
    If vParamTPV.Leer() = 1 Then
        SQL = "No se han podido cargar los Parámetros generales del TPV." & vbCrLf
        MsgBox SQL & "Debe configurar la aplicación.", vbExclamation
        Set vParamTPV = Nothing
        Unload Me
    ElseIf vParamTPV.Leer2(CStr(NumTermi)) = 1 Then
        SQL = "No se han podido cargar los Parámetros del terminal TPV." & vbCrLf
        MsgBox SQL & "Debe configurar la aplicación.", vbExclamation
        Set vParamTPV = Nothing
        Unload Me
    End If
    

    'Poner el trabajador que esta conectado
    CodTraba = PonerTrabajadorConectado(NomTraba)
    If CodTraba = "" Then
        SQL = "No se ha encontrado ningún trabajador con ese login." & vbCrLf
        SQL = SQL & "Compruebe que el trabajador tiene asignado su login."
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Almacen por defecto el del trabajador
    If CodTraba <> "" Then
        codAlmac = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", CodTraba, "N")
    Else
        codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    End If
    


    'Compruebo si tiene mas de una forma de pago
    '
    Set miRsAux = New ADODB.Recordset
    SQL = "SELECT count(*) from sforpa"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") = 1 Then SQL = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If SQL = "" Then vParamTPV.FormaPagoUnica = True


    NombreTabla = "scaven" 'tabla cabecera de la venta
    NomTablaLineas = "sliven" 'Tabla lineas de venta
    Ordenacion = " ORDER BY numtermi, numventa, fecventa "
    CadSelVenta = "" 'selecciona la venta de otro terminal

    'Poner los grid sin apuntar a nada
    LimpiarDataGrids

    'Cargamos el Data1 con las cabeceras
    Data1.ConnectionString = Conn

    'Inicializar Cabecera de Caja
    Me.label1(2).Caption = Format(Now, "dd-mm-yy  hh:mm")


    'Abrir Visor en puerto serie
    AbrirVisorPuerto

    'Iniciar Totales y los envia al visor del puerto
    ReiniciarVisor

    'Comprobar si hay ventas abiertas
    PonerVentasAbiertas


    If CInt(Me.label1(4).Caption) > 0 Then
        cargaCabeceras True
    Else
        cargaCabeceras False
    End If

    cargaComboVentas
    
End Sub



Private Sub PonerVentasAbiertas()
    SQL = "SELECT count(*) FROM " & NombreTabla & " WHERE numtermi=" & NumTermi
    If CadSelVenta <> "" Then SQL = SQL & " OR " & CadSelVenta
    Me.label1(4).Caption = CStr(NumRegistros(SQL)) 'nº ventas abiertas
End Sub


Private Sub cargaComboVentas()
Dim Rs As ADODB.Recordset
Dim N As Integer
    On Error GoTo ECargaCombo
    
    Me.cboNumVenta.visible = False
    Me.cboNumVenta.Clear

    SQL = "SELECT * FROM " & NombreTabla & " WHERE numtermi=" & NumTermi
    If CadSelVenta <> "" Then SQL = SQL & " OR " & CadSelVenta
    SQL = SQL & " ORDER BY numventa"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 0
    SQL = "                      "
    While Not Rs.EOF
        N = 21 - Len(CStr(Rs!NumVenta))
            
        Me.cboNumVenta.AddItem Right(SQL & Rs!NumVenta, N)
        cboNumVenta.ItemData(cboNumVenta.NewIndex) = i
        Me.cboNumVenta.RightToLeft = True
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Me.cboNumVenta.visible = True
    
ECargaCombo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos del combo.", Err.Description
End Sub



Private Sub cargaCabeceras(enlaza As Boolean)

    'ASignamos un SQL al DATA1
    If Not enlaza Then
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " where numventa=-1"
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE numtermi=" & NumTermi
        If CadSelVenta <> "" Then CadenaConsulta = CadenaConsulta & " OR " & CadSelVenta
        CadenaConsulta = CadenaConsulta & " ORDER BY numtermi,numventa,fecventa "
    End If
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
End Sub



Private Sub LimpiarCampos()
    On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Error1
    'cerramos el puerto serie del visor
    If Not vParamTPV Is Nothing Then
        If vParamTPV.HayVisor Then
            If Me.MSComm1.PortOpen Then
                Me.MSComm1.Output = Mid(vEmpresa.nomempre & Space(40), 1, 40)
                Me.MSComm1.PortOpen = False
            End If
        End If
        
        Set vParamTPV = Nothing
    End If
    
Error1:
    If Err.Number Then Err.Clear
End Sub


'Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'''Mantenimiento de Articulos
''    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
''    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
'End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If frmB.vTabla <> "sdirec" Then
        'ARTICULOS
        If ModificaLineas <> 1 Then Exit Sub
    
        If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            If ModificaLineas = 1 Then
                txtAux(2).Text = RecuperaValor(CadenaDevuelta, 1)
                'txtAux_LostFocus (2)
            End If
        End If
    Else
        'DIREC./DEPARTAMENTOS
         If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
'            If ModificaLineas = 1 Then
                Text1(2).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(2).Text = RecuperaValor(CadenaDevuelta, 2)
                'txtAux_LostFocus (2)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'clientes (busquedas)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod clien
    Text1(0).Text = Format(Text1(0).Text, "000000")
    
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub



Private Sub frmTraerVen_CargarVenta(cadSel As String, nVen As Long)
 'Traer venta de otro terminal
    CadSelVenta = cadSel
    
    If nVen < 0 Then
        'He mamdado a buscar las ventas del dia
        'Nven:      -1 Imprimir tick
        '           -2 Ver lineas ticket
        CadSelVenta = Abs(nVen) & "|" & cadSel & "|"
    
    Else
        
        PonerVentasAbiertas
        
        cargaCabeceras True
        cargaComboVentas
        
        
        PosicionarComboDes Me.cboNumVenta, CStr(nVen)

    End If

End Sub


Private Sub imgBuscar_Click(Index As Integer)

    On Error GoTo ErrImg
    
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 0 'clientes
            Set frmCli = New frmFacClientes2
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(0)
            
'            If CLng(Data1.Recordset!CodClien) <> CLng(Text1(0).Text) Then
'             'si se ha cambiado el cliente actualizar la cabecera venta (scaven)
'                ModificarVenta
'            End If
            
'           txtAux_LostFocus (1)

        Case 1 'Departamento
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(0).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                MandaBusquedaPrevia_dpto " codclien= " & Val(Text1(0).Text)
'                Indice = 12
             End If
    End Select
    Exit Sub
    
ErrImg:
    MuestraError Err.Number, "Busqueda de clientes", Err.Description
End Sub

Private Sub mnEliminar_Click()
'Elimina una venta
    BotonEliminar
End Sub



Private Sub mnLineasBuscar_Click()
Dim b As Boolean

    If Data1.Recordset.EOF Then Exit Sub
'    If ModificaLineas = 0 Then Exit Sub

    b = (Screen.ActiveControl.Name = "Text1")
    If b Then b = (Screen.ActiveControl.Index = 0) Or (Screen.ActiveControl.Index = 2)
    
    If b Then
        If (Screen.ActiveControl.Index = 0) Then
            imgBuscar_Click (0)
        ElseIf (Screen.ActiveControl.Index = 2) Then
            imgBuscar_Click (1)
        End If
    Else
        MandaBusquedaPrevia ""
    End If
End Sub

Private Sub mnLineasElim_Click()
'Elimina una linea de venta
    BotonEliminarLinea
End Sub

Private Sub mnNuevo_Click()
'Inicia una nueva venta
    BotonAnyadir
End Sub


Private Sub mnRevisarVentasDia_Click()
    RevisarEntradasDia
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String

    'Llamamos a al form
    '##A mano
    Cad = ""
    'tabla = "sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic "
    
    
     'SELECT `sartic`.`codartic`, `salmac`.`codalmac`, `slista`.`precioac`
    tabla = "(`sartic` `sartic` LEFT OUTER JOIN `salmac` `salmac` ON `sartic`.`codartic`=`salmac`.`codartic`)"
    tabla = tabla & " LEFT OUTER JOIN `slista` `slista` ON `sartic`.`codartic`=`slista`.`codartic`"
    
    
    
    Titulo = "Artículos"
    devuelve = "0|1|2|"
    Cad = Cad & "Cod. Artic.|sartic|codartic|T||22·"
    Cad = Cad & "Des. Artic.|sartic|nomartic|T||57·"
'    cad = cad & "Precio|sartic|preciove|N|###,##0.0000|15·"
    Cad = Cad & "Stock|salmac|canstock|N|#,###,###,##0.00|8·"
    Cad = Cad & "Precio|slista|precioac|N|#,###,###,##0.00|12·"
    
    If cadB = "" Then
        cadB = " codalmac = " & codAlmac
    Else
        cadB = cadB & " AND codalmac=" & codAlmac
    End If
    'La tarifa a buscar SIEMPRE es la de parametros
    cadB = cadB & " AND ( codlista = " & vParamAplic.CodTarifa & " OR codlista is null)"
           
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        '#
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
        If ModificaLineas = 1 Then
            If txtAux(2).Text <> "" Then
                PonerFoco txtAux(2)
                txtAux_LostFocus (2)
            Else
                PonerFoco txtAux(1)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub MandaBusquedaPrevia_dpto(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
'para busquedas de direc./depatamentos del cliente
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String
Dim Desc As String

    'Llamamos a al form
    '##A mano
    Cad = ""
    
    If vParamAplic.Departamento Then
        Titulo = "Dptos Cliente: "
        Desc = "Dpto."
    Else
        Titulo = "Direc. Cliente: "
        Desc = "Direc."
    End If
    Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
    Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
    Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
    tabla = "sdirec"
    devuelve = "0|1|"
    
    
    
   
'    Tabla = "sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic "
'    Titulo = "Artículos"
'    devuelve = "0|1|2|"
'    cad = cad & "Cod. Artic.|sartic|codartic|T||25·"
'    cad = cad & "Des. Artic.|sartic|nomartic|T||60·"
'    cad = cad & "Stock|salmac|canstock|N|#,###,###,##0.00|15·"
    
'    If cadB = "" Then
'        cadB = " codalmac = " & codAlmac
'    Else
'        cadB = cadB & " AND codalmac=" & codAlmac
'    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        '#
        frmB.label1.FontSize = 11
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
        
'        If ModificaLineas = 1 Then
'            If txtAux(2).Text <> "" Then
'                PonerFoco txtAux(2)
'                txtAux_LostFocus (2)
'            Else
'                PonerFoco txtAux(1)
'            End If
'        End If
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
'            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
'        PonerModo 2
'        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
    
    Modo = Kmodo
    
    
    b = (Modo = 3) Or (Modo = 4)
   
    'Si no es modo lineas Boquear los TxtAux
    For i = 1 To 5
        BloquearTxt txtAux(i), (Modo <> 3)
        txtAux(i).visible = (Modo = 3)
    Next i
    txtAux(12).visible = txtAux(5).visible
    txtAux2(2).visible = (Modo = 3)
    'El Importe siempre bloqueadoç
    BloquearTxt txtAux(5), True
    
    DataGrid1.Enabled = (Modo = 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
'    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    
    
    
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Antes de insertar una linea de venta comprueba que los datos son OK
Dim b As Boolean

    On Error GoTo EDatosOkLinea

    b = CompForm(Me, 3)
    
    
    If b Then
        If txtAux2(2).Text = "" And txtAux(2).Text <> "" Then
            MsgBox "Codigo artículo incorrecto", vbExclamation
            b = False
        End If

    End If
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub mnTotal_Click()
'Dim cCli As CCliente
Dim SQL As String
Dim bModif As Boolean

    'comprobar q el campo de cliente tiene valor
    If Trim(Text1(0).Text) = "" Then
        MsgBox "El campo Cliente debe tener valor.", vbInformation
        Exit Sub
    End If


    '-- Si el cliente esta bloqueado no permitimos generar ticket/Albaran/Factura
    '-- y no abrimos la pantalla de totales
    '-- si el cliente se ha modificado y no tiene la misma tarifa q para el q se crearon las lineas tampoco
    
    'obtenemos el cliente actual guardado en la cabecera de la venta
    SQL = ""
    If Not Data1.Recordset.EOF Then SQL = Data1.Recordset!CodClien
    If Not ClienteOK(Text1(0).Text, SQL, bModif) Then
        Text1(0).Text = SQL
        Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
        PonerFoco Text1(0)
        Exit Sub
    End If
'    If bModif Then ModificarVenta
    
    
    '-- si no hay lineas en la venta no permitimos generar ticket/Albaran/Factura
    '-- y no abrimos la pantalla de totales
    SQL = ObtenerWhereCP(False)
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    SQL = "SELECT COUNT(*) FROM " & NomTablaLineas & " WHERE " & SQL
    If Not RegistrosAListar(SQL) > 0 Then
        MsgBox "La venta debe tener líneas para totalizar.", vbExclamation
        Exit Sub
    End If
    
    cmdCancelar_Click
    
    Me.label1(2).Caption = Format(Now, "dd-mm-yy  hh:mm")
'    If CDate(Data1.Recordset!fecventa) <> CDate(Format(Now, "dd/mm/yyyy")) Then
        'Actualizar la fecha de la venta
        ModificarVenta
'    End If
    
    frmFacTPVTotal.Importe = Me.label1(6).Caption
    frmFacTPVTotal.cadSel = ObtenerWhereCP(False)
'    frmFacTPVTotal.NumTermi = NumTermi
    frmFacTPVTotal.Show vbModal


    If frmFacTPVTotal.cadSel = "1" Then 'Se genero correctamente el documento(ticket,alb,factu)
        'Refrescar los datos
        'Volvemos a cargar las ventas que quedan ahora
        PonerVentasAbiertas
        cargaCabeceras (True)
        cargaComboVentas
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        ReiniciarVisor
        LimpiarCampos
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    End If
End Sub



Private Sub mnTraerVenta_Click()
'Traer venta de otro terminal
    CadSelVenta = ""
    Set frmTraerVen = New frmFacTPVTraerVen
    frmTraerVen.parNumTermi = NumTermi
    frmTraerVen.Show vbModal
    Set frmTraerVen = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cCli As CCliente
Dim devuelve As String
Dim bModif As Boolean

    devuelve = ""
    
    Select Case Index
        Case 0 'cod CLIENTE
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                
                If Text2(Index).Text = "" Then
                    'no existe el cliente y no sale del campo
                    PonerFoco Text1(Index)
                Else
                    If Not (Me.Data1.Recordset.EOF And Me.Data1.Recordset.BOF) Then '(RAFA/ALZIRA 31082006) Al poner el foco directamente en el cliente este trozo falla cuando se pierde el foco sin que haya ninguna venta
                        devuelve = Me.Data1.Recordset!CodClien
                    End If
                        
                    If ClienteOK(Text1(Index).Text, devuelve, bModif, True) Then
                        If bModif Then ModificarVenta
                    Else
                        Text1(Index).Text = devuelve
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                        Text1(Index).Text = Format(Text1(Index).Text, "000000")
                        PonerFoco Text1(Index)
                    End If
'                    Set cCli = New CCliente
'                    If cCli.LeerDatos(Text1(Index).Text) Then
'                        'no permitimos ventas a clientes bloqueados
'                        If cCli.ClienteBloqueado Then
'                            PonerFoco Text1(Index)
'                            Set cCli = Nothing
'                            Exit Sub
'                        End If
'
'                        If Not (Me.Data1.Recordset.EOF And Me.Data1.Recordset.BOF) Then '(RAFA/ALZIRA 31082006) Al poner el foco directamente en el cliente este trozo falla cuando se pierde el foco sin que haya ninguna venta
'                            If CLng(Me.Data1.Recordset!CodClien) <> CLng(Text1(0).Text) Then 'OPP
'                                '--- Laura: 11/04/2007
'                                '--- comprobar q la tarifa del nuevo cliente es la misma q la del cliente q
'                                '--- habia antes siempre y cuando haya lineas de precios ya q si no no estariamos
'                                '--- aplicando la tarifa correcta al cliente
'                                If Not Me.Data2.Recordset.EOF Then 'si hay lineas
'                                    'obtener la tarifa del cliente actual
'                                    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", CStr(Me.Data1.Recordset!CodClien), "N")
'                                    If devuelve <> CStr(cCli.Tarifa) Then
'                                        devuelve = "No se puede seleccionar el cliente " & Text1(0).Text & " "
'                                        devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
'                                        devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
'                                        MsgBox devuelve, vbExclamation, "Comprobar cliente"
'                                        Text1(Index).Text = Me.Data1.Recordset!CodClien
'                                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
'                                        Text1(Index).Text = Format(Text1(Index).Text, "000000")
'                                        PonerFoco Text1(Index)
'                                        Set cCli = Nothing
'                                        Exit Sub
'                                    End If
'                                End If
'                                ModificarVenta
'                            End If
'                        End If
                    
                        'mostrar las observaciones del cliente
'                        cCli.MostrarObservaciones
'                    End If
'                    Set cCli = Nothing
                End If
            Else
                'si el formato no es numerico no lo aceptamos
                Text2(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 1 'OBSERVACIONES
            '---- Laura: 06/10/2006
            'Poner el foco en la linea
            If Screen.ActiveControl.Name <> "Text1" Then
                PonerFoco txtAux(1)
                ModificarVenta
            End If
            
        Case 2 'DIREC./DPTO
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                
                'Comprobar que el cliente seleccionada tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(0).Text, "N", , "coddirec", Text1(2).Text, "N")
                    If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
                Else
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Nueva venta
            mnNuevo_Click
            
        Case 4  'Borrar Venta
            mnEliminar_Click
            
        Case 6  'Total Venta
            mnTotal_Click
            
        Case 8 'traer venta de otro terminal
            mnTraerVenta_Click
            
        Case 12 'Buscar articulo
            mnLineasBuscar_Click
            
        Case 14  'Eliminar Linea
            mnLineasElim_Click
                    
        Case 17
            'Revision entradas dia
            mnRevisarVentasDia_Click
        
        Case 21    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If KeyAscii = 27 Then cmdCancelar_Click
    If cerrar Then Unload Me
End Sub

  

Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    vData.LockType = adLockOptimistic
    
    CargaGrid2 vDataGrid, vData
    DataGrid1.RowHeight = 320
    
    If Data2.Recordset.RecordCount = 0 Then DataGrid1_RowColChange 1, 1
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid2

    vDataGrid.ScrollBars = dbgNone
    vData.Refresh

    tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(1)|T|Cod. EAN|1550|;S|txtAux(2)|T|Código|1650|;S|cmdAux(1)|B||0|;S|txtAux2(2)|T|Desc. Artículo|4050|;"
    tots = tots & "S|txtAux(3)|T|Cantidad|900|;"
'    If vParamTPV.ModoCalculo = 1 Then
        tots = tots & "S|txtAux(4)|T|Precio|1100|;"
'    Else
'        tots = tots & "S|txtAux(6)|T|Precio|1200|;"
'    End If
    tots = tots & "S|txtAux(5)|T|Importe|1400|;N||||0|;N||||0|;"
       
    arregla tots, DataGrid1, Me
    
    'cantidad
    Me.DataGrid1.Columns(8).Alignment = dbgRight
'    Me.DataGrid1.Columns(8).NumberFormat = FormatoCantidad & " "
    
    'Precio
    Me.DataGrid1.Columns(9).Alignment = dbgRight
'    Me.DataGrid1.Columns(9).NumberFormat = FormatoPrecio
    
    'Importe
    Me.DataGrid1.Columns(10).Alignment = dbgRight
'    Me.DataGrid1.Columns(10).NumberFormat = FormatoImporte
   
   
   
   'Igual seria mejor meterlo en el arreglatots, pero como no lo domino
   Me.DataGrid1.Columns(13).Width = 450
   
   
    vDataGrid.ScrollBars = dbgAutomatic
    vDataGrid.Enabled = (Modo <> 3)
    
    Exit Sub
    
ECargaGrid2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        If Index = 5 Then
        End If
    Else
        If Index = 1 And (KeyCode = 37 Or KeyCode = 38) Then Exit Sub
        If Index = 4 And KeyCode = 39 Then Exit Sub
        KEYdownLineas KeyCode
    End If
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
Dim cant As String

    If KeyAscii = vbKeyTab Then
        If Index = 4 Then
        End If
    End If
    If KeyAscii = 42 Then '*
        cant = txtAux(Index).Text
        If EsNumerico(cant) Then
            txtAux(3).Text = cant
            PonerFormatoDecimal txtAux(3), 1
        End If
        KeyAscii = 0
        txtAux(Index).Text = ""
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 1 'Cod. ARTICULO EAN
            PonerArticuloEAN (txtAux(Index).Text)
                
        Case 2 'Cod. Articulo
           PonerArticuloCod (txtAux(Index).Text)
        
        Case 3 'CANTIDAD
          If PonerFormatoDecimal(txtAux(Index), 1) Then
                'txtAux(5).Text = ObtenerImporteLin2
                'PonerFormatoDecimal txtAux(5), 1
                FijarImportes False
          End If
            
        Case 4 'Precio
             If txtAux(Index).Text <> "" And txtAux(Index).Locked = False Then
                 'Tipo 2: Decimal(10,4)
                If PonerFormatoDecimal(txtAux(Index), 2) Then
                
                    
                
                   
                    
                    'si es articulo de varios y he modificado el precio del articulo
                    'el precio sin IVA del articulo habra que recalcularlo
                    txtAux(10).Text = ObtenerPrecioSinIVAvarios(txtAux(2).Text, txtAux(Index).Text)
                    
                     'txtAux(5).Text = ObtenerImporteLin2
                    'PonerFormatoDecimal txtAux(5), 1
                    FijarImportes True
                    
                    
                    
                    If Screen.ActiveControl.Name <> "txtAux" Then
                        GuardarLinea
                    ElseIf Screen.ActiveControl.Index > 4 Then
                        GuardarLinea
                    End If
                    'GuardarLinea
                End If
            End If
    End Select
    
'     If (Index = 3 Or Index = 4) Then  'Cant., Precio
'        If txtAux(1).Text = "" Then Exit Sub
''        txtAux(5).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'        PonerFormatoDecimal txtAux(5), 1
'    End If
End Sub




Private Function Eliminar() As Boolean
Dim b As Boolean
Dim MenError As String
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    vWhere = ObtenerWhereCP(False)
    
    MenError = "Eliminando tablas de venta."
    SQL = "DELETE FROM " & NomTablaLineas & " WHERE " & Replace(vWhere, NombreTabla, NomTablaLineas)
    Conn.Execute SQL
            
    SQL = "DELETE FROM " & NombreTabla & " WHERE " & vWhere
    Conn.Execute SQL
    
            
    'Devolvemos contador, si no estamos actualizando
    b = vParamTPV.DevolverContador(NumTermi, Data1.Recordset!NumVenta)
        
FinEliminar:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, MenError, Err.Description
    End If
    If Not b Then
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function PosicionarData() As Boolean
Dim vWhere As String

    PosicionarData = False
    
    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
'        vWhere = "numtermi= " & NumTermi & " and "
        vWhere = "numventa= " & Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
         If SituarDataMULTI(Data1, vWhere, "") Then
            PosicionarData = True
        Else
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
'        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
'        PonerCadenaBusqueda
    End If
End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    If Me.Data1.Recordset.EOF Then Exit Function
    
    SQL = NombreTabla & ".numtermi= " & Data1.Recordset!NumTermi & " and "
    SQL = SQL & NombreTabla & ".numventa= " & Data1.Recordset!NumVenta & " and " 'Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
    SQL = SQL & NombreTabla & ".fecventa= " & DBSet(Data1.Recordset!fecventa, "F")
    
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


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
    
    SQL = "SELECT numtermi, numventa,fecventa,horventa, numlinea, codigoea,codartic, nomartic, cantidad, precioiv, importel,precioar,codigiva "
    'NUEVO
    SQL = SQL & " ,if((dto1+dto2)>0,""*"","""") as dto"
    SQL = SQL & " FROM " & NomTablaLineas
    
    If enlaza Then
        SQL = SQL & " WHERE numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta 'Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
        SQL = SQL & " AND fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
    Else
        SQL = SQL & " WHERE numventa = -1 "
    End If
    SQL = SQL & Ordenacion & ",numlinea"  '" Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = SQL
End Function



   
Private Function EliminarLinea() As Boolean
Dim b As Boolean

    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
    
    'Actualizar el total de la cabecera
'    Me.Data2.Recordset.Delete
    
    'Inicilizar la clase para Actualizar los stocks
'    Set vCStock = New CStock
'    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    On Error GoTo EEliminarLinea
    Conn.BeginTrans
    
    Conn.Execute SQL 'Eliminar linea
'    b = vCStock.DevolverStock
'    Set vCStock = Nothing

    Me.label1(6).Caption = Format(ObtenerImporteTotal(True, b), FormatoImporte)

    b = Not b
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    EliminarLinea = b
End Function




Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
    Else
        LimpiarDataGrids
    End If
End Sub


Private Sub PonerArticuloEAN(codArt As String)
    codArt = Trim(codArt)
    If codArt <> "" Then
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "codartic", "codigoea", codArt, "T")
        If SQL <> "" Then
            PonerArticuloCod (SQL)
        Else
            MsgBox "No existe un artículo asociado al cód. EAN: " & codArt, vbInformation
            'Si el campo de codigo no esta vacio
            txtAux(1).Text = ""
            If txtAux(2).Text <> "" Then PonerFoco txtAux(2)
            
        End If
    End If
End Sub



Private Sub PonerArticuloCod(codArt As String)
Dim vArtic As CArticulo
Dim Dto1 As String
Dim Dto2 As String
Dim nCantidad As String
Dim nUnids As String


    On Error GoTo EArticulo
    
    codArt = Trim(codArt)
    
    If codArt <> "" Then
        Set vArtic = New CArticulo
        If vArtic.Existe(codArt) Then
            If vArtic.LeerDatos(codArt) Then
                txtAux(2).Text = vArtic.Codigo
                txtAux2(2).Text = UCase(vArtic.Nombre)
                If vArtic.TextoVentas <> "" Then vArtic.MostrarTextoVen
                Me.label1(7).Caption = vArtic.Nombre
                If txtAux(3).Text = "" Then 'cantidad
                    txtAux(3).Text = "1"
                    PonerFormatoDecimal txtAux(3), 1
                End If
                
                'Fijar tipo IVA de aritculo
                txtAux(7).Text = vArtic.TipoIVA
                txtAux(11).Text = vArtic.ObtenerPorceIVA
                
                nCantidad = txtAux(3).Text
                'Precio sin dtos
                txtAux(10).Text = vArtic.ObtPrecioParaCliente2(Text1(0).Text, nCantidad, CStr(Data1.Recordset!fecventa), Dto1, Dto2)
                
                                
                If CCur(nCantidad) <> CCur(txtAux(3).Text) Then
                'se puede vender por cajas y se insertan 2 lineas
                    nUnids = CCur(txtAux(3).Text) - CCur(nCantidad)
                    txtAux(3).Text = nCantidad
                    PonerFormatoDecimal txtAux(3), 1
                Else
                    nUnids = ""
                End If
                
                
                'Estos dos NO son visibles
                txtAux(8).Text = Dto1
                txtAux(9).Text = Dto2
                    
                If Val(Dto1) + Val(Dto2) > 0 Then txtAux(12).Text = "*"
                            
                
                'Fijar importes
                FijarImportes False
                

                

                

                Me.label1(8).Caption = txtAux(6).Text
                
                'recalculamos el importe total de la venta
                Me.label1(6).Caption = Format(ObtenerImporteTotal(False) + txtAux(5).Text, FormatoImporte)
                
               'Mostramos por el visor
               EnviarVisorPuerto Me.label1(7).Caption, label1(8).Caption, "TOTAL", label1(6).Caption
                
                
               'si es de varios el precio se puede modificar y lo desbloqueamos
               txtAux(4).Enabled = (vArtic.EsDeVarios = 1)
               txtAux(4).Locked = Not (vArtic.EsDeVarios = 1)
               If vArtic.EsDeVarios <> 0 Then
                    PonerFoco txtAux(4)
               Else
                    GuardarLinea
                    If nUnids <> "" Then
                        txtAux(3).Text = nUnids
                        PonerFormatoDecimal txtAux(3), 1
                        PonerArticuloCod codArt
                    End If
               End If
            Else
                MsgBox "No  se pudo leer el artículo", vbInformation
            End If
            
        Else
            'DE LEER EL ARTICULO. Ha dado error
            If txtAux(1).Text <> "" Then
                PonerFoco txtAux(1)
            Else
                txtAux2(2).Text = ""
            End If
        End If
        Set vArtic = Nothing
    Else
        txtAux2(2).Text = ""
    End If

EArticulo:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Poner datos del Artículo.", Err.Description
        Set vArtic = Nothing
    End If
End Sub




Private Sub LLamaLineas(alto As Single, xModo As Byte)
'Pone posicion TOP y LEFT de los controles en el form
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el alto
    For i = 0 To txtAux.Count - 1
        If ModificaLineas = 1 Then txtAux(i).Text = ""
        txtAux(i).Top = alto
        txtAux(i).Height = Me.DataGrid1.RowHeight - 10
    Next i
    If ModificaLineas = 1 Then txtAux2(2).Text = ""
    txtAux2(2).Top = alto
    txtAux2(2).Height = Me.DataGrid1.RowHeight - 10
    txtAux2(2).Enabled = False
    
    'El precio siempre bloqueado a no ser que sea articulo de varios
    txtAux(4).Locked = True
    txtAux(4).Enabled = False
    
    'El importe siempre bloqueado
    txtAux(5).Locked = True
    txtAux(5).Enabled = False
    
    
    'Pongo el del DTO ajustadito al precio
    txtAux(12).Left = txtAux(5).Left + txtAux(5).Width + 30
End Sub


Private Sub BotonAnyadir()
    On Error GoTo EVenta
                
    ReiniciarVisor
    If InsertarVenta Then
        PonerVentasAbiertas
    
        'Cargar Data 1 con las ventas abiertas (cabeceras)
        cargaCabeceras True
        
        'Cargar las ventas que hay abiertas en el combo
        cargaComboVentas
    
        'Nos situamos en la ultima venta
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    
        'Ponemos el importe total de la venta
'        If Not Data1.Recordset.EOF Then
'            SQL = DevuelveDesdeBDNew(conAri, "scaven", "imptotal", "numtermi", CStr(NumTermi), "N", , "numventa", CStr(Data1.Recordset!numventa), "N", "fecventa", CStr(Data1.Recordset!fecventa), "F")
'            Me.Label1(6).Caption = Format(CCur(SQL), FormatoImporte)
'        End If
        PonerModo 3
        DoEvents
        If vParamTPV.Rapida Then
            'Entrada rapidad  DAVID
            PonerFoco txtAux(1) ' Nos situamos automáticamente en el campo de la linea
        Else
            PonerFoco Text1(0) ' (RAFA/ALZIRA 31082006) Nos situamos automáticamente en el campo de cliente
        End If
    End If
    
EVenta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Venta.", Err.Description
    End If
End Sub




Private Sub BotonAnyadirLinea()
'Añade una nueva linea de Venta
Dim anc As Single

    On Error GoTo EAnyadirLinea
        
    ModificaLineas = 1
    
    AnyadirLinea DataGrid1, Data2
    
    
    If DataGrid1.Row < 0 Then
        anc = ObtenerAlto(DataGrid1, 50)
    Else
        anc = ObtenerAlto(DataGrid1, 20)
    End If
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(1)

EAnyadirLinea:
    If Err.Number > 0 Then MuestraError Err.Number, "Añadir linea.", Err.Description
End Sub




Private Function ObtenerImporteLin() As Currency
    'MODIFICADO.      Antes:  ObtenerImporteLin2(d1 As String, d2 As String
    'David  14  Mayo 08
    '
    'Antes pasabamos el descuento. Ahora debe estar en el campo txtaux 8 y 9
    
    ObtenerImporteLin = CalcularImporte4(txtAux(3).Text, txtAux(10).Text, txtAux(8).Text, txtAux(9).Text, vParamAplic.TipoDtos)
        
End Function

Private Function FijarImportes(HaCambiadoElPrecio As Boolean)
            txtAux(6).Text = ObtenerImporteLin
            'PonerFormatoDecimal txtAux(6), 1
            PonerFormatoDecimal txtAux(6), 2   '4 digitos
            
            SQL = CalcularDto(txtAux(6).Text, CStr(txtAux(11).Text))
            txtAux(5).Text = Round(CCur(ComprobarCero(txtAux(6).Text)) + CCur(ComprobarCero(SQL)), 4)
            PonerFormatoDecimal txtAux(5), 5
            
            
            If Not HaCambiadoElPrecio Then
                txtAux(4).Text = Round2(CCur(txtAux(5).Text) / CCur(txtAux(3).Text), 4)
                PonerFormatoDecimal txtAux(4), 2
            End If
            
                
           
                
End Function

Private Sub PosicionarComboVentas(venta As Long)

    For i = 0 To Me.cboNumVenta.ListCount - 1
        If CLng(Trim(Me.cboNumVenta.List(i))) = venta Then
            Me.cboNumVenta.ListIndex = i
            Exit For
        End If
    Next i
    
End Sub


Private Function ObtenerImporteTotal(Optional actualiza As Boolean, Optional Error As Boolean) As Currency
'Suma el total de las lineas y actualiza la tabla scaven con el valor correcto
Dim Rs As ADODB.Recordset
Dim total As Currency
        
    On Error GoTo ETotal
'    If Data2.Recordset.EOF Then Exit Function
    
'    SQL = Me.Data2.RecordSource
'    If SQL <> "" Then
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        total = 0
'        While Not RS.EOF
'            total = total + RS!ImporteL
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'    End If

    total = 0
    SQL = "select sum(importel) FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        total = DBLet(Rs.Fields(0).Value, "D")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    ObtenerImporteTotal = total
        
    If actualiza Then
        SQL = "UPDATE " & NombreTabla & " SET imptotal=" & DBSet(total, "N")
        SQL = SQL & ObtenerWhereCP(True)
        Conn.Execute SQL
    End If
    
ETotal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando Importe Total.", Err.Description
        Error = True
    Else
        Error = False
    End If
End Function



Private Function InsertarLinea() As Boolean
'Inserta una linea de venta en la tabla: sliven
Dim b As Boolean
    
    If Not DatosOkLinea Then Exit Function
    
    On Error GoTo EInsLinea
    Conn.BeginTrans
    'Anyadimos nuevo registro y rellenamos los campos clave (ocultos)
    'El campo de numventa no tiene valor es que aun se ha insertado la linea
    If Me.txtAux(0).Text = "" Then
        'numero de linea de la venta
        SQL = "numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta & " and fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
        SQL = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", SQL)
        txtAux(0).Text = SQL
        
        SQL = "INSERT INTO sliven(numtermi,numventa,fecventa,horventa,numlinea,codigoea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal) "
        SQL = SQL & " VALUES (" & Data1.Recordset!NumTermi & "," & Data1.Recordset!NumVenta & "," & DBSet(Data1.Recordset!fecventa, "F") & ","
        '& DBSet(Data1.Recordset!fecventa & " " & Format(Now, "hh:mm:ss"), "FH") & ","
        SQL = SQL & DBSet(Now, "FH") & ","
        SQL = SQL & txtAux(0).Text & "," & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T") & ","
        
        'DAVID.
        '   Antes pasaba el 6. Ahora paso el 10.  EL precio articulo esta en el 10
        '                  cantidad                        precioiv                             importe                             precioar
        SQL = SQL & DBSet(txtAux(3).Text, "N") & "," & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(10).Text, "N") & "," & txtAux(7).Text
        
        
        'Nuevo###       David
        'Para llevar so tiene descuentos en linea los guarda implineareal
        '                             dto1                              dto2                                 imp linea real
        SQL = SQL & "," & DBSet(txtAux(8).Text, "N", "N") & "," & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(6).Text, "N", "N") & ")"
        Conn.Execute SQL
    End If
    
    Me.label1(6).Caption = Format(ObtenerImporteTotal(True), FormatoImporte)
    b = True

EInsLinea:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Insertando linea.", Err.Description
    End If
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    InsertarLinea = b
End Function


Private Sub ReiniciarVisor()
    Me.label1(6).Caption = "0,00"
    Me.label1(7).Caption = "Precio"
    Me.label1(8).Caption = "0,00"
    If vParamTPV.HayVisor Then
        EnviarVisorPuerto label1(7).Caption, label1(8).Caption, label1(5).Caption, label1(6).Caption
    End If
End Sub


Public Sub EnviarVisorPuerto(lin1 As String, imp1 As String, lin2 As String, imp2 As String)
Dim Buffer1 As Variant
Dim Buffer2 As Variant

    If Me.MSComm1.PortOpen Then
        If Len(imp1) >= 8 Then
            Buffer1 = Left(Mid(lin1, 1, 11) & Space(20), 11) & Right(Space(20) & Mid(imp1, 1, 9), 9)
        Else
            Buffer1 = Left(Mid(lin1, 1, 13) & Space(20), 13) & Right(Space(20) & Mid(imp1, 1, 7), 7)
        End If
        Buffer2 = Left(Mid(lin2, 1, 7) & Space(20), 7) & Right(Space(20) & Mid(imp2, 1, 13), 13)
        Me.MSComm1.Output = Buffer1 & Buffer2
    End If
End Sub



Private Function InsertarVenta() As Boolean
'Inserta en tabla temporal la cabecera de una nueva venta
Dim cont As Long
Dim b As Boolean
Dim MenError As String

    On Error GoTo EInsVenta
    
    Conn.BeginTrans
    
    cont = vParamTPV.ConseguirContador(NumTermi)
    If Not IsNumeric(cont) Then
        MenError = "No se ha podido obtener nº de contador."
        b = False
    Else
        'Insertamos la cabecera de venta
        MenError = "Insertando en tabla de ventas."
        SQL = "INSERT INTO " & NombreTabla & " (numtermi,numventa,fecventa,horventa,codtraba,imcambio,imptotal,codclien) VALUES "
        SQL = SQL & "(" & NumTermi & "," & cont & "," & DBSet(Now, "F") & ", " & DBSet(Now, "FH") & "," & CodTraba & "," & ValorNulo & ",0,"
        'Poner el cliente que hay por defecto en los parametros
        If vParamTPV.Cliente = "" Then
            SQL = SQL & "0)"
        Else
            SQL = SQL & vParamTPV.Cliente & ")"
        End If
        
        Conn.Execute SQL
        
        
        MenError = "Incrementando contador de venta."
        b = vParamTPV.IncrementarContador(NumTermi)
    End If
    
EInsVenta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error insertando Venta." & vbCrLf & MenError, Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    InsertarVenta = b
End Function


Private Function ModificarVenta() As Boolean
'Modificar el cliente de una venta
Dim SQL As String
    On Error GoTo ErrModVen
    
    SQL = "UPDATE scaven SET "
    If Trim(Text1(0).Text) <> "" Then SQL = SQL & " codclien=" & DBSet(Text1(0).Text, "N") & ","
    SQL = SQL & "coddirec=" & DBSet(Text1(2).Text, "N", "S")
'    If CDate(Data1.Recordset!fecventa) <> CDate(Format(Now, "dd/mm/yyyy")) Then
''        SQL = SQL & ", fecventa=" & DBSet(Now, "F") '& ", horventa=" & DBSet(Now, "FH")
'        SQL = SQL & ", horventa=" & DBSet(Now, "FH")
'    End If
'    SQL = SQL & ", horventa=" & DBSet(Now, "FH")
    SQL = SQL & ", observa1=" & DBSet(Text1(1).Text, "T")
            
    SQL = SQL & " WHERE numtermi=" & Me.Data1.Recordset!NumTermi & " AND numventa=" & Me.Data1.Recordset!NumVenta
    SQL = SQL & " AND fecventa=" & DBSet(Me.Data1.Recordset!fecventa, "F")
    Conn.Execute SQL
    
    Data1.Refresh
    PosicionarData
    
    Exit Function
    
ErrModVen:
    ModificarVenta = False
    MuestraError Err.Number, "Modificar cabecera de la venta.", Err.Description
End Function


Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux2(Index)
End Sub



Private Sub GuardarLinea()
Dim PrimeraLin As Boolean

    If Data2.Recordset.EOF = True Then PrimeraLin = True
     If InsertarLinea Then
        DataGrid1.AllowAddNew = False
'        PonerModo 2
        
        If PrimeraLin Then
            CargaGrid DataGrid1, Data2, True
        Else
            CargaGrid2 DataGrid1, Data2
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            BotonAnyadirLinea
        Else
            PonerModo 2
        End If

        
    Else
        txtAux(0).Text = "" 'Limpiamos el num linea para volver a insertar tras corregir datos
    End If
End Sub



Private Function AbrirVisorPuerto() As Boolean
On Error GoTo EAbrirVisor

    'Establecemos el puerto de comunicaciones
    If vParamTPV.HayVisor Then
        Me.MSComm1.CommPort = vParamTPV.NumPuerto
        
        ' 9600 baudios, sin paridad, 8 bits de datos y 1 bit de parada.
        Me.MSComm1.Settings = vParamTPV.VelociPuerto & ",N,8,1"
        
        ' Indicar al control que lea todo el búfer al usar Input.
        MSComm1.InputLen = 0

        'Abrimos el puerto
        Me.MSComm1.PortOpen = True
                
        Me.MSComm1.Output = Mid(vEmpresa.nomempre & Space(40), 1, 40)
    End If
    
EAbrirVisor:
    If Err.Number <> 0 Then MuestraError Err.Number, "Abrir visor.", Err.Description
End Function




Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(0).Text
    
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(2).Text, NomDpto) Then
        Text2(2).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function



Private Function ClienteOK(newCli As String, oldCli As String, modifica As Boolean, Optional mostrarObs As Boolean) As Boolean
'(IN) newCli: cliente nuevo q queremos poner
'(IN) oldCli: cliente guardado actualmente si existe
Dim cCli As CCliente
Dim devuelve As String

    On Error GoTo ErrCliOK
    ClienteOK = False
    
    If newCli <> "" Then newCli = CStr(Val(newCli))
    Set cCli = New CCliente
    If cCli.LeerDatos(newCli) Then
        '-- Si el cliente esta bloqueado no permitimos este cliente para la venta
        If cCli.ClienteBloqueado Then
            Set cCli = Nothing
            Exit Function
        End If
        
        '-- si se ha modificado el cliente y si hay lineas de articulos:
        '   comprobar q el nuevo cliente tiene la misma tarifa q el cliente anterior
        '   sino no permitimos el nuevo cliente para la venta
        If (oldCli <> "") And (newCli <> oldCli) Then
            If Not Me.Data2.Recordset.EOF Then 'si hay lineas
                'obtener la tarifa del cliente actual
                devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", oldCli, "N")
                If devuelve <> CStr(cCli.Tarifa) Then
                    devuelve = "No se puede seleccionar el cliente " & newCli & " "
                    devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
                    devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
                    MsgBox devuelve, vbExclamation, "Comprobar cliente"
                    Set cCli = Nothing
                    Exit Function
                Else
                    modifica = True
                End If
            Else
                modifica = True
            End If
        End If
        ClienteOK = True
        
        'mostrar las observaciones del cliente
        If mostrarObs Then cCli.MostrarObservaciones
    End If
    
    Set cCli = Nothing
    Exit Function
    
ErrCliOK:
    MuestraError Err.Number, "Comprobar cliente correcto.", Err.Description
End Function



Private Sub RevisarEntradasDia()

    CadSelVenta = ""
    Set frmTraerVen = New frmFacTPVTraerVen
    frmTraerVen.parNumTermi = -1   'El menos 1 siginifac que me mostrara las ventas del dia
    frmTraerVen.Show vbModal
    Set frmTraerVen = Nothing
    If CadSelVenta <> "" Then
        
        i = RecuperaValor(CadSelVenta, 1)

        If i = 1 Then
        
            SQL = RecuperaValor(CadSelVenta, 2)
            CadSelVenta = RecuperaValor(CadSelVenta, 3)
        
            'IMPRESION
            If SQL = "FTI" Then
                 ImprimirTicketDirecto CadSelVenta, Now
            Else
                'Es una factura de venta
                ImprimirFAV
            End If
            
        Else
            SQL = RecuperaValor(CadSelVenta, 5)   'Tipo movimiento albaran
            'Ponemos el numero de albaran, NO el numero de factura
            CadSelVenta = RecuperaValor(CadSelVenta, 4)
            'Ver lineas detalles
            '----------------------------------
            With frmFacHcoFacturas
                 .hcoCodMovim = CadSelVenta
                 .hcoCodTipoM = SQL
                 .hcoFechaMov = Now
                 .Show vbModal
            End With
            
            
            
            
            
        End If
    End If
    CadSelVenta = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub ImprimirFAV()
Dim cadParam As String
Dim numParam As Byte
Dim nomDocu As String

'    '===================================================
'    '============ PARAMETROS ===========================
'    If FormatoFacturaTPV Then
'        indRPT = 18 'FACTURAS TPV
'    Else
'        indRPT = 12 'Facturas Clientes
'    End If
    If Not PonerParamRPT(18, cadParam, numParam, nomDocu) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu


    'EN sql tengo el numero de factura
    SQL = "({scafac.numfactu} = " & CadSelVenta & ")"
    
    'Tipo documento
    SQL = SQL & " AND ({scafac.codtipom}='FAV') "
    
    'fecha factu
    SQL = SQL & " AND ({scafac.fecfactu} =  Date(" & Year(Now) & "," & Month(Now) & "," & Day(Now) & ")" & ")"

    CadSelVenta = SQL
    'If Not HayRegParaInforme("scafac", CadSelVenta) Then Exit Sub

    With frmImprimir
            .FormulaSeleccion = SQL
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 53
            .Titulo = ""
            .Show vbModal
    End With
End Sub
