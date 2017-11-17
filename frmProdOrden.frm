VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProdOrden 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de produccion"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   9750
   Icon            =   "frmProdOrden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   5520
      MaxLength       =   16
      TabIndex        =   8
      Tag             =   "Lote"
      Text            =   "lot"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtComponentes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   6360
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
      Height          =   1455
      Left            =   1920
      TabIndex        =   28
      Top             =   6360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2566
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
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   3120
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
      TabIndex        =   7
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   7200
      MaxLength       =   16
      TabIndex        =   9
      Tag             =   "Cantidad"
      Text            =   "1,234,567,891.25"
      Top             =   3180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   25
      Tag             =   "Nombre Artículo"
      Text            =   "nomArtic"
      Top             =   3180
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   24
      ToolTipText     =   "Buscar almacen"
      Top             =   3180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   23
      ToolTipText     =   "Buscar artículo"
      Top             =   3180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   410
      Width           =   9375
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Fecha caducidad|F|S|||sordprod|feccaduca|dd/mm/yyyy|N|"
         Top             =   1080
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Nº Pedido|N|S|||sordprod|numpedcl|00000000|N|"
         Top             =   1080
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   1155
         Index           =   3
         Left            =   4680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "Obs|T|S|||sordprod|descripcion|||"
         Top             =   360
         Width           =   4545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha creación|F|N|||sordprod|feccreacion|dd/mm/yyyy|N|"
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
         Tag             =   "Nº ord produccion|N|S|0||sordprod|codigo|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha producción|F|S|||sordprod|fecproduccion|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2640
         Picture         =   "frmProdOrden.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. caducidad"
         Height          =   255
         Index           =   2
         Left            =   1590
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmProdOrden.frx":0097
         ToolTipText     =   "Buscar Nº Serie"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   26
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "F. creacion"
         Height          =   255
         Index           =   14
         Left            =   1590
         TabIndex        =   21
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2520
         Picture         =   "frmProdOrden.frx":0199
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   20
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmProdOrden.frx":0224
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. produccion"
         Height          =   255
         Index           =   51
         Left            =   3120
         TabIndex        =   19
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   7935
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8370
      TabIndex        =   12
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   8040
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
      TabIndex        =   16
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
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
            Object.ToolTipText     =   "Lineas produccion"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar cantidad en componentes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar orden produccion"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir con nºlote"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado produccion"
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
         TabIndex        =   17
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
      Left            =   8370
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProdOrden.frx":02AF
      Height          =   3720
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6562
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
      Caption         =   "Componentes"
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
      Left            =   120
      TabIndex        =   29
      Top             =   6360
      Width           =   1530
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   600
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Articulos producción"
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
      TabIndex        =   30
      Top             =   2040
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmProdOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const LimiteB = 100000

Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1


Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmPe As frmFacEntPedidos
Attribute frmPe.VB_VarHelpID = -1


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
Private Kcampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1






Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim OpcionConElPedido As Byte
    ' 0. NADA
    ' >1 traer los datos del pedido
    '   =1 AÑAIDR LOS DATOS
    '   =2 borrar los anteriores
    
    
    
Dim TrabajadorConectado As Integer

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
                    ActualizarLineasPedido
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    
                    'Llamo a los lotes de la cabecera
                    LlamaLotes
                    
                    'AHora despues Situaremos en las lineas
                    CargaGrid3 True
                    HacerLlamadaALotesLinea
                    'Para que quite el modo editar linea
                    cmdCancelar_Click
                    PonerModo 2
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
            
                'Para saber si he cambiado cantidad
                If ModificarLinea Then
                    PrimeraLin = CCur(txtAux(4).Text) <> Data2.Recordset!Cantidad
                    
           
                
                    TerminaBloquear
                    CargaTxtAux False, False
                    
                    'PrimeraLin
                    If PrimeraLin Then LlamaLotes
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True

                End If
                Me.DataGrid1.Enabled = True
            End If
            
            
        Case 6 'Modif cantidad componentes
            
            If DatosOkLineaCompo Then
                UpdateaCantidadComponentes
                CargaGrid3 True
                Me.DataGrid2.Enabled = True
                ModificaLineas = 0
            End If
            
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
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
            ModificaLineas = 0
            Me.txtComponentes.visible = False
            CargaGrid3 True
            DataGrid2.Enabled = True
            PonerBotonCabecera True
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

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    CargaGrid3 False
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", CStr(TrabajadorConectado), "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")

    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
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
            Text1(Kcampo).Text = ""
            Text1(Kcampo).BackColor = vbYellow
            PonerFoco Text1(Kcampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim C As String
'    LimpiarCampos
    C = "<"
    If vUsu.TrabajadorB Then C = ">"
    C = " codigo " & C & LimiteB
        
        

    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia C
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & " WHERE " & C & Ordenacion
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
    If Data2.Recordset.EOF Then Exit Sub
    
  '  vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
  '  vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
  '  If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False

    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
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

    If Not IsNull(Data1.Recordset!fecproduccion) Then
        MsgBox "Orden cerrada. No se puede eliminar", vbExclamation
        Exit Sub
    End If

    Cad = "Produccion." & vbCrLf
    Cad = Cad & "----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la orden de produccion:"
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    
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

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de produccion?     "
    SQL = SQL & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!Codigo
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        
        'Los lotes
        conn.Execute "DELETE FROM sliordprlotes " & SQL
        conn.Execute "DELETE FROM sliordpr2lotes " & SQL
        'Las sublineas
        conn.Execute "DELETE FROM sliordpr2 " & SQL
        'Las lineas
        conn.Execute "DELETE FROM sliordpr " & SQL
        
        
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
'        SituarDataTrasEliminar Data2, NumRegElim
        SituarDataPosicion Me.Data2, NumRegElim, SQL
        CargaGrid3 Not Data2.Recordset.EOF
        
'        CancelaADODC
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
        HabilitarModifCantidad False
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdCancelar.Cancel = True
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
            
            LlamaLotes
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






Private Sub DataGrid2_DblClick()

    If Modo = 6 Then  'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not data3.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Solo para las materias primas
            If data3.Recordset!FactorConversion <> 1 Then
                LlamaLotesLin
            Else
                If DevuelveDesdeBD(conAri, "trazabilidad", "sartic", "codartic", data3.Recordset!codarti2, "T") = "1" Then LlamaLotesLin
            End If
        Else
            
        End If
    End If
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

      
    LimpiarCampos   'Limpia los campos TextBox
   
   
    Ordenacion = PonerTrabajadorConectado(NombreTabla)
    If Ordenacion <> "" Then
        TrabajadorConectado = CInt(Ordenacion)
    Else
        TrabajadorConectado = -1
    End If
        

    NombreTabla = "sordprod"
    Ordenacion = " ORDER BY codigo "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    
    CadenaConsulta = "Select * from " & NombreTabla & " where codigo= "
    If DatosADevolverBusqueda2 = "" Then
        CadenaConsulta = CadenaConsulta & "-1"
        Label4.visible = vUsu.TrabajadorB
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

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub




Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
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

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmPe = New frmFacEntPedidos
    frmPe.DatosADevolverBusqueda2 = "0"
    frmPe.Show vbModal
    Set frmPe = Nothing

    
    
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


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub










Private Sub mnLineas_Click()
    BotonMtoLineas
End Sub


Private Sub mnModificar_Click()
        
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    ElseIf Modo = 6 Then
        ModificarCantidadComponentes
    Else
        'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
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
    Kcampo = Index
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
        Case 1, 2, 5 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)
            
            If Index = 1 And Text1(2).Text <> "" Then 'Fecha Entrega
                'Comprobar que es posterior a la del pedido
                If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, False, "") Then
                    Devuelve = "La Fecha de produccion debería ser posterior a la fecha de creacion." & vbCrLf & "¿Continuar?"
                    If MsgBox(Devuelve, vbQuestion + vbYesNo) = vbNo Then
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                        Exit Sub
                    End If
                End If
               
            End If
            
    
        Case 4 '
            If PonerFormatoEntero(Text1(Index)) Then

            Else
               
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(6).Text = Data1.Recordset!nifClien Then
'                    Exit Sub
'                End If
'            End If
'            PonerDatosClienteVario (Text1(Index).Text)
             
        Case 9 'Cod. Postal

            
 
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then cadB = cadB & " AND codigo "
    If vUsu.TrabajadorB Then
        cadB = cadB & " > "
    Else
        cadB = cadB & " < "
    End If
    cadB = cadB & LimiteB
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
        Cad = Cad & ParaGrid(Text1(0), 20, "Nº Orden")
        Cad = Cad & ParaGrid(Text1(1), 20, "Fecha creación")
        Cad = Cad & ParaGrid(Text1(2), 20, "Fecha producción")
        Tabla = NombreTabla
      
        Titulo = "Ordenes producción"
        Devuelve = "0|"

    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        Tabla = "sdirec"
        Devuelve = "0|1|"
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
'        frmB.vBuscaPrevia = chkVistaPrevia
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
          
            PonerFoco Text1(Kcampo)
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
    

       
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    

    
    Label4.visible = Val(Text1(0).Text) > LimiteB
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
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
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        

    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(2), b
    b = Modo = 0 Or Modo = 2 Or Modo >= 5
    BloquearTxt Text1(1), b
    BloquearTxt Text1(3), b
    BloquearTxt Text1(4), b
    BloquearTxt Text1(5), b
  
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
  
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    'Las imagenes añadimos el modo 6
    b = b And Modo <> 6
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    imgBuscar(0).visible = b


    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    'Solo en modificamos cantidad en modo6
    txtComponentes.visible = False
    
    
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
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, False) Then
        Devuelve = "La Fecha de produccion debería ser posterior a la fecha de creacion." & vbCrLf & "¿Continuar?"
        If MsgBox(Devuelve, vbQuestion + vbYesNo) = vbNo Then
            PonerFoco Text1(1)
            Exit Function
        End If
    End If
    
          
     'Si ha puesto el numero de pedido entonces
     'deberemos traer los datos
     OpcionConElPedido = 0
     If Text1(4).Text <> "" Then
     
        Devuelve = DevuelveDesdeBD(conAri, "numpedcl", "scaped", "numpedcl", Text1(4).Text)
        If Devuelve = "" Then
            MsgBox "No existe el pedido: " & Text1(4).Text, vbExclamation
            Exit Function
        End If
        If Modo = 3 Then
            OpcionConElPedido = 1 'INSERTAMOS Y A CORRER
        Else
            'Modificar. Si ya tenia datos entonces puede ser que quiera eliminar los datos anteriores
            'Si tenia pedido o no
            If Val(Text1(4).Text) <> DBLet(Data1.Recordset!numpedcl, "N") Then
                If Not Data2.Recordset.EOF Then
                    Devuelve = "Se van a insertar las lineas del pedido: " & Text1(4).Text
                    Devuelve = Devuelve & vbCrLf & "¿Desea eliminar las lineas anteriores?"
                    NumRegElim = Val(MsgBox(Devuelve, vbQuestion + vbYesNoCancel))
                    If CByte(NumRegElim) = vbCancel Then Exit Function
                    If CByte(NumRegElim) = vbYes Then
                        OpcionConElPedido = 2
                    Else
                        OpcionConElPedido = 1
                    End If
                Else
                    'EOF. insertamos
                    OpcionConElPedido = 1
                End If
            End If
        End If
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
Dim I As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For I = 0 To txtAux.Count - 1
        If txtAux(I).Text = "" And I <> 3 Then
            MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(I)
            Exit Function
        End If
    Next I
     
    If Not vUsu.TrabajadorB Then
        If Val(txtAux(0).Text) = vParamAplic.AlmacenB Then
            MsgBox "Almacen incorrecto(2)", vbExclamation
            Exit Function
        End If
    End If
    
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLineaCompo() As Boolean
    DatosOkLineaCompo = False
    
    If Me.txtComponentes.Text = "" Then
        MsgBox "Escriba la cantidad para el componente", vbExclamation
        Exit Function
    End If
    
    If Not IsNumeric(txtComponentes.Text) Then
        MsgBox "Cantidad debe ser numérica", vbExclamation
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
            'Modificar cantidad de componentes
            If Not IsNull(Data1.Recordset!fecproduccion) Then
                MsgBox "Orden cerrada", vbExclamation
                Exit Sub
            End If
            
            If data3.Recordset.RecordCount = 0 Then Exit Sub
            Me.DataGrid1.Enabled = False
            Me.DataGrid2.Enabled = True
            
                ModificaLineas = 0
                PonerModo 6
                PonerBotonCabecera True
            
            

            
        Case 12, 14, 15
            'IMPRIMIR (14,15)    y cerrar(12) orden produccion
            '--------------------------------------------------------------------
            
            If Modo <> 2 Then Exit Sub
            
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una orden de produccion", vbExclamation
                Exit Sub
            End If
            If Button.Index = 12 Then
            
                If Not IsNull(Data1.Recordset!fecproduccion) Then
                    MsgBox "La orden de produccion ya esta cerrada", vbExclamation
                    Exit Sub
                End If
                
                If Data2.Recordset.EOF Then
                    MsgBox "No tiene lineas la orden de produccion", vbExclamation
                    Exit Sub
                End If
                
                'Una comprobacion. Ver que todas las lineas (las que se crean) tienen indicados los numeros de LOTE
                Set miRsAux = New ADODB.Recordset
                b = ComprobarNumeroLoteLineas
                Set miRsAux = Nothing
                If Not b Then Exit Sub
                
                
                
                If BLOQUEADesdeFormulario(Me) Then
                
                    frmProduVarios.Intercambio = Data1.Recordset!Codigo & "|" & Data1.Recordset!feccreacion & "|"
                    frmProduVarios.Opcion = 0
                    frmProduVarios.Show vbModal
                
                    'TErminamos de bloquear
                    TerminaBloquear
                    
                    'Refrescamos
                    CadenaConsulta = Data1.RecordSource
                    Data1.Refresh
                    'Y ponemos los campos
                    PosicionarData
                  
                    
                    
                
                End If
                
            Else
                'Imprimir orden prod
                With frmImprimir
                    .ConSubInforme = True
                    .FormulaSeleccion = "{sordprod.codigo} = " & Data1.Recordset!Codigo
                    .Titulo = "Orden de produccion"
                    If Button.Index = 14 Then
                        .NombreRPT = "rordenproduccion.rpt" 'nromal
                    Else
                        .NombreRPT = "rordenproduccionT.rpt" 'lotes
                        .Titulo = .Titulo & " (Lotes)"
                    End If
                    .OtrosParametros = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
                    .NumeroParametros = 1
                    
                    .Opcion = 2003 'Esta libre
                    .Show vbModal
                End With
            End If
        Case 16
            'Imprimir listado produccion con UDS y LITROS
            frmListado2.Opcion = 29
            frmListado2.Show vbModal
        
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

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        SQL = "INSERT INTO sliordpr"
        SQL = SQL & "( codigo, codalmac, codartic ,cantidad,numlote ) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(4).Text, "S") & ","
        SQL = SQL & DBSet(txtAux(3).Text, "T", "S") & ")"
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        'Insertamos en lineas2
        ActualizarComponentes
        
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
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sliordpr set codalmac=" & txtAux(0).Text & " , codartic =" & DBSet(txtAux(1).Text, "T")
        SQL = SQL & ", cantidad = " & DBSet(txtAux(4).Text, "N")
        'SQL = SQL & ", numlote = " & DBSet(txtAux(3).Text, "T", "S")
        SQL = SQL & " WHERE codigo =" & Data1.Recordset!Codigo & " AND codalmac = " & Data2.Recordset!codAlmac
        SQL = SQL & " AND codartic =" & DBSet(Data2.Recordset!codartic, "T")
        
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        ActualizarComponentes
        
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

    SQL = "codigo = -1"


    If enlaza Then
       If Not Data2.Recordset.EOF Then
            SQL = " codigo = " & Data1.Recordset!Codigo
            SQL = SQL & " AND codalmac = " & Data2.Recordset!codAlmac
            SQL = SQL & " AND sliordpr2.codartic = " & DBSet(Data2.Recordset!codartic, "T")
            
       End If
    End If

    SQL = "select codarti2,nomartic,cantidad,factorconversion,trazabilidad  from sliordpr2,sartic where  sliordpr2.codarti2=sartic.codartic AND " & SQL
    data3.ConnectionString = conn
    data3.RecordSource = SQL
    data3.Refresh
    DataGrid2.AllowRowSizing = False
    If DataGrid2.DataSource Is Nothing Then DataGrid2.ClearFields
        
    Set DataGrid2.DataSource = data3
    DataGrid2.RowHeight = 290
    DataGrid2.Columns(0).Caption = "Codigo"
    DataGrid2.Columns(0).Width = 1900
    
    
    DataGrid2.Columns(1).Caption = "Articulo"
    DataGrid2.Columns(1).Width = 3700

    DataGrid2.Columns(2).Caption = "Cantidad"
    DataGrid2.Columns(2).Width = 1200
    DataGrid2.Columns(2).NumberFormat = "#,##0.00000"
    DataGrid2.Columns(2).Alignment = dbgRight
    
    'FactorConversion
    DataGrid2.Columns(3).visible = False
    DataGrid2.Columns(4).visible = False
    
End Sub



Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte

    On Error GoTo ECargaGrid

    vData.Refresh

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(0).Caption = "Alm."
                vDataGrid.Columns(0).Width = 500
                vDataGrid.Columns(0).NumberFormat = "000"
                
                vDataGrid.Columns(1).Caption = "Articulo"
                vDataGrid.Columns(1).Width = 1600

                
                vDataGrid.Columns(2).Caption = "Desc. Artículo"
                vDataGrid.Columns(2).Width = 3450

                'Numero de lote
                vDataGrid.Columns(3).Caption = "Nº Lote"
                vDataGrid.Columns(3).Width = 2150
                
                vDataGrid.Columns(4).Caption = "Cantidad"
                vDataGrid.Columns(4).Width = 1000
                vDataGrid.Columns(4).Alignment = dbgRight
                vDataGrid.Columns(4).NumberFormat = FormatoCantidad
             
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
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I).Text
                txtAux(I).Locked = False
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
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(0).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(1).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(2).Width - 10
        'Lote
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(3).Width - 10
        'Cantidad
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 10
        txtAux(4).Width = DataGrid1.Columns(4).Width - 10

        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
    'El Campo de Origen del precio se actualiza por programa al modificar el precio
    'LOTE SIEMPRE BLOQUEADO
    BloquearTxt txtAux(3), True
    
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(Kcampo, Index)
    Kcampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Modo6: Pasar de Pedido a Albaran
        If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    End If
End Sub




Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Modo 6: Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
'            PonerServidas
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As cStock
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            Devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = Devuelve
            'If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then 'Cod Artic
                txtAux(2).Text = "" 'Nom Artic
                Exit Sub
            End If
            If txtAux(0).Text = "" Then 'Cod Almacen
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If

            Devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then Devuelve = Data2.Recordset!codartic
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, "", ModificaLineas, Devuelve) Then
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            End If
            
        Case 2 'desc Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 4 'CANTIDAD
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 3) Then   'Tipo 3: FormatoCantidad
    
                Else
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
            
        
    End Select
    

End Sub


Private Sub BotonMtoLineas()
       If Not IsNull(Data1.Recordset!fecproduccion) Then
            MsgBox "Orden cerrada", vbExclamation
        Else
            ModificaLineas = 0
            PonerModo 5
            PonerBotonCabecera True
        End If
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean



    On Error GoTo FinEliminar

        conn.BeginTrans
        'Los lotes
        conn.Execute "DELETE FROM sliordprlotes where codigo =" & Text1(0).Text
        conn.Execute "Delete from sliordpr2 where codigo =" & Text1(0).Text
        conn.Execute "Delete from sliordpr where codigo =" & Text1(0).Text
        conn.Execute "Delete from sordprod where codigo =" & Text1(0).Text
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
    
    SQL = NombreTabla & ".codigo= " & Val(Text1(0).Text)
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
    
    SQL = "SELECT codalmac,sliordpr.codartic,nomartic,numlote,cantidad "
    SQL = SQL & " FROM sliordpr,sartic WHERE sliordpr.codartic=sartic.codartic AND "
    If enlaza Then
        SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, "sliordpr")
    Else
        SQL = SQL & " codigo = -1"
    End If
    SQL = SQL & " Order by codigo"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b Or (Modo = 6 And ModificaLineas = 0)
        Me.mnModificar.Enabled = b Or (Modo = 6 And ModificaLineas = 0)
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b

        
        Toolbar1.Buttons(13).Enabled = b And vParamAplic.ProduccionNueva
        
        
        
      
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
    CadenaConsulta = "<"
    If vUsu.TrabajadorB Then CadenaConsulta = ">"
    CadenaConsulta = "codigo " & CadenaConsulta & LimiteB
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codigo", CadenaConsulta)
    If vUsu.TrabajadorB Then
        If Text1(0).Text = "1" Then Text1(0).Text = LimiteB + 1
    End If
    
    If InsertarDesdeForm(Me) Then
    
            ActualizarLineasPedido
    
            'Si tiene pedido traeremos las lineas del pedido
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE codigo = " & Text1(0).Text & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas
            BotonAnyadirLinea
    
    Else
        CadenaConsulta = ""
    End If

End Sub

Private Sub ActualizarLineasPedido()
Dim SQL As String
Dim RT As ADODB.Recordset
Dim Cantidad As Currency

    If OpcionConElPedido = 0 Then Exit Sub
    
    'Si tiene que coger pero no tiene pedido (NO DEBERIA PASAR)
    If Text1(4).Text = "" Then Exit Sub
    
    If OpcionConElPedido = 2 Then
        'Eliminamos los que hubieren
        SQL = "DELETE FROM sliordpr where codigo = " & Text1(0).Text
        conn.Execute SQL
    End If
    

    SQL = "INSERT INTO sliordpr(codigo,codalmac,codartic,cantidad)"
    SQL = SQL & " select " & Text1(0).Text & ",sliped.codalmac,sliped.codartic,sum(cantidad) from sliped,sartic"
    SQL = SQL & " WHERE  sliped.codartic=sartic.codartic  AND trazabilidad=1 "
    SQL = SQL & " AND numpedcl = " & Text1(4).Text
    SQL = SQL & " group by 1,2,3 order by numlinea"
    conn.Execute SQL
    
    
    
       
    'Sub lineas
    SQL = "Select * from sliordpr where codigo = " & Text1(0).Text
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        SQL = "INSERT INTO sliordpr2"
        SQL = SQL & "( codigo, codalmac, codartic ,codarti2,cantidad ) "
        SQL = SQL & "select " & Val(Text1(0).Text) & ", " & Val(RT!codAlmac) & ","
        Cantidad = CCur(RT!Cantidad)
        SQL = SQL & DBSet(RT!codartic, "T") & ",sarti1.codarti1,round(cantidad * " & DBSet(Cantidad, "N")
        'Factor conversion. Solo se aplica cuando metamos en stock. Ahora no
        SQL = SQL & " * factorconversion,4) FROM sarti1,sartic WHERE sarti1.codarti1=sartic.codartic AND sarti1.codartic ="
        'SQL = SQL & " * 1 FROM sarti1,sartic WHERE sarti1.codarti1=sartic.codartic AND sarti1.codartic ="
        SQL = SQL & DBSet(RT!codartic, "T")
        conn.Execute SQL
    
        RT.MoveNext
    Wend
    RT.Close
    
    Set RT = Nothing

    
End Sub

Private Sub UpdateaCantidadComponentes()
Dim SQL As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    SQL = "UPDATE sliordpr2 SET cantidad = " & DBSet(txtComponentes.Text, "S")
    SQL = SQL & " WHERE codartic = " & DBSet(Data2.Recordset!codartic, "T")
    SQL = SQL & " and codigo=" & Data1.Recordset!Codigo
    SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
    SQL = SQL & " and codarti2=" & DBSet(data3.Recordset!codarti2, "T")
    conn.Execute SQL
    Espera 0.5
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Screen.MousePointer = vbDefault
End Sub

'ACutalizaremos las sublineas(componentes)
'Es decir. Si insertamos o modificamos un elemento que tiene componentes
'insertaremos en sliorpd
Private Sub ActualizarComponentes()
Dim SQL As String

    If ModificaLineas = 2 Then
        'BORRAMOS los datos que hubieren
        SQL = "DELETE FROM sliordpr2 WHERE codartic = " & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!Codigo
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        conn.Execute SQL
    End If
        
    Espera 0.2
    
    
    'MAYO 2010
    'Por algun motivo, que desconzco, ya no utilizaba aqui el factor de conversion
    'Lo vuelvo a poner
    SQL = "INSERT INTO sliordpr2"
    SQL = SQL & "( codigo, codalmac, codartic ,codarti2,cantidad ) "
    

    SQL = SQL & "select " & Val(Text1(0).Text) & ", " & Val(txtAux(0).Text) & ","
    SQL = SQL & DBSet(txtAux(1).Text, "T") & ",sarti1.codarti1,round(cantidad * " & DBSet(txtAux(4).Text, "N")
    'Factor conversion. Solo se aplica cuando metamos en stock. Ahora no
    SQL = SQL & " * factorconversion,"
    SQL = SQL & CStr(IIf(vParamAplic.QUE_EMPRESA = 4, 2, 4))
    SQL = SQL & ") FROM sarti1,sartic WHERE sarti1.codarti1=sartic.codartic AND sarti1.codartic ="
    SQL = SQL & DBSet(txtAux(1).Text, "T")
    conn.Execute SQL
    
End Sub






'Praparamos para modificar la cantidad de los compoenntes
Private Sub ModificarCantidadComponentes()
        If Me.data3.Recordset.EOF Then
            MsgBox "No hay cantidad para modificar", vbExclamation
            Exit Sub
        End If
        If ModificaLineas = 2 Then Exit Sub
        
        If DataGrid2.Row < 0 Then Exit Sub
        Me.txtComponentes.Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + 10
        Me.txtComponentes.Left = DataGrid2.Left + DataGrid2.Columns(2).Left
        Me.txtComponentes.Width = DataGrid2.Columns(2).Width
        txtComponentes.Text = DataGrid2.Columns(2).Value
        txtComponentes.visible = True
        HabilitarModifCantidad True
        ModificaLineas = 2
        Me.lblIndicador.Caption = "Camb. cantidad"
        PonerBotonCabecera False
        'PonerFoco txtComponentes
 
End Sub



Private Sub HabilitarModifCantidad(Habilitar As Boolean)
    If Habilitar Then
        DeseleccionaGrid DataGrid1
        DeseleccionaGrid DataGrid2
    End If
    DataGrid1.Enabled = Not Habilitar
    DataGrid2.Enabled = Not Habilitar
End Sub



Private Sub txtComponentes_GotFocus()
    ConseguirFoco txtComponentes, 3
End Sub

Private Sub txtComponentes_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtComponentes_LostFocus()
    PonerFormatoSingle txtComponentes, 5
End Sub



Private Sub LlamaLotes()
    If Modo <> 5 Then Exit Sub
    CadenaDesdeOtroForm = ""
    With frmProdLotes
        .vIdProd = Val(Text1(0).Text)
        
        If ModificaLineas <> 0 Then
            ' esta insertando o modificando cantidades
            .vCodArtic = txtAux(1).Text
            .vCodAlmac = txtAux(0).Text
            .vCantidad = ImporteFormateado(txtAux(4).Text)
        Else
            'Esta en edicion
            .vCodAlmac = Data2.Recordset!codAlmac
            .vCodArtic = Data2.Recordset!codartic
            .vCantidad = Data2.Recordset!Cantidad
        End If
        
        .Show vbModal
    End With
    If CadenaDesdeOtroForm <> "" And ModificaLineas = 1 Then
        EjecutaSQL conAri, "commit", False
        Espera 1
        'Cagamos el dagrid
        CargaGrid2 DataGrid1, Data2
        PosicionaData2
        Modo = 5
    End If
End Sub


Private Sub PosicionaData2()
    On Error GoTo EPos
    Data2.Recordset.Find "Codartic= " & DBSet(txtAux(1).Text, "T")
    
    Exit Sub
EPos:
    MuestraError Err.Number, "Posicionando data2"
End Sub


Private Sub LlamaLotesLin()
    If Modo <> 6 Then Exit Sub
    CadenaDesdeOtroForm = ""
    With frmProdLotesLin
        .vIdProd = Val(Text1(0).Text)
        .vCodAlmac = Data2.Recordset!codAlmac
        .vCodArtic = Data2.Recordset!codartic
        .vCodarti2 = data3.Recordset!codarti2
    
        
        If ModificaLineas <> 0 Then
        
            .vCantidad = ImporteFormateado(Me.txtComponentes.Text)
        Else
            'Esta en edicion
            .vCantidad = data3.Recordset!Cantidad
        End If
        
        .Show vbModal
    End With
    If CadenaDesdeOtroForm <> "" And ModificaLineas = 1 Then
        'Cagamos el dagrid
        CargaGrid2 DataGrid2, data3
        PosicionarData
        Modo = 6
    End If
End Sub



Private Function ComprobarNumeroLoteLineas() As Boolean
Dim SQL As String
Dim R2 As ADODB.Recordset
Dim SinASig As String
Dim Er As String
    ComprobarNumeroLoteLineas = False

    
    Data2.Recordset.MoveFirst
    While Not Data2.Recordset.EOF
        SQL = "Select sum(cantlote) from sliordprlotes WHERE "
        SQL = SQL & " codigo = " & Data1.Recordset!Codigo & " AND codalmac =" & Data2.Recordset!codAlmac
        SQL = SQL & " AND codArtic = '" & DevNombreSQL(Data2.Recordset!codartic) & "'"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If miRsAux.EOF Then
            SQL = "Sin lote asignado"
        Else
            If IsNull(miRsAux.Fields(0)) Then
                SQL = "Cantidad lote sin asignar"
            Else
                If miRsAux.Fields(0) <> Data2.Recordset!Cantidad Then SQL = "NO coinciden la suma de lotes a la de produccion"
            End If
        End If
        miRsAux.Close
        If SQL <> "" Then
            SQL = Data2.Recordset!codartic & "  " & Data2.Recordset!NomArtic & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Data2.Recordset.MoveFirst
            Exit Function
        End If
        
        Data2.Recordset.MoveNext
    Wend
    Data2.Recordset.MoveFirst
    
    'Para las sublineas comprobaremos que el de ACEITE materia prima tiene LOTE asignado
    Set R2 = New ADODB.Recordset
    SinASig = ""
    While Not Data2.Recordset.EOF
        'Para la primera produccion veremos las lineas con materia prima y sus LOTES
        SQL = "Select codarti2,nomartic,cantidad  from sliordpr2,sartic where  sliordpr2.codarti2=sartic.codartic"
        SQL = SQL & " AND codigo = " & Data1.Recordset!Codigo & " AND codalmac =" & Data2.Recordset!codAlmac
        SQL = SQL & " AND sliordpr2.codArtic = '" & DevNombreSQL(Data2.Recordset!codartic) & "'"
        SQL = SQL & " AND factorconversion <>1"
        R2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Er = ""
        While Not R2.EOF
            SQL = ""
            'Este es
            SQL = "Select sum(cantlote) from sliordpr2lotes WHERE "
            SQL = SQL & " codigo = " & Data1.Recordset!Codigo & " AND codalmac =" & Data2.Recordset!codAlmac
            SQL = SQL & " AND codArtic = '" & DevNombreSQL(Data2.Recordset!codartic) & "'"
            SQL = SQL & " AND codArti2 = '" & DevNombreSQL(R2!codarti2) & "'"
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If miRsAux.EOF Then
                SQL = "Sin lote asignado"
            Else
                If IsNull(miRsAux.Fields(0)) Then
                    SQL = "Cantidad lote sin asignar"
                Else
                    If miRsAux.Fields(0) <> R2!Cantidad Then SQL = "NO coinciden la suma de lotes a la de produccion"
                End If
            End If
            miRsAux.Close
            If SQL <> "" Then
                SQL = "        - " & R2!codarti2 & "  " & R2!NomArtic & ": " & SQL & vbCrLf
                Er = Er & SQL
            End If
            R2.MoveNext
        Wend
        R2.Close
        If Er <> "" Then
            'Ha habido errores
            SQL = String(50, "=") & vbCrLf & Data2.Recordset!codartic & " -  " & Data2.Recordset!NomArtic & vbCrLf
            SinASig = SinASig & SQL & Er
        End If
        Data2.Recordset.MoveNext

    Wend
     Set R2 = Nothing
    Data2.Recordset.MoveFirst
    
        If SinASig <> "" Then
            SQL = "Materia prima: " & vbCrLf & vbCrLf & SinASig & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
    
    ComprobarNumeroLoteLineas = True
End Function


Private Sub HacerLlamadaALotesLinea()
Dim Ya As Boolean
Dim AntModo As Byte
Dim AntModoLin As Byte
Dim LlamaLoteDesdeAqui As Boolean

    If data3.Recordset Is Nothing Then Exit Sub
    If data3.Recordset.EOF Then Exit Sub
    data3.Recordset.MoveFirst
    While Not Ya
        LlamaLoteDesdeAqui = False
        
        If data3.Recordset!FactorConversion <> 1 Then
            LlamaLoteDesdeAqui = True
        Else
            If Val(DBLet(data3.Recordset!Trazabilidad, "N")) = 1 Then LlamaLoteDesdeAqui = True
        End If
        If LlamaLoteDesdeAqui Then
            'Guardamos modos
            AntModoLin = ModificaLineas
            AntModo = Modo
            ModificaLineas = 0 'para que lea del data3.cantidad
            Modo = 6
            LlamaLotesLin
            Modo = AntModo
            ModificaLineas = AntModoLin
            
           
        End If
        data3.Recordset.MoveNext
        If data3.Recordset.EOF Then
            data3.Recordset.MoveFirst
            Ya = True
        End If

        
    Wend
    
End Sub


