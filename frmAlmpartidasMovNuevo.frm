VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmpartidasMovNuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento partidas/lotes"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12150
   Icon            =   "frmAlmpartidasMovNuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   10800
      TabIndex        =   28
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1980
      Left            =   120
      TabIndex        =   13
      Top             =   410
      Width           =   11775
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   5280
         MaxLength       =   60
         TabIndex        =   7
         Tag             =   "c|N|N|||spartidas|cantotal|||"
         Top             =   1560
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   4
         Tag             =   "c|T|N|||spartidas|numalbar|||"
         Top             =   960
         Width           =   2265
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   240
         MaxLength       =   60
         TabIndex        =   6
         Tag             =   "c|N|N|||spartidas|codprove|||"
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   1
         Tag             =   "c|T|N|||spartidas|numlote|||"
         Top             =   360
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   4
         Left            =   5280
         TabIndex        =   5
         Tag             =   "c|N|N|||spartidas|codalmac|||"
         Text            =   "Text1 7"
         Top             =   960
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6240
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7320
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   2
         Tag             =   "c|T|N|||spartidas|codartic|||"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   240
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "c|F|N|||spartidas|fecha|||"
         Top             =   960
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
         Tag             =   "c|N|N|||spartidas|id|||"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad sctock"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Albaran"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmAlmpartidasMovNuevo.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar articulo"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Lote"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   6000
         Picture         =   "frmAlmpartidasMovNuevo.frx":010E
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacen"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   6000
         Picture         =   "frmAlmpartidasMovNuevo.frx":0210
         Tag             =   "-1"
         ToolTipText     =   "Buscar articulo"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmAlmpartidasMovNuevo.frx":0312
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   14
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   2055
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1875
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10440
      Top             =   2880
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
      TabIndex        =   11
      Top             =   0
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas produccion"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar cantidad en componentes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Hacer coupage"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Imprimir actual"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Imprimir seleccion"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6120
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8040
      Top             =   2040
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
      Left            =   10800
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmpartidasMovNuevo.frx":039D
      Height          =   3840
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6773
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
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   7080
      Width           =   6615
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   8
         Left            =   4680
         MaxLength       =   60
         TabIndex        =   34
         Tag             =   "c|N|N|||spartidas|id|||"
         Top             =   0
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   315
         Index           =   10
         Left            =   2880
         MaxLength       =   60
         TabIndex        =   32
         Tag             =   "c|N|N|||spartidas|id|||"
         Top             =   0
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   9
         Left            =   840
         MaxLength       =   60
         TabIndex        =   30
         Tag             =   "c|N|N|||spartidas|id|||"
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Salida"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   33
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Entrada"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Movimientos"
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
      TabIndex        =   18
      Top             =   2520
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnGeneraFactura 
         Caption         =   "Generar factura"
         Shortcut        =   ^Q
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpPedido 
            Caption         =   "&Pedido"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnImpOrde 
            Caption         =   "&Orden Instalación"
            Shortcut        =   ^O
            Visible         =   0   'False
         End
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
Attribute VB_Name = "frmAlmpartidasMovNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios
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

'-------------------------------------------------------------------------


Dim PrimeraVez As Boolean

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



'================================================================================








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
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
       
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select id from " & NombreTabla & " " & Ordenacion
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









Private Sub cmdAceptar_Click()
    If Modo = 1 Then HacerBusqueda
End Sub

Private Sub cmdCancelar_Click()
    CargaGrid False
    LimpiarCampos
    PonerModo 2
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'cad = Data1.Recordset.Fields(0) & "|"
        'cad = cad & Data1.Recordset.Fields(1) & "|"
        Cad = Data1.Recordset.Fields(0)
       ' RaiseEvent DatoSeleccionado2(cad)
        Unload Me
    End If
End Sub









Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        cmdAceptar_Click
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 20
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 10 'Mto Lineas Ofertas
'        .Buttons(11).Image = 37 'Cambiar cantidad componentes
'
'        'Enero08
'        .Buttons(12).Image = 21 'Cerrar orden produccion
'
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(15).Image = 27 'Imprimir Orden Instalacion
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

      
    LimpiarCampos   'Limpia los campos TextBox
   

    NombreTabla = "spartidas"
    Ordenacion = " ORDER BY id "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select id from spartidas where id= -1"
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    


    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    
    PonerModo 1
    If Me.DatosADevolverBusqueda <> "" Then
        Text1(5).Text = RecuperaValor(DatosADevolverBusqueda, 1)
        Text1(2).Text = RecuperaValor(DatosADevolverBusqueda, 2)
        PrimeraVez = True
    
    End If
    
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
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
''Mantenimiento de Articulos
'    If EsCabecera Then
        Text2(0).Text = ""
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
'    Else
'        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
'    End If
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)

    
    CadenaDevuelta = RecuperaValor(CadenaDevuelta, 1)
    CadenaConsulta = "select * from " & NombreTabla & " WHERE id = " & CadenaDevuelta & " " & Ordenacion
    PonerCadenaBusqueda
    
    
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

Private Sub imgCuentas_Click(Index As Integer)
If Modo = 2 Or Modo = 0 Then Exit Sub

    If Index = 0 Then
        'articulo
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco Text1(2)
'
'    Else
'            Set frmAlm = New frmAlmAlPropios
'            frmAlm.DatosADevolverBusqueda = "0"
'            frmAlm.Show vbModal
'            Set frmAlm = Nothing
    End If
End Sub

'Private Sub imgBuscar_Click(Index As Integer)
'Dim Indice As Byte
'
'    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass
'    Set frmPe = New frmFacEntPedidos
'    frmPe.DatosADevolverBusqueda2 = "0"
'    frmPe.Show vbModal
'    Set frmPe = Nothing
'
'
'
'    Screen.MousePointer = vbDefault
'
'
'End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
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






















Private Sub mnSalir_Click()

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
Dim vArtic As CArticulo
       
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    'If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            
            PonerFormatoFecha Text1(Index)
            
        Case 2
            
            Text2(0).Text = ""
            If Text1(Index).Text = "" Then Exit Sub
            Set vArtic = New CArticulo
           
            If vArtic.LeerDatos(Text1(2).Text) Then
              Text2(0).Text = vArtic.Nombre
        
            Else
                MsgBox "No existe el artículo", vbExclamation
                Text1(2).Text = ""
            End If
            Set vArtic = Nothing
        Case 3
            Text2(Index).Text = ""
            If Text2(Index).Text <> "0" Then
                If PonerFormatoEntero(Text1(Index)) Then
                    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Text1(Index).Text)
                    If CadenaDesdeOtroForm = "" Then
                        MsgBox "No existe el proveedor: " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                    End If
                    Text2(Index).Text = CadenaDesdeOtroForm
                    
                Else
                    Text1(Index).Text = ""
                End If
            End If
        Case 4 '
            Text2(1).Text = ""
            If PonerFormatoEntero(Text1(Index)) Then
                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", Text1(Index).Text)
                If CadenaDesdeOtroForm = "" Then
                    MsgBox "No existe el almacén: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                End If
                Text2(1).Text = CadenaDesdeOtroForm
                
            Else
                Text1(Index).Text = ""
            End If
            
        
    End Select
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


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    
        Cad = Cad & ParaGrid(Text1(0), 8, "Id")
        Cad = Cad & ParaGrid(Text1(5), 20, "LOTE ")
        Cad = Cad & ParaGrid(Text1(2), 20, "Articulo")
        Cad = Cad & "Descripcion|sartic|nomartic|T||50·"
        tabla = NombreTabla & " inner join sartic on " & NombreTabla & ".codartic = sartic.codartic"
        
        Titulo = "Movimiento lotaje"
        devuelve = "0|"

   
    
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 3
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
    
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




Private Sub PonerCampos()
Dim T1 As Single
    If Data1.Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    T1 = Timer
    LimpiarCampos
    lblIndicador.Caption = "Leyendo BD"
    
    Me.Refresh
    
    'Pongo los campos
    PonerCampos2
    DataGrid1.Enabled = True


    'Reestablezco
      lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
  
    
    'Para que cuando vaya mas lento no de la impresion..
    T1 = Timer - T1
    If T1 < 0.5 Then Espera 0.2
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub PonerCampos2()
On Error Resume Next
Dim cP As cPartidas

    
    Set cP = New cPartidas
    If cP.Leer(CLng(Data1.Recordset!ID)) Then
        Text1(0).Text = cP.IdPartida
        Text1(5).Text = cP.NUmlote
        Text1(1).Text = cP.Fecha
        Text1(2).Text = cP.codartic
        Text1(4).Text = cP.codAlmac
        If cP.codProve > 0 Then Text1(3).Text = cP.codProve
        Text1(6).Text = cP.NumAlbar
        Text1(7).Text = Format(cP.Cantidad, FormatoCantidad)
    
        Conn.Execute "Delete from tmppartidas where codusu = " & vUsu.codigo
        
        
        'NUEVO. Carga directamente de la tabla smovalotes
        'cP.CargaDatosParaListar
        
        PonerDatosExtra
        
        
'        CadenaConsulta = "codusu=" & vUsu.Codigo & " and codartic=" & DBSet(cP.codartic, "T") & " AND 1"
'        CadenaConsulta = DevuelveDesdeBD(conAri, "sum(cantidad)", "tmppartidas", CadenaConsulta, "1")
'        If CadenaConsulta = "" Then CadenaConsulta = 0
'        Text1(8).Text = Format(CadenaConsulta, FormatoCantidad)
    End If
    Set cP = Nothing
    'Para que haga el losffocus bien
    CargaGrid True
    Modo = 3
    Text1_LostFocus 2
    Text1_LostFocus 4
    Text1_LostFocus 3
    Modo = 2
    
    
    

    
    
    
    '-- Esto permanece para saber donde estamos
    If Err.Number <> 0 Then Err.Clear

End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
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
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        If cmdRegresar.visible Then cmdRegresar.Cancel = True
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
     DataGrid1.Enabled = Modo = 2
        

    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True

    b = Modo = 0 Or Modo = 2 Or Modo >= 5
    For i = 1 To 7
        BloquearTxt Text1(i), b
    Next
    BloquearTxt Text1(8), True  'siempre bloqueado

    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)

    imgFecha(0).visible = b
     
    For i = 0 To Me.imgCuentas.Count - 1
        imgCuentas(i).visible = b
    Next i


    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Los kilos totatels
    b = Modo = 2 Or Modo = 4 Or Modo = 5

    
    
    'Abrir un coupage cerrado solo para admon
    b = False
    If Modo = 1 Then
        b = True
    Else
        If Modo = 4 Then b = vUsu.Nivel < 1
    End If
  
    
    

    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
















Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            
        Case 6  'Modificar
           
        Case 7  'Borrar
            
            
        Case 14, 15
            '10 Lineas 12, 14
            'IMPRIMIR (14)    y cerrar(12) orden produccion
            '--------------------------------------------------------------------
        
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una partida/lote", vbExclamation
                Exit Sub
            End If
            
            If Button.Index = 15 Then
                If Data1.Recordset.RecordCount > 1 Then
                    
                    CadenaConsulta = "Va a imprimir los datos de las " & Data1.Recordset.RecordCount & " lotes/partidas."
                    CadenaConsulta = CadenaConsulta & vbCrLf & vbCrLf & "El proceso puede llevar mucho tiempo. ¿Continuar?"
                    If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then CadenaConsulta = ""
                    If CadenaConsulta = "" Then Exit Sub
                    
                    'Generar todos los datos del data1.
                    GenerarDatosPartdas
                    CadenaConsulta = ""
                End If
            End If
           'Imprimir orden prod
           With frmImprimir
               .ConSubInforme = False
               .FormulaSeleccion = "{tmppartidas.codusu} = " & vUsu.codigo
               .NombreRPT = "morMovLotes.rpt"
               .OtrosParametros = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
               .NumeroParametros = 1
               .Titulo = "Movientos lotes/partidas"
               .Opcion = 2003 'Esta libre
               .Show vbModal
           End With
    

        Case 15 '
          
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
      
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    







Private Sub CargaGrid(enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    

    
    CargaGrid2
    DataGrid1.ScrollBars = dbgAutomatic
    
 
    
    
    If Not enlaza Then
    
        
        Text1(9).Text = ""
        Text1(10).Text = ""
        Text1(8).Text = ""
    End If
    
    
    DataGrid1.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Private Sub CargaGrid3(enlaza As Boolean)
'Dim SQL As String
'
'    SQL = "codigo = -1"
'
'
'    If enlaza Then
'       If Not Data2.Recordset.EOF Then
'            SQL = " codigo = " & Data1.Recordset!Codigo
'            SQL = SQL & " AND codalmac = " & Data2.Recordset!codAlmac
'            SQL = SQL & " AND sliordpr2.codartic = " & DBSet(Data2.Recordset!codArtic, "T")
'
'       End If
'    End If
'
'    SQL = "select codarti2,nomartic,cantidad  from sliordpr2,sartic where  sliordpr2.codarti2=sartic.codartic AND " & SQL
'    data3.ConnectionString = Conn
'    data3.RecordSource = SQL
'    data3.Refresh
'    If DataGrid2.DataSource Is Nothing Then DataGrid2.ClearFields
'
'    Set DataGrid2.DataSource = data3
'    DataGrid2.RowHeight = 290
'    DataGrid2.Columns(0).Caption = "Codigo"
'    DataGrid2.Columns(0).Width = 1900
'
'
'    DataGrid2.Columns(1).Caption = "Articulo"
'    DataGrid2.Columns(1).Width = 3700
'
'    DataGrid2.Columns(2).Caption = "Cantidad"
'    DataGrid2.Columns(2).Width = 1200
'    DataGrid2.Columns(2).NumberFormat = FormatoCantidad
'    DataGrid2.Columns(2).Alignment = dbgRight
'End Sub



Private Sub CargaGrid2()
Dim i As Byte

    On Error GoTo ECargaGrid

    Data2.Refresh


   ' fechamov,horamovi,if(tipomovi=0,""S"",""E""),detamovi,nomtipom,
   'cantidad,document,numlinea,linea2,codarti2,codprocliope,codalmac
                DataGrid1.Columns(0).Caption = "Fecha"
                DataGrid1.Columns(0).Width = 1350

                
                DataGrid1.Columns(1).Caption = "Hora."
                DataGrid1.Columns(1).Width = 1000
                DataGrid1.Columns(1).NumberFormat = "hh:mm:ss"

                DataGrid1.Columns(2).Caption = "E/S"
                DataGrid1.Columns(2).Width = 600
      
                DataGrid1.Columns(3).Caption = "Det"
                DataGrid1.Columns(3).Width = 900
      
                DataGrid1.Columns(4).Caption = "Movimiento"
                DataGrid1.Columns(4).Width = 2700
      
                
                DataGrid1.Columns(5).Caption = "Cantidad"
                DataGrid1.Columns(5).Width = 1400
                DataGrid1.Columns(5).Alignment = dbgRight
                DataGrid1.Columns(5).NumberFormat = FormatoCantidad
             
                DataGrid1.Columns(6).Caption = "Documento"
                DataGrid1.Columns(6).Width = 2400
                DataGrid1.Columns(6).Alignment = dbgCenter
                
                DataGrid1.Columns(7).Caption = "Linea"
                DataGrid1.Columns(7).Width = 600
                DataGrid1.Columns(7).Alignment = dbgCenter
                For i = 8 To 11
                    DataGrid1.Columns(i).visible = False
                Next
            For i = 0 To DataGrid1.Columns.Count - 1
                DataGrid1.Columns(i).Locked = True
                DataGrid1.Columns(i).AllowSizing = False
            Next i
            DataGrid1.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


















Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid False
    
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
'   devuelva.If Index = 1 Then
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    SQL = "select fechamov,horamovi,if(tipomovi=0,""S"",""E""),detamovi,nomtipom,"
    SQL = SQL & "cantidad,document,numlinea,linea2,codarti2,codprocliope,codalmac  from "
    SQL = SQL & " smovalotes,stipom where smovalotes.detamovi=stipom.codtipom"
    If enlaza Then

        SQL = SQL & " AND numlote = " & DBSet(Text1(5).Text, "T")
        SQL = SQL & " AND codartic = " & DBSet(Text1(2).Text, "T")
    Else
        SQL = SQL & " AND numlote = '@@@##@'"
    End If
    SQL = SQL & " Order by fechamov desc, horamovi desc"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = False
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        'Generar Albaran desde Pedido
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenAlbaran.Enabled = b
        
        Toolbar1.Buttons(12).Enabled = b
        Me.mnGeneraFactura.Enabled = b
        Toolbar1.Buttons(13).Enabled = b
        
        
        
      
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub







    
Private Sub GenerarDatosPartdas()
Dim cP As cPartidas
    Set miRsAux = Nothing
    NumRegElim = Data1.Recordset.AbsolutePosition
    lblIndicador.Caption = "Leyendo BD"
    lblIndicador.Refresh
    Data1.Recordset.MoveFirst
    Set cP = New cPartidas
    While Not Data1.Recordset.EOF
        Me.lblIndicador.Caption = "Gen: " & Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        Me.lblIndicador.Refresh
        If Val(Data1.Recordset!ID) <> Val(Text1(0).Text) Then
            If cP.Leer(CLng(Data1.Recordset!ID)) Then cP.CargaDatosParaListar
        End If
        Data1.Recordset.MoveNext
    Wend
    Set cP = Nothing
    Data1.Recordset.MoveFirst
    NumRegElim = NumRegElim - 1
    If NumRegElim > 0 Then Data1.Recordset.Move NumRegElim, 1
    Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub








Private Sub PonerDatosExtra()
Dim SQL As String
Dim Im As Currency
Dim RL As ADODB.Recordset

    'IMportes
    SQL = "select tipomovi,sum(cantidad) total from smovalotes"
    SQL = SQL & " WHERE numlote = " & DBSet(Text1(5).Text, "T")
    SQL = SQL & " AND codartic = " & DBSet(Text1(2).Text, "T")
    SQL = SQL & " group by tipomovi"
    Set RL = New ADODB.Recordset
    RL.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'En poner campos hay un limpiarcampos que pone esto
    'Text1(9).Text = ""
    'Text1(10).Text = ""
    'Text1(8).Text = ""
    Im = 0
    While Not RL.EOF
        
        If Not IsNull(RL!tipomovi) Then
            If Val(RL!tipomovi) = 0 Then
                'SALIDA
                Im = Im - RL!Total
                Text1(10).Text = Format(RL!Total, FormatoCantidad)
                
            Else
                'ENTRADA
                Im = Im + RL!Total
                Text1(9).Text = Format(RL!Total, FormatoCantidad)
            End If
        End If
        RL.MoveNext
    Wend
    RL.Close
    Set RL = Nothing
    Text1(8).Text = Format(Im, FormatoCantidad)
End Sub














