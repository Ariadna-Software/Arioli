VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmCoupage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Coupage"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   9825
   Icon            =   "frmAlmCoupage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAlmCoupage.frx":000C
      Left            =   7320
      List            =   "frmAlmCoupage.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7680
      TabIndex        =   25
      Text            =   "Text3"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   5520
      MaxLength       =   18
      TabIndex        =   7
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   3780
      Left            =   120
      TabIndex        =   17
      Top             =   410
      Width           =   9495
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   6
         Left            =   5160
         TabIndex        =   6
         Tag             =   "Deposito|N|S|||olicoupage|deposito|||"
         Text            =   "Text1 7"
         Top             =   3000
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3240
         MaxLength       =   60
         TabIndex        =   2
         Tag             =   "Lote|T|N|||olicoupage|numlote|||"
         Top             =   360
         Width           =   2625
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Tag             =   "Almacen|N|N|0||olicoupage|codalmac|||"
         Text            =   "Text1 7"
         Top             =   3000
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   3000
         Width           =   3735
      End
      Begin VB.CheckBox chkCoup 
         Caption         =   "Hecho"
         Height          =   255
         Left            =   6960
         TabIndex        =   11
         Tag             =   "Hecho|N|N|||olicoupage|YaCreado|||"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Tag             =   "Articulo|T|N|||olicoupage|codartic|||"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   795
         Index           =   3
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "Obs|T|S|||olicoupage|descripcion|||"
         Top             =   840
         Width           =   7545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha creación|F|N|||olicoupage|fecha|dd/mm/yyyy|N|"
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
         Tag             =   "Nº ord produccion|N|S|0||olicoupage|codigo|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Depósito"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   30
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Lote"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   735
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmAlmCoupage.frx":0022
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacen"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmAlmCoupage.frx":0124
         Tag             =   "-1"
         ToolTipText     =   "Buscar articulo"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "F. creacion"
         Height          =   255
         Index           =   14
         Left            =   1680
         TabIndex        =   19
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2640
         Picture         =   "frmAlmCoupage.frx":0226
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   18
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8250
      TabIndex        =   10
      Top             =   7560
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   7560
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2400
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
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar cantidad en componentes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Hacer coupage"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Orden Instal."
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
         Left            =   6240
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2280
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8250
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCoupage.frx":02B1
      Height          =   2640
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4657
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
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Total kilos"
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Lineas coupage"
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
      TabIndex        =   22
      Top             =   4320
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
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnGeneraFactura 
         Caption         =   "Generar factura"
         Shortcut        =   ^Q
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
Attribute VB_Name = "frmAlmCoupage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo para cuando viene desde pantalla smoval
Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios
Attribute frmAlm.VB_VarHelpID = -1
            
Private WithEvents frmB2 As frmBuscaGrid
Attribute frmB2.VB_VarHelpID = -1

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




'Cuando muestra el grid, lo añade direcamtene. Si pulsa cancelar, deberiamos borrarlo
Dim EsModificarDeAñadir As Boolean

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange





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
                    'ActualizarLineasPedido
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea Then
'
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
            
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    EsModificarDeAñadir = True
                End If
                Me.DataGrid1.Enabled = True
            End If
            
            

    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
            
            If EsModificarDeAñadir Then
                'Ha dado cancelar NADA mAs insertar la  la linea, es decir, NO quiere la linea
                ModificaLineas = 0
                BotonEliminarLinea False
                Espera 0.1
                ModificaLineas = 1
            End If
            
            CargaTxtAux False, False
           
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True

    End Select
End Sub


Private Sub PonerLote()
        CadenaConsulta = DevuelveDesdeBD(conAri, "contador + 1", "stipom", "codtipom", "LOV", "T")
        Text1(5).Text = "MOSTRA" & CadenaConsulta & "-"
        
        CadenaConsulta = Text1(1).Text
        If CadenaConsulta = "" Then CadenaConsulta = Now
        If Month(CDate(CadenaConsulta)) < 10 Then
            Text1(5).Text = Text1(5).Text & Year(Now) - 1
        Else
            Text1(5).Text = Text1(5).Text & Year(Now)
        End If
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        'Para La Vall,cogera el LOV
        PonerLote
        
        'Codalmac
        Text1(4).Text = "1"
        Text2(1).Text = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", Text1(4).Text)
        
        
        Text1(2).Text = DevuelveDesdeBD(conAri, "articMolturacion", "vallparam", "1", "1")
        Text2(0).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(2).Text, "T")
    Else
        CadenaConsulta = "numlote like '%C' AND 1"
        CadenaConsulta = DevuelveDesdeBD(conAri, "mid(numlote,1,length(numlote)-1)", "olicoupage", CadenaConsulta, "1 order by 1 desc")
        If CadenaConsulta <> "" Then
            If IsNumeric(CadenaConsulta) Then
                CadenaConsulta = Val(CadenaConsulta) + 1
                Text1(5).Text = CadenaConsulta & "C"
            End If
        End If
    End If
    

    Me.chkCoup.Value = 0
    txtTotal.Tag = 0

    Text1(1).Text = Format(Now, "dd/mm/yyyy hh:mm") 'Fecha
    If vParamAplic.QUE_EMPRESA = 4 Then
        PonerFoco Text1(6)
    Else
        PonerFoco Text1(1)
    End If
End Sub


Private Sub BotonAnyadirLinea()
    
    ObtenerLineCoupage False
    Exit Sub
    
    
    
    
    
    
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
   

    
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
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'    LimpiarCampos
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

    EsModificarDeAñadir = False
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If Val(Data1.Recordset!yacreado) = 1 Then
        If vUsu.Nivel > 0 Then
            MsgBox "Ya se ha hecho el coupage", vbExclamation
            Exit Sub
        End If
    End If

    cad = "Coupage." & vbCrLf
    cad = cad & "----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el coupage:"
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")
    cad = cad & vbCrLf & "Artículo:  " & Text1(2).Text & " - " & Text2(0).Text
    cad = cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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


Private Sub BotonEliminarLinea(HacerLaPregunta As Boolean)
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String
Dim QUitarTambienDeLineas As Boolean
Dim Eliminar_ As Boolean
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    Eliminar_ = False
    If HacerLaPregunta Then
        SQL = "¿Seguro que desea eliminar la línea de coupage?     "
        SQL = SQL & vbCrLf
        SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
        SQL = SQL & vbCrLf & "Kilos:  " & Format(Data2.Recordset!cantlote, FormatoPrecio)
        SQL = SQL & vbCrLf & "Deposito:  " & Format(Data2.Recordset!Deposito, "00")
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then Eliminar_ = True
    Else
        Eliminar_ = True
    End If
    
    If Eliminar_ Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!Codigo
        
        'Voy a ver si hay mas de una linea de ese articulo
        QUitarTambienDeLineas = True
        CadenaConsulta = Mid(SQL, 8) 'quito el where
        CadenaConsulta = DevuelveDesdeBD(conAri, "count(*)", "olicoupagelinlotes", CadenaConsulta & " AND 1", "1")
        If Val(CadenaConsulta) > 1 Then QUitarTambienDeLineas = False
        

        'Las lineas
        CadenaConsulta = "DELETE FROM olicoupagelinlotes  " & SQL
        CadenaConsulta = CadenaConsulta & " AND deposito = " & Data2.Recordset!Deposito
        conn.Execute CadenaConsulta
        
        If QUitarTambienDeLineas Then
            CadenaConsulta = "DELETE FROM olicoupagelin  " & SQL
            conn.Execute CadenaConsulta
        Else
            'UPDATEAMOS cantidad
            CadenaConsulta = "UPDATE olicoupagelin set kilos = kilos - " & DBSet(Data2.Recordset!cantlote, "N")
            CadenaConsulta = CadenaConsulta & SQL
            conn.Execute CadenaConsulta
        End If
        
        
        'Los kilos totaltes
        CambiaKilosTotales Data2.Recordset!cantlote, 0
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataPosicion Me.Data2, NumRegElim, SQL
        

    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub






Private Sub CambiaKilosTotales(KilosAntes As Currency, KilosAhora As Currency)
Dim K As Currency
    K = CCur(txtTotal.Tag)
    K = K - KilosAntes + KilosAhora
    txtTotal.Tag = K
    txtTotal.Text = Format(K, FormatoPrecio)
End Sub
Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

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
        cad = Data1.Recordset.Fields(0)
        RaiseEvent DatoSeleccionado2(cad)
        Unload Me
    End If
End Sub







Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If Modo = 5 And ModificaLineas = 0 Then
        'Modo lineas sin insertar ni modificar
        ' LanzaLote -1
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
    btnPrimero = 20
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
        
        
        .Buttons(14).Image = 16 'Imprimir Pedido
      '  .Buttons(15).Image = 27 'Imprimir Orden Instalacion
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

      
    LimpiarCampos   'Limpia los campos TextBox
   

    NombreTabla = "olicoupage"
    Ordenacion = " ORDER BY codigo "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & " where codigo= "
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
'Mantenimiento de Articulos
    If EsCabecera Then
        Text2(0).Text = ""
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    Else
       ' txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    End If
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









Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    CadenaConsulta = CadenaDevuelta
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
            EsCabecera = True
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco Text1(2)
    
    Else
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
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


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea True
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
    Else   'Modificar Pedido
         If Not IsNull(Me.Data1.Recordset) Then
            If Val(Data1.Recordset!yacreado) > 0 Then
                MsgBox "Ya se ha hecho el coupage", vbExclamation
                Exit Sub
            End If
        End If
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
            If Modo = 3 Then PonerLote
        Case 2
            
            Text2(0).Text = ""
            If Text1(Index).Text = "" Then Exit Sub
            Set vArtic = New CArticulo
            EsCabecera = False 'Para ver si bloquea
            If vArtic.LeerDatos(Text1(2).Text) Then
                vArtic.MostrarStatusArtic EsCabecera
                If Not EsCabecera Then Text2(0).Text = vArtic.Nombre
            Else
                MsgBox "No existe el artículo", vbExclamation
                Text1(2).Text = ""
            End If
            Set vArtic = Nothing
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
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, Devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 15, "Nº Orden")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha ")
        cad = cad & ParaGrid(Text1(2), 20, "Articulo")
        cad = cad & "Descripcion|sartic|nomartic|T||40·"
        Tabla = NombreTabla & " inner join sartic on " & NombreTabla & ".codartic = sartic.codartic"
        
        Titulo = "Coupages"
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
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        Tabla = "sdirec"
        Devuelve = "0|1|"
    End If
    
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
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
Dim b As Boolean
Dim Tot As Currency
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True
        
    'Total
    b = DataGrid1.Enabled
    DataGrid1.Enabled = False
    Tot = 0
    If Not Data2.Recordset.EOF Then
        While Not Data2.Recordset.EOF
            Tot = Tot + DBLet(Data2.Recordset!cantlote, "N")
            Data2.Recordset.MoveNext
        Wend
        Data2.Recordset.MoveFirst
    End If
    txtTotal.Tag = Tot
    txtTotal.Text = Format(Tot, FormatoPrecio)
    DataGrid1.Enabled = b
    
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
    
    'Para que haga el losffocus bien
    Modo = 3
    Text1_LostFocus 2
    Text1_LostFocus 4
    Modo = 2
    
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

    b = Modo = 0 Or Modo = 2 Or Modo >= 5
    For i = 1 To 6
        BloquearTxt Text1(i), b
    Next
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        'NO dejo cambiar el numero de lote del cupage
        BloquearTxt Text1(5), Modo <> 1
    End If
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    'Las imagenes añadimos el modo 6
    b = b And Modo <> 6
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i


    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Los kilos totatels
    b = Modo = 2 Or Modo = 4 Or Modo = 5
    txtTotal.visible = b
    lblTotal.visible = b
    
    
    'Abrir un coupage cerrado solo para admon
    b = False
    If Modo = 1 Then
        b = True
    Else
        If Modo = 4 Then b = vUsu.Nivel < 1
    End If
    Me.chkCoup.Enabled = b
    
    
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


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function


    'OCtubre 2014
    'Nuevas comprobaciones. Como dejamos que el deposito NO este vacio
    ' comprobamos que el articulo del deposito es igual que el del coupage
    '
    
    'El deposito destino NO puede ser mayor que maximo numero depositos
    If Modo = 3 Then
        If Val(Me.Text1(6).Text) > MaxNumDepositos_ Or Val(Me.Text1(6).Text) < 1 Then
            MsgBox "El deposito no existe (max: " & MaxNumDepositos_ & ")", vbExclamation
            Exit Function
        End If
    End If
    
    CadenaConsulta = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", Text1(2).Text, "T")
    If CadenaConsulta = "1" Then
        MsgBox "No es materia prima", vbExclamation
        Exit Function
    End If
    
    CadenaConsulta = DevuelveDesdeBD(conAri, "codartic", "proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote", "numdeposito", Text1(6).Text)
    If CadenaConsulta <> "" Then
        If CadenaConsulta <> Text1(2).Text Then
                MsgBox "Articulo diferente deposito - coupage", vbExclamation
                b = False
        End If
    End If
   
   
    
   
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim i As Byte
Dim C As String



    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For i = 1 To txtAux.Count
        If txtAux(i).Text = "" Then
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
        
  
    
    'Si la cantidad es mayor de lo que queda en el deposito, tiene que marcar fin cuba
    C = DevuelveDesdeBD(conAri, "kilos", "proddepositos", "numdeposito", CStr(Data2.Recordset!Deposito))
    If CCur(C) <= ImporteFormateado(txtAux(1).Text) Then
        If Combo1.ListIndex = 0 Then
            MsgBox "Cantidad mayor(o igual) de la del deposito. Debe marcar Fin deposito", vbExclamation
            b = False
        End If
    End If
    
    
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function









Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
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
            
        Case 10, 12, 14
            '10 Lineas 12, 14
            'IMPRIMIR (14)    y cerrar(12) orden produccion
            '--------------------------------------------------------------------
        
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una orden de produccion", vbExclamation
                Exit Sub
            End If
            
            
            If Button.Index < 14 Then
                'Modificar lineas o cerrar produccion
                If Val(Data1.Recordset!yacreado) > 0 Then
                    MsgBox "Ya se ha hecho el coupage", vbExclamation
                    Exit Sub
                End If
                
                If Button.Index = 10 Then
                        mnLineas_Click
                Else
                    If Me.Data2.Recordset.EOF Then
                        MsgBox "No tiene lineas de coupage", vbExclamation
                        Exit Sub
                    End If
                
                
                
                    If Not ComprobarNumeroLoteLineas Then Exit Sub
                    
                    


                    If BLOQUEADesdeFormulario(Me) Then
                
                        frmProduVarios.Intercambio = Data1.Recordset!Codigo & "|" & Data1.Recordset!Fecha & "|" & Data1.Recordset!codAlmac & "|"
                        frmProduVarios.Opcion = 1
                        frmProduVarios.Show vbModal
                    
                        'TErminamos de bloquear
                        TerminaBloquear
                        
                        'Refrescamos
                        CadenaConsulta = Data1.RecordSource
                        Data1.Refresh
                        'Y ponemos los campos
                        PosicionarData
                      
                    End If
                End If '=10
            Else
                'Imprimir orden prod
                With frmImprimir
                    .ConSubInforme = True
                    .FormulaSeleccion = "{olicoupage.codigo} = " & Data1.Recordset!Codigo
                    .NombreRPT = "rCoupage.rpt"
                    .OtrosParametros = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
                    .NumeroParametros = 1
                    .Titulo = "Coupage"
                    .Opcion = 2003 'Esta libre
                    .Show vbModal
                End With
            End If

        Case 15 'Imprimir Orden Instalacion
          
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
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String
Dim vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
'        'Conseguir el siguiente numero de linea
'        SQL = "INSERT INTO olicoupagelin"
'        SQL = SQL & "(`codigo`,`codartic`,`kilos`) "
'        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ","
'        SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(3).Text, "N") & ")"
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
'
'        CambiaKilosTotales 0, ImporteFormateado(txtAux(3).Text)
'
'
'        LanzaLote False
        
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
        SQL = "UPDATE olicoupagelinlotes set cantlote  = " & DBSet(txtAux(1).Text, "N")
        SQL = SQL & " ,FinCuba =" & Me.Combo1.ItemData(Combo1.ListIndex)
        SQL = SQL & " WHERE codigo =" & Data1.Recordset!Codigo
        SQL = SQL & " AND codartic =" & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " AND Linea =" & DBSet(Data2.Recordset!linea, "T")

        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        'Actualizamos los kilos en olicoupagelin
        Espera 0.5
        SQL = "codartic =" & DBSet(Data2.Recordset!codartic, "T") & " AND codigo "
        SQL = DevuelveDesdeBD(conAri, "sum(cantlote)", "olicoupagelinlotes", SQL, CStr(Data1.Recordset!Codigo))
        SQL = " SET kilos = " & TransformaComasPuntos(SQL)
        SQL = "UPDATE olicoupagelin " & SQL & " WHERE codartic =" & DBSet(Data2.Recordset!codartic, "T") & " AND codigo = " & Data1.Recordset!Codigo
        conn.Execute SQL
        
        CambiaKilosTotales Data2.Recordset!cantlote, ImporteFormateado(txtAux(1).Text)
        
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
    
 
    
    
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
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



Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte

    On Error GoTo ECargaGrid

    vData.Refresh

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                
                vDataGrid.Columns(0).Caption = "Articulo"
                vDataGrid.Columns(0).Width = 1500

                
                vDataGrid.Columns(1).Caption = "Desc. Artículo"
                vDataGrid.Columns(1).Width = 3800
                If vParamAplic.QUE_EMPRESA = 4 Then vDataGrid.Columns(1).Width = 3000

                vDataGrid.Columns(2).Caption = "depósito"
                vDataGrid.Columns(2).Width = 900
                vDataGrid.Columns(2).Alignment = dbgCenter
                vDataGrid.Columns(2).NumberFormat = "00"


                vDataGrid.Columns(3).Caption = "Lote"
                vDataGrid.Columns(3).Width = 1050
                If vParamAplic.QUE_EMPRESA = 4 Then vDataGrid.Columns(3).Width = 1850
                             
                vDataGrid.Columns(4).Caption = "Kilos"
                vDataGrid.Columns(4).Width = 1100
                vDataGrid.Columns(4).Alignment = dbgRight
                vDataGrid.Columns(4).NumberFormat = FormatoPrecio
             

                vDataGrid.Columns(5).Caption = "Fin"
                vDataGrid.Columns(5).Width = 700
                
                'Numliena
                vDataGrid.Columns(6).visible = False

    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
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
Dim i As Byte
  i = 1
    'On Error Resume Next
    On Error GoTo Quitar
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        'For I = 1 To txtAux2.Count 'TextBox
            txtAux(1).Top = 290
            txtAux(1).visible = False
        'Next I
        'cmdAux2(1).visible = visible
            Combo1.visible = False
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
         
                txtAux(1).Text = ""
                BloquearTxt txtAux(1), False
            'Next I
          
    
        Else 'Vamos a modificar
           ' For I = 1 To txtAux2.Count
                
                txtAux(1).Text = DataGrid1.Columns(4).Text
           '     If I < 3 Then BloquearTxt txtAux(I), True
           ' Next I
                If DBLet(Data2.Recordset!fincuba, "T") = "" Then
                    Combo1.ListIndex = 0
                Else
                    Combo1.ListIndex = 1
                End If
        End If
        

    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
       ' For I = 1 To txtAux.Count
     
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
       ' Next I
        'cmdAux(0).Top = alto
    '    cmdAux2(1).Top = alto
    '    cmdAux2(1).Height = DataGrid1.RowHeight

        'Fijamos anchura y posicion Left
        '--------------------------------
        
        
       ' cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 35
        'Nom Artic
        
        'Cantidad
        
        txtAux(1).Left = DataGrid1.Columns(4).Left + 150
        txtAux(1).Width = DataGrid1.Columns(4).Width - 30
        txtAux(1).visible = visible
        Combo1.Top = alto
        Combo1.Left = DataGrid1.Columns(5).Left + 150
        Combo1.Width = DataGrid1.Columns(5).Width - 30
        'cmdAux(1).visible = limpiar

    End If
    Combo1.visible = visible
Quitar:
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
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


    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    'Index SOLO puede ser 1
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 2) Then   'Tipo 3: FormatoCantidad
    
                Else
                    txtAux(Index).Text = Format(Data2.Recordset!cantlote, FormatoPrecio)
                    PonerFoco txtAux(Index)
                End If
            End If
        
    

End Sub


Private Sub BotonMtoLineas()
       
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean



    On Error GoTo FinEliminar

        conn.BeginTrans
        conn.Execute "Delete from olicoupagelinlotes where codigo =" & Text1(0).Text
        conn.Execute "Delete from olicoupagelin where codigo =" & Text1(0).Text
        conn.Execute "Delete from olicoupage where codigo =" & Text1(0).Text
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
    
'
'    SQL = "SELECT olicoupagelin.codartic,nomartic,kilos "
'    SQL = SQL & " FROM olicoupagelin,sartic WHERE olicoupagelin.codartic=sartic.codartic AND "

' JUNIO 2014
    SQL = "select  olicoupagelin.codartic,nomartic,deposito,numlote,olicoupagelinlotes.cantlote,"
    SQL = SQL & "  if(fincuba,'Si','') fincuba,linea"
    SQL = SQL & "  FROM olicoupagelin inner join sartic ON olicoupagelin.codartic=sartic.codartic"
    SQL = SQL & "  LEFT JOIN  olicoupagelinlotes ON olicoupagelin.codigo=olicoupagelinlotes.codigo  AND"
    SQL = SQL & "  olicoupagelin.codartic = olicoupagelinlotes.codartic WHERE "


    
    
    
    
    
    
    If enlaza Then
        SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, "olicoupagelin")
    Else
        SQL = SQL & " olicoupagelin.codigo = -1"
    End If
    SQL = SQL & " Order by deposito"
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







Private Sub InsertarCabecera()
Dim cT As CTiposMov
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codigo")
    If InsertarDesdeForm(Me) Then
    
            'ActualizarLineasPedido
            If vParamAplic.QUE_EMPRESA = 4 Then
                Set cT = New CTiposMov
                cT.Leer "LOV"
                cT.IncrementarContador cT.TipoMovimiento
                Set cT = Nothing
            End If
            'Si tiene pedido traeremos las lineas del pedido
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE codigo = " & Text1(0).Text & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            
            'Si e s la vall, y el depostio era destibno es el 18, añado YA la linea de deposito
            If vParamAplic.QUE_EMPRESA = 4 Then
                If Val(Text1(6).Text) = 18 Then
                    CadenaConsulta = DevuelveDesdeBD(conAri, "partida", "proddepositos", "numdeposito", "18")
                    If CadenaConsulta <> "" Then
                        Espera 0.3
                        ObtenerLineCoupage True
                    End If
                End If
            End If
            
            
            BotonMtoLineas
            BotonAnyadirLinea
    
    End If

End Sub







'>0  Nueva linea. Cogeremos los datos del txtaux
' -1:   dblclick en el datagrid
'Private Sub LanzaLote(DesdeGrid As Boolean)
'
'
'    'Articulos que NO llevan lote
''    If Linea < 0 Then
''        If Data2.Recordset!codArtic = vParamAplic.ArtReciclado Then Exit Sub
''        If Not EsArticuloTrazabilidad(CStr(Data2.Recordset!codArtic)) Then Exit Sub
''    Else
''        If Not ElArticulo.Trazabilidad Then Exit Sub
''    End If
'
'    frmAlmCoupLotes.vIdCup = Data1.Recordset!Codigo
'    frmAlmCoupLotes.vCodAlmac = Data1.Recordset!codalmac
'    '
'    If DesdeGrid Then
'        frmAlmCoupLotes.vCantidad = Data2.Recordset!Kilos
'        frmAlmCoupLotes.vCodArtic = Data2.Recordset!codartic
'    Else
'        'frmFacLotes.vCodAlmac = CInt(txtAux(0).Text)
'        frmAlmCoupLotes.vCantidad = ImporteFormateado(txtAux(3).Text)
'    End If
'    frmAlmCoupLotes.Show vbModal
'        frmAlmCoupLotes.vCodArtic = txtAux(1).Text
'End Sub




Private Function ComprobarNumeroLoteLineas() As Boolean
Dim SQL As String
Dim FinDeposito As String

    ComprobarNumeroLoteLineas = False

    
    
    'Para las sublineas comprobaremos que el de ACEITE materia prima tiene LOTE asignado
    
    'Cabcera. Tiene numero de lote y este NO existe
    SQL = DBLet(Data1.Recordset!NUmlote, "T")
    If SQL = "" Then
        MsgBox "Numero lote Coupage incorrecto", vbExclamation
        Exit Function
    End If
    
    'NO EXISTE en partdidas
    SQL = DevuelveDesdeBD(conAri, "numlote", "spartidas", "numlote", SQL, "T")
    If SQL <> "" Then
        MsgBox "Ya existe el numero de lote en partidas: " & SQL, vbExclamation
        Exit Function
    End If
    
    If Data2.Recordset.EOF Then
        MsgBox "No existen lineas a coupar", vbExclamation
        Exit Function
    End If
    
    
    SQL = DevuelveDesdeBD(conAri, "sum(kilos)", "olicoupagelin", "codigo", Data1.Recordset!Codigo)
    CadenaConsulta = DevuelveDesdeBD(conAri, "sum(cantlote)", "olicoupagelinlotes", "codigo", Data1.Recordset!Codigo)
    If SQL <> CadenaConsulta Then
        MsgBox "Error en Kilos.  " & vbCrLf & SQL & "    /     " & CadenaConsulta, vbExclamation
        Exit Function
    End If
    
    
    
    'Junio 2014
    'Comprobaremos tb que el deposito esta vacio, y caben lo kilos a couopar
    Dim cD As cDeposito
    Set cD = New cDeposito
    
    SQL = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", Data1.Recordset!codartic)
    
    If cD.LeerDatos(CInt(Data1.Recordset!Deposito), False) Then
'        If cD.NUmlote <> "" Then
'            'MsgBox "El deposito no esta vacio.", vbExclamation
'        Else
            
            If (cD.Capacidad * CCur(SQL)) < ImporteFormateado(CadenaConsulta) Then
                SQL = vbCrLf & vbCrLf & "Depósito: " & cD.Capacidad & " litros" & vbCrLf & "Kilos: " & (cD.Capacidad * CCur(SQL))
                SQL = SQL & vbCrLf & "Coupage: " & CadenaConsulta
                
                MsgBox "Cantidad excede del capacidad del deposito:" & SQL, vbExclamation
                
            Else
                ComprobarNumeroLoteLineas = True
            End If
       ' End If
    End If
    
    
            
    'COTUBRE 2014
    If ComprobarNumeroLoteLineas Then
        CadenaConsulta = ""
        'Octubre 2014
        'Si el deposito no esta vacio, el deposito tiene que aparecer en las lineas
        SQL = DevuelveDesdeBD(conAri, "codartic", "proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote", "numdeposito", cD.NumDeposito)
        If SQL <> "" Then
            If SQL <> Text1(2).Text Then CadenaConsulta = CadenaConsulta & vbCrLf & "-Articulo diferente deposito / coupage"

            'Si esta no esta vacio, en las lineas tiene que estar este deposito
            If SQL <> "" Then
                FinDeposito = "fincuba"
                SQL = "codigo = " & Data1.Recordset!Codigo & " AND deposito "
                SQL = DevuelveDesdeBD(conAri, "linea", "olicoupagelinlotes", SQL, cD.NumDeposito, "N", FinDeposito)
                If SQL = "" Then
                    CadenaConsulta = CadenaConsulta & vbCrLf & "-Deposito NO vacio. En las lineas deberia aparecer el deposito destino del cupage"
                Else
                    'OK. El deposito esta marcado destino y en las lineas.
                    'En las lineas tiene que aparecer con la marca findeposito (fincuba)
                    If FinDeposito = "0" Then CadenaConsulta = CadenaConsulta & vbCrLf & "-Aparece el deposito en origen/destino." & vbCrLf & "Debe tener la marca de fin deposito"
                    
                End If
            End If
        End If
        
        If CadenaConsulta <> "" Then
            ComprobarNumeroLoteLineas = False
            MsgBox CadenaConsulta, vbExclamation
            
        End If
    End If
    
    
    Set cD = Nothing
End Function



Private Sub ObtenerLineCoupage(Directo As Boolean)
Dim cad As String
Dim Depo As Integer
Dim b As Boolean
    
        
        
        Screen.MousePointer = vbHourglass
        If Not Directo Then
            Set frmB2 = New frmBuscaGrid
            'CAMPOS
            'numdeposito,nomartic,spartidas.codartic,spartidas.numlote,litros
            cad = "Deposito|proddepositos|numdeposito|N||5·"
            cad = cad & "Cod. art|spartidas|codartic|T||20·"
            cad = cad & "Articulo|sartic|nomartic|T||45·"
            cad = cad & "Lote|spartidas|numlote|T||12·"
            cad = cad & "Kilos||kilos|N|" & FormatoPrecio & "|16·"
            frmB2.vCampos = cad
            'TABLA
            cad = " proddepositos left join spartidas on spartidas.numlote=proddepositos.numlote"
            cad = cad & " inner join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
            frmB2.vTabla = cad
            'WHERE
            frmB2.vSQL = "not spartidas.numlote is null "  'and DepositoVtaDirecta = 0"
            HaDevueltoDatos = False
            frmB2.vDevuelve = "0|"
            frmB2.vTitulo = "Depositos"
            frmB2.vselElem = 0
            frmB2.vConexionGrid = conAri 'Conexión a BD: Ariges
            CadenaConsulta = ""
            frmB2.Show vbModal
            Set frmB2 = Nothing
        Else
            CadenaConsulta = "18|"
        End If
        If CadenaConsulta <> "" Then
        
            'Un par de comprobaciones
            'El deposito NO esta para este coupage
            CadenaConsulta = RecuperaValor(CadenaConsulta, 1)
            cad = "codigo = " & Text1(0).Text & " AND deposito"
            cad = DevuelveDesdeBD(conAri, "deposito", "olicoupagelinlotes", cad, CadenaConsulta)
        
            If cad <> "" Then
                MsgBox "Ya esta este deposito asignado al coupage actual", vbExclamation
            Else
        
        
                'Ha seleccionado el deposito. Ahora, con el deposito haremos la insercino
                'tanto en olicoupagelin   olicoupagelinlotes
                'teniendo en cuenta que si en la de lin, ya esta el articulo, no insert si no updatea la cantidad
                '
                Screen.MousePointer = vbHourglass
                Depo = CInt(Val(CadenaConsulta))
                conn.BeginTrans
                b = InsertarLineaCoupage(Depo)
                If Not b Then
                    conn.RollbackTrans
                    
                Else
                    conn.CommitTrans
                    Espera 0.2
                    'Ahora, situamos el datagrid en ese y pulsamos modificar
                    CargaGrid DataGrid1, Data2, True
                    CadenaConsulta = "deposito = " & Depo
                    If Not SituarData(Data2, CadenaConsulta, Me.lblIndicador) Then
                        PonerModo 2
                    Else
                        
                        BotonModificarLinea
                        EsModificarDeAñadir = True
                    End If
                    
                End If
                Screen.MousePointer = vbDefault
            End If
            
            
        End If

End Sub



Private Function InsertarLineaCoupage(Deposito As Integer) As Boolean
Dim cDe As cDeposito
Dim Aux As String
Dim Cantidad As Currency

    On Error GoTo eInsertarLineaCoupage
    InsertarLineaCoupage = False
    
    Set cDe = New cDeposito
    cDe.LeerDatos Deposito, False
    
    If cDe.NUmlote = "" Then Err.Raise 513, , "Error leyendo numero deposito"
    
    Set miRsAux = New ADODB.Recordset
    Aux = "select spartidas.*,factorconversion from spartidas inner join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
    Aux = Aux & " AND numlote = " & DBSet(cDe.NUmlote, "T")
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO pue ser eof
    Cantidad = Round(cDe.Kilos / miRsAux!FactorConversion, 2)
    If miRsAux!cantotal <> cDe.Kilos Then
        Aux = "Kilos partida: " & miRsAux!cantotal & vbCrLf
        Aux = Aux & "Deposito (L/Kg): " & Cantidad & " / " & cDe.Kilos
        Aux = "Diferencia cantidades" & vbCrLf & vbCrLf & Aux
        MsgBox Aux, vbInformation
    End If
    Dim C As cPartidas
     
    'YA TENGO ARTICULO - cantidad
    Cantidad = cDe.Kilos
    Aux = miRsAux!codartic
    miRsAux.Close
    
    
    'Veo si ya esta en olicoupagelin (podria existir YA)
    CadenaConsulta = "Select * from olicoupagelin where codigo = " & Text1(0).Text & " AND codartic=" & DBSet(Aux, "T")
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        'NIO existe, lo creo
        CadenaConsulta = "INSERT INTO olicoupagelin(codigo,codartic,kilos) VALUES (" & Text1(0).Text & ","
        CadenaConsulta = CadenaConsulta & DBSet(Aux, "T") & "," & DBSet(Cantidad, "N") & ")"
        
    Else
        'Si existe. Incremento la cantidad
        CadenaConsulta = "UPDATE olicoupagelin SET kilos=kilos + " & DBSet(Cantidad, "N")
        CadenaConsulta = CadenaConsulta & " WHERE codigo = " & Text1(0).Text & " AND codartic ="
        CadenaConsulta = CadenaConsulta & DBSet(Aux, "T")
    
    End If
    miRsAux.Close
    
    conn.Execute CadenaConsulta
    
    
    
    'En lineas de lote s de coupages
    CadenaConsulta = DevuelveDesdeBD(conAri, "max(linea)", "olicoupagelinlotes", "codigo", Text1(0).Text)
    NumRegElim = Val(CadenaConsulta) + 1
    CadenaConsulta = "INSERT INTO olicoupagelinlotes (codigo,codartic,linea,numlote,cantlote,fincuba,deposito) VALUES ("
    CadenaConsulta = CadenaConsulta & Text1(0).Text & "," & DBSet(Aux, "T") & "," & NumRegElim & ","
    CadenaConsulta = CadenaConsulta & DBSet(cDe.NUmlote, "T") & "," & DBSet(Cantidad, "N") & ",0," & cDe.NumDeposito & ")"
    conn.Execute CadenaConsulta
    
    
    CambiaKilosTotales 0, Cantidad
    
    
    InsertarLineaCoupage = True
    
eInsertarLineaCoupage:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set cDe = Nothing
    Set miRsAux = Nothing
End Function
