VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmRegListaRev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Articulos"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   ClipControls    =   0   'False
   Icon            =   "frmRegListaRev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
      Height          =   1455
      Left            =   9240
      TabIndex        =   28
      Top             =   5400
      Width           =   2415
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2280
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label5 
         Caption         =   "Salida"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   8
      Left            =   10200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   11
      Text            =   "numlinea"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "Operario|N|N|||smoval|codigope|000000|N|"
      Text            =   "codigope"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Importe|N|N|||smoval|impormov|#,###,###,##0.00|N|"
      Text            =   "importe"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   6
      Tag             =   "Cantidad|N|N|||smoval|cantidad|##,###,##0.00|N|"
      Text            =   "cantidad"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "hora"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   2505
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   2
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   6240
      Width           =   3975
   End
   Begin VB.ComboBox cboAux 
      Height          =   315
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Detalle Movimiento|T|N|||smoval|detamovi||N|"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   7
      Left            =   9120
      MaxLength       =   7
      TabIndex        =   10
      Tag             =   "Documento|T|N|||smoval|document||N|"
      Text            =   "documento"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   6
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   9
      Text            =   "letra ser"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   1200
      MaxLength       =   11
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||smoval|fechamov|dd/mm/yyyy|N|"
      Text            =   "fecha"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "Cod. Almacen|N|N|0|999|smoval|codalmac|000|N|"
      Text            =   "codalmac"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   960
      TabIndex        =   21
      ToolTipText     =   "Buscar almacen"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cboAux 
      Height          =   315
      Index           =   0
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo Movimiento|N|N|||smoval|tipomovi||N|"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Articulo|T|N|||smoval|codartic||N|"
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   3675
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   5910
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10515
      TabIndex        =   13
      Top             =   5910
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10515
      TabIndex        =   18
      Top             =   5910
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   480
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
      Left            =   9720
      Top             =   480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRegListaRev.frx":000C
      Height          =   4130
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7276
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Desc. Almacen"
      Height          =   255
      Left            =   2760
      TabIndex        =   27
      Top             =   6045
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente/Proveedor/Trabajador"
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   6045
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Cod. Art�culo"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1500
      Picture         =   "frmRegListaRev.frx":0021
      ToolTipText     =   "Buscar art�culo"
      Top             =   645
      Width           =   240
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
      TabIndex        =   16
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmRegListaRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmArtic As frmAlmArticulos  'Articulos
Attribute frmArtic.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer
Dim PrimeraVez As Boolean
Dim btnPrimero As Byte 'Variable que indica el N� del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim CadenaBusqueda As String 'Cadena para la consulta de de busqueda en Grid
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe
'---- Laura: 27/09/2006
'cadena para la SQL de los totales de cantida e importe por articulo mostrado
Dim cadSelGrid As String


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Private HaDevueltoDatos As Boolean


Private Sub cboAux_GotFocus(Index As Integer)
    With cboAux(Index)
        If Modo = 1 Then 'Modo 1: Busqueda
            .BackColor = vbYellow
        Else
            .BackColor = vbWhite
        End If
    End With
End Sub

Private Sub cboAux_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboAux_LostFocus(Index As Integer)
    If Modo = 1 Then cboAux(Index).BackColor = vbWhite
End Sub


Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    If Modo = 1 Then HacerBusqueda
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Imprimir()
Dim cad As String
Dim numParam As Byte

    'Resto parametros
    cad = ""
    cad = cad & "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
            
    With frmImprimir
        .NombreRPT = "rAlmMovim.rpt"
        .OtrosParametros = cad
        .NumeroParametros = numParam
        .FormulaSeleccion = cadSeleccion
        .EnvioEMail = False
        .Opcion = 9
        .Titulo = "Informe Movimientos Articulos"
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub cmdAux_Click()
'Abre Formulario de Mantenimiento de Almacenes Propios
    Set frmA = New frmAlmAlPropios
    frmA.DatosADevolverBusqueda = "0"
    frmA.Show vbModal
    Set frmA = Nothing
    PonerFoco txtAux(0)
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

   If Modo = 1 Then       'Buscar
        LimpiarCampos
        PonerModo 0
        CargaTxtAux False, False
    End If
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_DblClick()
'Abrir el formulario del Mantenimiento del que viene el Movimiento
'Se busca en hist�rico o en Form
Dim SQL As String

    Select Case Data2.Recordset!detamovi
        Case "TRA" 'traspaso de almacenes
            'Traspaso de Almacen
            With frmAlmTraspaso
                .EsHistorico = True
                .hcoCodMovim = Data2.Recordset!document
                .hcoFechaMovim = Data2.Recordset!Fechamov
                .Show vbModal
            End With
            
        Case "REG" 'Movimientos de Almacen
                    'Movimientos de Almacen
            With frmAlmMovimientos
                .EsHistorico = True
                .hcoCodMovim = Data2.Recordset!document
                .hcoFechaMovim = Data2.Recordset!Fechamov
                .Show vbModal
            End With

        Case "ALV", "ART", "ALM", "ALZ"
                                'ALV:Albaran de Venta (a clientes)
                                'ART: Albaran rectificativo
                                'ALM: ALbaran Mostrador
                                'ALZ: Albaranes "B"
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
            'si esta ya facturado abrir el hist�rico de facturas: frmFacHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", Data2.Recordset!detamovi, "T", , "numalbar", Data2.Recordset!document, "N")
            If SQL <> "" Then 'existe el Albaran
                 With frmFacEntAlbaranes
                    If EsNumerico(Data2.Recordset!document) Then
                        .hcoCodMovim = Format(Data2.Recordset!document, "0000000")
                    Else
                        .hcoCodMovim = Data2.Recordset!document
                    End If
                    .hcoCodTipoM = Data2.Recordset!detamovi
                    .RecuperarFactu = False
                    .Show vbModal
                End With
            Else 'No existe en albaran, abrir Historico Factura
                With frmFacHcoFacturas
                    If EsNumerico(Data2.Recordset!document) Then
                        .hcoCodMovim = Format(Data2.Recordset!document, "0000000")
                    Else
                        .hcoCodMovim = Data2.Recordset!document
                    End If
                    .hcoCodTipoM = Data2.Recordset!detamovi
                    .hcoFechaMov = Data2.Recordset!Fechamov
                    
                    .Show vbModal
                End With
            End If
            
        Case "ALR" 'Albaran de Reparacion (a clientes)
             With frmFacEntAlbaranes
                If EsNumerico(Data2.Recordset!document) Then
                    .hcoCodMovim = Format(Data2.Recordset!document, "0000000")
                Else
                    .hcoCodMovim = Data2.Recordset!document
                End If
                .hcoCodTipoM = Data2.Recordset!detamovi
                .RecuperarFactu = False
                .Show vbModal
            End With
            
        Case "ALC" 'Albaran de Compra (a Proveedores)
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmComEntAlbaranes
            'si esta ya facturado abrir el hist�rico de facturas: frmComHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalp
            SQL = DevuelveDesdeBDNew(conAri, "scaalp", "numalbar", "codprove", Data2.Recordset!codigope, "N", , "numalbar", Data2.Recordset!document, "T", "fechaalb", Data2.Recordset!Fechamov, "F")
            If SQL <> "" Then 'existe el Albaran
                With frmComEntAlbaranes
                    .hcoCodMovim = Data2.Recordset!document
                    .hcoFechaMovim = Data2.Recordset!Fechamov
                    .hcoCodProve = Data2.Recordset!codigope 'aqui es el proveedor
                    .Show vbModal
                End With
            Else        'No existe en albaran, abrir Historico Factura
                With frmComHcoFacturas
                    .hcoCodMovim = Data2.Recordset!document
                    .hcoFechaMovim = Data2.Recordset!Fechamov
                    .hcoCodProve = Data2.Recordset!codigope 'aqui es el proveedor
                    .Show vbModal
                End With
            End If
            
            
        '**********************************
        'Laura: modificado 11/09/06
'        Case "FTI" 'Factura Ticket de venta
        Case "ATI" 'Albaran Ticket de venta
        '**********************************
            'Abrir el historico de facturas
             With frmFacHcoFacturas
                If EsNumerico(Data2.Recordset!document) Then
                    .hcoCodMovim = Format(Data2.Recordset!document, "0000000")
                Else
                    .hcoCodMovim = Data2.Recordset!document
                End If
                .hcoCodTipoM = Data2.Recordset!detamovi
                .hcoFechaMov = Data2.Recordset!Fechamov
                .Show vbModal
                
            End With
            
        'Nuevo Marzo 2009
        'PRoduccion y coupage  ... para aceite morales
        Case "PRO"
            With frmProdOrden
                .DatosADevolverBusqueda2 = Data2.Recordset!document
                .Show vbModal
            End With
                    
            
            
        Case "CUP"
            With frmAlmCoupage
                .DatosADevolverBusqueda2 = Data2.Recordset!document
                .Show vbModal
            End With
    End Select
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Codigo As Long
Dim movim As String

    If Not Data2.Recordset.EOF Then
        'Poner descripcion del almacen
        Text2(1).Text = Data2.Recordset.Fields(2).Value
        
        'Poner descripcion del Cliente/Proveedor
        Codigo = Data2.Recordset!codigope
        movim = Data2.Recordset!detamovi
        Text2(2).Text = PonerNombreCliente(Codigo, movim)
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
   
    'ICONOS de La toolbar
    btnPrimero = 8 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 16 'Imprimir
        .Buttons(6).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    NombreTabla = "smoval"
    Ordenacion = " ORDER BY codartic," & NombreTabla & ".codalmac, fechamov desc, horamovi "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = -1"
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    PonerCampos
    PonerModo 0
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim SQL As String

    On Error GoTo ECarga

    B = DataGrid1.Enabled
     
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    DataGrid1.Columns(0).visible = False 'Cod. Artic
    DataGrid1.Columns(2).visible = False 'Nombre Almacen
    
    'Cod. Almac
    DataGrid1.Columns(1).Caption = "Almacen"
    DataGrid1.Columns(1).Width = 900
    DataGrid1.Columns(1).NumberFormat = "000"
    
    'Fecha Mov
    DataGrid1.Columns(3).Caption = "Fecha"
    DataGrid1.Columns(3).Width = 1050
    
    'Hora Movim
    DataGrid1.Columns(4).Caption = "Hora"
    DataGrid1.Columns(4).Width = 850
    DataGrid1.Columns(4).NumberFormat = "hh:mm:ss"
    
    'Tipo Movim
    DataGrid1.Columns(5).Caption = "Tipo"
    DataGrid1.Columns(5).Width = 600
    DataGrid1.Columns(5).Alignment = dbgCenter
    
    'Detalle Movim
    DataGrid1.Columns(6).Caption = "Detalle"
    DataGrid1.Columns(6).Width = 700
    DataGrid1.Columns(6).Alignment = dbgCenter
    
    'Cantidad
    DataGrid1.Columns(7).Caption = "Cantidad"
    DataGrid1.Columns(7).Width = 1400
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(7).NumberFormat = FormatoCantidad
    
    'Importe Movimiento
    DataGrid1.Columns(8).Caption = "Importe"
    DataGrid1.Columns(8).Width = 1400
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(8).NumberFormat = FormatoImporte
    
    
    'Cod. Cliente/Proveedor
    DataGrid1.Columns(9).Caption = "Cli./Prov."
    DataGrid1.Columns(9).Width = 900
    DataGrid1.Columns(9).Alignment = dbgCenter
    DataGrid1.Columns(9).NumberFormat = "000000"
    
    'Letra Serie
    DataGrid1.Columns(10).Caption = "Letra"
    DataGrid1.Columns(10).Width = 600
       
    'N� Documento
    DataGrid1.Columns(11).Caption = "N� Documento"
    DataGrid1.Columns(11).Width = 1400
    DataGrid1.Columns(11).Alignment = dbgCenter
    DataGrid1.Columns(11).NumberFormat = "0000000"
    
    'N� Linea
    DataGrid1.Columns(12).Caption = "N� Linea"
    DataGrid1.Columns(12).Width = 900
    DataGrid1.Columns(12).Alignment = dbgCenter
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Enabled = B
    If Modo = 2 Then DataGrid1.Enabled = True
    PrimeraVez = False
    
    CalcularTotales
    
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim i As Byte
Dim alto As Single

     'Los ponemos Visibles o No
    '--------------------------
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = visible
    Next i
    cmdAux.visible = visible
    cboAux(0).visible = visible
    cboAux(1).visible = visible


    

    If Not visible Then
        alto = 280
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
        Next i
        Me.cmdAux.Top = alto
        Me.cboAux(0).Top = alto
        Me.cboAux(1).Top = alto
    Else
        DeseleccionaGrid Me.DataGrid1
        CargarComboAux
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                txtAux(i).BackColor = vbWhite
                If (i = 0 Or i = 1 Or i = 3 Or i = 4 Or i = 5 Or i = 7) Then BloquearTxt txtAux(i), False 'TxtAux(i).Locked = False
            Next i
            cmdAux.Enabled = True
            cboAux(0).Enabled = True
            cboAux(0).ListIndex = -1
            cboAux(0).BackColor = vbWhite
            cboAux(1).Enabled = True
            cboAux(1).ListIndex = -1
            cboAux(1).BackColor = vbWhite
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 210
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posici�n Top
        '-------------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        cboAux(0).Top = alto
        cboAux(1).Top = alto
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux(0).Left = DataGrid1.Left + 340 'codalmac
        txtAux(0).Width = DataGrid1.Columns(1).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width  'fechamov
        txtAux(1).Width = DataGrid1.Columns(3).Width - 35
        i = 2 'hora mov
        txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 25
        txtAux(i).Width = DataGrid1.Columns(4).Width - 20
        'Tipo Movimiento
        cboAux(0).Left = txtAux(2).Left + txtAux(2).Width + 5
        cboAux(0).Width = DataGrid1.Columns(5).Width
        'Detalle Movimiento
        cboAux(1).Left = cboAux(0).Left + cboAux(0).Width
        cboAux(1).Width = DataGrid1.Columns(6).Width
        
        i = 3 'Cantidad
        txtAux(i).Left = cboAux(1).Left + cboAux(1).Width
        txtAux(i).Width = DataGrid1.Columns(7).Width - 25
        
        For i = 4 To txtAux.Count - 1
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 25
            txtAux(i).Width = DataGrid1.Columns(i + 4).Width - 25
        Next i
    End If

    

'    'Los ponemos Visibles o No
'    '--------------------------
'    For I = 0 To txtAux.Count - 1
'        txtAux(I).visible = visible
'    Next I
'    cmdAux.visible = visible
'    cboAux(0).visible = visible
'    cboAux(1).visible = visible
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArtic_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass

        cadB = ""
        cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
        PonerCadenaBusqueda
        
        cadB = RecuperaValor(CadenaDevuelta, 1)
        cadSeleccion = "{smoval.codartic}=""" & cadB & """"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    'Codigo Articulos
    If Index = 0 Then
        Set frmArtic = New frmAlmArticulos
        frmArtic.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
        frmArtic.Show vbModal
        Set frmArtic = Nothing
    End If
    PonerFoco Text1(0)
    Screen.MousePointer = vbDefault
End Sub









Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    If Trim(Text1(Index).Text) = "" Then
        Text2(Index).Text = ""
        Exit Sub
    ElseIf (Modo = 1) Then 'Busqueda
        Text2(0).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
    End If
End Sub




Private Sub txtAux_GotFocus(Index As Integer)
    If (Modo = 1 And (Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Or Index = 7)) Or (Modo <> 1) Then
        ConseguirFoco txtAux(Index), Modo
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
        
    Select Case Index
        Case 0 'cod. almacen
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(1).Text = PonerNombreDeCod(txtAux(Index), conAri, "salmpr", "nomalmac")
            Else
                Text2(1).Text = ""
            End If

        Case 1 'Fecha Movimiento
             If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
             
        Case 3 'cantidad
            PonerFormatoDecimal txtAux(Index), 3
        
        Case 4 'importe
            PonerFormatoDecimal txtAux(Index), 1
            
        Case 5 'Cliente/proveedor/trabajador
            If PonerFormatoEntero(txtAux(Index)) Then FormateaCampo txtAux(Index)
            
        Case 8
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 2 'Ver Todos
            BotonVerTodos
        Case 5 'Imprimir
            Imprimir
        Case 6  'Salir
            Unload Me
        Case 8 To 11 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg

   'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    Select Case Kmodo
    Case 0    'Modo Inicial
        Toolbar1.Buttons(5).Enabled = False 'Imprimir
        PonerBotonCabecera True
    Case 1 'Modo Buscar
        lblIndicador.Caption = "BUSQUEDA"
        Toolbar1.Buttons(5).Enabled = False 'Imprimir
        PonerBotonCabecera False
        PonerFoco Text1(0)
        
    Case 2    'Preparamos para que pueda Modificar
        PonerBotonCabecera True
    End Select
           
    B = Modo <> 0 And Modo <> 2
  
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i

    
    PonerLongCampos

    B = (Kmodo >= 3) Or Modo = 1
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub



Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    CargaGrid True
'    CalcularTotales
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim selSQL As String
Dim cadBuscar2 As String
Dim i As Integer

    cadSelGrid = ""

    selSQL = "SELECT smoval.codartic, smoval.codalmac, nomalmac, fechamov, horamovi, if(smoval.tipomovi=0,""S"",""E"") as tipomovi, detamovi, "
    selSQL = selSQL & "cantidad, impormov, codigope, letraser, document, numlinea "
    
    SQL = " FROM (smoval LEFT OUTER JOIN salmpr on smoval.codalmac=salmpr.codalmac)"
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            'LAura: 29/09/06
'            If Data1.Recordset.RecordCount > 1 Then
            'Si devuelve + de 1 registro en el DataGrid poner la info del primer articulo
                'quitar codartic de la cadena busqueda
'                i = InStr(CadenaBusqueda, "(smoval.codartic")
'                If i > 0 Then
'
'                End If
                
                SQL = SQL & CadenaBusqueda & " AND codartic=" & DBSet(Text1(0).Text, "T")
'            Else
'                SQL = SQL & CadenaBusqueda
'            End If
        Else
            SQL = SQL & " WHERE codartic = " & DBSet(Text1(0).Text, "T")
        End If
    Else
        SQL = SQL & " WHERE codartic = '-1'"
    End If
    
    If Not vUsu.TrabajadorB Then SQL = SQL & " AND smoval.codalmac <> " & vParamAplic.AlmacenB
    
    SQL = SQL & " " & Ordenacion & " DESC "
    '---- Laura: 27/09/2006
    cadSelGrid = SQL
    SQL = selSQL & SQL
    '----
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        CargaTxtAux True, True
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
Dim cadB As String

'Ver todos
    EsBusqueda = False
    
    cadB = ""
    If Not vUsu.TrabajadorB Then cadB = " codalmac <> " & vParamAplic.AlmacenB
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia cadB
        CargaGrid True
    Else
        If cadB <> "" Then cadB = " WHERE " & cadB
        CadenaConsulta = "Select codartic from " & NombreTabla & cadB & " group by codartic " & Ordenacion
        PonerCadenaBusqueda
        Toolbar1.Buttons(5).Enabled = False 'Imprimir
    End If
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
Dim bol As Boolean

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
    
    bol = (Modo = 1 Or Modo = 2)
    Me.Label3.visible = bol
    Me.Text2(1).visible = bol
    
    bol = (Modo = 2)
    Me.Label2.visible = bol
    Me.Text2(2).visible = bol
    
    '---- Laura: 27/09/2006
    'Total cantidad
    Me.Frame2.visible = bol
    Me.Label4.visible = bol
    Me.Text2(3).visible = bol
    'Total importe
    Me.Label5.visible = bol
    Me.Text2(4).visible = bol
    '----
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadB2 As String

    cadB = ObtenerBusqueda(Me, False)
'    If Me.Text1(0).Text <> "" Then
'        If cadB <> "" Then cadB = cadB & " AND "
'        cadB = cadB & "(codartic LIKE " & DBSet(Text1(0).Text, "T") & ")"
'    End If
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

'    If chkVistaPrevia = 1 Then
'        MandaBusquedaPrevia cadB
'    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            If Not vUsu.TrabajadorB Then
                cadB = cadB & " AND codalmac <> " & vParamAplic.AlmacenB
                cadSeleccion = cadSeleccion & " AND ({smoval.codalmac} <> " & vParamAplic.AlmacenB & ")"
            End If
        
            'Cadena para el Data1
            CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
            'Cadena para el Datagrid y el Data2
            'el codartic no se incluye en la cadB de las lineas pq siempre
            'se muestran las de un codartic concreto
            Text1(0).Text = ""
            cadB2 = ObtenerBusqueda(Me, False)
'            CadenaBusqueda = ""
            If cadB2 <> "" Then 'Para cargar la consulta del CargaGrid
                CadenaBusqueda = " WHERE " & cadB2
            Else
                CadenaBusqueda = ""
            End If
            
        Else
            'obtener todos los articulos
            CadenaConsulta = "select codartic from " & NombreTabla & " GROUP BY codartic " & Ordenacion
            CadenaBusqueda = ""
        End If
        PonerCadenaBusqueda
'    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim i As Byte
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta

    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de b�squeda", vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        'Limpiar los Campos Auxiliares
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
        Text2(1).Text = ""
        Me.cboAux(0).ListIndex = -1
        Me.cboAux(1).ListIndex = -1
        Exit Sub
    Else
        PonerModo 2
        Toolbar1.Buttons(5).Enabled = True 'Imprimir
        CargaTxtAux False, False
        PonerCampos
        CargaGrid True
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
    Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sartic", "nomartic")
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

Private Sub CargarComboAux()
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Entrada, 1-Salida
Dim Index As Byte, i As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
On Error GoTo ECargar

        Index = 0 'Combo Tipo Movimiento
        cboAux(Index).Clear
        cboAux(Index).AddItem "S"
        cboAux(Index).ItemData(cboAux(Index).NewIndex) = 0

        cboAux(Index).AddItem "E"
        cboAux(Index).ItemData(cboAux(Index).NewIndex) = 1
        
        Index = 1 'Combo Detalle Movimiento
        SQL = "select codtipom,nomtipom from stipom"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        i = 0
        cboAux(Index).Clear
        While Not RS.EOF
            cboAux(Index).AddItem RS.Fields(0).Value
            cboAux(Index).ItemData(cboAux(Index).NewIndex) = i
            i = i + 1
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
ECargar:
    If Err.Number <> 0 Then
        RS.Close
        Set RS = Nothing
        MuestraError Err.Number, "Cargando Combobox", Err.Description
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
            
    cad = cad & "C�digo|smoval|codartic|T||25�Denominacion|sartic|nomartic|T||70�"
    Tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ") "
    If cadB <> "" Then Tabla = Tabla & " WHERE " & cadB
    
    Tabla = Tabla & " GROUP BY smoval.codartic "
    Titulo = "Movimientos de Articulos"

           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = ""
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
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
            PonerFoco Text1(kCampo)
            Toolbar1.Buttons(5).Enabled = True 'Imprimir
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Function PonerNombreCliente(Codigo As Long, movim As String) As String
'Devuelve el nombre del Trabajador/Cliente/Proveedor para ponerlo en la caja de texto text2 en la parte inferior del form
Dim Nombre As String

    Select Case movim
        Case "TRA", "REG", "DFI"
            'Obtener nombre de la tabla de trabajadores
            Nombre = DevuelveDesdeBDNew(conAri, "straba", "nomtraba", "codtraba", CStr(Codigo), "N")
            Label2.Caption = "Trabajador"
        Case "ALV", "ALR", "ALM", "ART", "FAV", "FTI", "ATI"
            'Obtener nombre de la tabla de Clientes
            Nombre = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", CStr(Codigo), "N")
            Label2.Caption = "Cliente"
        Case "ALC"
            'Obtener el nombre de la tabla de Proveedores
            Nombre = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(Codigo), "N")
            Label2.Caption = "Proveedor"
    End Select
    PonerNombreCliente = Nombre
End Function



Private Sub CalcularTotales()
'calcula la cantidad total y el importe total para los
'registros mostrados de cada art�culo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Im As Currency
Dim Tot As Currency

    On Error GoTo ErrTotales
    If cadSelGrid = "" Then Exit Sub
    
    
    NumRegElim = InStr(1, cadSelGrid, "ORDER BY")
    If NumRegElim > 0 Then
        SQL = Mid(cadSelGrid, 1, NumRegElim - 1)
    Else
        SQL = cadSelGrid
    End If
    SQL = "SELECT tipomovi,sum(cantidad) as totCantidad " & SQL & " GROUP BY tipomovi"
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text2(3).Text = "": Text2(4).Text = ""
    Tot = 0
    While Not RS.EOF
            
        Im = DBLet(RS!totCantidad, "N")
        If Val(RS!tipomovi) = 1 Then
            Tot = Tot + Im
            Text2(3).Text = Format(Im, FormatoCantidad)
        Else
            Text2(4).Text = Format(Im, FormatoCantidad)
            Tot = Tot - Im
        End If
        RS.MoveNext
    Wend
    Text2(5).Text = Format(Tot, FormatoCantidad)
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ErrTotales:
    MuestraError Err.Number, "Calcular totales.", Err.Description
End Sub
