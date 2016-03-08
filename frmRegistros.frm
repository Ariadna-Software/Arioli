VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegistros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registros"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12030
   ClipControls    =   0   'False
   Icon            =   "frmRegistros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRegistros.frx":000C
      Left            =   4080
      List            =   "frmRegistros.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "Cantidad periodo|N|N|||sregistros|Perioricidad|||"
      Top             =   1695
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Tag             =   "rptPropio|T|S|||sregistros|ReportPropio|||"
      Text            =   "ReportPropio"
      Top             =   1695
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   2835
      Index           =   7
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "T|T|S|||sregistros|TextoAOfertar|||"
      Text            =   "frmRegistros.frx":004C
      Top             =   2520
      Width           =   7860
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   6000
      TabIndex        =   2
      Tag             =   "Aviso(dias)|N|S|||sregistros|DiasAviso|||"
      Text            =   "DiasAviso"
      Top             =   840
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1680
      TabIndex        =   4
      Tag             =   "Ultimo doc realizado|F|S|||sregistros|UltimoRealizado|dd/mm/yyyy||"
      Text            =   "UltimoRealizado"
      Top             =   1695
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Tag             =   "Cantidad periodo|N|S|||sregistros|NumPeriodo|||"
      Text            =   "NumPeriodo"
      Top             =   1695
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Tag             =   "1ºFecha|F|N|||sregistros|PrimeraFecha|dd/mm/yyyy||"
      Text            =   "PrimeraFecha"
      Top             =   1695
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||sregistros|Descripcion|||"
      Text            =   "Descripcion"
      Top             =   855
      Width           =   4140
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "idRegistro|N|N|||sregistros|idRegistro||S|"
      Text            =   "idRegistro"
      Top             =   855
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10680
      TabIndex        =   10
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10755
      TabIndex        =   11
      Top             =   5610
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   5440
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
         TabIndex        =   17
         Top             =   180
         Width           =   2115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir DOC"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRegistros.frx":005A
      Height          =   4575
      Left            =   8280
      TabIndex        =   12
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   4560
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
   Begin VB.Label lblDescCampo 
      Caption         =   "Histórico"
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   26
      Top             =   600
      Width           =   1260
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   5
      Left            =   2400
      Picture         =   "frmRegistros.frx":006F
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   960
      Picture         =   "frmRegistros.frx":05F9
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Informe"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   25
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Texto ofertado"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Aviso(dias)"
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   23
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Ultimo"
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   22
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Periodicidad"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   21
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "1ºFecha"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   19
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "idRegistro"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   1260
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
      TabIndex        =   14
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
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Lineas"
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
Attribute VB_Name = "frmRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Public CadenaSituarData As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean






Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then PosicionarData
                'Ponemos la cadena consulta
        End If
        
    Case 4 'MODIFICAR
        If DatosOk Then
             If ModificaDesdeFormulario(Me, 1) Then
                 DesBloqueaRegistroForm Text1(0)
                 PosicionarData
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
            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
        DesBloqueaRegistroForm Text1(0)
        PonerBotonCabecera False
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If Not (Modo = 2 Or Modo = 5) Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    
    
    AccionesLineas 2, Modo = 2
    
End Sub

'Private Sub cmdRegresar_Click()
''Este es el boton Cabecera
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
'    If CadenaSituarData <> "" Then
'        CadenaSituarData = ""
'        PonerModo 2
'        PonerCampos
'
'    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    

    
    
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
        .Buttons(10).Image = 10 'Imprimir
        
        .Buttons(13).Image = 16 'Eliminar
        .Buttons(14).Image = 40 'Imprimir
        
        .Buttons(16).Image = 15 'Salir
        
        
        
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "sregistros" 'Tabla Precios Especiales de Articulos
    Ordenacion = " ORDER BY idregistro"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE "
    
    
        CadenaConsulta = CadenaConsulta & " idregistro= -1" 'No recupera datos
  
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
        PonerModo 0
        CargaGrid (Modo = 2)

    
    'Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean

Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    


    DataGrid1.Columns(0).visible = False

    
    'Fecha Cambio
    DataGrid1.Columns(1).Caption = "Fecha "
    DataGrid1.Columns(1).Width = 1300
    DataGrid1.Columns(1).NumberFormat = "dd/mm/yyyy"
    DataGrid1.Columns(1).AllowSizing = False
    'Precio Unidad
    DataGrid1.Columns(2).Caption = "Firmado"
    DataGrid1.Columns(2).Width = 1600
    DataGrid1.Columns(2).AllowSizing = False
    
    DataGrid1.Enabled = b
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
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

    Text1(CInt(imgFecha(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub







Private Sub imgFecha_Click(Index As Integer)


   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass

   Set frmF = New frmCal
   frmF.Fecha = Now
   imgFecha(2).Tag = Index

   
   PonerFormatoFecha Text1(Index)
   If Text1(Index).Text <> "" Then frmF.Fecha = CDate(Text1(Index).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Index)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnLineas_Click()
    If BloqueaRegistroForm(Me) Then BotonMtoLineas
End Sub

Private Sub mnModificar_Click()

    
    If Modo = 5 Then
         AccionesLineas 2, False
    Else
        If BloqueaRegistroForm(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then
        'Añadir linea
        AccionesLineas 1, False
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
  If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
          
    
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 2, 5
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 0, 6, 4
            If Not PonerFormatoEntero(Text1(Index)) Then
                If Text1(Index).Text <> "" Then Text1(Index).Text = ""
            End If
            
        
    End Select
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
        
        Case 10
            'LINEAS
            mnLineas_Click
        
        Case 13
            LlamaImprimirGral "", "", 0, "morListReg.rpt", "Listado registros"
            
        Case 14
            If Data1.Recordset.EOF Then Exit Sub
            If Data2.Recordset.EOF Then Exit Sub
            
            CadenaConsulta = "{sregistros.idRegistro}=" & Text1(0).Text & " AND {sregistrosl.secuencial} = " & Data2.Recordset!secuencial
            
            CadenaDesdeOtroForm = "MorRegistro.rpt"
            If Text1(3).Text <> "" Then CadenaDesdeOtroForm = Text1(3).Text
            LlamaImprimirGral CadenaConsulta, "", 0, CadenaDesdeOtroForm, "Registro: " & Text1(0).Text & " - " & Data2.Recordset!Fecha
            
            CadenaConsulta = ""
            CadenaDesdeOtroForm = ""
        Case 16  'Salir
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



Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

     b = (Modo = 2) Or (Modo = 5)
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
    Toolbar1.Buttons(10).Enabled = b
    mnLineas.Enabled = b
    '===============================
    b = (Modo >= 3)
    'Insertar

    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    
    
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
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
Dim tabla As String
    
    tabla = "sregistrosl"
    SQL = "SELECT secuencial,fecha,if(firmado=0,""NO"",""si"") FROM " & tabla
    If enlaza Then
        SQL = SQL & " WHERE sregistrosl.idregistro=" & Data1.Recordset!idRegistro
    Else
        SQL = SQL & " WHERE idregistro = -1"
    End If
    SQL = SQL & " ORDER BY " & tabla & ".fecha desc"
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
    Combo1.ListIndex = 0
    Text1(4).Text = "0"
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
    
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    Text1(0).Text = SugerirCodigoSiguienteStr("sregistros", "idregistro")
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    PonerFoco Text1(2)
End Sub

Private Sub BotonEliminar()
    If Modo = 5 Then
        AccionesLineas 3, False
    Else
        BotonEliminarC
    End If
End Sub

Private Sub BotonEliminarC()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
  
    
  
    SQL = "Registros." & vbCrLf
    SQL = SQL & "--------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar registro:"
    SQL = SQL & vbCrLf & "Codigo : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Descr. : " & Text1(1).Text
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If Not Data2.Recordset.EOF Then
        SQL = "Tiene HISTORICO de registros. Va a eliminarlos tambien."
        SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
         If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    
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

   
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            'MsgBox Err.Number & " : " & Err.Description, vbExclamation
            MuestraError Err.Number, "Eliminar Precio Especial", Err.Description
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
    
        SQL = " WHERE idRegistro=" & Val(Data1.Recordset!idRegistro)
        
        'Lineas  sregistrosl
        Conn.Execute "Delete  from sregistrosl " & SQL
        
        'Cabeceras sregistros
        Conn.Execute "Delete  from sregistros " & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
On Error Resume Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    

    

    
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 10, "Cliente")
    cad = cad & "Nombre Cliente|sclien|nomclien|T||36·"
    cad = cad & ParaGrid(Text1(1), 15, "Cod. Artic")
    cad = cad & "Desc. Artic|sartic|nomartic|T||38·"
    
    tabla = "(" & NombreTabla & " LEFT JOIN sclien ON " & NombreTabla & ".codclien=sclien.codclien" & ")"
    tabla = tabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic"
    
    Titulo = "Precios Especiales"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
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

    

    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub









Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim SQL As String

    On Error Resume Next

    SQL = "UPDATE " & NombreTabla & " SET precioac=precionu, precioa1=precion1, dtoespec=dtoespe1, fechanue=null, precionu=0, precion1=0"
    SQL = SQL & " WHERE codclien=" & Data1.Recordset!CodClien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
    Conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
    Else
        ModificarCabecera = True
    End If
End Function


'Private Function InsertarLineasHistorico() As Boolean
'Dim SQL As String
'Dim NumF As String
'On Error Resume Next
'
'    'Obtenemos la siguiente numero de linea de tarifa
'    SQL = "codclien=" & Data1.Recordset!CodClien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
'    NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", SQL)
'
'    SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec)"
'    SQL = SQL & " VALUES (" & Data1.Recordset.Fields(0).Value & ", " & DBSet(Data1.Recordset.Fields(1).Value, "T") & ", "
'    SQL = SQL & NumF & ", " & DBSet(Text1(4).Text, "F") & ", "
'    SQL = SQL & DBSet(Data1.Recordset!precioac, "N") & ", " & DBSet(Data1.Recordset!precioa1, "N") & ", "
'    SQL = SQL & DBSet(Data1.Recordset!dtoespec, "N") & ") "
'    Conn.Execute SQL
'
'    If Err.Number <> 0 Then
'        'Hay error , almacenamos y salimos
'        InsertarLineasHistorico = False
'    Else
'        InsertarLineasHistorico = True
'    End If
'End Function


Private Sub BotonImprimir()
        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(idregistro=" & Text1(0).Text & ")"
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub



Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

     

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)

        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
  
  
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
    
    BloquearCmb Combo1, b
           
    '==============================
    b = Modo <> 0 And Modo <> 2
    
    
    Me.imgFecha(5).Enabled = b
    Me.imgFecha(2).Enabled = b
    
    
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    BloquearCmb Combo1, Not b
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    
    
    PonerModoOpcionesMenu    'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonMtoLineas()
        If Data1.Recordset.EOF Then Exit Sub
        
        PonerModo 5
        PonerBotonCabecera True
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
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas "
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


'1nue  2mod
Private Sub AccionesLineas(Accion As Byte, SoloVer As Boolean)
Dim F As Date
Dim Inc As Integer

    If Accion = 3 Then
        'Eliminar
        If Data2.Recordset.EOF Then Exit Sub
        CadenaDesdeOtroForm = "Va a eliminar del histórico:" & vbCrLf & "Fecha " & Data2.Recordset!Fecha
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "Firmado: " & Data2.Recordset.Fields(2)
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        CadenaDesdeOtroForm = "Delete from sregistrosl where idRegistro=" & Text1(0).Text & " AND secuencial =" & Data2.Recordset!secuencial
        If Not EjecutaSQL(conAri, CadenaDesdeOtroForm) Then Exit Sub
        
        'Veo cual es el ultimo ejecutado
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(fecha)", "sregistrosl", "idRegistro", Text1(0).Text)
        If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Format(CadenaDesdeOtroForm, FormatoFecha)
        
        If Text1(5).Text <> CadenaDesdeOtroForm Then
            Text1(5).Text = CadenaDesdeOtroForm
            'ha cambiado
            CadenaDesdeOtroForm = "UPDATE sregistro SET UltimoRealizado = " & DBSet(CadenaDesdeOtroForm, "F", "S")
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE idRegistro = " & Text1(0).Text & ""
            EjecutaSQL conAri, CadenaDesdeOtroForm
        End If
    Else
        If Accion = 2 Then
            If Data2.Recordset.EOF Then Exit Sub
            CadenaDesdeOtroForm = ""
            
            
            
        Else
            'Enviare la fecha prevista
            If Text1(5).Text = "" Then
                'Aun no ha hehco NINGUNO
                F = CDate(Text1(2).Text)
            Else
                'Ya tiene alguno hecho
                F = CDate(Text1(5).Text)
            End If
            Inc = CInt(Val(Text1(4).Text))
            If Inc > 0 Then
                'Periodicidad
                Select Case Combo1.ListIndex
                Case 2
                    'Semanas
                    CadenaDesdeOtroForm = "ww"
                Case 3
                    'MEs
                    CadenaDesdeOtroForm = "m"
                Case 4
                    CadenaDesdeOtroForm = "yyyy"
                Case Else
                    'dias
                    CadenaDesdeOtroForm = "d"
                End Select
                F = DateAdd(CadenaDesdeOtroForm, Inc, F)
            Else
                F = Now
            End If
            'Los diez primeros son la fehca
            CadenaDesdeOtroForm = Format(F, "dd/mm/yyyy") & Text1(7).Text
        End If
        
        With frmRegistroL2
            .EsMantenimientoPreventivo = False
            .InsMod = Not SoloVer
            .Nregistro2 = Data1.Recordset!idRegistro
            If Accion = 1 Then
                .linea = 0
            Else
                .linea = Data2.Recordset!secuencial
            End If
            
            .Show vbModal
        
    
        End With
        If CadenaDesdeOtroForm = "" Then Exit Sub
    End If

    'Regfrescamos el datagrid
    If Not SituarData(Data1, "(idregistro=" & Text1(0).Text & ")", "") Then
        PonerModo 0
    Else
        'La fecha
        Text1(5).Text = DBLet(Data1.Recordset!UltimoRealizado, "F")
        CargaGrid True
    End If
    
    
End Sub




