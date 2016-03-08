VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacRutasEnvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de carga"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12255
   ClipControls    =   0   'False
   Icon            =   "frmFacRutasEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   120
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "Vechiculo|N|S|0||sRepartoC|codtraba||N|"
      Text            =   "Text1"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   5
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   1800
      Width           =   3465
   End
   Begin VB.ComboBox cboEstado 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Situ|N|N|0||sRepartoC|situacion|||"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox cboConductor 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   120
      MaxLength       =   40
      TabIndex        =   21
      Tag             =   "Observaciones|T|S|||sRepartoC|conductor|||"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text"
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   3000
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||sRepartoC|fecha|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Index           =   1
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "Observaciones|T|S|||sRepartoC|observaciones|||"
      Text            =   "frmFacRutasEnvio.frx":000C
      Top             =   3600
      Width           =   5175
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10875
      TabIndex        =   8
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10875
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   4800
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
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   2400
      Width           =   3465
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   120
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "Vechiculo|N|N|0||sRepartoC|vehiculo||N|"
      Text            =   "Text1"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "Cod|N|N|0||sRepartoC|id|0000|S|"
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
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
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Iniciar expedicion"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "impresion expedicion"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir albaranes expedicion"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Resumen expedición"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   8880
         TabIndex        =   15
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   5040
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   5640
      TabIndex        =   20
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   5115
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Albarán / Factura"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador expedicion"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "frmFacRutasEnvio.frx":0035
      ToolTipText     =   "Buscar grupo plantilla"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Estado"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Conductor"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   2640
      Picture         =   "frmFacRutasEnvio.frx":0137
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmFacRutasEnvio.frx":01C2
      ToolTipText     =   "Buscar grupo plantilla"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Vehiculo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Id"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   375
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
      TabIndex        =   11
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
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
Attribute VB_Name = "frmFacRutasEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVe As frmFacVehiculos
Attribute frmVe.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1

Dim NombreTabla As String



Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CadenaConsulta As String
'Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub cboConductor_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    If Data1.Recordset.EOF Then Data1.RecordSource = "Select * from " & NombreTabla & " WHERE id = " & Text1(0).Text
                    PosicionarData
                    'BotonMtoLineas
                    'BotonAnyadirLinea
                    mnLineas_Click
                End If
            End If
        Case 4 'MODIFICAR
               If DatosOk Then
                    If ModificaDesdeFormulario(Me, 1) Then
                        TerminaBloquear
                        PosicionarData
                    End If
                End If
        Case 5 'InsertarModificar linea
                'Actualizar el registro en la tabla de lineas 'slipla' (Plantillas)
'                If ModificaLineas = 1 Then 'INSERTAR lineas
'                    If InsertarLinea Then
'                        CargaGrid True
'                        BotonAnyadirLinea
'                    End If
'                ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
'                    If ModificarLinea Then
'                        TerminaBloquear
'                        ModificaLineas = 0
'                        PonerBotonCabecera True
'                        CargaGrid True
'                        LLamaLineas 10
'                    End If
'                End If
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
        Case 5 'Lineas
'            TerminaBloquear
'            If ModificaLineas = 1 Then 'INSERTAR
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
'            DataGrid1.Enabled = True
'            ModificaLineas = 0
'            PonerBotonCabecera True
'            LLamaLineas 10
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Lineas Ofertas
        
        PonerModo 2
        Me.lblIndicador.Caption = ""
    End If
End Sub




Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    'ICONOS de La toolbar
    btnAnyadir = 5
    btnPrimero = 23 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 10 'Mto Lineas
        .Buttons(11).Image = 16 'impriir
        
        .Buttons(13).Image = 47 'Mto Lineas
        .Buttons(14).Image = 48   'imprimir
        .Buttons(15).Image = 40   'imprimir albaranes asociados
        .Buttons(16).Image = 26   'resumen expedicion
        
        
        .Buttons(21).Image = 15 'Salir
        .Buttons(23).Image = 6 'Primero
        .Buttons(24).Image = 7 'Anterior
        .Buttons(25).Image = 8 'Siguiente
        .Buttons(26).Image = 9 'Ultimo
    End With
    
    Me.ListView1.SmallIcons = frmppal.imgListComun
    
    LimpiarCampos   'Limpia los campos TextBox

    'PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "sRepartoC" 'Tabla Cabecera Plantillas

    
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE id = -1" 'No recupera datos
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaComobo
    
    PonerModo 0
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
Dim RT As ADODB.Recordset

    On Error GoTo ECarga
    
    Screen.MousePointer = vbHourglass
    Set RT = New ADODB.Recordset
    ListView1.ListItems.Clear

    If enlaza Then
    
        Set miRsAux = New ADODB.Recordset
        
        'Albaranes *****************
        SQL = "select s.numalbar elnumero,s.codtipom,s.fechaalb lafecha,nomclien,a.codtipom tienealbaran from srepartol s left join scaalb a on s.codtipom=a.codtipom and"
        SQL = SQL & " s.numalbar=a.numalbar and s.fechaalb=a.fechaalb where id= " & Data1.Recordset!ID
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then CargaItems True
        miRsAux.Close
        
        'FActuras*****
        'Por si acaso el hay albaranes ya facturados o, sacamos ordenes de transporte algo antiugas
        
        
        SQL = "select a.codtipom,numfactu,fecfactu,s.numalbar from srepartol s left join scafac1 a on s.codtipom=a.codtipoa"
        SQL = SQL & " and s.numalbar=a.numalbar and s.fechaalb=a.fechaalb where id=" & Data1.Recordset!ID & " ORDER BY codtipom,numfactu"
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RT.EOF
            If Not IsNull(RT!Codtipom) Then  'Esta en la factura
                SQL = "select codtipom,numfactu elnumero,fecfactu lafecha,nomclien , " & RT!NumAlbar & " elAlbaran from scafac where codtipom = '" & RT!Codtipom
                SQL = SQL & "' AND numfactu = " & RT!NumFactu & " AND fecfactu = '" & Format(RT!FecFactu, FormatoFecha) & "'"
                 miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then CargaItems False
                miRsAux.Close
            End If
            
            'Siguiente
            RT.MoveNext
        Wend
        RT.Close
    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos lineas ", Err.Description
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaItems(Albaranes As Boolean)
Dim It As ListItem
Dim Insertar As Boolean

    While Not miRsAux.EOF
        Insertar = True
        If Albaranes Then
            If IsNull(miRsAux!tienealbaran) Then Insertar = False
        Else
            Insertar = True
        End If
        If Insertar Then
            Set It = ListView1.ListItems.Add()
            It.Text = miRsAux!Codtipom
            It.SubItems(1) = miRsAux!elnumero
            It.SubItems(2) = miRsAux!LaFecha
            It.SubItems(3) = miRsAux!nomclien
            If Albaranes Then
                It.SmallIcon = 43
                
            Else
                It.SmallIcon = 44
                It.Tag = miRsAux!elalbaran
            End If
        End If
        miRsAux.MoveNext
    Wend
End Sub


'Private Sub LLamaLineas(alto As Single)
''Pone posicion TOP y LEFT de los controles en el form
'Dim jj As Integer
'Dim b As Boolean
'
'
'    DeseleccionaGrid Me.DataGrid1
'
'    'Fijamos el ancho
'    b = (Modo = 5 And ModificaLineas = 1 Or ModificaLineas = 2)
'
'    For jj = 0 To txtAux.Count - 1
'        txtAux(jj).Height = DataGrid1.RowHeight
'        txtAux(jj).Top = alto
'        txtAux(jj).visible = b
'        If b Then txtAux(jj).Text = ""
'    Next jj
'
'    jj = 0
'    Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
'    Me.cmdAux(jj).Top = alto
'    Me.cmdAux(jj).visible = b
'End Sub



'Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento Articulos
'    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
'    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub



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
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " ORDER BY id"
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub frmC_Selec(vFecha As Date)
    Text1(3).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVe_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0  '
            Set frmVe = New frmFacVehiculos
            frmVe.DatosADevolverBusqueda = "0|1|"
            frmVe.Show vbModal
            Set frmVe = Nothing
           PonerFoco Text1(2)
        Case 1
            
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
             PonerFoco Text1(5)
    End Select
   
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(3).Text)
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Plantilla
        ' BotonEliminarLinea
    Else   'Eliminar Plantilla
         BotonEliminar
    End If
End Sub

Private Sub mnLineas_Click()
    
    'BotonMtoLineas
    If BLOQUEADesdeFormulario(Me) Then
        frmFacRutasEnvioLineas.vCodigoCabcera = Data1.Recordset!ID
        frmFacRutasEnvioLineas.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            TerminaBloquear
            CargaGrid True
        End If
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         'BotonModificarLinea
    Else   'Modificar Cabecera Oferta
    
        If Val(Data1.Recordset!Situacion) <> 0 Then
            If vUsu.Nivel > 1 Then
                MsgBox "Solo se pueden modificar  la orden de carga esta en situación: abierta", vbExclamation
                Exit Sub
            Else
                If MsgBox("NO deberia modificar la orden de carga en esta situacion.    ¿Continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
        End If
    
    
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         'BotonAnyadirLinea
    Else 'Añadir Cabecera de Ofertas
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


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    With Text1(Index)
        Select Case Index
            Case 0 'Codigo Plantilla
                If PonerFormatoEntero(Text1(Index)) Then
                    'comprobar si ya existe el codigo de plantilla
                    If Modo = 3 Then 'Insertar
                        If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                    End If
                End If
                
            Case 2, 5 'Codigo
                If PonerFormatoEntero(Text1(Index)) Then
                    If Index = 2 Then
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "svehiculos", "descripcion", "codigo")
                    Else
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                    End If
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = ""
                End If
            Case 3
                PonerFormatoFecha Text1(Index)
        End Select
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    'If Button.Index = 10 Or Button.Index = 11 Or Button.Index = 13 Or Button.Index = 14 Then
    If Button.Index >= 10 And Button.Index <= 15 Then
        If Data1.Recordset Is Nothing Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
    End If


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
        Case 10, 11 'Mantenimiento Lineas
                
                            
                
                
                If Button.Index = 10 Then
                    'Si la situacion NO es abierta
                    If Val(Data1.Recordset!Situacion) <> 0 Then
                        MsgBox "Solo se pueden modificar los albaranes de la orden de carga si esta en situació: abierta", vbExclamation
                        Exit Sub
                    End If
                
                
                    mnLineas_Click
                Else
                    Imprimir False
                End If
        
        Case 13
            'Inciiar expedicion
            
            If Val(Data1.Recordset!Situacion) <> 0 Then
                MsgBox "Solo abiertas. Esta ya en situacion:" & Me.cboEstado.Text, vbExclamation
                Exit Sub
            End If
            'Si no hay albaranes asignados
            If ListView1.ListItems.Count = 0 Then
                MsgBox "No hay albaranes asignados", vbExclamation
                Exit Sub
            End If
            If MsgBox("Comienza proceso expedicion?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Screen.MousePointer = vbHourglass
                BotonExpedicion
                Screen.MousePointer = vbDefault
            End If
        Case 14
            ImprimirDatosExpedicion
            
        Case 15
            'Imprimir albaranes asociados. Si puede ser 3 copias
            If Modo = 2 Then ImprimirAlbaranesAsociados
            
        Case 16
            If Modo = 2 Then ImprimirResumen
        Case 21  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = 13 Then 'ENTER
'        KeyAscii = 0
'        SendKeys "{tab}"
'    ElseIf KeyAscii = 27 Then 'ESC
'        Select Case Modo
'            Case 0, 2: Unload Me
'            Case 1: cmdCancelar_Click 'Buscar
'            Case 5 'Lineas
'                If ModificaLineas = 0 Then PonerModo 2
'        End Select
'    End If
'    If Err.Number <> 0 Then Err.Clear
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim B As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    
    B = Modo = 3 Or Modo = 4
    cboConductor.visible = B
    If B And Me.cboConductor.ListCount = 0 Then CargaComboConductores
    Text1(4).visible = Not B
    
    
    
    '==============================
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    cmdRegresar.visible = (Modo = 5)
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = B
    Next I
    BloquearCmb Me.cboEstado, Not B
    
    'For i = 0 To Me.imgBuscar.Count - 1
    Me.imgFecha(3).Enabled = B
    'Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)
     
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '===============================
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

    B = (Modo = 2) Or (Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    B = (Modo = 2)
    'Lineas
    Toolbar1.Buttons(10).Enabled = B
    Me.mnLineas.Enabled = B

    B = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not B Or (Modo = 5)
    Me.mnNuevo.Enabled = Not B Or (Modo = 5)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    
    
    
    
    'orden expedicion
    B = Modo = 2 And vParamAplic.ProduccionNueva
    Toolbar1.Buttons(13).Enabled = B
    Toolbar1.Buttons(14).Enabled = B
    Toolbar1.Buttons(15).Enabled = B
    Toolbar1.Buttons(16).Enabled = B
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    cboEstado.ListIndex = -1
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
Dim Tabla As String
    
    Tabla = "slipla"
    SQL = "SELECT codplant,numlinea," & Tabla & ".codartic, sartic.nomartic, cantidad "
    SQL = SQL & " FROM " & Tabla & " LEFT JOIN sartic ON " & Tabla & ".codartic=sartic.codartic"
    If enlaza Then
        SQL = SQL & " WHERE codplant=" & Text1(0).Text 'Data1.Recordset!codPlant
    Else
        SQL = SQL & " WHERE codplant = -1"
    End If
    SQL = SQL & " ORDER BY " & Tabla & ".numlinea "
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
        CadenaConsulta = "Select * from " & NombreTabla & " ORDER BY id"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
           
    Me.cboConductor.Text = ""
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    Text1(0).Text = SugerirCodigoSiguienteStr("sRepartoC", "id")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
    
    cboEstado.ListIndex = 0
End Sub


'Private Sub BotonAnyadirLinea()
'Dim anc As Single
'
''    'Si no estaba modificando lineas salimos
''    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
''    If ModificaLineas = 2 Then Exit Sub
''
''    ModificaLineas = 1 'Ponemos Modo Añadir Linea
''
''    'Añadiremos el boton de aceptar y demas objetos para insertar
''    PonerBotonCabecera False
''    lblIndicador.Caption = "INSERTAR"
''
''    AnyadirLinea DataGrid1, Data2
''
''    anc = ObtenerAlto(DataGrid1)
''    LLamaLineas anc
''    PonerFoco txtAux(0)
'End Sub


'Private Sub BotonMtoLineas()
'On Error GoTo ErrorLineas
'    Screen.MousePointer = vbHourglass
'    PonerModo (5)
'    ModificaLineas = 0
'    PonerBotonCabecera True
'    CargaGrid True
'    Screen.MousePointer = vbDefault
'    Exit Sub
'ErrorLineas:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
'    Screen.MousePointer = vbDefault
'End Sub


Private Sub BotonModificar()




    'Escondemos el navegador y ponemos Modo Modificar
    
    PonerModo 4
    Me.cboConductor.Text = DBLet(Me.Data1.Recordset!conductor, "T")
    
    PonerFoco Text1(1)
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
        
            
    If Val(Data1.Recordset!Situacion) Then
        MsgBox "Situacion incorrecta", vbExclamation
        Exit Sub
    End If
        
        
    
    'Compruebo si tiene lineas de ya facturadas
    SQL = ""
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).SmallIcon = 44 Then
            'FACTURA
            SQL = SQL & "X"
        End If
    Next
    If SQL <> "" Then
        'La orden de carga tiene faturas. SOLO root puede borrarla
        SQL = Len(SQL)
        SQL = "La orden de carga tiene asociada albaranes ya facturados(" & SQL & ")"
        If vUsu.Codigo < 1 Then
            SQL = SQL & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then SQL = ""
        Else
            MsgBox SQL, vbExclamation
        End If
        If SQL <> "" Then Exit Sub
    End If

        
    
    SQL = "Rutas.                 " & vbCrLf
    SQL = SQL & "----------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar la orden de carga:"
    SQL = SQL & vbCrLf & "Código : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Desc : " & Text1(1).Text
    SQL = SQL & vbCrLf & "Fecha : " & Text1(3).Text
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
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
        MuestraError Err.Number, "Eliminar", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        If Data1.Recordset.EOF Then
            Eliminar = False
            Exit Function
        End If
        Conn.BeginTrans
        
        
        
        SQL = " WHERE id=" & Val(Data1.Recordset!ID)
        
        'Lineas
        Conn.Execute "Delete  from srepartolotcaj " & SQL
        
        Conn.Execute "Delete  from srepartol " & SQL
        
        'Cabeceras
        Conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim C As String

On Error Resume Next

    DatosOk = False
    
    

    Text1(4).Text = Me.cboConductor.Text
    
    
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    If B Then Text1(4).visible = True
    
    
    
    
    If Modo = 4 Then
        If B Then
            If vParamAplic.ProduccionNueva Then
                If Val(Data1.Recordset!Situacion) <> cboEstado.ItemData(cboEstado.ListIndex) Then
                    
                    'Si estaba en carga, podemos desahacer TODO y volver a que genere la tituacion
                    B = False
                    If vUsu.Nivel <= 1 Then
                        'Ha seleccionado nivel 0
                        If cboEstado.ListIndex = 0 Then
                            'HAbia nivel CARGA: 1
                            If Val(Data1.Recordset!Situacion) = 1 Then
                                
                               If MsgBox("Seguro que desea reeestablecer la orden de carga?", vbQuestion + vbYesNo) = vbYes Then
                                    'HAGO COSAS
                                    
                                    C = "DELETE FROM srepartolot WHERE idreparto =" & Text1(0).Text
                                    Conn.Execute C
                                    
                                    C = "UPDATE srepartol SET albexpedido=0 WHERE id =" & Text1(0).Text
                                    Conn.Execute C
                                    
                                    
                                    B = True
                                End If
                                    
                            End If
                        End If
                    End If
                
                    If Not B Then
                        MsgBox "No puede cambiar la situacion", vbExclamation
                        Exit Function
                    End If
                End If
            End If
        End If 'b=true
    End If 'modificar

    
    
    
    If Modo = 3 Then
        If cboEstado.ListIndex > 0 Then
            MsgBox "Al crear una orden nueva, el estado de la orden de carga SOLO puede ser abierto.", vbExclamation
            Exit Function
        End If
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
    'Registro de la tabla de cabeceras: scapla
    cad = cad & ParaGrid(Text1(0), 10, "Código")
    cad = cad & ParaGrid(Text1(3), 14, "Fecha")
    cad = cad & ParaGrid(Text1(1), 43, "Observaciones")
    cad = cad & "Conductor||conductor|T||33·"
    
    Tabla = NombreTabla
    Titulo = "Rutas envio"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
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
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " ORDER BY id"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim cadMen As String
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        cadMen = "No hay ningún registro en la tabla " & NombreTabla
        If Modo = 1 Then
            MsgBox cadMen & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox cadMen, vbInformation
        End If
        CargaGrid False
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
    '--Poner el nombre del Grupo Plantilla
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "svehiculos", "descripcion", "codigo")
    Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "straba", "nomtraba", "codtraba")
    CargaGrid True
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub



Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de plantillas
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


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = " id = " & Text1(0).Text
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        Indicador = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

'Private Function ObtenerWhereCP() As String
'Dim SQL As String
'
'    SQL = " WHERE codplant= " & Text1(0).Text
'    ObtenerWhereCP = SQL
'End Function


'Private Sub txtAux_GotFocus(Index As Integer)
'    ConseguirFocoLin txtAux(Index)
'End Sub
'
'Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'        KEYdown KeyCode
'End Sub
'
'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'        KEYpress KeyAscii
'End Sub
'
'
'Private Sub txtAux_LostFocus(Index As Integer)
'
'    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    Select Case Index
'        Case 0 'Cod Articulo
'           txtAux(1).Text = PonerNombreDeCod(txtAux(0), 1, "sartic", "nomartic", "codartic", " Artículo ", "T")
'           If txtAux(1).Text = "" And txtAux(0).Text <> "" Then PonerFoco txtAux(0)
'        Case 2 'Cantidad
'            If txtAux(Index).Text <> "" Then
'                PonerFormatoDecimal txtAux(Index), 1 'Tipo 1: Decimal(12,2)
'                PonerFocoBtn Me.cmdAceptar
'            End If
'    End Select
'End Sub
'

'Private Function InsertarLinea() As Boolean
''Inserta un registro en la tabla de lineas de Plantilla: slipla
'Dim SQL As String
'Dim numlinea As String, vWhere As String
'
'    On Error GoTo EInsertarLinea
'
'    InsertarLinea = False
'    SQL = ""
'    If DatosOkLinea Then 'Lineas de Ofertas
'        'Conseguir el siguiente numero de linea
'        vWhere = Mid(ObtenerWhereCP, 7)
'        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
'        SQL = "INSERT INTO " & NomTablaLineas
'        SQL = SQL & " (codplant, numlinea, codartic, cantidad) "
'        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & DBSet(txtAux(0).Text, "T") & ","
'        SQL = SQL & DBSet(txtAux(2).Text, "N") & ") "
'    End If
'
'    If SQL <> "" Then
'        Conn.Execute SQL
'        InsertarLinea = True
'    End If
'    Exit Function
'EInsertarLinea:
'    MuestraError Err.Number, "Insertar Lineas Plantillas" & vbCrLf & Err.Description
'End Function


'Private Function DatosOkLinea() As Boolean
'Dim b As Boolean
'Dim vArtic As CArticulo
'Dim SQL As String
'
'    On Error GoTo EDatosOkLinea
'
'    DatosOkLinea = False
'    b = True
'
'    If txtAux(0).Text = "" Then
'        MsgBox "El campo Cod. Articulo no puede ser nulo.", vbExclamation
'        b = False
'        PonerFoco txtAux(0)
'        Exit Function
'    End If
'    'If Not b Then Exit Function
'
'    'Comprobar que existe el articulo seleccionado
'    Set vArtic = New CArticulo
'    If Not vArtic.Existe(txtAux(0).Text) Then
'        b = False
'        PonerFoco txtAux(0)
'    ElseIf ModificaLineas = 1 Then
'        'si existe miramos si ya hay una linea con ese artículo antes de insertar
'        SQL = "SELECT COUNT(*) FROM " & NomTablaLineas & ObtenerWhereCP & " AND codartic=" & DBSet(txtAux(0).Text, "T")
'        If RegistrosAListar(SQL) > 0 Then
'            SQL = "Ya existe una línea en la plantilla con el Artículo: " & txtAux(0).Text & vbCrLf & "¿Desea añadir la linea?"
'            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then b = False
'        End If
'    End If
'    Set vArtic = Nothing
'
'    DatosOkLinea = b
'EDatosOkLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Function


'Private Sub BotonModificarLinea()
''Modificar una linea
'Dim vWhere As String
'Dim anc As Single
'Dim i As Byte
'
'    On Error GoTo EModificarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Then Exit Sub '1= Insertar
'
'    If Data2.Recordset.EOF Then Exit Sub
'
'    'Si BLOQUEA REGISTRO
'    vWhere = Mid(ObtenerWhereCP, 7) & " and numlinea=" & Data2.Recordset!numlinea
'    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
'
'    DataGrid1.Enabled = False
'
'    ModificaLineas = 2 'Modificar
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc
'
'    'cargamos los datos
'    For i = 0 To txtAux.Count - 1
'        txtAux(i).Text = DataGrid1.Columns(i + 2).Text
'    Next i
'
'    PonerFoco txtAux(0)
'
'EModificarLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub
'
'
'Private Function ModificarLinea() As Boolean
''Modifica un registro en la tabla de Lineas Plantillas: slipla
'Dim SQL As String
'
'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    SQL = ""
'    If DatosOkLinea Then
'        SQL = "UPDATE " & NomTablaLineas & " Set codartic = " & DBSet(txtAux(0).Text, "T") & ", "
'        SQL = SQL & " cantidad = " & DBSet(txtAux(2).Text, "N")
'        SQL = SQL & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea
'    End If
'
'    If SQL <> "" Then
'        Conn.Execute SQL
'        ModificarLinea = True
'    End If
'    Exit Function
'
'EModificarLinea:
'    MuestraError Err.Number, "Modificar Lineas Plantilla" & vbCrLf & Err.Description
'End Function
'
'
'
'Private Sub BotonEliminarLinea()
''Eliminar una linea De Mantenimiento. Tabla: slima1
'Dim SQL As String
'
'    On Error GoTo EEliminarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'    If Data2.Recordset.EOF Then Exit Sub
'
'    ModificaLineas = 3 'Eliminar
'    SQL = "¿Seguro que desea eliminar la línea de Plantilla?     " & vbCrLf
'    SQL = SQL & vbCrLf & "Plantilla: " & Text1(0).Text & " - " & Text1(1).Text
'    SQL = SQL & vbCrLf & "NumLinea: " & Data2.Recordset!numlinea
'    SQL = SQL & vbCrLf & "Articulo: " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
'
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'        'Hay que eliminar
'        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP
'        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
'        Conn.Execute SQL
'        ModificaLineas = 0
'        CargaGrid True
'
'        CancelaADODC Me.Data2
'    End If
'    PonerFocoBtn Me.cmdRegresar
'
'EEliminarLinea:
'        Screen.MousePointer = vbDefault
'        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
'End Sub


Private Sub Imprimir(SinVentanaOk As Boolean)
Dim I As Integer
Dim cad As String
Dim SQL As String
Dim Litros As Currency

    On Error GoTo EImprimir

    Conn.Execute "DELETE FROM tmprutas where codusu = " & vUsu.Codigo

   '`tmprutas`
   'insert into `tmprutas`
   ' (`codusu`,`idruta`,`codigo`,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,
   '`codartic`,`nomartic`,`cajas`,`fecha2`)


    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    For I = 1 To ListView1.ListItems.Count
        lblIndicador.Caption = 1 & " de " & ListView1.ListItems.Count
        lblIndicador.Refresh
        If ListView1.ListItems(I).SmallIcon = 43 Then
            '************************************   ALBARANES
            SQL = "select scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.codartic,slialb.nomartic,cantidad,cajas,nomclien,"
            SQL = SQL & "domclien, codpobla, pobclien, proclien,LitrosUnidad from slialb,scaalb,sartic where slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar"
            SQL = SQL & " AND slialb.codartic=sartic.codartic"
            SQL = SQL & " and slialb.codartic<>'" & vParamAplic.ArtReciclado & "'"  'Que no salgal el punto verde
            'Ahora el albaran en cuetion
            SQL = SQL & " AND scaalb.codtipom='" & ListView1.ListItems(I).Text & "' "
            SQL = SQL & " AND scaalb.numalbar=" & ListView1.ListItems(I).SubItems(1) & " "
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            cad = ""
            While Not miRsAux.EOF
               ' (`codusu`,`idruta`,,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,`fecha2`
               '`codigo` `codartic`,`nomartic`,`cajas`,)

                NumRegElim = NumRegElim + 1
                If cad = "" Then
                    cad = ", (" & vUsu.Codigo & "," & Text1(0).Text & ",'"
                    cad = cad & miRsAux!Codtipom & Format(miRsAux!NumAlbar, "0000000") & "',"
                    cad = cad & DBSet(miRsAux!nomclien, "T") & "," & DBSet(miRsAux!domclien, "T") & ","
                    cad = cad & DBSet(miRsAux!pobclien, "T") & ",'"
                    'cppos, provinci
                    cad = cad & DevNombreSQL(Trim(DBLet(miRsAux!codpobla, "T") & "   " & DBLet(miRsAux!proclien, "T"))) & "','"
                    cad = cad & Format(miRsAux!FechaAlb, FormatoFecha) & "',"
                End If
                'Faltan: `codigo` `codartic`,`nomartic`,`cajas`,)
                SQL = SQL & cad & NumRegElim & "," & DBSet(miRsAux!codartic, "T") & ","
                Litros = DBLet(miRsAux!LitrosUnidad, "N")
                Litros = Litros * miRsAux!Cantidad
                SQL = SQL & DBSet(miRsAux!NomArtic, "T") & "," & miRsAux!Cajas & "," & DBSet(Litros, "N") & ")"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If SQL <> "" Then
                'Tiene datos
                cad = "INSERT INTO tmprutas (`codusu`,`idruta`,`idalb`,`nomclien`,`domclien`,"
                cad = cad & "`pobclien`,`proclien`,`fecha2`,`codigo`,`codartic`,`nomartic`,`cajas`,litros) VALUES "
                cad = cad & Mid(SQL, 2) 'quito la primera coma

                Conn.Execute cad
                
                
                
                
            End If
        Else
            '************************************   FACTURAS
            SQL = "select scafac.codtipom,scafac.numfactu,scafac.fecfactu,slifac.codartic,slifac.nomartic,cantidad,nomclien,"
            SQL = SQL & "domclien, codpobla, pobclien, proclien,unicajas,LitrosUnidad from slifac,scafac,sartic where"
            SQL = SQL & " slifac.codtipom=scafac.codtipom and scafac.numfactu=slifac.numfactu"
            SQL = SQL & " AND slifac.codartic=sartic.codartic"
            SQL = SQL & " and slifac.codartic<>'" & vParamAplic.ArtReciclado & "'"  'Que no salgal el punto verde
            'Ahora el albaran en cuetion
            SQL = SQL & " AND scafac.codtipom='" & ListView1.ListItems(I).Text & "' "
            SQL = SQL & " AND scafac.numfactu=" & ListView1.ListItems(I).SubItems(1) & " "
            SQL = SQL & " AND scafac.fecfactu='" & Format(ListView1.ListItems(I).SubItems(2), FormatoFecha) & "' "
            SQL = SQL & " AND slifac.numalbar=" & ListView1.ListItems(I).Tag & " "
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            cad = ""
            While Not miRsAux.EOF
               ' (`codusu`,`idruta`,,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,`fecha2`
               '`codigo` `codartic`,`nomartic`,`cajas`,)

                NumRegElim = NumRegElim + 1
                If cad = "" Then
                    cad = ", (" & vUsu.Codigo & "," & Text1(0).Text & ",'"
                    cad = cad & miRsAux!Codtipom & Format(miRsAux!NumFactu, "0000000") & "',"
                    cad = cad & DBSet(miRsAux!nomclien, "T") & "," & DBSet(miRsAux!domclien, "T") & ","
                    cad = cad & DBSet(miRsAux!pobclien, "T") & ",'"
                    'cppos, provinci
                    cad = cad & DevNombreSQL(Trim(DBLet(miRsAux!codpobla, "T") & "   " & DBLet(miRsAux!proclien, "T"))) & "','"
                    cad = cad & Format(miRsAux!FecFactu, FormatoFecha) & "',"
                End If
                'Faltan: `codigo` `codartic`,`nomartic`,`cajas`,)
                SQL = SQL & cad & NumRegElim & "," & DBSet(miRsAux!codartic, "T") & ","
                SQL = SQL & DBSet(miRsAux!NomArtic, "T") & ","
                If DBLet(miRsAux!UniCajas, "N") = 0 Then
                    SQL = SQL & Round(miRsAux!Cantidad, 0)
                Else
                    SQL = SQL & CStr(CInt(miRsAux!Cantidad) \ CInt(miRsAux!UniCajas))
                End If
                Litros = DBLet(miRsAux!LitrosUnidad, "N")
                Litros = Litros * miRsAux!Cantidad
                SQL = SQL & "," & DBSet(Litros, "N")
                SQL = SQL & ")"
                    
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If SQL <> "" Then
                'Tiene datos
                cad = "INSERT INTO tmprutas (`codusu`,`idruta`,`idalb`,`nomclien`,`domclien`,"
                cad = cad & "`pobclien`,`proclien`,`fecha2`,`codigo`,`codartic`,`nomartic`,`cajas`,litros) VALUES "
                cad = cad & Mid(SQL, 2) 'quito la primera coma

                Conn.Execute cad
            End If
        
        
        End If
    
    
    Next
    
    If NumRegElim > 0 Then
            
            cad = DevuelveNombreReport(40)
            
    
            With frmImprimir
                .FormulaSeleccion = "{tmprutas.codusu} = " & vUsu.Codigo
                .OtrosParametros = ""
                .NumeroParametros = 0
        
                If SinVentanaOk Then
                    .SoloImprimir = True
                    .NumeroDeCopias = 1
                Else
                    .SoloImprimir = False
                End If
                .EnvioEMail = False
                .opcion = 2016
                .Titulo = Me.Caption
                .NombreRPT = cad
                .ConSubInforme = True
                .Show vbModal
            End With

    End If
    
    Exit Sub
EImprimir:
    MuestraError Err.Number
End Sub


Private Sub CargaComboConductores()
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select conductor from srepartoc group by  conductor", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then

        While Not miRsAux.EOF
            If DBLet(miRsAux!conductor, "T") <> "" Then Me.cboConductor.AddItem miRsAux!conductor
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Me.cboConductor.ListCount = 0 Then Me.cboConductor.AddItem "Sin asignar"
End Sub


Private Sub CargaComobo()
    '## Cuando contabilice, que valor pondra en el campo observaciones del
    '   la factura, tanto cliente como de proveedores

    Me.cboEstado.Clear
    Me.cboEstado.AddItem "Abierta"
    cboEstado.ItemData(cboEstado.NewIndex) = 0
    
    cboEstado.AddItem "En carga"
    cboEstado.ItemData(cboEstado.NewIndex) = 1

    cboEstado.AddItem "Finalizada la carga"
    cboEstado.ItemData(cboEstado.NewIndex) = 2

    cboEstado.AddItem "Expedida"
    cboEstado.ItemData(cboEstado.NewIndex) = 3

End Sub



Private Sub BotonExpedicion()
Dim SQL As String

         If BLOQUEADesdeFormulario(Me) Then
            
                'Hacemos el proceso de expedicion.
                'Es decir, generaremos los datos para la impresion
                SQL = DevuelveDesdeBD(conAri, "count(*)", "srepartolot", "idreparto", Text1(0).Text)
                If SQL = "" Then SQL = "0"
                If Val(SQL) > 0 Then
                    MsgBox "Ya habian datos. Que hacemos?  Avise soporte tecnico", vbExclamation
                Else
                    For NumRegElim = 1 To ListView1.ListItems.Count
                        'Para cada albaran meteremos la linea
                        SQL = "INSERT INTO srepartolot (idreparto , codtipom,numalbar,numlinea,codartic,cajas)"
                        'De momento todo lo que venga aqui sera ALV. Por eso pondremos un 1 como ALV para favorecer los codigos de barra
                        SQL = SQL & " select " & Text1(0).Text & ",1,numalbar,numlinea,slialb.codartic,cajas"
                        SQL = SQL & " from slialb,sartic where slialb.codartic=sartic.codartic and sartic.trazabilidad=1"
                        'Para cada albaran
                        SQL = SQL & " and codtipom='" & ListView1.ListItems(NumRegElim).Text & "' AND numalbar=" & ListView1.ListItems(NumRegElim).SubItems(1)
                        
                        Conn.Execute SQL
                    Next NumRegElim
                        
                        
                    'Maracamos que se esta "en carga
                    SQL = "UPDATE srepartoc set situacion=1 WHERE  id =" & Text1(0).Text
                    Conn.Execute SQL
                    
                    'Deberiamos lanzar la impresion
                    ImprimirDatosExpedicion
                    
                    Imprimir True   'Imprimos la antigua para que la firme el chofer
                    
                    ' ya tene cargada la tabla con los articulos que vamos a subir.
                    ' hay que ir asignado cajas, o bien una a una o bien con pistola
                        
                     
                        
                End If  'de val(sql)>0

                TerminaBloquear
                PosicionarData
                PonerCampos
        End If
        
        
End Sub


Private Sub ImprimirDatosExpedicion()
    Screen.MousePointer = vbHourglass
    If GeneraDatosImpresion Then
        With frmImprimir
                .FormulaSeleccion = "{srepartoc.id} = " & Me.Text1(0).Text
                .OtrosParametros = "|usu=" & vUsu.Codigo & "|"
                .NumeroParametros = 1
                .SoloImprimir = False
                .EnvioEMail = False
                .opcion = 2016
                .Titulo = "Expedicion"
                .NombreRPT = "MorOrdenCargaN.rpt"
                .ConSubInforme = True
                .Show vbModal
            End With
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub ImprimirResumen()
    If Val(Me.Data1.Recordset!Situacion) <> 3 Then
        MsgBox "No expedida", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    If GeneraDatosImpresion Then
        With frmImprimir
                .FormulaSeleccion = "{srepartoc.id} = " & Me.Text1(0).Text
                .OtrosParametros = "|usu=" & vUsu.Codigo & "|"
                .NumeroParametros = 1
                .SoloImprimir = False
                .EnvioEMail = False
                .opcion = 2016
                .Titulo = "Resumen expedicion"
                .NombreRPT = "MorOrdenCargaResumen.rpt"
                .ConSubInforme = True
                
                .Show vbModal
            End With
    End If
    Screen.MousePointer = vbDefault
End Sub


'para imprimir los datos
Private Function GeneraDatosImpresion() As Boolean
Dim SQL As String
Dim Udspal As Long
Dim palets As Long

    On Error GoTo EGeneraDatosImpresion
    GeneraDatosImpresion = False
    Set miRsAux = New ADODB.Recordset
    
    Conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM tmppartidas where codusu = " & vUsu.Codigo
    
    'Ahora iremos albaran por al
    SQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,importe1,importe2,porcen1) "
    SQL = SQL & " select " & vUsu.Codigo & ",numalbar,numlinea,srepartolot.codartic,nomartic,cajas,cajas,unicajas from"
    SQL = SQL & " srepartolot left join sartic on srepartolot.codartic=sartic.codartic where idreparto=" & Text1(0).Text
    SQL = SQL & " order by numalbar,numlinea"
    Conn.Execute SQL
    
    
    'Julio 2012.
    'Añadiremos Cajas x Palet, y palets   importe3 importe4
    Espera 0.5
    SQL = "Select * from tmpinformes where codusu = " & vUsu.Codigo & " ORDER BY nombre1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        palets = Val(DBLet(miRsAux!Importe1))
        If palets > 0 Then
            
            SQL = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", miRsAux!nombre1, "T")
            If SQL = "" Then SQL = "1"
            If SQL = "0" Then SQL = "1"
            Udspal = Val(SQL)
                                    
            palets = ((palets - 1) \ Udspal) + 1
            
            SQL = "UPDATE tmpinformes SET importe3=" & Udspal & ", importe4=" & palets
            SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND codigo1=" & miRsAux!Codigo1
            SQL = SQL & " AND campo1 = " & miRsAux!campo1
            Conn.Execute SQL
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    SQL = "select codartic,nomartic,cantidad from slialb where (numalbar,codtipom) in ("
    SQL = SQL & " select s.numalbar ,s.codtipom from srepartol s left join scaalb a on s.codtipom=a.codtipom and s.numalbar=a.numalbar"
    SQL = SQL & " and s.fechaalb=a.fechaalb where id= " & Text1(0).Text & " AND codartic <> '" & vParamAplic.ArtReciclado
    SQL = SQL & "'  group  by 1,2) GROUP BY 1,2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = "insert into tmppartidas(codusu,idpartida,codartic,Referencia,numlote,fecha,cantidad)"
        SQL = SQL & " select " & vUsu.Codigo & " codusu,id,spartidas.codartic,nomartic,numlote,fecha,"
        SQL = SQL & "  if(COALESCE(unicajas,0)=0,0,((cantotal-1) div unicajas)+1) cajas  "
        SQL = SQL & " from spartidas,sartic  where  spartidas.codartic=sartic.codartic AND"
        SQL = SQL & " spartidas.codartic=" & DBSet(miRsAux!codartic, "T") & " and cantotal>0 order by fecha asc"
        Conn.Execute SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'mas adelante miraremos los palets
    
    
    GeneraDatosImpresion = True
    
    
    
    
    
EGeneraDatosImpresion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Impresion datos carga"
    Set miRsAux = Nothing
End Function


Private Sub ImprimirAlbaranesAsociados()
Dim SQL As String
Dim I As Integer

    SQL = ""
    NumRegElim = 0
    For I = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Text = "ALV" Then
            SQL = SQL & ", ('" & ListView1.ListItems(I).Text & "'," & Val(ListView1.ListItems(I).SubItems(1)) & ")"
            NumRegElim = NumRegElim + 1
        End If
    Next
    
    
    If NumRegElim = 0 Then
        MsgBox "No hay albaranes para imprimir", vbExclamation
        Exit Sub
    End If
    
    If Val(Data1.Recordset!Situacion) <> 3 Then
        If MsgBox("Orden de carga NO finalizada. Albaranes sin numero de lote. Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    If MsgBox("Desea imprimir(por triplicado) los albaranes asociados a la expedicion?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Set miRsAux = New ADODB.Recordset
    lblIndicador.Caption = "Imp. albaranes"
    lblIndicador.Refresh
    SQL = Mid(SQL, 2) 'quito la primera coma
    SQL = "Select * from scaalb where (codtipom,numalbar) in (" & SQL & ")"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No se han encotrado los albaranes", vbExclamation
    Else
        Screen.MousePointer = vbHourglass
        While Not miRsAux.EOF
            lblIndicador.Caption = "ALV" & miRsAux!NumAlbar
            lblIndicador.Refresh
            SQL = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", CStr(miRsAux!CodClien), "N")
            If SQL = "" Then SQL = "0"
            ImprimirAlbaran 45, CStr(miRsAux!NumAlbar), CByte(SQL)
            miRsAux.MoveNext
            Me.Refresh
            DoEvents
            Espera 0.2
        Wend
        Screen.MousePointer = vbDefault
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub ImprimirAlbaran(opcion As Integer, NumAlbar As String, ImprimeValorado2 As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NomTabla As String
Dim Codtipom As String


    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NomTabla = "scaalb"
   
    '===================================================
    '============ PARAMETROS ===========================
   ' If (opcion = 45) Then
   '     If vUsu.TrabajadorB Then
   '         indRPT = 29   'Albaranes B
   '         Codtipom = "ALZ"
   '     Else
            indRPT = 10 'Albaran Clientes
            Codtipom = "ALV"
   '     End If
   ' End If
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
        Exit Sub
    End If

    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
                
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If NumAlbar <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NomTabla & ".codtipom}='" & Codtipom & "'" 'Val(txtCodigo(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{" & NomTabla & ".numalbar}=" & NumAlbar
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'select para insertar en tabla temporal
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
    End If
   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "codclien", "codtipom", Codtipom, "T", , "numalbar", NumAlbar, "N")
    If devuelve <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", devuelve, "N")
        
        If vUsu.TrabajadorB Then
            devuelve = "2"
        Else
            If devuelve = "3" Then devuelve = "2" 'El intracomunitario lo trato como exento
        End If
        
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
    End If
     
     
    'Si se imprimen importes y/o
''    If ImprimeValorado Then
''        devuelve = "0"
''    Else
''        devuelve = "2"
''    End If
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    'cadParam = cadParam & "Albarcon= " & devuelve & "|"
    cadParam = cadParam & "Albarcon= " & ImprimeValorado2 & "|"
    numParam = numParam + 1

     
     
     
     With frmImprimir
        
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .NumeroDeCopias = 3
            .SoloImprimir = True
            .EnvioEMail = False
            .opcion = opcion
            .Titulo = "Albaran de Cliente"
            .Show vbModal
    End With
End Sub

