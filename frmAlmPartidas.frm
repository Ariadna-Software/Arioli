VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmPartidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partidas (Lotes)"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13410
   ClipControls    =   0   'False
   Icon            =   "frmAlmPartidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   19
      ToolTipText     =   "Buscar trabajador"
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   6000
      TabIndex        =   3
      Tag             =   "Albara|T|S|||spartidas|numalbar||N|"
      Text            =   "nº albar"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   5160
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Prove|N|S|0||spartidas|codprove||N|"
      Text            =   "cantidad"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   10320
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "Cantidad|N|N|||spartidas|cantotal|#,##0.00|N|"
      Text            =   "cantidad"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   8280
      MaxLength       =   60
      TabIndex        =   5
      Tag             =   "Lote|T|N|||spartidas|numlote||N|"
      Text            =   "lote"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   7
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   5520
      Width           =   3525
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Tag             =   "codart|T|N|||spartidas|codartic|||"
      Text            =   "me"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   7080
      TabIndex        =   4
      Tag             =   "Fecha|F|N|||spartidas|fecha|dd/mm/yyyy|N|"
      Text            =   "Fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
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
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Buscar trabajador"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "Id|N|N|||spartidas|id|0|S|"
      Text            =   "id"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   5565
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   5565
      Width           =   1155
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12075
      TabIndex        =   12
      Top             =   5565
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Tag             =   "Almacen|N|N|0||spartidas|codalmac|0|N|"
      Text            =   "coda"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir etiquetas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   360
      Top             =   5520
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmPartidas.frx":000C
      Height          =   4725
      Left            =   240
      TabIndex        =   13
      Top             =   540
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8334
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
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   4080
      TabIndex        =   20
      Top             =   5520
      Width           =   735
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
      TabIndex        =   15
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
Attribute VB_Name = "frmAlmPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
    
Public BuscarNegativos As Boolean  'para cunado vengo a buscar, permitir que muestren negativos y  cero
    
    'Lleva
        '  * si viene para buscar el idpartida
        ' codartic si viene para que devuelva LOTE y cantidad
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public ParaMostrarDesdeNuevaProduccion As String


'Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos
Attribute frmArt.VB_VarHelpID = -1

Private NombreTabla As String
Private Ordenacion As String
Private Modo As Byte

Dim Kcampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros
Dim PrimVez As Boolean
Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean






Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long
Dim cP As cPartidas

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            
            Set cP = New cPartidas
            If DatosOk(cP) Then
                
                If cP.Insertar Then
                    'DEBIRAMOS AÑADIR UNA LINEA
                    'Si no, no cuadrara el total con la suma de lineas
                    AnyadirLineaLotaje
                    
                    txtAux(0).Text = cP.idPartida
                    InsertaLog
                    CargaGrid True
                    BotonAnyadir
                    

                End If
                
            End If
            Set cP = Nothing
        
        Case 4 'MODIFICAR
            Set cP = New cPartidas
            If DatosOk(cP) Then
                
                 If ModificaDesdeFormulario(Me, 3) Then
                     InsertaLog
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion Data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
            End If
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error1:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    HaDevueltoDatos = False
    
    Select Case Index
        Case 1 'cod. tecnico
            
            
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo busqueda
            frmArt.DeConsulta = True
            frmArt.Show vbModal
            Set frmArt = Nothing
            If HaDevueltoDatos Then PonerFoco txtAux(2)
            
'        Case 2 'cod. cliente
'            Set frmC = New frmFacClientes
'            frmC.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
'            frmC.Show vbModal
'            Set frmC = Nothing

        Case 0
            Set frmP = New frmComProveedores
            frmP.DatosADevolverBusqueda = "0|1|"
            frmP.Show vbModal
            Set frmP = Nothing
            If HaDevueltoDatos Then PonerFocoBtn Me.cmdAceptar
    End Select
    
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            limpiar Me
            PonerModo 0
            LLamaLineas 10
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
    End Select
    Exit Sub
    
ECancelar:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    If Me.DatosADevolverBusqueda = "*" Then
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
    Else
    
        If UCase(CStr(Data1.Recordset!codartic)) <> UCase(Me.DatosADevolverBusqueda) Then
            MsgBox "Error en articulo", vbExclamation
            Exit Sub
        End If
    
        Cad = Data1.Recordset!NUmlote & "|"
        Cad = Cad & Data1.Recordset!cantotal & "|"
    End If
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub Data1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim limpiar As Boolean
    limpiar = True
    If Modo = 2 Then
        If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
            txtAux(7).Text = DBLet(Data1.Recordset!codProve, "T")
            txtAux2(7).Text = DBLet(Data1.Recordset!nomprove, "T")
            limpiar = False
        End If
    End If
    If limpiar Then
        txtAux(7).Text = ""
        txtAux2(7).Text = ""
    End If
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then
        cmdRegresar_Click
    Else
        'JUN14
        'Abriremos el form de movimientos de partidas
        If Not Me.Data1.Recordset.EOF Then
            frmAlmpartidasMov.VerPartida = Data1.Recordset!ID
            frmAlmpartidasMov.Show vbModal   'igual e
        End If
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0 Or Modo = 2) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub

Private Sub Form_Activate()

    If PrimVez Then
        PrimVez = False
        If Modo = 1 Then
            Modo = 2
            BotonBuscar
            PonerFoco Me.txtAux(6)
        End If
    End If
        

    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Form_Load()
Dim enlaza As Boolean
    'Icono del formulario
    Me.Icon = frmppal.Icon
    PrimVez = True
    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
       
        .Buttons(10).Image = 10 'Generar mes aut.
        .Buttons(11).Image = 24 'Generar Norma 34
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(15).Image = 15 'Salir
    End With
    
    limpiar Me   'Limpia los campos TextBox
    DataGrid1.ClearFields 'limpiar el Grid

    
    NombreTabla = "spartidas" 'Tabla Nominas y Gastos
    
    'Ordenacion
    Ordenacion = " ORDER BY id"
    'Si esta buscando el lote ordenaremos por fecha
    If DatosADevolverBusqueda <> "" Then
        If DatosADevolverBusqueda <> "*" Then Ordenacion = " ORDER BY fecha"
    End If
    
    
    enlaza = False
    CadenaBusqueda = ""
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        'se le llama desde otro form
        If DatosADevolverBusqueda = "*" Then
            BotonBuscar
        Else
            'HA PUESTO UN CODARTIC
            enlaza = True
            EsBusqueda = True
           
            CadenaBusqueda = " AND spartidas.codartic = '" & DatosADevolverBusqueda & "'"
            'Le manda el numero de lote
            If ParaMostrarDesdeNuevaProduccion <> "" Then CadenaBusqueda = CadenaBusqueda & " AND spartidas.numlote = " & DBSet(ParaMostrarDesdeNuevaProduccion, "T")
        End If
    End If
    
    'CargarCombo_SiNo Me.CmbAux(0)
    'CargarCombo_SiNo Me.CmbAux(1)
    
    CargaGrid enlaza
    If enlaza Then
        PonerModo 2
    Else
        PonerModo 1
       
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim tots As String
    
    On Error GoTo ECarga
    
    tots = MontaSQLCarga(enlaza)
    
    CargaGridGnral DataGrid1, Me.Data1, tots, False
    
    

    
    tots = "S|txtAux(0)|T|id|850|;S|txtAux(1)|T|codartic|1700|;S|cmdAux(1)|B||0|;S|txtAux2(1)|T|Descripcion|3600|;"
    tots = tots & "S|txtAux(2)|T|Alma.|700|;S|txtAux(3)|T|Albaran|1240|;S|txtAux(4)|T|Fecha|1140|;"
    tots = tots & "S|txtAux(5)|T|Lote|1800|;S|txtAux(6)|T|Cantidad|1200|;N|||||;N|||||;"
    
    arregla tots, DataGrid1, Me

    DataGrid1.ScrollBars = dbgAutomatic
   
   
   
    If Not Data1.Recordset.EOF Then
   
          txtAux(7).Text = DBLet(Data1.Recordset!codProve, "T")
          txtAux2(7).Text = DBLet(Data1.Recordset!nomprove, "T")
    
    End If
   
   Exit Sub
   
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim b As Boolean

    On Error Resume Next
    
    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 2   'El ultimo siempre esta visible
        
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
    
    txtAux(0).Enabled = Modo = 1
    
    txtAux(7).Enabled = b  'siempre esta visible
    
    txtAux2(1).Height = DataGrid1.RowHeight
    txtAux2(1).Top = alto
    txtAux2(1).visible = b
    

    'boton de busqueda

    Me.cmdAux(1).Height = Me.DataGrid1.RowHeight
    Me.cmdAux(1).Top = alto
    Me.cmdAux(1).visible = b
    Me.cmdAux(1).Enabled = b And Modo <> 4
   
    Me.cmdAux(0).visible = b
    
    If Err.Number Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ParaMostrarDesdeNuevaProduccion = ""
    BuscarNegativos = False
End Sub

'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Trabajadores
'    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod traba
'    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom traba
'End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    HaDevueltoDatos = True
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(7).Text = RecuperaValor(CadenaSeleccion, 2)
    HaDevueltoDatos = True
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
Dim SQL As String

    If Me.Data1.Recordset.EOF Then Exit Sub
    
    SQL = " id=" & DBSet(Me.Data1.Recordset!ID, "N")
    
    If BloqueaRegistro(NombreTabla, SQL) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1: mnBuscar_Click 'Busqueda
        Case 2: mnVerTodos_Click 'Ver Todos
            
        Case 5: mnNuevo_Click 'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click 'Eliminar
            
        'Case 10: BotonGenerarMes 'generar automat. nuevas nominas
        'Case 11: BotonGenerarNorma34 'generar Norma 34
        Case 12:  ImpEtiquetas
        
        Case 15: mnSalir_Click  'Salir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim I As Byte
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
                      
    'Modo Buscar
    If Kmodo = 1 Then PonerFoco txtAux(0)
                      
    'Bloquear los campos de clave primaria al modificar
    For I = 0 To 2
        BloquearTxt txtAux(I), (Modo = 4)
    Next I
                      
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
    
    On Error Resume Next

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    b = (Modo = 2 Or Modo = 0)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '
    Toolbar1.Buttons(10).Enabled = b
    Toolbar1.Buttons(11).Enabled = b
    Toolbar1.Buttons(12).Enabled = b 'IMPRIMIR
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo opciones del menú.", Err.Description

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
Dim Orden As Boolean

    On Error GoTo ErrSQL
    
    SQL = "select id,spartidas.codartic,nomartic,codalmac,"
    SQL = SQL & "numalbar,fecha,numlote,cantotal,spartidas.codprove,nomprove from (spartidas left join sprove on spartidas.codprove = sprove.codprove)"
    SQL = SQL & ",sartic where spartidas.codartic=sartic.codartic "
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        ElseIf Modo = 3 Then
            If Data1.Recordset.RecordCount < 1 Then SQL = SQL & " AND id=" & txtAux(0).Text
        End If
    Else
        SQL = SQL & " AND  id = -1"
    End If
    Orden = False
    If Me.DatosADevolverBusqueda <> "" Then
        If DatosADevolverBusqueda <> "*" Then
            If Not Me.BuscarNegativos Then SQL = SQL & " AND cantotal>0"
            Orden = True
        End If
    End If
    SQL = SQL & Ordenacion
    If Orden Then SQL = SQL & " DESC"
    MontaSQLCarga = SQL
    Exit Function
    
ErrSQL:
   MuestraError Err.Number, "Cadena SQL", Err.Description
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    limpiar Me
    
    If Modo <> 1 Then
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 10)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(Kcampo).Text = ""
            PonerFoco txtAux(Kcampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    limpiar Me
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid Me.DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    limpiar Me 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
   
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
    'valores por defecto
    
    txtAux(0).Text = "AUTO."
    txtAux(2).Text = "1"
    txtAux(4).Text = Format(Now, "dd/mm/yyyy")
    
    
    'Si viene de buscar numero lote ponemos el codartic
    If DatosADevolverBusqueda <> "" Then
        If DatosADevolverBusqueda <> "*" Then
            txtAux(1).Text = DatosADevolverBusqueda
        End If
    End If
    PonerFoco txtAux(1)
End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    On Error Resume Next
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc

    'poner valores grabados
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    txtAux(1).Text = DBLet(DataGrid1.Columns(1).Value, "N")
    txtAux2(1).Text = DBLet(DataGrid1.Columns(2).Value, "T")
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "N")
    
    txtAux(3).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    txtAux(4).Text = DBLet(DataGrid1.Columns(5).Value, "N")
    txtAux(5).Text = DBLet(DataGrid1.Columns(6).Value, "N")
    txtAux(6).Text = DBLet(DataGrid1.Columns(7).Value, "N")
    txtAux(7).Text = DBLet(DataGrid1.Columns(8).Value, "N")
    txtAux2(7).Text = DBLet(DataGrid1.Columns(9).Value, "T")


    DataGrid1.Enabled = False
    PonerFoco txtAux(3)
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Botón modificar.", Err.Description
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
        
    On Error GoTo FinEliminar
        
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
        
    SQL = "¿Seguro que desea eliminar la partida/lote?" & vbCrLf
    SQL = SQL & vbCrLf & "Id: " & Data1.Recordset.Fields(0).Value
    SQL = SQL & vbCrLf & "Articulo: " & Data1.Recordset.Fields(1).Value & " - " & Data1.Recordset!NomArtic
    SQL = SQL & vbCrLf & "Lote: " & Data1.Recordset.Fields(6).Value
    SQL = SQL & vbCrLf & "Cantidad: " & Format(Data1.Recordset!cantotal, FormatoCantidad)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        conn.BeginTrans
        If Eliminar(SQL) Then
            conn.CommitTrans
            InsertaLog
            
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            
    '        SituarDataPosicion Me.Data1, NumRegElim, SQL
            SituarDataTrasEliminar Me.Data1, NumRegElim, True
        Else
            conn.RollbackTrans
        End If
    End If
    Exit Function
        
FinEliminar:
     Screen.MousePointer = vbDefault
     MuestraError Err.Number, "Eliminar Gastos Técnicos.", Err.Description
End Function

Private Function Eliminar(ByRef s As String) As Boolean
        
    Eliminar = False
    s = "Delete from spartidaslin where id=" & Data1.Recordset!ID
    If EjecutaSQL(1, s) Then
        s = "Delete from " & NombreTabla & " where id=" & Data1.Recordset!ID
        If EjecutaSQL(1, s) Then Eliminar = True
    End If
       

End Function


Private Function DatosOk(ByRef CP1 As cPartidas) As Boolean
Dim b As Boolean
    If Modo = "3" Then txtAux(0).Text = "0"
    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    
    Set miRsAux = New ADODB.Recordset
    CadenaBusqueda = "select * from spartidas where codartic = " & DBSet(txtAux(1).Text, "T")
    CadenaBusqueda = CadenaBusqueda & " and numlote = " & DBSet(txtAux(5).Text, "T")
    If Modo <> 3 Then CadenaBusqueda = CadenaBusqueda & " and id <> " & txtAux(0).Text
    miRsAux.Open CadenaBusqueda, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadenaBusqueda = ""
    While Not miRsAux.EOF
        CadenaBusqueda = CadenaBusqueda & "Y"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If CadenaBusqueda <> "" Then
        If MsgBox("Ya existe el numero de lote para el articulo. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then b = False
    End If
    CadenaBusqueda = ""
    
    
    
    If b And Modo = 3 Then
        With CP1
            .Cantidad = ImporteFormateado(txtAux(6).Text)
            .codalmac = txtAux(2).Text
            .codartic = txtAux(1).Text
            If txtAux(7).Text = "" Then
                .codProve = 0
            Else
                .codProve = txtAux(7).Text
            End If
            .Fecha = txtAux(4).Text
            .NumAlbar = txtAux(3)
            .NUmlote = txtAux(5).Text
        End With
    End If
    
    DatosOk = b
End Function


Private Sub HacerBusqueda()
Dim cadB As String

    On Error Resume Next

    cadB = ObtenerBusqueda(Me, False)

    If cadB <> "" Then 'Se muestran en el mismo form
'        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " AND " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        PonerCadenaBusqueda
        PonerFocoGrid Me.DataGrid1
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerCadenaBusqueda()
Dim Cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
      

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        Cad = "No hay ningún registro en la tabla " & NombreTabla
        If EsBusqueda Then Cad = Cad & " para ese criterio de Búsqueda."
        Screen.MousePointer = vbDefault
        MsgBox Cad, vbInformation
       
        PonerModo Modo
        Exit Sub
    Else
        CargaGrid True
        PonerModo 2
        DataGrid1_RowColChange 1, 1
'        DataGrid1.Refresh
'        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String

    On Error GoTo ErrFoco
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1
            
            'ARticulo
            Cad = ""
            If txtAux(Index).Text <> "" Then
                
                    Cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux(Index).Text, "T")
                    If Cad = "" Then
                        MsgBox "No existe el articulo:" & txtAux(Index).Text, vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(Index)
                    Else
                        PonerFoco txtAux(2)
                    End If
                End If
  
            txtAux2(Index).Text = Cad
            

        Case 2
            If txtAux(2).Text <> "" Then
                If txtAux(2).Text <> "1" Then
                    If PonerFormatoEntero(txtAux(Index)) Then
                        If DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", txtAux(Index).Text, "N") = "" Then
                            MsgBox "No existe el almacen", vbExclamation
                             txtAux(Index).Text = ""
                        End If
                    Else
                        txtAux(Index).Text = ""
                    End If
                End If
            End If
        Case 4
            PonerFormatoFecha txtAux(Index)
            
        Case 6 'Importes
             PonerFormatoDecimal txtAux(Index), 3
        
        Case 7
            
            Cad = ""
            If txtAux(Index).Text <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then
                    Cad = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text, "N")
                    If Cad = "" Then MsgBox "No existe el proveedor"
                End If
            End If
            txtAux2(7).Text = Cad
        
    End Select
    Exit Sub
    
ErrFoco:
    MuestraError Err.Number, "", Err.Description
End Sub



Public Sub AbrirListadoNomi(numero As Integer)
'Abre el Form con los listados de nominas
    Screen.MousePointer = vbHourglass
    frmListadoNomi.parOpcion = numero
    frmListadoNomi.Show vbModal
    Screen.MousePointer = vbDefault
End Sub




Private Sub InsertaLog()

    If Modo = 3 Then
        'INSERTANDO
        CadenaBusqueda = "Nuevo: " & txtAux(0).Text & "  - " & txtAux(1).Text & " - " & txtAux(5).Text & "   " & txtAux(6).Text
    Else
        If Modo = 4 Then
            'MODIFICANDO
            CadenaBusqueda = "Modif:"
        Else
            'ELiminar
            CadenaBusqueda = "Elim:"
        End If
        CadenaBusqueda = CadenaBusqueda & Data1.Recordset!ID & "  - " & Data1.Recordset!codartic & " - " & Data1.Recordset!NUmlote & "   " & Data1.Recordset!cantotal
          
    End If

    Set LOG = New cLOG
    LOG.Insertar 9, vUsu, CadenaBusqueda
    Set LOG = Nothing


    CadenaBusqueda = ""
End Sub


Private Sub AnyadirLineaLotaje()
Dim cLo As cLotaje
 
    Set cLo = New cLotaje
    
    cLo.Cantidad = ImporteFormateado(txtAux(6).Text)
    cLo.codalmac = Val(txtAux(2).Text)
    cLo.codarti2 = ""
    cLo.codartic = txtAux(1).Text
    cLo.DetaMov = "REG" 'regularicacion
    cLo.Documento = txtAux(3).Text
    cLo.Fechamov = txtAux(4).Text
    cLo.HoraMov = CDate(Format(txtAux(4).Text, "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss"))
    cLo.LineaDocu = 1
    cLo.NUmlote = txtAux(5).Text
    cLo.ProvCliTra = Val(txtAux(7).Text)
    cLo.SubLinea = 0
    cLo.tipoMov = 1
    
    
    cLo.InsertarLote
    Set cLo = Nothing
End Sub

Private Sub ImpEtiquetas()
Dim Col As Collection
        
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    Set Col = New Collection
    Col.Add CStr(Data1.Recordset!ID)
    
    ImpirmirEtiquetas2 Col, txtAux(7).Text & " " & txtAux2(7).Text, True, 1
    Set Col = Nothing
    
End Sub
