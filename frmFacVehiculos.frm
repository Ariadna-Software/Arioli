VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacVehiculos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehiculos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15885
   Icon            =   "frmFacVehiculos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   6
      Left            =   12000
      MaxLength       =   15
      TabIndex        =   14
      Tag             =   "Remolque|T|S|||svehiculos|matricularemolque||N|"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   10320
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "DNIConductor|T|S|||svehiculos|DNIConductor||N|"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   8640
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "C|T|S|||svehiculos|Conductor||N|"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   7080
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Empresa|T|S|||svehiculos|Empresa||N|"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "Descripcion|T|N|||svehiculos|matricula||N|"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "C�digo|N|N|0||svehiculos|codigo|0|S|"
      Text            =   "Dat"
      Top             =   4920
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||svehiculos|descripcion||N|"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   3075
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacVehiculos.frx":000C
      Height          =   4710
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   540
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8308
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6060
      TabIndex        =   7
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4860
      TabIndex        =   6
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6060
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   9
      Top             =   5340
      Width           =   1755
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
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15885
      _ExtentX        =   28019
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
Attribute VB_Name = "frmFacVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

Private CadenaConsulta As String

Dim Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
     
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(2).visible = Not b
    
    txtAux(3).visible = vParamAplic.QUE_EMPRESA = 4 And Not b
    txtAux(4).visible = vParamAplic.QUE_EMPRESA = 4 And Not b
    txtAux(5).visible = vParamAplic.QUE_EMPRESA = 4 And Not b
    txtAux(6).visible = vParamAplic.QUE_EMPRESA = 4 And Not b
    
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    
    PonerFoco txtAux(1)
    DataGrid1.Enabled = b

    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b

    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    b = b And Not DeConsulta
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(10).Enabled = b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
     PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
   
    anc = ObtenerAlto(Me.DataGrid1)

    txtAux(0).Text = SugerirCodigoSiguienteStr("svehiculos", "codigo")
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    txtAux(5).Text = ""
    txtAux(6).Text = ""
    
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(1)
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    CargaGrid "codigo= -1"  'para vaciar los datos del Grid
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    txtAux(5).Text = ""
    txtAux(6).Text = ""
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc, 1
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ning�n registro en la tabla de vehiculos", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim anc As Single
Dim I As Integer
    
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        txtAux(3).Text = DataGrid1.Columns(3).Text
        txtAux(4).Text = DataGrid1.Columns(4).Text
        txtAux(5).Text = DataGrid1.Columns(5).Text
        txtAux(6).Text = DataGrid1.Columns(5).Text
    End If
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc, 4
   
   'Como es modificar
'    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim jj As Byte
Dim b As Boolean
Dim N As Byte



    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    b = (xModo = 3 Or xModo = 4 Or xModo = 1) 'Insertar o Modificar Lineas
    N = 2
    If vParamAplic.QUE_EMPRESA = 4 Then N = 6
    For jj = 0 To N
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
  
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    'Comprobaremos que no esta vinculado a ninguna hoja de rutas
    SQL = DevuelveDesdeBD(conAri, "vehiculo", "sRepartoC", "vehiculo", CStr(Adodc1.Recordset.Fields(0)))
    If SQL <> "" Then
        MsgBox "No puede eliminar el vehiculo", vbExclamation
        Exit Sub
    End If
    
    '### a mano
    SQL = "�Seguro que desea eliminar el vehicul0?" & vbCrLf
    SQL = SQL & vbCrLf & "C�digo: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descrp: " & Adodc1.Recordset.Fields(1) & "    " & DBLet(Adodc1.Recordset!matricula, "T")
        
        
        
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from svehiculos where codigo=" & DBSet(Adodc1.Recordset!Codigo, "N")
        conn.Execute SQL
        CancelaADODC Me.Adodc1
        CargaGrid ""
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Ubicaciones", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3  'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
            
        Case 4  'Modificar
             If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                      TerminaBloquear
                      I = Adodc1.Recordset.Fields(0)
'                      LLamaLineas Modo, 0
                      PonerModo 2
                      CancelaADODC Me.Adodc1
                      CargaGrid
                      Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & I)
                  End If
                  DataGrid1.SetFocus
            End If
            
        Case 1  'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Adodc1.Recordset.EOF Then
                Adodc1.Recordset.MoveFirst
            Else
                Me.lblIndicador.Caption = ""
            End If
        Case 4 'Modificar
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
            TerminaBloquear
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
ECancelar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
On Error GoTo ERegresar

    If Adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
ERegresar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim K As Integer

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Busqueda
        .Buttons(2).Image = 2   'Bot�n Recuperar Todos
        .Buttons(5).Image = 3   'Bot�n A�adir Nuevo Registro
        .Buttons(6).Image = 4   'Bot�n Modificar Registro
        .Buttons(7).Image = 5   'Bot�n Borrar Registro
        .Buttons(10).Image = 16  'Bot�n Imprimir
        .Buttons(11).Image = 15  'Bot�n Salir
    End With
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    K = 7305
    If vParamAplic.QUE_EMPRESA = 4 Then K = 15220
    Me.Width = K
    
    
    Me.DataGrid1.Width = K - 330
    Me.cmdAceptar.Left = K - 2445
    Me.cmdCancelar.Left = K - 1245
    Me.cmdRegresar.Left = cmdCancelar.Left
    
    
    'Cadena consulta
    CadenaConsulta = "Select * from svehiculos"
    CargaGrid
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
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
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5: mnNuevo_Click
        Case 6: mnModificar_Click
        Case 7: mnEliminar_Click
        Case 10 'Bot�n Imprimir Listado
'                Me.Hide
'                AbrirListado (110) 'Opci�n 110 de los Listados
'                Me.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim b As Boolean
Dim tots As String

    b = DataGrid1.Enabled
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codigo"
    
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False
    
    tots = "S|txtAux(0)|T|Cod.|700|;S|txtAux(1)|T|Nombre|3150|;S|txtAux(2)|T|Matricula|1450|;"
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        tots = tots & "S|txtAux(3)|T|Empresa|2900|;S|txtAux(4)|T|Conductor|3350|;S|txtAux(5)|T|DNI|1350|;S|txtAux(6)|T|Remolque|1350|;"
    Else
        tots = tots & "N|||||;N|||||;N|||||;N|||||;"
    End If
    arregla tots, DataGrid1, Me
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
   
   
   
   
   
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    If Index = 0 Then
        If Not PonerFormatoEntero(txtAux(0)) Then
            txtAux(0).Text = ""
            If Modo = 3 Then PonerFoco txtAux(0)
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de Almacen en la tabla y no esperar al btn Aceptar
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
