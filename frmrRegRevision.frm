VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegRevision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de revision"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10830
   ClipControls    =   0   'False
   Icon            =   "frmrRegRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   6960
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Desc|T|N|||srevaspectos|denominacion|||"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Area|N|N|||srevaspectos|orden|0|S|"
      Text            =   "numlote"
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   2535
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
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Area|N|N|||srevaspectos|codarea|0|S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   5325
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9675
      TabIndex        =   4
      Top             =   5325
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9675
      TabIndex        =   5
      Top             =   5325
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
      Top             =   5280
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
      Bindings        =   "frmrRegRevision.frx":000C
      Height          =   4430
      Left            =   240
      TabIndex        =   6
      Top             =   710
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7805
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
      TabIndex        =   8
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
Attribute VB_Name = "frmRegRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB1 As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB1.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid





Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    Indicador = txtAux(0).Text
                    CadenaBusqueda = " AND  codarea=" & txtAux(0).Text
                    CargaGrid True
                    BotonAnyadir
                    txtAux(0).Text = Indicador
                    Indicador = ""
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
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
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub





Private Sub cmdAux_Click(Index As Integer)
              
        
        Screen.MousePointer = vbHourglass
        CadenaBusqueda = ""
        Set frmB1 = New frmBuscaGrid
        frmB1.vCampos = "Código|srevarea|codigo|T||10·Descripción|srevarea|Descripcion|T||70·"
        frmB1.vTabla = "srevarea"
        frmB1.vSQL = ""
        frmB1.vCargaFrame = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB1.vDevuelve = "0|1|"
        frmB1.vTitulo = "Areas revision"
        frmB1.vselElem = 0
        frmB1.vConexionGrid = conAri
        frmB1.Show vbModal
        Set frmB1 = Nothing
        If CadenaBusqueda <> "" Then
            txtAux(0).Text = RecuperaValor(CadenaBusqueda, 1)
            txtAux2(0).Text = RecuperaValor(CadenaBusqueda, 2)
            PonerFoco txtAux(0)
            CadenaBusqueda = ""
        End If
End Sub

Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
   
           
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
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
    End With

    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    

    Ordenacion = " ORDER BY codarea,orden "
    CadenaConsulta = MontaSQLCarga2(True)
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta & Ordenacion
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga2(enlaza) & Ordenacion
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    
    tots = "S|txtAux(0)|T|Area|600|;S|cmdAux(0)|B||0|;S|txtAux2(0)|T|Descripcion|4100|;S|txtAux(1)|T|Orden|700|;S|txtAux(2)|T|Aspecto|4200|;"
    arregla tots, DataGrid1, Me

    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
    txtAux2(0).Height = Me.DataGrid1.RowHeight
    txtAux2(0).Top = alto
    txtAux2(0).visible = b
    
    
    For jj = 0 To Me.cmdAux.Count - 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = b
    Next jj
End Sub








Private Sub frmB1_Selecionado(CadenaDevuelta As String)
    CadenaBusqueda = CadenaDevuelta
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
        Case 10 'Imprimir
'            BotonImprimir
            
        Case 11  'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
    
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
                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    Me.cmdAux(0).Enabled = (Modo <> 4)
                      
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

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
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
    
    b = ((Modo >= 3))
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga2(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "select codigo,descripcion,orden,denominacion from srevarea,srevaspectos where codarea = codigo"
    If enlaza Then
        If CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND codarea = -1"
    End If
   
    MontaSQLCarga2 = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar

    If Modo <> 1 Then
        CadenaBusqueda = ""
        LimpiarCampos
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
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next


    LimpiarCampos
    CadenaBusqueda = ""
    CadenaConsulta = MontaSQLCarga2(True) & Ordenacion
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1

    
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
    'codartic
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")
    txtAux2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
    'numlote
    txtAux(1).Text = DBLet(Me.DataGrid1.Columns(2).Value, "T")
    'fecha entrada
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "F")
    FormateaCampo txtAux(2)
    

    

    DataGrid1.Enabled = False
    PonerFoco txtAux(1)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar
    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
    
    SQL = "¿Seguro que desea eliminar el registro?" & vbCrLf
    SQL = SQL & vbCrLf & "Area: " & Data1.Recordset.Fields(0).Value & "  " & Data1.Recordset.Fields(1).Value
    SQL = SQL & vbCrLf & "Rev: " & Data1.Recordset.Fields(2).Value & "  " & Data1.Recordset.Fields(3).Value
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = "Delete from srevaspectos WHERE codarea=" & DBSet(Data1.Recordset!Codigo, "N") & " AND orden=" & DBSet(Data1.Recordset!Orden, "N")
        
        Conn.Execute SQL
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
        SituarDataPosicion Me.Data1, NumRegElim, SQL
    End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Nº de lote", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cArt As CArticulo

    On Error GoTo ErrDatosOK

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    
    'comprobar que el articulo insertado tiene control de numero de serie
    If Modo = 3 Then
        Set cArt = New CArticulo
        If cArt.LeerDatos(txtAux(0).Text) Then
            If Not cArt.TieneNumLote Then
                b = False
                MsgBox "El artículo no tiene control de nº de lote.", vbInformation
            End If
        End If
        Set cArt = Nothing
    End If
    
    DatosOk = b
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
'    cad = cad & ParaGrid(txtAux(0), 20, "Código")
'    cad = cad & ParaGrid(txtAux(1), 80, "Descripción")
'    Tabla = NombreTabla
'    Titulo = "Tipos de Contrato"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco txtAux(kCampo)
''            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = MontaSQLCarga2(True)
        CadenaConsulta = CadenaConsulta & " AND " & cadB & Ordenacion
        CadenaBusqueda = " AND " & cadB
        PonerCadenaBusqueda
        CadenaBusqueda = ""
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10
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
                If Index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub Siguiente()
Dim SQL As String
    On Error GoTo Si
    
    Set miRsAux = New ADODB.Recordset
    
    If txtAux(0).Text <> "" Then
        SQL = "select descripcion,orden from srevarea left join srevaspectos on codarea = codigo WHERE codigo = " & txtAux(0).Text & " order by orden desc"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "No existe area: " & txtAux(0).Text, vbExclamation
            SQL = ""
        Else
            'Si esta
            txtAux2(0).Text = miRsAux!Descripcion
            txtAux(1).Text = DBLet(miRsAux!Orden, "N") + 1
            PonerFoco txtAux(2)
        End If
        miRsAux.Close
    Else
        SQL = ""
    End If
   
    If SQL = "" Then
        If txtAux(0).Text <> "" Then PonerFoco txtAux(0)
        txtAux(0).Text = ""
        txtAux2(0).Text = ""
        txtAux(1).Text = ""
    End If
Si:
    
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub
Private Sub txtAux_LostFocus(Index As Integer)
    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod. Articulo
            If Modo = 3 Then Siguiente
            
        Case 1 'cantidad entrada
            PonerFormatoEntero txtAux(Index)
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub

