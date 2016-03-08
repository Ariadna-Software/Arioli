VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComPedLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lotes"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmComPedLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   7
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "Dat"
      Top             =   4080
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmComPedLotes.frx":000C
      Height          =   3270
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5768
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
      Left            =   8400
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7140
      TabIndex        =   3
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Volver"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   9
      Top             =   4380
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
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
   Begin VB.Label lblDesc 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   8415
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
Attribute VB_Name = "frmComPedLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vPedido As Long
Public vProve As String
Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos


Dim Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas



Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = False
    txtAux(1).visible = False
    txtAux(2).visible = False
    txtAux(3).visible = Not b
    txtAux(4).visible = Not b
    txtAux(5).visible = Not b
    cmdAceptar.visible = Not b
    
    DataGrid1.Enabled = b
    cmdCancelar.visible = True
    cmdRegresar.visible = Modo = 2
    
    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = False
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
   
    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = False
    Me.mnNuevo.Enabled = False
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(7).Enabled = False
    Me.mnEliminar.Enabled = False
    
    'Imprimir
    Toolbar1.Buttons(10).Enabled = False
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub





Private Sub BotonBuscar()
    CargaGrid
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""

    LLamaLineas 770, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ningún registro en la tabla smarca", vbInformation
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
Dim C As String
On Error GoTo EModificar

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 0)
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    txtAux(4).Text = DataGrid1.Columns(4).Text
    txtAux(5).Text = DataGrid1.Columns(5).Text
    
    
    
    C = DevuelveDesdeBD(conAri, "if(factorconversion<1,1,0)", "sartic", "codartic", adodc1.Recordset!codartic, "T")
    I = Val(C)
    
    
    LLamaLineas anc, 4
    BloquearTxt txtAux(1), True
    BloquearTxt txtAux(2), True
    PonerFoco txtAux(3)
    
    txtAux(5).visible = I > 0
    
    
    Screen.MousePointer = vbDefault
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar", Err.Description
End Sub




Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(4).Top = alto
    txtAux(5).Top = alto
  
End Sub





Private Sub cmdAceptar_Click()
Dim I As Integer
'Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3   'INSERTAR
           
        Case 4  'MODIFICAR
            If DatosOk Then
                
                If InsertarModificar_() Then
                   
                    I = adodc1.Recordset.AbsolutePosition
                    PonerModo 2
                    CancelaADODC Me.adodc1
                    CargaGrid
                    If I >= adodc1.Recordset.RecordCount Then
                        DataGrid1.SetFocus
                    Else
                        adodc1.Recordset.Move I
                        BotonModificar
                    End If
                 End If
            End If
            
        Case 1 'BUSQUEDA
'            cadB = ObtenerBusqueda(Me, False)
'            If cadB <> "" Then
'                'Encuentra registros
'                PonerModo 2
'                CargaGrid
'                DataGrid1.SetFocus
'            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    Me.lblIndicador.Caption = ""
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        Case 1 'Buscar
            CargaGrid
        Case 2
            'Cancela LA edicion. NO debe seguir adelante  con le proceso
            Unload Me
            Exit Sub
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
End Sub







Private Sub cmdRegresar_Click()
    CadenaDesdeOtroForm = ""
    If LotesCorrectos Then Unload Me
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()

    If Me.lblDesc.Caption = "" Then
     
        Me.lblDesc.Caption = "Pedido: " & Me.vPedido & " de " & vProve
        
        If Len(Me.lblDesc.Caption) > 48 Then
            Me.lblDesc.FontSize = 11
        Else
            Me.lblDesc.FontSize = 12
        End If
            adodc1.Recordset.MoveFirst
            BotonModificar
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
     '   .Buttons(1).Image = 1 'Buscar
     '   .Buttons(2).Image = 2 'VerTodos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4  'Modificar
        .Buttons(7).Image = 5  'Eliminar
    '    .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    

    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields

    CadAncho = False
    cmdRegresar.visible = False
    PonerModo 2
    Me.lblDesc.Caption = ""
    CargaGrid
    CadenaDesdeOtroForm = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

'Private Sub mnEliminar_Click()
'    BotonEliminar
'End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub
'
'Private Sub mnNuevo_Click()
'    BotonAnyadir True
'End Sub

Private Sub mnSalir_Click()
    If Modo <> 2 Then Exit Sub
    
    cmdRegresar_Click
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        'Case 5: mnNuevo_Click
        Case 6: mnModificar_Click
        'Case 7: mnEliminar_Click
        Case 10   'Imprimir Listado de Marcas
'            Me.Hide
'            AbrirListado (1) 'OpcionListado=1
'            Me.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub

Private Function DevWHERE() As String
    DevWHERE = " "
End Function

Private Sub CargaGrid()
Dim I As Byte
Dim b As Boolean
Dim SQL As String
    b = DataGrid1.Enabled

    
    SQL = "select slippr.numlinea ,nomartic,slipprlotes.cantidad,numlote,etiquetas,deposito,codartic from slippr,slipprlotes where slippr.numpedpr=slipprlotes.numpedpr and"
    SQL = SQL & " slippr.numlinea = slipprlotes.numlinea AND  slippr.numpedpr = " & Me.vPedido

    
    SQL = SQL & " ORDER BY  slipprlotes.numlinea"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    
    'Nombre producto
    
    DataGrid1.Columns(0).visible = False
    
    I = 1
        DataGrid1.Columns(I).Caption = "Articulo"
        DataGrid1.Columns(I).Width = 3100
        
    I = 2
        DataGrid1.Columns(I).Caption = "Cantidad"
        DataGrid1.Columns(I).Width = 1150
        DataGrid1.Columns(I).NumberFormat = FormatoCantidad
        DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 3
        DataGrid1.Columns(I).Caption = "Lote"
        DataGrid1.Columns(I).Width = 2000
        DataGrid1.Columns(I).Alignment = dbgCenter
            
    I = 4
        DataGrid1.Columns(I).Caption = "Etiquetas"
        DataGrid1.Columns(I).Width = 950
        DataGrid1.Columns(I).Alignment = dbgRight
            
    I = 5
        DataGrid1.Columns(I).Caption = "Depósito"
        DataGrid1.Columns(I).Width = 1250
        DataGrid1.Columns(I).Alignment = dbgRight
        
    I = 6 'cdoartic
        DataGrid1.Columns(I).visible = False
            
    'Fiajamos el cadancho
    If Not CadAncho Then
       ' txtAux(1).Left = DataGrid1.Columns(1).Left + 120
       ' 'La primera vez fijamos el ancho y alto de  los txtaux
       ' txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        
       ' 'cantidad
       ' txtAux(2).Left = DataGrid1.Columns(2).Left + 120
       ' txtAux(2).Width = DataGrid1.Columns(2).Width
        
        'cantidad
        txtAux(3).Left = DataGrid1.Columns(3).Left + 120
        txtAux(3).Width = DataGrid1.Columns(3).Width - 30
        
        txtAux(4).Left = DataGrid1.Columns(4).Left + 120
        txtAux(4).Width = DataGrid1.Columns(4).Width - 15
        
        txtAux(5).Left = DataGrid1.Columns(5).Left + 120
        txtAux(5).Width = DataGrid1.Columns(5).Width - 30
        
        
        CadAncho = True
    End If
   
   'No permitir cambiar tamaño de columnas
   For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
   Next I
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        'Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        'mnEliminar.Enabled = Not adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
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
    Select Case Index
    Case 0
        PonerFormatoEntero txtAux(Index) 'codmarca
    Case 2
        If txtAux(Index).Text <> "" Then
            If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
    Case 5
        If PonerFormatoEntero(txtAux(Index)) Then
            If Val(txtAux(Index).Text) < 1 Or Val(txtAux(Index).Text) > MaxNumDepositos_ Then
                If Val(DevuelveDesdeBD(conAri, "numDeposito", "proddepositos", "numdeposito", txtAux(Index).Text)) = 0 Then
                    MsgBox "Nº Deposito incorrecto", vbExclamation
                    txtAux(Index).Text = ""
                End If
            End If
        Else
            txtAux(Index).Text = ""
        End If
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim cDep As cDeposito
Dim Cantidad As Currency

    DatosOk = False
    If txtAux(3).Text = "" Then
        MsgBox "Nº lote obligado", vbExclamation
        PonerFoco txtAux(3)
        Exit Function
    End If
    
    If txtAux(4).Text = "" Then
        MsgBox "Nº etiquetas obligado (Uno valor por defecto)", vbExclamation
        PonerFoco txtAux(4)
        Exit Function
    End If
    
    
    DatosOk = True
    'Si tiene deposito, tenemos que ver que esta asignado
    If txtAux(5).visible Then
        CadenaConsulta = ""
        If txtAux(5).Text = "" Then
            CadenaConsulta = "Numero de deposito obligado"
        Else
            Set cDep = New cDeposito
            If cDep.LeerDatos(CInt(Val(txtAux(5).Text)), False) Then
                
                'Vemos si esta vacio
                If cDep.NUmlote <> "" Then
                    CadenaConsulta = "El deposito NO esta vacio"
                Else
                    'Factor conversion
                    CadenaConsulta = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", adodc1.Recordset!codartic, "T")
                    If CadenaConsulta = "" Then
                        CadenaConsulta = "Error leyendo factor conversion"
                    Else
                        If ImporteFormateado(CadenaConsulta) >= 1 Then
                            CadenaConsulta = "Error factor conversion: " & CadenaConsulta
                        Else
                            Cantidad = ImporteFormateado(CadenaConsulta)
                            'Vemos si cabe
                            If Cantidad = 0 Then
                                CadenaConsulta = "Error factor conversion. =0"
                            Else
                                Cantidad = adodc1.Recordset!Cantidad / Cantidad
                                If Cantidad > cDep.Capacidad Then
                                    CadenaConsulta = "Cantidad excede de la capacidad del deposito"
                                
                                Else
                                    'POR FIN, todo bien
                                    CadenaConsulta = ""
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                CadenaConsulta = "Error leyendo deposito"
            End If
            Set cDep = Nothing
        End If
        
        If CadenaConsulta <> "" Then
            MsgBox CadenaConsulta, vbExclamation
            PonerFoco txtAux(5)
            DatosOk = False
        End If
        CadenaConsulta = ""
    End If
    
    
End Function




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub






Private Function InsertarModificar_() As Boolean
Dim C As String

    InsertarModificar_ = False
    C = "update `slipprlotes` set numlote=" & DBSet(txtAux(3).Text, "T") & ""
    C = C & ", etiquetas=" & DBSet(txtAux(4).Text, "N") & ""
    C = C & ", deposito ="
    If Me.txtAux(5).visible Then
        C = C & txtAux(5).Text
    Else
        C = C & "NULL"
    End If
    C = C & " where `numpedpr`=" & Me.vPedido & " AND `numlinea`=" & Me.adodc1.Recordset!numlinea
    
    If EjecutaSQL(conAri, C, True) Then
        Espera 0.15
        InsertarModificar_ = True
    End If
    
End Function



Private Function LotesCorrectos() As Boolean

    Set miRsAux = New ADODB.Recordset
    LotesCorrectos = True
    miRsAux.Open "Select numlote from slipprlotes where numpedpr=" & Me.vPedido, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Trim(DBLet(miRsAux!NUmlote, "T")) = "" Then LotesCorrectos = False
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If LotesCorrectos Then
        'Pongo la marca de correcot
        CadenaDesdeOtroForm = "OK"
        Exit Function
    End If
    
    'Si no esta bien preguntar si cont
    If MsgBox("No estan indicados todos los lotes. Salir y cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then LotesCorrectos = True
End Function
