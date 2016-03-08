VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProdEtiquetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas cajas"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "frmProdEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   4920
      TabIndex        =   9
      Text            =   "Dat"
      Top             =   5040
      Width           =   1755
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Tag             =   "fichero|T|N|||prodparametiq|archivo||N|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   1755
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código Marca|N|N|0|9999|prodparametiq|codmarca||S|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      HelpContextID   =   35
      Index           =   1
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||prodparametiq|descripcion|||"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProdEtiquetas.frx":000C
      Height          =   4710
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   540
      Width           =   7935
      _ExtentX        =   13996
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
      Left            =   7080
      TabIndex        =   4
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   5940
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
         TabIndex        =   7
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   7815
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
Attribute VB_Name = "frmProdEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Dim FormatoCod As String 'formato del campo de codigo
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
Dim B As Boolean
    
    Modo = vModo
    B = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = Not B
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    txtAux(3).visible = Not B
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B

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
Dim B As Boolean

    B = (Modo = 2)
    
    'Buscar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = B
    Me.mnVerTodos.Enabled = B
   
    B = B And vUsu.Nivel = 0
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    

End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
      
    anc = ObtenerAlto(DataGrid1, 10)
    
    'Obtenemos la siguiente numero de código de Marca
    txtAux(0).Text = "": txtAux(1).Text = ""
    txtAux(2).Text = "": txtAux(3).Text = ""
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "prodparametiq.codmarca= -1"
    'Buscar
    limpiar Me
    LLamaLineas 770, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
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
Dim Cad As String
Dim anc As Single
Dim I As Integer
On Error GoTo EModificar

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    Cad = ""
    For I = 0 To 1
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    LLamaLineas anc, 4
   
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

    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 45
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 45
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    '### a mano
    SQL = "¿Seguro que desea eliminar la etiqueta?" & vbCrLf
    SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "MARCA: " & DBLet(adodc1.Recordset.Fields(3), "T")
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = "Delete from prodparametiq where codmarca=" & adodc1.Recordset!codmarca
        Conn.Execute SQL
        CancelaADODC adodc1
        CargaGrid ""
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar etiqueta", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3   'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CancelaADODC Me.adodc1
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
                DataGrid1.SetFocus
            End If
            
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                'Encuentra registros
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    DataGrid1.SetFocus
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1 'Buscar
        .Buttons(2).Image = 2 'VerTodos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4  'Modificar
        .Buttons(7).Image = 5  'Eliminar
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields

    CadAncho = False
    PonerModo 2
      
    Label1.Caption = "El archivo tiene que existir en el servidor impresion dentro de \ariges\etiquetas      "
      
    'Cadena consulta
    'prodparametiq  codmarca descripcion archivo tipo
    CadenaConsulta = "Select prodparametiq.codmarca,descripcion,archivo,nommarca from prodparametiq left join smarca on prodparametiq.codmarca=smarca.codmarca"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
        Case 11: mnSalir_Click
    End Select
End Sub



Private Sub CargaGrid(Optional SQL As String)
Dim I As Byte
Dim B As Boolean
    
    B = DataGrid1.Enabled

    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codmarca"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Id"
        DataGrid1.Columns(I).Width = 800
        DataGrid1.Columns(I).NumberFormat = FormatoCod
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Descripcion"
        DataGrid1.Columns(I).Width = 1930
            
    I = 2
        DataGrid1.Columns(I).Caption = "Archivo"
        DataGrid1.Columns(I).Width = 1930
            
    I = 3
        DataGrid1.Columns(I).Caption = "Marca"
        DataGrid1.Columns(I).Width = 2330
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
' txtAux(0).Width = DataGrid1.Columns(0).Width - 60
' txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        
        For I = 0 To DataGrid1.Columns.Count - 1
            txtAux(I).Width = DataGrid1.Columns(I).Width - 60
        Next I
        
        CadAncho = True
    End If
   
   'No permitir cambiar tamaño de columnas
   For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
   Next I
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        mnEliminar.Enabled = Not adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = B
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
    If Index = 0 Then
        txtAux(3).Text = ""
        If PonerFormatoEntero(txtAux(Index)) Then
            txtAux(3).Text = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", txtAux(0).Text)
            If txtAux(3).Text = "" Then
                MsgBox "No existe la marca: " & txtAux(0).Text, vbExclamation
                txtAux(0).Text = ""
                PonerFoco txtAux(0)
            End If
        Else
            txtAux(0).Text = ""
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
    'Comprobar si ya existe el cod de marca en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then B = False
    End If
    
    DatosOk = B
End Function


Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
