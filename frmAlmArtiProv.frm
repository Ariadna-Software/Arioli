VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmArtiProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arti prov"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   7635
   Icon            =   "frmAlmArtiProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdprov 
      Caption         =   "+"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   135
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "Precio|N|N|||sarti5|precio|#,##0.0000|N|"
      Text            =   "Dato2"
      Top             =   1920
      Width           =   1155
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   720
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "Código|T|N|||sarti5|codprove|000|S|"
      Text            =   "Dat"
      Top             =   1920
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   4
      Tag             =   "Proveedor|T|N|||sarti5|nomprove||N|"
      Text            =   "Dato2"
      Top             =   1920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmArtiProv.frx":000C
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
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
      Left            =   6240
      TabIndex        =   3
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   2520
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
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   120
      Top             =   840
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
End
Attribute VB_Name = "frmAlmArtiProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public articulo2 As String  'codartic|nomartic|


Private CadenaConsulta As String
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

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
Dim B As Boolean

    Modo = vModo
    B = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
     
    txtAux(0).visible = Not B
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    cmdprov.visible = Modo = 3
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    
    
    DataGrid1.Enabled = B

    cmdRegresar.visible = False

    
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    BloquearTxt txtAux(1), True 'siempre
    
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
    'Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = B
    'Me.mnVerTodos.Enabled = b
    

    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    'Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    'Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(7).Enabled = B
    'Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(10).Enabled = B
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
    
    anc = ObtenerAlto(Me.DataGrid1)

    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


'Private Sub BotonBuscar()
'Dim anc As Single
'
'    CargaGrid "codubica= -1"  'para vaciar los datos del Grid
'    'Buscar
'    txtAux(0).Text = ""
'    txtAux(1).Text = ""
'
'    anc = ObtenerAlto(Me.DataGrid1)
'    LLamaLineas anc, 1
'    PonerFoco txtAux(0)
'End Sub

Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ningún registro en la tabla sarti5", vbInformation
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
    
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

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
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc, 4
   
   'Como es modificar
'    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim jj As Byte
Dim B As Boolean

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    B = (xModo = 3 Or xModo = 4 Or xModo = 1) 'Insertar o Modificar Lineas
    
    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = B
    Next jj
    Me.cmdprov.Top = alto
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    SQL = "¿Seguro que desea eliminar el precio del proveeedor para el articulo?" & vbCrLf
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Proveedor: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sarti5 where codprove=" & DBSet(adodc1.Recordset!codProve, "N")
        SQL = SQL & " AND codartic = " & DBSet(articulo2, "T")
        Conn.Execute SQL
        CancelaADODC Me.adodc1
        CargaGrid ""
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
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
                If InsertarModificar Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
            
        Case 4  'Modificar
             If DatosOk Then
                 If InsertarModificar Then
                      TerminaBloquear
                      I = adodc1.Recordset.Fields(0)
'                      LLamaLineas Modo, 0
                      PonerModo 2
                      CancelaADODC Me.adodc1
                      CargaGrid
                      adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " = " & I)
                  End If
                  DataGrid1.SetFocus
            End If
            
        Case 1  'HacerBusqueda
'            cadB = ObtenerBusqueda(Me, False)
'            If cadB <> "" Then
'                PonerModo 2
'                CargaGrid cadB
'                DataGrid1.SetFocus
'            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
            TerminaBloquear
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
ECancelar:
    If Err.Number <> 0 Then Err.Clear
End Sub






Private Sub cmdprov_Click()
    Me.Tag = ""
    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Código|sprove|codprove|T|000|25·Nombre|sprove|nomprove|T||60·"
    frmB.vTabla = "sprove"
    frmB.vTitulo = "Proveedores"
    frmB.vConexionGrid = conAri     'Conexión a BD: Ariges
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 0
    frmB.Show vbModal
    If Me.Tag <> "" Then
        txtAux(0).Text = RecuperaValor(Me.Tag, 1)
        txtAux(1).Text = RecuperaValor(Me.Tag, 2)
        PonerFoco txtAux(2)
    End If
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
        .Buttons(1).Image = 1   'Botón Busqueda
        .Buttons(2).Image = 2   'Botón Recuperar Todos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4   'Botón Modificar Registro
        .Buttons(7).Image = 5   'Botón Borrar Registro
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    

    PonerModo 2
    'Cadena consulta
    CadenaConsulta = "Select sarti5.codprove,nomprove,precio from sarti5,sprove where sarti5.codprove=sprove.codprove AND codartic = " & DBSet(articulo2, "T")
    CargaGrid
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

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Me.Tag = CadenaDevuelta
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
        Case 2: mnVerTodos_Click
        Case 5: mnNuevo_Click
        Case 6: mnModificar_Click
        Case 7: mnEliminar_Click
        Case 10 'Botón Imprimir Listado
                Me.Hide
                AbrirListado (110) 'Opción 110 de los Listados
                Me.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim B As Boolean
Dim tots As String

    B = DataGrid1.Enabled
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codprove"
    
    CargaGridGnral DataGrid1, Me.adodc1, SQL, False
    
    tots = "S|txtAux(0)|T|Cod.|1000|;S|txtAux(1)|T|Proveedor|4550|;S|txtAux(2)|T|Precio|1250|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
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
        txtAux(1).Text = ""
        If PonerFormatoEntero(txtAux(Index)) Then txtAux(1).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(0).Text)
        
    ElseIf Index = 2 Then
        PonerFormatoDecimal txtAux(Index), 2
    End If
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    
    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
 
    
    DatosOk = B
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function InsertarModificar() As Boolean
Dim SQL As String
On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If Modo = 3 Then
        'INSERT
        SQL = "INSERT INTO sarti5 (codartic,codprove,precio) VALUES (" & DBSet(articulo2, "T") & "," & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(2).Text, "N") & ")"
    Else
        SQL = "UPDATE sarti5 Set precio=" & DBSet(txtAux(2).Text, "N") & " WHERE codartic=" & DBSet(articulo2, "T") & " AND codprove= " & DBSet(txtAux(2).Text, "N")
    End If
    Conn.Execute SQL
    'Aqui... todo bien
    InsertarModificar = True
    
EInsertarModificar:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function
