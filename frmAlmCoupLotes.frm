VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmCoupLotes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmAlmCoupLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   2
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
      TabIndex        =   0
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
      TabIndex        =   1
      Text            =   "Dato2"
      Top             =   4080
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCoupLotes.frx":000C
      Height          =   3270
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5175
      _ExtentX        =   9128
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
      Left            =   4200
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3060
      TabIndex        =   3
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Volver"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   6
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
         TabIndex        =   7
         Top             =   240
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   5175
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
Attribute VB_Name = "frmAlmCoupLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vIdCup As Long
Public vCodAlmac As Integer
Public vCantidad As Currency   'Cantidad de la linea de produccion
Public vCodartic As String     'articulo para buscar en los lotes

    'insert into `olicoupagelinlotes` (`codigo`,`codartic`,`linea`,`numlote`,`cantlote`)

    'Aqui no mueve ne partidas. Lo hara cuando cierra la produccion.
    '   Entonces YA no podra tocar las lineas
    

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

Dim Suma As Currency  'Para saber lo que suma las lineas

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = False
    txtAux(1).visible = Not b
    txtAux(2).visible = Not b
    Me.Combo1.visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b

    cmdRegresar.visible = Modo = 2
    
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

    b = False
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
   
    b = (Modo = 2)
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


Private Sub BotonAnyadir(DesdeToolbar As Boolean)
Dim anc As Single
    
    
    
    'Obtenemos la siguiente numero de c�digo de Marca
    Set miRsAux = New ADODB.Recordset
    CadenaConsulta = "select sum(cantlote) total ,max(linea) ultimo from olicoupagelinlotes " & DevWHERE
    miRsAux.Open CadenaConsulta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        'NInugni, no hay ninguno
        NumRegElim = 0
        Suma = 0
        
    Else
        NumRegElim = DBLet(miRsAux!ultimo, "N")
        Suma = DBLet(miRsAux!total, "N")
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Suma = vCantidad Then
        Me.DataGrid1.AllowAddNew = False
        If DesdeToolbar Then
            MsgBox "No se pueden a�adir mas lineas", vbExclamation
            
        Else
            PonerModo 2
        End If
        Exit Sub
    End If
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
      
    anc = ObtenerAlto(DataGrid1, 10)
    
    
    txtAux(0).visible = False
    CadenaConsulta = ""
    txtAux(1).Text = ""
    NumRegElim = NumRegElim + 1
    txtAux(0).Text = NumRegElim
    txtAux(2).Text = Format(vCantidad - Suma, FormatoCantidad)
    Combo1.ListIndex = 0
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(1)
End Sub


Private Sub BotonBuscar()
    CargaGrid
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    Combo1.ListIndex = -1
    
    LLamaLineas 770, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ning�n registro en la tabla smarca", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()

Dim anc As Single
Dim i As Integer
On Error GoTo EModificar

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    
    If DataGrid1.Columns(3) = "SI" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    LLamaLineas anc, 4
    
    Suma = TotalLineas
    
    Screen.MousePointer = vbDefault
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar", Err.Description
End Sub

Private Function TotalLineas() As Currency
Dim cad As String
    cad = DevWHERE
    cad = Mid(cad, 8)
    cad = cad & " AND 1 "
    cad = DevuelveDesdeBD(conAri, "sum(cantlote)", "olicoupagelinlotes", cad, "1")
    If cad = "" Then cad = "0"
    TotalLineas = CCur(cad)
End Function


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    Me.Combo1.Top = alto - 30
  '  txtAux(0).Left = DataGrid1.Left + 340
  '  txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    

    '### a mano
    SQL = "�Seguro que desea eliminar el numero de lote "
    
    SQL = SQL & adodc1.Recordset.Fields(1) & "?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = DevWHERE
        SQL = "Delete FROM olicoupagelinlotes " & SQL & " AND linea = " & adodc1.Recordset!linea
        Conn.Execute SQL
        CancelaADODC adodc1
        CargaGrid
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Marca", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3   'INSERTAR
            If DatosOk Then
                '   notinhg , lote   cantidad   linea
                If InsertarModificar_() Then
                    CargaGrid
                    BotonAnyadir False
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                
                If InsertarModificar_() Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CancelaADODC Me.adodc1
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
                DataGrid1.SetFocus
            End If
            
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                'Encuentra registros
                PonerModo 2
                CargaGrid
                DataGrid1.SetFocus
            End If
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
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
End Sub







Private Sub cmdRegresar_Click()

    Suma = TotalLineas
    If Suma <> vCantidad Then
        
        If MsgBox("Total lineas distinto del total de la linea de cupage" & vbCrLf & "�Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        If Suma > vCantidad Then Exit Sub
    End If
                
                
    'Updatear la descipcio de los numeros de lote
'    If Suma > 0 Then
'        UpdatearCadenaDescripcionLote
'        CadenaDesdeOtroForm = "OK"
'    End If
        
    Unload Me
End Sub


'Private Sub UpdatearCadenaDescripcionLote()
'    Adodc1.Recordset.MoveFirst
'    CadenaConsulta = ""
'    While Not Me.Adodc1.Recordset.EOF
'        CadenaConsulta = CadenaConsulta & ",  " & DevNombreSQL(Adodc1.Recordset!Numlote)
'        Adodc1.Recordset.MoveNext
'    Wend
'    If CadenaConsulta <> "" Then
'        CadenaConsulta = Trim(Mid(CadenaConsulta, 2))
'        'UPDATEAMOS LAS LINEA DE PRODUCCION
'        CadenaConsulta = "UPDATE sliordpr SET numlote = '" & DevNombreSQL(CadenaConsulta) & "' "
'        CadenaConsulta = CadenaConsulta & DevWHERE
'        EjecutaSQL conAri, CadenaConsulta
'    End If
'End Sub

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
        
        Me.lblDesc.Caption = "CUP: " & Format(Me.vIdCup, "000000") & " - " & Me.vCodartic
        Me.lblDesc.Caption = Me.lblDesc.Caption & "   Tot: " & Format(vCantidad, FormatoCantidad)
        If adodc1.Recordset.RecordCount = 0 Then
            If Not buscarNumerosLotes Then
                BotonAnyadir False
            Else
                BotonVerTodos
            End If
        Else
            adodc1.Recordset.MoveLast
            BotonModificar
        End If
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
        .Buttons(5).Image = 3   'Bot�n A�adir Nuevo Registro
        .Buttons(6).Image = 4  'Modificar
        .Buttons(7).Image = 5  'Eliminar
    '    .Buttons(10).Image = 16  'Bot�n Imprimir
        .Buttons(11).Image = 15  'Bot�n Salir
    End With
    Me.Caption = "Lotes" & Space(20) & "COUPAGE"

    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields
    DataGrid1.RowHeight = Combo1.Height
    
    CargarCombo_SiNo Me.Combo1
    
    CadAncho = False
    cmdRegresar.visible = False
    PonerModo 2
    Me.lblDesc.Caption = ""
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
    BotonAnyadir True
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
        Case 10   'Imprimir Listado de Marcas
            Me.Hide
            AbrirListado (1) 'OpcionListado=1
            Me.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub

Private Function DevWHERE() As String
    '(`codigo`,`codalmac`,`codartic
    'DevWHERE = " where codigo=" & Me.vIdCup & " AND codalmac =" & Me.vCodAlmac & " AND codartic =" & DBSet(vCodartic, "T")
    DevWHERE = " where codigo=" & Me.vIdCup & " AND codartic =" & DBSet(vCodartic, "T")
End Function

Private Sub CargaGrid()
Dim i As Byte
Dim b As Boolean
Dim SQL As String
    b = DataGrid1.Enabled

    SQL = DevWHERE
    ''sliordprlotes` (`codigo`,`codalmac`,`codartic`,`linea`,`numlote`,`cantlote`)
    SQL = "Select linea,numlote,cantlote,if(fincuba=1,""SI"",""no"") from olicoupagelinlotes" & SQL
    
    SQL = SQL & " ORDER BY linea"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    
    'Nombre producto
    
    DataGrid1.Columns(0).visible = False
    
    i = 1
        DataGrid1.Columns(i).Caption = "Lote"
        DataGrid1.Columns(i).Width = 2800
        
    i = 2
        DataGrid1.Columns(i).Caption = "Cantidad"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).NumberFormat = FormatoCantidad
        DataGrid1.Columns(i).Alignment = dbgRight
    
    i = 3
        DataGrid1.Columns(i).Caption = "Fin"
        DataGrid1.Columns(i).Width = 800
    
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        txtAux(1).Height = DataGrid1.RowHeight
        txtAux(1).Left = DataGrid1.Columns(1).Left + 120
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        
        'cantidad
        txtAux(2).Height = DataGrid1.RowHeight
        txtAux(2).Left = DataGrid1.Columns(2).Left + 120
        txtAux(2).Width = DataGrid1.Columns(2).Width
        
        Combo1.Left = DataGrid1.Columns(3).Left + 120
        Combo1.Width = DataGrid1.Columns(3).Width
        CadAncho = True
    End If
   
   'No permitir cambiar tama�o de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        mnEliminar.Enabled = Not adodc1.Recordset.EOF
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
    If Index = 0 Then
        PonerFormatoEntero txtAux(Index) 'codmarca
    ElseIf Index = 2 Then
        If txtAux(Index).Text <> "" Then
            If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim C As Currency
Dim Au2 As String
    DatosOk = False
    If txtAux(1).Text = "" Or txtAux(2).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
    Else
    
        C = ImporteFormateado(txtAux(2).Text)  'Cantidad
        If Modo = 4 Then
            'Modificando
            C = Suma - adodc1.Recordset!cantlote + C
            If C > vCantidad Then
                MsgBox "Importe excede del total." & vbCrLf & "Albaran: " & Format(vCantidad, FormatoCantidad) & _
                     vbCrLf & "Lotes: " & Format(C, FormatoCantidad) & vbCrLf & "Dif: " & Format(C - vCantidad), vbExclamation
                Exit Function
            End If
            
            
            
        Else
            C = C + Suma
            If C > vCantidad Then
                MsgBox "Excede del total", vbExclamation
                Exit Function
            End If
            
        End If
        
        
        
        
        'Comprobamos que exista el lote
        Au2 = "codalmac = " & Me.vCodAlmac & " AND numlote = " & DBSet(Me.txtAux(1).Text, "T")
        Au2 = Au2 & " AND codartic "
        Au2 = DevuelveDesdeBD(conAri, "id", "spartidas", Au2, Me.vCodartic, "T")
        If Au2 = "" Then
            MsgBox "No existe el lote: " & txtAux(1).Text & " para el articulo " & Me.vCodartic, vbExclamation
            Exit Function
        End If
        
        
        DatosOk = True
    End If
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

Private Function InsertarModificar_() As Boolean

        '
        InsertarModificar_ = InsertarModificar(txtAux(1).Text, txtAux(2).Text, txtAux(0).Text)
        
    
        
    
End Function


'Si idPartida <0 entonces estoy insertando a mano
Private Function InsertarModificar(Lote As String, Cantidad As String, linea As Integer) As Boolean
Dim SQL As String
Dim Leido As Boolean
Dim Can As Currency

    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    
    
    'insert into `olicoupagelinlotes` (`codigo`,`codartic`,`linea`,`numlote`,`cantlote`)
    If Modo = 3 Then
        SQL = "insert into olicoupagelinlotes (`codigo`,`codartic`,`linea`,`numlote`,`cantlote`,`fincuba` ) values ("
        SQL = SQL & Me.vIdCup & ",'" & Me.vCodartic & "',"
        SQL = SQL & linea & ",'" & DevNombreSQL(Lote) & "'," & DBSet(Cantidad, "N") & ","
        If Combo1.ListIndex = 1 Then
            SQL = SQL & "1"
        Else
            SQL = SQL & "0"
        End If
        SQL = SQL & ")"
    Else
        'Modificar

        SQL = "UPDATE olicoupagelinlotes SET numlote = '" & DevNombreSQL(Lote) & "' "
        SQL = SQL & ", `fincuba` = " & Combo1.ListIndex
        SQL = SQL & ", cantlote = " & DBSet(Cantidad, "N") & " " & DevWHERE
        SQL = SQL & " AND linea = " & linea
    End If
    Conn.Execute SQL
    InsertarModificar = True
    
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, Err.Description
End Function





Private Function buscarNumerosLotes() As Boolean
Dim Rc As Byte
Dim cL As Collection
Dim Par As cPartidas
Dim cad As String
Dim i As Integer

    buscarNumerosLotes = False
    Set Par = New cPartidas
    Rc = Par.RecuperarLotes(Me.vCodartic, vCodAlmac, vCantidad, cL)
    Set Par = Nothing
    If Rc = 2 Then
        'Error. NO hay ningun numero de lote para el articulo/almacen
        
    Else
        'Mensajito
        cad = ""
        For i = 1 To cL.Count
            Suma = RecuperaValor(cL(i), 2)
            cad = cad & RecuperaValor(cL(i), 3) & Space(10) & Format(Suma, FormatoCantidad) & vbCrLf
        Next i
        Suma = 0
        cad = "Asignar los siguientes numeros de lote: " & vbCrLf & vbCrLf & cad & vbCrLf
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Function
            
        
        'Si que vamos a asignar los numeros de lote
        Modo = 3
        For i = 1 To cL.Count
            cad = RecuperaValor(cL(i), 1)
            NumRegElim = CLng(cad)
            Set Par = New cPartidas
            If Not Par.Leer(NumRegElim) Then
                MsgBox "Error insesperado leyendo partidas", vbExclamation
            Else
                Suma = RecuperaValor(cL(i), 2)
                InsertarModificar Par.Numlote, CStr(Suma), i
            End If
        Next i
        Suma = 0
        Modo = 2
        If Rc = 1 Then
            'Significa que aun quedan lotes por asignar
            CargaGrid
            Espera 0.1
        Else
            'Todo oK
            buscarNumerosLotes = True
        End If
    End If


End Function

