VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lotes"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmFacLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBus 
      Caption         =   "+"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   3
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
      Bindings        =   "frmFacLotes.frx":000C
      Height          =   3270
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4335
      _ExtentX        =   7646
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
      Left            =   3360
      TabIndex        =   5
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2220
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Volver"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   7
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
         TabIndex        =   8
         Top             =   240
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
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
      TabIndex        =   11
      Top             =   600
      Width           =   4455
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
Attribute VB_Name = "frmFacLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCodAlmac As Integer
Public vCodtipom As String
Public vNumalbar As Long
Public vNumlinea As Integer
Public vCantidad As Currency   'Cantidad de la linea de albaran
Public vCodartic As String     'articulo para buscar en los lotes
Public vFecha As Date          'para la clase Lotaje. Sera fecha mov





Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Private WithEvents frmL As frmAlmPartidas
Attribute frmL.VB_VarHelpID = -1

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

Dim DatosDevueltos As String


Private Sub PonerModo(vModo As Byte)
Dim B As Boolean
    
    Modo = vModo
    B = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = False
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    'Me.cmdBus.visible = Not B And vEmpresa.codempre <> EmpresaAVAB
    Me.cmdBus.visible = Not B And vParamAplic.Produccion
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B

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
Dim B As Boolean

    B = False
    'Buscar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = B
    Me.mnVerTodos.Enabled = B
   
    B = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    Toolbar1.Buttons(10).Enabled = B
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
    CadenaConsulta = "select sum(cantidad) total ,max(linea) ultimo from slialblotes " & DevWHERE
    miRsAux.Open CadenaConsulta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        'NInugni, no hay ninguno
        NumRegElim = 0
        Suma = 0
        
    Else
        NumRegElim = DBLet(miRsAux!ultimo, "N")
        Suma = DBLet(miRsAux!Total, "N")
        
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
    LLamaLineas anc, 3
    BloquearTxt txtAux(1), False
    'Ponemos el foco
    PonerFoco txtAux(1)
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
Dim SQL As String
Dim anc As Single
Dim i As Integer
On Error GoTo EModificar

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    
   If vParamAplic.Produccion Then
        SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", adodc1.Recordset!NUmlote, "T")
        If SQL <> "" Then
            MsgBox "No puede modificar lote en deposito. Elimine y vuelva a insertar", vbExclamation
            Exit Sub
        End If
    End If
    
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
    LLamaLineas anc, 4
    BloquearTxt txtAux(1), True
    Me.cmdBus.visible = False
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
    cad = DevuelveDesdeBD(conAri, "sum(cantidad)", "slialblotes", cad, "1")
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
    Me.cmdBus.Top = alto - 30
  '  txtAux(0).Left = DataGrid1.Left + 340
  '  txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
Dim SQL As String
Dim cP As cPartidas
Dim cLo As cLotaje
Dim cDep As cDeposito
On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    

    '### a mano
    SQL = "�Seguro que desea eliminar el numero de lote "
    
    SQL = SQL & adodc1.Recordset.Fields(1) & "?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Reestablecemos en partidas
        Set cP = New cPartidas
        SQL = CStr(adodc1.Recordset!NUmlote)
        If cP.LeerDesdeArticulo(vCodartic, vCodAlmac, SQL) Then
            SQL = CStr(adodc1.Recordset!Cantidad)
            If SQL = "" Then SQL = 0
            cP.IncrementarCantidad CCur(SQL)
        End If
        
        
        SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", adodc1.Recordset!NUmlote, "T")
        If SQL <> "" Then
            Set cDep = New cDeposito
            cDep.LeerDatos CInt(SQL), True
            If cDep.NUmlote <> "" Then
                If cDep.IdPartida = cP.IdPartida Then
                    cDep.VariacionKilosDeposito adodc1.Recordset!Cantidad
                    MsgBox "Falta eliminar en hcodepositos. Consulte soporte tecnico", vbExclamation
                End If
            End If
            Set cDep = Nothing
        
        End If
        
        
        
        
        Set cP = Nothing
        
        'LOTAJE. Movimientos smovalotes
            Set cLo = New cLotaje
            AsignarLotaje cLo  'asignar valores fijos
            cLo.NUmlote = CStr(adodc1.Recordset!NUmlote)
            cLo.SubLinea = Val(adodc1.Recordset!Linea) 'La sublinea del lote 'Normalmente 1 o 2
            If cLo.Leer Then cLo.EliminarMovimArticulosLotaje False
            Set cLo = Nothing
            
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = DevWHERE
        SQL = "Delete FROM slialblotes " & SQL & " AND linea = " & adodc1.Recordset!Linea
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
                If InsertarModificar_(Nothing, 0, 0) Then
                    CargaGrid
                    BotonAnyadir False
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                
                If InsertarModificar_(Nothing, 0, 0) Then
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


Private Sub cmdBus_Click()
Dim C As Currency

    DatosDevueltos = ""
    Set frmL = New frmAlmPartidas
    frmL.DatosADevolverBusqueda = vCodartic
    frmL.Show vbModal
    Set frmL = Nothing
    If DatosDevueltos <> "" Then
        'Comprobamos cantidad
        
        
        C = CCur(RecuperaValor(DatosDevueltos, 2))
        If C < 0 Then
            MsgBox "Cantidad negativa.", vbExclamation
        Else
            txtAux(1).Text = RecuperaValor(DatosDevueltos, 1)
            If C > vCantidad - Suma Then
                'Tengo mas. Solo cojo lo que necesito
                txtAux(2).Text = vCantidad - Suma
            End If
            
        End If
    End If
    
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
        
        If MsgBox("Total lineas distinto del total de la linea del albaran" & vbCrLf & "�Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        If Suma > vCantidad Then Exit Sub
    End If
            
    Unload Me
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
     
        Me.lblDesc.Caption = vCodtipom & Format(vNumalbar, "000000") & " - " & vNumlinea
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
    

    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields

    CadAncho = False
    cmdRegresar.visible = False
    PonerModo 2
    Me.lblDesc.Caption = ""
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmL_DatoSeleccionado(CadenaSeleccion As String)
    DatosDevueltos = CadenaSeleccion
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
    DevWHERE = " where codtipom='" & Me.vCodtipom & "' AND numalbar =" & vNumalbar & " AND numlinea =" & vNumlinea
End Function

Private Sub CargaGrid()
Dim i As Byte
Dim B As Boolean
Dim SQL As String
    B = DataGrid1.Enabled

    SQL = DevWHERE
    SQL = "Select linea,numlote,cantidad from slialblotes" & SQL
    
    SQL = SQL & " ORDER BY linea"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False
    
    
    'Nombre producto
    DataGrid1.RowHeight = txtAux(1).Height
    
    DataGrid1.Columns(0).visible = False
    
    i = 1
        DataGrid1.Columns(i).Caption = "Lote"
        DataGrid1.Columns(i).Width = 2800
        
    i = 2
        DataGrid1.Columns(i).Caption = "Cantidad"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).NumberFormat = FormatoCantidad
        DataGrid1.Columns(i).Alignment = dbgRight
    

            
    'Fiajamos el cadancho
    If Not CadAncho Then
        txtAux(1).Left = DataGrid1.Columns(1).Left + 120
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        
        'cantidad
        txtAux(2).Left = DataGrid1.Columns(2).Left + 120
        txtAux(2).Width = DataGrid1.Columns(2).Width
        
        Me.cmdBus.Height = txtAux(1).Height + 60
        Me.cmdBus.Left = txtAux(2).Left + 60 - Me.cmdBus.Width
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
    If Index = 1 Then
        'PonerFormatoEntero txtAux(Index) 'codmarca
        If txtAux(1).Text <> "" Then PonerFoco txtAux(2)
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
        If C <= 0 Then
            MsgBox "Importe debe ser mayor que cero", vbExclamation
            Exit Function
        End If
        If Modo = 4 Then
            'Modificando
            C = Suma - adodc1.Recordset!Cantidad + C
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
        
        If Not vParamAplic.Produccion Then
            DatosOk = True
            Exit Function
        End If
        
        'Comprobamos que exista el lote
        DatosDevueltos = "cantotal"
        Au2 = "codalmac = " & Me.vCodAlmac & " AND numlote = " & DBSet(Me.txtAux(1).Text, "T")
        Au2 = Au2 & " AND codartic "
        Au2 = DevuelveDesdeBD(conAri, "id", "spartidas", Au2, Me.vCodartic, "T", DatosDevueltos)
        If Au2 = "" Then
            Au2 = "No existe el lote: " & txtAux(1).Text & " para el articulo " & Me.vCodartic & vbCrLf
            Au2 = Au2 & "�Continuar?"
            If MsgBox(Au2, vbQuestion + vbYesNo) = vbNo Then Exit Function
            
        Else
            C = ImporteFormateado(txtAux(2).Text)
            If CCur(DatosDevueltos) < C Then
                'No tengo tantos
                Au2 = "Cantidad insuficente     Lote: " & txtAux(1).Text & "       Articulo " & Me.vCodartic & vbCrLf
                Au2 = Au2 & "Necesaria: " & Format(C, FormatoCantidad) & vbCrLf
                C = CCur(DatosDevueltos)
                Au2 = Au2 & "Existente: " & Format(C, FormatoCantidad) & vbCrLf
                Au2 = Au2 & "�Continuar?"
                If MsgBox(Au2, vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
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

Private Function InsertarModificar_(cPar As cPartidas, Lin As Integer, CantidadSiEsDesdepartidas As Currency) As Boolean
Dim B As Boolean

    InsertarModificar_ = False
    Conn.BeginTrans
    
    If cPar Is Nothing Then
        'Desde boton a�adir
        B = InsertarModificar(Nothing, txtAux(1).Text, txtAux(2).Text, txtAux(0).Text)
    Else
        'Desde la obtencion de los numeros de lotes en el load
        B = InsertarModificar(cPar, cPar.NUmlote, CStr(CantidadSiEsDesdepartidas), Lin)
    End If
    If B Then
        Conn.CommitTrans
        InsertarModificar_ = True
    Else
        Conn.RollbackTrans
    End If
End Function



Private Sub AsignarLotaje(ByRef cL As cLotaje)


    cL.codArtic = vCodartic
    cL.codAlmac = vCodAlmac
    cL.DetaMov = vCodtipom
    cL.LineaDocu = Me.vNumlinea
    cL.Documento = Me.vNumalbar
    cL.tipoMov = 0
    
End Sub

'Si idPartida <0 entonces estoy insertando a mano
Private Function InsertarModificar(cP As cPartidas, Lote As String, Cantidad As String, Linea As Integer) As Boolean
Dim SQL As String
Dim Leido As Boolean
Dim Can As Currency
Dim cLot As cLotaje
Dim InsertarLotaje As Boolean
Dim cDep As cDeposito


    On Error GoTo EInsertarModificar
    InsertarModificar = False
    '---------------------
    Set cLot = New cLotaje
    AsignarLotaje cLot  'Ponemos los campos
    cLot.NUmlote = Lote
    cLot.Cantidad = ImporteFormateado(Cantidad)
    cLot.SubLinea = Linea 'La sublinea del lote 'Normalmente 1 o 2
    InsertarLotaje = True
    If Modo = 3 Then
        SQL = "insert into `slialblotes` (`codtipom`,`numalbar`,`numlinea`,`linea`,`numlote`,cantidad) values ('"
        SQL = SQL & Me.vCodtipom & "'," & vNumalbar & "," & vNumlinea & ","
        'SQL = SQL & txtAux(0).Text & ",'" & DevNombreSQL(txtAux(1).Text) & "'," & DBSet(txtAux(2).Text, "N") & ")"
        'Ahora
        SQL = SQL & Linea & ",'" & DevNombreSQL(Lote) & "'," & DBSet(Cantidad, "N") & ")"
        
    Else
        If cLot.Leer Then InsertarLotaje = False
    
        'Modificar
        
        SQL = "UPDATE slialblotes SET numlote = '" & DevNombreSQL(Lote) & "' "
        SQL = SQL & ", cantidad = " & DBSet(Cantidad, "N") & " " & DevWHERE
        SQL = SQL & " AND linea = " & Linea
    End If
    Conn.Execute SQL
    
    If InsertarLotaje Then
        'Hay k rellenar el resto de valores
        cLot.Fechamov = vFecha
        cLot.HoraMov = Now
        cLot.InsertarLote
    Else
        cLot.HoraMov = Now   'ha si guardo la modificacion
        cLot.ModificarMovimArticulosLotaje True
    End If
    
    InsertarModificar = True  'Ya ponemos que esta bien
                               'aunque de errores bajo
    
   

    
    Set cLot = Nothing

    'Si no hay produccion No metememos en partidas
    '--------------------------------------------
    If Not vParamAplic.Produccion Then Exit Function
    
    
    If cP Is Nothing Then
        
        If Modo = 4 Then
            'Modificar.  Habria que ver si ha cambiado el numero de LOTE
    
    
    
        End If
        Set cP = New cPartidas
        If cP.LeerDesdeArticulo(vCodartic, vCodAlmac, Lote) Then
            'SI existe el lote
            Can = Cantidad
            If Modo = 4 Then Can = Cantidad - adodc1.Recordset!Cantidad
            
            cP.IncrementarCantidad -Can
        Else
            'NO existe el lote. Lo creamos en negativo?
            
            
        End If
    
    Else
        'Ya tenemos el lote
        Can = -1 * ImporteFormateado(Cantidad)
        cP.IncrementarCantidad Can
        
    End If
    
    
    
    SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", cP.NUmlote, "T")
    If SQL <> "" Then
        'Venta directa de un deposito
        Set cDep = New cDeposito
        cDep.LeerDatos CInt(SQL), True
        If cDep.NUmlote <> "" Then
            
            If cDep.IdPartida = cP.IdPartida Then
                cDep.VariacionKilosDeposito -Cantidad
                cDep.InsertarEnHco 6, vCodtipom & vNumalbar
            Else
                SQL = "Venta con numero de lote en deposito, pero distinta Partida:"
                SQL = SQL & "Dep " & cDep.IdPartida & "    partida: " & cP.IdPartida
                SQL = SQL & " El proceso continuar�. Llame a soporte t�cnico"
                MsgBox SQL, vbExclamation
            End If
        End If
        Set cDep = Nothing
    End If
    
    Set cLot = Nothing
    
    
    
    
    
    Set cP = Nothing
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
                InsertarModificar_ Par, i, Suma
            End If
        Next i
        Suma = 0
        Modo = 2
        If Rc = 1 Then
            'Significa que aun quedan lotes por asignar
            
            
        Else
            'Todo oK
            buscarNumerosLotes = True
        End If
    End If


End Function

