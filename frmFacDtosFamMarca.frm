VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacDtosFamMarca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dtos. Familia/Marca Cliente"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10950
   ClipControls    =   0   'False
   Icon            =   "frmFacDtosFamMarca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. Cliente|N|N|0|999999|sdtofm|codclien|000000|S|"
      Text            =   "cod. clien"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha Aplicación|F|N|||sdtofm|fechadto|dd/mm/yyyy|N|"
      Text            =   "fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   27
      ToolTipText     =   "Buscar cliente"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   26
      ToolTipText     =   "Buscar familia artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "Buscar marca"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   4
      Left            =   7320
      TabIndex        =   24
      ToolTipText     =   "Buscar tarifa"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   3
      Left            =   3960
      TabIndex        =   23
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "Cod. Tarifa|N|S|0|999|sdtofm|codlista|000|N|"
      Text            =   "codt"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Descuento 1|N|N|0|99.90|sdtofm|dtoline1|#0.00|N|"
      Text            =   "Dto 1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Descuento 2|N|N|0|99.90|sdtofm|dtoline2|#0.00|N|"
      Text            =   "Dto 2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "Dto. 1 Cajas|N|S|0|99.90|sdtofm|dtocaja1|#0.00|N|"
      Text            =   "dto1 Cajas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   6000
      MaxLength       =   5
      TabIndex        =   7
      Tag             =   "Dto. 2 Cajas|N|S|0|99.90|sdtofm|dtocaja2|#0.00|N|"
      Text            =   "dto2 Cajas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Cod. Marca|N|N|0|9999|sdtofm|codmarca|0000|S|"
      Text            =   "codmarca"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   6975
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   5000
      Width           =   2565
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9675
      TabIndex        =   10
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9675
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   5000
      Width           =   3645
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   4207
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   5000
      Width           =   2445
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Cod. Familia|N|N|0|9999|sdtofm|codfamia|0000|S|"
      Text            =   "cod.famia"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
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
         Left            =   6600
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
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
      Bindings        =   "frmFacDtosFamMarca.frx":000C
      Height          =   4110
      Left            =   240
      TabIndex        =   12
      Top             =   580
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7250
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
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
   Begin VB.Label Label11 
      Caption         =   "Desc. Marca"
      Height          =   255
      Left            =   6975
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Desc. Familia"
      Height          =   255
      Left            =   4207
      TabIndex        =   16
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Desc. Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Width           =   975
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacDtosFamMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFam As frmAlmFamiliaArticulo  'Form Mantenimiento Familias Articulos
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas  'Form Mantenimiento Marcas
Attribute frmM.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas  'Form Mantenimiento Tarifas
Attribute frmT.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim WhereConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

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
                CargaGrid True
                BotonAnyadir
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
'                    CargaTxtAux False, False
                    LLamaLineas 10
                    SituarDataPosicion Data1, NumReg, Indicador
                End If
                lblIndicador.Caption = Indicador
                DataGrid1.SetFocus
            End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Cod Cliente
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
        Case 1 'Cod Familia
            Set frmFam = New frmAlmFamiliaArticulo
            frmFam.DatosADevolverBusqueda = "0"
            frmFam.Show vbModal
            Set frmFam = Nothing
        Case 2 'Cod Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3 'Fecha aplicacion
            Screen.MousePointer = vbHourglass
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index).Text <> "" Then frmF.Fecha = CDate(txtAux(Index).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
        Case 4 'Tarifas
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    
    If Index = 4 Then
        PonerFoco txtAux(8)
    Else
        PonerFoco txtAux(Index)
    End If
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else 'No hay Registros en la Tabla
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            DeseleccionaGrid Me.DataGrid1
            PonerModo 2
            LLamaLineas 10
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.AbsolutePosition > 0 Then
            'Poner descripcion del Cliente
            Text2(0).Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", CStr(Data1.Recordset!CodClien), "N")
            
            'Poner descripcion de Familia
            Text2(1).Text = DevuelveDesdeBDNew(conAri, "sfamia", "nomfamia", "codfamia", CStr(Data1.Recordset!codfamia), "N")
            
            'Poner descripcion de Familia
            Text2(2).Text = DevuelveDesdeBDNew(conAri, "smarca", "nommarca", "codmarca", CStr(Data1.Recordset!codmarca), "N")
            
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        End If
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
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
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    NombreTabla = "sdtofm" 'Tabla Descuentos Familia/Marca
    Ordenacion = " ORDER BY codclien, codfamia, codmarca "
    WhereConsulta = " WHERE codclien = -1"
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
On Error GoTo ECarga

    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    CargaGrid2
    DataGrid1.Enabled = (Modo = 2)
    Me.DataGrid1.ScrollBars = dbgAutomatic
    PrimeraVez = False
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    'SQL = "SELECT codclien, codfamia, codmarca, fechadto, dtoline1, dtoline2, dtocaja1, dtocaja2, " & tabla & ".codlista, " & " Tarifas.nomlista "
    tots = "S|txtAux(0)|T|Cliente|850|;S|cmdAux(0)|B||0|;S|txtAux(1)|T|Familia|750|;S|cmdAux(1)|B||0|;S|txtAux(2)|T|Marca|700|;S|cmdAux(2)|B||0|;"
    tots = tots & "S|txtAux(3)|T|Fecha dto.|1150|;S|cmdAux(3)|B||0|;S|txtAux(4)|T|Dto 1|800|;S|txtAux(5)|T|Dto 2|800|;S|txtAux(6)|T|Dto Caja1|1000|;S|txtAux(7)|T|Dto Caja2|1000|;"
    tots = tots & "S|txtAux(8)|T|Tarif. alt|900|;S|cmdAux(4)|B||0|;S|txtAux2|T|Desc. Tarifa|1900|;"
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(8).Alignment = dbgCenter

ECarga2:
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
        txtAux2.Height = Me.DataGrid1.RowHeight
        txtAux2.Top = alto
        txtAux2.visible = b
        
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = b
        Next jj
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
            Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtAux(3).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Clientes
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Familias
    txtAux(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento MARCAS
    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento TARIFAS
    txtAux(8).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtAux2.Text = RecuperaValor(CadenaSeleccion, 2)
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
            AbrirListado (54) '54: Listado Descuentos Familia/Marca
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
    
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(0)
        Case 2    'Preparamos para que pueda Modificar
            Me.cmdRegresar.visible = False
    End Select
                            
    BloquearClavesP (Modo = 4) ' si modificar
           
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    Me.DataGrid1.Enabled = (Modo = 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    'Modo 2. Hay datos y estamos visualizandolos
    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = (b Or (Modo = 0))
    Me.mnNuevo.Enabled = (b Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
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
    
    Tabla = "sdtofm"
    SQL = "SELECT codclien, codfamia, codmarca, fechadto, dtoline1, dtoline2, dtocaja1, dtocaja2, " & Tabla & ".codlista, " & " Tarifas.nomlista "
    SQL = SQL & " FROM " & Tabla & " LEFT JOIN starif AS Tarifas ON " & Tabla & ".codlista ="
    SQL = SQL & " Tarifas.codlista"

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        ElseIf CadenaConsulta = "" Then
            If CadenaBusqueda <> "" Then
                CadenaBusqueda = CadenaBusqueda & " OR " & "(codclien=" & txtAux(0).Text & " and codfamia=" & txtAux(1).Text & " and codmarca=" & txtAux(2).Text & ")"
            Else
                CadenaBusqueda = " WHERE (codclien=" & txtAux(0).Text & " and codfamia=" & txtAux(1).Text & " and codmarca=" & txtAux(2).Text & ")"
            End If
            SQL = SQL & CadenaBusqueda
        End If
    Else
        SQL = SQL & " WHERE codclien = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1)
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
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
'        CadenaConsulta = ""
'    End If
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    anc = ObtenerAlto(Me.DataGrid1)
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
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    'poner valores grabados
    For i = 0 To 2
        txtAux(i).Text = DBLet(DataGrid1.Columns(i).Value, "N")
        FormateaCampo txtAux(i)
    Next i
    txtAux(3).Text = DBLet(Me.DataGrid1.Columns(3).Value, "F")
    For i = 4 To 8
        txtAux(i).Text = DBLet(DataGrid1.Columns(i), "N")
        FormateaCampo txtAux(i)
    Next i
    txtAux2.Text = Me.DataGrid1.Columns(i)

    PonerFoco txtAux(3)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        SQL = "¿Seguro que desea eliminar el Descuento para?" & vbCrLf
        SQL = SQL & vbCrLf & "Cliente: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Text2(0).Text
        SQL = SQL & vbCrLf & "Familia: " & Format(Data1.Recordset.Fields(1).Value, "0000") & " - " & Text2(1).Text
        SQL = SQL & vbCrLf & "Marca : " & Format(Data1.Recordset.Fields(2).Value, "0000") & " - " & Text2(2).Text
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            SQL = "Delete from sdtofm where codclien=" & Val(Data1.Recordset!CodClien)
            SQL = SQL & " and codfamia=" & Val(Data1.Recordset!codfamia) & " and codmarca=" & Val(Data1.Recordset!codmarca)
            Conn.Execute SQL
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataTrasEliminar Data1, NumRegElim
            CargaGrid2
        End If
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
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
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(txtAux(0), 40, "Cod. Clien.")
    cad = cad & ParaGrid(txtAux(1), 20, "Cod. Artic")
    Tabla = NombreTabla
    Titulo = "Precios Especiales"

    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
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
                PonerFoco txtAux(kCampo)
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
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        PonerModo Modo
        CargaGrid False
         MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        If EsBusqueda Then CadenaConsulta = ""
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
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub BloquearClavesP(bol As Boolean)
'Si BloquearClavesPrimarias=true deshablilita los textbox de codigos y lo pone amarillo
'y habilita el resto de campos para introducir nuevos valores
'Si BloquearClavesPrimarias=false habilita los textbox de codigos para introducir
Dim i As Byte

    For i = 0 To 2 'Codigos
        BloquearTxt txtAux(i), bol
        Me.cmdAux(i).Enabled = Not bol
    Next i
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
'    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0 'cod. Cliente
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "sclien", "nomclien")
            Else
                Text2(Index).Text = ""
            End If
        Case 1 'Cod. Familia
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "sfamia", "nomfamia")
            Else
                Text2(Index).Text = ""
            End If
        Case 2 'Cod. Marca
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "smarca", "nommarca")
            Else
                Text2(Index).Text = ""
            End If
        Case 3 'Fecha Descuento
            PonerFormatoFecha txtAux(Index)
        Case 4 To 7 'Descuentos
            'Formato tipo 4: Decimal(4,2)
            PonerFormatoDecimal txtAux(Index), 4
        Case 8 'Cod. Tarifa
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2.Text = PonerNombreDeCod(txtAux(Index), conAri, "starif", "nomlista")
            Else
                txtAux2.Text = ""
            End If
    End Select
End Sub
