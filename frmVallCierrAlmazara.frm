VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallCierrAlmazara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre y asignacion rendimientos de ALMAZARA"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13380
   ClipControls    =   0   'False
   Icon            =   "frmVallCierrAlmazara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   720
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   5160
      Width           =   1245
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "existencia"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5595
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12075
      TabIndex        =   2
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVallCierrAlmazara.frx":000C
      Height          =   3795
      Left            =   120
      TabIndex        =   5
      Top             =   1305
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6694
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   5520
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Litros"
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
      Index           =   1
      Left            =   10680
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
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
      Index           =   0
      Left            =   8400
      TabIndex        =   14
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Deposito"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   5760
      Width           =   2055
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
      TabIndex        =   7
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmVallCierrAlmazara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long

Private Modo As Byte

Dim kCampo As Integer

Dim CadenaConsulta As String

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim Cad As String

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    
    lblInfInv.Caption = ""
    Select Case Modo
        Case 2
            'Cerrar proceso. Asignar rendimientos. Cerrar parte
            Cad = DevuelveDesdeBD(conAri, "numalbar", "tmpnlotes", "cantidad=0 and codusu", vUsu.Codigo)
            If Cad <> "" Then
                MsgBox "Lineas sin asignar rendimiento", vbExclamation
                
            Else
                
                If Combo1.ListIndex < 0 Then
                    MsgBox "Seleccione el deposito destino", vbExclamation
                    Cad = "No"
                Else
                    Cad = ""
                End If
            End If
            
            If Cad = "" Then
                
                'Preguntaremos si cerramos el parte
                Cad = "Almazara." & vbCrLf & "Nº Albaranes: " & Data1.Recordset.RecordCount & vbCrLf & "Kg olivas:  " & Text2(0).Text
                Cad = Cad & vbCrLf & "Litros producidos: " & Text2(1).Text & vbCrLf & "Destino: " & Combo1.Text
                Cad = Cad & "       " & "Lote :"
                'Veremos si el deposito esta vacio. Crearemos nuevo lote
                
                'Veremos la capaciadad que hay mas la estimada
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open "Select * from proddepositos where numdeposito=" & Combo1.ListIndex + 1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If IsNull(miRsAux!NUmlote) Then
                    Cad = Cad & "*** Nuevo ****"
                    CadenaConsulta = "0"
                
                Else
                    Cad = Cad & miRsAux!NUmlote & " (Se creará nuevo)"
                    CadenaConsulta = miRsAux!Litros
                End If
                Cad = Cad & vbCrLf & "Litros deposito: " & Format(CadenaConsulta, FormatoCantidad) & "   Maximo: " & miRsAux!Capacidad
                CadenaConsulta = CCur(CadenaConsulta) + ImporteFormateado(Text2(1).Text)
                If CCur(CadenaConsulta) > miRsAux!Capacidad Then Cad = Cad & vbCrLf & "         --EXCEDE--"
                miRsAux.Close
                Set miRsAux = Nothing
                Cad = Cad & vbCrLf & vbCrLf & "¿Cerrar proceso almazara?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                    NumRegElim = 0
                    'hacemos el proceso de obtencion de aceite
                    If cerrarProcesoMolturacion Then
                        Me.Refresh
                        Espera 0.5
                        conn.Execute "commit"
                        Espera 0.5
                        'Si hemos creado un coupage, lo cerramos
                        If NumRegElim > 0 Then
                            frmProduVarios.Intercambio = NumRegElim & "|1|"
                            frmProduVarios.opcion = 5
                            frmProduVarios.Show vbModal
                            Espera 0.5
                        End If
                        
                        CadenaDesdeOtroForm = "OK"
                        Unload Me
                    End If
                End If
                
                
             End If
        Case 4 'Modificar Existencia Real (Introducir Valores Reales)
            CargaTxtAux False, False
            PonerModo 2
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Set miRsAux = Nothing
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar
    
     
    lblInfInv.Caption = ""
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid
        Case 2
            'Salir sin cerrar parte
            If MsgBox("Desea salir sin cerrar el proceso de almazara?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Unload Me
        Case 4  ' 4: Modificar
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub Combo1_Click()
    If Combo1.ListIndex < 0 Then
        Me.txtLote.Text = ""
    Else
        CadenaConsulta = DevuelveDesdeBD(conAri, "Numlote", "proddepositos", "numDeposito", Combo1.ListIndex + 1)
        If CadenaConsulta = "" Then
            'NUEVO. No existe lote
            txtLote.Text = "VACIO"
            txtLote.FontBold = True
            txtLote.ForeColor = vbRed
        Else
            txtLote.Text = CadenaConsulta
            txtLote.FontBold = False
            txtLote.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Me.Tag = "NO" Then
        Me.Tag = ""
        BotonModificar
        PonerFoco txtAux
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    Me.Tag = "NO"
    
    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(4).Image = 4 'Modificar
        .Buttons(5).Image = 15 'Salir
    End With

   
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True


    CadenaDesdeOtroForm = ""
    For NumRegElim = 0 To 16
        Combo1.AddItem "Deposito " & NumRegElim + 1
    Next

    PonerModo 2
    CargaGrid
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()
Dim I As Byte
Dim SQL As String
On Error GoTo ECarga

    gridCargado = False

    SQL = MontaSQLCarga()
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    PrimeraVez = False
        
    'numalbar,fechaalb,codprove,numlinea,codartic,nomartic,numlotes,cantidad
    DataGrid1.Columns(0).Caption = "NºAlb"
    DataGrid1.Columns(0).Width = 850
    DataGrid1.Columns(1).Caption = "Fecha"
    DataGrid1.Columns(1).Width = 1150
  
    DataGrid1.Columns(2).Caption = "Cod"
    DataGrid1.Columns(2).Width = 800
    DataGrid1.Columns(2).NumberFormat = "0000"
        
    DataGrid1.Columns(3).Caption = "Proveedor"
    DataGrid1.Columns(3).Width = 2600
    
    'Cod artic
    DataGrid1.Columns(4).Caption = "Art."
    DataGrid1.Columns(4).Width = 1100
    
    'Cod artic
    DataGrid1.Columns(5).Caption = "Oliva"
    DataGrid1.Columns(5).Width = 2400
    
    DataGrid1.Columns(6).Caption = "Kilos"
    DataGrid1.Columns(6).Width = 1100
    DataGrid1.Columns(6).Alignment = dbgRight
    
    DataGrid1.Columns(7).Caption = "Rdto"
    DataGrid1.Columns(7).Width = 800
    DataGrid1.Columns(7).Alignment = dbgRight
    
    DataGrid1.Columns(8).Caption = "Litros"
    DataGrid1.Columns(8).Width = 1100
    DataGrid1.Columns(8).Alignment = dbgRight
    
    
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
        DataGrid1.Columns(I).Locked = True
    Next I
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    
    SQL = "SELECT sum(numlotes+0),sum( Round(((numlotes + 0) * Cantidad) / 100, 2)) from tmpnlotes where codusu = " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic
    Me.Text2(0).Text = "": Me.Text2(1).Text = ""
    If Not miRsAux.EOF Then
        Me.Text2(0).Text = Format(miRsAux.Fields(0), FormatoCantidad)
        Me.Text2(1).Text = Format(miRsAux.Fields(1), FormatoCantidad)
    End If
    miRsAux.Close
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = Data1.Recordset!Cantidad
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(7).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(7).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
        txtAux.SelStart = 0
        txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub




Private Sub ImageCombo1_Change()

End Sub

Private Sub txtAux_GotFocus()
    ConseguirFocoLin txtAux
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
'                If DataGrid1.Row > 0 Then
'                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
''                elseif
'                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        ModificarExistencia
        PasarSigReg
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        If Not PonerFormatoDecimal(txtAux, 4) Then .Text = ""
    End With

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 4 'Modificar
            BotonModificar
        Case 5 'Salir
            Unload Me
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    'b = (Kmodo = 2)
   
    
    b = (Modo = 0)
    PonerBotonCabecera b
   
    If Modo = 2 Then
        Me.lblIndicador.Caption = "Cerrar"
    Else
        Me.lblIndicador.Caption = "Modificar"
    End If

           
    b = Modo <> 0 And Modo <> 2 And Modo <> 4

    b = (Modo = 1)
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(4).Enabled = Not b And (Not (Modo = 0 Or Modo = 4))

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Function MontaSQLCarga() As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

    SQL = "SELECT NumAlbar , FechaAlb, tmpnlotes.codProve, nomprove, codartic, NomArtic, numlotes, Cantidad,"
    SQL = SQL & " Round(((numlotes + 0) * Cantidad) / 100, 2) from tmpnlotes,sprove where tmpnlotes.codprove="
    SQL = SQL & " sprove.codprove and codusu = " & vUsu.Codigo
    
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()

    Modo = 4
    cmdCancelar_Click
    
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
    PonerFoco txtAux
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux.Text = Trim(txtAux.Text)

    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
        If PonerFormatoDecimal(txtAux, 4) Then
            DatosOk = True
        Else
            DatosOk = False
        End If
        'DatosOk = True
    Else
        DatosOk = False
    End If
End Function


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarExistencia(Cantidad As String) As Boolean
Dim C As String
    C = "UPDATE tmpnlotes SET cantidad=" & DBSet(Cantidad, "N")
    C = C & " WHERE codusu =" & vUsu.Codigo & " AND numalbar=" & DBSet(Data1.Recordset!NumAlbar, "T")
    C = C & " AND fechaalb =" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND codprove=" & Data1.Recordset!codProve
    ActualizarExistencia = EjecutaSQL(conAri, C, True)
    
End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOk Then
        If ActualizarExistencia(txtAux.Text) Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function


Private Function cerrarProcesoMolturacion() As Boolean
    conn.BeginTrans
    If cerrarProcesoAlm Then
        cerrarProcesoMolturacion = True
        conn.CommitTrans
    Else
        cerrarProcesoMolturacion = False
        conn.RollbackTrans
        NumRegElim = 0
    End If
End Function


Private Function cerrarProcesoAlm() As Boolean
Dim CP1 As cPartidas  'La ppal me la guardo por si hay coupage ir ams rapido
Dim cP2 As cPartidas
Dim cStock As cStock
Dim cD As cDeposito
Dim Kilos As Currency
Dim cLot As cLotaje
Dim NuevaPartida As Boolean
Dim Cad As String
Dim RT As ADODB.Recordset
Dim vT As CTiposMov
Dim articuloTalco As String
Dim NumDeposito As Integer
Dim KilosTotales As Currency


    On Error GoTo eCerrarProcesoAlm
    
    cerrarProcesoAlm = False
    
    'Generara una entrada en el deposito, o añadira la cantidad
    Set CP1 = New cPartidas
    Set cStock = New cStock
    Set cD = New cDeposito
    Set cLot = New cLotaje
    

    'ALTA en deposito
    'Si el deposito NO esta vacio, entonces produzo sobre el deposito 100 y luego coupo sobre el destino y el 100 en el destino
    NumDeposito = Combo1.ListIndex + 1
    If txtLote.Text <> "VACIO" Then NumDeposito = 100
 
    If Not cD.LeerDatos(NumDeposito, True) Then Err.Raise 513, , "Leyendo deposito: " & NumDeposito
    
    
    If cD.NumDeposito = 100 And cD.NUmlote <> "" Then Err.Raise 513, , "Deposito produccion NO esta vacio"
    
    
    If cD.NUmlote = "" Then
        'Asignaremos un lote nuevo.. y una nueva partida
        Set vT = New CTiposMov
        vT.Leer "LOV"
        CP1.NUmlote = vT.ConseguirContador(vT.TipoMovimiento)
        
        CP1.NUmlote = "MOSTRA" & CP1.NUmlote & "-"
        If Month(Now) < 11 Then
            CP1.NUmlote = CP1.NUmlote & Year(Now) - 1
        Else
            CP1.NUmlote = CP1.NUmlote & Year(Now)
        End If
        NuevaPartida = True
       
        cD.idPartida = CP1.Siguiente
        cD.Kilos = 0
        cD.NUmlote = CP1.NUmlote
        vT.IncrementarContador vT.TipoMovimiento
        
    Else
        NuevaPartida = False
        If Not CP1.Leer(cD.idPartida) Then Err.Raise 513, , "Leyendo partida: " & cD.idPartida
    End If
    
    Kilos = ImporteFormateado(Text2(1).Text) * 0.916
    
    'Incrementamos la cantidad del deposito
    cD.VariacionKilosDeposito Kilos   'Si tenia o no, da lo mismo. Esta sumando los nuevos
    If Not cD.InsertarEnDeposito(10) Then Err.Raise 513, , "Insertando datos nuevos deposito: " & cD.NUmlote
    
    'Metemos los moviimientos
    'Tanto en smoval como en smovallotes
    cStock.DetaMov = "MLT" 'Molturacion
    cStock.codAlmac = 1
    Cad = DevuelveDesdeBD(conAri, "articMolturacion", "vallparam", "1", "1")
    cStock.codartic = Cad
    cStock.Documento = Format(ID, "00000")
    cStock.Fechamov = Format(Now, "dd/mm/yyyy")
    cStock.HoraMov = Now
    cStock.Importe = 0
    cStock.LineaDocu = 1
    cStock.tipoMov = "E"
    cStock.Trabajador = vUsu.CodigoTrabajador
    cLot.codAlmac = cStock.codAlmac
    cLot.codartic = cStock.codartic
    cLot.DetaMov = cStock.DetaMov
    cLot.Documento = cStock.Documento
    cLot.Fechamov = cStock.Fechamov
    cLot.HoraMov = cStock.HoraMov
    cLot.LineaDocu = cStock.LineaDocu
    cLot.tipoMov = 1 'entrada
    cLot.NUmlote = CP1.NUmlote
    cLot.ProvCliTra = cStock.Trabajador
  
    cStock.Cantidad = Kilos
    cLot.Cantidad = Kilos
    
    
    If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock"
    If Not cLot.InsertarLote Then Err.Raise 513, , "Actualizando stock lotes"
    
    
    
    If NuevaPartida Then
        CP1.Cantidad = 0  'Luego los incremento
        CP1.codAlmac = cStock.codAlmac
        CP1.codartic = cStock.codartic
        CP1.codProve = 1
        CP1.Fecha = cStock.Fechamov
        CP1.NumAlbar = cStock.Documento
        'cp.NUmlote = lo he asginado arriba
        If Not CP1.Insertar Then Err.Raise 513, , "Creando partida "
    End If
    CP1.IncrementarCantidad Kilos
    
    
    'El orujo
    '------------------------------------------------------
    'Que es el resto de producto
    'Crearemos partida nueva
    Set cP2 = New cPartidas
    
    Kilos = ImporteFormateado(Text2(0).Text) - ImporteFormateado(Text2(1).Text)
    cP2.NUmlote = "Orujo" & Format(ID, "0000")
    cP2.Cantidad = Kilos
    cP2.codAlmac = cStock.codAlmac
    Cad = DevuelveDesdeBD(conAri, "articOrujo", "vallparam", "1", "1")
    cStock.codartic = Cad
    cStock.Cantidad = Kilos
    cStock.LineaDocu = 1
    cP2.codartic = cStock.codartic
    cP2.codProve = 1
    cP2.Fecha = cStock.Fechamov
    cP2.NumAlbar = cStock.Documento
    'cp.NUmlote = lo he asginado arriba
    If Not cP2.Insertar Then Err.Raise 513, , "Creando partida orujo "
    If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock orujo"
    
    
    
    
    
    
    'Las lineas de oliva, movimiento nuevo para dar de baja
    Cad = "select * from tmpnlotes where codusu=" & vUsu.Codigo
    Set RT = New ADODB.Recordset
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cStock.Documento = Format(ID, "00000")
    cStock.Importe = 0
    cStock.LineaDocu = 1
    cStock.Trabajador = vUsu.CodigoTrabajador
    While Not RT.EOF
        cStock.codartic = RT!codartic
        Kilos = CCur(TransformaPuntosComas((RT!numlotes)))
        cStock.Cantidad = Kilos
        cStock.tipoMov = "S"
        Cad = "Alb: " & RT!NumAlbar & " " & RT!FechaAlb & " .> " & RT!NomArtic
        If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock: " & Cad
        
        
        
        'El rendimiento aplicado a cada linea
        Cad = " numalbar=" & DBSet(RT!NumAlbar, "T") & " AND codartic = " & DBSet(RT!codartic, "T") & " AND 1"
        Cad = DevuelveDesdeBD(conAri, "entrada", "vallentradacamionlineas", Cad, "1 ORDER BY entrada desc")  'No deberia haber mas de una
        If Cad = "" Then Err.Raise 513, , "NO se encuentra la entrada de camion para el albarán: " & RT!NumAlbar & " Art: " & RT!codartic
        
        Cad = " WHERE  numalbar=" & DBSet(RT!NumAlbar, "T") & " AND entrada =" & Cad
        Cad = ",rdtoRea=" & DBSet(RT!Cantidad, "N") & Cad
        Cad = "UPDATE vallentradacamionlineas set rendimiento=" & DBSet(RT!Cantidad, "N") & Cad
        
        
        
        If Not EjecutaSQL(conAri, Cad, False) Then Err.Raise 513, , "Actualizando albaran entrada camion: " & Cad
        
        
        RT.MoveNext
    Wend
    RT.Close
    
    
    
    
    
    'Faltaria ver el talco
    '-----------------------------------------------------------------
    articuloTalco = ""
    Cad = "Select dosis , numlote from vallalmazaraproceso where id =" & ID
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT!dosis) Then
            articuloTalco = DevuelveDesdeBD(conAri, "arttalco", "vallparam", "1", "1")
            If articuloTalco = "" Then
                MsgBox "Pone dosis talco y no esta configurado proceso continua sin tratar talco", vbExclamation
            Else
                Kilos = ImporteFormateado(Text2(0).Text)
                Kilos = Round((Kilos * RT!dosis) / 100, 2)
                Set cP2 = Nothing
                Set cP2 = New cPartidas
                
                If Not cP2.LeerDesdeArticulo(articuloTalco, 1, CStr(RT!NUmlote)) Then Err.Raise 513, , "Leyendo lote talco: " & articuloTalco & " " & RT!NUmlote
               
                cP2.Cantidad = Kilos
                cStock.codartic = cP2.codartic
                cStock.Cantidad = Kilos
                cStock.LineaDocu = 1
                cStock.tipoMov = "S"
                'cp.NUmlote = lo he asginado arriba
                cP2.IncrementarCantidad -Kilos
                If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock talco"
            
            
            
            
            End If
        End If
    End If
    RT.Close
    
    
    NumRegElim = 0
    If NumDeposito = 100 Then
        
        If vT Is Nothing Then
            Set vT = New CTiposMov
            vT.Leer "LOV"
        End If
    
        'Ahora tengo que HACER un nuevo COUPAGE automatico entre el deposito 100 y el deposito destino
        Cad = "Select max(codigo) from olicoupage"
        RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = DBLet(RT.Fields(0), "N") + 1
        RT.Close
        
        Set cP2 = Nothing
        Set cP2 = New cPartidas
        cP2.NUmlote = vT.ConseguirContador(vT.TipoMovimiento)
        cP2.NUmlote = "MOSTRA" & cP2.NUmlote & "-"
        If Month(Now) < 11 Then
            cP2.NUmlote = cP2.NUmlote & Year(Now) - 1
        Else
            cP2.NUmlote = cP2.NUmlote & Year(Now)
        End If
        
        
        Cad = "INSERT INTO olicoupage(codigo,codartic,fecha,descripcion,YaCreado,codalmac,numlote,Deposito) VALUES ("
        Cad = Cad & NumRegElim & ",'" & CP1.codartic & "'," & DBSet(Now, "FH") & ",'"
        Cad = Cad & "Molturacion. ID: " & CP1.NumAlbar & "',0, " & CP1.codAlmac & "," & DBSet(cP2.NUmlote, "T") & ","
        Cad = Cad & Combo1.ListIndex + 1 & ")"
        conn.Execute Cad
        
        
        
        
        'Leemos lo que hay en el deposito AHORA
        If Not cD.LeerDatos(Combo1.ListIndex + 1, True) Then Err.Raise 513, , "Leyendo deposito coupage: " & Combo1.ListIndex + 1
        Kilos = Round(ImporteFormateado(Text2(1).Text) * 0.916, 2) 'De la nueva produccion + lo que habia
        KilosTotales = cD.Kilos + Kilos
        Cad = "INSERT INTO olicoupagelin(codigo,codartic,kilos) VALUES (" & NumRegElim & ","
        Cad = Cad & DBSet(CP1.codartic, "T") & "," & DBSet(KilosTotales, "N") & ")"
        conn.Execute Cad
        
        
        
        Set cP2 = Nothing
        Set cP2 = New cPartidas
        If Not cP2.Leer(cD.idPartida) Then Err.Raise 513, , "Leyendo partida deposito: " & cD.NumDeposito & " ->" & cD.NUmlote
        
        'Las dos linea del coupage con lotes
        Cad = "INSERT INTO olicoupagelinlotes(codigo,codartic,linea,numlote,cantlote,fincuba,deposito) VALUES ("
        Cad = Cad & NumRegElim & ",'" & cP2.codartic & "',1," & DBSet(cP2.NUmlote, "T") & "," & DBSet(cP2.Cantidad, "N")
        Cad = Cad & ",1," & cD.NumDeposito & "),(" & NumRegElim & "," & DBSet(CP1.codartic, "T") & ",2,"
        Cad = Cad & DBSet(CP1.NUmlote, "T") & "," & DBSet(Kilos, "N") & ",0,100)"
        
        conn.Execute Cad
        
            
        vT.IncrementarContador vT.TipoMovimiento
        
    
    
    End If
    
    
    'En la tablaproceso guardo el dato de deposito y fecha fin
    Cad = "UPDATE vallalmazaraproceso SET HoraFin =" & DBSet(Now, "H") & ",  deposito=" & Combo1.ListIndex + 1
    'Litros producidos y kilos utlizados
    Cad = Cad & ", kilos =" & DBSet(Text2(0).Text, "N")
    Cad = Cad & ", Litros =" & DBSet(Text2(1).Text, "N")
    If articuloTalco = "" Then
        articuloTalco = "NULL"
    Else
        articuloTalco = DBSet(articuloTalco, "T")
    End If
    Cad = Cad & ", articuloTalco =" & articuloTalco
    
    'Lote y articulo
    Cad = Cad & ", loteproducido =" & DBSet(CP1.NUmlote, "T")
    Cad = Cad & ", artproducido =" & DBSet(CP1.codartic, "T")
     
    
    
    Cad = Cad & " WHERE id=" & Me.ID
    conn.Execute Cad
    

    
    
    cerrarProcesoAlm = True
    
eCerrarProcesoAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set CP1 = Nothing
    Set cP2 = Nothing
    Set cStock = Nothing
    Set cD = Nothing
    Set cLot = Nothing
    Set RT = Nothing
    Set vT = Nothing
End Function
