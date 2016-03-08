VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProdPalets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Palets"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmProdPalets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTipoImpresion 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Tipo|N|N|0||prodpalets|TipoImpresion|||"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   240
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Cajas|N|N|0||prodpalets|cajasprod|0|N|"
      Text            =   "Text1"
      Top             =   1680
      Width           =   1365
   End
   Begin VB.CheckBox chkMostrar 
      Caption         =   "Mostrar"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   1680
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   1320
      TabIndex        =   15
      Top             =   2280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cajas"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Articulo"
         Object.Width           =   5645
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Id|N|N|0||prodpalets|idpalet|00000|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Tag             =   "Inicio|H|S|||prodpalets|fhFin|dd/mm/yyyy hh:mm:ss||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1725
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Inicio|H|N|||prodpalets|fhinicio|dd/mm/yyyy hh:mm:ss||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   3135
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   210
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5700
      TabIndex        =   7
      Top             =   5715
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   5715
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   240
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir seleccion"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5280
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame FrameBtn 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   2400
      Width           =   975
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Refrescar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nueva caja"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Quitar caja"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Impresión"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   27
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas en palet"
      Height          =   195
      Index           =   7
      Left            =   1920
      TabIndex        =   26
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas produccion"
      Height          =   195
      Index           =   6
      Left            =   5400
      TabIndex        =   25
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas expedidas"
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   24
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   6720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   "CAJAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6840
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Fin"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   10
      Top             =   600
      Width           =   495
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
Attribute VB_Name = "frmProdPalets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean



Private Sub chkMostrar_Click()
    
    
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    CargaPalets False
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    PonerModo 0
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    cad = "(idpalet=" & Text1(0).Text & ")"
                    If SituarData(Data1, cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                        PonerFoco Text1(0)
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("prodpalets", "idpalet")
    Text1(1).Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        '### A mano
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    '### A mano

    
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String
    If Modo <> 2 Then Exit Sub
    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    'De momento NO dejo
    
    cad = DevuelveDesdeBD(conAri, "count(*)", "prodcajasprod", "idpalet", CStr(Data1.Recordset!IdPalet))
    If cad = "" Then cad = "0"
    If Val(cad) > 0 Then
        MsgBox "No se puede borrar(produccion)!!!!", vbExclamation
        Exit Sub
    End If
        
  
    'prodpaletstraza
    cad = DevuelveDesdeBD(conAri, "count(*)", "prodpaletstraza", "idpalet", CStr(Data1.Recordset!IdPalet))
    If cad = "" Then cad = "0"
    If Val(cad) > 0 Then
        MsgBox "El palet esta asociado a trazabilidad", vbExclamation
        Exit Sub
    End If
    
    
    
    
    '### a mano
    cad = "¿Seguro que desea eliminar el palet?" & vbCrLf
    cad = cad & vbCrLf & "Id: " & Format(Data1.Recordset.Fields(0), "0000")
    cad = cad & vbCrLf & "Fecha inicio: " & Data1.Recordset.Fields(2)
    If Text2(0).Text <> "" Then cad = cad & vbCrLf & vbCrLf & "Cajas: " & Text2(0).Text
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        Screen.MousePointer = vbHourglass
        
        Conn.Execute "UPDATE prodcajas set idpalet=NULL where idpalet=" & Text1(0).Text
        Conn.Execute "DELETE FROM prodpalets where idpalet=" & Text1(0).Text
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Forma de Pago", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    If IsNull(Data1.Recordset!fhFin) Then
        MsgBox "Palet no esta cerrado", vbExclamation
        Exit Sub
    End If


    'IDPALET

    
    cad = "Id: " & Format(Data1.Recordset.Fields(0), "0000") & "   "
    cad = cad & Data1.Recordset.Fields(2) & "  Cajas: " & Data1.Recordset.Fields(4)
    
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 16
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 16  'impr
        .Buttons(11).Image = 48  'imprimir seleecion
        .Buttons(13).Image = 15  'Salir
        .Buttons(16).Image = 6  'Primero
        .Buttons(17).Image = 7  'Anterior
        .Buttons(18).Image = 8  'Siguiente
        .Buttons(19).Image = 9  'Último
    End With
    
    With Me.Toolbar3
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 16   'Borrar
        .Buttons(3).Image = 3   'Insertar Nuevo
        .Buttons(4).Image = 5   'Modificar
    End With
    
    
    
    LimpiarCampos

    '## A mano
    NombreTabla = "prodpalets"
    Ordenacion = " ORDER BY idPalet"
           
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    'COMBO
    CargaComboTipoImpresionPalet Me.cboTipoImpresion
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    '## A mano
    Data1.RecordSource = "Select * from " & NombreTabla & " where idpalet=-1"
    Data1.Refresh
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub



Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboTipoImpresion.ListIndex = -1
    CargaPalets True
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Integer

    If CadenaDevuelta <> "" Then
    
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
    
    End If
End Sub
    


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
   
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod Forma de Pago
           If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
           End If
            
        Case 2, 3 'FH
          '  If fechaok Then
                
        Case 4  'nº cajas
            PonerFormatoEntero Text1(Index)
        
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String


    'Cambio los CHK para que pueda bsucar por fechas
    Me.Text1(1).Tag = "Inicio|F|N|||prodpalets|fhinicio|dd/mm/yyyy hh:mm:ss||"
    Me.Text1(2).Tag = "Fin|F|S|||prodpalets|fhFin|dd/mm/yyyy hh:mm:ss||"
    

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then
        
        cadB = Replace(cadB, "prodpalets.fhinicio", "date(prodpalets.fhinicio)")
        cadB = Replace(cadB, "prodpalets.fhFin", "date(prodpalets.fhFin)")
        
    End If
    
    Me.Text1(1).Tag = "Inicio|H|N|||prodpalets|fhinicio|dd/mm/yyyy hh:mm:ss||"
    Me.Text1(2).Tag = "Fin|H|S|||prodpalets|fhFin|dd/mm/yyyy hh:mm:ss||"
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim cad As String

        'Llamamos a al form
        '##A mano
        cad = ""
        cad = cad & ParaGrid(Text1(0), 30, "Código")
        cad = cad & ParaGrid(Text1(1), 70, "Fecha incio")
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = cadB
            
            HaDevueltoDatos = False
            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Mto. palets"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1 'Conexión a BD: Ariges
'            If imgFPago(0).Tag = -1 Then
'                frmB.vBuscaPrevia = chkVistaPrevia
'            Else
'                frmB.vBuscaPrevia = True
'            End If
            frmB.vCargaFrame = False
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
''            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                PonerFoco Text1(kCampo)
'                PonerModo Modo
'            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
'         MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If

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
    PonerCamposForma Me, Me.Data1

    CargaPalets False
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = B
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 0
    If Not Data1.Recordset.EOF Then
        NumReg = 1
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    FrameBtn.visible = Modo = 2 And NumReg > 0
    
    '----------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    Me.cboTipoImpresion.Enabled = cmdAceptar.visible
    
    
    BloquearText1 Me, Modo
    
 
    B = (Modo = 3) 'Insertar
    'Campos Importe Mínimo y % Adelantado
    If B Then
        For i = 8 To 9
            BloquearTxt Text1(i), True
        Next i
    End If

     chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean

    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    B = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnNuevo.Enabled = Not B
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    
    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
     
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If
     
    If Not B Then Exit Function
    DatosOk = B
End Function


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 9
            If Modo <> 2 Then Exit Sub
            If Me.Data1.Recordset Is Nothing Then Exit Sub
            If Me.Data1.Recordset.EOF Then Exit Sub
            
            ImprimirEtiquetaPalet
        Case 11
            ImprimirEtiquetaPaletSELECCION
            
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub CargaPalets(limpiar As Boolean)
Dim R As ADODB.Recordset
Dim It As ListItem
Dim Articulo As String
Dim byt As Byte
Dim CajasMetidas As String
Dim K As Integer
Dim Aux As String


    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Me.ListView1.ListItems.Clear
    
    If limpiar Then Exit Sub
    
    
    Set R = New ADODB.Recordset
    R.Open "select * from prodcajas where idpalet=" & Text1(0).Text & " ORDER BY lotetraza,idcaja", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    ListView1.Tag = 0 'lotetraza
    While Not R.EOF
        NumRegElim = NumRegElim + 1
        If Me.chkMostrar.Value = 1 Then
            Set It = ListView1.ListItems.Add()
            It.Text = Format(R!lotetraza, "00000000") & Format(R!idcaja, "00000")
            CajasMetidas = CajasMetidas & It.Text & "|"
            If ListView1.Tag <> R!lotetraza Then
                Articulo = " prodtrazlin.codigo=prodlin.codigo and prodtrazlin.idlin= prodlin.idlin and prodlin.codartic=sartic.codartic AND lotetraza"
                Articulo = DevuelveDesdeBD(conAri, "nomartic", "prodtrazlin,prodlin,sartic", Articulo, R!lotetraza)
                byt = 1
                It.SubItems(1) = Articulo
                ListView1.Tag = R!lotetraza
            Else
                If byt > 9 Then
                    byt = 0
                    It.SubItems(1) = Articulo
                End If
                byt = byt + 1
            End If
        End If
        R.MoveNext
    Wend
    R.Close
    Text2(0).Text = Format(NumRegElim, "0000")
    NumRegElim = 0
    
    
    
    'Cajas expedidas
    R.Open "select * from srepartolotcaj where idpalet=" & Text1(0).Text & " ORDER BY lotetraza,idcaja", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    ListView1.Tag = 0 'lotetraza
    While Not R.EOF
        NumRegElim = NumRegElim + 1
        If Me.chkMostrar.Value = 1 Then
        
            Aux = Format(R!lotetraza, "00000000") & Format(R!idcaja, "00000")
            If InStr(1, CajasMetidas, Aux) > 0 Then
                For K = 1 To Me.ListView1.ListItems.Count
                    If Me.ListView1.ListItems(K).Text = Aux Then
                        'Este es el item, nos salimos
                        Me.ListView1.ListItems(K).ForeColor = vbBlue
                        Exit For
                    End If
                Next
            Else
            
                Set It = ListView1.ListItems.Add()
                It.Text = Aux
                If R!idreparto < 0 Then
                    'Perdida o baja
                    It.ForeColor = vbMagenta
                Else
                    It.ForeColor = vbBlue
                End If
                If ListView1.Tag <> R!lotetraza Then
                    Articulo = " prodtrazlin.codigo=prodlin.codigo and prodtrazlin.idlin= prodlin.idlin and prodlin.codartic=sartic.codartic AND lotetraza"
                    Articulo = DevuelveDesdeBD(conAri, "nomartic", "prodtrazlin,prodlin,sartic", Articulo, R!lotetraza)
                    byt = 1
                    It.SubItems(1) = Articulo
                    ListView1.Tag = R!lotetraza
                Else
                    If byt > 9 Then
                        byt = 0
                        It.SubItems(1) = Articulo
                    End If
                    byt = byt + 1
                End If
        
            End If
        End If
        R.MoveNext
    Wend
    R.Close
    Text2(1).Text = Format(NumRegElim, "0000")
    NumRegElim = 0
    
    
    R.Open "select count(*) from prodcajasprod where idpalet=" & Text1(0).Text
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then Text2(2).Text = R.Fields(0)
    End If
    R.Close
    
    Set R = Nothing
    

    
End Sub





Private Sub ImprimirEtiquetaPalet()
Dim cPal As CPalet
Dim C As Collection
    Set cPal = New CPalet
    If cPal.Leer(Val(Text1(0).Text)) Then
    
        If Val(Text2(0).Text) = 0 Then
            MsgBox "No tiene cajas el palet", vbExclamation
        Else
            cPal.CargaDatosPalet C, True, kCampo, False
            
            kCampo = CByte(MsgBox("Utilizar impresora SATO?", vbQuestion + vbYesNoCancel))
            If kCampo <> vbCancel Then
                If kCampo = vbNo Then
                    'OK YA ESTA CERRADO. Ahora a imprimir la sabana santa
                    CadenaConsulta = "{tmppartidas.codusu}=" & vUsu.Codigo
                    LlamaImprimirGral CadenaConsulta, "", 0, "EtiqPalet.rpt", "Etiqueta palet: " & cPal.ID
                    CadenaConsulta = Data1.RecordSource
                
                Else
                    ImprimirPalet cPal.ID, cPal.TipoImpresion
                
                End If
            End If
        End If
    End If
    Set cPal = Nothing
    kCampo = 0
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    'no dejo hacer nada

'
'    If Modo <> 2 Then Exit Sub
'
'    Select Case Button.index
'    Case 0
'            CargaPalets False
'
'    Case 3
'            MsgBox "Solo desde el lector de barras"
'    Case 4
'            'Insertar o eliminar cajas
'            If Me.ListView1.ListItems.Count = 0 Then Exit Sub
'            If ListView1.SelectedItem Is Nothing Then Exit Sub
'
'            'Deberia
'
'            If IsNull(Me.Data1.Recordset!fhFin) Then
'                If MsgBox("El palet NO esta cerrado. Desea eliminar la caja: " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'            Else
'                If MsgBox("Desea eliminar la caja: " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'            End If
'            EjecutaSQL conAri, "DELETE FROM prodcajas where lotetraza= " & Val(Mid(ListView1.SelectedItem.Text, 1, 8)) & " AND idcaja = " & Val(Mid(ListView1.SelectedItem.Text, 9)) & " AND idpalet =" & Text1(0).Text, True
'
'            ListView1.ListItems.Remove ListView1.SelectedItem.index
'
'            If Not IsNull(Me.Data1.Recordset!fhFin) Then
'                'Hay que actualizar las cajas
'                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", "idpalet", Text1(0).Text)
'                If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
'                CadenaDesdeOtroForm = "UPDATE prodpalets set cajasprod=" & CadenaDesdeOtroForm & " WHERe idpalet=" & Text1(0).Text
'                EjecutaSQL conAri, CadenaDesdeOtroForm, True
'
'
'            End If
'            Text2.Text = Val(Text2(0).Text) - 1
'            Text1(3).Text = Text2.Text
'    End Select
End Sub





Private Sub ImprimirEtiquetaPaletSELECCION()
Dim cPal As CPalet
Dim C As Collection
Dim Aux As String
Dim SATO As Boolean

    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset Is Nothing Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    
    
    Aux = "Desea imprimir las " & Data1.Recordset.RecordCount & " etiqueta(s) de palet "
    Aux = Aux & "por la impresora SATO?"
    
    Aux = Aux & vbCrLf & vbCrLf & "(Las etiquetas de palets sin cajas NO se imprimiran)    "
    
    
    Aux = MsgBox(Aux, vbQuestion + vbYesNoCancel)
    If CByte(Aux) = vbCancel Then Exit Sub
    SATO = False
    If CByte(Aux) = vbYes Then SATO = True
    
    
    If SATO Then
        If Data1.Recordset.RecordCount > 10 Then
            If MsgBox("Son muchas etiquetas, continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    
    
    
    Set cPal = New CPalet
    NumRegElim = Data1.Recordset.AbsolutePosition
    'Preparamos
    Conn.Execute "DELETE FROM tmppartidas  WHERE codusu = " & vUsu.Codigo
    
    Data1.Recordset.MoveFirst
    
    While Not Data1.Recordset.EOF
    
    
        If cPal.Leer(Val(Data1.Recordset.Fields(0))) Then
            
        
        
        
            cPal.CargaDatosPalet C, True, kCampo, Not SATO
            
            If SATO Then
                ImprimirPalet cPal.ID, cPal.TipoImpresion
                Espera 0.85
                
                
            Else
                

            
            
                
            
            End If
            
        End If
        Data1.Recordset.MoveNext
    Wend
    Data1.Recordset.MoveFirst
    Data1.Recordset.Move NumRegElim - 1
    Screen.MousePointer = vbDefault
    If Not SATO Then
         Aux = "{tmppartidas.codusu}=" & vUsu.Codigo
         LlamaImprimirGral Aux, "", 0, "EtiqPalet.rpt", "Etiquetas palets "
    End If
    
    
    Set cPal = Nothing
    kCampo = 0
    
End Sub




