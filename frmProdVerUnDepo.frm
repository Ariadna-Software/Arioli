VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProdVerUnDepo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposito"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos depósito"
      TabPicture(0)   =   "frmProdVerUnDepo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ListView1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdModKilos"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmProdVerUnDepo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(63)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "imgFecha(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtFecha(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdVer"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdVer 
         Caption         =   "Movimientos"
         Height          =   375
         Left            =   -71520
         TabIndex        =   24
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdModKilos 
         Height          =   375
         Left            =   5880
         Picture         =   "frmProdVerUnDepo.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ajustar cantidad en deposito"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   -74040
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   6480
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Linea"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lote"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Previsto"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   3120
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   7
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   5
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   9551
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lote"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UltimaLlevaElorden"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   -74280
         Picture         =   "frmProdVerUnDepo.frx":0A3A
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   -74880
         TabIndex        =   22
         Top             =   6480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "  Produccion     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   6120
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Partida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   15
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Capacidad(L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Litros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4200
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Deposito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   7575
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmProdVerUnDepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumDepo As Integer
Public idProd  As String ' codigo|linea|"

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim miSQL As String
Dim PrimVe As Boolean

Private Sub CargarUnDeposito()
Dim cad As String
Dim PorcentajeLleno As Currency
Dim It As ListItem
    
    
    cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote,spartidas.id"
    cad = cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    cad = cad & " left join sartic on spartidas.codartic=sartic.codartic WHERE numdeposito = " & NumDepo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

        Me.Text1(2).Text = DBLet(miRsAux!NUmlote, "T")
        Text1(0).Text = NumDepo
        
        
        Text1(3).Text = Format(miRsAux!Capacidad, "#,##0")
        
        
      
        If Not IsNull(miRsAux!ID) Then
            Text1(6).Text = miRsAux!ID
            Text1(7).Text = miRsAux!codartic
            Text1(1).Text = miRsAux!NomArtic
            PorcentajeLleno = miRsAux!Kilos / miRsAux!FactorConversion
        Else
            PorcentajeLleno = 0
            SSTab1.TabVisible(0) = False
        End If
        Text1(4).Text = Format(PorcentajeLleno, FormatoCantidad)
        Text1(5).Text = Format(miRsAux!Kilos, FormatoCantidad)
        
        
        'PorcentajeLleno = miRsAux!Kilos / miRsAux!FactorConversion   --> esta arriba
        PorcentajeLleno = Round((PorcentajeLleno * 100) / miRsAux!Capacidad, 2)
        If PorcentajeLleno > 100 Then
            PorcentajeLleno = 100
        ElseIf PorcentajeLleno < 0 Then PorcentajeLleno = 0
        End If
        
        
        PorcentajeLleno = CInt((Me.ShFondo(0).Height * PorcentajeLleno / 100))
        ShDeposito(0).Height = PorcentajeLleno
        PorcentajeLleno = ShFondo(0).Height - PorcentajeLleno
        ShDeposito(0).Top = ShFondo(0).Top + PorcentajeLleno
    
        

    miRsAux.Close
    Me.ListView1.ListItems.Clear
    
    If idProd <> "" Then
        'Esta en produccion
        
        cad = "select prodlin.codartic,cantesti,lineaprod,lotetraza,nomartic,cantesti from prodlin,prodtrazcompo,sartic"
        cad = cad & " where prodlin.codigo= prodtrazcompo.codigo AND prodlin.codartic=sartic.codartic AND prodlin.idlin = prodtrazcompo.idlin"
        cad = cad & " AND numlote = " & DBSet(Text1(2).Text, "T") & " and cantutili is null"
        'Cad = Cad & " and  prodlin.codigo =" & RecuperaValor(idProd, 1)
        'Cad = Cad & " and  prodlin.idlin =" & RecuperaValor(idProd, 2)
       cad = cad & " ORDER BY lineaprod"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        ListView1.ColumnHeaders(4).Width = 0
        
        While Not miRsAux.EOF
              Set It = ListView1.ListItems.Add()
              It.Text = miRsAux!lineaprod
              It.SubItems(1) = miRsAux!NomArtic
              It.SubItems(2) = miRsAux!lotetraza
              'IT.SubItems(3) = Format(miRsAux!cantesti, "#,##0")
        
              miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    Set miRsAux = Nothing
End Sub

Private Sub cmdModKilos_Click()
    If vUsu.Nivel > 0 Then Exit Sub
    If SSTab1.TabVisible(0) = False Then Exit Sub
    
    
    CadenaDesdeOtroForm = ""
    frmVarios.opcion = 12
    frmVarios.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'OK. Ajustamos
        cmdModKilos.Tag = 1
        
        'Uopdateamos slotes y prodeposito
        UpdatearKilos
        
        Unload Me
    End If
    
    
    
End Sub


Private Function UpdatearKilos() As Boolean


    'El deposito
    miSQL = TransformaComasPuntos(ImporteFormateado(CStr(CadenaDesdeOtroForm)))
    miSQL = "UPDATE proddepositos SET kilos = " & miSQL
    miSQL = miSQL & " WHERE numdeposito=" & Text1(0).Text
    conn.Execute miSQL
    
    'La partida
    miSQL = TransformaComasPuntos(ImporteFormateado(CadenaDesdeOtroForm))
    miSQL = "UPDATE spartidas SET cantotal = " & miSQL
    miSQL = miSQL & " WHERE id=" & Text1(6).Text 'idpartida
    conn.Execute miSQL
    
    
    
    'Y el log
    
    Set LOG = New cLOG
    
    miSQL = "Deposito: " & Text1(0).Text & "      " & Text1(1).Text & " [" & Text1(7).Text & "]"
    miSQL = miSQL & vbCrLf & "Partida: " & Text1(6).Text & "    LOTE: " & Text1(2).Text
    miSQL = miSQL & vbCrLf & "Kilos. Anterior: " & Text1(5).Text & "      Actualizado: " & CadenaDesdeOtroForm
    
    Set LOG = New cLOG
    LOG.Insertar 12, vUsu, miSQL
    Set LOG = Nothing
    
End Function


Private Sub cmdVer_Click()
    frmAlmpartidasMov.VerPartida = Val(Text1(6).Text)
    frmAlmpartidasMov.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVe Then
        PrimVe = False
        CargarUnDeposito
        CargaHco
        If Val(Text1(0).Text) >= 100 Then
            Me.Label1(0).ForeColor = &H4080&
        Else
            Me.Label1(0).ForeColor = &H80000012
        End If
        cmdModKilos.Tag = 0
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimVe = True
    Screen.MousePointer = vbHourglass
    Me.Icon = frmppal.Icon
    PrimVe = True
    limpiar Me
    Me.ListView1.ListItems.Clear
    Me.ListView2.ListItems.Clear
    
    
    
    cmdModKilos.visible = vUsu.Nivel <= 0
    
    LeerGuardarFecha True
    txtFecha(0).Tag = txtFecha(0).Text
    Me.cmdVer.visible = False
    If vParamAplic.QUE_EMPRESA = 4 Then
        Me.Text1(2).FontSize = 14
        Me.cmdVer.visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LeerGuardarFecha False
    
    If cmdModKilos.Tag = "1" Then
        CadenaDesdeOtroForm = "MOD"
    Else
        CadenaDesdeOtroForm = ""
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    miSQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
   
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   miSQL = ""
   frmC.Show vbModal
   Set frmC = Nothing
   If miSQL <> "" Then
        txtFecha(Index).Text = miSQL
        txtFecha_LostFocus Index
    End If
        
    
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
            'txtFecha(0).Tag = T
            CargaHco
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(0).Text = txtFecha(0).Tag
        End If
    End If
    
End Sub




Private Sub CargaHco()
Dim It As ListItem

    Screen.MousePointer = vbHourglass
    ListView2.ListItems.Clear
    'La columna esta oculta y lleva ekl campo fechahora en formato yyyymmddhhnnss
    ListView2.SortKey = 3
    ListView2.Sorted = True

    Set miRsAux = New ADODB.Recordset
    miSQL = "select horamovi,numlote,tipoaccion from proddepositoshco where numdeposito=" & NumDepo
    miSQL = miSQL & " AND horamovi >='" & Format(Me.txtFecha(0).Text, "yyyy-mm-dd") & " 00:00:00'"
    miSQL = miSQL & " ORDER BY horamovi"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
            Set It = ListView2.ListItems.Add
            It.Text = Format(miRsAux!horamovi, "dd/mm/yyyy hh:nn:ss")
            It.SubItems(3) = Format(miRsAux!horamovi, "yyyymmddhhnnss")
            '***** VER InsertarEnHco enla clase DEPOSITO
            '   0 .- Albaran de compra
            '   1 .- Coupage Entrada
            '   2 .-  "      salida
            '   3 .- Trasiego entrada
            '   4 .-    "     salida
            '   5 .-  Produccion
            '   6 .- Venta directa
            '   7 .- Forzar vaciado
            '   8 .- FIltrado entrada
            '   9 .-   "    salida
            Select Case miRsAux!tipoaccion
            Case 0
                'ALBARAN COMPRA
                miSQL = "Compra"
            Case 1
                'COUPAGE ENTRADA
                miSQL = "Coupage entrada"
            Case 2
                'COU salida
                miSQL = "Coupage salida"
            Case 3
                'TRASIEGO E
                miSQL = "Trasiego Entrada"
            Case 4
                'TRAS SAL
                miSQL = "Trasiego Salida"
                
            Case 5
                'NO esta en hco, esta en protrazlin
                miSQL = "Parte produccion"
            Case 6
                'Venta directa
                miSQL = "Venta directa"
            Case 7
                'VACIADO
                miSQL = "Vaciado"
            
            Case 8
                'VACIADO
                miSQL = "Filtrado entrada"
            Case 9
                'VACIADO
                miSQL = "Filtrado salida"
            End Select
            It.SubItems(1) = miSQL
            It.SubItems(2) = miRsAux!NUmlote
            If vParamAplic.QUE_EMPRESA = 4 Then It.SubItems(2) = Mid(miRsAux!NUmlote, 7)
            
            miRsAux.MoveNext
            
    Wend
    miRsAux.Close
    
    
    
    'Metemos las producciones
    miSQL = "select fhinicio,prodlin.codigo,prodlin.idlin,lotetraza"
    miSQL = miSQL & " from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin"
    miSQL = miSQL & "  and prodtrazlin.depositol = " & NumDepo & "  AND fhinicio >='" & Format(Me.txtFecha(0).Text, "yyyy-mm-dd") & " 00:00:00'"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add
        It.Text = Format(miRsAux!fhinicio, "dd/mm/yyyy hh:nn:ss")
        It.SubItems(3) = Format(miRsAux!fhinicio, "yyyymmddhhnnss")
        It.SubItems(1) = "Produccion " & miRsAux!Codigo & "/" & miRsAux!idlin
        'IT.SubItems(2) = miRsAux!lotetraza
        
        miRsAux.MoveNext

    Wend
    miRsAux.Close
    Screen.MousePointer = vbDefault
End Sub


Private Sub LeerGuardarFecha(Leer As Boolean)
Dim NF As Integer
Dim F1 As Date
    On Error GoTo eLeerGruadarFecha

    miSQL = App.Path & "\fecdep.dat"
    NF = FreeFile
    If Leer Then
        F1 = "01/06/2014"
        If Dir(miSQL, vbArchive) <> "" Then
            Open miSQL For Input As #NF
            Line Input #NF, miSQL
            Close #NF
            
            If Trim(miSQL) <> "" Then
                If IsDate(miSQL) Then F1 = CDate(miSQL)
            End If
        
        End If
        
        Me.txtFecha(0).Text = Format(F1, "dd/mm/yyyy")

    Else
        If Me.txtFecha(0).Text <> Me.txtFecha(0).Tag Then
            Open miSQL For Output As #NF
            Print #NF, Me.txtFecha(0).Text
            Close #NF
        End If
    End If
    Exit Sub
eLeerGruadarFecha:
    Err.Clear
End Sub
