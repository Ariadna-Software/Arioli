VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProdVerUnDepoVall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposito (COOP)"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12840
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos depósito"
      TabPicture(0)   =   "frmProdVerUnDepoVall.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdModKilos"
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(2)=   "Text1(1)"
      Tab(0).Control(3)=   "Text1(6)"
      Tab(0).Control(4)=   "Text1(7)"
      Tab(0).Control(5)=   "Text1(3)"
      Tab(0).Control(6)=   "Text1(4)"
      Tab(0).Control(7)=   "Text1(5)"
      Tab(0).Control(8)=   "lblNoMolturar"
      Tab(0).Control(9)=   "Label1(10)"
      Tab(0).Control(10)=   "Line1"
      Tab(0).Control(11)=   "Label1(1)"
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(14)=   "Label1(3)"
      Tab(0).Control(15)=   "Label1(4)"
      Tab(0).Control(16)=   "Label1(5)"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmProdVerUnDepoVall.frx":001C
      Tab(1).ControlEnabled=   -1  'True
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
      TabCaption(2)   =   "Albaranes"
      TabPicture(2)   =   "frmProdVerUnDepoVall.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TreeView1"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   9551
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Movimientos"
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdModKilos 
         Height          =   375
         Left            =   -69120
         Picture         =   "frmProdVerUnDepoVall.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ajustar cantidad en deposito"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   6480
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   19
         Top             =   4080
         Width           =   10575
         _ExtentX        =   18653
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
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2040
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
         Left            =   -74760
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
         Left            =   -72960
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2040
         Width           =   2415
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
         Left            =   -74760
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
         Left            =   -72600
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
         Left            =   -70800
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5415
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   10575
         _ExtentX        =   18653
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LOTE"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Datos"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UltimaLlevaElorden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Arrastrado"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblNoMolturar 
         Alignment       =   1  'Right Justify
         Caption         =   "No molturar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   -67440
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmProdVerUnDepoVall.frx":0A56
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   120
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
         Left            =   -74640
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -64320
         Y1              =   3240
         Y2              =   3240
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
         Left            =   -70200
         TabIndex        =   17
         Top             =   1680
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
         Left            =   -74760
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
         Left            =   -72960
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
         Left            =   -74760
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
         Left            =   -72600
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
         Left            =   -70800
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
Attribute VB_Name = "frmProdVerUnDepoVall"
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
Dim Cad As String
Dim PorcentajeLleno As Currency
Dim It As ListItem
    
    
    Cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote,spartidas.id,nomolturar"
    Cad = Cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    Cad = Cad & " left join sartic on spartidas.codartic=sartic.codartic WHERE numdeposito = " & NumDepo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

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
    
        
        lblNoMolturar.visible = Val(miRsAux!nomolturar) = 1
        
    miRsAux.Close
    Me.ListView1.ListItems.Clear
    
    If idProd <> "" Then
        'Esta en produccion
        
        Cad = "select prodlin.codartic,cantesti,lineaprod,lotetraza,nomartic,cantesti from prodlin,prodtrazcompo,sartic"
        Cad = Cad & " where prodlin.codigo= prodtrazcompo.codigo AND prodlin.codartic=sartic.codartic AND prodlin.idlin = prodtrazcompo.idlin"
        Cad = Cad & " AND numlote = " & DBSet(Text1(2).Text, "T") & " and cantutili is null"
        'Cad = Cad & " and  prodlin.codigo =" & RecuperaValor(idProd, 1)
        'Cad = Cad & " and  prodlin.idlin =" & RecuperaValor(idProd, 2)
       Cad = Cad & " ORDER BY lineaprod"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
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
    frmVarios.Opcion = 12
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
        
        CargarUnDeposito
        CargaHco
        PrimVe = False
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
    lblNoMolturar.visible = False
    Me.ListView1.ListItems.Clear
    Me.ListView2.ListItems.Clear
    
    
    
    cmdModKilos.visible = vUsu.Nivel <= 0
    
    LeerGuardarFecha True
    txtFecha(0).Tag = txtFecha(0).Text
    Me.cmdVer.visible = False
    SSTab1.TabVisible(2) = False
    If vParamAplic.QUE_EMPRESA = 4 Then
        Me.Text1(2).FontSize = 14
        Me.cmdVer.visible = True
        SSTab1.TabVisible(2) = True
        Me.ListView2.ColumnHeaders(6).Width = 1200
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

Private Sub ListView2_DblClick()
    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    
    
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
            '10 molt
    Screen.MousePointer = vbHourglass
    Select Case ListView2.SelectedItem.Tag
    Case 5
            miSQL = ListView2.SelectedItem.SubItems(3)
            miSQL = Mid(miSQL, 6)
            miSQL = Trim(Mid(miSQL, 1, InStr(1, miSQL, "-") - 1))
            With frmProdOrden
                .DatosADevolverBusqueda2 = miSQL
                .Show vbModal
            End With
                    
            
            
    Case 1, 2
            miSQL = ListView2.SelectedItem.SubItems(3)
            If Mid(miSQL, 1, 3) = "CUP" Then
                miSQL = Mid(miSQL, 4)
                    
                With frmAlmCoupage
                    .DatosADevolverBusqueda2 = miSQL
                    .Show vbModal
                End With
                
            End If
            
    Case 10
            frmVallAlmazara.DatosADevolverBusqueda2 = Val(ListView2.SelectedItem.SubItems(3))
            frmVallAlmazara.Show vbModal
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    If Not PrimVe Then
        Screen.MousePointer = vbHourglass
        If Me.TreeView1.Nodes.Count = 0 Then PonerCamposTraza
        Screen.MousePointer = vbDefault
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
Dim arrastrado As Currency

    Screen.MousePointer = vbHourglass
    ListView2.ListItems.Clear
    'La columna esta oculta y lleva ekl campo fechahora en formato yyyymmddhhnnss
    ListView2.SortKey = 4
    ListView2.Sorted = True

    Set miRsAux = New ADODB.Recordset
    miSQL = "select horamovi,numlote,tipoaccion,CantidadMov,descripcion from proddepositoshco where numdeposito=" & NumDepo
    miSQL = miSQL & " AND horamovi >='" & Format(Me.txtFecha(0).Text, "yyyy-mm-dd") & " 00:00:00'"
    miSQL = miSQL & " ORDER BY horamovi"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    arrastrado = 0
    
    While Not miRsAux.EOF
            Set It = ListView2.ListItems.Add
            It.Text = Format(miRsAux!horamovi, "dd/mm/yyyy hh:nn:ss")
            It.SubItems(4) = Format(miRsAux!horamovi, "yyyymmddhhnnss")
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
            Case 10
                miSQL = "Molturacion"
            End Select
            It.SubItems(2) = miSQL
            It.SubItems(1) = miRsAux!NUmlote
            
            If Not IsNull(miRsAux!Descripcion) Then
                It.SubItems(3) = miRsAux!Descripcion
            Else
                It.SubItems(3) = Mid(miRsAux!NUmlote, 7)
            End If
            
            It.SubItems(5) = " " & Format(miRsAux!CantidadMov, FormatoCantidad)
            
            If miRsAux!horamovi >= vParamAplic.FechaActiva Then
                arrastrado = arrastrado + miRsAux!CantidadMov
                It.SubItems(6) = " " & Format(arrastrado, FormatoCantidad)
                It.ListSubItems(6).ForeColor = &H808080
            Else
                It.SubItems(6) = " "
            End If
            It.Tag = miRsAux!tipoaccion
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
        It.SubItems(4) = Format(miRsAux!fhinicio, "yyyymmddhhnnss")
        It.SubItems(2) = "Produccion " & miRsAux!Codigo & "/" & miRsAux!idlin
        'IT.SubItems(2) = miRsAux!lotetraza
        
        miRsAux.MoveNext

    Wend
    miRsAux.Close
    Screen.MousePointer = vbDefault
    
    If ListView2.ListItems.Count > 0 Then
        ListView2.ListItems(ListView2.ListItems.Count).EnsureVisible
    End If
End Sub


Private Sub LeerGuardarFecha(Leer As Boolean)
Dim NF As Integer
Dim F1 As Date
    On Error GoTo eLeerGruadarFecha

    If vParamAplic.QUE_EMPRESA = 4 Then
        If Leer Then txtFecha(0).Text = Format(vParamAplic.FechaActiva, "dd/mm/yyyy")
        
        Exit Sub
    End If

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


'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'
'       Ver albaranes para generear ese aceite PonerCamposTraza
'
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Private Sub PonerCamposTraza()
Dim SQL As String
Dim cP As cPartidas
Dim N
Dim CargaDesdeTmpTraza As Boolean

    
    TreeView1.Nodes.Clear
    Set cP = New cPartidas
    
    conn.Execute "DELETE FROM tmptraza"
    If cP.LeerDesdeArticulo(Text1(7).Text, 1, Text1(2).Text) Then
        'cP.GeneracionHastaMolturacion_
        cP.TrazbilidadDesdeVenta False, False
    End If
    
    
    
    Set miRsAux = New ADODB.Recordset
    SQL = DBLet(cP.NumAlbar, "T")
    If SQL <> "" Then
        'AQUI VERE SI ES UN COUPAGE, PRODUCCION u otro
        CargaDesdeTmpTraza = True
    
        If CargaDesdeTmpTraza Then
                'PRODUCCION
                'Cargar datos produccion
                CargarDatosProduccion
        Else
                SQL = cP.NumAlbar
                
                Set N = TreeView1.Nodes.Add(, , "C" & CStr(TreeView1.Nodes.Count + 1), SQL)
        End If
        
    
    End If
    
    Set miRsAux = Nothing
    Set cP = Nothing
End Sub






Private Sub CargarArbol(padre, NivelPintando As Integer)
Dim N
Dim C As String
Dim Aux As String
Dim contador As Integer
Dim Fin As Boolean
Dim NivelActual As Integer
   
Dim NOdoErroneo As Boolean
Dim Fin2 As Boolean
    
            Fin = False
            Do
                
                If Not miRsAux.EOF Then
                    If NivelPintando = miRsAux!nivle Then
                        
                        C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
                        If vParamAplic.QUE_EMPRESA = 4 Then
                            If miRsAux!nivle >= 1 Then
                                C = miRsAux!idoperacion & " [" & miRsAux!NUmlote2 & "]"
                                If InStr(1, miRsAux!idoperacion, "Id:") > 0 And InStr(1, miRsAux!idoperacion, "Dep:") > 0 Then
                                  C = "MOLT." & miRsAux!idoperacion & " [" & miRsAux!NUmlote2 & "]"
                                End If
                            End If
                        End If
                         
                        Aux = ""
                        If Mid(miRsAux!idoperacion, 3, 1) = "/" Then
                            If Mid(miRsAux!idoperacion, 6, 1) = "/" Then
                                Aux = "MOLT"
                                C = miRsAux!NUmlote2 & "(" & miRsAux!Cantidad & ") " & miRsAux!nomclien
                            End If
                        End If
                        If Aux = "" Then
                            Aux = DevuelveAlbaran(miRsAux!NUmlote2, miRsAux!artic2)
                            C = DevuelveCadena(C, Aux, NivelPintando)
                        End If
                       
                        NOdoErroneo = False
                        If LCase(Mid(miRsAux!idoperacion, 1, 3)) = "err" Then
                            'NO LO PINTO
                            NOdoErroneo = True
                            'Y ademas, lo borro
                            C = "DELETE from tmptraza where codusu =" & vUsu.Codigo & " AND contador =" & miRsAux!contador
                            conn.Execute C
                               
                            
                        End If
                        
                        If Not NOdoErroneo Then
                            
                            contador = TreeView1.Nodes.Count + 1
                            If padre Is Nothing Then
                                Set N = TreeView1.Nodes.Add(, , , C)
                            Else
                                Set N = TreeView1.Nodes.Add(padre, tvwChild, "C" & contador, C)
                            End If
                            N.Tag = miRsAux!contador 'Clave
                           
                        End If
                        
                        NivelActual = miRsAux!nivle
                        miRsAux.MoveNext
                    Else
                        'Stop
                    
                        Fin = True
                    End If
                End If
                If miRsAux.EOF Then
                    Fin = True
                Else
                
                    If miRsAux!nivle > NivelActual Then
                        'Si habia nodo erroneo, entonces tengo que matar los sub nodos ya que no los voy a pintar
                        If Not NOdoErroneo Then
                        CargarArbol N, miRsAux!nivle
                        Fin = False
                        
                        Else
                            Fin2 = False
                            While Not Fin2
                                C = "DELETE from tmptraza where codusu =" & vUsu.Codigo & " AND contador =" & miRsAux!contador
                                conn.Execute C
                                miRsAux.MoveNext
                                
                                If miRsAux.EOF Then
                                    Fin2 = True
                                    Fin = True
                                Else
                                    If miRsAux!nivle > NivelActual Then
                                        'QUe siga en este borrado
                                    Else
                                        Fin2 = True
                                    End If
                                End If
                            Wend
                        End If
                    Else
                       ' Stop
                       ' Fin = True
                    End If
                End If
            Loop Until Fin
        
End Sub



Private Sub CargarDatosProduccion()
Dim C As String
Dim N
Dim contador As Integer
Dim Nivel As Integer
Dim padre As String
Dim Aux As String


    

    C = "select tmptraza.*,nomartic from tmptraza,sartic where codartic=artic2 AND codusu =" & vUsu.Codigo
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Nivel = -1
        
    If Mid(DBLet(miRsAux!idoperacion, "T"), 1, 3) = "Id:" Then
        'Empieza con una molturacion
        Set N = Nothing
    Else
        'Normal
        C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
        Aux = DevuelveAlbaran(miRsAux!NUmlote2, miRsAux!artic2)
        C = DevuelveCadena(C, Aux, 0)
    
                
        contador = TreeView1.Nodes.Count + 1
        Set N = TreeView1.Nodes.Add(, , "C" & contador, C)
        miRsAux.MoveNext
    End If
    Nivel = 0
    
    If Not miRsAux.EOF Then
        Nivel = miRsAux!nivle
        CargarArbol N, Nivel
        'Llamamos a un recursivo para cargar el arbol
    End If
            
    miRsAux.Close
    
    If Not N Is Nothing Then N.EnsureVisible
    
    'If ElAceite <> "" Then CargaCoupageRecursivo RecuperaValor(ElAceite, 1), RecuperaValor(ElAceite, 2), N.Key, EsCou
    
End Sub







Private Function DevuelveAlbaran(NUmlote As String, vArtic As String) As String
'Dim RT As ADODB.Recordset
'Dim Cad As String
'Dim PalWhere As String  'numalbar
'    DevuelveAlbaran = ""
'    Set RT = New ADODB.Recordset
'    Cad = "select * from spartidas where numlote=" & DBSet(NUmlote, "T") & " and codartic='" & vArtic & "'"
'    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Cad = ""
'    If Not RT.EOF Then
'
'        Cad = "select nomprove ,scafpc.numfactu idDoc ,scafpc.fecfactu fecha from scafpc,slifpc where scafpc.codprove=slifpc.codprove and"
'        Cad = Cad & " scafpc.numfactu=slifpc.numfactu and scafpc.fecfactu=slifpc.fecfactu"
'        Cad = Cad & " AND slifpc.numalbar=" & DBSet(RT!NumAlbar, "T") & " and codartic=" & DBSet(RT!codartic, "T")
'        Cad = Cad & " AND scafpc.codprove =" & RT!codProve
'
'    End If
'    RT.Close
'
'
'    If Cad <> "" Then
'        RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        If RT.EOF Then
'
'            RT.Close
'
'            Cad = Mid(RT.Source, InStr(1, UCase(RT.Source), "WHERE") + 6)
'            'Reemplazamos
'
'            Cad = Replace(Cad, "scafpc", "scaalp")
'            Cad = Replace(Cad, "slifpc", "slialp")
'            Cad = Replace(Cad, "fecfactu", "fechaalb")
'            Cad = Replace(Cad, "numfactu", "numalbar")
'            Cad = " from scaalp,slialp where " & Cad
'            Cad = "select nomprove ,scaalp.numalbar idDoc ,scaalp.fechaalb fecha " & Cad
'
            
'            RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'
'        End If
        
'        If Not RT.EOF Then
'            DevuelveAlbaran = "Alb: " & RT!iddoc & "  " & RT!Fecha & "   " & RT!nomprove
'
'
'
'
'        End If
'        RT.Close
'    End If
    
'    Set RT = Nothing
End Function


Private Function DevuelveCadena(Cadena As String, cad2 As String, Nivel As Integer) As String
Dim J As Integer
    
        
    DevuelveCadena = cad2
    J = 124 - (Nivel * 5)
    
    J = J - Len(DevuelveCadena) - Len(Cadena)
    If J < 0 Then J = 0
    DevuelveCadena = Cadena & Space(J) & DevuelveCadena
    
End Function


