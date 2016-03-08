VERSION 5.00
Begin VB.Form frmFichaTecnicaMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha tecnica materia prima"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11760
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton cmdVerDatos 
         Height          =   375
         Left            =   10560
         Picture         =   "frmFichaTecnicaMP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modificar datos"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   375
         Left            =   11040
         Picture         =   "frmFichaTecnicaMP.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   6015
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Text5"
            Top             =   120
            Width           =   4335
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   1200
            Picture         =   "frmFichaTecnicaMP.frx":7254
            ToolTipText     =   "Buscar marca"
            Top             =   120
            Width           =   240
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ACEITE"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   310
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blanqueta"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   310
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8175
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   4440
         Picture         =   "frmFichaTecnicaMP.frx":7C56
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   10200
         TabIndex        =   11
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   8880
         TabIndex        =   10
         Top             =   7680
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Index           =   0
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmFichaTecnicaMP.frx":8658
         Top             =   240
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Index           =   1
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmFichaTecnicaMP.frx":865E
         Top             =   1080
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   975
         Index           =   2
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmFichaTecnicaMP.frx":8664
         Top             =   1920
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Index           =   3
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmFichaTecnicaMP.frx":866A
         Top             =   3000
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Index           =   4
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmFichaTecnicaMP.frx":8670
         Top             =   3600
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   1335
         Index           =   5
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmFichaTecnicaMP.frx":8676
         Top             =   4200
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   1095
         Index           =   6
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmFichaTecnicaMP.frx":867C
         Top             =   5640
         Width           =   7935
      End
      Begin VB.Label Label2 
         Caption         =   "Imagenes asociadas a la categoria para la ficha técnica"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   7200
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   7560
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frmFichaTecnicaMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1



Private Sub cmdActualizarImportes1_Click(Index As Integer)
    If Me.Option1(1).Value Then
        'Esta pidinedo categoria. Tiene que haber uba
        MsgBox "Seleccione la categoria", vbExclamation
        Exit Sub
    End If
    
    
    
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVerDatos_Click()
    If Me.Option1(1).Value Then
        If Text1.Text = "" Then
            MsgBox "Ponga la categoria", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.Option1(0).Value Then
        'Cargamos datos DE BLANQUETA
        Text1.Text = ""
        txtNombre.Text = ""
        CargarDatos -1
    Else
        'De las otras familias
            
        If Not CargarDatos(CInt(Text1.Text)) Then Exit Sub
    
    End If



    PonerFrameVisible Me.Frame3
End Sub

Private Sub Command1_Click(Index As Integer)
Dim i As Integer
Dim SQL As String

    If Index = 0 Then
        'UPDATEAMOS
        For i = 0 To 6
             SQL = "UPDATE sfichtecnicamp set intro =" & DBSet(Text2(i).Text, "T", "S")
             SQL = SQL & ", texto=" & DBSet(Text3(i).Text, "T", "S")
             If Val(Label1.Tag) < 0 Then
                'Blanqueta
                SQL = SQL & " WHERE marca = 3 and linea = " & i + 1
            Else
                SQL = SQL & " WHERE marca = 0 and categoria = " & Label1.Tag & " and linea = " & i + 1
            End If
            Conn.Execute SQL
        Next i
        
    End If
    PonerFrameVisible Frame1

End Sub

Private Sub Command2_Click()
    frmFichaTecIMG_.EsArticulo = False
    frmFichaTecIMG_.vDatos = Text1.Text & "|" & txtNombre.Text & "|"
    frmFichaTecIMG_.Show vbModal
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Me.Top = 1800
    Me.Left = 1000
    limpiar Me
    PonerFrameVisible Me.Frame1
End Sub

Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    Me.Height = F.Height + 550
    Me.Width = F.Width + 350
    If LCase(F.Name) = "frame1" Then
        Me.cmdSalir.Cancel = True
        Frame3.visible = False
    Else
        Frame1.visible = False
        Me.Command1(1).Cancel = True
    End If
End Sub




Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)
    'Abrimos el frmde familia
    Set frmMtoFamilia = New frmAlmFamiliaArticulo
    frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
    frmMtoFamilia.Show vbModal
    Set frmMtoFamilia = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    Me.Frame2.visible = Me.Option1(1).Value
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
    
End Sub

Private Sub Text1_LostFocus()
    txtNombre.Text = ""
    If PonerFormatoEntero(Text1) Then
        txtNombre.Text = PonerNombreDeCod(Text1, conAri, "sfamia", "nomfamia", "codfamia")
    Else
        Text1.Text = ""
    End If
                
            
End Sub

Private Function DevSQL(Codfamia As Integer) As String
    DevSQL = "from sfichtecnicamp WHERE "
    If Codfamia < 0 Then
        DevSQL = DevSQL & "  marca = 3 "  'BLANQUETA
    Else
        DevSQL = DevSQL & " categoria = " & Codfamia  'CATEGo
    End If

End Function


Private Function CargarDatos(Codfamia As Integer) As Boolean
Dim SQL As String
Dim Aux As String
Dim i As Integer
    
    'Tengo que asegurarme que parar blanqueta SIEMRE EXISTE
    CargarDatos = False
    For i = 0 To 6
        Text2(i).Text = ""
        Text3(i).Text = ""
    Next i
    Set miRsAux = New ADODB.Recordset
    
    If Codfamia > 0 Then
        SQL = "Select count(*) from sartic where conjunto=1 and codfamia =" & Codfamia
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        If Not miRsAux.EOF Then i = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        If i = 0 Then
            MsgBox "No existe articulos venta asociado  a esta categoria", vbExclamation
            Exit Function
        End If
    End If
        
    SQL = "select count(*) " & DevSQL(Codfamia)
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not miRsAux.EOF Then i = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If i = 0 Then
        SQL = "No existen datos para la categoria: " & Text1.Text & " " & txtNombre.Text & "   ¿Crear?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            Aux = "INSERT INTO sfichtecnicamp (marca ,categoria ,linea ,intro)"
            Aux = Aux & " Select 0," & Codfamia & ",linea, intro from sfichtecnicamp where marca = 3"
            Conn.Execute Aux
        Else
            Exit Function
        End If
    End If
    
    
    SQL = "select * " & DevSQL(Codfamia) & " ORDER BY linea"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        i = miRsAux!Linea - 1
        Text2(i).Text = DBLet(miRsAux.Fields!intro, "T")
        Text3(i).Text = DBLet(miRsAux.Fields!Texto, "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Codfamia >= 0 Then
        Me.Label1.Caption = txtNombre.Text
    Else
        Me.Label1.Caption = "Blanqueta"
    End If
    Label1.Tag = Codfamia
    CargarDatos = True
End Function

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub
