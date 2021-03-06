VERSION 5.00
Begin VB.Form frmFacTrazabilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado trazabilidad(I)"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdListado 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Albaranes"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Facturas"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   8
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Origen facturas- albaranes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   195
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   465
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1515
      Picture         =   "frmFacTrazabilidad.frx":0000
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   4
      Left            =   3795
      Picture         =   "frmFacTrazabilidad.frx":008B
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   465
   End
End
Attribute VB_Name = "frmFacTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim SQL As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdListado_Click()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    CargaDatos
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If NumRegElim > 0 Then
    
        EstablecerParametros   'Sql y numregelim
    
        With frmImprimir
            .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
            .OtrosParametros = SQL
            .NumeroParametros = CInt(NumRegElim)
    
            .SoloImprimir = False
            .EnvioEMail = False
            .opcion = 2002
            .Titulo = "Trazabilidad"
            .NombreRPT = DevuelveNombreReport(38) ' "rLotes1.rp"
            .ConSubInforme = False
            .Show vbModal
        End With
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
End Sub

Private Sub frmF_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmF = New frmCal
    frmF.Fecha = Now
    NumRegElim = Index
    If txtCodigo(NumRegElim).Text <> "" Then frmF.Fecha = CDate(txtCodigo(NumRegElim).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    If SQL <> "" Then txtCodigo(NumRegElim).Text = SQL
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
     ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    If txtCodigo(Index).Text = "" Then Exit Sub
    If Index <= 1 Then
        PonerFormatoFecha txtCodigo(Index)
    Else
        If Not PonerFormatoEntero(txtCodigo(Index)) Then
            txtCodigo(Index).Text = ""
            PonerFoco txtCodigo(Index)
        End If
    End If
    
End Sub



Private Function CargaDatos() As Boolean
Dim Aux As String
Dim J As Integer
Dim C As String

    On Error GoTo ECargaDatos
    CargaDatos = False
    
    SQL = "DELETE from tmpinformes WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL
    NumRegElim = 0
    J = 0
    If Me.Check1(0).Value = 1 Then
        'Albaranes
        SQL = "select c.numalbar,c.fechaalb,nomclien,l.numlinea,linea,numlote,nomclien,codartic,c.codtipom"
        SQL = SQL & " from scaalb c,slialb l,slialblotes t where"
        SQL = SQL & " c.codtipom=l.codtipom and c.numalbar=l.numalbar and"
        SQL = SQL & " C.codTipoM = T.codTipoM And C.NumAlbar = T.NumAlbar And L.numlinea = T.numlinea"
        'WHERE
        If txtCodigo(0).Text <> "" Then SQL = SQL & " AND c.fechaalb>='" & Format(txtCodigo(0).Text, FormatoFecha) & "'"
        If txtCodigo(1).Text <> "" Then SQL = SQL & " AND c.fechaalb<='" & Format(txtCodigo(1).Text, FormatoFecha) & "'"
        If txtCodigo(2).Text <> "" Then SQL = SQL & " AND c.numalbar>=" & txtCodigo(2).Text
        If txtCodigo(3).Text <> "" Then SQL = SQL & " AND c.numalbar<=" & txtCodigo(3).Text

        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                                             'factura o albaran
        Aux = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,campo2,"  'para agrupar en el informe
        '       cliente   codartic  lote      numero    linea      fechafra alb
        Aux = Aux & "`nombre1`,`nombre2`,`nombre3`,`importe1`,`importe2`,`fecha1`)"
        Aux = Aux & " values (" & vUsu.Codigo & ","
        C = ""
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            
            SQL = miRsAux!codTipoM & Format(miRsAux!NumAlbar, "000000")
            If SQL <> C Then
                J = J + 1
                C = SQL
            End If
            SQL = NumRegElim & ",0," & J & ",'"
            SQL = SQL & DevNombreSQL(miRsAux!nomclien) & "','" & miRsAux!codArtic & "','" & DevNombreSQL(miRsAux!Numlote) & "',"
            SQL = SQL & miRsAux!NumAlbar & "," & miRsAux!numlinea & ",'" & Format(miRsAux!FechaAlb, FormatoFecha) & "')"
            SQL = Aux & SQL
            Conn.Execute SQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    
    'Facturas
    If Me.Check1(1).Value = 1 Then
        

        SQL = "select  c2.numalbar,c2.fechaalb,nomclien,l.numlinea,linea,numlote,nomclien,codartic,c.numfactu"
        SQL = SQL & " from scafac c,scafac1 c2,slifac l, slifaclotes t where"
        SQL = SQL & " c.codtipom=c2.codtipom and c.numfactu=c2.numfactu and c.fecfactu=c2.fecfactu and"
        SQL = SQL & " l.codtipom=c2.codtipom and l.numfactu=c2.numfactu and l.fecfactu=c2.fecfactu and l.codtipoa=c2.codtipoa and l.numalbar=c2.numalbar and"
        SQL = SQL & " L.codTipoM = T.codTipoM And L.NumFactu = T.NumFactu And L.FecFactu = T.FecFactu And L.codtipoa = T.codtipoa And L.NumAlbar = T.NumAlbar And L.numlinea = T.numlinea"
        
        
        'WHERE
        If txtCodigo(0).Text <> "" Then SQL = SQL & " AND c.FecFactu>='" & Format(txtCodigo(0).Text, FormatoFecha) & "'"
        If txtCodigo(1).Text <> "" Then SQL = SQL & " AND c.FecFactu<='" & Format(txtCodigo(1).Text, FormatoFecha) & "'"
        If txtCodigo(2).Text <> "" Then SQL = SQL & " AND c.numfactu>=" & txtCodigo(2).Text
        If txtCodigo(3).Text <> "" Then SQL = SQL & " AND c.numfactu<=" & txtCodigo(3).Text

        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                                             'factura o albaran   campo2, para poder agrupar comodamente
        Aux = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,campo2,"
        '       cliente   codartic  lote      numero    linea      fechafra alb
        Aux = Aux & "`nombre1`,`nombre2`,`nombre3`,`importe1`,`importe2`,`fecha1`)"
        Aux = Aux & " values (" & vUsu.Codigo & ","
        C = ""
        While Not miRsAux.EOF
        
            SQL = Format(miRsAux!FechaAlb, "yymmdd") & Format(miRsAux!NumFactu, "000000")
            If SQL <> C Then
                J = J + 1
                C = SQL
            End If
        
        
        
            NumRegElim = NumRegElim + 1
            SQL = NumRegElim & ",1," & J & ",'"
            SQL = SQL & DevNombreSQL(miRsAux!nomclien) & "','" & miRsAux!codArtic & "','" & DevNombreSQL(miRsAux!Numlote) & "',"
            SQL = SQL & miRsAux!NumFactu & "," & miRsAux!numlinea & ",'" & Format(miRsAux!FechaAlb, FormatoFecha) & "')"
            SQL = Aux & SQL
            Conn.Execute SQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    
    If NumRegElim > 0 Then
        CargaDatos = True
    Else
        MsgBox "Ningun dato con estos valores", vbExclamation
    End If
    Exit Function
ECargaDatos:
    MuestraError Err.Number, Err.Description
End Function



Private Sub EstablecerParametros()
    
        CadenaDesdeOtroForm = ""
        NumRegElim = 0
        
        SQL = ""
        If txtCodigo(0).Text <> "" Then SQL = SQL & " desde " & txtCodigo(0).Text
        If txtCodigo(1).Text <> "" Then SQL = SQL & "  hasta " & txtCodigo(1).Text
        If SQL <> "" Then CadenaDesdeOtroForm = "Fecha.  " & SQL
            
        SQL = ""
        If txtCodigo(2).Text <> "" Then SQL = SQL & " desde " & txtCodigo(2).Text
        If txtCodigo(3).Text <> "" Then SQL = SQL & " hasta " & txtCodigo(3).Text
        If SQL <> "" Then SQL = String(20, " ") & "Numero.  " & SQL
        SQL = CadenaDesdeOtroForm & SQL
        CadenaDesdeOtroForm = ""
        
        
        If Check1(0).Value = 0 Or Check1(1).Value = 0 Then
            'Solo ha seleccionado facturas o albaranes
            SQL = SQL & String(15, " ")
            If Check1(0).Value = 0 Then
                SQL = SQL & "FACTURAS"
            Else
                SQL = SQL & "ALBARANES"
            End If
            
        End If
        If SQL <> "" Then
            CadenaDesdeOtroForm = "pDHFecha= """ & Trim(SQL) & """|"
            NumRegElim = NumRegElim + 1
        End If

        SQL = "|" & CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
End Sub

