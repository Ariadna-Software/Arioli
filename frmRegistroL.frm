VERSION 5.00
Begin VB.Form frmRegistroL2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico registros"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmRegistroL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpr 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Documento firmado"
         Height          =   195
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmRegistroL.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "Firma 1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "Firma 2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmRegistroL.frx":0596
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmRegistroL.frx":0698
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3195
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmRegistroL.frx":079A
      Top             =   2280
      Width           =   7860
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   5760
      Width           =   975
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
      TabIndex        =   9
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label lblDescCampo 
      Caption         =   "Texto"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   19
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmRegistroL2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public EsMantenimientoPreventivo As Boolean

Public Linea As Integer
Public Nregistro2 As String   '
Public InsMod As Boolean  'Que se pueda insertar / modificar

Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim SQL As String



Private Function RealizarAccionReg() As Boolean
Dim NL As Integer
Dim ActualizoFech As Boolean
    RealizarAccionReg = False
    
    NL = 0
    SQL = ""
    If Text1(0).Text = "" Then SQL = SQL & "- Fecha" & vbCrLf
    If Text1(1).Text = "" Then SQL = SQL & "- Trabajador 1" & vbCrLf
    If SQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & SQL, vbExclamation
        Exit Function
    End If
    
    If Linea = 0 Then
        'NUEVO

        SQL = "idregistro=" & Nregistro2
        SQL = SugerirCodigoSiguienteStr("sregistrosL", "secuencial", SQL)

        NL = CInt(SQL)
        
        'insert into `sregistrosl` (`idRegistro`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`,`codtraba2`)
        SQL = Nregistro2 & "," & SQL & ",'" & Format(Text1(0).Text, FormatoFecha) & "',"
        SQL = SQL & DBSet(Text1(3).Text, "T", "S") & ","
        SQL = SQL & Abs(Check1.Value) & ","
        SQL = SQL & DBSet(Text1(1).Text, "N") & ","
        SQL = SQL & DBSet(Text1(2).Text, "N", "S") & ")"
        SQL = "insert into `sregistrosl` (`idRegistro`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`,`codtraba2`) VALUES (" & SQL
        
    Else
        'MODIFICAR
        SQL = "UPDATE sregistrosl SET fecha=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & ", firmado = " & Abs(Me.Check1.Value)
        SQL = SQL & ", texto = " & DBSet(Text1(3).Text, "T", "S")
        SQL = SQL & ", codtraba1 = " & DBSet(Text1(1).Text, "N", "N")
        SQL = SQL & ", codtraba2 = " & DBSet(Text1(2).Text, "S", "S")
        SQL = SQL & " WHERE idregistro=" & Nregistro2 & " AND secuencial = " & Linea
    End If
    If Not EjecutaSQL(conAri, SQL, True) Then Exit Function
    
    
    'Updatea tb el campo
    'Veo si es la ultima fecha
    ActualizoFech = False
    If Linea = 0 Then
        ActualizoFech = True
        
    Else
        SQL = DevuelveDesdeBD(conAri, "UltimoRealizado", "sregistros", "idRegistro", CStr(Nregistro2))
        If SQL = "" Then
            'Actualio la fecha maxima
            ActualizoFech = True
        Else
            If CDate(SQL) < CDate(Text1(0).Text) Then ActualizoFech = True
        End If
    
    End If
    If ActualizoFech Then
        SQL = "UPDATE sregistros SET UltimoRealizado=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & " WHERE idregistro=" & Nregistro2
        EjecutaSQL conAri, SQL, True
    End If
    
    
    If NL > 0 And Linea = 0 Then
        'Nuevo. Por si queire imprimir
        Espera 0.5
        ImprimirLinea NL
    End If
    
    RealizarAccionReg = True
End Function





Private Function RealizarAccionMantePrev() As Boolean
Dim NL As Integer
Dim ActualizoFech As Boolean
    RealizarAccionMantePrev = False
    
    NL = 0
    SQL = ""
    If Text1(0).Text = "" Then SQL = SQL & "- Fecha" & vbCrLf
    If Text1(1).Text = "" Then SQL = SQL & "- Trabajador 1" & vbCrLf
    If SQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & SQL, vbExclamation
        Exit Function
    End If
    
    If Linea = 0 Then
        'NUEVO

        SQL = "idregistro=" & DBSet(Nregistro2, "T")
        SQL = SugerirCodigoSiguienteStr("sregistrosL", "secuencial", SQL)

        NL = CInt(SQL)
        
        'insert into `sregistrosl` (`idRegistro`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`,`codtraba2`)
        SQL = Nregistro2 & "," & SQL & ",'" & Format(Text1(0).Text, FormatoFecha) & "',"
        SQL = SQL & DBSet(Text1(3).Text, "T", "S") & ","
        SQL = SQL & Abs(Check1.Value) & ","
        SQL = SQL & DBSet(Text1(1).Text, "N") & ","
        SQL = SQL & DBSet(Text1(2).Text, "N", "S") & ")"
        SQL = "insert into `sregistrosl` (`idRegistro`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`) VALUES (" & SQL
        
    Else
        'MODIFICAR
        SQL = "UPDATE sregistrosl SET fecha=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & ", firmado = " & Abs(Me.Check1.Value)
        SQL = SQL & ", texto = " & DBSet(Text1(3).Text, "T", "S")
        SQL = SQL & ", codtraba1 = " & DBSet(Text1(1).Text, "N", "N")
        SQL = SQL & ", codtraba2 = " & DBSet(Text1(2).Text, "S", "S")
        SQL = SQL & " WHERE idregistro=" & DBSet(Nregistro2, "T") & " AND secuencial = " & Linea
    End If
    If Not EjecutaSQL(conAri, SQL, True) Then Exit Function
    
    
    'Updatea tb el campo
    'Veo si es la ultima fecha
    ActualizoFech = False
    If Linea = 0 Then
        ActualizoFech = True
        
    Else
        SQL = DevuelveDesdeBD(conAri, "UltimoRealizado", "sregistros", "codigo", CStr(Nregistro2))
        If SQL = "" Then
            'Actualio la fecha maxima
            ActualizoFech = True
        Else
            If CDate(SQL) < CDate(Text1(0).Text) Then ActualizoFech = True
        End If
    
    End If
    If ActualizoFech Then
        SQL = "UPDATE sregistros SET UltimoRealizado=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & " WHERE idregistro=" & Nregistro2
        EjecutaSQL conAri, SQL, True
    End If
    
    
    If NL > 0 And Linea = 0 Then
        'Nuevo. Por si queire imprimir
        Espera 0.5
        ImprimirLinea NL
    End If
    
    RealizarAccionMantePrev = True
End Function







Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub ImprimirLinea(Lin As Integer)

        
        If EsMantenimientoPreventivo Then
       
       
       
       
       Else
       
             SQL = "{sregistros.idRegistro}=" & Nregistro2 & " AND {sregistrosl.secuencial} = " & Lin
                  
             CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "ReportPropio", "sregistros", "idRegistro", CStr(Nregistro2))
             If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "MorRegistro.rpt"
            
             
             LlamaImprimirGral SQL, "", 0, CadenaDesdeOtroForm, "Registro: " & Nregistro2 & " - " & Text1(0).Text
        
        End If
        
        
End Sub

Private Sub cmdImpr_Click()
    ImprimirLinea Linea
End Sub

Private Sub Command1_Click(Index As Integer)
Dim B As Boolean
    If Index = 1 Then
        CadenaDesdeOtroForm = ""  'Garantizo que esta variable esta vacia
        Unload Me
    Else
        If EsMantenimientoPreventivo Then
            B = RealizarAccionPreventivo
        Else
            B = RealizarAccionReg
        End If
        If B Then
            CadenaDesdeOtroForm = "OK"   'para que refresque datos
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    
    
    
    limpiar Me
    
    lblDescCampo(3).visible = Not EsMantenimientoPreventivo
    imgArticulo(2).visible = Not EsMantenimientoPreventivo
    Text1(2).visible = Not EsMantenimientoPreventivo
    txtDesc(2).visible = Not EsMantenimientoPreventivo
    
    Me.cmdImpr.visible = Linea > 0 And Not InsMod
    If Linea = 0 Then
        Me.Label1.Caption = "NUEVO"
        InsMod = True
        Text1(0).Text = Mid(CadenaDesdeOtroForm, 1, 10)
        
        Text1(3).Text = Mid(CadenaDesdeOtroForm, 11)
        CadenaDesdeOtroForm = ""
        'limpiar Me
    Else
        If InsMod Then
            Me.Label1.Caption = "MODIFICAR"
            
        Else
            Label1.Caption = "VER"
        End If
        Me.Frame1.Enabled = InsMod
        Text1(3).Locked = Not InsMod
        Set miRsAux = New ADODB.Recordset
        
        If EsMantenimientoPreventivo Then
            'MATENIMIENTO PREENTIVO
            SQL = "Select * from sregistrosmanprevl where codigo=" & DBSet(Nregistro2, "T") & " AND "
            SQL = SQL & "secuencial = " & Linea
            
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "No se ha encotrado el registro", vbExclamation
                Me.Command1(0).Enabled = False
                
            Else
                Me.Command1(0).Enabled = InsMod
                Me.Text1(0).Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
                SQL = ""
                If Not IsNull(miRsAux!Codtraba1) Then
                    'DATOS TRABAJADOR
                    Text1(1).Text = miRsAux!Codtraba1
                    SQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(1).Text)
                    Me.txtDesc(1).Text = SQL
                End If
                
        
                Me.Check1.Value = DBLet(miRsAux!firmado, "N")
                Text1(3).Text = DBLetMemo(miRsAux!Texto)
            End If
        
        Else
            SQL = "Select * from sregistrosl where idregistro=" & Nregistro2 & " AND "
            SQL = SQL & "secuencial = " & Linea
            
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "No se ha encotrado el registro", vbExclamation
                Me.Command1(0).Enabled = False
                
            Else
                Me.Command1(0).Enabled = InsMod
                Me.Text1(0).Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
                SQL = ""
                If Not IsNull(miRsAux!Codtraba1) Then
                    'DATOS TRABAJADOR
                    Text1(1).Text = miRsAux!Codtraba1
                    SQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(1).Text)
                    Me.txtDesc(1).Text = SQL
                End If
                
                If Not IsNull(miRsAux!Codtraba2) Then
                    'DATOS TRABAJADOR
                    Text1(2).Text = miRsAux!Codtraba2
                    SQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(2).Text)
                    Me.txtDesc(2).Text = SQL
                End If
                Me.Check1.Value = DBLet(miRsAux!firmado, "N")
                Text1(3).Text = DBLetMemo(miRsAux!Texto)
                
            End If
            miRsAux.Close
        End If
        Set miRsAux = Nothing
    End If
    

End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
    
End Sub

Private Sub imgArticulo_Click(Index As Integer)
    SQL = ""
    Set frmT = New frmAdmTrabajadores
    frmT.DatosADevolverBusqueda = "0|1|"
    frmT.Show vbModal
    Set frmT = Nothing
    If SQL <> "" Then
        Text1(Index).Text = RecuperaValor(SQL, 1)
        txtDesc(Index).Text = RecuperaValor(SQL, 2)
        PonerFoco Text1(Index + 1)
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   SQL = ""
   Set frmC = New frmCal
   frmC.Fecha = Now
   If Text1(Index).Text <> "" Then
        If IsDate(Text1(Index).Text) Then frmC.Fecha = CDate(Text1(Index).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
   If SQL <> "" Then
        Text1(Index).Text = Format(SQL, "dd/mm/yyyy")
        PonerFoco Text1(1)
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index = 3 Then Exit Sub
    ConseguirFoco Text1(Index), 3
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 3 Then Exit Sub
    KEYpressGnral KeyAscii, 3, False
End Sub
 



Private Sub Text1_LostFocus(Index As Integer)
  
        
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 1, 2
            SQL = ""
            If Text1(Index).Text = "" Then
            
            Else
                If Not PonerFormatoEntero(Text1(Index)) Then
                    If Text1(Index).Text <> "" Then Text1(Index).Text = ""
                Else
                    SQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(Index).Text)
                    If SQL = "" Then
                        MsgBox "No existe el trabajador: " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                    End If
                End If
                If Text1(Index).Text = "" Then PonerFoco Text1(Index)
            End If
            
            Me.txtDesc(Index).Text = SQL
        
    End Select
End Sub




Private Function RealizarAccionPreventivo() As Boolean
Dim NL As Integer
Dim ActualizoFech As Boolean
    RealizarAccionPreventivo = False
    
    NL = 0
    SQL = ""
    If Text1(0).Text = "" Then SQL = SQL & "- Fecha" & vbCrLf
    If Text1(1).Text = "" Then SQL = SQL & "- Trabajador 1" & vbCrLf
    If SQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & SQL, vbExclamation
        Exit Function
    End If
    
    If Linea = 0 Then
        'NUEVO

        SQL = "codigo=" & DBSet(Nregistro2, "T")
        SQL = SugerirCodigoSiguienteStr("sregistrosmanprevl", "secuencial", SQL)

        NL = CInt(SQL)
        
        'insert into `sregistrosl` (`idRegistro`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`,`codtraba2`)
        SQL = DBSet(Nregistro2, "T") & "," & SQL & ",'" & Format(Text1(0).Text, FormatoFecha) & "',"
        SQL = SQL & DBSet(Text1(3).Text, "T", "S") & ","
        SQL = SQL & Abs(Check1.Value) & ","
        SQL = SQL & DBSet(Text1(1).Text, "N") & ")"
        SQL = "insert into `sregistrosmanprevl` (`codigo`,`secuencial`,`Fecha`,`texto`,`Firmado`,`codtraba1`) VALUES (" & SQL
        
    Else
        'MODIFICAR
        SQL = "UPDATE sregistrosmanprevl SET fecha=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & ", firmado = " & Abs(Me.Check1.Value)
        SQL = SQL & ", texto = " & DBSet(Text1(3).Text, "T", "S")
        SQL = SQL & ", codtraba1 = " & DBSet(Text1(1).Text, "N", "N")
        SQL = SQL & " WHERE codigo=" & DBSet(Nregistro2, "T") & " AND secuencial = " & Linea
    End If
    If Not EjecutaSQL(conAri, SQL, True) Then Exit Function
    
    
    'Updatea tb el campo
    'Veo si es la ultima fecha
    ActualizoFech = False
    If Linea = 0 Then
        ActualizoFech = True
        
    Else
        SQL = DevuelveDesdeBD(conAri, "UltimoRealizado", "sregistrosmanprev", "codigo", DBSet(Nregistro2, "T"))
        If SQL = "" Then
            'Actualio la fecha maxima
            ActualizoFech = True
        Else
            If CDate(SQL) < CDate(Text1(0).Text) Then ActualizoFech = True
        End If
    
    End If
    If ActualizoFech Then
        SQL = "UPDATE sregistrosmanprev SET UltimoRealizado=" & DBSet(Text1(0).Text, "F")
        SQL = SQL & " WHERE codigo=" & DBSet(Nregistro2, "T")
        EjecutaSQL conAri, SQL, True
    End If
    
    
'    If NL > 0 And Linea = 0 Then
'        'Nuevo. Por si queire imprimir
'        Espera 0.5
'        ImprimirLinea NL
'    End If
    
    RealizarAccionPreventivo = True
End Function

