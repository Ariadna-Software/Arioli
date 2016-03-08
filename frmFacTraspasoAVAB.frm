VERSION 5.00
Begin VB.Form frmFacTraspasoAVAB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso facturas"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblDpto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha factura"
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
      Index           =   34
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1155
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1680
      Picture         =   "frmFacTraspasoAVAB.frx":0000
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmFacTraspasoAVAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim IndiceImg As Integer


Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Not HacerTrapaso Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Label1.Caption = "" 'indicador
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
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
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub





Private Function HacerTrapaso() As Boolean
Dim Col As Collection
Dim SQL As String
Dim i As Integer
Dim Trab As String
Dim CadenaDatosCabceraFra As String
Dim vPro As CProveedor

    HacerTrapaso = False
    
    If Me.txtFecha(0).Text = "" Then
        MsgBox "Ponga la fecha", vbExclamation
        Exit Function
    End If
    
    Label1.Caption = "Obteniendo facturas a traspasar"
    Label1.Refresh
    Set Col = New Collection
    Set miRsAux = New ADODB.Recordset
    SQL = "Select codtipom,numfactu,fecfactu from ariges" & EmprMorales & ".scafac where fecfactu>='2010-01-01'"
    
     SQL = SQL & " AND codtipom<>'FAZ'"
    SQL = SQL & " AND fecfactu <='" & Format(txtFecha(0).Text, FormatoFecha)
    SQL = SQL & "' AND codclien = 1 " 'AVAB
    'El meollo
    SQL = SQL & " AND not (codtipom,numfactu,fecfactu) IN (select codtipom,numfactu,fecfactu FROM straspasofra)"
    SQL = SQL & " ORDER BY 1,2,3"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
            SQL = miRsAux!codTipoM & "|" & miRsAux!NumFactu & "|" & miRsAux!FecFactu & "|"
            Col.Add SQL
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    SQL = "" 'Para poner a nothing los objetos
    If Col.Count > 0 Then
        SQL = DevuelveDesdeBD(conAri, "nomempre", "ariges" & EmprMorales & ".sparam", "codigo", "1") & "  "
    
        SQL = "Va a traspasar desde " & SQL & Col.Count & " facturas. ¿Continuar? "
        If MsgBox(SQL, vbExclamation + vbYesNo) = vbNo Then SQL = ""
    Else
        MsgBox "Ninguna factura pendiente de traspasar", vbExclamation
    End If
    If SQL = "" Then
        Set Col = Nothing
        Set miRsAux = Nothing
        Label1.Caption = ""
        Exit Function
    End If
    
    
    Trab = PonerTrabajadorConectado(SQL)
    If Trab = "" Then
        MsgBox "Imposible asignar trabajador conectado", vbExclamation
        Exit Function
    End If
    
    
    'nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codforpa"
    CadenaDatosCabceraFra = ""
    SQL = "Select nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprov1,codforpa from sprove where codprove=5"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "NO existe codprove=5", vbExclamation
    Else
        CadenaDatosCabceraFra = DBSet(miRsAux!nomprove, "T") & "," & DBSet(miRsAux!domprove, "T") & ","
        CadenaDatosCabceraFra = CadenaDatosCabceraFra & DBSet(miRsAux!codpobla, "T") & "," & DBSet(miRsAux!pobprove, "T") & ","
        CadenaDatosCabceraFra = CadenaDatosCabceraFra & DBSet(miRsAux!proprove, "T") & "," & DBSet(miRsAux!nifProve, "T") & ","
        CadenaDatosCabceraFra = CadenaDatosCabceraFra & DBSet(miRsAux!telprov1, "T") & "," & DBSet(miRsAux!codforpa, "N") & ","
    End If
    miRsAux.Close
    
    

    Set vPro = New CProveedor
    vPro.LeerDatos 5
    'Voy a fijar el banco por defecto
    SQL = DevuelveDesdeBD(conAri, "codmacta", "sbanpr", "codbanpr", vPro.BancoPropio)
    vPro.Observaciones = SQL 'Cta contable del banco
    
    If CadenaDatosCabceraFra = "" Then Exit Function
    CadenaDesdeOtroForm = ""
    '-------------- Vamos p'alla
    If BloqueoManual("trasfra", "1") Then
        
        For i = 1 To Col.Count
            SQL = Col.Item(i)
            Label1.Caption = i & " de " & Col.Count & "          Fra: " & SQL
            Label1.Refresh
            
            
            ConnConta.BeginTrans
            Conn.BeginTrans
            If TraspasarUnaFactura(SQL, Trab, CadenaDatosCabceraFra, vPro) Then
                Conn.CommitTrans
                ConnConta.CommitTrans
            Else
                Conn.RollbackTrans
                ConnConta.RollbackTrans
                SQL = "¿Continuar con el proceso? [NO]"
                If MsgBox(SQL, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then i = Col.Count
            End If
            Espera 0.1
            DoEvents
        
        Next i
        DesBloqueoManual "trasfra"
        Label1.Caption = ""
        Unload Me
        
    End If
End Function



Private Function TraspasarUnaFactura(StrFactura As String, Trabajador As String, DatosCabcera As String, vPro As CProveedor) As Boolean
Dim SQL As String
Dim Insert As String
Dim NumFacPRO As String
Dim R As ADODB.Recordset


    On Error GoTo eTraspasarUnFactura
        TraspasarUnaFactura = False
        
        
        

        
        'Primero insertaremos la factura
        '
        '
        SQL = "codprove,numfactu,fecfactu,fecrecep,nomprove,domprove,codpobla,pobprove,proprove,"
        SQL = SQL & "nifprove,telprove,codforpa,codtraba,brutofac,dtoppago,dtognral,impppago,impgnral,"
        SQL = SQL & "baseiva1,baseiva2,baseiva3,tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,"
        SQL = SQL & "porciva3,porcrec1,porcrec2,porcrec3,impoiva1,impoiva2,impoiva3,"
        SQL = SQL & "imporec1,imporec2,imporec3,totalfac,presupuesto"
        Insert = "INSERT INTO scafpc(" & SQL & ") SELECT "
        SQL = RecuperaValor(StrFactura, 2)
        NumRegElim = Val(SQL)
        'Veamos que letra de serie es esa factura. Para no tener k leer de la BD cada vez
        SQL = RecuperaValor(StrFactura, 1)
        If SQL <> RecuperaValor(CadenaDesdeOtroForm, 1) Then
            SQL = DevuelveDesdeBD(conAri, "letraser", "ariges" & EmprMorales & ".stipom", "codtipom", SQL, "T")
            CadenaDesdeOtroForm = RecuperaValor(StrFactura, 1) & "|" & SQL & "|"
        End If
        
        'NumFacPRO = RecuperaValor(StrFactura, 1) & Format(NumRegElim, "000000")
        NumFacPRO = RecuperaValor(CadenaDesdeOtroForm, 2) & Format(NumRegElim, "000000")
        SQL = "5,'" & NumFacPRO & "',fecfactu,fecfactu,"
        'SQL = SQL & "nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,codforpa," & Trabajador & ","
        SQL = SQL & DatosCabcera
        SQL = SQL & Trabajador & ","
        SQL = SQL & "brutofac,dtoppago,dtognral,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3,codigiv1,codigiv2,codigiv3,"
        SQL = SQL & "porciva1,porciva2,porciva3,porciva1re,porciva2re,porciva3re,imporiv1,imporiv2,imporiv3,"
        SQL = SQL & "imporiv1re,imporiv2re,imporiv3re,totalfac,0 "
        SQL = SQL & " FROM ariges" & EmprMorales & ".scafac WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        SQL = Insert & SQL
        Conn.Execute SQL
        
        
        'La de albaranes x fra
        SQL = "codprove,numfactu,fecfactu,numalbar,fechaalb,numpedpr,fecpedpr,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5"
        Insert = "INSERT INTO scafpa(" & SQL & ") SELECT "
        SQL = "5,'" & NumFacPRO & "',fecfactu,"
        SQL = SQL & "concat(codtipoa,numalbar), FechaAlb, numpedcl, fecpedcl," & Trabajador & "," & Trabajador & ", observa1, observa2, observa3, observa4, observa5"
        SQL = SQL & " FROM ariges" & EmprMorales & ".scafac1 WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        SQL = Insert & SQL
        Conn.Execute SQL
        
        'Las lineas
        
        'Primero compruebo k existe en salmac
        Set R = New ADODB.Recordset
        SQL = "Select codalmac,codartic,sum(cantidad) FROM ariges" & EmprMorales & ".slifac WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        SQL = SQL & " GROUP BY 1,2"
        R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not R.EOF
            SQL = DevuelveDesdeBD(conAri, "codartic", "salmac", "codalmac=" & R!codalmac & " AND codartic ", CStr(R!codartic), "T")
            If SQL = "" Then
                SQL = "insert INTO salmac(codartic,codalmac,ubialmac,canstock,statusin,preciomp,precioma,preciouc,preciost) "
                SQL = SQL & " VALUES ('" & R!codartic & "'," & R!codalmac & ",''," & TransformaComasPuntos(DBLet(R.Fields(2), "N"))
                SQL = SQL & ",0,0,0,0,0)"
                Conn.Execute SQL
            End If
            'Insertamos en smoval ?
            
            'Sig
            R.MoveNext
        Wend
        R.Close
       
        
        SQL = "codprove,numfactu,fecfactu,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,"
        SQL = SQL & "precioar , dtoline1, dtoline2, ImporteL"
        Insert = "INSERT INTO slifpc(" & SQL & ") SELECT "
        SQL = "5,'" & NumFacPRO & "',fecfactu,"
        SQL = SQL & "concat(codtipoa,numalbar),  numlinea , codalmac, codartic, NomArtic, ampliaci, Cantidad, precioar, dtoline1, dtoline2, ImporteL"
        SQL = SQL & " FROM ariges" & EmprMorales & ".slifac WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        SQL = Insert & SQL
        Conn.Execute SQL
        
        
                
                
                
        'Insert into smoval. Lo hace la funciona para que no de error
        SQL = "codartic,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,letraser,document,numlinea"
        Insert = "INSERT INTO smoval(" & SQL & ") SELECT "
        SQL = "codartic,codalmac,fecfactu,'" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & " " & Format(Now, "hh:mm:ss") & "'"
        SQL = SQL & ",1,'ALC',cantidad,importel,5,' ',concat(codtipoa,numalbar),  numlinea "
        SQL = SQL & " FROM ariges" & EmprMorales & ".slifac WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        SQL = Insert & SQL
        If Not EjecutaSQL(conAri, SQL) Then MsgBox "Error insertando en movimientos. El proceso continua", vbExclamation
        
            

        SQL = "insert into `straspasofra` (`codtipom`,`numfactu`,`fecfactu`,`codusu`,`fectraspaso`) "
        SQL = SQL & " values ('" & RecuperaValor(StrFactura, 1) & "'," & Format(NumRegElim)
        SQL = SQL & ",'" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'," & (vUsu.Codigo Mod 1000) & ",'" & Format(Now, FormatoFechaHora) & "')"
        Conn.Execute SQL
        
        
        
        
        
        
        
        'En tesoreria, en los pagos
        'vPro
        
        SQL = "Select totalfac FROM ariges" & EmprMorales & ".scafac WHERE codtipom='" & RecuperaValor(StrFactura, 1) & "' AND numfactu= " & Format(NumRegElim)
        SQL = SQL & " AND fecfactu = '" & Format(RecuperaValor(StrFactura, 3), FormatoFecha) & "'"
        R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Insert = CStr(R!TotalFac)
        R.Close
        
        SQL = Format(RecuperaValor(StrFactura, 3), FormatoFecha)
        '`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`codforpa`,`fecefect`
        SQL = DBSet(vPro.CuentaCble, "T") & ",'" & NumFacPRO & "','" & SQL & "',1," & vPro.ForPago & ",'" & SQL
        '`impefect`,`fecultpa`,`imppagad`,`ctabanc1`,`ctabanc2`,            Llevo la codmacta del banco
        SQL = SQL & "'," & TransformaComasPuntos(Insert) & ",NULL,NULL,'" & vPro.Observaciones & "',NULL"
        '`emitdocum`,`contdocu`,`text1csb`,`text2csb`,`entidad`,`oficina`,`CC`,`cuentaba`,`transfer`,`estacaja`,`referencia`)
        SQL = SQL & ",0,0,NULL,NULL,'"
        SQL = SQL & Format(vPro.Banco, "0000") & "','" & Format(vPro.Sucursal, "0000") & "','" & vPro.DigControl
        SQL = SQL & "','" & vPro.CuentaBan & "',NULL,0,NULL)"
        Insert = "INSERT INTO spagop (`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`codforpa`,`fecefect`,"
        Insert = Insert & "`impefect`,`fecultpa`,`imppagad`,`ctabanc1`,`ctabanc2`,"
        Insert = Insert & "`emitdocum`,`contdocu`,`text1csb`,`text2csb`,`entidad`,`oficina`,`CC`,`cuentaba`,`transfer`,`estacaja`,`referencia`) values ("
        SQL = Insert & SQL
        ConnConta.Execute SQL
        
        
         Set R = Nothing
        TraspasarUnaFactura = True
    Exit Function
eTraspasarUnFactura:
    MuestraError Err.Number, Err.Description & vbCrLf & SQL
     Set R = Nothing
End Function




