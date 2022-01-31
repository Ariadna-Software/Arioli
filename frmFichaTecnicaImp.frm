VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFichaTecnicaImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión ficha técnica"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   240
      Top             =   7200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
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
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin MSComctlLib.TreeView tv1 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9340
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   5760
         TabIndex        =   2
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4320
         TabIndex        =   1
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label lblCarga 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmFichaTecnicaImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CarpetaIMG = "ImgFicFT"   'si cambio aqui, cambiar tambien en impresion ficha tencina producto venta

Public vCodArtic As String


Private PrimeraVez As Boolean

Private Sub Command1_Click(Index As Integer)
Dim MatePrima As Byte
Dim ImpImag As Byte
Dim Aux As String
Dim ListaPdfs As String
Dim ImpreDirPDF As Boolean
    If Index = 1 Then
        Unload Me
    Else
    
        'Cargar tabla de documentos que se visualizan
        ImpImag = VerImgAImprimir(ListaPdfs)
    
        If ListaPdfs <> "" Then
            'Ok. Lleva fichas tecnicas en formato PDF
            'Veremos si las ha seleccionado
            Aux = "Hay fichas técnicas en formato PDF." & vbCrLf & " ¿Desea imprimirlos directamente?" & vbCrLf & vbCrLf
            Aux = Aux & "       -SI     : Imprime por la impresora por defecto " & vbCrLf
            Aux = Aux & "       -No  : Ver en visor PDFs" & vbCrLf
            MatePrima = MsgBox(Aux, vbQuestion + vbYesNoCancel + vbDefaultButton3)
            If MatePrima = vbCancel Then Exit Sub
            ImpreDirPDF = IIf(MatePrima = vbYes, True, False)
            
        
        End If
    
    
    
        'Si mostramos el aceite
        MatePrima = "0"
        If Mid(tv1.Nodes(1).Key, 2, 1) = "0" Then    'si es aceite
            If tv1.Nodes(1).Children > 0 Then   'si tiene hijos
                If tv1.Nodes(1).Child.Checked Then MatePrima = "1"  'si esta checkeado
            End If
        End If
        With frmImprimir
            .FormulaSeleccion = "{sartic.codartic}=""" & Me.vCodArtic & """"
            .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|pMP=" & MatePrima & "|vcodusu=" & vUsu.Codigo & "|MostrarImg=" & ImpImag & "|"
            .NumeroParametros = 4
            'select , from ariges9.scryst
            Aux = DevuelveDesdeBD(conAri, "codigrev", "scryst", "codcryst", "37")
            If Aux <> "" Then
                If IsNumeric(Aux) Then
                    .OtrosParametros = .OtrosParametros & "pCodigoRev=""" & Aux & """|"
                    .NumeroParametros = .NumeroParametros + 1
                End If
            End If
            Aux = DevuelveDesdeBD(conAri, "codigiso", "scryst", "codcryst", "37")
            If Aux <> "" Then
                If IsDate(Aux) Then
                    .OtrosParametros = .OtrosParametros & "pCodigoISO=""" & Aux & """|"
                    .NumeroParametros = .NumeroParametros + 1
                End If
            End If

            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 95
            .Show vbModal
        End With
        
        Me.lblCarga.Caption = "PDFs"
        Me.lblCarga.Refresh
        While ListaPdfs <> ""
            NumRegElim = InStr(1, ListaPdfs, "|")
            
            If NumRegElim = 0 Then
                ListaPdfs = ""
            Else
                Aux = Mid(ListaPdfs, 1, NumRegElim - 1)
                ListaPdfs = Mid(ListaPdfs, NumRegElim + 1)
                Aux = App.Path & "\" & CarpetaIMG & "\" & Aux
                
                Me.lblCarga.Caption = "...." & Right(Aux, 12)
                lblCarga.Refresh
                Espera 0.5
                If ImpreDirPDF Then
                    lanzaImpresionShellDirecta Me.hwnd, Aux
                Else
                    LanzaVisorMimeDocumento Me.hwnd, Aux
                End If
            End If
        Wend
     lblCarga.Caption = ""
    End If
   
End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        'Cargamos lw
        lblCarga.Caption = "Cargando desde BD"
        lblCarga.Refresh
        CargaLW
        
        
        CargaArchivosPDfs
        
        lblCarga.Caption = ""
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
    
    Screen.MousePointer = vbHourglass
    lblCarga.Caption = ""
    Label1.Caption = RecuperaValor(vCodArtic, 2)
    vCodArtic = RecuperaValor(vCodArtic, 1)
End Sub


Private Sub CargaLW()
Dim cad As String
Dim R As ADODB.Recordset
Dim N As Node
Dim YaInsertado As String

    tv1.Nodes.Clear
    
        
    cad = "SELECT `sarti1`.`codarti1`,nomartic, `sartic`.`codmarca`, `stipfamia`.`desctipfamia`,tipfamia"
    cad = cad & " FROM   `stipfamia` `stipfamia` inner JOIN  `sartic` `sartic` ON `stipfamia`.`tipfamia`=`sartic`.`tipartic`"
    cad = cad & " INNER JOIN `sarti1` `sarti1` ON `sartic`.`codartic`=`sarti1`.`codarti1` and sarti1.codartic = '" & vCodArtic & "'"
    cad = cad & " order by orden"
    YaInsertado = "|"
    Set R = New ADODB.Recordset
    R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not R.EOF
        
        cad = "|" & CStr(R!tipfamia) & "|"
            
        If InStr(1, YaInsertado, cad) = 0 Then
            Set N = tv1.Nodes.Add(, , "T" & CStr(R!tipfamia))
            N.Text = R!desctipfamia
            YaInsertado = YaInsertado & CStr(R!tipfamia) & "|"
        End If
        
        
        'insertamos el nodo
        If Val(R!tipfamia) = 0 Then
            'ACEITE
            Set N = tv1.Nodes.Add("T" & CStr(R!tipfamia), tvwChild, "A" & R!codarti1)
            N.Text = R!NomArtic
            N.Checked = True
            N.Parent.Checked = True
            N.EnsureVisible
        Else
            'insertamos el nodo
            cad = DevuelveDesdeBD(conAri, "count(*)", "sfichtecdocs", "codartic", R!codarti1, "T")
            If cad = "" Then cad = "0"
            If Val(cad) > 0 Then
                Set N = tv1.Nodes.Add("T" & CStr(R!tipfamia), tvwChild, "A" & CStr(R!codarti1))
                N.Text = R!NomArtic
                N.EnsureVisible
            End If
         
                
        End If
        
        
    
        R.MoveNext
    Wend
    R.Close
    
    
   
    
End Sub




Private Sub tv1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
    If Not Node.Parent Is Nothing Then Exit Sub
    Set N = Node.Child
    While Not N Is Nothing
        N.Checked = Node.Checked
        Set N = N.Next
    Wend
    
End Sub


'Devolvera
'   0 si NO imprime images
'   1 SI
Private Function VerImgAImprimir(ByRef ListaDocsPDF As String) As Byte

Dim Orden As Integer
Dim RS As ADODB.Recordset
Dim cad As String
Dim N As Node
Dim Nod As Node
    
    
    ListaDocsPDF = ""
    
    conn.Execute "DELETE from tmprutas where codusu = " & vUsu.Codigo
    Orden = 0
    VerImgAImprimir = 0
    
    'Ahora para cada check ire mirando...
    Set RS = New ADODB.Recordset
        
    Set N = tv1.Nodes(1).Next  'El siguiente NODO al del aceite Materia prima
    
    While Not N Is Nothing
    
        
        'Vamos a los hijos del NODO
        Set Nod = N.Child
        While Not Nod Is Nothing
                If Nod.Checked Then
                    'ESTE ARTICULO de lineas CARGO SUS IMAGENES
                    cad = "Select codigo,esPDF from sfichtecdocs where codartic = '" & Mid(Nod.Key, 2) & "' ORDER BY orden"
                    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    cad = ""
                    While Not RS.EOF
                        'insert into `tmprutas` (`codusu`,``,`codigo`,idruta
                        'Si es PDF lo concateno
                        If DBLet(RS!EsPDF, "N") = 1 Then ListaDocsPDF = ListaDocsPDF & Format(DBLet(RS!Codigo, "N"), "000000") & ".pdf|"
                            
                        Orden = Orden + 1
                        cad = cad & ", (" & vUsu.Codigo & "," & RS!Codigo & "," & Orden & ")"
                        RS.MoveNext
                        
                    Wend
                    RS.Close
                    
                    If cad <> "" Then
                        cad = Mid(cad, 2)
                        'insert into `tmprutas` (`codusu`,``,`codigo`,idruta
                        cad = "insert into `tmprutas` (`codusu`,`codigo`,`idruta`) VALUE " & cad
                        conn.Execute cad
                    End If
                End If
        
            Set Nod = Nod.Next
        Wend
        Set N = N.Next
    Wend
    If Orden > 0 Then VerImgAImprimir = 1
    Set RS = Nothing
End Function


Private Sub CargaArchivosPDfs()
Dim C As String

    
    On Error GoTo eCargaArchivosPDfs
  
    C = "Select codarti1 from sarti1 where codartic='" & vCodArtic & "'  and espdf=1 "
    C = " Select * from sfichtecdocs where codartic IN (" & C & ") ORDER BY orden"
    
    
    Me.lblCarga.Caption = "Leyendo desde BD "
    Me.lblCarga.Refresh
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = C
    Adodc1.Refresh
    
    
    While Not Adodc1.Recordset.EOF
        
        lblCarga.Caption = "Leyendo desde BD " & DBLet(Adodc1.Recordset!Codigo, "N") & "->       " & Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        lblCarga.Refresh
        C = App.Path & "\" & CarpetaIMG & "\" & Format(DBLet(Adodc1.Recordset!Codigo, "N"), "000000") & ".pdf"   'cuidado al cambar. En VerImgAImprimir hay que cambiarlo
        If Not LeerBinary(Adodc1.Recordset!campo, C) Then Err.Raise 513, , "Copiando fichero desde BD. CargaArchivosPDfs"
        
        
        
        Adodc1.Recordset.MoveNext
    Wend

    
    

    Exit Sub
eCargaArchivosPDfs:
    MuestraError Err.Number, , Err.Description
    Me.Command1(0).Enabled = False
End Sub
