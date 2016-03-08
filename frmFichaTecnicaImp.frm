VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaTecnicaImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión ficha técnica"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin MSComctlLib.TreeView tv1 
         Height          =   5055
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8916
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   495
         Index           =   1
         Left            =   5280
         TabIndex        =   2
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   495
         Index           =   0
         Left            =   4080
         TabIndex        =   1
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Top             =   120
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmFichaTecnicaImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public vCodArtic As String

Private PrimeraVez As Boolean

Private Sub Command1_Click(Index As Integer)
Dim MatePrima As Byte
Dim ImpImag As Byte
Dim Aux As String

    If Index = 1 Then
        Unload Me
    Else
    
        'Cargar tabla de documentos que se visualizan
        ImpImag = VerImgAImprimir
    
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
    End If
End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        'Cargamos lw
        CargaLW
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    PrimeraVez = True
    
    Screen.MousePointer = vbHourglass
    Me.Label1.Caption = RecuperaValor(vCodArtic, 2)
    vCodArtic = RecuperaValor(vCodArtic, 1)
End Sub


Private Sub CargaLW()
Dim Cad As String
Dim R As ADODB.Recordset
Dim N As Node
Dim YaInsertado As String

    tv1.Nodes.Clear
    
        
    Cad = "SELECT `sarti1`.`codarti1`,nomartic, `sartic`.`codmarca`, `stipfamia`.`desctipfamia`,tipfamia"
    Cad = Cad & " FROM   `stipfamia` `stipfamia` inner JOIN  `sartic` `sartic` ON `stipfamia`.`tipfamia`=`sartic`.`tipartic`"
    Cad = Cad & " INNER JOIN `sarti1` `sarti1` ON `sartic`.`codartic`=`sarti1`.`codarti1` and sarti1.codartic = '" & vCodArtic & "'"
    Cad = Cad & " order by orden"
    YaInsertado = "|"
    Set R = New ADODB.Recordset
    R.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not R.EOF
        
        Cad = "|" & CStr(R!tipfamia) & "|"
            
        If InStr(1, YaInsertado, Cad) = 0 Then
            Set N = tv1.Nodes.Add(, , "T" & CStr(R!tipfamia))
            N.Text = R!desctipfamia
            YaInsertado = YaInsertado & CStr(R!tipfamia) & "|"
        End If
        
        
        'insertamos el nodo
        If Val(R!tipfamia) = 0 Then
            'ACEITE
            Set N = tv1.Nodes.Add("T" & CStr(R!tipfamia), tvwChild, "A" & R!codArti1)
            N.Text = R!NomArtic
            N.Checked = True
            N.Parent.Checked = True
            N.EnsureVisible
        Else
            'insertamos el nodo
            Cad = DevuelveDesdeBD(conAri, "count(*)", "sfichtecdocs", "codartic", R!codArti1, "T")
            If Cad = "" Then Cad = "0"
            If Val(Cad) > 0 Then
                Set N = tv1.Nodes.Add("T" & CStr(R!tipfamia), tvwChild, "A" & CStr(R!codArti1))
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
Private Function VerImgAImprimir() As Byte

Dim Orden As Integer
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim N As Node
Dim Nod As Node
    Conn.Execute "DELETE from tmprutas where codusu = " & vUsu.Codigo
    Orden = 0
    VerImgAImprimir = 0
    
    'Ahora para cada check ire mirando...
    Set Rs = New ADODB.Recordset
        
    Set N = tv1.Nodes(1).Next  'El siguiente NODO al del aceite Materia prima
    
    While Not N Is Nothing
    
        
        'Vamos a los hijos del NODO
        Set Nod = N.Child
        While Not Nod Is Nothing
                If Nod.Checked Then
                    'ESTE ARTICULO de lineas CARGO SUS IMAGENES
                    Cad = "Select codigo from sfichtecdocs where codartic = '" & Mid(Nod.Key, 2) & "' ORDER BY orden"
                    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Cad = ""
                    While Not Rs.EOF
                        'insert into `tmprutas` (`codusu`,``,`codigo`,idruta
                        Orden = Orden + 1
                        Cad = Cad & ", (" & vUsu.Codigo & "," & Rs!Codigo & "," & Orden & ")"
                        Rs.MoveNext
                    Wend
                    Rs.Close
                    
                    If Cad <> "" Then
                        Cad = Mid(Cad, 2)
                        'insert into `tmprutas` (`codusu`,``,`codigo`,idruta
                        Cad = "insert into `tmprutas` (`codusu`,`codigo`,`idruta`) VALUE " & Cad
                        Conn.Execute Cad
                    End If
                End If
        
            Set Nod = Nod.Next
        Wend
        Set N = N.Next
    Wend
    If Orden > 0 Then VerImgAImprimir = 1
    Set Rs = Nothing
End Function
