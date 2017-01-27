VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVallTrazaDesdeEntrada 
   Caption         =   "Datos desde entrada oliva"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   19403
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   5760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Peso"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   8
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Origen"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Albaran"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblIndi 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmVallTrazaDesdeEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
     If Text1.Text = "" Then Exit Sub
    Set miRsAux = New ADODB.Recordset
   
    If PonerDatosAlbaran Then
        CargaArbol
    Else
        LimpiarCampos
    End If
    Set miRsAux = Nothing
    lblIndi.Caption = ""
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    limpiar Me
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text1_LostFocus
End Sub

Private Sub Text1_LostFocus()
    
    
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        LimpiarCampos
    Else
        If Not IsNumeric(Text1.Text) Then
            MsgBox "Albaran incorrecto", vbExclamation
            LimpiarCampos
        Else
            PonerFocoBtn Me.Command1
            
        End If
    End If
    
End Sub

Private Function PonerDatosAlbaran() As Boolean
Dim Cad As String
Dim cP As cPartidas

    PonerDatosAlbaran = False
    Cad = "select neto , nomprove,numalbar,EntradaFinalizada from vallentradacamionlineas,vallentradacamion,sprove"
    Cad = Cad & " Where vallentradacamionlineas.entrada = vallentradacamion.entrada And sprove.codProve = vallentradacamion.codProve"
    Cad = Cad & " and EntradaFinalizada=1 and  numalbar=" & Text1.Text
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No existe albaran o no esta generado", vbExclamation
        
    Else
        Text2(0).Text = miRsAux!nomprove
        Text2(1).Text = miRsAux!Neto
        ListView1.ListItems.Clear
        miRsAux.Close
        
        Cad = "select loteproducido,artproducido from vallalmazaraprocesoalb,vallalmazaraproceso where "
        Cad = Cad & " vallalmazaraprocesoalb.ID = vallalmazaraproceso.ID and numalbar=" & Text1.Text
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "Error grave. No se encuentra porduccion almazara", vbExclamation
        Else
            conn.Execute "DELETE from tmptraza where codusu =" & vUsu.Codigo
            
            Set cP = New cPartidas
            If cP.LeerDesdeArticulo(miRsAux!artproducido, 1, miRsAux!loteproducido) Then
                PonerDatosAlbaran = True
                cP.TrazbilidadDesdeCompra lblIndi, True
            End If
            Set cP = Nothing
            
         End If
         miRsAux.Close
    End If
End Function



Private Sub CargaArbol()
Dim Cad As String
Dim UltDeposito As String
Dim UltNivel As Integer
Dim Nodo As Integer

    Cad = "select * from tmptraza where codusu =" & vUsu.Codigo & " order by contador"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    UltNivel = -1
    While Not miRsAux.EOF
        Cad = Space(miRsAux!nivle * 4) & miRsAux!idoperacion
        ListView1.ListItems.Add , , Cad
        If Mid(miRsAux!idoperacion, 1, 3) = "Cou" Then
            If UltNivel < miRsAux!nivle Then
                UltNivel = miRsAux!nivle
                Nodo = InStr(1, miRsAux!idoperacion, "(")
                Cad = Mid(miRsAux!idoperacion, Nodo + 1)
                Nodo = InStr(1, Cad, ")")
                Cad = Mid(Cad, 1, Nodo - 1)
                UltDeposito = CInt(Cad)
                Nodo = ListView1.ListItems.Count
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If UltDeposito <> "" Then
        Cad = "select * from olicoupage where codigo=" & UltDeposito
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            Cad = "  DEP: " & miRsAux!Deposito
            
            ListView1.ListItems(Nodo).Bold = True
            ListView1.ListItems(Nodo).Text = ListView1.ListItems(Nodo).Text & Cad
        End If
        miRsAux.Close
    End If
End Sub

Private Sub LimpiarCampos()
Dim I As Integer
    
    For I = 0 To Text2.Count - 1
        Text2(I).Text = ""
    Next
    ListView1.ListItems.Clear
End Sub
