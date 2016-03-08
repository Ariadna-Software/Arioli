VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacCopiarObservaciones2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmFacCopiarObservaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2325
      Index           =   6
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4200
      Width           =   7155
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   0
      Left            =   120
      MaxLength       =   80
      TabIndex        =   0
      Top             =   4200
      Width           =   7845
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   1
      Left            =   120
      MaxLength       =   80
      TabIndex        =   1
      Top             =   4470
      Width           =   7845
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   2
      Left            =   120
      MaxLength       =   80
      TabIndex        =   2
      Top             =   4740
      Width           =   7845
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   3
      Left            =   120
      MaxLength       =   80
      TabIndex        =   3
      Top             =   5010
      Width           =   7845
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   4
      Left            =   120
      MaxLength       =   80
      TabIndex        =   4
      Top             =   5280
      Width           =   7845
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Index           =   5
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5640
      Width           =   7845
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
   End
End
Attribute VB_Name = "frmFacCopiarObservaciones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IdCliente As Long
Public PackingList As Boolean
    'False: Normal
    'true:  Veremos las observaciones del packing list
    
'Llevara empipadas TODAS las observaciones
Public Event DatoSeleccionado(Datos As String)

Dim PrimeraVez As Boolean
Dim rs As ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        SQL = ""
        If Me.PackingList Then
            SQL = Text1(6).Text
            
        Else
            For IdCliente = 0 To 5
                SQL = SQL & Text1(IdCliente).Text & "|"
            Next IdCliente
        End If
        RaiseEvent DatoSeleccionado(SQL)
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Me.PackingList Then
            CargaDatos 0
            CargaDatos 1
        End If
        CargaDatos 2
        If TreeView1.Nodes.Count > 0 Then
            TreeView1.Nodes(1).Selected = True
            Set TreeView1.SelectedItem = TreeView1.Nodes(1)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
     PrimeraVez = True
    Me.Icon = frmppal.Icon
    Caption = "Copiar observaciones"
    Label1.Caption = "Observaciones"
    If Me.PackingList Then
        If vParamAplic.EsAVAB Then
            SQL = " packing list"
        Else
            SQL = " auxiliares"
        End If
    Else
        SQL = ""
    End If
    Text1(6).Width = Text1(0).Width
    Text1(6).visible = Me.PackingList
    
    Caption = Caption & SQL
    Label1.Caption = Label1.Caption & SQL
    'Iconos
    TreeView1.ImageList = frmppal.ImgListPpal
    'Text1(5).visible = vEmpresa.codempre = EmpresaAVAB
    Text1(5).visible = vParamAplic.EsAVAB
    
End Sub

'0-Pedidos
'1-Albaranes
'2-Facturas
Private Sub CargaDatos(vOpc As Byte)

Dim PadreInsertado As Boolean
Dim nodX

    On Error GoTo EC
    'Cargamos Facturas y albaranes
    Set rs = New ADODB.Recordset
    
    Select Case vOpc
    Case 0
        SQL = "Select numpedcl, fecpedcl FROM scaped WHERE "
    Case 1
        SQL = "Select codtipom, numalbar ,fechaalb FROM scaalb WHERE "
    

    Case 2
        If Me.PackingList Then
            SQL = "Select scafac.codtipom ,scafac.numfactu ,scafac.fecfactu ,packingobs FROM scafac WHERE "
            'SQL = SQL & "scafac.codtipom=scafac1.codtipom and scafac.fecfactu=scafac1.fecfactu and scafac.numfactu=scafac1.numfactu AND"
            SQL = SQL & " packingobs <>"""" AND "
        Else
            SQL = "Select scafac.codtipom ,scafac.numfactu ,scafac.fecfactu ,codtipoa ,numalbar FROM scafac,scafac1 WHERE "
            SQL = SQL & "scafac.codtipom=scafac1.codtipom and scafac.fecfactu=scafac1.fecfactu and scafac.numfactu=scafac1.numfactu AND "
        End If
    End Select
    
    
    
    SQL = SQL & "  codclien = " & IdCliente & " ORDER BY "
    Select Case vOpc
    Case 0
        SQL = SQL & "numpedcl desc"
    Case 1
        SQL = SQL & "codtipom,numalbar desc ,fechaalb desc "
    Case 2
        SQL = SQL & "codtipom ,numfactu desc ,fecfactu desc "
        If Not PackingList Then SQL = SQL & ",codtipoa ,numalbar"
    End Select
    
    'iconos
    '    .Buttons(9).Image = 5   'Ofertas Clientes
    '    .Buttons(10).Image = 6   'Pedidos Clientes
    '    .Buttons(11).Image = 7   'Albaranes Clientes
    '    .Buttons(12).Image = 8   'Hist. Albaranes Clientes (Facturas)

    
    rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PadreInsertado = False
    SQL = ""
    While Not rs.EOF
        If Not PadreInsertado Then
            'Insertamos el padre
            Set nodX = TreeView1.Nodes.Add(, , "N" & CStr(vOpc))
            Select Case vOpc
            Case 0
                nodX.Text = "PEDIDOS"
                nodX.Image = 6
            Case 1
                
                nodX.Text = "ALBARANES"
                nodX.Image = 7
            Case 2
                nodX.Text = "FACTURAS"
                nodX.Image = 8
            End Select
            PadreInsertado = True
        End If
    
        'Insertamos los nodos
        
        Set nodX = TreeView1.Nodes.Add("N" & CStr(vOpc), tvwChild)
        Select Case vOpc
        
        Case 0
            SQL = Format(rs!numpedcl, "000000")
            SQL = SQL & "  " & Format(rs!fecpedcl, "dd/mm/yyyy") & "  "
    
        
            nodX.Tag = "numpedcl = " & rs!numpedcl
            nodX.Image = 5
        
        Case 1
            SQL = rs!Codtipom & Format(rs!NumAlbar, "000000")
            SQL = SQL & "  " & Format(rs!FechaAlb, "dd/mm/yyyy") & "  "
    
        
            nodX.Tag = "numalbar = " & rs!NumAlbar & " AND codtipom = '" & rs!Codtipom & "'"
            nodX.Image = 5
        Case 2
            'fra
            'codtipom numfactu fecfactu codtipoa numalbar
            SQL = rs!Codtipom & Format(rs!NumFactu, "000000")
            SQL = SQL & "  " & Format(rs!FecFactu, "dd/mm/yyyy")
            nodX.Tag = "codtipom = '" & rs!Codtipom & "' AND numfactu = " & rs!NumFactu & " AND fecfactu = '" & Format(rs!FecFactu, FormatoFecha) & "' "
            If Not Me.PackingList Then
                SQL = SQL & "   (" & rs!NumAlbar & ")"
                nodX.Tag = nodX.Tag & " AND codtipoa = '" & rs!Codtipoa & "'  AND numalbar = " & rs!NumAlbar
            End If
            nodX.Image = 5
        End Select
        
        nodX.Text = SQL
        

        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
    Set rs = Nothing
End Sub



Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub



Private Sub limpiar()
    'Ya no la utilizo mas
    For IdCliente = 0 To 5
        Text1(IdCliente).Text = ""
    Next IdCliente
    
End Sub




Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node.Parent Is Nothing Then
        limpiar
        Exit Sub
    End If
    Set rs = New ADODB.Recordset

    
    If Right(Node.Parent.Key, 1) = "0" Then
        'PEDIDO
        SQL = "observa01,observa02,observa03,observa04,observa05,observa6 FROM scaped"
    ElseIf Right(Node.Parent.Key, 1) = "1" Then
        SQL = "observa01,observa02,observa03,observa04,observa05,observa6 FROM scaalb"
    Else
        If Me.PackingList Then
            SQL = "packingobs FROM scafac "
        Else
            SQL = "observa1,observa2,observa3,observa4,observa5,observa6 FROM scafac1"
        End If
    End If
    
    SQL = "Select " & SQL & " WHERE " & Node.Tag
    rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rs.EOF Then
        limpiar
    Else
        If PackingList Then
            'Del avab
            Text1(IdCliente).Text = DBLetMemo(rs!packingobs)
        Else
            For IdCliente = 0 To 5
                Text1(IdCliente).Text = DBLet(rs.Fields(IdCliente), "T")
            Next IdCliente
            If Not Text1(5).visible Then Text1(5).Text = ""
    
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub
