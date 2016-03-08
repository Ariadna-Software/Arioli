VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRevisonInven 
   Caption         =   "Revision Inventario"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "30/10/2009"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Corregir STOCK"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      ToolTipText     =   "Check"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Corregir DFI"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "codartic"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nomartic"
         Object.Width           =   5293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Antes mov"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Inventari"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "MovPosterior"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "canstoc"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "st teori"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "DFI"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "DFI cal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Difst"
         Object.Width           =   899
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Almacen"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   7680
      Width           =   4335
   End
End
Attribute VB_Name = "frmRevisonInven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PrimV As Boolean
'Dim miA As miArticulo



'Private Sub TrataEstructura(A_cadena As Boolean, ByRef Cadena As String, ByRef cArt As miArticulo)
'
'    If A_cadena Then
'        Cadena = cArt.cantidadactual & "|" & cArt.cantidainven & "|" & cArt.codartic & "|" & cArt.DFI & "|" & cArt.MovPosterior & "|"
'    Else
'        'Al obejto
'        cArt.cantidadactual = RecuperaValor(Cadena, 1)
'        cArt.cantidainven = RecuperaValor(Cadena, 2)
'        cArt.codartic = RecuperaValor(Cadena, 3)
'        cArt.DFI = RecuperaValor(Cadena, 4)
'        cArt.MovPosterior = RecuperaValor(Cadena, 5)
'    End If
'End Sub

Private Sub HacerACciones()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Codfamia As Integer
Dim Posteriores As Currency
Dim Actual As Currency
Dim MovimientodDFI As Currency
Dim Suma As Currency
Dim AnteriorSmoval As Currency
Dim TieneMovDFI As Boolean
Dim Rojos As Integer
Dim Azul As Integer

Dim It As ListItem


    ListView1.ListItems.Clear
SQL = "select salmac.*,sartic.codfamia,nomartic,sfamia.nomfamia from salmac,sartic,sfamia where "
SQL = SQL & " salmac.codartic=sartic.codartic and sfamia.codfamia=sartic.codfamia AND"
SQL = SQL & " fechainv='" & Format(Text2.Text, FormatoFecha) & "' and codalmac= " & Text1.Text & "  ORDER by sartic.codfamia,nomartic"

Set RS = New ADODB.Recordset

    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Codfamia = -1
    Rojos = 0
    Azul = 0
    While Not RS.EOF
        Label1.Caption = RS!codartic
        Label1.Refresh
        If Codfamia <> RS!Codfamia Then
            Set It = ListView1.ListItems.Add(, "F" & RS!Codfamia, "******")
            It.SubItems(1) = RS!nomfamia
            It.Bold = True
            It.ForeColor = &H8000&
            It.ListSubItems(1).ForeColor = &H8000&
            Codfamia = RS!Codfamia
        End If
        
        
        'If RS!codartic = "002400112602" Then Stop
        
        Set It = ListView1.ListItems.Add(, , CStr(RS!codartic))
        
        With It
            .Text = RS!codartic
            .SubItems(1) = RS!nomartic
            '.SubItems(2) = Format(RS!canstock, FormatoCantidad)
            .SubItems(3) = Format(RS!stockinv, FormatoCantidad)
            Actual = RS!canstock
            'If RS!stockinv < 0 Then Stop
        
            FijarCantidadesPosteriores It, Posteriores, MovimientodDFI, AnteriorSmoval, TieneMovDFI

            
            .SubItems(2) = Format(AnteriorSmoval, FormatoCantidad)
            .SubItems(4) = Format(Posteriores, FormatoCantidad)
            .SubItems(5) = Format((RS!canstock), FormatoCantidad)
            Actual = RS!stockinv + Posteriores
            .SubItems(6) = Format((Actual), FormatoCantidad)
            .SubItems(7) = ""
            If TieneMovDFI Then
                .SubItems(7) = Format((MovimientodDFI), FormatoCantidad)
            Else
                If MovimientodDFI <> 0 Then Stop 'NO DEBE PASAR
            End If
            Suma = RS!canstock - Actual
            .SubItems(9) = ""
            If Suma <> 0 Then .SubItems(9) = Format((Suma), FormatoCantidad)
                
                
            .SubItems(8) = ""
            
            
            Suma = AnteriorSmoval + MovimientodDFI
            Suma = RS!stockinv - Suma
            
            If Suma <> 0 Then
                 
                
                 'If RS!codartic = "000700030306" Then Stop
                
                Suma = RS!stockinv - AnteriorSmoval
                
                 
                .SubItems(8) = Format(Suma, FormatoCantidad)
            
                'Si ademas no couincide el stcok
                If Actual <> RS!canstock Then
                    It.ForeColor = vbCyan
                    It.ListSubItems(1).ForeColor = vbCyan
                    Azul = Azul + 1
                Else
                    'SOlo no coincde el DFI
                    Rojos = Rojos + 1
                    It.ForeColor = vbRed
                    It.ListSubItems(1).ForeColor = vbRed
                End If
            End If
        
        End With
        RS.MoveNext
    Wend
     RS.Close
     
       Label1.Caption = "Rojos: " & Rojos & "      Azules: " & Azul
     
     
End Sub

Private Sub FijarCantidadesPosteriores(ByRef Itm As ListItem, ByRef Posteriores As Currency, ByRef MovDFI As Currency, ByRef AnteriorEnSmoval As Currency, ByRef TieneMovimientoDFI As Boolean)
Dim C As String
Dim HoraInv As String
Dim RT As ADODB.Recordset
    
    
    'El movimiento de diferencia de inventario
    Set RT = New ADODB.Recordset
    C = "Select * from smoval where codalmac= " & Text1.Text & "  and detamovi='DFI' and codartic ='" & Itm.Text & "'"
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MovDFI = 0
    HoraInv = "'" & Format(Text2.Text, FormatoFecha) & " 23:59:59'"
    TieneMovimientoDFI = False
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then
            HoraInv = "'" & Format(RT!horamovi, "yyyy-mm-dd hh:nn:ss") & "'"
            TieneMovimientoDFI = True
            MovDFI = RT!cantidad
            'If MovDFI < 0 Then Stop
            If RT!tipomovi = 0 Then MovDFI = -MovDFI
            
        End If
    End If
    RT.Close
    Set RT = Nothing
    
    

    C = "Select sum(if(tipomovi=1,cantidad,-cantidad)) from smoval where codalmac= " & Text1.Text & "  and codartic='" & Itm.Text & "'"
    'C = C & " AND fechamov>'2009-10-30'"
    C = C & " AND (fechamov>'" & Format(Text2.Text, FormatoFecha) & "' or (fechamov='" & Format(Text2.Text, FormatoFecha) & "' and horamovi>" & HoraInv & "))"
    Set RT = New ADODB.Recordset
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Posteriores = 0
    
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then Posteriores = RT.Fields(0)
    End If
    RT.Close

    
    
    
    C = "Select sum(if(tipomovi=1,cantidad,-cantidad)) from smoval where codalmac= " & Text1.Text & "  and codartic='" & Itm.Text & "'"
    'C = C & " AND fechamov<'2009-10-30' or fechamofand detamovi<>'DFI'"
    C = C & " AND (fechamov<'" & Format(Text2.Text, FormatoFecha) & "' or (fechamov='" & Format(Text2.Text, FormatoFecha) & "' and horamovi<" & HoraInv & "))"
    Set RT = New ADODB.Recordset
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AnteriorEnSmoval = 0
    
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then AnteriorEnSmoval = RT.Fields(0)
    End If
    RT.Close

    
    
    
End Sub

Private Sub Command1_Click()
Dim I As Integer
    If MsgBox("Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).ForeColor = vbCyan Or ListView1.ListItems(I).ForeColor = vbRed Then
                If ListView1.ListItems(I).SubItems(8) <> "" Then
                    Corregir I
                    ListView1.ListItems(I).ForeColor = vbBlack
                    ListView1.ListItems(I).Checked = False
                End If
                    
            End If
        End If
    Next
    
End Sub


Private Sub Corregir(ByRef Indice As Integer)
Dim Im As Currency
Dim SQL As String
Dim S2 As String
Dim RS As ADODB.Recordset

    Im = ImporteFormateado(ListView1.ListItems(Indice).SubItems(8))
    
    
    
            If ListView1.ListItems(Indice).SubItems(7) = "" Then
                'NO EXISTIA
                
                Set RS = New ADODB.Recordset
                SQL = "Select * from smoval where codalmac= " & Text1.Text & "  and fechamov='" & Format(Text2.Text, FormatoFecha) & "' and codartic = '" & ListView1.ListItems(Indice).Text & "'"
                SQL = SQL & " ORDER BY horamovi desc"
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = "'" & Format(Text2.Text, FormatoFecha) & " 23:59:59'"
                If Not RS.EOF Then
                    Stop
                End If
                RS.Close
                Set RS = Nothing
                'INSERTARLO
                S2 = "INSERT INTO smoval(codartic,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,letraser,document,numlinea) VALUES ('"
                S2 = S2 & ListView1.ListItems(Indice).Text & "',1,'" & Format(Text2.Text, FormatoFecha) & "'," & SQL & ","
                'tipomovi
                If Im < 0 Then
                    S2 = S2 & "0"
                    Im = Abs(Im)
                Else
                    S2 = S2 & "1"
                End If
                'impormov,codigope,letraser,document,numlinea)             2:ramon
                S2 = S2 & ",'DFI'," & TransformaComasPuntos(CStr(Im)) & ",0,2,'',1,1000)"
                SQL = S2
                
            Else
                'EXISTE
            
                
                SQL = "UPDATE smoval set tipomovi= "
                If Im < 0 Then
                    SQL = SQL & "0"
                    Im = Abs(Im)
                Else
                    SQL = SQL & "1"
                End If
                SQL = SQL & " ,cantidad=" & TransformaComasPuntos(CStr(Im))
                SQL = SQL & " WHERE detamovi='DFI' and codalmac= " & Text1.Text & "  and codartic ='" & ListView1.ListItems(Indice).Text & "'"
                SQL = SQL & " AND fechamov='" & Format(Text2.Text, FormatoFecha) & "' "
                
            End If
            Conn.Execute SQL
    
   
    
    
End Sub


Private Sub Command2_Click()
Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
       ' If ListView1.ListItems(I).ForeColor = vbCyan Or ListView1.ListItems(I).ForeColor = vbRed Then
            ListView1.ListItems(I).Checked = Not ListView1.ListItems(I).Checked
       ' Else
        '    ListView1.ListItems(I).Checked = False
       ' End If
    Next I
End Sub

Private Sub Command3_Click()
Dim I As Integer
Dim C As String
Dim Im As Currency
    'El stock
   For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).SubItems(9) <> "" Then
                'Actualizmos stock
                 Im = ImporteFormateado(ListView1.ListItems(I).SubItems(6))
                
                C = "UPDATE  salmac set canstock = " & TransformaComasPuntos(CStr(Im))
                C = C & " WHERE codalmac= " & Text1.Text & "  and codartic= '" & ListView1.ListItems(I).Text & "'"
                Conn.Execute C
            End If
        End If
        ListView1.ListItems(I).Checked = False
    Next I
End Sub

Private Sub Command4_Click()
    If Text1.Text = "" Then Exit Sub
    If Not IsNumeric(Text1.Text) Then Exit Sub
    If Val(Text1.Text) = 0 Then Exit Sub
    If Text2.Text = "" Then Exit Sub
    HacerACciones
End Sub

Private Sub Form_Activate()
    If PrimV Then
        PrimV = False
        
    End If
End Sub

Private Sub Form_Load()
    PrimV = True
    Me.Icon = frmPpUtuil.Icon
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    InputBox ListView1.SelectedItem.SubItems(1), , ListView1.SelectedItem.Text
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Trim(Text2.Text)
    If Text2.Text = "" Then Exit Sub
    
    If Not EsFechaOKTex(Text2) Then
        Text2.Text = ""
        PonerFoco Text2
    End If
End Sub
