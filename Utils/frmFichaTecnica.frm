VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFichaTecnica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "pesos ficha técnica"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Q"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Quitar seleccion"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Seleccionar todo"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lim"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Añadir articulos"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "quitar articulos"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdGuardaBD 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdArti 
      Caption         =   "+"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Añadir articulos"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11280
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   12255
      _ExtentX        =   21616
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "codartic"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nomartic"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Aceite N"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Aceite B"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Caja N"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Caja B"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Palet N"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Palet B"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   4575
   End
End
Attribute VB_Name = "frmFichaTecnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdArti_Clic()
    
    
End Sub

Private Sub cmdArti_Click()
    CadenaDesdeOtroForm = ""
    frmArti.Show vbModal
    If CadenaDesdeOtroForm <> "" Then CargaArticulo
        
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Function InsertarL(C As String, ByRef It) As Boolean
    On Error Resume Next
    Set It = ListView1.ListItems.Add(, "'" & C & "'")
    If Err.Number <> 0 Then
        InsertarL = False
        Err.Clear
    Else
        InsertarL = True
    End If
End Function

Private Sub CargaArticulo()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Unidades As Integer
Dim TotalCajasPalet As Integer
Dim PesoNetoAceite As Currency
Dim It As ListItem


    SQL = "Select sartic.*,sarti4.* from sarti4,sartic where sarti4.codartic=sartic.codartic and sartic.codartic in (" & CadenaDesdeOtroForm & ")"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not Rs.EOF
        Label1.Caption = Rs!codartic
        Label1.Refresh
        SQL = Rs!codartic
        If InsertarL(SQL, It) Then
            It.Text = SQL
            It.SubItems(1) = Rs!nomartic
            It.Bold = True
            It.ForeColor = &H8000&
            It.ListSubItems(1).ForeColor = &H8000&
            It.Tag = 0
             
            Unidades = DBLet(Rs!pal_udbas, "N")
            TotalCajasPalet = DBLet(Rs!pal_udalt, "N")
            TotalCajasPalet = TotalCajasPalet * Unidades
            Unidades = Rs!UniCajas
            PesoNetoAceite = DBLet(Rs!pesoneto, "N")
            CargaSarti1 SQL, Unidades, TotalCajasPalet, PesoNetoAceite, It
        
        End If
        Rs.MoveNext
    Wend
     Rs.Close
     
       Label1.Caption = ""
     
     
End Sub
Private Sub CargaSarti1(cod As String, UniCajas As Integer, CajasPalet As Integer, PesoNetoAceite As Currency, ByRef It As ListItem)
Dim R As ADODB.Recordset
Dim SQL As String
Dim PesoTapon As Currency
Dim PesoBotella As Currency
Dim OtrosPesos As Currency
Dim PesoBrutoBotella As Currency
Dim aUX As Currency
Dim CajaVacia As Currency
Dim PesoBrutoCaja As Currency
Dim PesoNetoCaja As Currency
Dim PesoRetractil As Currency

    PesoTapon = 0
    PesoBotella = 0
    CajaVacia = 0
    OtrosPesos = 0
    PesoRetractil = 0
    Set R = New ADODB.Recordset
    SQL = "select sarti4.*,tipartic,cantidad,nomartic from sarti4,sarti1,sartic where sarti4.codartic=sarti1.codarti1 and sartic.codartic = sarti4.codartic and sarti1.codartic=" & cod
    R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not R.EOF
        Select Case Val(R!tipartic)
        Case 3
            PesoTapon = DBLet(R!pesoneto, "N")
        Case 2
            PesoBotella = DBLet(R!pesoneto, "N")
           
        Case 8
            'retractil
            PesoRetractil = DBLet(R!pesoneto, "N")    'va por cada botella
        Case Else
            If Val(R!tipartic) = 6 Then
                'Caja
               CajaVacia = DBLet(R!caj_vacia, "N")
            Else
                If Val(R!tipartic) = 9 Then
                    'ESTUCHE
                    aUX = UniCajas
                    aUX = aUX * DBLet(R!pesoneto, "N")
                    OtrosPesos = OtrosPesos + aUX
                Else
                
                    If Val(R!tipartic) > 1 Then
                        aUX = 1  'R!cantidad
                        aUX = aUX * DBLet(R!pesoneto, "N")
                        OtrosPesos = OtrosPesos + aUX
                    End If
                End If
            End If
        End Select
        R.MoveNext
    Wend
    R.Close
    
    If PesoTapon = 0 Or PesoBotella = 0 Then It.ForeColor = vbRed
    PesoBrutoBotella = PesoTapon + PesoBotella + PesoNetoAceite + PesoRetractil
    It.SubItems(2) = Format(PesoNetoAceite, FormatoPrecio)
    It.SubItems(3) = Format(PesoBrutoBotella, FormatoPrecio)
    
    'En la columna pondremos este peso
    
    'CAJA
    'Vamos a calcular la caja
    'neto caja
    PesoNetoCaja = PesoNetoAceite * UniCajas
    
    'bruto caja
    PesoBrutoCaja = (PesoBrutoBotella * UniCajas) + CajaVacia + OtrosPesos
    It.SubItems(4) = Format(PesoNetoCaja, FormatoPrecio)
    It.SubItems(5) = Format(PesoBrutoCaja, FormatoPrecio)
    
    
    
    
    'Ahora vamos con el palet
    aUX = PesoNetoCaja * CajasPalet
    It.SubItems(6) = Format(aUX, FormatoPrecio)
    aUX = PesoBrutoCaja * CajasPalet
    It.SubItems(7) = Format(aUX, FormatoPrecio)
    
    
End Sub

Private Sub cmdGuardaBD_Click()
Dim I As Integer
Dim SQL As String


    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Exit For
    Next
    If I > ListView1.ListItems.Count Then
        MsgBox "Selecciona alguno", vbExclamation
        Exit Sub
        
    Else
        If MsgBox("¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'Para ada referencia....
    Screen.MousePointer = vbHourglass
    For I = 1 To ListView1.ListItems.Count
    
        If ListView1.ListItems(I).Checked Then
            Label1.Caption = ListView1.ListItems(I).Text
            Label1.Refresh
            
            ActualizarReferencia I
        End If
    Next I
    Screen.MousePointer = vbDefault
    
    

End Sub


Private Sub ActualizarReferencia(Indice As Integer)
Dim C As String
Dim A As String
    'Updaeatemos del codaric de venta peso bruto, peso bruto palet y pesoneto palet
    A = ImporteFormateado(ListView1.ListItems(Indice).SubItems(3))
    A = TransformaComasPuntos(A)
    C = "UPDATE sarti4 set pesobruto=" & A
    
    'pal_pneto pal_pbruto
    A = ImporteFormateado(ListView1.ListItems(Indice).SubItems(6))
    A = TransformaComasPuntos(A)
    C = C & ", pal_pneto =" & A
    A = ImporteFormateado(ListView1.ListItems(Indice).SubItems(7))
    A = TransformaComasPuntos(A)
    C = C & ", pal_pbruto =" & A
    
    'IMPORTANTE.Copiar esto en Arioli
    'Para el producto venta , los campos de sarti4  ret_medid ret_seriT
    'seran el peso brut y el peso neto de la caja
     A = ImporteFormateado(ListView1.ListItems(Indice).SubItems(4))
    A = TransformaComasPuntos(A)
    A = "'" & A & " Kg'"
    C = C & ", ret_medid = " & A
     A = ImporteFormateado(ListView1.ListItems(Indice).SubItems(5))
    A = TransformaComasPuntos(A)
    A = "'" & A & " Kg'"
    C = C & ", ret_seriT =" & A
    
    C = C & " WHERE codartic = " & ListView1.ListItems(Indice).Text
    
    Conn.Execute C
End Sub


Private Sub Command1_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub Command2_Click(Index As Integer)
Dim I As Integer
Dim B As Boolean
    If ListView1.ListItems.Count = 0 Then Exit Sub
    B = Index = 0
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = B
    Next
        
End Sub

Private Sub Command3_Click()
    ListView1.ListItems.Clear
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpUtuil.Icon
End Sub

Private Sub ListView1_DblClick()
Dim SQL As String
Dim Rs As ADODB.Recordset

    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    SQL = "select codarti1,nomartic,desctipfamia,pesoneto from sarti1,sartic,sarti4,stipfamia where"
    SQL = SQL & " sarti4.codartic=sarti1.codarti1 and sarti1.codarti1=sartic.codartic and sartic.tipartic=tipfamia"
    SQL = SQL & " and sarti1.codartic = '" & ListView1.SelectedItem.Text & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not Rs.EOF
        SQL = SQL & Mid(Rs!desctipfamia & Space(30), 1, 20) & "        " & Rs!codarti1 & "   " & Rs!nomartic
        SQL = SQL & Right(Space(30) & Format(DBLet(Rs!pesoneto, "N"), FormatoPrecio), 30) & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    SQL = ListView1.SelectedItem.Text & "  " & ListView1.SelectedItem.SubItems(1) & vbCrLf & String(50, "-") & vbCrLf & SQL
    MsgBox SQL, vbInformation
    

End Sub
