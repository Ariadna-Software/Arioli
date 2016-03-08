VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStock 
   Caption         =   "Stock"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCorregir 
      Caption         =   "Corregir"
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Seleccionar/Des."
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Almacen"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Text            =   "1"
      ToolTipText     =   "Almacen"
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "codartic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nomartic"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Sotck"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "smoval"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Diferencia"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Familia"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Almacen"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   3015
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimV As Boolean

Private Sub HacerACciones()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Codfamia As Integer
Dim Actual As Currency
Dim movSalm As Currency
    ListView1.ListItems.Clear

Dim it As ListItem
SQL = "select salmac.*,sartic.codfamia,nomartic,sfamia.nomfamia from salmac,sartic,sfamia where "
SQL = SQL & " salmac.codartic=sartic.codartic and sfamia.codfamia=sartic.codfamia "
SQL = SQL & " and codalmac=" & txtNumero(0).Text & " AND ctrstock=1"
If txtNumero(1).Text <> "" Then SQL = SQL & " AND sartic.codfamia = " & txtNumero(1).Text

SQL = SQL & " ORDER by sartic.codfamia,nomartic"

Set Rs = New ADODB.Recordset

    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Codfamia = -1
    While Not Rs.EOF
        Label1.Caption = Rs!codartic
        Label1.Refresh
        If Codfamia <> Rs!Codfamia Then
            Set it = ListView1.ListItems.Add(, "F" & Rs!Codfamia, "******  " & Rs!Codfamia)
            it.SubItems(1) = Rs!nomfamia
            it.Bold = True
            it.ForeColor = &H8000&
            it.ListSubItems(1).ForeColor = &H8000&
            it.Tag = 0
            Codfamia = Rs!Codfamia
        End If
        
       ' If RS!codartic = "000300010103" Then Stop
        
        Set it = ListView1.ListItems.Add(, , CStr(Rs!codartic))
        
        With it
            .Text = Rs!codartic
            .SubItems(1) = Rs!nomartic
            
            Actual = Rs!canstock
            .SubItems(2) = Format(Actual, FormatoCantidad)
        
            FijarCantidades it.Index, movSalm
            
            .SubItems(3) = Format(movSalm, FormatoCantidad)
            
            movSalm = Actual - movSalm

            
            If movSalm <> 0 Then
                .SubItems(4) = Format(movSalm, FormatoCantidad)
                'Stop
                it.ForeColor = vbRed
                it.ListSubItems(1).ForeColor = vbRed
            End If
            it.Tag = 1  'los articulos
        End With
        Rs.MoveNext
    Wend
     Rs.Close
     
       Label1.Caption = ""
     
     
End Sub

Private Sub FijarCantidades(ind As Integer, ByRef Totalsmoval As Currency)
Dim C As String
Dim RT As ADODB.Recordset
    

    C = "Select sum(if(tipomovi=1,cantidad,-cantidad)) from smoval where codalmac=" & Me.txtNumero(0).Text & " and codartic='" & ListView1.ListItems(ind).Text & "'"
    'C = C & " AND fechamov>'2009-10-30'"
    Set RT = New ADODB.Recordset
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Totalsmoval = 0
    
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then Totalsmoval = RT.Fields(0)
    End If
    RT.Close

End Sub

Private Sub cmdCorregir_Click()
Dim I As Integer
    If MsgBox("Corregir los seleccionados?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Tag = 1 And ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).ForeColor = vbRed Then
                ActualizarSalmac I
                ListView1.ListItems(I).ForeColor = vbBlack
            End If
        End If
    Next
    
End Sub

Private Sub Command1_Click()
    If txtNumero(0).Text = "" Then
        MsgBox "Almacen"
        Exit Sub
    End If
    
    HacerACciones
End Sub

Private Sub Command2_Click()
Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Tag = 1 Then
            ListView1.ListItems(I).Checked = Not ListView1.ListItems(I).Checked
        End If
    Next
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


Private Sub txtNumero_LostFocus(Index As Integer)
    If txtNumero(Index).Text <> "" Then
        If Not IsNumeric(txtNumero(Index).Text) Then
            MsgBox "Numero"
            txtNumero(Index).Text = ""
        End If
    End If
End Sub


Private Sub ActualizarSalmac(Indice As Integer)
Dim SQL As String
Dim Im As Currency
    On Error GoTo EF

    Im = ImporteFormateado(ListView1.ListItems(Indice).SubItems(3))
    SQL = "UPDATE salmac set canstock=" & TransformaComasPuntos(CStr(Im))
    SQL = SQL & " WHERE codalmac = " & txtNumero(0).Text
    SQL = SQL & " AND codartic = '" & ListView1.ListItems(Indice).Text & "'"
    Conn.Execute SQL
    Exit Sub
EF:
    MuestraError Err.Number, Err.Description
End Sub
