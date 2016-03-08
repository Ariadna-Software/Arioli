VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacRutasEnvioLineas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar albaranes ruta"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7890
   Icon            =   "frmFacRutasEnvioLineas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11033
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   5115
      EndProperty
   End
End
Attribute VB_Name = "frmFacRutasEnvioLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public vCodigoCabcera As Long

Dim PrimeraVez As Boolean
Dim SQL As String


Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        SQL = ""
        For NumRegElim = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(NumRegElim).Checked Then
                SQL = "X"
                Exit For
            End If
        Next
        
        If SQL = "" Then
            SQL = "Deberia seleccionar algún albaran.  ¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
        End If
    
        Conn.BeginTrans
        If Not HacerInsercion Then
            Conn.RollbackTrans
            Exit Sub
        Else
            Conn.CommitTrans
        End If
    
        CadenaDesdeOtroForm = "OK" 'PAra que refresque el listview
    Else
        CadenaDesdeOtroForm = ""
    End If
    'SAlimos
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        DoEvents
        Set miRsAux = New ADODB.Recordset
        CargaDatos
        Set miRsAux = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    PrimeraVez = True
        
End Sub


Private Sub CargaDatos()

    'ALbaranes
    
    SQL = "select a.numalbar,a.codtipom,a.fechaalb,nomclien,id from scaalb a left join srepartol s on s.codtipom=a.codtipom and"
    SQL = SQL & " s.numalbar=a.numalbar and s.fechaalb=a.fechaalb"
    SQL = SQL & " WHERE a.codtipom='ALV'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then CargaItems True
    miRsAux.Close


    'Veremos de meter las facturas aunque no se puedan modficar
    
End Sub



 
Private Sub CargaItems(Albaranes As Boolean)
Dim IT As ListItem
Dim Insertar As Boolean
    While Not miRsAux.EOF
        Insertar = True
        If Not IsNull(miRsAux!id) Then
            If miRsAux!id <> vCodigoCabcera Then Insertar = False
        End If
        If Insertar Then
            Set IT = ListView1.ListItems.Add()
            IT.Text = miRsAux!codTipoM
            IT.SubItems(1) = miRsAux!NumAlbar
            IT.SubItems(2) = miRsAux!FechaAlb
            IT.SubItems(3) = miRsAux!nomclien
            IT.Tag = DBLet(miRsAux!id, "N")
            If Val(IT.Tag) > 0 Then
                IT.Checked = True
                IT.Bold = True
            End If
        End If
        miRsAux.MoveNext
    Wend
End Sub



Private Function HacerInsercion() As Boolean

    On Error GoTo EHacerInsercion
    HacerInsercion = False



    'QUitare los que ya estaban
    'No puedo hacer delete de las lineas pq hay lines que referencian a facturas.
'        For NumRegElim = 1 To ListView1.ListItems.Count
'        If ListView1.ListItems(NumRegElim).Checked Then




    SQL = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then
            'Si es mayor que cero significa que ya estaba. Luego NO lo inserto
            If Val(ListView1.ListItems(NumRegElim).Tag) = 0 Then
                SQL = SQL & ", (" & Me.vCodigoCabcera & ",'" & ListView1.ListItems(NumRegElim).Text & "',"
                SQL = SQL & ListView1.ListItems(NumRegElim).SubItems(1) & ",'" & Format(ListView1.ListItems(NumRegElim).SubItems(2), FormatoFecha) & "')"
            End If
        Else
            'NO esta seleccionado. Veo si lo estaba
            If Val(ListView1.ListItems(NumRegElim).Tag) > 0 Then
                'Lo borro. Ahora no deberia estar
                CadenaDesdeOtroForm = "DELETE from srepartol WHERE id = " & Me.vCodigoCabcera & " AND codtipom='" & ListView1.ListItems(NumRegElim).Text & "'"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND numalbar= " & ListView1.ListItems(NumRegElim).SubItems(1)
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND FechaAlb= '" & Format(ListView1.ListItems(NumRegElim).SubItems(2), FormatoFecha) & "'"
                Conn.Execute CadenaDesdeOtroForm
            End If
        End If
    Next
    If SQL <> "" Then
                                                                                       'quitamos la primera coma
        SQL = "insert into `srepartol` (`id`,`codtipom`,`numalbar`,`FechaAlb`) values " & Mid(SQL, 2)
        Conn.Execute SQL
    End If

    HacerInsercion = True
    Exit Function
EHacerInsercion:
    MuestraError Err.Number
End Function
