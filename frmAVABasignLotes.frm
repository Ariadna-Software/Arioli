VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAVABasignLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar lotes "
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   Icon            =   "frmAVABasignLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   10680
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Docum."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Factura"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Albarán"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Articulo"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   6985
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "L."
         Object.Width           =   661
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   1561
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "LOTE"
         Object.Width           =   2558
      EndProperty
   End
End
Attribute VB_Name = "frmAVABasignLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim miSQL As String
Dim It As ListItem



Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        If Not ImportarLotes Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaDatos False
         
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
End Sub




Private Sub CargaDatos(SoloArticulos As Boolean)
     Screen.MousePointer = vbHourglass
     Set miRsAux = New ADODB.Recordset
     If Not SoloArticulos Then CargaAlbaranes
     CargaLineas
     Set miRsAux = Nothing
     Screen.MousePointer = vbDefault
End Sub

Private Sub CargaAlbaranes()
    
    
    If Me.Option1(0).Value Then
        miSQL = "Select codtipom,numalbar,fechaalb from "
        miSQL = miSQL & "ariges" & EmprMorales & ".scaalb WHERE codtipom='ALV' AND "
    Else
       
        miSQL = "Select c.codtipom,c.numfactu,c.fecfactu,numalbar from "
        miSQL = miSQL & "ariges" & EmprMorales & ".scafac c, ariges" & EmprMorales & ".scafac1 c2"
        miSQL = miSQL & " where c.codtipom=c2.codtipom and c.numfactu=c2.numfactu and c.fecfactu=c2.fecfactu "
        miSQL = miSQL & " and c.fecfactu>= '" & Format(DateAdd("m", -33, Now), FormatoFecha) & "' AND"
        miSQL = miSQL & " c.codtipom='FAV' AND "
    End If
    miSQL = miSQL & " codclien = 1"  'AVAB
    miRsAux.Open miSQL, Conn, adOpenForwardOnly, adCmdText
    ListView1.ListItems.Clear
    
    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add()
        It.Text = miRsAux.Fields(0)
        It.SubItems(1) = miRsAux.Fields(1)
        It.SubItems(2) = miRsAux.Fields(2)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If ListView1.ListItems.Count > 0 Then Set ListView1.SelectedItem = ListView1.ListItems(1)
    
    
    
End Sub



Private Sub CargaLineas()

    ListView2.ListItems.Clear
    
    
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
    
    
    If Me.Option1(0).Value Then
    
        miSQL = "select codartic,nomartic,linea,slialblotes.cantidad cantidad,numlote"
        
        miSQL = miSQL & " from  ariges" & EmprMorales & ".slialblotes, ariges" & EmprMorales & ".slialb where"
        miSQL = miSQL & " slialblotes.codTipoM = slialb.codTipoM And slialb.NumAlbar = slialblotes.NumAlbar"
        miSQL = miSQL & " and slialb.numlinea=slialblotes.numlinea"
        miSQL = miSQL & " AND slialb.codtipom='" & ListView1.SelectedItem.Text
        miSQL = miSQL & "' AND slialb.numalbar= " & ListView1.SelectedItem.SubItems(1)
        
    Else
        miSQL = "select codartic,nomartic,linea,t.cantidad cantidad,numlote"
        miSQL = miSQL & " from ariges" & EmprMorales & ".scafac1 c2,ariges" & EmprMorales & ".slifac l, ariges" & EmprMorales & ".slifaclotes t where"
        miSQL = miSQL & " l.codtipom=c2.codtipom and l.numfactu=c2.numfactu and l.fecfactu=c2.fecfactu and l.codtipoa=c2.codtipoa and l.numalbar=c2.numalbar and"
        miSQL = miSQL & " L.codTipoM = T.codTipoM And L.NumFactu = T.NumFactu And L.FecFactu = T.FecFactu And L.codtipoa = T.codtipoa And L.NumAlbar = T.NumAlbar And L.numlinea = T.numlinea"
        'El link
        miSQL = miSQL & " AND c2.codtipom='" & ListView1.SelectedItem.Text
        miSQL = miSQL & "' AND c2.NumFactu= " & ListView1.SelectedItem.SubItems(1)
        miSQL = miSQL & " AND c2.fecFactu= " & DBSet(ListView1.SelectedItem.SubItems(2), "F")
        
    End If
    
    
    miRsAux.Open miSQL, Conn, adOpenForwardOnly, adCmdText
    
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add()
        It.Text = miRsAux.Fields(0)
        It.SubItems(1) = miRsAux.Fields(1)
        It.SubItems(2) = miRsAux.Fields(2)
        It.SubItems(3) = miRsAux.Fields(3)
        It.SubItems(4) = miRsAux.Fields(4)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CargaDatos True
End Sub

Private Sub Option1_Click(Index As Integer)
    CargaDatos False
End Sub



Private Function ImportarLotes() As Boolean
    
    On Error GoTo EImportarLotes
    ImportarLotes = False
    
    If Me.ListView2.ListItems.Count = 0 Then Exit Function
    
    
    'Veremos si el albaran origen tiene
    

EImportarLotes:
    MuestraError Err.Number, ": ImportarLotes"
End Function
