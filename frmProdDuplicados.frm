VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProdDuplicados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cajas duplicadas"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "frmProdDuplicados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdReestablecer 
      Caption         =   "Reestablecer"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lote"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Caja"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Linea"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Articulo"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "frmProdDuplicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReestablecer_Click()
    If MsgBox("Seguro que desea reestablecer?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Conn.Execute "Delete from prodcajasduplicadas"
    
    
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    CargaDatos
    CadenaDesdeOtroForm = ""
End Sub


Private Sub CargaDatos()
On Error GoTo Ec
Dim IT
    CadenaDesdeOtroForm = "Select * from prodcajasduplicadas"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!lotetraza
        IT.SubItems(1) = miRsAux!idcaja
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'Cargo las lineas
   
        

    CadenaDesdeOtroForm = "select lotetraza,lineaprod,nomartic from prodlin,prodtrazlin,sartic "
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " = prodtrazlin.idlin AND prodlin.codartic=sartic.codartic"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND lotetraza in (Select lotetraza from prodcajasduplicadas)"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For NumRegElim = 1 To Me.ListView1.ListItems.Count
            If Val(ListView1.ListItems(NumRegElim).Text) = Val(miRsAux!lotetraza) Then
             
                 ListView1.ListItems(NumRegElim).SubItems(2) = miRsAux!lineaprod
                 ListView1.ListItems(NumRegElim).SubItems(3) = miRsAux!NomArtic
            End If
        Next
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
Ec:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando cajas duplicadas"
    Set miRsAux = Nothing
End Sub
