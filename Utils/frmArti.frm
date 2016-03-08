VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Articuls"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Seleccionar todos"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Quitar seleccion"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   7920
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   13785
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Articulo"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Command2_Click()
Dim I As Integer
    CadenaDesdeOtroForm = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & ListView1.ListItems(I).Key
    Next I
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Nada seleccionado", vbExclamation
    Else
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        Unload Me
    End If
End Sub

Private Sub Command3_Click(Index As Integer)
Dim I As Integer
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = (Index = 1)
    Next
End Sub

Private Sub Form_Load()
    Dim Rs As ADODB.Recordset
    Dim It
    
     Me.Icon = frmPpUtuil.Icon
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from sartic where conjunto=1", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set It = ListView1.ListItems.Add(, "'" & CStr(Rs!codartic) & "'", CStr(Rs!codartic))

        It.Text = Rs!nomartic
        'It.Checked = True
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub
