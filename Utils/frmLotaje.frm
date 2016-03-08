VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLotaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   945
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   945
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5953
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
         Text            =   "codartic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "nomartic"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Incid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descrip"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   $"frmLotaje.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmLotaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Codfamia As Integer
Dim Actual As Currency
Dim movSalm As Currency
    ListView1.ListItems.Clear

Dim It As ListItem
SQL = "select salmac.*,sartic.codfamia,nomartic,sfamia.nomfamia from salmac,sartic,sfamia where "
SQL = SQL & " salmac.codartic=sartic.codartic and sfamia.codfamia=sartic.codfamia "


Set Rs = New ADODB.Recordset

    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Codfamia = -1
    While Not Rs.EOF
        Label1.Caption = Rs!codartic
        Label1.Refresh
        If Codfamia <> Rs!Codfamia Then
            Set It = ListView1.ListItems.Add(, "F" & Rs!Codfamia, "******  " & Rs!Codfamia)
            It.SubItems(1) = Rs!nomfamia
            It.Bold = True
            It.ForeColor = &H8000&
            It.ListSubItems(1).ForeColor = &H8000&
            It.Tag = 0
            Codfamia = Rs!Codfamia
        End If
        
       
        
        Set It = ListView1.ListItems.Add(, , CStr(Rs!codartic))
        
        With It
            .Text = Rs!codartic
            .SubItems(1) = Rs!nomartic
            
            Actual = Rs!canstock
            .SubItems(2) = Format(Actual, FormatoCantidad)
        
            
            
            .SubItems(3) = Format(movSalm, FormatoCantidad)
            
            movSalm = Actual - movSalm

            
            If movSalm <> 0 Then
                .SubItems(4) = Format(movSalm, FormatoCantidad)
                'Stop
                It.ForeColor = vbRed
                It.ListSubItems(1).ForeColor = vbRed
            End If
            It.Tag = 1  'los articulos
        End With
        Rs.MoveNext
    Wend
     Rs.Close
     
       Label1.Caption = ""
     
     
End Sub



Private Sub Text1_LostFocus(Index As Integer)
    If Text1(Index).Text <> "" Then
        If Not EsFechaOKTex(Text1(Index)) Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
    End If
End Sub
