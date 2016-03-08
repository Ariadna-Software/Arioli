VERSION 5.00
Begin VB.Form frmAyudaDemo 
   Caption         =   "AYUDA DEMO"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Palet"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Caja fin"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Caja Inicio"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Lote traza"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Simular lectura cajas desde poste paletizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmAyudaDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String
Dim I As Integer

Private Sub cmdGenerar_Click()

    On Error GoTo EcmdGenerar
    
    'Generara una entrada en la tabla
    If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Then Exit Sub
    
    If Val(Text1(0).Text) = 0 Or Val(Text1(1).Text) = 0 Or Val(Text1(2).Text) = 0 Then Exit Sub
    
    'Si pone palet lo insertaremos con es ID de palet
    If Text1(3).Text <> "" Then
            If Val(Text1(3).Text) = 0 Then Exit Sub
    End If
    
    SQL = ""
    For I = Val(Text1(1).Text) To Val(Text1(2).Text)
        'prodcajas(lotetraza,idcaja,idpalet fcreacion)
        SQL = SQL & ", (" & Text1(0).Text & "," & I
        If Text1(3).Text <> "" Then
            SQL = SQL & "," & Text1(3).Text & ",'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        Else
            SQL = SQL & ",NULL,NULL"
        End If
        SQL = SQL & ")"
    Next I
    
    If SQL = "" Then
        MsgBox "No se generar nada", vbExclamation
    Else
        SQL = Mid(SQL, 2)
        SQL = "INSERT INTO prodcajas(lotetraza,idcaja,idpalet,fcreacion)  VALUES " & SQL
        Conn.Execute SQL
    End If
    Exit Sub
EcmdGenerar:
    MsgBox Err.Description, vbExclamation
    Err.Clear
End Sub


Private Sub Form_Load()
    Me.Icon = frmPpUtuil.Icon
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    'keyrpressgnral
End Sub
