VERSION 5.00
Begin VB.Form frmComPreciosActualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar precios proveedor"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCambiarPrecio 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblPrecio 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha cambio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image imgFecha 
      Height          =   255
      Index           =   0
      Left            =   1560
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmComPreciosActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim cad As String 'multiproposito


Private Sub CargarIconos()
Dim I As Image


    'For Each i In Me.imgArticulo
    '     i.Picture = frmppal.imgListComun.ListImages(19).Picture
    '     i.ToolTipText = "Articulo"
    'Next
    For Each I In Me.imgFecha
         I.Picture = frmppal.imgListComun.ListImages(23).Picture
         I.ToolTipText = "fecha"
    Next
End Sub




Private Sub cmdCambiarPrecio_Click()
    If txtFecha(0).Text = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    '-----------------
    HacerCambioPrecios
    lblPrecio.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    limpiar Me
    CargarIconos
    lblPrecio.Caption = ""
    
    txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then txtFecha(Index).Text = cad
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        cad = txtFecha(Index).Text
        If Not EsFechaOK(cad) Then
            MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        Else
            txtFecha(Index).Text = cad
        End If
    End If
End Sub



Private Sub HacerCambioPrecios()
Dim I As Integer
    'Comprobaremos si hay datos para realizar el cambio
    cad = "select count(*) from slispr where fechanue<='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim = 0 Then
        MsgBox "Ningún precio para actualizar", vbExclamation
    Else
        cad = "Va a actualizar " & NumRegElim & " registro(s)"
        cad = cad & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
    End If
    If NumRegElim = 0 Then Exit Sub
    
    'Ok, vamos p'alla
    cad = "select slispr.*,nomartic from slispr,sartic where"
    cad = cad & " sartic.codartic=slispr.codartic AND fechanue<='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        
        lblPrecio.Caption = miRsAux!NomArtic & " (" & I & " de " & NumRegElim & ")"
        lblPrecio.Refresh
        CambiarUnaReferenciaPrecios
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    'Nos salimos
    Unload Me
End Sub

'EN mirsaux esta el registro
'Tendra todos los datos necesarios. No hace falta pasr parametros
Private Sub CambiarUnaReferenciaPrecios()
Dim R As ADODB.Recordset
Dim C As String
Dim NumLin As Integer

On Error GoTo ECambiarUnaReferenciaPrecios

    Set R = New ADODB.Recordset
    'codartic codprove numlinea fechacam precioac
    
    C = "Select max(numlinea) from slisp1 "
    C = C & " WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
    C = C & " AND codprove = " & miRsAux!codProve
    R.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumLin = 0
    If Not R.EOF Then NumLin = DBLet(R.Fields(0), "N")
    R.Close
    NumLin = NumLin + 1
    
    'Insertamos en la siguiente
    C = "insert into slisp1 (codartic,codprove,numlinea,fechacam,precioac) VALUES ("
    C = C & DBSet(miRsAux!codArtic, "T") & "," & miRsAux!codProve & ","
    C = C & NumLin & ",'" & Format(miRsAux!fechanue, FormatoFecha)
    C = C & "'," & TransformaComasPuntos(CStr(DBLet(miRsAux!precioac, "N"))) & ")"
    Conn.Execute C

    'QUITAMOS LOS VALORES en la tabla "cabeceras"
    C = "UPDATE slispr SET "
    C = C & " precioac= precionu, fechanue=NULL, precionu=NULL"
    C = C & " WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
    C = C & " AND codprove = " & miRsAux!codProve
    Conn.Execute C
    
    
ECambiarUnaReferenciaPrecios:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set R = Nothing
End Sub
