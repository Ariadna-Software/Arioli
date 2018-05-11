VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVallPrecioLiquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidacion"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3000
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2400
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Albaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Oliva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6840
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   7575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   2400
         ToolTipText     =   "Buscar cliente"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "BD acces con liquidaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   10695
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7646
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
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   7800
         TabIndex        =   9
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   9240
         TabIndex        =   8
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   5280
         Width           =   4095
      End
      Begin VB.Label lblError 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmVallPrecioLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Acces As Connection
Dim RS As ADODB.Recordset
Dim Modo As Byte
Dim Cad As String
Dim I As Integer

Private Sub cmdAceptar_Click(Index As Integer)
    Select Case Index
    Case 0
        If Text1.Text = "" Then Exit Sub
        If Dir(Text1.Text, vbArchive) = "" Then
            MsgBox "No existe el archivo ", vbExclamation
            Exit Sub
        End If
        
        If LCase(Right(Text1.Text, 4)) <> ".mdb" Then
            MsgBox "No es extension Access", vbExclamation
            Exit Sub
        End If
         Set RS = New ADODB.Recordset
        If AbrirAccess Then
            PonerModo 1
        Else
            Set Acces = Nothing
        End If
         Set RS = New ADODB.Recordset
    Case 1
        Cad = ""
        For I = 1 To 4
            If Combo1(I).ListIndex < 0 Then Cad = "N"
        Next
        If Cad <> "" Then
            MsgBox "seleccione campos importacion", vbExclamation
            Exit Sub
        End If
        
        Cad = "|"
        For I = 1 To 4
            If InStr(1, Cad, Format(Combo1(I).ListIndex, "0000") & "|") > 0 Then
                MsgBox "No puede seleccinar el mismo campo dos veces", vbExclamation
                Exit Sub
            Else
                Cad = Cad & Format(Combo1(I).ListIndex, "0000") & "|"
            End If
        Next
        
        Screen.MousePointer = vbHourglass
        lblI.Caption = ""
        If CuadrarDatos Then
            PonerModo 2
            
            cmdAceptar(2).Enabled = NumRegElim = 0
                    
            If NumRegElim = 0 Then
                'CargaDatosOk
                CargaDatos
            Else
                'Carga ERRORES
                CargaErrores
            End If
        End If
        Screen.MousePointer = vbDefault
        lblI.Caption = ""
    Case 2
    
        If MsgBox("Desea realizar la actualizacion del precio/importe del albaran en Ariges?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
        If RealizaActualizacion Then Unload Me
        
    End Select
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index < 2 Then
        Unload Me
    Else
        PonerModo 1
    End If
End Sub

Private Sub Combo1_Click(Index As Integer)
    
    
    If Index > 0 Then Exit Sub
     For I = 1 To 4
        Combo1(I).Clear
    Next
    
    If Combo1(0).ListIndex <= 0 Then Exit Sub
    
    Set RS = New ADODB.Recordset
    RS.Open "select * from " & Combo1(0).Text & " WHERE false", Acces, adOpenKeyset, adLockPessimistic, adCmdText
    Cad = ""
    For I = 0 To RS.Fields.Count - 1
        Combo1(1).AddItem RS.Fields(I).Name
        Combo1(2).AddItem RS.Fields(I).Name
        Combo1(3).AddItem RS.Fields(I).Name
        Combo1(4).AddItem RS.Fields(I).Name
    Next
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Text1.Text = ""
    Me.imgBuscar(0).Picture = frmppal.imgListComun.ListImages(19).Picture
    
    PonerModo 0
End Sub


Private Sub PonerModo(vNuevoModo As Byte)
Dim H As Long
Dim W As Long

    Me.Frame1.visible = vNuevoModo = 0
    Me.Frame2.visible = vNuevoModo = 1
    Me.Frame3.visible = vNuevoModo = 2
    
    Me.cmdCancelar(0).Cancel = vNuevoModo = 0
    Me.cmdCancelar(1).Cancel = vNuevoModo = 1
    Me.cmdCancelar(2).Cancel = vNuevoModo = 2
    
    
    
    If vNuevoModo = 0 Then
        W = Frame1.Width
        H = Frame1.Height
    
    ElseIf vNuevoModo = 1 Then
        W = Frame2.Width
        H = Frame2.Height
    Else
        W = Frame3.Width
        H = Frame3.Height
    End If
    H = H + 240
    W = W + 240
    Me.Width = W
    Me.Height = H
    
    lblTotal.Caption = ""
    Modo = vNuevoModo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Acces.Close
    Set Acces = Nothing
    Err.Clear
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    On Error Resume Next
    
    frmppal.CommonDialog1.DefaultExt = ".mdb" 'extension por defecto
    frmppal.CommonDialog1.Filter = "Acces BD |*.mdb|" 'extensiones a mostrar
    frmppal.CommonDialog1.FilterIndex = 1
    frmppal.CommonDialog1.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear
    Else
        If frmppal.CommonDialog1.FileName <> "" Then Text1.Text = frmppal.CommonDialog1.FileName
    End If
End Sub

Private Function AbrirAccess() As Boolean


    On Error GoTo eAbrirAccess
    AbrirAccess = False
    
    For I = 0 To 4
        Combo1(I).Clear
    Next
    Combo1(0).AddItem ""
    
    Cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text1.Text
    Set Acces = New Connection
    Acces.Open Cad
    
    Cad = "tablas"
   
    Set RS = Acces.OpenSchema(adSchemaTables)
    Cad = ""
    While Not RS.EOF
        For I = 0 To 3
            
            Cad = Cad & DBLet(RS.Fields(I).Name, "T") & ": " & DBLet(RS.Fields(I).Value, "T") & "    "
        Next
        'Deb ug.Print Cad
        Cad = ""
        If RS!TABLE_TYPE = "TABLE" Then Combo1(0).AddItem RS!TABLE_NAME
        RS.MoveNext
    Wend
    RS.Close
    
    
    AbrirAccess = True
    Exit Function
eAbrirAccess:
    MuestraError Err.Number, Cad & vbCrLf & Err.Description
End Function

Private Function CuadrarDatos() As Boolean
Dim C As String
Dim J As Integer

    On Error GoTo eCuadrarDatos

    CuadrarDatos = False
    
    lblI.Caption = "Olivas"
    lblI.Refresh
    
    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    
    
    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    'Veremos los articulos si son olivas lo que toca
    'Las oliVAS
    Cad = "Select distinct `" & Combo1(1).Text & "` FROM " & Combo1(0).Text
    RS.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    NumRegElim = 0
    If RS.EOF Then
        Err.Raise 513, , "Ningun dato en BD Access"
    Else
        While Not RS.EOF
            Cad = Cad & ", " & DBSet(RS.Fields(0), "T")
            NumRegElim = NumRegElim + 1
            RS.MoveNext
        Wend
    End If
    RS.Close
    
    
    If Cad = "" Then Err.Raise 513, , "Ningun articulo " & Combo1(1).Text
    
    C = "Select codartic,codfamia from sartic where codfamia=100 AND codartic IN (" & Mid(Cad, 2) & ")"
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        NumRegElim = NumRegElim - 1
        RS.MoveNext
    Wend
    RS.Close
    
    If NumRegElim > 0 Then Err.Raise 513, , "Algun articulos no es OLIVA " & Cad
    
    
    
    '----------------------------------------------------------------------------
    lblI.Caption = "Importes y cantidad"
    lblI.Refresh
    Cad = "Select  `" & Combo1(3).Text & "` ,`" & Combo1(4).Text & "` "
    Cad = Cad & " FROM " & Combo1(0).Text
    RS.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    C = ""
    While Not RS.EOF
        'Campos vacios
        If IsNull(RS.Fields(0)) Then
            Cad = Cad & "X"
        Else
            'NO numerico
            If Not IsNumeric(RS.Fields(0)) Then C = C & "X"
        End If
        
        If IsNull(RS.Fields(1)) Then
            Cad = Cad & "X"
        Else
            'NO numerico
            If Not IsNumeric(RS.Fields(1)) Then C = C & "X"
        End If
        
        
        RS.MoveNext
    Wend
    RS.Close
    
    
    If Cad <> "" Then Cad = "Campos vacios  (" & Len(Cad) & ")"
    If C <> "" Then Cad = Cad & vbCrLf & "Campos no numericos (" & Len(C) & ")"
    
    If Cad <> "" Then Err.Raise 513, , "Datos incorrectos " & Cad
    
    
    
    
    
    
    '----------------------------------------------------------------------------
    lblI.Caption = "Comprobacion albaranes"
    lblI.Refresh
    
    'Cargamos todos los albaranes
    Cad = "select numalbar,slialp.codartic,cantidad from slialp,sartic where sartic.codartic=slialp.codartic and codfamia=100"
    RS.Open Cad, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    
    Cad = "Select  `" & Combo1(1).Text & "` ,`" & Combo1(2).Text & "` "
    Cad = Cad & " , `" & Combo1(3).Text & "` ,`" & Combo1(4).Text & "` "
    Cad = Cad & " FROM " & Combo1(0).Text
    miRsAux.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        lblI.Caption = "Albaran: " & DBLet(miRsAux.Fields(1), "T")
        lblI.Refresh
    
        RS.Find "numalbar = " & DBSet(miRsAux.Fields(1), "T"), , adSearchForward, 1
        
        'c="insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`,`importe1`
        C = ""
        If RS.EOF Then
            '`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`
            C = "1,0,'No existe albaran'," & DBSet(miRsAux.Fields(1), "T") & ",null"
            
        Else
        
            'Misma OLIVA
            If miRsAux.Fields(0) <> RS!codartic Then
                C = "2,0,'No mismo articulo albaran'," & DBSet(miRsAux.Fields(1), "T") & "," & DBSet(RS!codartic, "T")
            Else
                'o-----
                'Ok. Si existe el albaran. y mismo articulo
                'Que la cantidad es la misma
                
                If miRsAux.Fields(2) <> RS!Cantidad Then C = "3,0,'Cantidad distinta'," & DBSet(miRsAux.Fields(1), "T") & "," & DBSet(RS!Cantidad, "T")
            End If
        End If
        
        If C <> "" Then
            NumRegElim = NumRegElim + 1
            C = " (" & vUsu.Codigo & "," & NumRegElim & "," & C & ")"
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`) VALUES " & C
            conn.Execute C
        End If
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    RS.Close
    
    
    
    
    '----------------------------------------------------------------------------
    lblI.Caption = "Comprobacion pendientes "
    lblI.Refresh
    
    'Cargamos todos los albaranes
    Cad = "Select  `" & Combo1(1).Text & "` ,`" & Combo1(2).Text & "` "
    Cad = Cad & " , `" & Combo1(3).Text & "` ,`" & Combo1(4).Text & "` "
    Cad = Cad & " FROM " & Combo1(0).Text
    RS.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Cad = "select numalbar,slialp.codartic,cantidad from slialp,sartic where sartic.codartic=slialp.codartic and codfamia=100"
    Cad = Cad & " AND fechaalb >=" & DBSet(vParamAplic.FechaActiva, "F")
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
   
    While Not miRsAux.EOF
        lblI.Caption = "Albaran Acce: " & DBLet(miRsAux.Fields(0), "T")
        lblI.Refresh
        
        C = "[" & Combo1(2).Text & "] = " & DBSet(miRsAux.Fields(0), "T")
        RS.Find C, , adSearchForward, 1
        
        'c="insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`,`importe1`
        C = ""
        If RS.EOF Then
            '`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`
            C = "3,0,'No existe albaran en Access'," & DBSet(miRsAux.Fields(0), "T") & ",null"
                        
            NumRegElim = NumRegElim + 1
            C = " (" & vUsu.Codigo & "," & NumRegElim & "," & C & ")"
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`) VALUES " & C
            conn.Execute C
        

        
        End If
        
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    RS.Close
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Vamos a cargar el LISTVIEW
    lblI.Caption = "Cargando datos"
    lblI.Refresh
    
    
    ListView1.ColumnHeaders.Clear
    
    If NumRegElim > 0 Then
        'OK. TOdo perfecto
        lblError.Caption = "Errores"
        lblError.ForeColor = vbRed
        I = 3
        Cad = "Albaran|Error|Dato ariges|"
        C = "1800|4800|3000|"
    Else
        'Han habido errorres
        lblError.Caption = "Datos"
        lblError.ForeColor = vbBlue
        Cad = "Proveedor|Albaran|Refer|Cantidad|precio|Importe|"
        C = "4229|1154|1379|900|1200|1005|"
        I = 6
    End If
    
    For J = 1 To I
        ListView1.ColumnHeaders.Add , , RecuperaValor(Cad, J), Val(RecuperaValor(C, J)), IIf(J >= 4, lvwColumnRight, lvwColumnLeft)
        
    Next J
    
    
    
    CuadrarDatos = True
    
    
    
    
eCuadrarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing
End Function



Private Sub CargaErrores()
    
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    Cad = "Select * from tmpinformes where codusu =" & vUsu.Codigo & " ORDER BY codigo1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        ListView1.ListItems.Add , , miRsAux!nombre2
        ListView1.ListItems(NumRegElim).SubItems(1) = miRsAux!nombre1
        ListView1.ListItems(NumRegElim).SubItems(2) = DBLet(miRsAux!nombre3, "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
End Sub


Private Sub CargaDatos()
Dim Importe As Currency
Dim Total  As Currency

    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    Cad = "select numalbar,slialp.codartic,cantidad,slialp.codprove from slialp,sartic where sartic.codartic=slialp.codartic and codfamia=100"
    RS.Open Cad, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    ListView1.ListItems.Clear
    
    Cad = "Select  `" & Combo1(1).Text & "` ,`" & Combo1(2).Text & "` "
    Cad = Cad & " , `" & Combo1(3).Text & "` ,`" & Combo1(4).Text & "` "
    Cad = Cad & " FROM " & Combo1(0).Text
    miRsAux.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    NumRegElim = 0
    Total = 0
    
    Me.Refresh
    DoEvents
    
    
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
    
        RS.Find "numalbar = " & DBSet(miRsAux.Fields(1), "T"), , adSearchForward, 1
        
        'NO PUEDE FALLAR , ya hemos comprobado antes
        Cad = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", RS!codProve)
        'Cad = "Proveedor|Albaran|Refer|Cantidad|precio|Importe|"
        ListView1.ListItems.Add , , Cad
        ListView1.ListItems(NumRegElim).SubItems(1) = miRsAux.Fields(1)
        ListView1.ListItems(NumRegElim).SubItems(2) = RS!codartic
        ListView1.ListItems(NumRegElim).SubItems(3) = RS!Cantidad
        Importe = Round(miRsAux.Fields(3) / RS.Fields(2), 4)
        ListView1.ListItems(NumRegElim).SubItems(4) = Format(Importe, FormatoPrecio)
        ListView1.ListItems(NumRegElim).SubItems(5) = Format(miRsAux.Fields(3), FormatoImporte)
        Total = Total + miRsAux.Fields(3)

        miRsAux.MoveNext
        If (NumRegElim Mod 150) = 0 Then
            Screen.MousePointer = vbHourglass
            ListView1.Refresh
            DoEvents
        End If
    Wend

    miRsAux.Close
    RS.Close
    
    lblTotal.Caption = "Total imponible: " & Format(Total, FormatoImporte)
    
    Set RS = Nothing
    Set miRsAux = Nothing
    

End Sub



Private Function RealizaActualizacion() As Boolean
    Dim Importe As Currency
    
    
    On Error GoTo eRealizaActualizacion
    
    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    
    lblTotal.Caption = "Preparando datos actualizacion"
    lblTotal.Refresh
    Cad = "select numalbar,slialp.codartic,cantidad,slialp.codprove,numlinea,fechaalb from slialp,sartic where sartic.codartic=slialp.codartic and codfamia=100"
    RS.Open Cad, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
   
    
    Cad = "Select  `" & Combo1(1).Text & "` ,`" & Combo1(2).Text & "` "
    Cad = Cad & " , `" & Combo1(3).Text & "` ,`" & Combo1(4).Text & "` "
    Cad = Cad & " FROM " & Combo1(0).Text
    miRsAux.Open Cad, Acces, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    
  
    
    
    While Not miRsAux.EOF
    
    
        lblTotal.Caption = "Albaran " & miRsAux.Fields(1)
        lblTotal.Refresh
    
        RS.Find "numalbar = " & DBSet(miRsAux.Fields(1), "T"), , adSearchForward, 1
        
        'NO PUEDE FALLAR , ya hemos comprobado antes
        Importe = Round2(miRsAux.Fields(3), 2)
        
        Cad = "UPDATE slialp SET precioar=" & DBSet(Round2(Importe / RS.Fields(2), 4), "N") & ", importel=" & DBSet(Importe, "N")
        Cad = Cad & " WHERE codprove=" & RS!codProve & " AND numalbar=" & DBSet(RS!NumAlbar, "T")
        Cad = Cad & " And Fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND numlinea=" & RS!numlinea
        conn.Execute Cad
        
        miRsAux.MoveNext
    Wend
    RS.Close
    miRsAux.Close
    
    
    'Grabo log
    Set LOG = New cLOG
    Cad = "BD: " & Text1.Text & vbCrLf & "Tabla: " & Combo1(0).Text & vbCrLf
    Cad = Cad & "Oliva: " & Combo1(1).Text & vbCrLf
    Cad = Cad & "Albaran: " & Combo1(2).Text & vbCrLf
    Cad = Cad & "Cantidad: " & Combo1(3).Text & vbCrLf
    Cad = Cad & "Importe: " & Combo1(4).Text
    
    LOG.Insertar 15, vUsu, Cad
    Set LOG = Nothing
    
    
    MsgBox "Proceso finalizado correctamente", vbInformation
    RealizaActualizacion = True
    
eRealizaActualizacion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Cad
    Set RS = Nothing
    Set miRsAux = Nothing
End Function
