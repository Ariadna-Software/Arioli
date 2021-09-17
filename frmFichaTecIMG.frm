VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFichaTecIMG_ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha técnica. IMAGENES"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar cambios"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1320
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nueva imagen"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar imagen"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Subir"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bajar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "imagen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "desdeBD"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "pathCompleto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "codigoBD"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Lista imágenes"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblCarga2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11535
   End
   Begin VB.Shape Shape1 
      Height          =   7215
      Left            =   5400
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   7020
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4963
   End
   Begin VB.Label Label1 
      Caption         =   "Imagen"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmFichaTecIMG_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const CarpetaIMG = "ImgFicFT"
Public vDatos As String 'codartic|nomartic|


Public EsArticulo As Boolean

Dim InsertandoImg As Boolean
Dim PrimeraVez As Boolean


Dim It As ListItem
Dim contador As Integer


Private Sub InsertarDesdeFichero()
Dim Cadena As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer


    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos Jpg|*.jpg|Archivos Png|*.png|Archivos TIFF|*.tif"
    cd1.ShowOpen
    cd1.MaxFileSize = 256
    cd1.CancelError = False
    
    If cd1.FileName = "" Then Exit Sub
    
    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass
    InsertandoImg = True
    
    J = InStr(1, cd1.FileName, Chr(0))
    Cadena = cd1.FileName
    If J = 0 Then
        'Solo hay un archivo es decir c:\..\eje.txt
        AnyadirAlListview Cadena, False
        
    Else
        Carpeta = Mid(Cadena, 1, J - 1)
        If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta & "\"
        Cadena = Mid(Cadena, J + 1)
        
        While Cadena <> ""
            'Empiezo por la derecch
            J = InStrRev(Cadena, Chr(0))
            If J = 0 Then
                Aux = Cadena
                Cadena = ""
            Else
                Aux = Mid(Cadena, J + 1)
                Cadena = Mid(Cadena, 1, J - 1)
            End If
            AnyadirAlListview Carpeta & Aux, False
        Wend
    End If
    
    
    
    'La ultima es la que voy a previsualizar
    J = lw1.ListItems.Count
    Set lw1.SelectedItem = lw1.ListItems(J)
    Cadena = lw1.ListItems(J).SubItems(2)
    'Cargamos la imagene ''''PREVISUALIZAR
        
    CargarIMG (Cadena)
    cmdGuardar.visible = True
    InsertandoImg = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
Dim J As Integer
Dim Aux As String
    If Dir(vpaz, vbArchive) = "" Then
        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
        'List1.AddItem vpaz
        Set It = lw1.ListItems.Add()
        It.SmallIcon = 23
        
        If DesdeBD Then
            J = InStrRev(vpaz, "\") + 1
            Aux = Mid(vpaz, J)
            It.Text = "Código " & Aux
            If Not IsNumeric(Aux) Then It.SmallIcon = 9
            It.SubItems(3) = Aux
                
        Else
            contador = contador + 1
            It.Text = "Nuevo " & contador
        End If
        
        It.SubItems(1) = Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        It.SubItems(2) = vpaz
        
        Set It = Nothing
    End If
End Sub

Private Function CargarIMG(Archivo As String) As Boolean
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    lblCarga2.Caption = "Cargando ..."
    lblCarga2.Refresh
    CargarIMG = False
    Me.Image1.Picture = LoadPicture(Archivo)

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    Else
        CargarIMG = True
    End If
    lblCarga2.Caption = lblCarga2.Tag
    Screen.MousePointer = vbDefault
End Function

Private Sub MoverItem(Index As Integer)
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim Seleccionado As Integer

    If lw1.SelectedItem Is Nothing Then Exit Sub
    'Si no hay, o hay uno nos salimos
    'If List1.ListCount <= 1 Then Exit Sub
    If lw1.ListItems.Count <= 1 Then Exit Sub

    
    Seleccionado = lw1.SelectedItem.Index
    
    
    'Subir // BAJAR
    If Index = 1 Then
        'SUBIR
        
        'ESTAMOS EN EL ULTIMO
        If Seleccionado >= lw1.ListItems.Count Then Exit Sub
        I = Seleccionado + 1
        
    Else
        'ESTAMOS EN EL primero
        If Seleccionado = 1 Then Exit Sub
        I = Seleccionado - 1

    End If
    InsertandoImg = True  'para que no recarge
    
   
    Aux = lw1.ListItems(I).Text & "|" & lw1.ListItems(I).SubItems(1) & "|" & lw1.ListItems(I).SubItems(2) & "|" & lw1.ListItems(I).SubItems(3) & "|"
    lw1.ListItems(I).Text = lw1.SelectedItem.Text
    lw1.ListItems(I).SubItems(1) = lw1.SelectedItem.SubItems(1)
    lw1.ListItems(I).SubItems(2) = lw1.SelectedItem.SubItems(2)
    lw1.ListItems(I).SubItems(3) = lw1.SelectedItem.SubItems(3)
         
    lw1.ListItems(Seleccionado).Text = RecuperaValor(Aux, 1)
    lw1.ListItems(Seleccionado).SubItems(1) = RecuperaValor(Aux, 2)
    lw1.ListItems(Seleccionado).SubItems(2) = RecuperaValor(Aux, 3)
    lw1.ListItems(Seleccionado).SubItems(3) = RecuperaValor(Aux, 4)
    
    Set lw1.SelectedItem = lw1.ListItems(I)
    
    lw1.SetFocus
    
    InsertandoImg = False
    cmdGuardar.visible = True
    
    
    
End Sub

Private Sub EliminarIMg()
Dim C As String
Dim I As Integer
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub
     
    C = "Seguro que desea eliminar de la lista la imagen: " & lw1.SelectedItem.Text & "?"
    If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    
    
    I = lw1.SelectedItem.Index
    
    lw1.ListItems.Remove I
    If lw1.ListItems.Count = 0 Then
        I = 0
    Else
        If lw1.SelectedItem Is Nothing Then
            I = 0
        Else
            I = 1
        End If
    End If
        
        
    
    
    If I = 0 Then
        CargarIMG ""
    Else
        'i = i - 1
        'If i < 0 Then i = 0
        'If i >= List1.ListCount Then i = List1.ListCount - 1
        InsertandoImg = True
        'lw1.SelectedItem
        CargarIMG lw1.SelectedItem.SubItems(2)
        InsertandoImg = False
    End If
    cmdGuardar.visible = True
End Sub


Private Sub cmdGuardar_Click()
Dim RS As ADODB.Recordset
Dim C As String
Dim L As Long
Dim K As Integer
Dim Eliminar As Boolean

    AbrirConexion
    
    C = "Select max(codigo) from sfichtecdocs"
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then L = RS.Fields(0)
    End If
    L = L + 1
    RS.Close
    
    
    'Por si acaso ha borrado alguno haremios un delete
    'para codartic=vdatos y  no este el codigo
    C = ""
    Eliminar = True
    For K = 1 To lw1.ListItems.Count
        If Val(lw1.ListItems(K).SubItems(1)) > 0 Then
            'Por sia caso hubo error en la caerga
            If lw1.ListItems(K).SmallIcon = 9 Then
                Eliminar = False
            Else
                C = C & ", " & lw1.ListItems(K).SubItems(3)
            End If
        End If
    Next
    If Eliminar Then
        If C <> "" Then
            C = Mid(C, 2) 'quito la primera coma
            C = " not codigo IN (" & C & ")"
            C = "codartic = '" & RecuperaValor(vDatos, 1) & "' AND " & C
            C = "DELETE FROM sfichtecdocs WHERE " & C
            conn.Execute C
        
        End If
    End If
    If lw1.ListItems.Count = 0 Then
        'Ha quitado todos
        C = "codartic = '" & RecuperaValor(vDatos, 1) & "'"
        C = "DELETE FROM sfichtecdocs WHERE " & C
        conn.Execute C
        Me.cmdGuardar.visible = False
        Exit Sub
    End If
    For K = 1 To lw1.ListItems.Count
        
        If Val(lw1.ListItems(K).SubItems(1)) > 0 Then
            C = "UPDATE sfichtecdocs set orden=" & K
            C = C & " WHERE codigo =" & lw1.ListItems(K).SubItems(3)
            conn.Execute C
            
        Else
            'ES NUEVO
            C = "Insert into sfichtecdocs(codigo,codartic,orden) VALUES (" & L & ",'" & RecuperaValor(vDatos, 1) & "'," & K & ")"
            conn.Execute C
            Espera 0.2
            
            'Abro parar guardar el binary
            C = "Select * from sfichtecdocs where codigo =" & L
            Adodc1.ConnectionString = conn
            Adodc1.RecordSource = C
            Adodc1.Refresh
'
            If Adodc1.Recordset.EOF Then
                'MAAAAAAAAAAAAL

            Else
                'Guardar
                 InsertandoImg = True
                CargarIMG lw1.ListItems(K).SubItems(2)
                GuardarBinary Adodc1.Recordset!campo, lw1.ListItems(K).SubItems(2)
                Adodc1.Recordset.Update
            End If

            L = L + 1
        End If
    Next

    Me.cmdGuardar.visible = False
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        ProcesarCarpetaImagenes
        
        CargarArchivos
        
        
        lblCarga2.Caption = lblCarga2.Tag
        
        cmdGuardar.visible = False
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Icon = frmppal.Icon
    PrimeraVez = True
      ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 3   'Botón
        .Buttons(2).Image = 5   'Botón
        .Buttons(5).Image = 7   'Botón
        .Buttons(6).Image = 8  'Botón
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    
    If EsArticulo Then
        'lblCarga2.BackColor = vbWindowBackground
        'lblCarga2.ForeColor = vbBlack
    Else
        lblCarga2.BackColor = vbBlue
        lblCarga2.ForeColor = vbWhite
    End If
    
    
    lw1.ColumnHeaders(1).Width = lw1.Width - 520
    Set lw1.SmallIcons = frmppal.ImgListPpal
    Me.lblCarga2.Tag = RecuperaValor(Me.vDatos, 2)
    If Not EsArticulo Then lblCarga2.Tag = " Categoría: " & lblCarga2.Tag
    
    lblCarga2.Caption = "Leyendo datos BD"
End Sub




Private Sub ProcesarCarpetaImagenes()
Dim C As String
    On Error GoTo EProcesarCarpetaImagenes
    C = App.Path & "\" & CarpetaIMG
    If Dir(C, vbDirectory) = "" Then
        MkDir C
    Else
        If Dir(C & "\*.*", vbArchive) <> "" Then Kill C & "\*.*"
    End If
    
    
    


    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
End Sub



Private Sub CargarArchivos()
Dim C As String
Dim L As Long

    C = "Select * from sfichtecdocs where codartic='" & RecuperaValor(vDatos, 1) & "' ORDER BY orden"
    Me.lblCarga2.Caption = "Leyendo desde BD "
    Me.lblCarga2.Refresh
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = C
    Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
        'NO HAY NINGUNA
    
    Else
        'LEEMOS LAS IMAGENES
        InsertandoImg = True
        While Not Adodc1.Recordset.EOF
            L = Adodc1.Recordset!Codigo
            Me.lblCarga2.Caption = "Leyendo desde BD " & L & "       " & Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
            lblCarga2.Refresh
            C = App.Path & "\" & CarpetaIMG & "\" & L
            If LeerBinary(Adodc1.Recordset!campo, C) Then AnyadirAlListview C, True
            
            Adodc1.Recordset.MoveNext
        Wend
    
    
        
        InsertandoImg = False
        If lw1.ListItems.Count > 0 Then CargarIMG lw1.ListItems(1).SubItems(2)
    End If

    Set Adodc1.Recordset = Nothing
End Sub

Private Sub lw1_Click()
    
    If InsertandoImg Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub
    CargarIMG lw1.SelectedItem.SubItems(2)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'añadir imagen
        InsertarDesdeFichero
        
        
    Case 2
        'quitar img
        EliminarIMg
    Case 5, 6
        'Subir
        MoverItem Button.Index - 5
    Case 10
        Imprimir
        
    Case 11
        'if ha cambiado....
        If cmdGuardar.visible Then
            If MsgBox("No ha guardado los cambios. Desea salir?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        Unload Me
    End Select
End Sub


Private Sub Imprimir()


        With frmImprimir
            If EsArticulo Then
                .NombreRPT = "morImgArt.rpt"
                .FormulaSeleccion = "{sartic.codartic}=""" & RecuperaValor(vDatos, 1) & """"
            Else
                .NombreRPT = "morImgCate.rpt"
                .FormulaSeleccion = "{sfichtecdocs.codartic}= """ & RecuperaValor(vDatos, 1) & """"
            End If
            .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
            .Titulo = "Imágenes adjuntas"
            .NumeroParametros = 1
            .SoloImprimir = False
            .EnvioEMail = False
            
            .Opcion = 2015
            .Show vbModal
        End With
End Sub
