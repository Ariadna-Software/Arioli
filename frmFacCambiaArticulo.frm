VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacCambiaArticulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio referencia"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkKilos 
      Caption         =   "Kilos en M.P."
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   2520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCantidad 
      Height          =   375
      Left            =   4560
      Picture         =   "frmFacCambiaArticulo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Modificar cantidad"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmFacCambiaArticulo.frx":0A02
      Left            =   1200
      List            =   "frmFacCambiaArticulo.frx":0A0C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtDescArticulo 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtArticulo 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtDescArticulo 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtArticulo 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   2655
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1200
      Picture         =   "frmFacCambiaArticulo.frx":0A27
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDpto 
      AutoSize        =   -1  'True
      Caption         =   "Fechas"
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
      Index           =   28
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3720
      TabIndex        =   15
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "Desde"
      Height          =   195
      Index           =   63
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   465
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "frmFacCambiaArticulo.frx":0AB2
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDpto 
      AutoSize        =   -1  'True
      Caption         =   "Datos"
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
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image imgArticulo 
      Height          =   240
      Index           =   1
      Left            =   1200
      Picture         =   "frmFacCambiaArticulo.frx":0B3D
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Nuevo"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label lblDpto 
      AutoSize        =   -1  'True
      Caption         =   "Artículo"
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
      Index           =   26
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   660
   End
   Begin VB.Image imgArticulo 
      Height          =   240
      Index           =   0
      Left            =   1200
      Picture         =   "frmFacCambiaArticulo.frx":0C3F
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Antiguo"
      Height          =   195
      Index           =   60
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   705
   End
End
Attribute VB_Name = "frmFacCambiaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim IndiceImg As Integer
Dim T As String
'Dim FechaProd As Date

Dim DiferenciaCantidad As Currency   'Para cuand es una variacion sobre un mismo articulo

Private Sub cmdCambiar_Click()
Dim I As Integer




    If Me.txtArticulo(1).Text = "" Or Me.txtArticulo(0).Text = "" Then
        MsgBox "Seleccione el articulos para el cambio", vbExclamation
        Exit Sub
    End If
    
    
    If lw1.Tag = "0" Then
        If Me.txtArticulo(1).Text = Me.txtArticulo(0).Text Then
            MsgBox "Son el mismo articulo", vbExclamation
            Exit Sub
        End If
    End If
    
    T = ""
    For IndiceImg = 1 To lw1.ListItems.Count
        If lw1.ListItems(IndiceImg).Checked Then T = T & "O"
    Next IndiceImg
    
    If T = "" Then
        MsgBox "Seleccione algun dato para actualizar", vbExclamation
        Exit Sub
    End If
    
    T = Len(T)
    If MsgBox("Va a realizar el cambio de " & T & " artículo(s). ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set miRsAux = Nothing
    
    'Comprobamos
    Screen.MousePointer = vbHourglass
    For I = lw1.ListItems.Count To 1 Step -1
        If lw1.ListItems(I).Checked Then
        
            If lw1.Tag = "0" Then
                '*****************************************  FACTURAS
                      'Si comprobar
                      Set miRsAux = New ADODB.Recordset
                      If Comprobar(I) Then
                          Conn.BeginTrans
                
                          If ACtualizarReferenciaFactu(I) Then
                              Conn.CommitTrans
                              lw1.ListItems.Remove I
                              
                              'Aqui deberiamos ofertar NUEVO lotes
                              
                          Else
                              Conn.RollbackTrans
                          End If
                      End If
                      
                      Set miRsAux = Nothing
            Else
                '*****************************************  PRODUCCION
                '1ºComprobar

                

                Set miRsAux = New ADODB.Recordset
                    If Quitarproduccion(I) Then
                
                        Conn.BeginTrans
                                                    
                        If RealizarCambios(I) Then
                            Conn.CommitTrans
                            lw1.ListItems.Remove I
                        Else
                            Conn.RollbackTrans
                        End If

                        
                    End If
                Set miRsAux = Nothing
            End If
        End If
    Next I
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCantidad_Click()
    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    If Not lw1.SelectedItem.Checked Then
        If MsgBox("No esta marcado. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    T = "Desea cambiar la cantidad para la produccion: " & lw1.SelectedItem.Text & " -  " & lw1.SelectedItem.SubItems(1) & "?"
    
    T = InputBox(T, , lw1.SelectedItem.SubItems(4))
    If T = "" Then Exit Sub
    If InStr(1, T, ",") = 0 Then
        T = Format(T, FormatoPrecio)
    End If
    If IsNumeric(T) Then
    
        If InStr(1, T, "-") Then
            MsgBox "Campo no puede ser negativo", vbExclamation
        Else
            lw1.SelectedItem.SubItems(4) = T
        End If
    Else
        MsgBox "Campo NO numerico", vbExclamation
    End If
    
        
    
End Sub

Private Sub Combo1_Click()
    cmdCantidad.visible = Me.Combo1.ListIndex = 1
    chkKilos.visible = Me.Combo1.ListIndex = 1
    Datos
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Form_Activate()
    If IndiceImg < 0 Then
        'Primera vez
        IndiceImg = 0
        limpiar Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    IndiceImg = -1
End Sub

Private Sub Datos()
    
    If Combo1.ListIndex < 0 Then Exit Sub
    Set miRsAux = New ADODB.Recordset
    CargaColumnas Combo1.ListIndex
    CargaDatosLW
    Set miRsAux = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    txtArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgArticulo_Click(Index As Integer)
    IndiceImg = Index
    Set frmMtoArticulos = New frmAlmArticulos
    frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
End Sub


Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)

    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    
    T = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T")
    
    If T = "" Then
        MsgBox "No existe el articulo: " & Me.txtArticulo(Index).Text, vbExclamation
        Me.txtArticulo(Index).Text = ""
        PonerFoco Me.txtArticulo(Index)
    End If
    Me.txtDescArticulo(Index).Text = T
    
    Datos
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
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            PonerFoco txtFecha(Index)
        End If
    End If
    Datos
End Sub

Private Function AñadeFecha(campo As String) As String
    AñadeFecha = ""
    If txtFecha(0).Text <> "" Then AñadeFecha = AñadeFecha & " AND " & campo & " >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then AñadeFecha = AñadeFecha & " AND " & campo & " <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
End Function

Private Sub CargaDatosLW()
Dim IT As ListItem
    If Me.Combo1.ListIndex = 0 Then
        'Cargamos las facturas que tenga ese articulo entre esas fechas
        T = "Select * from slifac WHERE codartic =" & DBSet(Me.txtArticulo(0).Text, "T")
        T = T & AñadeFecha("fecfactu")
    Else
        'Produccion
        T = "select sordprod.*,sliordpr.* from sordprod,sliordpr  where sordprod.codigo=sliordpr.codigo"
        T = T & " AND fecproduccion >='1900-01-01'"  'YA ESTA PRODUCIDO
        T = T & " AND codartic =" & DBSet(Me.txtArticulo(0).Text, "T")
        T = T & AñadeFecha("fecproduccion")
    End If
    
    lw1.ListItems.Clear
    
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add
        If Me.Combo1.ListIndex = 0 Then
            'codtipom numfactu fecfactu cantidad
            IT.Text = miRsAux!codTipoM
            IT.SubItems(1) = Format(miRsAux!NumFactu, "000000")
            IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!codalmac
            IT.SubItems(4) = Format(miRsAux!Cantidad, FormatoCantidad)
            T = "numfactu = " & miRsAux!NumFactu & " and fecfactu = '" & Format(miRsAux!FecFactu, FormatoFecha) & "'"
            T = T & " AND numalbar = " & miRsAux!NumAlbar & " AND codtipoa ='" & miRsAux!codTipoa & "' and numlinea = " & miRsAux!numlinea
            IT.Tag = T
        Else
            IT.Text = miRsAux!Codigo
            IT.SubItems(1) = Format(miRsAux!fecproduccion, "dd/mm/yyyy")
            IT.SubItems(2) = Format(miRsAux!feccreacion, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!codalmac
            IT.SubItems(4) = Format(miRsAux!Cantidad, FormatoCantidad)
            T = "codigo = " & miRsAux!Codigo
            'T = T & " AND numalbar = " & miRsAux!NumAlbar & " AND codtipoa ='" & miRsAux!codtipoa & "' and numlinea = " & miRsAux!numlinea
            IT.Tag = T
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub



Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader
    If lw1.Tag <> "" Then
        If lw1.Tag = Val(OpcionList) Then Exit Sub
    End If
    Select Case OpcionList
    Case 0
        'Facturas
       ' codtipom numfactu fecfactu cantidad
        Columnas = "Tipo|Numero|fecha|Alm|Cantidad|"
        Ancho = "750|1200|1400|650|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|2|0|0|2|"
        'Formatos
        Formato = "||||" & FormatoPrecio & "|"
        Ncol = 5
    
    Case 1
        'PRECIOS ESPECIALES
        'Label2(0).Caption = "Precios especiales"
        Columnas = "Cod Prod.|Fecha|F. Producion|Alm|Cantidad|"
        Ancho = "1000|1100|1100|650|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "000|||" & FormatoImporte & "|"
        Ncol = 5
  
    End Select
    
  

    'Guardo la opcion en el tag
    'lw1.Tag = OpcionList & "|" & Ncol & "|"
    lw1.Tag = OpcionList
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub




Private Function Comprobar(Indice As Integer) As Boolean
    On Error GoTo EC
    
    Comprobar = False
    
    
    If lw1.Tag = 0 Then
        If ComprobarFacturado(Indice) Then
            Comprobar = True
            
        End If
    Else
    
        Comprobar = True
    
    End If
    
    Exit Function
EC:
    MuestraError Err.Number
End Function



'Comprobar FACTURADO
'Para tirar atras un elemto facturado...
'Tenemos que encontrar la linea de smoval
'correspodiente a este elemento
Private Function ComprobarFacturado(Indice As Integer) As Boolean
Dim Nlinea As Integer
Dim Cad As String
Dim cP As cPartidas

     ComprobarFacturado = False
     
     
    'Si tiene en salmac
    T = "select * from salmac where codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        T = ""
    End If
    miRsAux.Close
    If T = "" Then
        MsgBox "No hay datos en salmac para el articulo"
        Exit Function
     End If
     
     
     
     T = "Select * from slifac WHERE " & lw1.ListItems(Indice).Tag
     miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     T = ""
     If Not miRsAux.EOF Then
        Nlinea = miRsAux!numlinea
        T = "Select * from smoval where detamovi = 'ALV' and document = '" & Format(miRsAux!NumAlbar, "0000000") & "'"
        T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
        T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
        T = T & " AND tipomovi =0 " 'Siempre salida
        T = T & " AND numlinea = " & Nlinea
     End If
     miRsAux.Close
     If T = "" Then
        MsgBox "No se ha encotrado la linea de factura: " & lw1.ListItems(Indice).Text & " - " & Nlinea, vbExclamation
        Exit Function
    End If
     
     miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     T = "EOF"
     If Not miRsAux.EOF Then
        T = "NULL"
        If Not IsNull(miRsAux!detamovi) Then
            T = "NO ALV"
            If miRsAux!detamovi = "ALV" Then T = ""
        End If
    End If
    miRsAux.Close
    
    If T <> "" Then
        MsgBox "Error recuperando datos en smoval. " & T, vbExclamation
        Exit Function
    End If
    
    
    'Comprobamos los LOTES
    '-----------------------------------------------------------------
    T = "Select * from slifaclotes WHERE " & lw1.ListItems(Indice).Tag
     miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     T = ""
     Cad = ""
     If Not miRsAux.EOF Then
        'Ahora vere si el lote lo encuentra
        T = DBLet(miRsAux!Numlote)
        If T = "" Then
            'esto NO deberia pasar
            Cad = "No se ha encontrado el LOTE en lotes/fras(faclotes) para el articulo: " & Me.txtArticulo(0).Text
        Else
            Set cP = New cPartidas
            If Not cP.LeerDesdeArticulo(Me.txtArticulo(0).Text, CInt(lw1.ListItems(Indice).SubItems(3)), T) Then
                Cad = "No se ha encontrado la partida asociada al lote/articulo: " & T & "   /   " & Me.txtArticulo(0).Text
            End If
        End If
        
     Else
        Cad = "No se ha encontrado el LOTE para: " & lw1.ListItems(Indice).Text & lw1.ListItems(Indice).SubItems(1) & vbCrLf
     End If
     miRsAux.Close
     If Cad <> "" Then
        Cad = Cad & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Function
      
    End If
    
    
    
    
    
    
    
    'OK
    ComprobarFacturado = True
        
End Function




Private Function ACtualizarReferenciaFactu(Indice As Integer) As Boolean
Dim Cantidad As Currency
Dim cP As cPartidas

On Error GoTo EACtualizarReferenciaFactu
    ACtualizarReferenciaFactu = False

    'Updateamos smoval
        
     T = "Select * from slifac WHERE " & lw1.ListItems(Indice).Tag
     miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     T = ""
     
     'codartic codalmac fechamov horamovi tipomovi detamovi cantidad impormov codigope letraser document numlinea
        Cantidad = miRsAux!Cantidad
        T = "Select * from smoval where detamovi = 'ALV' and document = '" & Format(miRsAux!NumAlbar, "0000000") & "'"
        T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
        T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
        T = T & " AND tipomovi =0 " 'Siempre salida
        T = T & " AND numlinea = " & miRsAux!numlinea
     
     miRsAux.Close
     Espera 0.2
     
     miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     T = "UPDATE smoval Set codartic = " & DBSet(Me.txtArticulo(1).Text, "T")
     T = T & " WHERE codalmac = " & miRsAux!codalmac
     T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
     T = T & " AND tipomovi =0 " 'Siempre salida
     T = T & " AND numlinea = " & miRsAux!numlinea
     T = T & " AND detamovi = 'ALV' and document = '" & miRsAux!document & "'"
     T = T & " AND Fechamov = '" & Format(miRsAux!Fechamov, FormatoFecha) & "'"
    'Cambio  el articulo en smoval
    Conn.Execute T
    
    
    'Updateamos la linea de factura
    T = "UPDATE slifac set codartic = " & DBSet(Me.txtArticulo(1).Text, "T")
    T = T & ", nomartic = " & DBSet(Me.txtDescArticulo(1).Text, "T") 'nomartic tb lo cambio
    T = T & " WHERE " & lw1.ListItems(Indice).Tag
    Conn.Execute T
    
    
    'Le sumo la cantida al artiuclo 2 y se la resto al articulo 1
    T = "UPDATE salmac set canstock= canstock + " & TransformaComasPuntos(CStr(Cantidad))
    T = T & " WHERE codartic = " & DBSet(Me.txtArticulo(0).Text, "T")
    T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
    Conn.Execute T
    
    T = "UPDATE salmac set canstock= canstock - " & TransformaComasPuntos(CStr(Cantidad))
    T = T & " WHERE codartic = " & DBSet(Me.txtArticulo(1).Text, "T")
    T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
    Conn.Execute T
    
    'LOTES
    '----------------------------------------------------------------------------
      
    'Comprobamos los LOTES
    '-----------------------------------------------------------------
    miRsAux.Close
    T = "Select * from slifaclotes WHERE " & lw1.ListItems(Indice).Tag
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        'Ahora vere si el lote lo encuentra
        T = DBLet(miRsAux!Numlote)
        If T <> "" Then
            Set cP = New cPartidas
            If cP.LeerDesdeArticulo(Me.txtArticulo(0).Text, CInt(lw1.ListItems(Indice).SubItems(3)), T) Then
               cP.IncrementarCantidad Cantidad
            End If
        End If
        
     End If
     miRsAux.Close
     
    
    'Añadimos LOG
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    T = "Articulo en factura" & vbCrLf
    T = T & "Inic.: " & Me.txtArticulo(0).Text & " " & Me.txtDescArticulo(0).Text & vbCrLf
    T = T & "Final: " & Me.txtArticulo(1).Text & " " & Me.txtDescArticulo(1).Text & vbCrLf
    'Que referencia estamos cambiando
    T = T & "Factura : " & lw1.ListItems(Indice).Text & lw1.ListItems(Indice).SubItems(1)
    T = T & " " & lw1.ListItems(Indice).SubItems(2) & "  Cant: " & lw1.ListItems(Indice).SubItems(4)
    LOG.Insertar 8, vUsu, T
    Set LOG = Nothing
    Espera 0.2
    ACtualizarReferenciaFactu = True
    Exit Function
EACtualizarReferenciaFactu:
    MuestraError Err.Number, "Actualizando: " & lw1.ListItems(Indice).SubItems(1)
End Function






'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'   Cambio articulo en produccion
'
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------


Private Function Quitarproduccion(Indice As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim OK As Boolean
Dim Cantidad As Currency
Dim TieneFactor As Boolean


'Comprobara tb si es el mismo articulo si hay cantidad suficinete
Dim MismoArticulo As Currency
Dim CantidadOriginal As Currency
Dim MensajeLote As String
Dim Diferencia As Currency
Dim CuantasLineasProduccion As Integer

Dim cP As cPartidas

    'BORRAR la var
    
    On Error GoTo EQuitarproduccion
    Quitarproduccion = False
    
    
    
    T = "codigo=" & lw1.ListItems(Indice).Text & " AND codartic  "
    'T = "codartic= '000700010701' and codigo=220 and codalmac=1"
    T = DevuelveDesdeBD(conAri, "count(*)", "sliordprlotes", T, txtArticulo(0).Text, "T")
    If T <> "" Then
        If Val(T) > 1 Then
            T = "Existe mas de un numero de lote asignado a este articulo en este parte de produccion(" & T & ")"
            MsgBox T, vbExclamation
            T = ""
            Exit Function
        End If
    End If
    
    
    
    
    
    MismoArticulo = txtArticulo(0).Text = txtArticulo(1).Text
    If Not MismoArticulo Then
        T = "Solo deberia cambiar la cantidad. El proceso de lotaje no se puede reconstruir" & vbCrLf & vbCrLf
        T = T & "¿Continuar?"
        If MsgBox(T, vbQuestion + vbYesNo) = vbNo Then Exit Function
        
        If MsgBox("Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    T = "Select * from smoval where codartic='" & txtArticulo(0).Text & "' and detamovi='PRO' "
    T = T & " AND document = '" & lw1.ListItems(Indice).Text & "'"

    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    T = ""
    If miRsAux.EOF Then
        T = "Error leyendo articulo produccion en movimientos"
    Else
        CantidadOriginal = miRsAux!Cantidad
    End If
    miRsAux.Close
    
    If T <> "" Then
        MsgBox T, vbExclamation
        Exit Function
    End If
    
    Set rs = New ADODB.Recordset
    
    'Para cada SUBCOMPONENTE veremos si localizo el movimiento en smoval
    T = "select sliordpr2.*,sartic.factorconversion,CtrStock,nomartic from sliordpr2,sartic    where"
    T = T & " sliordpr2.codarti2=sartic.codartic "
    T = T & " AND codigo=" & lw1.ListItems(Indice).Text
    T = T & " AND sliordpr2.codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OK = True
    CuantasLineasProduccion = 0
    While Not miRsAux.EOF
        If miRsAux!CtrStock = 1 Then
            Cantidad = DBLet(miRsAux!FactorConversion, "N")
            If Cantidad = 0 Then Cantidad = 1
            TieneFactor = (Cantidad <> 1)
            If TieneFactor Then
                If Me.chkKilos.Value = 1 Then
                    'Mod nuevo. Graba los kilos en los movimientos
                    Cantidad = miRsAux!Cantidad
                Else
                    Cantidad = Round2(Cantidad * miRsAux!Cantidad, 2)
                End If
            Else
                Cantidad = Round2(Cantidad * miRsAux!Cantidad, 2)
            End If
            
            
            T = "Select * from smoval where codartic='" & miRsAux!codarti2 & "' and detamovi='PRO' "
            'Movimiento
            T = T & " AND document = '" & lw1.ListItems(Indice).Text & "'"
            'Y LA CANTIDAD
            If TieneFactor Then
                T = T & " and cantidad between " & TransformaComasPuntos(CStr(Cantidad - 1))
                T = T & " and " & TransformaComasPuntos(CStr(Cantidad + 1))
            Else
                T = T & " AND cantidad = " & TransformaComasPuntos(CStr(Cantidad))
            End If
        
            rs.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If rs.EOF Then
                
                T = miRsAux!codarti2 & "  " & miRsAux!NomArtic & "        Can: " & Cantidad & vbCrLf
                MensajeLote = MensajeLote & T
            Else
                CuantasLineasProduccion = CuantasLineasProduccion + 1
            End If
            rs.Close
            
        End If
        If OK Then miRsAux.MoveNext
        
    Wend
    miRsAux.Close


    If MensajeLote <> "" Then
        MensajeLote = "Error buscando movimientos articulos: " & vbCrLf & vbCrLf & MensajeLote
        MensajeLote = MensajeLote & " ¿Continuar?"
        If MsgBox(MensajeLote, vbQuestion + vbYesNo) = vbNo Then OK = False
    End If
    If Not OK Then Exit Function
    
    
    'Comprobare tb los numeros de LOTE y si se han asignado
    MensajeLote = ""
    
    'Incremento de cantidad
    Diferencia = ImporteFormateado(lw1.ListItems(Indice).SubItems(4))
    Diferencia = Diferencia - CantidadOriginal
    
    If Diferencia = 0 Then
        If MsgBox("NO hay cambio de cantidad. ¿Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    
    MensajeLote = ""
    T = "Select * from sliordprlotes,sartic where sartic.codartic= sliordprlotes.codartic AND sliordprlotes.codartic='" & txtArticulo(0).Text & "'  "
    T = T & " AND codigo = " & lw1.ListItems(Indice).Text

    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    T = ""
    OK = True
    If miRsAux.EOF Then
        T = "Error leyendo articulo produccion en lotes" & vbCrLf
        OK = False
    Else
        If CantidadOriginal <> miRsAux!cantlote Then T = "Cantidades distintas en lotes / produccion" & vbCrLf
        
        Set cP = New cPartidas
        If cP.LeerDesdeArticulo(miRsAux!codartic, miRsAux!codalmac, miRsAux!Numlote) Then
                
            'T = T & "No se ha encotrado la partida asociada al producto: " & miRsAux!codartic & vbCrLf
            If Diferencia < 0 Then
                'Va a quitar cantidad producida. Comprobaremos si aun queda cantidad sin asignar
                Cantidad = cP.Cantidad + Diferencia  '+ pq la cantidad YA es negativa
            
                If Cantidad < 0 Then
                    T = T & "**No queda suficiente cantidad : " & vbCrLf
                    T = T & miRsAux!codartic & " - " & miRsAux!NomArtic & vbCrLf
                    T = T & "       Partida: " & Format(cP.Cantidad, FormatoCantidad) & vbCrLf
                    T = T & "       Mov. Prod: " & Format(miRsAux!cantlote, FormatoCantidad) & vbCrLf
                    T = T & "       Cant. necesaria: " & Format(Abs(Cantidad), FormatoCantidad) & vbCrLf
                End If
            End If
        Else
            T = T & "No se ha encotrado la partida asociada al producto: " & miRsAux!codartic & vbCrLf
        End If
        Set cP = Nothing

    End If
    miRsAux.Close
    
    
    If T <> "" Then MensajeLote = MensajeLote & vbCrLf & T
    If Not OK Then Exit Function  'Si no encuentra la linea de produccion NO continuamos
    
    

    
    T = "select sliordpr2lotes.*,sartic.factorconversion,cantidad,nomartic  from sliordpr2lotes,sarti1,sartic   where"
    T = T & " sliordpr2lotes.codarti2= sarti1.codarti1 and sarti1.codartic='" & Me.txtArticulo(0).Text & "' and"
    T = T & " sliordpr2lotes.codarti2=sartic.codartic "
    T = T & " AND sarti1.codarti1= sartic.codartic"
    T = T & " AND codigo=" & lw1.ListItems(Indice).Text
    T = T & " AND sliordpr2lotes.codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OK = True


    'Lo que para la materia producida, la materia necesari va en signo contrario.
    'Esto es.
    'Si produzco mas disminuyo las lineas(tapones botellas...)
    'si produzco menos aunmento "        "
    Diferencia = -1 * Diferencia
    
    T = ""
    While Not miRsAux.EOF
        
        If Me.chkKilos.Value = 0 Then
            Cantidad = DBLet(miRsAux!FactorConversion, "N")
        Else
            Cantidad = 1
        End If
        If Cantidad = 0 Then Cantidad = 1  'PARA QUE NO DE ERROR
        Cantidad = Round2(Cantidad * miRsAux!Cantidad, 5)
        Cantidad = Round2(Diferencia * Cantidad, 5)
    
        Set cP = New cPartidas
        
        If cP.LeerDesdeArticulo(miRsAux!codarti2, miRsAux!codalmac, miRsAux!Numlote) Then
            
            'Aqui incrementaremos la cantidad, con lo cual No hay que comprobar si
            If Diferencia < 0 Then
            
                Cantidad = cP.Cantidad + Cantidad '+ pq la cantidad YA es negativa
                If Cantidad < 0 Then
                    T = "   -> " & miRsAux!codartic & " - " & miRsAux!NomArtic & vbCrLf
                    T = T & "       Partida: " & Format(cP.Cantidad, FormatoCantidad) & vbCrLf
                    T = T & "       Mov. Prod: " & Format(miRsAux!cantlote, FormatoCantidad) & vbCrLf
                    T = T & "       Cant. necesaria: " & Format(Abs(Cantidad), FormatoCantidad) & vbCrLf
                    MensajeLote = MensajeLote & T
                    
                End If
            End If
        
        
        
        
        Else
            T = T & "No se ha encontrado la Partida para materia prima: " & miRsAux!codartic & " " & miRsAux!Numlote
        End If
        Set cP = Nothing
    
        
    
    
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    
    If MensajeLote <> "" Then
        MensajeLote = MensajeLote & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(MensajeLote, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
        
    
    
    
    
    Quitarproduccion = True
EQuitarproduccion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Quitar produccion"
    Set rs = Nothing
End Function


'OK. Pasos.
' 1 Quitaremos de smvoal los movmientos relacionados
' 2 Quitaremos de slior y sliord2 los datos
' 3 insertaremos en slior y sliord2
' 4 insertaremos en smoval
' 5 actualizaremos sctocks



Private Function RealizarCambios(Indice As Integer) As Boolean
Dim Multiplicador As Currency
Dim rs As ADODB.Recordset
Dim OK As Boolean
Dim Canti As Currency
Dim TieneFactor As Boolean
Dim Trab As Long
Dim F As Date
Dim H As Date
Dim Insert As String
Dim Diferencia2 As Currency
Dim CantiActual As Currency
Dim Aux As Currency
Dim cP As cPartidas
Dim CantidadSlior As Currency

    On Error GoTo EQuitarproduccion
    RealizarCambios = False
    
    
    
    
    Set rs = New ADODB.Recordset

    'Mayo 2010
    'CAmbio. Los cambios de produccion serán SOBRE le mismo articulo
    'Con lo cual ya no tengo que borrar
    'Ahora updear
''''    T = "Select * from sliordpr WHERE  codigo=" & lw1.ListItems(Indice).Text
''''    T = T & " AND sliordpr.codartic = " & DBSet(txtArticulo(0).Text, "T")
''''    T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
''''    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''    'NO PUEDE SER NUL, eso ya lo hemos comprobado arriba
''''    Canti = miRsAux!Cantidad
'    miRsAux.Close
    CantiActual = ImporteFormateado(lw1.ListItems(Indice).SubItems(4))

    'Diferencia tengo la cantidad de diferencia
    'En CantiActual la cantidad actual

    
    
    
    'NOOOOOOOOOOOOOOOOOOOOOOO borro, ahora se UPDATEA
    T = "select sliordpr2.*,sartic.factorconversion,sarti1.cantidad as cuanto from "
    T = T & " sliordpr2,sarti1,sartic where"
    T = T & " sliordpr2.codarti2= sarti1.codarti1 AND sarti1.codartic='" & Me.txtArticulo(0).Text & "'"
    T = T & " AND sliordpr2.codarti2=sartic.codartic "
    T = T & " AND sarti1.codarti1= sartic.codartic"
    T = T & " AND sliordpr2.codarti2=sartic.codartic "
    T = T & " AND codigo=" & lw1.ListItems(Indice).Text
    T = T & " AND sliordpr2.codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OK = True
    While Not miRsAux.EOF
        Canti = DBLet(miRsAux!FactorConversion, "N")
        If Canti = 0 Then Canti = 1
        TieneFactor = (Canti <> 1)
        'Para cuando sea materia prima
        Aux = DBLet(miRsAux!Cantidad, "N")
        If TieneFactor Then
            If Me.chkKilos.Value = 0 Then Aux = Aux * Canti
        
            Multiplicador = Round2(Canti * miRsAux!cuanto, 5)
            
            
            
        Else
            Multiplicador = miRsAux!cuanto
            Aux = Multiplicador * Aux
        End If
        'Para buscar el movimiento en sliord esta AUX
        
        
        Canti = Round2(Multiplicador * CantiActual, 2)
        'Cantidad para la sliordpr
        If TieneFactor Then
                If Me.chkKilos.Value = 1 Then
                    CantidadSlior = Round2(CantiActual * miRsAux!cuanto * miRsAux!FactorConversion, 2)
                Else
                    'Si esta en litros NO tengo k multiplicar por factor conversion
                    CantidadSlior = CantiActual * miRsAux!cuanto
                End If
        Else
            CantidadSlior = CantiActual * miRsAux!cuanto
        End If
        T = "UPDATE smoval "
        T = T & " SET cantidad = " & TransformaComasPuntos(CStr(Canti))  'ahora tendra esta cantidad
        T = T & " where codartic='" & miRsAux!codarti2 & "' and detamovi='PRO' "
        'Movimiento
        T = T & " AND document = '" & lw1.ListItems(Indice).Text & "'"
        'Y LA CANTIDAD
        If TieneFactor Then

            T = T & " and cantidad between " & TransformaComasPuntos(CStr(Aux - 5))
            T = T & " and " & TransformaComasPuntos(CStr(Aux + 5))
        Else
            
            T = T & " AND cantidad = " & TransformaComasPuntos(CStr(Aux))
        End If
        Conn.Execute T
        

        
        
        'Aumentamos en salmac
        'En aux tenemos lo que habia
        Diferencia2 = Aux - Canti
        T = "UPDATE salmac set canstock= canstock  "
        If Diferencia2 >= 0 Then T = T & " + "
        T = T & TransformaComasPuntos(CStr(Diferencia2))
        T = T & " WHERE codartic = " & DBSet(miRsAux!codarti2, "T")
        T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
        Conn.Execute T
    
    
    
        'Borramos de sliordp2
        T = "UPDATE sliordpr2 "
        T = T & " SET cantidad = " & TransformaComasPuntos(CStr(CantidadSlior))
        T = T & " WHERE  codigo=" & lw1.ListItems(Indice).Text
        T = T & " AND sliordpr2.codartic = " & DBSet(txtArticulo(0).Text, "T")
        T = T & " AND codarti2 = " & DBSet(miRsAux!codarti2, "T")
        Conn.Execute T
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    '----------------------------------------------------------------------
    'Los lotes de la materia auxiliar
    T = "select sliordpr2lotes.*,sartic.factorconversion ,sarti1.cantidad as cuanto from sliordpr2lotes,sarti1,sartic    where"
    
    T = T & " sliordpr2lotes.codarti2= sarti1.codarti1 and sarti1.codartic='" & Me.txtArticulo(0).Text & "'"
    T = T & " AND sliordpr2lotes.codarti2=sartic.codartic "
    T = T & " AND codarti2=sartic.codartic "
    T = T & " AND codigo=" & lw1.ListItems(Indice).Text
    T = T & " AND sliordpr2lotes.codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not miRsAux.EOF
    
        
        Canti = DBLet(miRsAux!FactorConversion, "N")
        If Canti = 0 Then Canti = 1
        TieneFactor = (Canti <> 1)
        
    
        
        'Para cuando sea materia prima
        Aux = DBLet(miRsAux!cantlote, "N")
        If TieneFactor Then
            If Me.chkKilos.Value = 1 Then
                'Aux = Aux * Canti
            Else
                
                 Aux = Aux * Canti
            End If
            
        
            Multiplicador = Round2(Canti * miRsAux!cuanto, 5)
            
            
            
        Else
            Multiplicador = miRsAux!cuanto
            Aux = Multiplicador * Aux
        End If
        'Para buscar el movimiento en sliord esta AUX
        
        
        Canti = Round2(Multiplicador * CantiActual, 2)
        'Cantidad para la sliordpr
        
        CantidadSlior = CantiActual * miRsAux!cuanto
        If Me.chkKilos.Value = 1 And TieneFactor Then CantidadSlior = Round2(CantiActual * miRsAux!cuanto * miRsAux!FactorConversion, 2)   'La cantiDAD
        
        
        
        
        'Ahora updateo en sliord2lotes
        T = "UPDATE sliordpr2lotes SET cantlote = " & TransformaComasPuntos(CStr(CantidadSlior))
        T = T & " WHERE  codigo=" & lw1.ListItems(Indice).Text
        T = T & " AND sliordpr2lotes.codartic = " & DBSet(txtArticulo(0).Text, "T")
        T = T & " AND sliordpr2lotes.codarti2 = " & DBSet(miRsAux!codarti2, "T")
        
        Conn.Execute T
        
        T = DBLet(miRsAux!Numlote, "T")
        If T <> "" Then
            Canti = Aux - Canti
         
            Set cP = New cPartidas
            If cP.LeerDesdeArticulo(miRsAux!codarti2, miRsAux!codalmac, T) Then
                cP.IncrementarCantidad Canti
            Else
                MsgBox "No ha encotrado la partida para " & miRsAux!codarti2 & "   " & T
            End If
            Set cP = Nothing
        End If
                
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    '**********************************************************************+
    ' Articulo ppal



    T = "select sliordpr.* from sliordpr where codigo=" & lw1.ListItems(Indice).Text
    T = T & " AND sliordpr.codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Canti = miRsAux!Cantidad
    miRsAux.Close
    
    T = "UPDATE smoval SET cantidad = " & TransformaComasPuntos(CStr(CantiActual))
    T = T & " where codartic='" & txtArticulo(0).Text & "' and detamovi='PRO' "
    T = T & " AND document = '" & lw1.ListItems(Indice).Text & "'"
    Conn.Execute T
    
    T = "UPDATE salmac set canstock= canstock "
    Diferencia2 = CantiActual - Canti   'lo que hay menos lo que habia
    If Diferencia2 >= 0 Then T = T & " + "
    T = T & TransformaComasPuntos(CStr(Diferencia2))
    T = T & " WHERE codartic = " & DBSet(Me.txtArticulo(0).Text, "T")
    T = T & " AND codalmac = " & lw1.ListItems(Indice).SubItems(3)
    Conn.Execute T

    
    T = "UPDATE sliordpr  SET cantidad = " & TransformaComasPuntos(CStr(CantiActual))
    T = T & " Where codigo = " & lw1.ListItems(Indice).Text
    T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
    Conn.Execute T
    
    
    'Las aprtidas
    T = "select * from sliordprlotes "
    T = T & " Where codigo = " & lw1.ListItems(Indice).Text
    T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
    miRsAux.Open T, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set cP = New cPartidas
        T = miRsAux!Numlote
        If cP.LeerDesdeArticulo(miRsAux!codartic, miRsAux!codalmac, T) Then
            'Aqui deberiamos ir descontado de los lotes
            cP.IncrementarCantidad Diferencia2
            
            'Descontamos de un lote
            'updateamos sliorpr2
            T = "UPDATE sliordprlotes SET cantlote = " & TransformaComasPuntos(CStr(CantiActual))
            T = T & " Where codigo = " & lw1.ListItems(Indice).Text
            T = T & " AND codartic = " & DBSet(txtArticulo(0).Text, "T")
            T = T & " AND numlote = " & DBSet(cP.Numlote, "T")
            Conn.Execute T
            
        Else
            MsgBox "Lote materia embasado NO encontrado", vbExclamation
        End If
        Set cP = Nothing
        
        
        
        miRsAux.MoveNext
    Wend
    Set cP = Nothing
    
    



    Set LOG = New cLOG
    T = "Articulo orden de producción nº:" & lw1.ListItems(Indice).Text & "  " & lw1.ListItems(Indice).SubItems(1) & " Can: " & lw1.ListItems(Indice).SubItems(4) & vbCrLf
    T = T & "Inic.: " & Me.txtArticulo(0).Text & " " & Me.txtDescArticulo(0).Text & vbCrLf
    T = T & "Final: " & Me.txtArticulo(1).Text & " " & Me.txtDescArticulo(1).Text & vbCrLf

    'Que referencia estamos cambiando
    LOG.Insertar 8, vUsu, T
    Set LOG = Nothing
    Espera 0.2


    RealizarCambios = True
EQuitarproduccion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Quitar produccion" & Err.Description
    Set rs = Nothing
End Function
    


Private Sub ContruyeSQL_Delete(ByRef R As ADODB.Recordset)
    'delete from `smoval` where `codartic`='003700441513' and `codalmac`='1' and `fechamov`='0000-00-00' and `horamovi`='0000-00-00 00:00:00' and `tipomovi`='0' and `detamovi`='PRO' and `cantidad`='0.00' and `impormov`='0.00' and `codigope`='0' and `letraser` IS NULL and `document`='0' and `numlinea`='0'
    T = "delete from `smoval` where `codartic`='" & R!codartic & "' and `codalmac`=" & R!codalmac
    T = T & " and `fechamov`='" & Format(R!Fechamov, FormatoFecha) & "'  and `tipomovi`=" & R!tipomovi & " and `detamovi`='PRO'"
    T = T & " AND `cantidad`=" & TransformaComasPuntos(CStr(R!Cantidad)) & " and `document`='" & R!document & "' and `numlinea`=" & R!numlinea
End Sub



