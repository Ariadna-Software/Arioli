VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComImprimirEtiquetas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion etiquetas albaranes proveedor"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmComImprimirEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAyuda 
      Height          =   375
      Left            =   6360
      Picture         =   "frmComImprimirEtiquetas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CheckBox chkUltimas 
      Caption         =   "Ultimas"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkImpresoraEtiquetas 
      Caption         =   "Impresora RED"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox chkImprEtiqueAlba 
      Caption         =   "Imprime etiqueta albarán"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   240
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmComImprimirEtiquetas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vTexto As String  'Texto a mostrar
Public GuardarEImprimir As Boolean

Dim PrimeraVez As Boolean

Dim gridCargado As Boolean
Dim HaModificado2 As String







Private Sub chkImpresoraEtiquetas_KeyPress(KeyAscii As Integer)
      KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub chkUltimas_KeyPress(KeyAscii As Integer)
      KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmdAyuda_Click()
Dim C As String

    C = "Cuando se desee imprimir unas etiquetas más para bultos que se han quedado"
    C = C & vbCrLf & "sin etiquetar, modificaremos la cantidad de etiq. poniendo una mayor con "
    C = C & vbCrLf & "la suma de las que habian mas las nuevas, y marcaremos el check de 'ultimas' " & vbCrLf
    
    MsgBox C, vbInformation
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Cuantos As Integer
Dim C As String
Dim Cp As cPartidas
Dim ColPartidas As Collection
Dim i As Integer
Dim Diferencia As Integer   'por si quiere esta marcado imprimir las utlimas
Dim CuantasCajas As Integer
    
    If Index = 0 Then
        Set miRsAux = New ADODB.Recordset
        If HaModificado2 <> "" Then
            C = "Ha modificado la cantidad de etiquetas. Se guardaran los cambios. " & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then
                'Unload Me
                Exit Sub
            End If
            
            Diferencia = 0
            vTexto = MontaSQLCarga2 & " ORDER BY numlinea"
            miRsAux.Open vTexto, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                Cuantos = miRsAux!Cantidad
                vTexto = MontaWHERE(miRsAux)
                If InStr(1, HaModificado2, vTexto) > 0 Then
                    Set Cp = New cPartidas
  
                    
                    
                    If Cp.LeerDesdeArticulo(miRsAux!codArtic, CInt(miRsAux!codalmac), miRsAux!Numlotes) Then
                                          
                        If Me.chkUltimas.Value = 1 Then
                            Diferencia = Cp.CuantasEtiquetas
                            Diferencia = Cuantos - Diferencia
                        End If
                    
                    
                        Cp.EstablecerEtiquetas Cuantos
                    Else
                        MsgBox "Error leyendo Partida modificada", vbExclamation
                    End If
                    Set Cp = Nothing
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
        
        
        Screen.MousePointer = vbHourglass
        Conn.Execute "DELETE from tmppartidas where codusu = " & vUsu.Codigo
        
        
        vTexto = MontaSQLCarga2 & " ORDER BY numlinea"
        miRsAux.Open vTexto, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set ColPartidas = New Collection
        Set Cp = New cPartidas
        NumRegElim = 0
        While Not miRsAux.EOF
            If Not Cp.LeerDesdeArticulo(miRsAux!codArtic, CInt(miRsAux!codalmac), miRsAux!Numlotes) Then
                'No lee la partida
                MsgBox "Error leyendo idPart/Lote", vbExclamation
                
            Else
                Cuantos = miRsAux!Cantidad
                If Cuantos > 0 Then ColPartidas.Add Cp.IdPartida
                NumRegElim = NumRegElim + Cuantos
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set Cp = Nothing
        
        
        If NumRegElim = 0 Then
            MsgBox "Cantidad=0", vbExclamation
            Exit Sub
        End If
            
        NumRegElim = 0
        'Para cada PARTIDA veo en partidaslineas las etiquetas que hay y las imprimo
        If ColPartidas.Count = 0 Then
            MsgBox "Ningun dato generado", vbExclamation
            Exit Sub
        Else
            Set Cp = New cPartidas
            For i = 1 To ColPartidas.Count
                    Cp.Leer CLng(ColPartidas(i))
                    vTexto = "Select * from spartidaslin where id = " & Cp.IdPartida
                    CuantasCajas = Cp.CuantasEtiquetas
                    
                    vTexto = vTexto & " ORDER BY bulto"
                    If Me.chkUltimas.Value = 1 Then
                        If Diferencia > 0 Then vTexto = vTexto & " desc limit 0," & Diferencia
                    End If
                    miRsAux.Open vTexto, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    C = ""
                    vTexto = ""
                    Cuantos = 0
                    While Not miRsAux.EOF
                        NumRegElim = NumRegElim + 1 'Pa saber que ha impreso
                        
                        '(`idpartida`,`codusu`,`codartic`,`numlote`,referencia,idReferencia,fecha,,idOperacion,idNumOperacion)
                        '(2,22000,'003200190111','PR127','TAPON VERDE MAR 20 ML.',0,'000002001')
                        If vTexto = "" Then
                            vTexto = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Cp.codArtic, "T")
                            vTexto = DBSet(Cp.codArtic, "T") & "," & DBSet(Cp.Numlote, "T") & "," & DBSet(vTexto, "T")
                            vTexto = ", (" & Cp.IdPartida & "," & vUsu.Codigo & "," & vTexto & ","
                            vTexto = vTexto & Cp.codProve & "," & DBSet(Cp.Fecha, "F") & "," & CuantasCajas & ","
                        End If
                            
                        If chkImprEtiqueAlba.Value Then
                            'Imprime una etiqueta CERO del albaran
                            If Cuantos = 0 Then
                                C = C & vTexto & "'ALBARAN','" & Format(Cp.IdPartida, "000000") & "000',0)"
                                Cuantos = 1
                            End If
                        End If
                        
                        
                        C = C & vTexto & "'','" & Format(Cp.IdPartida, "000000") & Format(miRsAux!bulto, "000") & "'," & miRsAux!bulto & " )"
                        
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
                    
                    
                    C = Mid(C, 2) 'quito la primera coma
                    vTexto = "insert into `tmppartidas` (`idpartida`,`codusu`,`codartic`,`numlote`,referencia,idReferencia,fecha,abs_cantidad,idOperacion,idNumOperacion,cantidad) VALUES " & C
                    Conn.Execute vTexto
            Next i
                
        End If
        
        
        Set miRsAux = Nothing
            
        If NumRegElim > 0 Then
        
                'Si imprime con el commander de bartender o imprime en local. Hara unas cosas u otras
                
                If chkImpresoraEtiquetas.Value = 1 Then
                    'Esta imprime por comandos de BARTENDER
                    ImprimeEtiquetasMateriaAuxiliar
                Else
                     'Esta es un rpt normal y corriente
                     With frmImprimir
                        .FormulaSeleccion = "{tmppartidas.codusu} = " & vUsu.Codigo
                        .OtrosParametros = ""
                        .NumeroParametros = 0
                        .SoloImprimir = False
                        .EnvioEMail = False
                        .Opcion = 2014
                        .Titulo = "Etiq. recepcion lotes"
                        .NombreRPT = "EtiquetaRecep.rpt"
                        .ConSubInforme = False
                        .Show vbModal
                    End With
                End If
        Else
            Exit Sub
        End If
            

    End If
    
    Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not gridCargado Then Exit Sub 'Esta gargando datos entodavia
    
    If (Not Data1.Recordset.BOF) And (Not Data1.Recordset.EOF) Then
        If gridCargado Then
            If GuardarEImprimir Then CargaTxtAux True, True
          
        End If
    Else
        Data1.Recordset.MoveLast
    End If
    
    
    
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Label1.Caption = vTexto
        Label1.Refresh
        PonerFocoBtn Command1(0)
        CargaGrid
        Me.txtAux.visible = GuardarEImprimir
        If GuardarEImprimir Then
            CargaTxtAux True, True
            PonerFoco txtAux
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    PrimeraVez = True
    HaModificado2 = ""
    chkImprEtiqueAlba.visible = Not GuardarEImprimir
End Sub


Private Sub CargaGrid()
Dim i As Byte
Dim SQL As String
    
    On Error GoTo ECarga


    gridCargado = False
    
    SQL = MontaSQLCarga2()
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
 
        
    DataGrid1.Columns(0).visible = False 'codusu
    DataGrid1.Columns(1).visible = False 'numalbar
    DataGrid1.Columns(2).visible = False 'fecha albar
    DataGrid1.Columns(3).visible = False 'codprove
    DataGrid1.Columns(4).visible = False 'codalmac

    
    DataGrid1.Columns(5).Caption = "NºLin."
    DataGrid1.Columns(5).Width = 600
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(5).NumberFormat = "#" & " "
    
    DataGrid1.Columns(6).Caption = "Cod.Art."
    DataGrid1.Columns(6).Width = 1700
    
    DataGrid1.Columns(7).Caption = "Desc. Articulo"
    DataGrid1.Columns(7).Width = 3200
       
    DataGrid1.Columns(8).Caption = "Nº Lote"
    DataGrid1.Columns(8).Width = 900
       
    DataGrid1.Columns(9).Caption = "Cantid."
    DataGrid1.Columns(9).Width = 800
    DataGrid1.Columns(9).Alignment = dbgRight
    
       
    DataGrid1.Columns(10).Caption = "Etiquetas"
    DataGrid1.Columns(10).Width = 800
    DataGrid1.Columns(10).Alignment = dbgRight
    DataGrid1.Columns(10).NumberFormat = "# "
    
    
     
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    
    gridCargado = True
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Function MontaSQLCarga2() As String
Dim SQL As String

     SQL = "SELECT codusu,numalbar,fechaalb,tmpnlotes.codprove,codalmac,numlinea,tmpnlotes.codartic,sartic.nomartic,"
                           'lleva la cantidad  etiquetas
     SQL = SQL & "numlotes,tmpnlotes.nomartic,cantidad FROM "
     SQL = SQL & " tmpnlotes,sartic WHERE tmpnlotes.codartic=sartic.codartic AND  codusu=" & vUsu.Codigo
     

     MontaSQLCarga2 = SQL
End Function


Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            txtAux.Text = DBLet(Data1.Recordset!Cantidad, "N")
            txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(10).Left + DataGrid1.Left + 10 'Nº Lotes
        txtAux.Width = DataGrid1.Columns(10).Width - 10
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux.visible = visible
    
    'PonerFoco txtAux
End Sub


Private Sub txtAux_GotFocus()
    txtAux.SelStart = 0
    txtAux.SelLength = Len(txtAux.Text)
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    txtAux_LostFocus
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAux True, True
                End If
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                    txtAux_LostFocus
                    DataGrid1.Row = DataGrid1.Row + 1
                    CargaTxtAux True, True
                Else
                    PonerFocoBtn Me.Command1(0)
                End If
    End Select
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtAux_KeyDown 40, 0
End Sub

Private Sub txtAux_LostFocus()

    txtAux.Text = Trim(txtAux.Text)
'    If Screen.ActiveControl.Name = "cmdAceptar" Then Exit Sub
    
    
    If txtAux.Text <> "" Then
        If PonerFormatoEntero(txtAux) Then
            GuardarLinea
        Else
            txtAux.Text = "0"
            PonerFoco txtAux
        End If
    Else
        MsgBox "Cantidad etiquetas debe tener valor", vbInformation
        PonerFoco txtAux
    End If
End Sub


Private Sub GuardarLinea()

'    If Data1.Recordset.EOF = True Then PrimeraLin = True
     If ActualizarLinea Then

        NumRegElim = Data1.Recordset.AbsolutePosition
        
        CargaGrid

        If SituarDataPosicion(Data1, NumRegElim, "") Then
'            Data1.Recordset.MoveNext
        End If
    End If
End Sub

Private Function MontaWHERE(ByRef R As ADODB.Recordset) As String
    MontaWHERE = " WHERE codusu=" & R!CodUsu & " AND numalbar=" & DBSet(R!NumAlbar, "T") & " AND fechaalb=" & DBSet(R!FechaAlb, "F")
    MontaWHERE = MontaWHERE & " AND codprove=" & R!codProve & " AND numlinea=" & R!numlinea
    'MontaWHERE = " WHERE codusu=" & data1.Recordset!CodUsu & " AND numalbar=" & DBSet(data1.Recordset!NumAlbar, "T") & " AND fechaalb=" & DBSet(data1.Recordset!FechaAlb, "F")
    'MontaWHERE = MontaWHERE & " AND codprove=" & data1.Recordset!codProve & " AND numlinea=" & data1.Recordset!numlinea

End Function



Private Function ActualizarLinea() As Boolean
Dim SQL As String

'    If Not DatosOkLinea Then Exit Function
    
    On Error GoTo ErrActLinea
    
'    Conn.BeginTrans
    
    
    If txtAux.Text = "" Then txtAux.Text = "0"
    
    If Val(Data1.Recordset!Cantidad) <> Val(txtAux.Text) Then
        If Val(Data1.Recordset!Cantidad) > Val(txtAux.Text) Then
            MsgBox "No se puede guardar menos etiquetas de las que habian", vbExclamation
            txtAux.Text = Data1.Recordset!Cantidad
        Else
    
            SQL = MontaWHERE(Data1.Recordset)
            If InStr(1, HaModificado2, SQL) = 0 Then HaModificado2 = HaModificado2 & SQL & "|"
            
            
            SQL = "UPDATE tmpnlotes SET cantidad=" & DBSet(txtAux.Text, "N") & "  " & SQL
            
            
            Conn.Execute SQL
            
            If Val(Data1.Recordset!Cantidad) < Val(txtAux.Text) Then
                Me.chkUltimas.visible = True
                PonerFocoChk chkUltimas
            End If
        End If
            
    End If

     
    ActualizarLinea = True
    Exit Function

ErrActLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando linea.", Err.Description
    End If
'    If b Then
'        Conn.CommitTrans
'    Else
'        Conn.RollbackTrans
'    End If
    ActualizarLinea = False
End Function




