VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFacActPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Precios"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmFacActPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
      Begin VB.CheckBox chkPreuEsp 
         Caption         =   "Precios especiales"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkPreuAct 
         Caption         =   "Precios actuales"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblProgreso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   6105
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblProgreso 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Actualizar Tarifas de Precios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de cambio"
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
      Index           =   21
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   1410
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   2355
      Picture         =   "frmFacActPrecios.frx":000C
      Top             =   960
      Width           =   240
   End
End
Attribute VB_Name = "frmFacActPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1


Private menErrProceso As String 'mensaje final del proceso actualizacion de precios

Private Sub chkPreuAct_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
'actualizar los nuevos precios actuales  y/o especiales
Dim cadSel As String
Dim SQL As String
Dim totRegPA As Long 'total registros a cambiar de precios actuales
Dim totRegPE As Long 'total registros a cambia de precios especiales


    '--- COMPROBACIONES DE DATOS
    '-----------------------------
    
    '- comprobar q se ha seleccionado fecha de cambio
    If txtCodigo(0).Text = "" Then
        MsgBox "El campo fecha de cambio debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    '- comprobar que es una fecha valida
    PonerFormatoFecha txtCodigo(0)
    If txtCodigo(0).Text = "" Then
        Exit Sub
    End If
    
    '- comprobar q se ha seleccionado al menos un check
    If Me.chkPreuAct.Value <> 1 And Me.chkPreuEsp <> 1 Then
        MsgBox "Debe seleccionar al menos un precio para actualizar.", vbExclamation
        Exit Sub
    End If
    
    
    
    '--- COMPROBAR Q HAY REGISTROS A PROCESAR
    '------------------------------------------
    
    '- obtener la cadena de seleccion de registros de tarifas de precio q se van
    '    a actualizar: los q cumplan q slista.fechanue <= valor_introducido
    cadSel = "fechanue"
    cadSel = CadenaDesdeHastaBD("", txtCodigo(0).Text, cadSel, "F")
    
    '- comprabar q existen registros para ese criterio de seleccion
    totRegPA = 0
    totRegPE = 0
    If Me.chkPreuAct.Value = 1 Then
        'si marcado actualizar PRECIOS ACTUALES
        SQL = "SELECT COUNT(*) FROM slista WHERE " & cadSel
        totRegPA = TotalRegistros(SQL)
        
        If Not (totRegPA > 0) Then
            If Me.chkPreuEsp.Value = 1 Then
                'comprobar si se actualizar precios especiales
                SQL = "SELECT COUNT(*) FROM sprees WHERE " & cadSel
                totRegPE = TotalRegistros(SQL)
                If Not (totRegPE > 0) Then
                    MsgBox "No hay tarifas de precios ni precios especiales a actualizar para esa fecha.", vbExclamation
                    Exit Sub
                End If
            ElseIf Me.chkPreuEsp.Value <> 1 Then
                'no hay registros a procesar y fin
                MsgBox "No hay tarifas de precios a actualizar para esa fecha.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Me.chkPreuEsp.Value = 1 Then
        'comprobar si se actualizar PRECIOS ESPECIALES
        SQL = "SELECT COUNT(*) FROM sprees WHERE " & cadSel
        totRegPE = TotalRegistros(SQL)
        
        If Not ((totRegPE) > 0) And totRegPA = 0 Then
            MsgBox "No hay precios especiales a actualizar para esa fecha.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    '--- ACTUALIZAR LOS PRECIOS
    '---------------------------------
    menErrProceso = ""
    
    '-- Bloquear para que nadie mas pueda actualizar precios
    DesBloqueoManual ("ACTPRE")
    If Not BloqueoManual("ACTPRE", "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'PRECIOS ACTUALES
    If Me.chkPreuAct.Value = 1 And totRegPA > 0 Then
        '-- bloquear los registros a actualizar
        If Not BloqueaRegistro("slista", " not isnull(fechanue) and " & cadSel) Then
            MsgBox "No se ha podido actualizar precios actuales.", vbExclamation
        Else
            '-- proceso actualizar precios actuales
            Screen.MousePointer = vbHourglass
            ProcesoActualizarPrecios_Actuales cadSel, totRegPA
            Screen.MousePointer = vbDefault
        End If
        TerminaBloquear
    End If
    
    
    'PRECIOS ESPECIALES
    If Me.chkPreuEsp.Value = 1 And totRegPE > 0 Then
        '-- bloquear los registros a actualizar
        If Not BloqueaRegistro("sprees", " not isnull(fechanue) and " & cadSel) Then
            MsgBox "No se ha podido actualizar precios especiales.", vbExclamation
        Else
            '-- proceso actualizar precios especiales
            Screen.MousePointer = vbHourglass
            ProcesoActualizarPrecios_Especiales cadSel, totRegPE
            Screen.MousePointer = vbDefault
        End If
        TerminaBloquear
    End If
    
    DesBloqueoManual ("ACTPRE")
    
    If menErrProceso <> "" Then MsgBox menErrProceso, vbInformation
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Me.ProgressBar1.visible = False
    Me.lblProgreso(0).visible = False
    Me.lblProgreso(1).visible = False
    Me.Height = 4100
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgFecha(0).Tag = Index
    Set frmF = New frmCal
    frmF.Fecha = Now
    
    PonerFormatoFecha txtCodigo(Index)
    If txtCodigo(Index).Text <> "" Then frmF.Fecha = CDate(txtCodigo(Index).Text)
   
    Screen.MousePointer = vbDefault
    frmF.Show vbModal
    Set frmF = Nothing
    PonerFoco txtCodigo(Index)
End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Index = 0 Then 'fecha de cambio
        If txtCodigo(Index).Text <> "" Then
           PonerFormatoFecha txtCodigo(Index)
        End If
    End If
End Sub




Private Sub ProcesoActualizarPrecios_Actuales(cadWhere As String, totReg As Long)
'Actualizar los precios Actuales de las Tarifas
'(IN) cadWHERE: cadena seleccion de tarifas a actualizar
'Para cada tarifa a actualizar:
'   - insertar. en historico (slist1) linea con slista.fechanue y con el slista.precioac
'   - actualizar slista con slista.precioac=slista.precionu
'   - si slista.codlista es la tarifa de los parametros de la aplicacion: actualizar PVP del articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim hayErr As Boolean

    On Error GoTo ErrActPreu
   
    '-- iniciar la barra de progreso
    Me.Height = 4600
    Me.lblProgreso(0).Caption = "Actualizando precios actuales."
    Me.lblProgreso(0).visible = True
    Me.lblProgreso(1).visible = True
    CargarProgresNew Me.ProgressBar1, 100
    i = 0
    Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
    Me.ProgressBar1.visible = True
    
    
    '-- seleccionar todos los registros actuales a procesar
    SQL = "SELECT * FROM slista WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa a cambiar
    hayErr = False
    While Not RS.EOF
        '-- actualizar tarifas precios y PVP si corresponde
        If Not ActualizarTarifa(DBLet(RS!codArtic, "T"), DBLet(RS!codlista, "N")) Then
            hayErr = True
        End If
        
        '-- actualizar la progress bar
'        IncrementarProgresNew Me.ProgressBar1, 1
        i = i + 1
        Me.ProgressBar1.Value = CInt((i * 100) / totReg)
        Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
        Me.lblProgreso(0).Caption = "Actualizando precios actuales.     (" & i & " de " & totReg & ")"
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    
    Screen.MousePointer = vbDefault
    If Not hayErr Then
        Me.lblProgreso(0).Caption = "Proceso finalizado correctamente.     (" & i & " de " & totReg & ")"
'        MsgBox "Proceso actualización precios actuales finalizado correctamente.", vbInformation
        menErrProceso = "Proceso actualización precios actuales finalizado correctamente." & vbCrLf
    Else
        Me.lblProgreso(0).Caption = "Proceso finalizado con errores.     (" & i & " de " & totReg & ")"
        'MsgBox "Algunos precios actuales no se actualizaron correctamente.", vbExclamation
        menErrProceso = "Algunos precios actuales no se actualizaron correctamente." & vbCrLf
    End If
    Espera 0.2
    
    Exit Sub
    
ErrActPreu:
    MuestraError Err.Number, "Actualizar precios actuales.", Err.Description
End Sub




Private Function ActualizarTarifa(codArt As String, codLis As Integer) As Boolean
Dim cadErr As String
Dim cTar As CTarifaArt
Dim b As Boolean
Dim margen As Currency
Dim newPrecio As Currency

    On Error GoTo ErrAct
    Conn.BeginTrans
    
    
    Set cTar = New CTarifaArt
    b = cTar.LeerDatos(codArt, codLis)
    
    If b Then
        'actualizar la tarifa precios
        b = cTar.ActualizarPrecios(cTar.FechaCambio, cTar.PrecioNuevo, cTar.PrecioCajaNuevo, cadErr, True)
        
        'si tarifa es la de parametros actualizar PVP del articulo
        If b And codLis = vParamAplic.CodTarifa Then
            b = BloquearArticulo(codArt)
            If b Then
                margen = Round2(cTar.MargenComercial / 100, 4)
                newPrecio = Round2((cTar.PrecioNuevo / (margen + 1)), 4)
                b = ActualizarPVPArticulo(codArt, newPrecio)
            End If
        End If
    End If
    Set cTar = Nothing
    
    
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    
    ActualizarTarifa = b
    If Not b And cadErr <> "" Then MsgBox cadErr, vbExclamation
    Exit Function
    
ErrAct:
    Conn.RollbackTrans
    MuestraError Err.Number, "Actualizar precio actual tarifa.", Err.Description
End Function



Private Function BloquearArticulo(codArt As String) As Boolean
Dim cadWhere As String

    cadWhere = "codartic=" & DBSet(codArt, "T")
    BloquearArticulo = BloqueaRegistro("sartic", cadWhere)
End Function




Private Function ActualizarPVPArticulo(codArt As String, newPreu As Currency) As Boolean
Dim SQL As String
    
    On Error GoTo ErrActPVP
    ActualizarPVPArticulo = False
    
    SQL = "UPDATE sartic SET preciove=" & DBSet(newPreu, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(codArt, "T")
    Conn.Execute SQL
    
    ActualizarPVPArticulo = True
    Exit Function
    
ErrActPVP:
    ActualizarPVPArticulo = False
    MuestraError Err.Number, "Actualizar precio PVP del articulo.", Err.Description
End Function







Private Sub ProcesoActualizarPrecios_Especiales(cadWhere As String, totReg As Long)
'Actualizar los precios especiales de las Tarifas
'(IN) cadWHERE: cadena seleccion de precios a actualizar
'Para cada precio especial a actualizar:
'   - insertar. en historico (spree1) linea con sprees.fechanue y con el sprees.precioac
'   - actualizar sprees con sprees.precioac=sprees.precionu
'   - poner a nulos los valores nuevos
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim hayErr As Boolean

    On Error GoTo ErrActPreu
    
    '-- iniciar la barra de progreso
    Me.Height = 4600
    Me.lblProgreso(0).Caption = "Actualizando precios especiales."
    Me.lblProgreso(0).visible = True
    Me.lblProgreso(1).visible = True
    CargarProgresNew Me.ProgressBar1, 100 'CInt(totReg)
    i = 0
    Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
    Me.ProgressBar1.visible = True
    
    
    '-- seleccionar todos los registros actuales a procesar
    SQL = "SELECT * FROM sprees WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada precio especial a cambiar
    hayErr = False
    While Not RS.EOF
        '-- actualizar precios especiales
        If Not ActualizarPrecioEspec(RS!CodClien, RS!codArtic) Then
            'procesar errores!!!!!!!!!!!
            hayErr = True
        End If
        
        '-- actualizar la progress bar
'        IncrementarProgresNew Me.ProgressBar1, 1
        i = i + 1
        Me.ProgressBar1.Value = CInt((i * 100) / totReg)
        Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
        Me.lblProgreso(0).Caption = "Actualizando precios especiales.     (" & i & " de " & totReg & ")"
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    Screen.MousePointer = vbDefault
    If Not hayErr Then
        Me.lblProgreso(0).Caption = "Proceso finalizado correctamente.     (" & i & " de " & totReg & ")"
        'MsgBox "Proceso actualización precios especiales finalizado correctamente.", vbInformation
        menErrProceso = menErrProceso & "Proceso actualización precios especiales finalizado correctamente."
    Else
        Me.lblProgreso(0).Caption = "Proceso finalizado con errores.     (" & i & " de " & totReg & ")"
        'MsgBox "Algunos precios especiales no se actualizaron correctamente.", vbExclamation
         menErrProceso = menErrProceso & "Algunos precios especiales no se actualizaron correctamente."
    End If
    
    Exit Sub
    
ErrActPreu:
    MuestraError Err.Number, "Actualizar precios especiales.", Err.Description
End Sub




Private Function ActualizarPrecioEspec(codCli As Long, codArt As String) As Boolean
'actualizar precio especial
Dim SQL As String
Dim RS As ADODB.Recordset
Dim numF As String

    On Error GoTo ErrAct
    
    Conn.BeginTrans
    ActualizarPrecioEspec = False
    
    SQL = "SELECT * FROM sprees WHERE codclien=" & codCli & " AND codartic=" & DBSet(codArt, "T")
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        '-- Insertar en el historico spree1
        'numero de linea
        numF = SugerirCodigoSiguienteStr("spree1", "numlinea", "codartic=" & DBSet(codArt, "T") & " AND codclien=" & codCli)
    
        SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec) "
        SQL = SQL & " VALUES (" & codCli & "," & DBSet(codArt, "T") & "," & numF & ","
        SQL = SQL & DBSet(RS!fechanue, "F") & "," & DBSet(RS!precioac, "N") & "," & DBSet(DBLet(RS!precioa1, "N"), "N") & "," & DBSet(RS!dtoespec, "N") & ")"
        Conn.Execute SQL
        
        
        '-- Actualizar precios actuales con nuevo y resetear valores nuevos
        SQL = "UPDATE sprees SET precioac=" & DBSet(RS!precionu, "N")
        '    SQL = SQL & "," & " precioa1=" & DBSet(newPrecioA1, "N")
        SQL = SQL & ", dtoespec=" & DBSet(RS!dtoespe1, "N")
        SQL = SQL & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo & ", precion1=" & ValorNulo
        SQL = SQL & ", dtoespe1=" & ValorNulo
        SQL = SQL & " WHERE codclien=" & codCli & " and codartic=" & DBSet(codArt, "T")
        Conn.Execute SQL
    End If
    RS.Close
    Set RS = Nothing


    Conn.CommitTrans
    ActualizarPrecioEspec = True
    Exit Function
    
ErrAct:
    ActualizarPrecioEspec = False
    Conn.RollbackTrans
    MuestraError Err.Number, "Actualizar precio especial.", Err.Description
End Function

