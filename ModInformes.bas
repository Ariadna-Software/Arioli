Attribute VB_Name = "ModInformes"
Option Explicit


'============================================================
'====== FUNCIONES GENERALES  ================================


Public Sub AbrirListado(numero As Integer)
    Screen.MousePointer = vbHourglass
    frmListado.OpcionListado = numero
    frmListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function

Public Function AnyadirAFormulaOr(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
' La modificación es que la concatenación de criterios se hace con OR si se utiliza esta
' función [SERVICIOS]
    If arg = "Error" Then
        AnyadirAFormulaOr = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " OR (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormulaOr = True
End Function

Public Function NumRegistros(vSQL As String, Optional vBD As Byte) As Integer
'Devuelve si hay registros con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    If vBD = conConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    NumRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then
            NumRegistros = RS.Fields(0).Value
'            If RS.Fields(0).Value = 1 Then
'                RegistrosAListar = 1  'Solo es para saber que hay registros que mostrar
'            Else
'                RegistrosAListar = 2  'Solo es para saber que hay registros que mostrar
'            End If
        End If
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        NumRegistros = 0
        Err.Clear
    End If
End Function



Public Function RegistrosAListar(vSQL As String, Optional vBD As Byte) As Byte
'Devuelve si hay algun registro para mostrar en el Informe con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0 y no abrirá el informe
Dim RS As ADODB.Recordset

    On Error GoTo ErrReg

    Set RS = New ADODB.Recordset
    If vBD = conConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    RegistrosAListar = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then
            If RS.Fields(0).Value = 1 Then
                RegistrosAListar = 1  'Solo es para saber que hay registros que mostrar
            Else
                RegistrosAListar = 2  'Solo es para saber que hay registros que mostrar
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing

    Exit Function
    
ErrReg:
    RegistrosAListar = 0
    MuestraError Err.Number, "Comprobar si hay registros seleccionados", Err.Description
End Function



'Para que no muestre el mensaje de NO hay datos
'   optional: por defecto FALSE
Public Function HayRegParaInforme(cTabla As String, cWhere As String, Optional OcultarMSG As Boolean) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        SQL = SQL & " WHERE " & cWhere
    End If
    If RegistrosAListar(SQL) = 0 Then
        'Por defecto SI que lo muestra
        If Not OcultarMSG Then MsgBox "No hay datos para mostrar.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function


Public Function CadenaDesdeHasta(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String, Optional nomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim cadAux As String
On Error GoTo ErrDH

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                Case "FH"
                    cadAux = campo & " >= DateTime(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & "," & Hour(cadDesde) & "," & Minute(cadDesde) & "," & Second(cadDesde) & ")"
                    
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                        
                    Case "FH"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                                   
                            cadAux = cadAux & " AND " & campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                cadAux = cadAux & "23,59,59"
                            Else
                                cadAux = cadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            cadAux = cadAux & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                    Case "FH"
                            cadAux = campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                cadAux = cadAux & "23,59,59"
                            Else
                                cadAux = cadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            cadAux = cadAux & ")"
                        
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHasta = cadAux
ErrDH:
    If Err.Number <> 0 Then CadenaDesdeHasta = "Error"
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
                Case "FH"
                    If Len(cadDesde) = 10 Then cadDesde = cadDesde & " 00:00:00"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFechaHora) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and (" & campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                    Case "FH"
                        If Len(cadHasta) = 10 Then cadHasta = cadHasta & " 23:59:59"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = "(" & campo & " <= '" & Format(cadHasta, FormatoFechaHora) & "')"
                        End If

                    

                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHastaBD = cadAux
End Function



Public Function QuitarCaracterACadena(cadForm As String, Caracter As String) As String
'IN: [cadForm] es la cadena en la que se eliminara todos los caractes iguales a la vble [Caracter]
'OUT: cadena sin los caracteres
Dim i As Long
Dim J As Long
Dim Aux As String

    Aux = cadForm
    i = InStr(1, Aux, Caracter, vbTextCompare)
    While i > 0
        i = InStr(1, Aux, Caracter, vbTextCompare)
        If i > 0 Then
            J = Len(Caracter)
            Aux = Mid(Aux, 1, i - 1) & Mid(Aux, i + J, Len(Aux) - 1)
        End If
    Wend
    QuitarCaracterACadena = Aux
End Function


Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles

    vFrame.visible = visible
    If visible = True Then
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        vFrame.Top = -90
        vFrame.Left = 0
        vFrame.Width = W
        vFrame.Height = H
    End If
End Sub



'Public Function SustituirCadenas(CADENA As String, cad1 As String, cad2 As String) As String
''IN: Cadena es la cadena de texto en la que se va a sustituir la cad1 por la cad2
''OUT: cadena con la sustitucion
'
''EJEMPLO: cadena = "scaalb.codtipom='ALV' AND scaalb.numalbar=1"
''         cad1 = "scaalb"
''         cad2 = "slialb"
'
''         Resultado = "slialb.codtipom='ALV' AND slialb.numalbar=1"
'
'Dim i As Integer
'Dim J As Integer
'Dim Aux As String
'
'    Aux = CADENA
'    Do
'        i = InStr(1, Aux, cad1, vbTextCompare)
'        If i > 0 Then
'            J = Len(cad1)
'            Aux = Mid(Aux, 1, i - 1) & cad2 & Mid(Aux, i + J, Len(Aux) - 1)
'        End If
'    Loop Until i <= 0
'    SustituirCadenas = Aux
'End Function



'============================================================
'====== FUNCIONES PARA ARIGES  ==============================

Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.OpcionListado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Sub AbrirListadoPed(numero As Integer)
'Abre el Form con los listados de Pedidos
    Screen.MousePointer = vbHourglass
    frmListadoPed.OpcionListado = numero
    frmListadoPed.Show vbModal
    Screen.MousePointer = vbDefault
End Sub



Public Function PonerParamEmpresa(cadParam As String, numParam As Byte) As Boolean
Dim DomiEmp As String
Dim WebEmp As String
Dim Cad As String

        DomiEmp = vParam.DomicilioEmpresa & " - " & vParam.CPostal & " " & vParam.Poblacion
        If vParam.Provincia <> vParam.Poblacion Then DomiEmp = DomiEmp & " " & vParam.Provincia
        DomiEmp = DomiEmp & " - Telf. " & vParam.Telefono & " - Fax. " & vParam.Fax
        WebEmp = "Internet: " & vParam.WebEmpresa & " - E-mail: " & vParam.MailEmpresa
        'Resto parametros
        Cad = ""
        Cad = Cad & "pNomEmpre=""" & vParam.NombreEmpresa & """|"
        Cad = Cad & "pDomEmpre=""" & DomiEmp & """|"
        Cad = Cad & "pWebEmpre=""" & WebEmp & """|"
        
        numParam = numParam + 3
        cadParam = cadParam & Cad
        PonerParamEmpresa = True
End Function


Public Function PonerParamRPT(Indice As Byte, cadParam As String, numParam As Byte, nomDocu As String, Optional ImpresionDirecta As Boolean) As Boolean
Dim vParamRpt As CParamRpt 'Tipos de Documentos
Dim Cad As String

    Set vParamRpt = New CParamRpt
    
    If vParamRpt.Leer(Indice) = 1 Then
        Cad = "No se han podido cargar los Parámetros de Tipos de Documentos." & vbCrLf
        MsgBox Cad & "Debe configurar la aplicación.", vbExclamation
        Set vParamRpt = Nothing
        PonerParamRPT = False
        Exit Function
    Else
        If cadParam = "" Then
            Cad = "|"
        Else
            Cad = ""
        End If
        Cad = Cad & "pCodigoISO=""" & vParamRpt.CodigoISO & """|"
        If vParamRpt.CodigoRevision = -1 Then
            Cad = Cad & "pCodigoRev=""" & "" & """|"
        Else
            Cad = Cad & "pCodigoRev=""" & Format(vParamRpt.CodigoRevision, "00") & """|"
        End If
        numParam = numParam + 2
        If vParamRpt.LineaPie1 <> "" Then
            Cad = Cad & "pLinea1=""" & vParamRpt.LineaPie1 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie2 <> "" Then
            Cad = Cad & "pLinea2=""" & vParamRpt.LineaPie2 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie3 <> "" Then
            Cad = Cad & "pLinea3=""" & vParamRpt.LineaPie3 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie4 <> "" Then
            Cad = Cad & "pLinea4=""" & vParamRpt.LineaPie4 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie5 <> "" Then
            Cad = Cad & "pLinea5=""" & vParamRpt.LineaPie5 & """|"
            numParam = numParam + 1
        End If
        cadParam = cadParam & Cad
        nomDocu = vParamRpt.Documento
        ImpresionDirecta = vParamRpt.ImprimeDirecto
        PonerParamRPT = True
        Set vParamRpt = Nothing
    End If
End Function


Public Sub PonerParamCadOferta(cadParam As String, numParam As Byte, cadSelect As String)
'Concatena los Nº de Ofertas que se van a imprimir, y lo añade como parametro
' a los parametros que se pasaran al Report.
'Añade el parametro: pCadOfertas= 0000001, 0000002, ...
'RPT que lo utiliza: AriOfertas.rpt
Dim cadOfertas As String
Dim SQL As String
Dim i As Byte
Dim RS As ADODB.Recordset

    On Error GoTo EPonParam
    
    cadOfertas = ""
    SQL = "scapre"

    i = InStr(1, cadSelect, "scapre")
    If Not (i > 0) Then SQL = "schpre"

    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")

    SQL = "SELECT distinct numofert from  " & SQL
    SQL = SQL & " WHERE " & cadSelect
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not RS.EOF
        If Len(cadOfertas) > 75 Then
            If InStr(cadOfertas, "...") > 0 Then
                RS.MoveNext
            Else
                cadOfertas = cadOfertas & ", ..."
            End If
            
        Else
            If cadOfertas = "" Then
                cadOfertas = Format(RS.Fields(0).Value, "0000000")
            Else
                cadOfertas = cadOfertas & ", " & Format(RS.Fields(0).Value, "0000000")
            End If
            RS.MoveNext
        End If
    Wend
    RS.Close
    Set RS = Nothing

    cadParam = cadParam & "pCadOfertas=""" & cadOfertas & """|"
    numParam = numParam + 1
    
EPonParam:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo paramétros del informe.", Err.Description
End Sub





Public Function PonerNombreImpresora() As String
On Error Resume Next
    PonerNombreImpresora = Printer.DeviceName
    If Err.Number <> 0 Then
        PonerNombreImpresora = "No hay impresora instalada"
        Err.Clear
    End If
End Function


Public Sub EstablecerImpresora(Nombre As String)
Dim X As Printer
    For Each X In Printers
       If X.DeviceName = Nombre Then
          ' La define como predeterminada del sistema.
          Set Printer = X
          ' Sale del bucle.
          Exit For
       End If
    Next

End Sub
  


Public Function NombreImpresoraTicket(nTermi As Integer) As String
    On Error GoTo ErrNomImp
    
    If vParamTPV Is Nothing Then
    
        'Establecemos la impresora de ticket
        Set vParamTPV = New CParamTPV
        If vParamTPV.Leer2(CStr(nTermi)) = 0 Then
             NombreImpresoraTicket = vParamTPV.NomImpresora
    '        If vParamTPV.NomImpresora <> "" Then
    
    '            If Printer.DeviceName <> vParamTPV.NomImpresora Then
    '                NomImpre = Printer.DeviceName
    '                EstablecerImpresora vParamTPV.NomImpresora
    '            End If
    '        End If
        End If
        Set vParamTPV = Nothing
    Else
        NombreImpresoraTicket = vParamTPV.NomImpresora
    End If
    
    Exit Function
ErrNomImp:
    MuestraError Err.Number, "Obtener nombre impresora de Ticket", Err.Description
End Function



Public Function ObtenerTerminal() As Integer
Dim SQL As String
    
    On Error GoTo ErrTermi

    'Obtener que terminal es
    'Terminal con el que trabajaremos, leemos el nombre del ordenador
    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
    If Not IsNumeric(SQL) Then
        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
'    Else
'        bImpre = True
    End If
    ObtenerTerminal = CInt(SQL)
    Exit Function
    
ErrTermi:
    MuestraError Err.Number, "Obtener terminal.", Err.Description
    ObtenerTerminal = 0
End Function



Public Function SaltosDeLinea(ByVal Cadena As String) As String
    Dim Devu As String
    Dim i As Integer
    
    Devu = ""
    Do
        i = InStr(1, Cadena, vbCrLf)
        If i > 0 Then
            If Devu <> "" Then Devu = Devu & """ + chr(13) + """
            Devu = Devu & Mid(Cadena, 1, i - 1)
            Cadena = Mid(Cadena, i + 2)
            
       Else
            Devu = Devu & Cadena
       End If
    Loop While i > 0
    SaltosDeLinea = Devu
End Function





'MAYO 2010
Public Sub LlamaImprimirGral(cadFormula As String, cadParam As String, numParam As Integer, Nomrpt As String, Titulo As String)
            With frmImprimir
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .Titulo = Titulo
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 2009
                .NombreRPT = Nomrpt
                .ConSubInforme = True
                .Show vbModal
            End With

End Sub



''DesdeDonde
''   0.-Albaran proveedor
''   1.-Desde partidas
''   2.-Movimientos partidas
''   3.- DESDE FACTURA DE PROVEEDOR
Public Sub ImpirmirEtiquetas(WhereAlb As String, Proveedor As String, MostrarMsgNohay As Boolean, DesdeDonde As Byte)
Dim C As String


    C = "DELETE FROM tmpnlotes where codusu = " & vUsu.Codigo
    Conn.Execute C


    'En tmpnlotes.nomartic, metere la cantidad
    'cuando haga el select lo cruzare
    C = "insert into tmpnlotes(codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic,cantidad, numlotes) "

    C = C & "select " & vUsu.Codigo & ",numalbar,fechaalb,slialp.codprove,numlinea,slialp.codartic,codalmac"
    C = C & ",format(cantidad,2),etiquetas "
    C = C & ",numlotes from slialp,sartic where slialp.codartic=sartic.codartic and  trazabilidad =1 and factorconversion=1"
    C = C & " AND " & WhereAlb

    If EjecutaSQL(conAri, C, True) Then
        Espera 0.3
        C = DevuelveDesdeBD(conAri, "count(*)", "tmpnlotes", "codusu", vUsu.Codigo)
        If C = "" Then C = "0"
        If Val(C) = 0 Then
            If MostrarMsgNohay Then MsgBox "No hay artículos de trazabilidad de materia auxiliar", vbExclamation
        Else
            '-----------------------------------------
            frmComImprimirEtiquetas2.GuardarEImprimir = DesdeDonde = 1
            frmComImprimirEtiquetas2.vTexto = Proveedor
            frmComImprimirEtiquetas2.Show vbModal


        End If

    End If
End Sub

Public Sub ImpirmirEtiquetas2(ByRef ColP As Collection, Proveedor As String, MostrarMsgNohay As Boolean, DesdeDonde As Byte)
Dim C As String
Dim i As Integer
Dim Cp As cPartidas
Dim J As Integer

    C = "DELETE FROM tmpnlotes where codusu = " & vUsu.Codigo
    Conn.Execute C
    
    Set Cp = New cPartidas
    C = ""
    For i = 1 To ColP.Count
        If Cp.Leer(Val(ColP.Item(i))) Then
            J = Cp.CuantasEtiquetas                                                         'codprove=idPARTIDA
            'C = "insert into tmpnlotes(codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic,cantidad, numlotes) "
            C = C & ", (" & vUsu.Codigo & "," & DBSet(Cp.NumAlbar, "T") & "," & DBSet(Cp.Fecha, "F") & "," & Cp.IdPartida & "," & i & ","
            '                                                           en nomartic va la cant linea
            C = C & DBSet(Cp.codArtic, "T") & "," & Cp.codalmac & ",'" & Format(Cp.Cantidad, FormatoCantidad)
            C = C & "'," & J & "," & DBSet(Cp.Numlote, "T") & ")"
        End If
    Next i
    
    If C = "" Then
        
        Exit Sub
    End If
    C = Mid(C, 2) 'la 1ªcoma
    C = "insert into tmpnlotes(codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic,cantidad, numlotes) VALUES " & C

    
    If EjecutaSQL(conAri, C, True) Then
         Espera 0.3
        C = DevuelveDesdeBD(conAri, "count(*)", "tmpnlotes", "codusu", vUsu.Codigo)
        If C = "" Then C = "0"
        If Val(C) = 0 Then
            If MostrarMsgNohay Then MsgBox "No hay artículos de trazabilidad de materia auxiliar", vbExclamation
        Else
            '-----------------------------------------
            'ANTES
            'frmComImprimirEtiquetas2.GuardarEImprimir = DesdeDonde = 1
            'AHORA
            frmComImprimirEtiquetas2.GuardarEImprimir = False
            If DesdeDonde = 1 Then
                If vUsu.nivel < 1 Then
                    frmComImprimirEtiquetas2.GuardarEImprimir = True
                Else
                    If Not Cp.TieneEtiquetasYAenProduccion Then frmComImprimirEtiquetas2.GuardarEImprimir = True
                End If
            End If
            
            frmComImprimirEtiquetas2.vTexto = Proveedor
            frmComImprimirEtiquetas2.Show vbModal
            
            
        End If
        
    End If
End Sub


'Devolvera el nombre del report, puro y duro. Es un devuelve desde BD
Public Function DevuelveNombreReport(Indice As Integer) As String
    DevuelveNombreReport = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", CStr(Indice), "T")
End Function
