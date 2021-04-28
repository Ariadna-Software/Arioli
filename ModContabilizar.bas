Attribute VB_Name = "ModContabilizar"
Option Explicit


'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private TotalFac As Currency
Private CCoste2 As String

Private vCCos As Byte
    'Para cuando pasamos en la contabilizacion de las facturas
    'Sera 2:    tiene mas de un centro de coste. Habra que agrupar por CC
    '     1:  o solo es un trabajador o tienen el mismo CC, con lo cual no hace falta agrupar por CC
    '     0:  no habra CC.  Si vpara.. tieneanalitica = false

Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

'Para pasar a contabilidad facturas de proveedor
Private AnyoFacPr As Integer 'año factura proveedor, es el ano de fecha_recepcion

'Modificacion Centro de coste.
'La factura cogera el Centro de coste del trabajador del albaran




'llevara: codmacta_proveedor | impo_retencion |
Private DatosRetencion As String
Private DatosAportacion As String


' Nueva contabilizacion
' Los IVAS de la cabceera se los paso a las lineas
' ya que agrupara por cuenta, y codigiva
Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency





Public Function CrearTMPFacturas(cadTabla As String, cadWhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    SQL = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTabla = "scafac" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10)  NOT NULL ,"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute SQL
     
     
    If cadTabla = "scafac" Then
        SQL = "SELECT codtipom, numfactu, fecfactu"
    Else
        SQL = "SELECT codprove, numfactu, fecfactu"
    End If
    SQL = SQL & " FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere
    
    'DAVID###
    'Si son de proveedores el orden es MUY importante para
    'que vayan ordenaditas por fecha recepcion
    'ademas, por si tiene mas de una por prove añado los dos campos
    If cadTabla <> "scafac" Then SQL = SQL & " ORDER BY fecrecep,codprove,numfactu"

    
    SQL = " INSERT INTO tmpFactu " & SQL
    conn.Execute SQL

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        MuestraError Err.Number, "", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadWhere As String)
Dim SQL As String

    On Error Resume Next
    SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    SQL = SQL & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
    conn.Execute SQL
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    SQL = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTabla = "scafac" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10) NOT NULL ,"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "error varchar(200) NULL )"
    conn.Execute SQL
     
     CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrFac;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTabla As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTabla = "scafac" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        SQL = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open SQL, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        SQL = "select distinct scafac.codtipom from " & cadTabla
        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'        SQL = SQL & cadWHERE
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        b = True
        While Not RS.EOF And b
            'comprobar que todas las letras serie existen en Ariges
            SQL = "letraser"
            devuelve = DevuelveDesdeBDNew(conAri, "stipom", "codtipom", "codtipom", RS!Codtipom, "T", SQL)
            If devuelve = "" Then
                b = False
                cad = RS!Codtipom & " en BD de Gestión."
            ElseIf SQL <> "" Then
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(SQL, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    cad = SQL & " en BD de Contabilidad."
                End If
            End If
            If b Then cad = cad & DBSet(RS!Codtipom, "T") & ","
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not b Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            SQL = "select count(*) from stipom "
            SQL = SQL & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(SQL) > 0 Then
                SQL = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                SQL = SQL & "Comprobar en la tabla de tipos de movimiento: " & cad
                MsgBox SQL, vbExclamation
                Exit Function
            End If
        End If
        ComprobarLetraSerie = True
    Else
        ComprobarLetraSerie = True
    End If

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function


Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim SQLconta As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
    SQLconta = "SELECT count(*) FROM cabfact WHERE "
        
        
        

        If vParamAplic.ContabilidadNueva Then
            SQLconta = "SELECT count(*) FROM factcli WHERE "
        Else
            SQLconta = "SELECT count(*) FROM cabfact WHERE "
        End If
        
        
        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "

        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not RS.EOF And b
            If vParamAplic.ContabilidadNueva Then
                SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND numfactu=" & DBSet(RS!NumFactu, "N") & " AND anofactu=" & Year(RS!Fecfactu) & ")"
            Else
                SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!Fecfactu) & ")"
            End If
            SQL = SQLconta & SQL
            If RegistrosAListar(SQL, conConta) Then
                b = False
                SQL = "          Letra Serie: " & DBSet(RS!LetraSer, "T") & vbCrLf
                SQL = SQL & "          Nº Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & RS!Fecfactu
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
'    Else
'        ComprobarNumFacturas_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function




'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte) As Boolean
''Comprobar que todas las ctas contables de los distintos clientes de las facturas
''que vamos a contabilizar existan en la contabilidad
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'Dim cadG As String
'
'    On Error GoTo ECompCta
'
'    ComprobarCtaContable = False
'
'    If Opcion = 3 Then 'si hay analitica comprobar que todas las cuentas
'                        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
'        cadG = "grupovta"
'        SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
'        If SQL <> "" And cadG <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
'        ElseIf SQL <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%')"
'        ElseIf cadG <> "" Then
'            SQL = " AND (codmacta like '" & cadG & "%')"
'        End If
'        cadG = SQL
'    End If
'
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If
'
'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
'
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
'
'            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            If Opcion <> 3 Then
'                SQL = "No existe la cta contable " & SQL
'            Else
'                SQL = "La cuenta " & SQL & " no es del nivel correcto."
'            End If
'            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarCtaContable = False
'        Else
'            ComprobarCtaContable = True
'        End If
'    Else
'        ComprobarCtaContable = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompCta:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
'    End If
'End Function






Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable_new = False
    
    cadG = ""
    
    
    If Opcion = 3 Then
            'si hay analitica comprobar que todas las cuentas
            'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
            cadG = "grupovta"
            SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
            If SQL <> "" And cadG <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
            ElseIf SQL <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%')"
            ElseIf cadG <> "" Then
                SQL = " AND (codmacta like '" & cadG & "%')"
            End If
            cadG = SQL
    End If
    
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTabla = "scafac" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            
            SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
            SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
            SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
        Else
            'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
            SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
            SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
            SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
        End If
    
    ElseIf Opcion = 2 Or Opcion = 3 Then
        SQL = "SELECT distinct "
        If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
        If cadTabla = "scafac" Then
            SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
            SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
        Else
            SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
            SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
        End If
        SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
        
    ElseIf Opcion = 4 Then
        'opcion para la contabilizacion de tickets AGRUPADA  FTG
        
        
        Set RS = New ADODB.Recordset
      
        
        cadG = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND  codtipom='FTG' "
        RS.Open cadG, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cadG = ""
        Do
            cadG = cadG & "," & RS!NumFactu
            RS.MoveNext
        Loop Until RS.EOF
        RS.Close
        Set RS = Nothing
        cadG = Mid(cadG, 2)
        'Monto el SELECT , igual que el de arriba, pero partiendo de los FTIs
         SQL = "SELECT distinct  sartic.codfamia, sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1"
         SQL = SQL & " from (slifac   INNER JOIN sartic ON slifac.codartic=sartic.codartic)  LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia"
         SQL = SQL & " WHERE  codtipom='FTI' and numfactu IN (" & cadG & ")"
         cadG = ""
         'Fuerzo para que haga las mismas comprobaciones que si fuera la opcion 2
         Opcion = 2
         
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    b = True

    While Not RS.EOF And b
        SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Codmacta, "T")
        
        If Not (RegistrosAListar(SQL, conConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "scafac" Then
                    SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
                Else
                    SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
                End If
            ElseIf Opcion = 2 Then
                SQL = RS!Codmacta & " de la familia " & Format(RS!Codfamia, "0000")
            ElseIf Opcion = 3 Then
                SQL = RS!Codmacta
            End If
        End If
        
        
        If Opcion = 2 Or Opcion = 3 Then
            'Comprobar que ademas de existir la cuenta de ventas exista tambien
            'la cuenta ABONO ventas (sfamia.aboventa)
            '---------------------------------------------
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctaabono, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
            If Not (RegistrosAListar(SQL, conConta) > 0) Then
                b = False 'no encontrado
                If Opcion = 2 Then
                    SQL = RS!ctaabono & " de la familia " & Format(RS!Codfamia, "0000")
                ElseIf Opcion = 3 Then
                    SQL = RS!ctaabono
                End If
            End If
            
            
            'comprobar cuentas alternativas solo para facturacion a CLIENTES
            '----------------------------------------------------------------
            If cadTabla = "scafac" Then
                ' Comprobar cuenta VENTA alternativa
                If DBLet(RS!ctavent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctavent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = RS!ctavent1 & " de la familia " & Format(RS!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = RS!ctavent1
                        End If
                    End If
                Else
                    b = False
                    SQL = " o la familia no tiene asignada cuenta venta alternativa."
                End If
                
                ' Comprobar cuenta de ABONO alternativa
                If DBLet(RS!abovent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!abovent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = RS!abovent1 & " de la familia " & Format(RS!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = RS!abovent1
                        End If
                    End If
                Else
                    b = False
                    SQL = " o la familia no tiene asignada cuenta abono alternativa."
                End If
            End If
            
        End If
        
        RS.MoveNext
    Wend
    
    
        
        
        
        If Not b Then
            If Opcion <> 3 Then
                SQL = "No existe la cta contable " & SQL
            Else
                SQL = "La cuenta " & SQL & " no es del nivel correcto. (Familias de artículos)."
            End If
            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarCtaContable_new = False
        Else
            ComprobarCtaContable_new = True
        End If
        
        
        
        
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function




Public Function ComprobarCtabasesMoixent(esCliente As Boolean, Opcion As Byte) As Boolean
'Las bases ahora NO van con las ctas de sfamia, van con las de sfamcontab
Dim SQL As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim SQLcuentas As String
Dim cadG As String
Dim CadenaErroresCta As String
Dim Nulos As Boolean

    On Error GoTo ECompCta

    ComprobarCtabasesMoixent = False
    
    cadG = ""
    
   
   
    
    'si hay analitica comprobar que todas las cuentas
    'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
    cadG = "grupovta"
    SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
    If SQL <> "" And cadG <> "" Then
        SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
    ElseIf SQL <> "" Then
        SQL = " AND (codmacta like '" & SQL & "%')"
    ElseIf cadG <> "" Then
        SQL = " AND (codmacta like '" & cadG & "%')"
    End If
    cadG = SQL

    
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    

    SQL = "SELECT distinct "
    If Opcion = 2 Then SQL = SQL & " sartic.codfamia,sartic.codmarca,sartic.codunida,sartic.codtipar,"
    If esCliente Then
        SQL = SQL & " sfamcontab.ctaventa as codmacta,sfamcontab.aboventa as ctaabono, sfamcontab.ctavent1,sfamcontab.abovent1 from ((slifac "
        SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
        SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
    Else
        SQL = SQL & " sfamcontab.ctacompr as codmacta,sfamcontab.abocompr as ctaabono from ((slifpc "
        SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
        SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
    End If
 
 
    SQL = SQL & " LEFT OUTER JOIN sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
    SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    b = True
    CadenaErroresCta = ""
    
    While Not RS.EOF And b
    
        Nulos = IsNull(RS!Codmacta) Or IsNull(RS!ctaabono)
        If esCliente Then Nulos = Nulos Or IsNull(RS!ctavent1) Or IsNull(RS!abovent1)
            
        
    
        If Nulos Then
                
                'Campo NULO para:
                SQL = ""
                If IsNull(RS!Codmacta) Then SQL = SQL & "/ cuenta venta"
                If IsNull(RS!ctaabono) Then SQL = SQL & "/ cuenta abono"
                If esCliente Then
                    If IsNull(RS!ctavent1) Then SQL = SQL & "/ alt. cuenta"
                    If IsNull(RS!abovent1) Then SQL = SQL & "/ alt. abono"
                End If
                SQL = Mid(SQL, 2)
                SQL = " NULO: " & RS!Codfamia & "(F)  " & RS!codmarca & "(M)    " & RS!CodUnida & "(U)  " & RS!codTipar & "(T) -> " & SQL
                
                CadenaErroresCta = CadenaErroresCta & vbCrLf & SQL
        Else
    
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Codmacta, "T")
            
            If Not (RegistrosAListar(SQL, conConta) > 0) Then
            'si no lo encuentra
                b = False 'no encontrado
                
                If Opcion = 2 Then
                    SQL = RS!Codmacta & " de la familia " & Format(RS!Codfamia, "0000")
                'ElseIf opcion = 3 Then
                Else
                    SQL = RS!Codmacta
                End If
            End If
            
            
            
                'Comprobar que ademas de existir la cuenta de ventas exista tambien
                'la cuenta ABONO ventas (sfamia.aboventa)
                '---------------------------------------------
                SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctaabono, "T")
    '            RSconta.MoveFirst
    '            RSconta.Find (SQL), , adSearchForward
    '            If RSconta.EOF Then
                If Not (RegistrosAListar(SQL, conConta) > 0) Then
                    b = False 'no encontrado
                    If Opcion = 2 Then
                        SQL = RS!ctaabono & " de la familia " & Format(RS!Codfamia, "0000")
                    'ElseIf opcion = 3 Then
                    Else
                        SQL = RS!ctaabono
                    End If
                End If
                
                
                'comprobar cuentas alternativas solo para facturacion a CLIENTES
                '----------------------------------------------------------------
                If esCliente Then
                     ' Comprobar cuenta VENTA alternativa
                     If DBLet(RS!ctavent1, "T") <> "" Then
                         SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctavent1, "T")
        
                         If Not (RegistrosAListar(SQL, conConta) > 0) Then
                             b = False 'no encontrado
                             If Opcion = 2 Then
                                 SQL = RS!ctavent1 & " de la familia " & Format(RS!Codfamia, "0000")
                             'ElseIf opcion = 3 Then
                             Else
                                 SQL = RS!ctavent1
                             End If
                         End If
                     Else
                         b = False
                         SQL = " o la familia no tiene asignada cuenta venta alternativa."
                     End If
                     
                     ' Comprobar cuenta de ABONO alternativa
                     If DBLet(RS!abovent1, "T") <> "" Then
                         SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!abovent1, "T")
         '                    RSconta.MoveFirst
         '                    RSconta.Find (SQL), , adSearchForward
         '                    If RSconta.EOF Then
                         If Not (RegistrosAListar(SQL, conConta) > 0) Then
                             b = False 'no encontrado
                             If Opcion = 2 Then
                                 SQL = RS!abovent1 & " de la familia " & Format(RS!Codfamia, "0000")
                             'ElseIf opcion = 3 Then
                             Else
                                 SQL = RS!abovent1
                             End If
                         End If
                     Else
                         b = False
                         SQL = " o la familia no tiene asignada cuenta abono alternativa."
                     End If
            End If ' de escliente
        End If 'de nulls

        
        RS.MoveNext
    Wend
    
    
    If CadenaErroresCta <> "" Then b = False
        
        
        If Not b Then
        
            
        
            If Opcion <> 3 Then
                SQL = "Error cuentas contabilizacion" & vbCrLf & CadenaErroresCta
            Else
                SQL = "La cuenta " & SQL & " no es del nivel correcto. (Familias de artículos)."
                SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL & CadenaErroresCta
                
            End If
            
            MsgBox SQL, vbExclamation
            
           
        Else
            ComprobarCtabasesMoixent = True
        End If
        
        
        
        
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function






Public Function ComprobarTiposIVA(cadTabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim I As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For I = 1 To 3
            If cadTabla = "scafac" Then
                SQL = "SELECT DISTINCT scafac.codigiv" & I
                SQL = SQL & " FROM scafac "
                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(codigiv" & I & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            Else
                SQL = "SELECT DISTINCT scafpc.tipoiva" & I
                SQL = SQL & " FROM " & cadTabla
                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(tipoiva" & I & ")"
'                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not RS.EOF And b
                SQL = "codigiva= " & DBSet(RS.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "Tipo de IVA: " & RS.Fields(0)
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
        
            If Not b Then
                SQL = "No existe el " & SQL
                SQL = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & SQL
            
                MsgBox SQL, vbExclamation
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next I
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function

'La comprobacion del centro de coste ha cambiado
'El centro de coste lo cojera de CADA factura donde tiene
'un trabajador asignado. Luego ya no necesito cadCC
'Comprobaremos:
'           que todas las facturas el trabajador asignado tiene CC
'           y que es distintos, puesto que si es el mismo CC no hare la fiesta
Public Function ComprobarCCoste2(cadSQL As String, Clientes As Boolean) As Byte
Dim SQL As String
Dim cadCC As String 'Para quitar


    On Error GoTo ECCoste

    ComprobarCCoste2 = 0
    Set miRsAux = New ADODB.Recordset
    
    If Clientes Then
        SQL = "select codccost from scafac , scafac1, straba "
        SQL = SQL & " WHERE scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and"
        SQL = SQL & " scafac.fecfactu=scafac1.fecfactu  and scafac1.codtraba=straba.codtraba"
        
    Else
        'PROVEEDORES
        SQL = "select codccost from scafpc ,scafpa, straba WHERE"
        SQL = SQL & " scafpc.codProve = scafpa.codProve And scafpc.NumFactu = scafpa.NumFactu And"
        SQL = SQL & " scafpc.FecFactu = scafpa.FecFactu AND codtrab2=straba.codtraba"
        
        
    End If
    
    If cadSQL <> "" Then SQL = SQL & " AND " & cadSQL
    'Grouo
    SQL = SQL & " GROUP BY codccost"
    
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            'MAL MAL. NO puede ser NULO
            MsgBox "Centro de coste de alguno de los trabajadores es NULO", vbExclamation
            SQL = ""
            miRsAux.MoveLast
            miRsAux.MoveNext
        Else
            SQL = SQL & "1"
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    'Ahora.
    'Si =""  Tiene un CC a NULL
    'If SQL = "" Then
        'NO hacemos nada, pq ya esta puesto el error
    If SQL <> "" Then
        If Len(SQL) = 1 Then
            'Todos los CEntros de coste son el mismo. Con lo cual NO hara falta agrupar por trabajador
            ComprobarCCoste2 = 1
        Else
            'Tiene CC distintos. SI agruparemos por Trabajador
            ComprobarCCoste2 = 2
        End If
    End If
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Cento de Coste", Err.Description
    End If
    Set miRsAux = Nothing
End Function


'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
'

Public Function PasarFactura(cadWhere As String, CodCCost As Byte, EsContabilizacionAgrupadaTickets As Boolean, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada

'EsContabilizacionAgrupadaTickets:  La diferencia es en las lineas de la factura.
'                                   Si false: procedimeineto normal
'                                       true: Las lineas hare los select de otra forma
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim TipoIvaFactura As Byte '0 Normal   1 R.E     2 Exento

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
     'MARZO 2011
    'lAs facturas internas NO van al registro
    'Solo meten el apunte
    If InStr(1, cadWhere, "'FAI'") > 0 Then
        'Estamos contabilizando una factura FAI, INTERNA
        'No entra en el registro de IVA de la contabilidad, solo el apunte
        vCCos = CodCCost
        
        b = ContabilizaFAI(cadMen, cadWhere, vContaFra)
        cadMen = "Insertando Cab. Factura interna: " & cadMen
    
    
    Else
        'Insertar en la conta Cabecera Factura
        b = InsertarCabFact(cadWhere, cadMen, vContaFra, TipoIvaFactura) '
        cadMen = "Insertando Cab. Factura: " & cadMen
        vCCos = CodCCost
        If b Then
     
           ' 'Insertar lineas de Factura en la Conta
           If EsContabilizacionAgrupadaTickets Then
                MsgBox "Funcion. No deberia haber entrado aqui", vbExclamation
           '     'Tickets agrupados
           '     b = InsertarLinFact_TicketsAgrupados("scafac", cadWhere, cadMen, False)
            Else
                'Normal. Esta es la forma NORMAL NORMAL de hacerlo
                b = InsertarLinFact("scafac", cadWhere, cadMen, False, 0, "", TipoIvaFactura)   'TipoIvaFactura
            End If
            cadMen = "Insertando Lin. Factura: " & cadMen
    
            If vContaFra.RealizarContabilizacion Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            End If

        End If
    End If

    If b Then
        'Poner intconta=1 en ariges.scafac
        b = ActualizarCabFact("scafac", cadWhere, cadMen)
        cadMen = "Actualizando Factura: " & cadMen
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    

    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFactura = False
        'Inserto en errores, DESPUES del rollback. Si no no lo refleja, y al hacer el rollback
        'tira atras la insercion
        SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "tmpFactu")
        conn.Execute SQL
        
    End If
        

        
    
End Function



' TipoIvaFactura As Byte '0 Normal   1 R.E     2 Exento
Private Function InsertarCabFact(cadWhere As String, cadErr As String, ByRef vCF As cContabilizarFacturas, ByRef QueTipoDeIVA As Byte) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
 Dim Aux As String
 
'Nueva contabilidad
Dim ImporAux As Currency
Dim TipoOpera As Byte   'Nueva contabilidad. Tipo operacion (usuarios.wtipopera)
Dim CadenaInsertFaclin2 As String
Dim SQL2 As String

    On Error GoTo EInsertar
    
    SQL = SQL & " SELECT stipom.letraser,numfactu,fecfactu, sclien.codmacta,sclien.cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,tipoiva,scafac.codtipom"
    If vParamAplic.ContabilidadNueva Then
        SQL = SQL & " ,scafac.codagent,scafac.nomclien,scafac.domclien,scafac.codpobla,scafac.pobclien,"
        SQL = SQL & " scafac.proclien,scafac.nifclien, codpais,scafac.coddirec,scafac.codforpa"
    End If
    
    
    
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    QueTipoDeIVA = 0
    If Not RS.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = RS!DtoPPago
        DtoGnral = RS!DtoGnral
        BaseImp = RS!baseimp1 + CCur(DBLet(RS!baseimp2, "N")) + CCur(DBLet(RS!baseimp3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = RS!TotalFac
        DatosAportacion = ""
        If RS!Aportacion > 0 Then
            'Deberia dar error si vparam.ctaaportacion=""
            DatosAportacion = RS!Codmacta & "|" & RS!Aportacion & "|"
        Else
            
        End If
        '----
        conCtaAlt = RS!cliAbono
        
        
        'Guardamos los valores de la factura que estoy integrando
        If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura RS!NumFactu, Year(RS!Fecfactu), RS!LetraSer
        
        SQL = ""
        SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & DBSet(RS!Codmacta, "T") & "," & Year(RS!Fecfactu) & ","
        Select Case vParamAplic.ObsFactura
        Case 0
            'Vacio
            SQL = SQL & ValorNulo
        Case 1
            'Nº Factura
            SQL = SQL & "'" & DevNombreSQL("N/Fra " & RS!NumFactu) & "'"
        Case 2
            'Fecha integracion
            SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
        End Select
        Nulo2 = "N"
        Nulo3 = "N"
        If DBLet(RS!codigiv2, "N") = 0 Then Nulo2 = "S"
        If DBLet(RS!codigiv3, "N") = 0 Then Nulo3 = "S"
    
    
    
    
    If vParamAplic.ContabilidadNueva Then
     
            'totbases ,totbasesret,
            ImporAux = RS!baseimp1 + DBLet(RS!baseimp2, "N") + DBLet(RS!baseimp3, "N")
            SQL = SQL & "," & DBSet(ImporAux, "N") & "," & ValorNulo & ","
            'totivas
            ImporAux = RS!imporiv1 + DBLet(RS!imporiv2, "N") + DBLet(RS!imporiv3, "N")
            SQL = SQL & DBSet(ImporAux, "N") & ","
            ',totrecargo,totfaccl,retfaccl,trefaccl,cuereten,tiporeten,
            ImporAux = DBLet(RS!imporiv1re, "N") + DBLet(RS!imporiv2re, "N") + DBLet(RS!imporiv3re, "N")
            If ImporAux <> 0 Then QueTipoDeIVA = 1
            SQL = SQL & DBSet(ImporAux, "N") & "," & DBSet(RS!TotalFac, "N")
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
            
            
                        
            'fecliqcl,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,dpto,codagente,codforpa,escorrecta,
            SQL = SQL & DBSet(RS!Fecfactu, "F") & "," & DBSet(RS!nomclien, "T") & "," & DBSet(RS!domclien, "T", "S") & ","
            SQL = SQL & DBSet(RS!codpobla, "T", "S") & "," & DBSet(RS!pobclien, "T", "S") & "," & DBSet(RS!proclien, "T", "S") & ","
            SQL = SQL & DBSet(RS!nifClien, "T", "S") & "," & DBSet(RS!codpais, "T", "S") & "," & DBSet(RS!CodDirec, "T", "S") & ","
            SQL = SQL & DBSet(RS!codagent, "N", "S") & "," & RS!codforpa & ",1,"
            
            
            'codopera,codconce340,codintra
            '*****
            ' Tipo de operacion
            '  GENERAL // INTRACOMUNITARIA // EXPORT. - IMPORT. //   INTERIOR EXENTA   // INV. SUJETO PASIVO   // R.E.A.
            'Si es una factura con IVA 0%
            If RS!porciva1 = 0 And IsNull(RS!porciva2) And IsNull(RS!porciva3) Then
                'IVA ES CERO
                Aux = DBLet(RS!codpais, "T")
                If Aux = "" Then Aux = "ES"
                
                
                If Aux = "ES" Then
                    'NACIONAL. Facturas exenta de iva
                    TipoOpera = 3
                    QueTipoDeIVA = 2
                Else
                    Aux = DevuelveDesdeBD(conConta, "intracom", "paises", "codpais", Aux, "T")
                    If Aux = "1" Then
                        'intracomunitaria
                        TipoOpera = 1
                    Else
                        'Exstranjero
                        TipoOpera = 2
                    End If
                    QueTipoDeIVA = 2
                End If
            Else
                'Factura NORMAL
                TipoOpera = 0
            End If
            
            'Concepto 340
            '---------------------
            ' 0 Habitual                B  Ticuet agrupado         C  Varios tipos impositivos
            ' D Rectificativa           I Sujeto pasivo             J Tikets
            ' P adquisiciones de bienes y servicios
            Select Case RS!Codtipom
            Case "FTG"
                Aux = "B"
            Case "FTI"
                Aux = "J"
            Case "FRT"
                Aux = "D"
            Case Else
                'HABITUAL
                If Not IsNull(RS!porciva2) Then
                    Aux = "C" 'varios tipos de iVA
                Else
                    Aux = "0"
                End If
            End Select
            
            'codopera,codconce340,codintra
            SQL = SQL & TipoOpera & "," & DBSet(Aux, "T") & ","
            Aux = ValorNulo
            If TipoOpera = 1 Then Aux = "'E'" 'Entregas intracomunitarias extenas de IVA
            SQL = SQL & Aux
            
            
            'Lineas importes iva
            'factcli_totales(numserie,numfactu,fecfactu,anofactu........
            
        
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            SQL2 = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & Year(RS!Fecfactu) & ","
            SQL2 = SQL2 & "1," & DBSet(RS!baseimp1, "N") & "," & RS!codigiv1 & "," & DBSet(RS!porciva1, "N") & ","
            SQL2 = SQL2 & DBSet(RS!porciva1re, "N") & "," & DBSet(RS!imporiv1, "N") & "," & DBSet(RS!imporiv1re, "N")
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
            
            'para las lineas
            vTipoIva(0) = RS!codigiv1
            vPorcIva(0) = RS!porciva1
            vPorcRec(0) = DBLet(RS!porciva1re, "N")
            vImpIva(0) = RS!imporiv1
            vImpRec(0) = DBLet(RS!imporiv1re, "N")
            vBaseIva(0) = RS!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(RS!porciva2) Then
                SQL2 = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & Year(RS!Fecfactu) & ","
                SQL2 = SQL2 & "2," & DBSet(RS!baseimp2, "N") & "," & RS!codigiv2 & "," & DBSet(RS!porciva2, "N") & ","
                SQL2 = SQL2 & DBSet(RS!porciva2re, "N") & "," & DBSet(RS!imporiv2, "N") & "," & DBSet(RS!imporiv2re, "N")
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(1) = RS!codigiv2
                vPorcIva(1) = RS!porciva2
                vPorcRec(1) = DBLet(RS!porciva2re, "N")
                vImpIva(1) = RS!imporiv2
                vImpRec(1) = DBLet(RS!imporiv2re, "N")
                vBaseIva(1) = RS!baseimp2
            End If
            If Not IsNull(RS!porciva3) Then
                SQL2 = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & Year(RS!Fecfactu) & ","
                SQL2 = SQL2 & "3," & DBSet(RS!baseimp3, "N") & "," & RS!codigiv3 & "," & DBSet(RS!porciva3, "N") & ","
                SQL2 = SQL2 & DBSet(RS!porciva3re, "N") & "," & DBSet(RS!imporiv3, "N") & "," & DBSet(RS!imporiv3re, "N")
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(2) = RS!codigiv3
                vPorcIva(2) = RS!porciva3
                vPorcRec(2) = DBLet(RS!porciva3re, "N")
                vImpIva(2) = RS!imporiv3
                vImpRec(2) = DBLet(RS!imporiv3re, "N")
                vBaseIva(2) = RS!baseimp3
            End If
            
        Else
    
        
        
            'SQL = SQL & "," & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & ","
            SQL = SQL & "," & DBSet(RS!baseimp1, "N") & ","
            If IsNull(RS!baseimp2) Then
                If Nulo2 = "S" Then
                    SQL = SQL & DBSet(RS!baseimp2, "N", Nulo2) 'Que ponga un NULL
                Else
                    'NO puede ser NULO. Que ponga un cero
                    SQL = SQL & DBSet(0, "N", Nulo2)
                End If
            Else
                SQL = SQL & DBSet(RS!baseimp2, "N", Nulo2)
            End If
            
            
            
            SQL = SQL & "," & DBSet(RS!baseimp3, "N", Nulo3)
            
            
            SQL = SQL & "," & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3)
            
            
            'RECARGO EQUIVALENCIA
            'ANTES
            'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo
            'AHORA
            SQL = SQL & "," & DBSet(RS!porciva1re, "N", "S") & "," & DBSet(RS!porciva2re, "N", "S") & "," & DBSet(RS!porciva3re, "N", "S")
            
            'Nuevo Abril 2009
            SQL = SQL & "," & DBSet(RS!imporiv1, "N", "N") & ","
            
            
            
            If IsNull(RS!imporiv2) Then
                If Nulo2 = "S" Then
                    SQL = SQL & DBSet(RS!imporiv2, "N", Nulo2) 'Que ponga un NULL
                Else
                    'NO puede ser NULO. Que ponga un cero
                    SQL = SQL & DBSet(0, "N", Nulo2)
                End If
            Else
                SQL = SQL & DBSet(RS!imporiv2, "N", Nulo2)
            End If
            
            
            
            SQL = SQL & "," & DBSet(RS!imporiv3, "N", Nulo3)
            
            'ANTES
            'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & "," & DBSet(RS!imporiv1re, "N", "S") & "," & DBSet(RS!imporiv2re, "N", "S") & "," & DBSet(RS!imporiv3re, "N", "S") & ","
            
            
            SQL = SQL & DBSet(RS!TotalFac, "N") & "," & DBSet(RS!codigiv1, "N") & "," & DBSet(RS!codigiv2, "N", Nulo2) & "," & DBSet(RS!codigiv3, "N", Nulo3) & ","
            
            'INTRACOM
            If RS!TipoIVA = 3 Then
                'Tipo de iva intrcomunitatro
                SQL = SQL & "1"
            Else
                SQL = SQL & "0"
            End If
            
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(RS!Fecfactu, "F")
        End If 'NuevaContab
        cad = cad & "(" & SQL & ")"
    
    
    End If
    RS.Close
    Set RS = Nothing
    
    
    'Insertar en la contabilidad
        'Insertar en la contabilidad
    If vParamAplic.ContabilidadNueva Then
        SQL = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,totbases ,totbasesret,totivas,totrecargo,"
        SQL = SQL & "totfaccl,retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla"
        SQL = SQL & ",despobla,desprovi,nifdatos,codpais,dpto,codagente,codforpa,escorrecta,codopera,codconce340,codintra) "
    
    Else
    
    
        SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
        SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
        SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    End If
    SQL = SQL & " VALUES " & cad
    ConnConta.Execute SQL
    
    
    If vParamAplic.ContabilidadNueva Then
        SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
         ConnConta.Execute SQL
    End If

    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        cadErr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function



'Private Function InsertarLinFact(cadTabla As String, cadWhere As String, cadErr As String, Optional numRegis As Long) As Boolean
''cadWHere: selecciona un registro de scafac
''codtipom=x and numfactu=y and fecfactu=z
'Dim SQL As String
'Dim SQLaux As String
'Dim SQL2 As String
'Dim RS As ADODB.Recordset
'Dim Cad As String, Aux As String
'Dim I As Byte
'Dim TotImp As Currency, ImpLinea As Currency
'
'    On Error GoTo EInLinea
'
'    If cadTabla = "scafac" Then
'        SQL = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
'        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
'        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "slifac")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    Else
'        SQL = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
'        SQL = SQL & " FROM (slifpc  "
'        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "slifpc")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    End If
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Cad = ""
'    I = 1
'    TotImp = 0
'    SQLaux = ""
'    While Not RS.EOF
'        SQLaux = Cad
'        'calculamos la Base Imp del total del importe para cada cta cble ventas
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CalcularPorcentaje(RS!Importe, DtoPPago, 2)
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CalcularPorcentaje(RS!Importe, DtoGnral, 2)
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
'
'        'concatenamos linea para insertar en la tabla de conta.linfact
'        SQL = ""
'        SQL2 = ""
'        If cadTabla = "scafac" Then
'            SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & "," & I & ","
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
'        Else
'            SQL = numRegis & "," & Year(RS!FecFactu) & "," & I & ","
'            If ImpLinea >= 0 Then
'                SQL = SQL & DBSet(RS!ctacompr, "T")
'            Else
'                SQL = SQL & DBSet(RS!abocompr, "T")
'            End If
'        End If
'        SQL2 = SQL & ","
'        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
'
'        If CCoste = "" Then
'            SQL = SQL & ValorNulo
'        Else
'            SQL = SQL & DBSet(CCoste, "T")
'        End If
'
'        Cad = Cad & "(" & SQL & ")" & ","
'
'        I = I + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If TotImp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        TotImp = BaseImp - TotImp
'        TotImp = ImpLinea + TotImp '(+- diferencia)
'        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
'        If CCoste = "" Then
'            SQL2 = SQL2 & ValorNulo
'        Else
'            SQL2 = SQL2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & SQL2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & SQL2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If
'
'
'    'Insertar en la contabilidad
'    If Cad <> "" Then
'        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
'        If cadTabla = "scafac" Then
'            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
'        Else
'            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
'        End If
'        SQL = SQL & " VALUES " & Cad
'        ConnConta.Execute SQL
'    End If
'
'EInLinea:
'    If Err.Number <> 0 Then
'        InsertarLinFact = False
'        cadErr = Err.Description
'    Else
'        InsertarLinFact = True
'    End If
'End Function
'

'TipoIVAfra:
Private Function InsertarLinFact(cadTabla As String, cadWhere As String, cadErr As String, vLlevaRetencion As Boolean, numRegis As Long, Intracom As String, TipoIvaFactura As Byte) As Boolean
    If vParamAplic.ContabilidadNueva Then
        InsertarLinFact = InsertarLinFactContaNueva(cadTabla, cadWhere, cadErr, vLlevaRetencion, numRegis, Intracom, TipoIvaFactura)
    Else
        InsertarLinFact = InsertarLinFactContaAntigua(cadTabla, cadWhere, cadErr, vLlevaRetencion, numRegis)
    End If
End Function



Private Function InsertarLinFactContaAntigua(cadTabla As String, cadWhere As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String


    On Error GoTo EInLinea
    

    '
    '   Habra que ver en funcion de CC que tenga si agrupo, o no, por  codtraba
    '
    If cadTabla = "scafac" Then 'VENTAS
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        'JULIO 2011
        If vParamAplic.ContabilizacionMoixent Then cadCampo = Replace(cadCampo, "sfamia", "sfamcontab")
        
        SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",CodTraba"
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        
        
        'Aunque en teoria SOLO lo utiliza moixent, lo dejare preparado
        If vParamAplic.ContabilizacionMoixent Then
            SQL = SQL & " inner join sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
            SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
            
        Else
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        End If
     
        
        'David.
        'Lleva anal. Necesitare el trabajador para obtener el CC
        If vCCos > 0 Then SQL = SQL & " ,scafac1 "
        
        SQL = SQL & " WHERE "
        
        'Si lleva analitica
        If vCCos > 0 Then
            'Linkamos la tabla
            SQL = SQL & " slifac.codTipoM = scafac1.codTipoM And slifac.NumFactu = scafac1.NumFactu And slifac.FecFactu = scafac1.FecFactu"
            SQL = SQL & " and slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar AND "
        End If
        
        SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
    Else 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sfamia.ctacompr"
        Else
            cadCampo = "sfamia.abocompr"
        End If
        
        'JULIO 2011
        If vParamAplic.ContabilizacionMoixent Then cadCampo = Replace(cadCampo, "sfamia", "sfamcontab")
        
        
        
        SQL = "SELECT slifpc.codprove,slifpc.numfactu,slifpc.fecfactu," & cadCampo & " as cuenta, sum(importel) as importe  "
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",CodTrab2 as codtraba"
                
        
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        
        
        'Aunque en teoria SOLO lo utiliza moixent, lo dejare preparado
        If vParamAplic.ContabilizacionMoixent Then
            SQL = SQL & " inner join sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
            SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
        Else
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        End If

        
        If vCCos > 0 Then SQL = SQL & ",scafpa "
        
        SQL = SQL & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
            
        SQL = SQL & Replace(cadWhere, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
        
        
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    TotImp = 0
    SQLaux = ""
    Aux = ""
    While Not RS.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        
        If cadTabla = "scafac" Then 'VENTAS a clientes
            'En aux guardaremos el trozo comun de las lineas (letra/numero/anño
            If Aux = "" Then Aux = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!Fecfactu) & ","
            SQL = Aux & I & ","
            SQL = SQL & DBSet(RS!Cuenta, "T")

        Else 'COMPRAS
            'Laura 24/10/2006
            'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
            SQL = numRegis & "," & AnyoFacPr & "," & I & ","
            
'            If ImpLinea >= 0 Then
                SQL = SQL & DBSet(RS!Cuenta, "T")
'            Else
'                SQL = SQL & DBSet(RS!abocompr, "T")
'            End If
        End If
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If vCCos = 0 Then
            SQL = SQL & ValorNulo
        Else
            'Obtendremos el centro de coste a partir del trabajador
            CCoste2 = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", RS!CodTraba)
            If CCoste2 = "" Then
                cadErr = "ERROR en el centro de coste del trabajador: " & RS!CodTraba
                'CIerro el rs y salgo por patas
                RS.Close
                Set RS = Nothing
    
            End If
            SQL = SQL & DBSet(CCoste2, "T")
        End If
        
'        If CCoste = "" Then
'            SQL = SQL & ValorNulo
'        Else
'            SQL = SQL & DBSet(CCoste, "T")
'        End If
        
        cad = cad & "(" & SQL & ")" & ","
        
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If vLlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        I = I + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        cad = cad & SQL
        I = I + 1
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        cad = cad & SQL
        
    End If

    
    
    
    'Facturas clientes. Ver si lleva aportacion al terminal
    If cadTabla = "scafac" Then
        If DatosAportacion <> "" Then
            
            
            SQL = "(" & Aux & I & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
            'Dejo en DatosAportacion solo el importe
            DatosAportacion = TransformaComasPuntos(RecuperaValor(DatosAportacion, 2))
            SQL = SQL & DatosAportacion & ",NULL),"
            cad = cad & SQL
            I = I + 1                                                                                   'Importe en negativo
            SQL = "(" & Aux & I & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
            cad = cad & SQL
        
        
        
    
        End If
    End If

    Set RS = Nothing

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafac" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        SQL = SQL & " VALUES " & cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactContaAntigua = False
        cadErr = Err.Description
    Else
        InsertarLinFactContaAntigua = True
    End If
End Function







Private Function ActualizarCabFact(cadTabla As String, cadWhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE " & cadTabla & " SET intconta=1 "
    SQL = SQL & " WHERE " & cadWhere

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS PROVEEDOR
'----------------------------------------------------------------------
'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador


'Ahora la retencion puede llevarla CUALQUIERA de las facturas.
'   0. Retencion NORMAL
'   1. Retencion SOCIOS

Public Function PasarFacturaProv(cadWhere As String, CodCCost As Byte, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada

'Modificacion Enero2008.   Tipo cooperativas (Ej. Terrasana)
'                          Si lleva retencion la factura, y el preoveedore es tipo REA
'                          entonces  a la contabilidad
'                          El importe de la factura es totfac + retencion
'                          y a las lineas van dos lineas mas
'                          proveedor     -impret
'                          ctareten      +impret

Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim Mc As Contadores
Dim vLlevaRetencion As Boolean
Dim I As Integer

Dim TipoIvaFactura As Byte '0 Normal   1 R.E     2 Exento

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    vLlevaRetencion = False 'Si llevara retencion me lo devolvera la fucion insertar
    TipoIvaFactura = 0
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactProv(cadWhere, cadMen, Mc, FechaFin, vLlevaRetencion, vContaFra, TipoIvaFactura)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        
        'Veremos que opcion de CC es la que hay que pasar (agrupar o no agrupar)
        vCCos = CodCCost
        '---- Insertar lineas de Factura en la Conta
        b = InsertarLinFact("scafpc", cadWhere, cadMen, vLlevaRetencion, Mc.contador, "", TipoIvaFactura) 'El ultinmo seria intracom
        
        cadMen = "Insertando Lin. Factura: " & cadMen

        
        If b Then
            If vContaFra.RealizarContabilizacion Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            End If
        End If
        
        If b Then
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("scafpc", cadWhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
        

        
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaProv = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaProv = False

        InsertarTMPErrFac cadMen, cadWhere
        
        'Si es correcto entonces creo una entrada en tmp para luego listar los resultados de
        'la contabilizacion
         If Mc.contador > 0 Then
            SQL = "DELETE from tmpinformes where codusu = " & vUsu.Codigo & " AND codigo1= " & Mc.contador
            conn.Execute SQL
        End If
    
    End If
End Function


Private Function InsertarCabFactProv(cadWhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef LlevaRetencion As Boolean, ByRef vCF As cContabilizarFacturas, ByRef QueTipoDeIVA As Byte) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Aux As String
Dim TipoOpera  As String
Dim SQL2 As String
Dim CadenaInsertFaclin2 As String
Dim ImporAux As Currency
Dim CodIntra As String
    
    On Error GoTo EInsertar
       
    
    SQL = SQL & " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,sprove.codmacta,"
    SQL = SQL & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "totalfac,tipoiva1,tipoiva2,tipoiva3,tipprove,impret,scafpc.nomprove,scafpc.codprove,tiporet,PorRet,impret "   'Modificacion facturas socios
    
    If vParamAplic.ContabilidadNueva Then SQL = SQL & " ,scafpc.nomprove,scafpc.domprove,scafpc.codpobla,scafpc.pobprove,scafpc.proprove,scafpc.nifprove,scafpc.codforpa,codpais"
    
    SQL = SQL & " FROM " & "scafpc "
    SQL = SQL & "INNER JOIN " & "sprove ON scafpc.codprove=sprove.codprove "
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not RS.EOF Then
    
        If Mc.ConseguirContador("1", (RS!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
                
                    
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = RS!DtoPPago
            DtoGnral = RS!DtoGnral
            BaseImp = RS!BaseIVA1 + CCur(DBLet(RS!BaseIVA2, "N")) + CCur(DBLet(RS!BaseIVA3, "N"))
            TotalFac = RS!TotalFac
            AnyoFacPr = RS!anofacpr
            
            'Para que contabilice las facturas automaticamente
            If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Mc.contador, AnyoFacPr, ""
            
            'SI es facutra socio y tiene retencion
            DatosRetencion = ""
            If RS!TipoRet = 1 Then 'FACTURA SOCIO, con retencion
                If DBLet(RS!ImpRet, "N") <> 0 Then
                    'El total factura es totafac+ retencion
                    DatosRetencion = RS!Codmacta & "|" & RS!ImpRet & "|"
                    TotalFac = TotalFac + RS!ImpRet  'Luego en las lineas va la resta de este importe
                    LlevaRetencion = True
                Else
                    DatosRetencion = ""
                End If
            End If
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(RS!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(RS!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            SQL = ""
            If vParamAplic.ContabilidadNueva Then SQL = "'" & SerieFraPro & "',"
            SQL = SQL & Mc.contador & "," & DBSet(RS!Fecfactu, "F") & "," & RS!anofacpr & "," & DBSet(RS!FecRecep, "F") & ","
            If vParamAplic.ContabilidadNueva Then SQL = SQL & DBSet(RS!FecRecep, "F") & ","
            SQL = SQL & DBSet(RS!NumFactu, "T") & "," & DBSet(RS!Codmacta, "T") & ","
            
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'Nº Factura
                SQL = SQL & "'" & DevNombreSQL("S/Fra " & RS!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            
            
            If Not vParamAplic.ContabilidadNueva Then

                
                SQL = SQL & "," & DBSet(RS!BaseIVA1, "N") & "," & DBSet(RS!BaseIVA2, "N", "S") & "," & DBSet(RS!BaseIVA3, "N", "S") & ","
                SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3) & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!impoiva1, "N") & "," & DBSet(RS!impoiva2, "N", Nulo2) & "," & DBSet(RS!impoiva3, "N", Nulo3) & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                'ANTES era dbset de Rs!totalfac, ahora lo haremos de la variabele totalfac
                SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(RS!TipoIVA1, "N") & "," & DBSet(RS!TipoIVA2, "N", Nulo2) & "," & DBSet(RS!TipoIVA3, "N", Nulo3) & ",0,"
                
            
                'RETENCION.   29 MAYO 2008
                ' retfacpr,trefacpr,cuereten              Las facturas pueden llevar retencion
                Nulo2 = ""
                If RS!TipoRet = 0 Then
                    If Not IsNull(RS!PorRet) And Not IsNull(RS!ImpRet) Then Nulo2 = "O"
                End If
                If Nulo2 = "" Then
                    'NULOS
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Else
                    'TIene valor
                    SQL = SQL & DBSet(RS!PorRet, "N") & "," & DBSet(RS!ImpRet, "N") & ",'" & vParamAplic.CtaReten & "',"
                End If
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!FecRecep, "F") & ",0"
                
            
            
            
            Else
                'Contabilidad NUEVA
                'fecliqcl,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,dpto,codagente,codforpa,escorrecta,
                SQL = SQL & "," & DBSet(RS!nomprove, "T") & "," & DBSet(RS!domprove, "T", "S") & ","
                SQL = SQL & DBSet(RS!codpobla, "T", "S") & "," & DBSet(RS!pobprove, "T", "S") & "," & DBSet(RS!proprove, "T", "S") & ","
                SQL = SQL & DBSet(RS!nifProve, "T", "S") & "," & DBSet(RS!codpais, "T", "S") & ","
                SQL = SQL & RS!codforpa & ","
                
  
                'codopera,codconce340,codintra
                '*****
                ' Tipo de operacion
                ' 0 General   1 Intracom    2  Export import    3 Interior exenta    4   ISP    5 REA
                '  GENERAL // INTRACOMUNITARIA // EXPORT. - IMPORT. //   INTERIOR EXENTA   // INV. SUJETO PASIVO   // R.E.A.
                'Si es una factura con IVA 0%
                TipoOpera = 0
                CodIntra = ""
                'IVA ES CERO
                If RS!tipprove = 1 Then
                    'intracomunitaria
                    TipoOpera = 1
                    CodIntra = "A"
                    QueTipoDeIVA = 2 'exento
                Else
                    'Exstranjero
                     If RS!tipprove = 2 Then
                        TipoOpera = 2
                        QueTipoDeIVA = 2 'exento
                    End If
                End If
                
                'Concepto 340
                '---------------------
                ' 0 Habitual                 C  Varios tipos impositivos
                ' D Rectificativa           I Sujeto pasivo
                ' P adquisiciones intracomunitarias de bienes y servicios
                'IMPORTACION(  NO salen en el 340)
                Aux = "0"
                If RS!TotalFac < 0 Then
                    Aux = "D"
                Else
                    If RS!tipprove = 1 Then
                        Aux = "P"
                    Else
                        If Not IsNull(RS!TipoIVA2) Then Aux = "C"
                    End If
                End If
            
                
                'codopera,codconce340,codintra
                SQL = SQL & TipoOpera & "," & DBSet(Aux, "T") & "," & DBSet(CodIntra, "T", "S") & ","
                
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.contador & "," & DBSet(RS!FecRecep, "F") & "," & RS!anofacpr & ","
                
                SQL2 = Aux & "1," & DBSet(RS!BaseIVA1, "N") & "," & RS!TipoIVA1 & "," & DBSet(RS!porciva1, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(RS!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
                vTipoIva(0) = RS!TipoIVA1
                vPorcIva(0) = RS!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = RS!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = RS!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(RS!porciva2) Then
                    SQL2 = Aux & "2," & DBSet(RS!BaseIVA2, "N") & "," & RS!TipoIVA2 & "," & DBSet(RS!porciva2, "N") & ","
                    SQL2 = SQL2 & ValorNulo & "," & DBSet(RS!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                    vTipoIva(1) = RS!TipoIVA2
                    vPorcIva(1) = RS!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = RS!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = RS!BaseIVA2
                
                End If
                If Not IsNull(RS!porciva3) Then
                    SQL2 = Aux & "3," & DBSet(RS!BaseIVA3, "N") & "," & RS!TipoIVA3 & "," & DBSet(RS!porciva3, "N") & ","
                    SQL2 = SQL2 & ValorNulo & "," & DBSet(RS!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                    vTipoIva(2) = RS!TipoIVA3
                    vPorcIva(2) = RS!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = RS!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = RS!BaseIVA3
                End If
                
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = RS!BaseIVA1 + DBLet(RS!BaseIVA2, "N") + DBLet(RS!BaseIVA3, "N")
                SQL = SQL & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = RS!impoiva1 + DBLet(RS!impoiva2, "N") + DBLet(RS!impoiva3, "N")
                SQL = SQL & DBSet(ImporAux, "N") & "," & DBSet(RS!TotalFac, "N") & ","
                        
                
                  
                  
                  
                'EsFacturaIntracom2 = ""
                'If DBLet(RS!tipprove, "N") = 1 Then
                '    'OK es intracomunitaria
                '    EsFacturaIntracom2 = RS!TipoIVA1
                'End If
                  
                Aux = ""
                If DatosRetencion = "" Then
                    
                    'NULOS
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    If vParamAplic.ContabilidadNueva Then SQL = SQL & "0"
                Else
                    ' retfacpr , trefacpr, cuereten,     tiporeten   'SOLO EN LA NUEVA
                    'TIene valor
                    SQL = SQL & DBSet(RecuperaValor(DatosRetencion, 2), "T") & "," & DBSet(RecuperaValor(DatosRetencion, 1), "N") & ",'" & vParamAplic.CtaReten & "',"
                    If vParamAplic.ContabilidadNueva Then SQL = SQL & "0"
                End If
            
            
            
            
            End If
            cad = cad & "(" & SQL & ")"
            
            'Insertar en la contabilidad
            If vParamAplic.ContabilidadNueva Then
                SQL = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                SQL = SQL & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                SQL = SQL & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten)"
            Else
            
                SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                
            End If
            SQL = SQL & " VALUES " & cad
            ConnConta.Execute SQL
            
            
            If vParamAplic.ContabilidadNueva Then
                'Las  lineas de IVA
                SQL = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                SQL = SQL & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute SQL
                
            End If
            
            
            
            
            'Para saber el numreo de registro que le asigna a la factrua
            SQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.contador
            SQL = SQL & ",'" & DevNombreSQL(RS!NumFactu) & " @ " & Format(RS!Fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(RS!nomprove) & "'," & RS!codProve & ")"
            conn.Execute SQL
        End If
    End If
    RS.Close
    Set RS = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Public Sub FechasEjercicioConta(FIni As String, FFin As String)
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EFechas
'
'    FIni = "Select fechaini,fechafin From parametros"
'    Set RS = New ADODB.Recordset
'    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        FIni = DBLet(RS!FechaIni, "F")
'        FFin = DBLet(RS!FechaFin, "F")
'    End If
'    RS.Close
'    Set RS = Nothing
'
'EFechas:
'    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function InsertarLinFact_TicketsAgrupados(cadTabla As String, cadWhere As String, cadErr As String, LlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String


    On Error GoTo EInLinea
    
        
    
            
            
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        
        
        'Monto el WHERE buscando los tikets que estan asociados a este numfact FTG
        SQLaux = Replace(cadWhere, "scafac.", "")
        SQLaux = Replace(SQLaux, "numfactu", "numfacftg")
        SQLaux = Replace(SQLaux, "fecfactu", "fecfacftg")
        SQLaux = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND " & SQLaux
    
    
    
    
        Set RS = New ADODB.Recordset
        RS.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = RS!NumFacftg & " as numfactu ,'" & Format(RS!FecFacftg, FormatoFecha) & "' as fecfactu,"
        'En aux guardare el codtraba
        Aux = RS!CodTraba
        SQLaux = ""
        Do
            SQLaux = SQLaux & "," & RS!NumFactu
            RS.MoveNext
        Loop Until RS.EOF
        RS.Close
        
        
        
        
        
        SQL = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FTG", "T")
        SQL = " SELECT '" & SQL & "' as LetraSer,slifac.codtipom," & cad & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & "," & Aux & " as CodTraba"
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        'David.
        'Lleva anal. Necesitare el trabajador para obtener el CC
        If vCCos > 0 Then SQL = SQL & " ,scafac1 "
        
        SQL = SQL & " WHERE "
        
        'Si lleva analitica
        If vCCos > 0 Then
            'Linkamos la tabla
            SQL = SQL & " slifac.codTipoM = scafac1.codTipoM And slifac.NumFactu = scafac1.NumFactu And slifac.FecFactu = scafac1.FecFactu"
            SQL = SQL & " and slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar AND "
        End If
        

        
                
        
        
        
        SQLaux = Mid(SQLaux, 2)
        SQLaux = "   slifac.codtipom='FTI' AND slifac.numfactu IN (" & SQLaux & ")"
        SQL = SQL & SQLaux
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
    
    

    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    TotImp = 0
    SQLaux = ""
    While Not RS.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        

        SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!Fecfactu) & "," & I & ","
        SQL = SQL & DBSet(RS!Cuenta, "T")
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If vCCos = 0 Then
            SQL = SQL & ValorNulo
        Else
            'Obtendremos el centro de coste a partir del trabajador
            CCoste2 = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", RS!CodTraba)
            If CCoste2 = "" Then
                cadErr = "ERROR en el centro de coste del trabajador: " & RS!CodTraba
                'CIerro el rs y salgo por patas
                RS.Close
                Set RS = Nothing
    
            End If
            SQL = SQL & DBSet(CCoste2, "T")
        End If
        
        
        cad = cad & "(" & SQL & ")" & ","
        
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If LlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        I = I + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        cad = cad & SQL
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & I + 1 & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        cad = cad & SQL
    End If





    Set RS = Nothing

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        SQL = SQL & " VALUES " & cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_TicketsAgrupados = False
        cadErr = Err.Description
    Else
        InsertarLinFact_TicketsAgrupados = True
    End If
End Function
















'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'
'   Contabilizacion de las facturas internas (FAI)
'
'   - No inserta en cabfact(ni linfact).  Mete un apunte
'       43000   contra las 70000 que deriven de las familias
'
Private Function ContabilizaFAI(cadError As String, cadWhere As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
Dim Rc As ADODB.Recordset
Dim RL As ADODB.Recordset
Dim SQL As String
Dim Aux As String

    On Error GoTo eContabilizaFAI
    ContabilizaFAI = False
    cadError = ""
    SQL = " SELECT stipom.letraser,numfactu,fecfactu, sclien.codmacta,sclien.cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,tipoiva"
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWhere
    Set Rc = New ADODB.Recordset
    Rc.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DtoPPago = Rc!DtoPPago
    DtoGnral = Rc!DtoGnral
    BaseImp = Rc!baseimp1
    TotalFac = Rc!TotalFac
    DatosAportacion = ""
    conCtaAlt = Rc!cliAbono
    
    
    
    'Para las lineas de factura
    '-------------------------------
    If conCtaAlt Then
        'utilizamos sfamia.ctavent1 o sfamia.abovent1
        If TotalFac >= 0 Then
            Aux = "sfamia.ctavent1"
        Else
            Aux = "sfamia.abovent1" 'si es negativa es un abono
        End If
    Else
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            Aux = "sfamia.ctaventa"
        Else
            Aux = "sfamia.aboventa"
        End If
    End If
    
    
    'JULIO 2011
    If vParamAplic.ContabilizacionMoixent Then Aux = Replace(Aux, "sfamia", "sfamcontab")
        
    
    
    
    SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & Aux & " as cuenta,sum(importel) as importe"
    'Tiene analitica. Luego el codtraba tiene que aparecer
    If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
    SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
    SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
    
    'Aunque en teoria SOLO lo utiliza moixent, lo dejare preparado
    If vParamAplic.ContabilizacionMoixent Then
        SQL = SQL & " inner join sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
        SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
        
    Else
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
    End If
    SQL = SQL & " WHERE "
    SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
    SQL = SQL & " GROUP BY "
    'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
    If vCCos > 0 Then SQL = SQL & " codccost, "
    'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
    SQL = SQL & Aux
    Set RL = New ADODB.Recordset
    RL.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    cadError = vContaFra.IntegraLaFacturaClienteINTERNA(Rc, RL)

    
eContabilizaFAI:
    If Err.Number <> 0 Then
        cadError = Err.Description
        Err.Clear
    End If
    If cadError = "" Then ContabilizaFAI = True
    
End Function











'Nueva contabiliacion lineas factura
'FraIntraCom2. Si es <>"" entonces lleva el tipo de iva que va la factura
'TipoIVA 0 Normal   1 R.E     2 Exento
Private Function InsertarLinFactContaNueva(cadTabla As String, cadWhere As String, cadErr As String, vLlevaRetencion As Boolean, numRegis As Long, FraIntraCom As String, TipoIvaFra As Byte) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim cadCampo As String
Dim LineaCentroCoste As Boolean
Dim NumeroIVA As Byte
Dim ImpImva As Currency
Dim ImpRec As Currency
Dim ImpLinea As Currency
Dim HayQueAjustar As Boolean
Dim K As Byte
Dim PrimerCodigiva As Integer
Dim IvaABuscar As Integer
    On Error GoTo EInLinea
    

    'Agrupa tambien por tipo de iva
    
    If cadTabla = "scafac" Then 'VENTAS
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        
        
        'PARA las FAS, si viene la cuenta de venta, leida desde advparametros, entonces pone esa
        If InStr(1, cadWhere, "'FAS'") > 0 Then
            If conCtaAlt Then
                cadCampo = "sfamia.ctavent1"
            Else
                'cadCampo = "sfamia.ctavtaseralt"
                'Juli2020
                cadCampo = "sfamia.ctavent1"
            End If
        End If
        
        'JULIO 2011
        If vParamAplic.ContabilizacionMoixent Then cadCampo = Replace(cadCampo, "sfamia", "sfamcontab")
        
            
        SQL = "SELECT "
        
        If TipoIvaFra = 2 Then SQL = SQL & " 1 " 'Para que solo meta un tipo iva. Lo reempplazaara por el que le toque
        SQL = SQL & " codigiva,stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
        
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        
        
        
        'Aunque en teoria SOLO lo utiliza moixent, lo dejare preparado
        If vParamAplic.ContabilizacionMoixent Then
            SQL = SQL & " inner join sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
            SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
            
        Else
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        End If
        
        
        
        

        
        
        SQL = SQL & " WHERE "
        
        
        SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos > 0 Then SQL = SQL & " codccost, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo & ", codigiva ORDER BY codigiva ," & cadCampo
        
    Else 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sfamia.ctacompr"
        Else
            cadCampo = "sfamia.abocompr"
        End If
        
        
        If vParamAplic.ContabilizacionMoixent Then cadCampo = Replace(cadCampo, "sfamia", "sfamcontab")
        
        
        If FraIntraCom <> "" Then
            SQL = " SELECT " & FraIntraCom
        Else
            SQL = " SELECT"
        
        End If
        
        SQL = SQL & " codigiva,slifpc.codprove,slifpc.numfactu,FecFactu," & cadCampo & " as cuenta, sum(importel) as importe  "
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",CodTrab2 as codtraba"
                
        
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        
        If vParamAplic.ContabilizacionMoixent Then
            SQL = SQL & " inner join sfamcontab ON sartic.codfamia=sfamcontab.codfamia and sartic.codmarca=sfamcontab.codmarca"
            SQL = SQL & " AND sartic.codtipar=sfamcontab.codtipar and sartic.codunida=sfamcontab.codunida"
        Else
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        End If
        
        If vCCos > 0 Then SQL = SQL & ",scafpa "
        
        SQL = SQL & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
            
        SQL = SQL & Replace(cadWhere, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, , "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo & ", codigiva ORDER BY codigiva ," & cadCampo
        
        
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    SQL2 = ""
    Aux = ""
    PrimerCodigiva = -1
    While Not RS.EOF
        
        'Preparamos todo menos los importes
        'factcli_lineas(
        ' numserie,numfactu,fecfactu,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost
        If cadTabla = "scafac" Then
            'VENTAS a clientes
            SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & Year(RS!Fecfactu) & "," & I & ","
            'Por si lleva datosa`portacion o lo que sea
            Aux = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!Fecfactu, "F") & "," & Year(RS!Fecfactu) & ","
        Else
            'Compras
            'factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost)
            
            SQL = "'" & SerieFraPro & "'," & numRegis & "," & DBSet(RS!Fecfactu, "F") & "," & AnyoFacPr & "," & I & ","
        
            
        End If
        SQL = SQL & DBSet(RS!Cuenta, "T")
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For K = 0 To 2
            
            IvaABuscar = RS!codigiva
            ''0 Normal   1 R.E     2 Exento
            If TipoIvaFra = 1 Then
                If IvaABuscar = vParamAplic.TipoIVA1 Then IvaABuscar = vParamAplic.TipoIVAre1
                If IvaABuscar = vParamAplic.TipoIVA2 Then IvaABuscar = vParamAplic.TipoIVAre2
                If IvaABuscar = vParamAplic.TipoIVA3 Then IvaABuscar = vParamAplic.TipoIVAre3
            Else
                If TipoIvaFra = 2 Then
                    'Solo tiene un IVA
                    IvaABuscar = vTipoIva(K)
                    
                End If
            End If
            
        
        
            If IvaABuscar = vTipoIva(K) Then
                NumeroIVA = K
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, , "Error obteniendo IVA: " & RS!codigiva
        If PrimerCodigiva < 0 Then PrimerCodigiva = K
        
        'Importe
        '----------------------------------------------------------------------------
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            RS.MoveNext
            If RS.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If RS!codigiva <> vTipoIva(0) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            RS.MovePrevious
        End If
        
        'codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost)
        'codigiva,porciva,porcrec,
        SQL = SQL & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            'S top
        
        
        End If
        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpRec = 0
        Else
            ImpRec = vPorcRec(NumeroIVA) / 100
            ImpRec = Round2(ImpLinea * ImpRec, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpRec
        
        'baseimpo , impoiva, imporec, aplicret, CodCCost
        SQL = SQL & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpRec, "N", "S")
        SQL = SQL & ",0,"
        
        
        'CENTRO DE COSTE
        LineaCentroCoste = False
        If vCCos = 0 Then
            'NO NECESTIA CENTRO DE COSTE.. seguro
           ' SQL = SQL & ValorNulo
        Else
            LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(RS!Cuenta))
            
        End If
        If LineaCentroCoste Then
            CCoste2 = RS!CodCCost
            SQL = SQL & DBSet(CCoste2, "T")
        Else
            SQL = SQL & ValorNulo
        End If
        
        
        
        

        
        
        
        cad = cad & "(" & SQL & ")" & ","
        
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close

    
    


    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If vLlevaRetencion Then
        's top
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        I = I + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        cad = cad & SQL
        I = I + 1
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        cad = cad & SQL
        
    End If

    
    
    
    'Facturas clientes. Ver si lleva aportacion al terminal
    If cadTabla = "scafac" Then
        If DatosAportacion <> "" Then
            
             ' numserie,numfactu,fecfactu,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost
            SQL = RecuperaValor(DatosAportacion, 2)
            ImpLinea = CCur(SQL)
            ImpImva = vPorcIva(PrimerCodigiva) / 100
            ImpImva = Round2(ImpLinea * ImpImva, 2)
            If vPorcRec(PrimerCodigiva) = 0 Then
                ImpRec = 0
            Else
                ImpRec = vPorcRec(PrimerCodigiva) / 100
                ImpRec = Round2(ImpLinea * ImpRec, 2)
            End If
            
                       
            SQL = " (" & Aux & I & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
            SQL = SQL & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
            'baseimpo , impoiva, imporec, aplicret, CodCCost
            SQL = SQL & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpRec, "N", "S") & ",0,NULL) ,"
            cad = cad & SQL
            
            'Dejo en DatosAportacion solo el importe
            
            I = I + 1                                                                                   'Importe en negativo
            'SQL = ", (" & Aux & i & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
            SQL = "(" & Aux & I & ",'" & vParamAplic.ctaAportacion & "'"
            SQL = SQL & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
            'baseimpo , impoiva, imporec, aplicret, CodCCost
            SQL = SQL & DBSet(-ImpLinea, "N") & "," & DBSet(-ImpImva, "N") & "," & DBSet(-ImpRec, "N", "S") & ",0,NULL) ,"
        
            cad = cad & SQL
        
        
        
        
        
    
        End If
    End If

    Set RS = Nothing

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafac" Then
            SQL = "INSERT INTO  factcli_lineas(numserie,numfactu,fecfactu,anofactu,numlinea,codmacta,codigiva,porciva,porcrec,"
            SQL = SQL & " baseimpo,impoiva,imporec,aplicret,codccost)"
        Else
            SQL = "INSERT INTO factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,codigiva,porciva,porcrec,"
            SQL = SQL & " baseimpo,impoiva,imporec,aplicret,codccost)"
        End If
        SQL = SQL & " VALUES " & cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactContaNueva = True
    End If
End Function


Private Function CuentaNecesitaCentroCoste(cta As String) As Boolean
Dim I As Integer
Dim C As String
    
    CuentaNecesitaCentroCoste = False
    
'    'vEmpresa.RaizAnalitica    lleva: gripo gasto |grupo vta| otros grupo
'    For i = 1 To 3
'        C = RecuperaValor(vEmpresa.RaizAnalitica, i)
'        If i < 3 Then
'            'UN DIGITO
'            If Mid(Cta, 1, 1) = C Then
'                CuentaNecesitaCentroCoste = True
'                Exit Function
'            End If
'        Else
'            'Subgrupo a tres digitos
'            If Mid(Cta, 1, 3) = C Then
'                CuentaNecesitaCentroCoste = True
'                Exit Function
'            End If
'        End If
'    Next i
    
End Function


