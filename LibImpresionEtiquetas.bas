Attribute VB_Name = "LibImpresionEtiquetas"
Option Explicit


'                                                para rehutilizaar la variable
Public Function ImprimeEtiquetasCajas(linea As Byte, Lote As Long, ByVal Articulo As String, Incio As Long, Cantidad As Long, DatosLineaExtra As String, DatosLineaExtra2 As String) As Boolean
Dim NF As Integer
Dim Cad As String
Dim Conta As Long
Dim Linea1 As String
Dim CodFamilia As String
Dim Aux As String

    On Error GoTo EImprimeEtiquetasCajas:
    ImprimeEtiquetasCajas = False
    
    
    Set miRsAux = New ADODB.Recordset
    Cad = "Select carpetaSRV,pathArchBartender,extension,CajaL" & linea & " impresora FROM  prodparamimpr"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'NO PUEDE SER EOF
    
    'registro es la marca
    CodFamilia = "codfamia"
    Cad = DevuelveDesdeBD(conAri, "codmarca", "sartic", "codartic", Articulo, "T", CodFamilia)
    If Cad = "" Then Err.Raise 513, "Error insesperado obteniendo marca-familia del articulo"
    
    'NUEVO Sept. 2018
    'Tenemos marca famimilia
    'Veremos si hay una etiqueta marca-familia
    Aux = "codmarca = " & Cad & " AND codfamia "
    Aux = DevuelveDesdeBD(conAri, "archivo", "prodparametiq", Aux, CodFamilia)
    
    If Aux <> "" Then
        'Etiqueta especfiica para marca /categoria(famlia)
        Cad = Aux
    Else
        Cad = "codfamia is null and (codmarca = " & Cad & " OR codmarca"
        Cad = DevuelveDesdeBD(conAri, "archivo", "prodparametiq", Cad, " 0) ORDER BY codmarca DESC")
    End If
        
        
    'No deberia pasar, por que la cero la tiene que traer
    If Cad = "" Then Err.Raise 513, "Error insesperado obteniendo nombre archivo etiqueta"
    
    'La primera linea es la de "ordenes"
    'Ejmplo
    '%BTW% /AF="C:\MisDoc\BarTender\Formats\ParaComanderMarca.btw" /P /D="%Trigger File Name%" /PRN="HP DeskJet 1220C" /R=3 /P %END%
    Cad = "%BTW% /AF=""" & miRsAux!pathArchBartender & "\" & Cad & """  /D=""%Trigger File Name%"" /PRN="""
    'Ahora la impresora de la linea       r3: en que linea empiezan los datos
    Cad = Cad & miRsAux!impresora & """ /R=3 /P "
    Linea1 = Cad
    'Print #NF, "%END%"
    
    
    'Nomatic
    Cad = Articulo
    Articulo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Articulo, "T")
    'Quito las comas , por si acaso
    Articulo = Replace(Articulo, ",", ".")
    Articulo = Articulo & ","
    Cad = DevuelveDesdeBD(conAri, "caj_codun", "sarti4", "codartic", Cad, "T")
    If Cad = "" Then Cad = "0"
    Articulo = Articulo & Cad & ","
    
    'Comprobamos que la carpeta esta accesible
    Cad = miRsAux!carpetaSRV & "\*.txt"
    Cad = Dir(Cad, vbArchive) 'Si no estuviera daria error
    
    
    
        
        NF = -1
    
        'Elimino el anterior
        If Dir(App.Path & "\datoseti.txt", vbArchive) <> "" Then Kill App.Path & "\datoseti.txt"
        
        'YA tengo los datos que necesito. Vamo p'alla
        'Creo el archivo en local y luego hago un file copy
        NF = FreeFile
        Open App.Path & "\datoseti.txt" For Output As #NF
        
    
        'Va, primero el nomartic, luego codigodun, y despues... el de la caja
        Print #NF, Linea1
        Print #NF, "%END%"
        
        For Conta = Incio + 1 To Incio + Cantidad
            Cad = Format(Lote, "00000000") & Format(Conta, "00000")
            'Nuevo. 29 Marzo 2012
            'Pondremos separado el numero de lote  y una descripcion que lleva la produccion
            Cad = Cad & "," & Lote & "," & Replace(DatosLineaExtra, ",", "") & "," & Replace(DatosLineaExtra2, ",", "")
            Print #NF, Articulo & Cad
        Next Conta
    
        Close #NF
        NF = -1 'para que no vuelva a hacer close
    
        'Ahora lo copiamos donde diga el path
        Cad = Format(Now, "yymmdd_hhnnss") & vUsu.PC & "."
        Cad = Cad & miRsAux!extension
        Cad = miRsAux!carpetaSRV & "\" & Cad
    
        FileCopy App.Path & "\datoseti.txt", Cad
    
    

    
    ImprimeEtiquetasCajas = True
    
    
    
    miRsAux.Close
    
EImprimeEtiquetasCajas:
    If Err.Number <> 0 Then MuestraError Err.Number
    If NF >= 0 Then Close #NF
    Set miRsAux = Nothing
End Function







Public Function ImprimeEtiquetasMateriaAuxiliar() As Boolean
Dim NF As Integer
Dim Cad As String
Dim Prove As String
Dim Destino As String

    On Error GoTo EImprimeEtiquetasMA:
    ImprimeEtiquetasMateriaAuxiliar = False

    NF = -1
    
    'Elimino el anterior
    If Dir(App.Path & "\datosMA.txt", vbArchive) <> "" Then Kill App.Path & "\datosMA.txt"
    Set miRsAux = New ADODB.Recordset
    Cad = "Select carpetaSRV,pathArchBartender,extension,MateriaAux impresora FROM  prodparamimpr"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Ahora lo copiamos donde diga el path
    Cad = Format(Now, "yymmdd_hhnn") & vUsu.PC & "."
    Cad = Cad & miRsAux!extension
    Destino = miRsAux!carpetaSRV & "\" & Cad
    
    'No deberia pasar, por que la cero la tiene que traer
    'If cad = "" Then Err.Raise 513, "Error insesperado obteniendo nombre archivo etiqueta"
    
        
    'YA tengo los datos que necesito. Vamo p'alla
    'Creo el archivo en local y luego hago un file copy
    NF = FreeFile
    Open App.Path & "\datosMA.txt" For Output As #NF
        
    'La primera linea es la de "ordenes"
    'Ejmplo
    '%BTW% /AF="C:\MisDoc\BarTender\Formats\ParaComanderMarca.btw" /P /D="%Trigger File Name%" /PRN="HP DeskJet 1220C" /R=3 /P %END%
    Cad = "MateriaAux.btw"
    Cad = "%BTW% /AF=""" & miRsAux!pathArchBartender & "\" & Cad & """  /D=""%Trigger File Name%"" /PRN="""
    'Ahora la impresora de la linea       r3: en que linea empiezan los datos
    Cad = Cad & miRsAux!impresora & """ /R=3 /P "
    Print #NF, Cad
    Print #NF, "%END%"
    miRsAux.Close
    
    
    'Los registros estan en la tabla tmppartidas
    Cad = "Select * from tmppartidas where codusu = " & vUsu.Codigo & " order by idpartida,idnumoperacion"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'NO PUEDE SER EOF, si no no hubiera entrado aqui

    Prove = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", miRsAux!idreferencia, "T")
    'Quito las comas , por si acaso
    Prove = " - " & Replace(Prove, ",", ".")
    
    While Not miRsAux.EOF
        'Grabacion en fichero
        '002400113502,"BOTELLA DE CRISTAL","12/04/1972","13/04/1972","123-A","0001-Proveedor",002989001,2989,"ALBARAN","Campo extra2"
        
        'qUE SE TRADUCE EN
        'codartic,Referencia,fecha,NOW,numlote,PROVEE,idnumoperacion,idpartida,idOperacion
        Cad = miRsAux!codartic & ",""" & TransformaComasPuntos(miRsAux!Referencia) & """,""" & Format(miRsAux!Fecha, "dd/mm/yyyy")
        Cad = Cad & """,""" & Format(Now, "dd/mm/yyyy") & """,""" & miRsAux!numLote & ""","""
        
        
        Cad = Cad & Format(miRsAux!idreferencia, "0000") & Prove & """,""" & miRsAux!idnumoperacion & ""","""
        
        
        
        'Nuevo.
        'Modificacion 2012. Pondra etiqueta numeti/toteti
        '
        Cad = Cad & Format(miRsAux!idPartida, "0000") & """,""" & DBLet(miRsAux!idoperacion, "T") & """,""" 'el ultimo vacio
        Cad = Cad & Format(miRsAux!Cantidad, "0000") & "/" & Format(miRsAux!abs_cantidad, "0000") & """,""" & DBLet(miRsAux!idoperacion, "T") & """,""""" 'el ultimo vacio
        
        Print #NF, Cad
    
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    Close #NF
    NF = -1 'para que no vuelva a hacer close
    
    FileCopy App.Path & "\datosMA.txt", Destino
    
    ImprimeEtiquetasMateriaAuxiliar = True
    
EImprimeEtiquetasMA:
    If Err.Number <> 0 Then MuestraError Err.Number
    If NF >= 0 Then Close #NF
    
End Function


'Marzo 2013
'Habran 2 tipos de impresion de palets
' 0,. Etiqueta normal
' 1.- OLIVE LINE

'Mayo 2014.  Se imprimmiran la misma etqiqueta pero sin poner OLIVE LINE
'       Es decir, el proceso sigue desdoblandose igual

Public Function ImprimirPalet(IdPalet As Long, TipoDeImpresion As Byte) As Boolean
    If TipoDeImpresion = 0 Then
        'MAYO 2014
        'ImprimirPaletNORMAL IdPalet
        ImprimirPaletNUEVO IdPalet
    Else
        'OLIVE LINE
        ImprimirPaletOLIVELINE IdPalet
    End If
End Function

'Private Function ImprimirPaletNORMAL(IdPalet As Long) As Boolean
'Dim MismaFechaProduccion As Boolean
'
'
'Dim NF As Integer
'Dim cad As String
'Dim Destino As String
'Dim Aux As String
'Dim C2 As String
'Dim i As Integer
'Dim J As Integer
'Dim Col As Collection
'Dim CC As Byte
'
'    On Error GoTo EImprimeEtiquetasPA:
'    ImprimirPaletNORMAL = False
'
'    NF = -1
'
'    'Elimino el anterior
'    If Dir(App.Path & "\datosMA.txt", vbArchive) <> "" Then Kill App.Path & "\datosMA.txt"
'    Set miRsAux = New ADODB.Recordset
'    cad = "Select carpetaSRV,pathArchBartender,extension,palets impresora FROM  prodparamimpr"
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    'Ahora lo copiamos donde diga el path
'    cad = Format(Now, "yymmdd_hhnnss") & vUsu.PC & "."
'    cad = cad & miRsAux!Extension
'    Destino = miRsAux!carpetaSRV & "\" & cad
'
'
'    NF = FreeFile
'    Open App.Path & "\datosMA.txt" For Output As #NF
'
'    'La primera linea es la de "ordenes"
'    'Ejmplo
'    '%BTW% /AF="C:\MisDoc\BarTender\Formats\ParaComanderMarca.btw" /P /D="%Trigger File Name%" /PRN="HP DeskJet 1220C" /R=3 /P %END%
'    cad = "PaletBu.btw"
'    cad = "%BTW% /AF=""" & miRsAux!pathArchBartender & "\" & cad & """  /D=""%Trigger File Name%"" /PRN="""
'    'Ahora la impresora de la linea       r3: en que linea empiezan los datos
'    cad = cad & miRsAux!impresora & """ /R=3 /P "
'    Print #NF, cad
'    Print #NF, "%END%"
'    miRsAux.Close
'
'
'    'Julio 2012
'    MismaFechaProduccion = True
'
'    Aux = "select fhinicio from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo"
'    Aux = Aux & " AND prodlin.idlin = prodtrazlin.idlin and lotetraza In ("
'    Aux = Aux & "Select idpartida from tmppartidas where codusu = " & vUsu.Codigo & " order by idpartida)"
'    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Aux = ""
'    While Not miRsAux.EOF
'        If Aux = "" Then
'            'Primera aparicion
'            Aux = Format(miRsAux!fhinicio, "dd/mm/yyyy")
'        Else
'            If Aux <> Format(miRsAux!fhinicio, "dd/mm/yyyy") Then MismaFechaProduccion = False
'        End If
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'    'Si NO tienen todos la misma fecha de produccion no la ponemos arriba en el encabezado
'    Aux = "Select fhinicio,fhFin,CajasProd from prodpalets where idpalet = " & IdPalet
'    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'    'Trozo fijo
'    ''"ID12","fecha inicio","100",10000001,"1/2",
'    cad = """" & Format(IdPalet, "0000") & ""","""
'    If MismaFechaProduccion Then cad = cad & Format(miRsAux!fhinicio, "dd/mm/yyyy")   'Si no es la misma fecha no la pondre
'    cad = cad & ""","
'    cad = cad & miRsAux!Cajasprod & ","""
'    cad = cad & "1" & Format(IdPalet, "00000000") & "1"","   'luego pondre si procede el 1/2 o el 4/5....
'    miRsAux.Close
'
'
'    Aux = "Select * from tmppartidas where codusu = " & vUsu.Codigo & " order by idpartida,idnumoperacion"
'    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'    Set Col = New Collection 'Total artic/lotre que iran
'    i = 0
'    While Not miRsAux.EOF
'            i = i + 1
'            '"MAESTRO OLIVA a.o.V.E. MAR O+L 500 ML","(02)18436024291213(37)0035","(10)003676(17)140402","(00)8436024291216",""
'
'            'Marzo 2013
'            'Va el codigo DUN
'            C2 = DevuelveDesdeBD(conAri, "caj_codun", "sarti4", "codartic", miRsAux!codartic, "T")
'            If C2 = "" Then
'                MsgBox "No se ha encontrado el codigo DUN para el articulo: " & DBLet(miRsAux!NumLote, "T"), vbExclamation
'                C2 = DevuelveDesdeBD(conAri, "codigoea", "sartic", "codartic", miRsAux!codartic, "T")
'                If C2 = "" Then C2 = miRsAux!codartic
'            End If
'
'
'            Aux = ",""" & miRsAux!NumLote & ""","""
'
'            C2 = "(02)" & C2 & "(37)" & Format(miRsAux!Cantidad, "000") & """"
'            Aux = Aux & C2
'
'
'
'            C2 = ",""(10)" & Format(miRsAux!IdPartida, "000000") & "(17)" & Format(miRsAux!Fecha, "yymmdd") & ""","
'            Aux = Aux & C2
'
'
'
'
'            'Marzo 2013   [SSCC]
'            'El (00) es el sscc
'            ' (00)38412594xxxxxxxxxC     '38412594=id Morales
'            '      morales  palet  C:ontrol
'            'Dejaremos los 2 priemros digitos  disponibles . Sera de momento 00 para produccion, pero nunca se sabe
'            '          los 7 siguientes seran para idpalet
'
'            C2 = "38412594" & "00" & Format(IdPalet, "0000000")
'            CC = DevuelveDigitoControlSSCC(C2)
'
'
'            C2 = """(00)" & C2 & CC
'            Aux = Aux & C2
'
'
'            C2 = "prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin and lotetraza"
'            C2 = DevuelveDesdeBD(conAri, "fhinicio", "prodlin,prodtrazlin", C2, miRsAux!IdPartida)
'            If C2 <> "" Then C2 = "(11)" & Format(C2, "yymmdd")
'
'            C2 = C2 & ""","""""
'            Aux = Aux & C2
'
'            Col.Add Aux
'
'        miRsAux.MoveNext
'    Wend
'
'    'Es impar, añado una vacia
'    If (i Mod 2) = 1 Then Col.Add """"","""","""","""","""""
'
'
'    'En col tenemos ya las lineas. Veremos cuantas hacen falta
'    J = Col.Count \ 2
'
'
'
'    For i = 1 To J
'        Aux = cad & """"
'        If J > 1 Then Aux = Aux & i & "/" & J
'        Aux = Aux & """"
'
'        'El primer trozo de linea
'        Aux = Aux & Col.item((i * 2) - 1)
'
'        'El segundo trozo de linea
'        Aux = Aux & Col.item((i * 2))
'        Print #NF, Aux
'    Next
'
'
'
'    miRsAux.Close
'
'    Close #NF
'    NF = -1 'para que no vuelva a hacer close
'
'    FileCopy App.Path & "\datosMA.txt", Destino
'
'    ImprimirPaletNORMAL = True
'
'EImprimeEtiquetasPA:
'    If Err.Number <> 0 Then MuestraError Err.Number
'    If NF >= 0 Then Close #NF
'
'
'End Function


'Habra una etiqueta para cada trazabilidad del palet
'Es decir , si el palet tiene 50 cajas 35 L1
'                                      15 L2  sacara 2 etiquetas cada una con sus datos pero el mismos SSCC
Private Function ImprimirPaletOLIVELINE(IdPalet As Long) As Boolean
Dim NF As Integer
Dim Cad As String
Dim Destino As String
Dim Aux As String
Dim C2 As String
Dim I As Integer
Dim J As Integer
Dim CC As Byte
Dim TotalCajas As Integer
Dim Kilos As Long
Dim ParaCodigoBarras As String
    On Error GoTo EImprimeEtiquetasPA:
    ImprimirPaletOLIVELINE = False

    NF = -1
    
    'Elimino el anterior
    If Dir(App.Path & "\datosMA.txt", vbArchive) <> "" Then Kill App.Path & "\datosMA.txt"
    Set miRsAux = New ADODB.Recordset
    Cad = "Select carpetaSRV,pathArchBartender,extension,palets impresora FROM  prodparamimpr"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Ahora lo copiamos donde diga el path
    Cad = Format(Now, "yymmdd_hhnnss") & vUsu.PC & "."
    Cad = Cad & miRsAux!extension
    Destino = miRsAux!carpetaSRV & "\" & Cad
    
 
    NF = FreeFile
    Open App.Path & "\datosMA.txt" For Output As #NF
        
    'La primera linea es la de "ordenes"
    'Ejmplo
    '%BTW% /AF="C:\MisDoc\BarTender\Formats\ParaComanderMarca.btw" /P /D="%Trigger File Name%" /PRN="HP DeskJet 1220C" /R=3 /P %END%
    Cad = "PaletBuOliLine.btw"
    Cad = "%BTW% /AF=""" & miRsAux!pathArchBartender & "\" & Cad & """  /D=""%Trigger File Name%"" /PRN="""
    'Ahora la impresora de la linea       r3: en que linea empiezan los datos
    Cad = Cad & miRsAux!impresora & """ /R=3 /P "
    Print #NF, Cad
    Print #NF, "%END%"
    miRsAux.Close
    
    
    'Si NO tienen todos la misma fecha de produccion no la ponemos arriba en el encabezado
    Aux = "Select fhinicio,fhFin,CajasProd from prodpalets where idpalet = " & IdPalet
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    
    TotalCajas = miRsAux!Cajasprod
    
    miRsAux.Close
    
    
    Aux = "Select * from tmppartidas where codusu = " & vUsu.Codigo & " order by idpartida,idnumoperacion"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
   
 
    I = 0
    While Not miRsAux.EOF
            I = I + 1
          
          
            ' 15 Abril 2013
            ' Los barrados NO llevan los ids(parentesis), ademas
            ' DUN y SSCC no llevan el codigo de control, lo pone el
            '"MAESTRO OLIVA a.o.V.E. LATA 500 ML","05/04/2013","05/10/2014",
            '"005804","18436024291171","073","000","384125940000058764","00816",
            '"130405","141005","005804","1843602429117","073","000",
            '"38412594000005876","0816"
          
          
            'Grabaremos: nomartic,inicio(11),caduca(17),lote,dun,cajaspalet,refer[000],kilos(3300),sscc
                 

            Aux = """" & miRsAux!numLote & ""","""
            'Inicio
            C2 = "prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin and lotetraza"
            C2 = DevuelveDesdeBD(conAri, "fhinicio", "prodlin,prodtrazlin", C2, miRsAux!idPartida)
            If C2 = "" Then
                MsgBox "No se encuentra fecha produccion", vbExclamation
                '------------------------
                C2 = Now
            End If
            C2 = Format(C2, "dd/mm/yyyy")
            Aux = Aux & C2 & ""","""
            ParaCodigoBarras = ",""" & Format(C2, "yymmdd") & ""","   'Antes ",""(11)" & Format(C2, "yymmdd") & ""","
            
            'Fecha cad y lote
            Aux = Aux & Format(miRsAux!Fecha, "dd/mm/yyyy") & """,""" & Format(miRsAux!idPartida, "000000") & ""","""
            'ParaCodigoBarras = ParaCodigoBarras & """(17)" & Format(miRsAux!Fecha, "yymmdd") & """,""(10)" & Format(miRsAux!IdPartida, "000000") & ""","""
            ParaCodigoBarras = ParaCodigoBarras & """" & Format(miRsAux!Fecha, "yymmdd") & """,""" & Format(miRsAux!idPartida, "000000") & ""","""
           
           
            C2 = DevuelveDesdeBD(conAri, "caj_codun", "sarti4", "codartic", miRsAux!codartic, "T")
            If C2 = "" Then
                MsgBox "No se ha encontrado el codigo DUN para el articulo: " & DBLet(miRsAux!numLote, "T"), vbExclamation
                C2 = DevuelveDesdeBD(conAri, "codigoea", "sartic", "codartic", miRsAux!codartic, "T")
                If C2 = "" Then C2 = miRsAux!codartic
                
            End If
            Aux = Aux & C2 & """,""" & Format(miRsAux!Cantidad, "000") & """,""000"","""
            'Abril2013. Para el ean , el dun (y el SSCC) van sin digito de control
            ' Y todo va sin los ids
            If Len(C2) > 0 Then C2 = Mid(C2, 1, Len(C2) - 1)
            'ParaCodigoBarras = ParaCodigoBarras & "(02)" & C2 & """,""(37)" & Format(miRsAux!Cantidad, "000") & """,""(240)000"","""
            ParaCodigoBarras = ParaCodigoBarras & "" & C2 & """,""" & Format(miRsAux!Cantidad, "000") & """,""000"","""
           
            'DUN y cajas palet
           
            'Marzo 2013
            'El (00) es el sscc
            ' (00)38412594xxxxxxxxxC     '38412594=id Morales
            '      morales  palet  C:ontrol
            'Dejaremos los 2 priemros digitos  disponibles . Sera de momento 00 para produccion, pero nunca se sabe
            '          los 7 siguientes seran para idpalet
            
            C2 = "38412594" & "00" & Format(IdPalet, "0000000")
            CC = DevuelveDigitoControlSSCC(C2)
          
            Aux = Aux & C2 & CC & ""","""
            'ParaCodigoBarras = ParaCodigoBarras & "(00)" & C2 & ""","""
            ParaCodigoBarras = ParaCodigoBarras & "" & C2 & ""","""
            
            Kilos = DevuelvePesoPalet(miRsAux!codartic, CCur(miRsAux!Cantidad))
            
            Aux = Aux & Format(Kilos, "00000") & """"
            'ParaCodigoBarras = ParaCodigoBarras & "(3300)" & Format(Kilos, "0000") & """"
            ParaCodigoBarras = ParaCodigoBarras & "" & Format(Kilos, "0000") & """"

            Aux = Aux & ParaCodigoBarras
            Print #NF, Aux
            
            
        miRsAux.MoveNext
    Wend
    
   
    
    
    
    miRsAux.Close
    
    Close #NF
    NF = -1 'para que no vuelva a hacer close
    
    FileCopy App.Path & "\datosMA.txt", Destino
    
    ImprimirPaletOLIVELINE = True
    
EImprimeEtiquetasPA:
    If Err.Number <> 0 Then MuestraError Err.Number
    If NF >= 0 Then Close #NF


End Function



'***  Es una copia de impresion de palet OLIVE LINE
'
'
'Habra una etiqueta para cada trazabilidad del palet
'Es decir , si el palet tiene 50 cajas 35 L1
'                                      15 L2  sacara 2 etiquetas cada una con sus datos pero el mismos SSCC
Private Function ImprimirPaletNUEVO(IdPalet As Long) As Boolean
Dim NF As Integer
Dim Cad As String
Dim Destino As String
Dim Aux As String
Dim C2 As String
Dim I As Integer
Dim J As Integer
Dim CC As Byte
Dim TotalCajas As Integer
Dim Kilos As Long
Dim ParaCodigoBarras As String
    On Error GoTo EImprimeEtiquetasPA:
    ImprimirPaletNUEVO = False

    NF = -1
    
    'Elimino el anterior
    If Dir(App.Path & "\datosMA.txt", vbArchive) <> "" Then Kill App.Path & "\datosMA.txt"
    Set miRsAux = New ADODB.Recordset
    Cad = "Select carpetaSRV,pathArchBartender,extension,palets impresora FROM  prodparamimpr"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Ahora lo copiamos donde diga el path
    Cad = Format(Now, "yymmdd_hhnnss") & vUsu.PC & "."
    Cad = Cad & miRsAux!extension
    Destino = miRsAux!carpetaSRV & "\" & Cad
    
 
    NF = FreeFile
    Open App.Path & "\datosMA.txt" For Output As #NF
        
    'La primera linea es la de "ordenes"
    'Ejmplo
    '%BTW% /AF="C:\MisDoc\BarTender\Formats\ParaComanderMarca.btw" /P /D="%Trigger File Name%" /PRN="HP DeskJet 1220C" /R=3 /P %END%
    'cad = "PaletBuOliLine.btw"
    'MAYO 2014
    Cad = "PaletBu.btw"
    
    Cad = "%BTW% /AF=""" & miRsAux!pathArchBartender & "\" & Cad & """  /D=""%Trigger File Name%"" /PRN="""
    'Ahora la impresora de la linea       r3: en que linea empiezan los datos
    Cad = Cad & miRsAux!impresora & """ /R=3 /P "
    Print #NF, Cad
    Print #NF, "%END%"
    miRsAux.Close
    
    
    'Si NO tienen todos la misma fecha de produccion no la ponemos arriba en el encabezado
    Aux = "Select fhinicio,fhFin,CajasProd from prodpalets where idpalet = " & IdPalet
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    
    TotalCajas = miRsAux!Cajasprod
    
    miRsAux.Close
    
    
    Aux = "Select * from tmppartidas where codusu = " & vUsu.Codigo & " order by idpartida,idnumoperacion"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
   
 
    I = 0
    While Not miRsAux.EOF
            I = I + 1
          
          
            ' 15 Abril 2013
            ' Los barrados NO llevan los ids(parentesis), ademas
            ' DUN y SSCC no llevan el codigo de control, lo pone el
            '"MAESTRO OLIVA a.o.V.E. LATA 500 ML","05/04/2013","05/10/2014",
            '"005804","18436024291171","073","000","384125940000058764","00816",
            '"130405","141005","005804","1843602429117","073","000",
            '"38412594000005876","0816"
          
          
            'Grabaremos: nomartic,inicio(11),caduca(17),lote,dun,cajaspalet,refer[000],kilos(3300),sscc
                 

            Aux = """" & miRsAux!numLote & ""","""
            'Inicio
            C2 = "prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin and lotetraza"
            C2 = DevuelveDesdeBD(conAri, "fhinicio", "prodlin,prodtrazlin", C2, miRsAux!idPartida)
            If C2 = "" Then
                MsgBox "No se encuentra fecha produccion", vbExclamation
                '------------------------
                C2 = Now
            End If
            C2 = Format(C2, "dd/mm/yyyy")
            Aux = Aux & C2 & ""","""
            ParaCodigoBarras = ",""" & Format(C2, "yymmdd") & ""","   'Antes ",""(11)" & Format(C2, "yymmdd") & ""","
            
            'Fecha cad y lote
            Aux = Aux & Format(miRsAux!Fecha, "dd/mm/yyyy") & """,""" & Format(miRsAux!idPartida, "000000") & ""","""
            'ParaCodigoBarras = ParaCodigoBarras & """(17)" & Format(miRsAux!Fecha, "yymmdd") & """,""(10)" & Format(miRsAux!IdPartida, "000000") & ""","""
            ParaCodigoBarras = ParaCodigoBarras & """" & Format(miRsAux!Fecha, "yymmdd") & """,""" & Format(miRsAux!idPartida, "000000") & ""","""
           
           
            C2 = DevuelveDesdeBD(conAri, "caj_codun", "sarti4", "codartic", miRsAux!codartic, "T")
            If C2 = "" Then
                MsgBox "No se ha encontrado el codigo DUN para el articulo: " & DBLet(miRsAux!numLote, "T"), vbExclamation
                C2 = DevuelveDesdeBD(conAri, "codigoea", "sartic", "codartic", miRsAux!codartic, "T")
                If C2 = "" Then C2 = miRsAux!codartic
                
            End If
            Aux = Aux & C2 & """,""" & Format(miRsAux!Cantidad, "000") & """,""000"","""
            'Abril2013. Para el ean , el dun (y el SSCC) van sin digito de control
            ' Y todo va sin los ids
            If Len(C2) > 0 Then C2 = Mid(C2, 1, Len(C2) - 1)
            'ParaCodigoBarras = ParaCodigoBarras & "(02)" & C2 & """,""(37)" & Format(miRsAux!Cantidad, "000") & """,""(240)000"","""
            ParaCodigoBarras = ParaCodigoBarras & "" & C2 & """,""" & Format(miRsAux!Cantidad, "000") & """,""000"","""
           
            'DUN y cajas palet
           
            'Marzo 2013
            'El (00) es el sscc
            ' (00)38412594xxxxxxxxxC     '38412594=id Morales
            '      morales  palet  C:ontrol
            'Dejaremos los 2 priemros digitos  disponibles . Sera de momento 00 para produccion, pero nunca se sabe
            '          los 7 siguientes seran para idpalet
            
            C2 = "38412594" & "00" & Format(IdPalet, "0000000")
            CC = DevuelveDigitoControlSSCC(C2)
          
            Aux = Aux & C2 & CC & ""","""
            'ParaCodigoBarras = ParaCodigoBarras & "(00)" & C2 & ""","""
            ParaCodigoBarras = ParaCodigoBarras & "" & C2 & ""","""
            
            Kilos = DevuelvePesoPalet(miRsAux!codartic, CCur(miRsAux!Cantidad))
            
            Aux = Aux & Format(Kilos, "00000") & """"
            'ParaCodigoBarras = ParaCodigoBarras & "(3300)" & Format(Kilos, "0000") & """"
            ParaCodigoBarras = ParaCodigoBarras & "" & Format(Kilos, "0000") & """"

            Aux = Aux & ParaCodigoBarras
            Print #NF, Aux
            
            
        miRsAux.MoveNext
    Wend
    
   
    
    
    
    miRsAux.Close
    
    Close #NF
    NF = -1 'para que no vuelva a hacer close
    
    FileCopy App.Path & "\datosMA.txt", Destino
    
    ImprimirPaletNUEVO = True
    
EImprimeEtiquetasPA:
    If Err.Number <> 0 Then MuestraError Err.Number
    If NF >= 0 Then Close #NF


End Function










Private Function DevuelvePesoPalet(ByRef codartic As String, CajasPal As Currency) As Long
Dim RN As ADODB.Recordset
Dim Aux As String
Dim Cajas As Currency

    Set RN = New ADODB.Recordset
    DevuelvePesoPalet = 0
    Aux = "Select * from sarti4 where codartic='" & codartic & "'"
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RN.EOF Then
        Cajas = DBLet(RN!pal_udbas, "N")
        Cajas = Cajas * DBLet(RN!pal_udalt, "N")
        If Cajas > 0 Then
            'Los transformamos en entero
            Cajas = DBLet(RN!pal_pbruto, "N") / Cajas
            Cajas = Cajas * CajasPal 'peso cajas por suma cajas palet
            
            'Cajas = Cajas + DBLet(RN!pal_pvaci, "N") 'Nunca lo he sumado. Por si hiciera falta lo dejo comentado
            DevuelvePesoPalet = Val(Cajas)
        End If
    End If
    RN.Close
    Set RN = Nothing
End Function
'Digitos17b  string de lenth=17
Private Function DevuelveDigitoControlSSCC(ByRef Digitos17b As String) As Byte
Dim I As Byte
Dim Tot As Integer
Dim N As Byte

    Tot = 0
    For I = 1 To 17 'SIEMPRE 17
        N = CByte(Mid(Digitos17b, I, 1))
        If (I Mod 2) = 0 Then
            'Par
            Tot = Tot + N
        Else
            'impar
            Tot = Tot + (3 * N)
        End If
    Next I
    
    Tot = 10 - (Tot Mod 10)
    If Tot = 10 Then Tot = 0
    
    DevuelveDigitoControlSSCC = CByte(Tot)
    
End Function
